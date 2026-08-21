"""
現場照片報告生成器
------------------
需求套件（requirements.txt）:
    streamlit
    python-docx
    pillow

執行方式：
    streamlit run app.py

拖曳功能說明（重要）：
    這一版拖曳是「真的縮圖」可拖曳，用的是原生 SortableJS（透過 CDN 載入，支援手機觸控），
    塞進一個 st.components.v1.html 的 iframe 裡執行。因為 Streamlit 官方沒有提供
    「純 Python 就能雙向溝通的自訂 HTML 元件」這種東西（那需要一整套 React 編譯流程），
    所以這裡用了一個業界常見但非官方保證的技巧：
        1. JS 把目前的拖放結果整理成 JSON
        2. 用「原生 setter + dispatchEvent」的方式，寫進一個隱藏的 st.text_input
        3. Streamlit 偵測到該欄位值改變 -> 觸發 rerun -> Python 讀出 JSON 更新 session_state
    這個技巧依賴瀏覽器 / Streamlit 前端版本的內部行為，不是 100% 官方保證穩定，
    如果部署後發現拖曳完沒有反應，最可能是這個同步環節在你的環境行為不同，
    需要打開瀏覽器 DevTools Console 看有沒有噴錯，再回來一起調整。

字型：預設嘗試載入「微軟正黑體 msjh.ttc」，若部署主機（例如 Linux 伺服器）沒有該字型，
請將 msjh.ttc 或 TaipeiSansTCBeta-Bold.ttf 等中文字型檔案放在與 app.py 同一資料夾，
或修改 FONT_PATHS 指向實際字型路徑，否則文字會退回系統預設字型（可能無法顯示中文）。
"""

import base64
import io
import json
import zipfile
from datetime import datetime

import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from PIL import Image, ImageDraw, ImageFont

st.set_page_config(page_title="現場照片報告生成器", page_icon="📱", layout="centered")

FONT_PATHS = ["msjh.ttc", "msjhbd.ttc", "TaipeiSansTCBeta-Bold.ttf", "NotoSansTC-Regular.otf"]

# -----------------------------------------------------------------------------
# 1. Session State
# -----------------------------------------------------------------------------
if "rows_count" not in st.session_state:
    st.session_state.rows_count = 5
if "assignments" not in st.session_state:
    st.session_state.assignments = {}  # {"01": ["P01", "P03"], "02": [], ...}
if "uploaded_signature" not in st.session_state:
    st.session_state.uploaded_signature = None
if "dnd_sync_last" not in st.session_state:
    st.session_state.dnd_sync_last = ""


# -----------------------------------------------------------------------------
# 2. 照片壓印文字功能：白色 60% 透明底、靠左對齊、微軟正黑體
# -----------------------------------------------------------------------------
def process_photo_with_text(image_bytes, text):
    image = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
    width, height = image.size

    overlay = Image.new("RGBA", image.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)

    # 字級放大 10 倍（原本 0.028 比例 / 最小 16px → 現在 0.28 比例 / 最小 160px）
    font_size = max(160, int(min(width, height) * 0.28))
    font = None
    for path in FONT_PATHS:
        try:
            font = ImageFont.truetype(path, font_size)
            break
        except IOError:
            continue
    if font is None:
        font = ImageFont.load_default()

    margin = int(min(width, height) * 0.025)
    padding = max(10, int(font_size * 0.2))

    lines = text.split("\n") if text else [" "]
    line_heights = [font.getbbox(line)[3] - font.getbbox(line)[1] for line in lines]
    max_line_width = max(font.getbbox(line)[2] - font.getbbox(line)[0] for line in lines)

    box_w = max_line_width + (padding * 2)
    box_h = sum(line_heights) + (padding * 2) + (len(lines) - 1) * 5

    x1 = margin
    y2 = height - margin
    x2 = x1 + box_w
    y1 = y2 - box_h

    # 白色 60% 透明背景 (alpha = 153)
    draw.rectangle([x1, y1, x2, y2], fill=(255, 255, 255, 153))

    curr_y = y1 + padding
    for line in lines:
        # 靠左對齊：文字一律從 x1 + padding 開始畫
        draw.text((x1 + padding, curr_y), line, fill=(0, 0, 0, 255), font=font)
        curr_y += (font.getbbox(line)[3] - font.getbbox(line)[1]) + 5

    combined = Image.alpha_composite(image, overlay)
    output = io.BytesIO()
    combined.convert("RGB").save(output, format="JPEG", quality=95)
    output.seek(0)
    return output


def set_cell_border_none(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for border_name in ["top", "left", "bottom", "right"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "none")
        tcBorders.append(border)
    tcPr.append(tcBorders)


# -----------------------------------------------------------------------------
# 3. 產生拖曳區的 HTML/JS（真縮圖、SortableJS、觸控可用）
# -----------------------------------------------------------------------------
def build_dnd_html(photo_map, assignments, row_keys):
    assigned_keys = {k for keys in assignments.values() for k in keys}
    pool_keys = [k for k in photo_map if k not in assigned_keys]

    def chip(key):
        b64 = base64.b64encode(photo_map[key].getvalue()).decode()
        return f'<div class="photo-chip" data-photo="{key}"><img src="data:image/jpeg;base64,{b64}"></div>'

    pool_html = "".join(chip(k) for k in pool_keys)

    rows_html = ""
    for rk in row_keys:
        keys = assignments.get(rk, [])
        items_html = "".join(chip(k) for k in keys if k in photo_map)
        rows_html += f"""
        <div class="row-block">
            <div class="row-num">{rk}</div>
            <div class="dnd-list row-drop" data-row="{rk}">{items_html}</div>
        </div>"""

    return f"""
    <style>
      html, body {{ margin:0; padding:0; font-family: "Microsoft JhengHei", -apple-system, sans-serif; }}
      .pool-wrap {{ display:flex; flex-wrap:wrap; gap:8px; padding:8px 4px 14px 4px;
          border-bottom:3px solid #2b7cff; margin-bottom:10px; min-height:60px; }}
      .pool-wrap:empty::after {{ content:"（全部照片已指派完畢）"; color:#9aa4b2; font-size:13px; }}
      .photo-chip {{ width:60px; height:60px; border-radius:6px; overflow:hidden;
          border:1px solid #ccc; cursor:grab; flex:none; touch-action:none; }}
      .photo-chip img {{ width:100%; height:100%; object-fit:cover; display:block; pointer-events:none; }}
      .row-block {{ display:flex; align-items:flex-start; gap:10px; border-bottom:1px solid #eee; padding:10px 4px; }}
      .row-num {{ flex:0 0 34px; font-size:20px; font-weight:700; color:#111827; padding-top:16px; }}
      .row-drop {{ flex:1; min-height:64px; background:#f3f5f8; border-radius:8px;
          padding:6px; display:flex; flex-wrap:wrap; gap:6px; touch-action:none; }}
      .row-drop:empty::after {{ content:"尚未指派照片"; color:#9aa4b2; font-size:14px; line-height:64px; }}
      .sortable-ghost {{ opacity:0.3; }}
    </style>
    <div class="pool-wrap dnd-list" data-row="__pool__">{pool_html}</div>
    <div>{rows_html}</div>
    <script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.0/Sortable.min.js"></script>
    <script>
    (function () {{
      function syncState() {{
        var state = {{}};
        document.querySelectorAll('.row-drop').forEach(function (el) {{
          var row = el.getAttribute('data-row');
          state[row] = Array.from(el.querySelectorAll('.photo-chip')).map(function (c) {{
            return c.getAttribute('data-photo');
          }});
        }});
        var jsonStr = JSON.stringify(state);
        try {{
          var doc = window.parent.document;
          var target = doc.querySelector('input[aria-label="dnd_sync_field"]');
          if (target) {{
            var setter = Object.getOwnPropertyDescriptor(window.parent.HTMLInputElement.prototype, 'value').set;
            setter.call(target, jsonStr);
            target.dispatchEvent(new Event('input', {{ bubbles: true }}));
          }}
        }} catch (e) {{
          console.error('dnd sync failed', e);
        }}
      }}
      function initSortable() {{
        document.querySelectorAll('.dnd-list').forEach(function (el) {{
          new Sortable(el, {{
            group: 'photos',
            animation: 150,
            ghostClass: 'sortable-ghost',
            onEnd: syncState
          }});
        }});
      }}
      initSortable();
    }})();
    </script>
    """


# -----------------------------------------------------------------------------
# 4. CSS：窄版手機畫面 + 隱藏同步用的 text_input
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main .block-container {
        max-width: 550px !important;
        padding-top: 0.5rem !important;
        padding-bottom: 2rem;
    }
    .frozen-header {
        position: sticky;
        top: 0;
        background-color: #ffffff;
        z-index: 9999;
        padding-top: 10px;
        padding-bottom: 6px;
        border-bottom: 3px solid #2b7cff !important;
        margin-bottom: 15px;
    }
    input[aria-label="dnd_sync_field"] {
        opacity: 0;
        position: absolute;
        height: 1px;
        width: 1px;
        pointer-events: none;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# 5. 凍結區塊：上傳按鈕（藍色線以上）
# -----------------------------------------------------------------------------
st.markdown('<div class="frozen-header">', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "照片上傳 (1~30張)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="uploader",
)

photo_map = {}  # {"P01": UploadedFile, ...}

if uploaded_files:
    if len(uploaded_files) > 30:
        st.warning("最多支援 30 張照片，僅取前 30 張。")
        uploaded_files = uploaded_files[:30]

    signature = tuple((f.name, f.size) for f in uploaded_files)
    if signature != st.session_state.uploaded_signature:
        st.session_state.uploaded_signature = signature
        st.session_state.assignments = {}

    for idx, f in enumerate(uploaded_files):
        photo_map[f"P{idx + 1:02d}"] = f

    st.caption(f"已上傳 {len(uploaded_files)} 張照片，請在下方拖曳縮圖到各編號列。")

st.markdown("</div>", unsafe_allow_html=True)  # 藍色線結束點

# -----------------------------------------------------------------------------
# 6. 拖曳指派 + 文字敘述
# -----------------------------------------------------------------------------
if uploaded_files:
    row_keys = [f"{i:02d}" for i in range(1, st.session_state.rows_count + 1)]

    # 隱藏欄位：JS 拖放結果透過它同步回 Python（不做整頁 reload，避免清空其他欄位資料）
    st.text_input("dnd_sync_field", key="dnd_sync_raw", label_visibility="collapsed")

    raw = st.session_state.get("dnd_sync_raw", "")
    if raw and raw != st.session_state.dnd_sync_last:
        try:
            parsed = json.loads(raw)
            cleaned = {
                rk: [p for p in v if p in photo_map]
                for rk, v in parsed.items()
                if rk in row_keys
            }
            st.session_state.assignments = cleaned
            st.session_state.dnd_sync_last = raw
        except (json.JSONDecodeError, AttributeError):
            pass

    assignments = st.session_state.assignments

    col_drop, col_desc = st.columns([2.3, 2.7])

    with col_drop:
        html_code = build_dnd_html(photo_map, assignments, row_keys)
        pool_count = len(photo_map) - sum(len(v) for v in assignments.values())
        pool_lines = max(1, -(-max(pool_count, 1) // 6))
        total_height = 30 + pool_lines * 70 + len(row_keys) * 92 + 20
        components.html(html_code, height=total_height, scrolling=False)

    with col_desc:
        for i in range(1, st.session_state.rows_count + 1):
            row_num = f"{i:02d}"
            st.text_area(
                f"敘述 {row_num}",
                placeholder="文字敘述...",
                key=f"desc_{i}",
                height=90,
                label_visibility="collapsed",
            )
else:
    st.info("請先上傳照片")

if st.button("➕ 新增項目"):
    st.session_state.rows_count += 1
    st.rerun()

st.divider()

# -----------------------------------------------------------------------------
# 7. 匯出下載
# -----------------------------------------------------------------------------
st.subheader("📥 下載專區")
today_str = datetime.now().strftime("%Y%m%d")

dl_mode = st.radio(
    "照片下載模式：",
    ["📱 手機模式 (逐張下載，方便直接檢視)", "💻 電腦模式 (下載 ZIP 壓縮檔)"],
    horizontal=False,
)

col_dl_photo, col_dl_word = st.columns(2)

with col_dl_photo:
    if dl_mode == "💻 電腦模式 (下載 ZIP 壓縮檔)":
        if st.button("下載照片"):
            if not photo_map or not any(st.session_state.assignments.values()):
                st.warning("請先上傳照片並指派到列表！")
            else:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for idx in range(1, st.session_state.rows_count + 1):
                        row_num = f"{idx:02d}"
                        pkeys = st.session_state.assignments.get(row_num, [])
                        desc = st.session_state.get(f"desc_{idx}", "")
                        for n, pkey in enumerate(pkeys, start=1):
                            if pkey not in photo_map:
                                continue
                            img_out = process_photo_with_text(photo_map[pkey].getvalue(), desc)
                            suffix = f"-{n}" if len(pkeys) > 1 else ""
                            zf.writestr(f"{today_str}-{row_num}{suffix}.jpg", img_out.getvalue())
                zip_buffer.seek(0)
                st.download_button(
                    "⬇️ 下載照片 ZIP",
                    data=zip_buffer,
                    file_name=f"Photos_{today_str}.zip",
                    mime="application/zip",
                )
    else:
        for idx in range(1, st.session_state.rows_count + 1):
            row_num = f"{idx:02d}"
            pkeys = st.session_state.assignments.get(row_num, [])
            desc = st.session_state.get(f"desc_{idx}", "")
            for n, pkey in enumerate(pkeys, start=1):
                if pkey not in photo_map:
                    continue
                img_out = process_photo_with_text(photo_map[pkey].getvalue(), desc)
                suffix = f"-{n}" if len(pkeys) > 1 else ""
                st.download_button(
                    label=f"⬇️ 下載 {today_str}-{row_num}{suffix}.jpg",
                    data=img_out,
                    file_name=f"{today_str}-{row_num}{suffix}.jpg",
                    mime="image/jpeg",
                    key=f"dl_single_{idx}_{n}",
                )

with col_dl_word:
    if st.button("下載word", use_container_width=True):
        if not photo_map or not any(st.session_state.assignments.values()):
            st.warning("請先上傳照片並指派到列表！")
        else:
            doc = Document()
            for section in doc.sections:
                section.top_margin = Inches(0.5)
                section.bottom_margin = Inches(0.5)
                section.left_margin = Inches(0.5)
                section.right_margin = Inches(0.5)

            items_to_print = []
            for idx in range(1, st.session_state.rows_count + 1):
                row_num = f"{idx:02d}"
                pkeys = st.session_state.assignments.get(row_num, [])
                desc = st.session_state.get(f"desc_{idx}", "")
                for pkey in pkeys:
                    if pkey in photo_map:
                        img_out = process_photo_with_text(photo_map[pkey].getvalue(), desc)
                        items_to_print.append(img_out)

            if not items_to_print:
                st.warning("目前沒有已指派照片的項目！")
            else:
                for page_start in range(0, len(items_to_print), 4):
                    if page_start > 0:
                        doc.add_page_break()

                    page_items = items_to_print[page_start: page_start + 4]
                    table = doc.add_table(rows=2, cols=2)
                    table.autofit = False

                    for cell_idx, img_bytes in enumerate(page_items):
                        r = cell_idx // 2
                        c = cell_idx % 2
                        cell = table.cell(r, c)
                        set_cell_border_none(cell)

                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p.add_run()
                        run.add_picture(img_bytes, width=Inches(3.4))

                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)

                st.download_button(
                    label=f"⬇️ 下載 {today_str}_01.docx",
                    data=doc_io,
                    file_name=f"{today_str}_01.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
