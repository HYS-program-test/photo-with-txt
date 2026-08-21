"""
現場照片報告生成器
------------------
需求套件（requirements.txt）:
    streamlit
    python-docx
    pillow

執行方式（注意：這次多了一個本地元件資料夾，要跟 app.py 放在同一層）：
    app.py
    photo_dnd_component/
        __init__.py
        frontend/
            index.html
    streamlit run app.py

拖曳＋文字敘述：
    這次改用「真正的 Streamlit 自訂元件」（photo_dnd_component），純靜態 HTML/JS，
    不需要 Node/React 建置流程，用 Streamlit 官方的 postMessage 通訊協定跟 Python 溝通。
    好處是：資料同步時不會把整個 iframe 砍掉重建，畫面不會閃爍，拖曳/打字也不會被打斷。
    編號、拖曳目的地、文字敘述都在同一份 HTML 裡用 flex 排在同一列，保證對齊；
    縮圖池用 position:sticky 貼在一個內部可捲動箱子的頂端，捲動找編號時縮圖池會固定在頂端。

字型：預設嘗試載入「微軟正黑體 msjh.ttc」，若伺服器沒有該字型會退回 Pillow 的可縮放
預設字型（不支援中文）。要讓中文字放大生效，請把中文字型檔案放到跟 app.py 同一層。
"""

import base64
import hashlib
import io
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

from photo_dnd_component import photo_dnd

st.set_page_config(page_title="現場照片報告生成器", page_icon="📱", layout="centered")

FONT_PATHS = ["msjh.ttc", "msjhbd.ttc", "TaipeiSansTCBeta-Bold.ttf", "NotoSansTC-Regular.otf"]
MAX_DIMENSION = 1920  # 輸出照片長邊上限
THUMB_SIZE = 160  # 拖曳用縮圖的邊長，縮小一點可以減少傳輸資料量

# -----------------------------------------------------------------------------
# 1. Session State
# -----------------------------------------------------------------------------
if "rows_count" not in st.session_state:
    st.session_state.rows_count = 5
if "assignments" not in st.session_state:
    st.session_state.assignments = {}  # {"01": ["Pxxxxxxxxxx", ...], ...}
if "desc_text" not in st.session_state:
    st.session_state.desc_text = {}  # {"01": "文字敘述...", ...}


def make_photo_key(f):
    h = hashlib.md5(f"{f.name}|{f.size}".encode("utf-8")).hexdigest()[:10]
    return f"P{h}"


def make_thumb_b64(f):
    img = Image.open(io.BytesIO(f.getvalue())).convert("RGB")
    img.thumbnail((THUMB_SIZE, THUMB_SIZE))
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=80)
    return base64.b64encode(out.getvalue()).decode()


# -----------------------------------------------------------------------------
# 2. 照片壓印文字功能：白色 60% 透明底、靠左對齊、微軟正黑體
# -----------------------------------------------------------------------------
def process_photo_with_text(image_bytes, text):
    image = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
    width, height = image.size

    if max(width, height) > MAX_DIMENSION:
        scale = MAX_DIMENSION / max(width, height)
        image = image.resize((int(width * scale), int(height * scale)), Image.LANCZOS)
        width, height = image.size

    overlay = Image.new("RGBA", image.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)

    font_size = max(32, int(min(width, height) * 0.056))
    font = None
    font_used = None
    for path in FONT_PATHS:
        try:
            font = ImageFont.truetype(path, font_size)
            font_used = path
            break
        except IOError:
            continue
    if font is None:
        try:
            font = ImageFont.load_default(size=font_size)
            font_used = "PIL 內建可縮放預設字型（Pillow>=10.1；不支援中文字元）"
        except TypeError:
            font = ImageFont.load_default()
            font_used = "PIL 內建固定大小預設字型（Pillow 版本太舊，不吃 size，也不支援中文）"
    st.session_state["_last_font_used"] = font_used

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

    draw.rectangle([x1, y1, x2, y2], fill=(255, 255, 255, 153))

    curr_y = y1 + padding
    for line in lines:
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


def build_download_all_html(files):
    """files: list of {"name":..., "data": base64_jpeg_str}。一顆按鈕依序觸發每張照片各自下載。"""
    import json as _json
    files_json = _json.dumps(files)
    return f"""
    <meta charset="utf-8">
    <style>
      body {{ margin:0; font-family:"Microsoft JhengHei", -apple-system, sans-serif; overflow:visible; }}
      #dl-all-btn {{ width:100%; padding:12px; font-size:15px; font-weight:600; color:#111827;
          background:#f3f5f8; border:1px solid #d8dde5; border-radius:8px; cursor:pointer;
          box-sizing:border-box; line-height:1.4; }}
      #dl-all-btn:active {{ background:#e7ebf1; }}
      #dl-all-status {{ font-size:12px; color:#6b7280; margin-top:8px; min-height:14px; }}
    </style>
    <button id="dl-all-btn">⬇️ 一鍵下載全部照片（共 {len(files)} 張）</button>
    <div id="dl-all-status"></div>
    <script>
      var files = {files_json};
      document.getElementById('dl-all-btn').addEventListener('click', function () {{
        var statusEl = document.getElementById('dl-all-status');
        var i = 0;
        function next() {{
          if (i >= files.length) {{
            statusEl.textContent = '完成，共下載 ' + files.length + ' 張照片。';
            return;
          }}
          var f = files[i];
          statusEl.textContent = '下載中… (' + (i + 1) + '/' + files.length + ') ' + f.name;
          var a = document.createElement('a');
          a.href = 'data:image/jpeg;base64,' + f.data;
          a.download = f.name;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          i += 1;
          setTimeout(next, 350);
        }}
        next();
      }});
    </script>
    """


# -----------------------------------------------------------------------------
# 3. CSS：窄版手機畫面
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
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# 4. 凍結區塊：上傳按鈕（藍色線以上）
# -----------------------------------------------------------------------------
st.markdown('<div class="frozen-header">', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "照片上傳 (1~30張)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="uploader",
)

photo_map = {}  # {photo_key: UploadedFile, ...}

if uploaded_files:
    if len(uploaded_files) > 30:
        st.warning("最多支援 30 張照片，僅取前 30 張。")
        uploaded_files = uploaded_files[:30]

    for f in uploaded_files:
        photo_map[make_photo_key(f)] = f

    st.caption(f"已上傳 {len(uploaded_files)} 張照片，請在下方拖曳縮圖到各編號列。")

st.markdown("</div>", unsafe_allow_html=True)  # 藍色線結束點

# -----------------------------------------------------------------------------
# 5. 拖曳指派 + 文字敘述（自訂元件）
# -----------------------------------------------------------------------------
if uploaded_files:
    row_keys = [f"{i:02d}" for i in range(1, st.session_state.rows_count + 1)]

    # 清掉已被移除照片的殘留指派
    pruned = {
        rk: [p for p in v if p in photo_map]
        for rk, v in st.session_state.assignments.items()
    }
    if pruned != st.session_state.assignments:
        st.session_state.assignments = pruned

    photos_b64 = {k: make_thumb_b64(f) for k, f in photo_map.items()}

    result = photo_dnd(
        photos=photos_b64,
        assignments=st.session_state.assignments,
        descriptions=st.session_state.desc_text,
        row_keys=row_keys,
        key="photo_dnd_widget",
    )

    if result:
        new_assign = result.get("assignments", {})
        new_desc = result.get("descriptions", {})
        st.session_state.assignments = {
            rk: [p for p in v if p in photo_map]
            for rk, v in new_assign.items()
            if rk in row_keys
        }
        st.session_state.desc_text = {
            rk: v for rk, v in new_desc.items() if rk in row_keys
        }

    with st.expander("🔧 除錯資訊（測試用，確認沒問題後可以請我刪掉）"):
        st.write("元件最新回傳值：")
        st.json(result)
        st.write("目前 session_state.assignments：")
        st.json(st.session_state.assignments)
        st.write("目前 session_state.desc_text：")
        st.json(st.session_state.desc_text)
        st.write("目前 photo_map 的 key：")
        st.code(list(photo_map.keys()))
else:
    st.info("請先上傳照片")

if st.button("➕ 新增項目"):
    st.session_state.rows_count += 1
    st.rerun()

st.divider()

# -----------------------------------------------------------------------------
# 6. 匯出下載
# -----------------------------------------------------------------------------
st.subheader("📥 下載專區")
if "_last_font_used" in st.session_state:
    st.caption(f"（上次壓字使用的字型：{st.session_state['_last_font_used']}）")
today_str = datetime.now().strftime("%Y%m%d")

dl_mode = st.radio(
    "照片下載模式：",
    ["📱 手機模式 (一鍵下載全部照片，非壓縮檔)", "💻 電腦模式 (下載 ZIP 壓縮檔)"],
    horizontal=False,
)


def render_word_download_button():
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
                desc = st.session_state.desc_text.get(row_num, "")
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


if dl_mode == "💻 電腦模式 (下載 ZIP 壓縮檔)":
    col_dl_photo, col_dl_word = st.columns(2)
    with col_dl_photo:
        if st.button("下載照片"):
            if not photo_map or not any(st.session_state.assignments.values()):
                st.warning("請先上傳照片並指派到列表！")
            else:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for idx in range(1, st.session_state.rows_count + 1):
                        row_num = f"{idx:02d}"
                        pkeys = st.session_state.assignments.get(row_num, [])
                        desc = st.session_state.desc_text.get(row_num, "")
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
    with col_dl_word:
        render_word_download_button()
else:
    if not photo_map or not any(st.session_state.assignments.values()):
        st.warning("請先上傳照片並指派到列表！")
    else:
        files_payload = []
        for idx in range(1, st.session_state.rows_count + 1):
            row_num = f"{idx:02d}"
            pkeys = st.session_state.assignments.get(row_num, [])
            desc = st.session_state.desc_text.get(row_num, "")
            for n, pkey in enumerate(pkeys, start=1):
                if pkey not in photo_map:
                    continue
                img_out = process_photo_with_text(photo_map[pkey].getvalue(), desc)
                suffix = f"-{n}" if len(pkeys) > 1 else ""
                fname = f"{today_str}-{row_num}{suffix}.jpg"
                files_payload.append({
                    "name": fname,
                    "data": base64.b64encode(img_out.getvalue()).decode(),
                })
        if files_payload:
            components.html(build_download_all_html(files_payload), height=150, scrolling=True)
    st.write("")
    render_word_download_button()
