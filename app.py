"""
現場照片報告生成器
------------------
需求套件（requirements.txt）:
    streamlit
    streamlit-sortables
    python-docx
    pillow

執行方式：
    streamlit run app.py

字型：預設嘗試載入「微軟正黑體 msjh.ttc」，若部署主機（例如 Linux 伺服器）沒有該字型，
請將 msjh.ttc 或 TaipeiSansTCBeta-Bold.ttf 等中文字型檔案放在與 app.py 同一資料夾，
或修改 FONT_PATHS 指向實際字型路徑，否則文字會退回系統預設字型（可能無法顯示中文）。
"""

import io
import zipfile
from datetime import datetime

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from PIL import Image, ImageDraw, ImageFont
from streamlit_sortables import sort_items

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
# 3. CSS：窄版手機畫面 + 頂部凍結區塊（下緣藍線）
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
    .thumb-scroll {
        display: flex;
        overflow-x: auto;
        gap: 8px;
        padding: 6px 0 4px 0;
    }
    .thumb-scroll img {
        width: 56px;
        height: 56px;
        object-fit: cover;
        border-radius: 6px;
        border: 1px solid #ddd;
        flex: none;
    }
    .pool-caption { font-size: 12px; color: #888; margin: 2px 0 0 0; }
    </style>
    """,
    unsafe_allow_html=True,
)

sortable_style = """
.sortable-component { padding: 0; }
.sortable-container { display: flex; align-items: flex-start; gap: 10px;
    border-bottom: 1px solid #eee; padding: 10px 0; min-height: 78px; }
.sortable-container-header { flex: 0 0 46px; font-size: 20px; font-weight: 700;
    color: #111827; padding-top: 6px; }
.sortable-container-body { flex: 1; background-color: #f3f5f8; border-radius: 8px;
    padding: 8px; min-height: 60px; display: flex; flex-wrap: wrap; gap: 6px; align-content: flex-start; }
.sortable-container-body:empty::after { content: "尚未指派照片"; color: #9aa4b2; font-size: 14px; }
.sortable-item { background-color: #e7edfb; border: 1px solid #b9c8ee; border-radius: 6px;
    padding: 4px 10px; font-size: 13px; color: #333; }
"""

# -----------------------------------------------------------------------------
# 4. 凍結區塊：上傳 + 縮圖預覽（藍色線以上）
# -----------------------------------------------------------------------------
st.markdown('<div class="frozen-header">', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "照片上傳 (1~30張)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="uploader",
)

photo_map = {}   # {"P01": UploadedFile, ...}
photo_thumb_label = {}

if uploaded_files:
    if len(uploaded_files) > 30:
        st.warning("最多支援 30 張照片，僅取前 30 張。")
        uploaded_files = uploaded_files[:30]

    signature = tuple((f.name, f.size) for f in uploaded_files)
    if signature != st.session_state.uploaded_signature:
        # 新的一批照片，重置已分配的項目
        st.session_state.uploaded_signature = signature
        st.session_state.assignments = {}

    for idx, f in enumerate(uploaded_files):
        key = f"P{idx + 1:02d}"
        photo_map[key] = f
        photo_thumb_label[key] = f"照片{idx + 1:02d}"

    # 靜態縮圖列（僅供辨識用，實際拖曳用下方的可拖曳清單）
    thumb_htmls = [
        f'<img src="data:image/jpeg;base64,{__import__("base64").b64encode(f.getvalue()).decode()}" title="照片{idx+1:02d}">'
        for idx, f in enumerate(uploaded_files)
    ]
    st.markdown(f'<div class="thumb-scroll">{"".join(thumb_htmls)}</div>', unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)  # 藍色線結束點

# -----------------------------------------------------------------------------
# 5. 拖曳指派 + 文字敘述：左側「編號＋拖曳目的地（尚未指派照片）」／右側「文字敘述」
#    每一列的拖曳目的地就是「尚未指派照片」文字出現的位置，可放單張或多張照片
# -----------------------------------------------------------------------------
if uploaded_files:
    assignments = st.session_state.assignments
    assigned_keys = {k for keys in assignments.values() for k in keys}
    pool_items = [photo_thumb_label[k] for k in photo_map if k not in assigned_keys]

    # 反查：顯示文字 -> photo key
    label_to_key = {v: k for k, v in photo_thumb_label.items()}

    row_keys = [f"{i:02d}" for i in range(1, st.session_state.rows_count + 1)]
    containers = [{"header": "待分配", "items": pool_items}]
    for rk in row_keys:
        current_keys = assignments.get(rk, [])
        items = [photo_thumb_label[k] for k in current_keys if k in photo_map]
        containers.append({"header": rk, "items": items})

    col_drop, col_desc = st.columns([2.3, 2.7])

    with col_drop:
        result = sort_items(containers, multi_containers=True, direction="horizontal", custom_style=sortable_style)

        new_assignments = {}
        if result:
            for container in result[1:]:
                row_key = container["header"]
                new_assignments[row_key] = [
                    label_to_key[lbl] for lbl in container["items"] if lbl in label_to_key
                ]

        if new_assignments != assignments:
            st.session_state.assignments = new_assignments
            st.rerun()

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
