import base64
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

st.set_page_config(
    page_title="現場照片報告生成器", page_icon="📱", layout="centered"
)

# -----------------------------------------------------------------------------
# 1. Session State 初始化
# -----------------------------------------------------------------------------
if "rows_count" not in st.session_state:
    st.session_state.rows_count = 5


# -----------------------------------------------------------------------------
# 2. 影像壓印文字 (白色 60% 透明度背景 + 微軟正黑體)
# -----------------------------------------------------------------------------
def process_photo_with_text(image_bytes, text):
    image = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
    width, height = image.size

    overlay = Image.new("RGBA", image.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)

    font_size = max(18, int(min(width, height) * 0.038))
    font_paths = ["msjh.ttc", "msjhbd.ttc", "TaipeiSansTCBeta-Bold.ttf"]
    font = None
    for path in font_paths:
        try:
            font = ImageFont.truetype(path, font_size)
            break
        except IOError:
            continue
    if font is None:
        font = ImageFont.load_default()

    margin = int(min(width, height) * 0.03)
    padding = 12

    lines = text.split("\n") if text else [" "]
    line_heights = [
        font.getbbox(line)[3] - font.getbbox(line)[1] for line in lines
    ]
    max_line_width = max(
        [font.getbbox(line)[2] - font.getbbox(line)[0] for line in lines]
    )

    box_w = max_line_width + (padding * 2)
    box_h = sum(line_heights) + (padding * 2) + (len(lines) - 1) * 6

    x1 = margin
    y2 = height - margin
    x2 = x1 + box_w
    y1 = y2 - box_h

    # 60% 透明度白色背景
    draw.rectangle([x1, y1, x2, y2], fill=(255, 255, 255, 153))

    curr_y = y1 + padding
    for line in lines:
        draw.text((x1 + padding, curr_y), line, fill=(0, 0, 0, 255), font=font)
        curr_y += (font.getbbox(line)[3] - font.getbbox(line)[1]) + 6

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
# 3. CSS 注入：確保藍色分割線上方完全固定凍結
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main .block-container {
        max-width: 550px !important;
        padding-top: 10px !important;
        padding-bottom: 2rem;
    }
    
    /* 藍色分割線上方凍結容器 */
    .frozen-top-container {
        position: sticky;
        top: 0;
        background-color: #ffffff;
        z-index: 9999;
        padding-top: 10px;
        padding-bottom: 10px;
        border-bottom: 3px solid #3388ff; /* 藍色分隔線 */
    }

    /* 單行滾動縮圖 */
    .thumb-scroll-container {
        display: flex;
        overflow-x: auto;
        gap: 8px;
        padding: 5px 0;
    }
    .thumb-scroll-container img {
        width: 65px;
        height: 65px;
        object-fit: cover;
        border-radius: 6px;
        border: 1px solid #ccc;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# 4. 藍色線以上：凍結區域 (上傳 + 縮圖)
# -----------------------------------------------------------------------------
st.markdown('<div class="frozen-top-container">', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "照片上傳 (1~30張)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="uploader",
)

if uploaded_files:
    thumb_html_list = []
    for idx, f in enumerate(uploaded_files[:30]):
        b64_str = base64.b64encode(f.getvalue()).decode("utf-8")
        thumb_html_list.append(
            f'<img src="data:image/jpeg;base64,{b64_str}" title="照片 {idx+1:02d}">'
        )

    st.markdown(
        f'<div class="thumb-scroll-container">{"".join(thumb_html_list)}</div>',
        unsafe_allow_html=True,
    )

st.markdown("</div>", unsafe_allow_html=True)  # 凍結區結束 (藍色線)

# -----------------------------------------------------------------------------
# 5. 藍色線下方：數字清單、照片分配與文字框
# -----------------------------------------------------------------------------
st.write(" ")
photo_options = ["(未選擇)"] + [
    f"照片 {i+1:02d} - {f.name}" for i, f in enumerate(uploaded_files or [])
]

for i in range(1, st.session_state.rows_count + 1):
    row_num = f"{i:02d}"

    col_num, col_select, col_txt = st.columns([0.6, 1.8, 2.6])

    with col_num:
        st.markdown(f"### {row_num}")

    with col_select:
        selected = st.selectbox(
            f"選擇照片 {row_num}",
            options=photo_options,
            key=f"select_{i}",
            label_visibility="collapsed",
        )

        if selected != "(未選擇)" and uploaded_files:
            img_idx = photo_options.index(selected) - 1
            st.image(uploaded_files[img_idx], use_container_width=True)

    with col_txt:
        desc = st.text_area(
            f"敘述 {row_num}",
            placeholder="文字敘述...",
            key=f"desc_{i}",
            height=85,
            label_visibility="collapsed",
        )

# '+' 按鈕增加列數
if st.button("➕"):
    st.session_state.rows_count += 1
    st.rerun()

st.divider()

# -----------------------------------------------------------------------------
# 6. 下載區
# -----------------------------------------------------------------------------
st.subheader("📥 下載專區")
today_str = datetime.now().strftime("%Y%m%d")

dl_mode = st.radio(
    "照片下載模式：",
    ["📱 手機模式 (直接下載單張)", "💻 電腦模式 (下載 ZIP 壓縮檔)"],
    horizontal=True,
)

col_dl_photo, col_dl_word = st.columns(2)

# 下載照片
with col_dl_photo:
    if dl_mode == "💻 電腦模式 (下載 ZIP 壓縮檔)":
        if st.button("下載照片"):
            if not uploaded_files:
                st.warning("請先上傳照片！")
            else:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(
                    zip_buffer, "w", zipfile.ZIP_DEFLATED
                ) as zf:
                    for idx in range(1, st.session_state.rows_count + 1):
                        sel = st.session_state.get(f"select_{idx}", "(未選擇)")
                        desc = st.session_state.get(f"desc_{idx}", "")
                        if sel != "(未選擇)":
                            img_idx = photo_options.index(sel) - 1
                            img_out = process_photo_with_text(
                                uploaded_files[img_idx].getvalue(), desc
                            )
                            zf.writestr(
                                f"{today_str}-{idx:02d}.jpg", img_out.getvalue()
                            )

                zip_buffer.seek(0)
                st.download_button(
                    "⬇️ 點此下載照片 (ZIP)",
                    data=zip_buffer,
                    file_name=f"Photos_{today_str}.zip",
                    mime="application/zip",
                )
    else:
        for idx in range(1, st.session_state.rows_count + 1):
            sel = st.session_state.get(f"select_{idx}", "(未選擇)")
            desc = st.session_state.get(f"desc_{idx}", "")
            if sel != "(未選擇)" and uploaded_files:
                img_idx = photo_options.index(sel) - 1
                img_out = process_photo_with_text(
                    uploaded_files[img_idx].getvalue(), desc
                )
                st.download_button(
                    label=f"⬇️ 下載 {today_str}-{idx:02d}.jpg",
                    data=img_out,
                    file_name=f"{today_str}-{idx:02d}.jpg",
                    mime="image/jpeg",
                    key=f"dl_single_{idx}",
                )

# 下載 Word
with col_dl_word:
    if st.button("下載word", use_container_width=True):
        if not uploaded_files:
            st.warning("請先上傳照片！")
        else:
            doc = Document()

            for section in doc.sections:
                section.top_margin = Inches(0.5)
                section.bottom_margin = Inches(0.5)
                section.left_margin = Inches(0.5)
                section.right_margin = Inches(0.5)

            items_to_print = []
            for idx in range(1, st.session_state.rows_count + 1):
                sel = st.session_state.get(f"select_{idx}", "(未選擇)")
                desc = st.session_state.get(f"desc_{idx}", "")
                if sel != "(未選擇)":
                    img_idx = photo_options.index(sel) - 1
                    img_out = process_photo_with_text(
                        uploaded_files[img_idx].getvalue(), desc
                    )
                    items_to_print.append(img_out)

            for page_start in range(0, len(items_to_print), 4):
                if page_start > 0:
                    doc.add_page_break()

                page_items = items_to_print[page_start : page_start + 4]
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
                label=f"⬇️ 點此下載 {today_str}_01.docx",
                data=doc_io,
                file_name=f"{today_str}_01.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
