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

# -----------------------------------------------------------------------------
# 1. Streamlit 頁面設定 (專為手機優化寬度)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="現場照片報告生成器", page_icon="📱", layout="centered"
)

# 載入 CSS：固定頂部凍結區塊、單行橫向滾動
st.markdown(
    """
    <style>
    /* 調整主要容器內襯，留出頂部固定區空間 */
    .main .block-container {
        max-width: 500px !important;
        padding-top: 170px !important; /* 留出空間給固定頂部 */
        padding-bottom: 2rem;
    }

    /* 頂部固定凍結區塊 */
    .fixed-top-header {
        position: fixed;
        top: 0;
        left: 50%;
        transform: translateX(-50%);
        width: 100%;
        max-width: 500px;
        background-color: #ffffff;
        z-index: 99999;
        padding: 10px 15px;
        box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.1);
        border-bottom: 2px solid #007bff;
    }

    /* 單行滾動縮圖區 */
    .thumb-scroll-container {
        display: flex !important;
        flex-wrap: nowrap !important;
        overflow-x: auto !important;
        gap: 8px;
        padding: 5px 0;
        -webkit-overflow-scrolling: touch;
    }

    .thumb-card {
        flex: 0 0 auto;
        width: 60px;
        height: 60px;
        border-radius: 8px;
        object-fit: cover;
        border: 2px solid #ccc;
        cursor: grab;
    }

    /* 按鈕優化 */
    .stButton button {
        width: 100%;
        border-radius: 8px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# 2. Session State 初始化
# -----------------------------------------------------------------------------
if "rows_count" not in st.session_state:
    st.session_state.rows_count = 5


# -----------------------------------------------------------------------------
# 3. 輔助函式：壓印文字 (白色 60% 透明度背景 + 微軟正黑體)
# -----------------------------------------------------------------------------
def process_photo_with_text(image_bytes, text):
    # 【修復重點】將 raw bytes 使用 io.BytesIO 包裝，避免 Image.open 讀取錯誤
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

    # 60% 透明度白色背景 (Alpha = 153)
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
# 4. 上半部：完全凍結區域 (照片上傳區 + 單行縮圖區)
# -----------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "上傳現場照片 (1~30張)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="uploader",
)

# 使用 HTML 注入構建凍結頂部與橫向單行縮圖
if uploaded_files:
    thumb_html_list = []
    for idx, f in enumerate(uploaded_files[:30]):
        # 【修復重點】正確進行 Base64 編碼
        b64_str = base64.b64encode(f.getvalue()).decode("utf-8")
        thumb_html_list.append(
            f'<img src="data:image/jpeg;base64,{b64_str}" class="thumb-card" title="照片 {idx+1:02d}" draggable="true">'
        )

    thumb_imgs_html = "".join(thumb_html_list)

    st.markdown(
        f"""
        <div class="fixed-top-header">
            <div style="font-size:12px; font-weight:bold; color:#555; margin-bottom:3px;">
                📷 照片縮圖 (按住可拖曳或直接於下方下拉選單選取)：
            </div>
            <div class="thumb-scroll-container">
                {thumb_imgs_html}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# -----------------------------------------------------------------------------
# 5. 下半部：清單編輯區 (01, 02, 03...)
# -----------------------------------------------------------------------------
st.subheader("📋 報告項目編輯")

photo_options = ["(未選擇)"] + [
    f"照片 {i+1:02d} - {f.name}" for i, f in enumerate(uploaded_files or [])
]

for i in range(1, st.session_state.rows_count + 1):
    row_num = f"{i:02d}"

    with st.expander(f"📌 項目 {row_num}", expanded=True):
        col_img, col_txt = st.columns([1, 1.3])

        with col_img:
            selected = st.selectbox(
                f"選擇照片 {row_num}",
                options=photo_options,
                key=f"select_{i}",
                label_visibility="collapsed",
            )

            if selected != "(未選擇)" and uploaded_files:
                img_idx = photo_options.index(selected) - 1
                st.image(
                    uploaded_files[img_idx],
                    use_container_width=True,
                )

        with col_txt:
            desc = st.text_area(
                f"敘述 {row_num}",
                placeholder="輸入現場敘述...",
                key=f"desc_{i}",
                height=90,
                label_visibility="collapsed",
            )

if st.button("➕ 新增項目 (+)"):
    st.session_state.rows_count += 1
    st.rerun()

st.divider()

# -----------------------------------------------------------------------------
# 6. 最下方：下載區 (手機 / 電腦模式)
# -----------------------------------------------------------------------------
st.subheader("📥 匯出檔案")
today_str = datetime.now().strftime("%Y%m%d")

dl_mode = st.radio(
    "下載模式選取：",
    ["📱 手機模式 (直接下載單張)", "💻 電腦模式 (打包 ZIP 壓縮檔)"],
    horizontal=True,
)

col_dl_photo, col_dl_word = st.columns(2)

# --- 下載照片 ---
with col_dl_photo:
    if dl_mode == "💻 電腦模式 (打包 ZIP 壓縮檔)":
        if st.button("📦 下載照片 ZIP"):
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
                    "⬇️ 下載 ZIP",
                    data=zip_buffer,
                    file_name=f"Photos_{today_str}.zip",
                    mime="application/zip",
                )
    else:
        st.write("點選下載單張照片：")
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

# --- 下載 Word ---
with col_dl_word:
    if st.button("📄 生成 Word 檔", use_container_width=True):
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
                    run.add_picture(img_bytes, width=Inches(3.5))

            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)

            st.download_button(
                label=f"⬇️ 下載 {today_str}_01.docx",
                data=doc_io,
                file_name=f"{today_str}_01.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
