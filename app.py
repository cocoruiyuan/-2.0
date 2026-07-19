from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import random
import zipfile

import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont


# ============================================================
# 基础设置
# ============================================================
PAGE_WIDTH = 1240
PAGE_HEIGHT = 1754

LEFT_MARGIN = 120
RIGHT_MARGIN = 100
TOP_MARGIN = 130
BOTTOM_MARGIN = 110

BASE_DIR = Path(__file__).resolve().parent
FONT_DIR = BASE_DIR / "fonts"

SUPPORTED_FILE_TYPES = ["pdf", "docx", "txt"]


@dataclass(frozen=True)
class RenderSettings:
    font_path: str
    font_size: int
    line_spacing: int
    paper_type: str
    ink_name: str
    randomness: int
    seed: int
    quality_scale: int = 2


# ============================================================
# 通用工具
# ============================================================
def clamp(value: int) -> int:
    return max(0, min(255, value))


def normalize_text(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()


def get_font_files() -> dict[str, Path]:
    """
    只扫描 fonts 文件夹。
    返回：{显示名称: 字体路径}
    """
    FONT_DIR.mkdir(exist_ok=True)

    files = sorted(
        [
            *FONT_DIR.glob("*.ttf"),
            *FONT_DIR.glob("*.otf"),
            *FONT_DIR.glob("*.TTF"),
            *FONT_DIR.glob("*.OTF"),
        ],
        key=lambda path: path.name.lower(),
    )

    return {path.name: path for path in files}


@st.cache_resource(show_spinner=False)
def load_font(font_path: str, font_size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(font_path, font_size)


# ============================================================
# 文档文字提取
# ============================================================
def extract_text_from_txt(file_bytes: bytes) -> str:
    """
    尝试用常见编码读取 TXT。
    """
    encodings = ["utf-8-sig", "utf-8", "cp1251", "windows-1251"]

    for encoding in encodings:
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue

    raise ValueError("无法识别 TXT 文件编码。请将文件另存为 UTF-8 后再上传。")


def extract_text_from_docx(file_bytes: bytes) -> str:
    """
    提取 DOCX 中的段落和表格文字。
    """
    document = Document(BytesIO(file_bytes))
    result: list[str] = []

    for element in document.iter_inner_content():
        if hasattr(element, "text") and not hasattr(element, "rows"):
            result.append(element.text)

        elif hasattr(element, "rows"):
            for row in element.rows:
                cells = [cell.text.strip() for cell in row.cells]
                result.append(" | ".join(cells))

    return "\n".join(result)


def extract_text_from_pdf(file_bytes: bytes) -> str:
    """
    提取文字型 PDF 的文字。
    扫描版 PDF 不会自动 OCR。
    """
    result: list[str] = []

    with fitz.open(stream=file_bytes, filetype="pdf") as document:
        for page_number, page in enumerate(document, start=1):
            page_text = page.get_text("text").strip()

            if page_text:
                result.append(page_text)

            if page_number < len(document):
                result.append("")

    return "\n".join(result)


def extract_uploaded_file(uploaded_file) -> str:
    """
    根据扩展名选择提取方法。
    """
    file_bytes = uploaded_file.getvalue()
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".txt":
        return extract_text_from_txt(file_bytes)

    if suffix == ".docx":
        return extract_text_from_docx(file_bytes)

    if suffix == ".pdf":
        return extract_text_from_pdf(file_bytes)

    raise ValueError("暂不支持这个文件格式。")


# ============================================================
# 纸张背景
# ============================================================
def make_texture(
    width: int,
    height: int,
    rng: random.Random,
) -> Image.Image:
    """
    使用低分辨率随机纹理放大，生成轻微纸张质感。
    """
    small_width = max(40, width // 10)
    small_height = max(40, height // 10)

    pixels = bytes(
        rng.randint(235, 255)
        for _ in range(small_width * small_height)
    )

    texture = Image.frombytes(
        "L",
        (small_width, small_height),
        pixels,
    )

    return texture.resize(
        (width, height),
        Image.Resampling.BILINEAR,
    ).convert("RGB")


def create_paper_background(
    paper_type: str,
    rng: random.Random,
) -> Image.Image:
    base = Image.new(
        "RGB",
        (PAGE_WIDTH, PAGE_HEIGHT),
        (248, 246, 240),
    )

    texture = make_texture(PAGE_WIDTH, PAGE_HEIGHT, rng)
    image = Image.blend(base, texture, 0.12)
    draw = ImageDraw.Draw(image)

    for _ in range(90):
        x1 = rng.randint(0, PAGE_WIDTH - 1)
        y1 = rng.randint(0, PAGE_HEIGHT - 1)
        x2 = x1 + rng.randint(-15, 15)
        y2 = y1 + rng.randint(-15, 15)

        draw.line(
            (x1, y1, x2, y2),
            fill=(239, 236, 229),
            width=1,
        )

    if paper_type == "横线纸":
        for y in range(TOP_MARGIN, PAGE_HEIGHT - BOTTOM_MARGIN, 62):
            draw.line(
                (80, y, PAGE_WIDTH - 80, y),
                fill=(208, 220, 235),
                width=2,
            )

        draw.line(
            (
                LEFT_MARGIN - 25,
                70,
                LEFT_MARGIN - 25,
                PAGE_HEIGHT - 70,
            ),
            fill=(235, 175, 175),
            width=2,
        )

    elif paper_type == "方格纸":
        spacing = 55

        for x in range(60, PAGE_WIDTH - 60, spacing):
            draw.line(
                (x, 60, x, PAGE_HEIGHT - 60),
                fill=(220, 227, 235),
                width=1,
            )

        for y in range(60, PAGE_HEIGHT - 60, spacing):
            draw.line(
                (60, y, PAGE_WIDTH - 60, y),
                fill=(220, 227, 235),
                width=1,
            )

    return image.convert("RGBA")


# ============================================================
# 自动换行与分页
# ============================================================
def split_long_word(
    word: str,
    draw: ImageDraw.ImageDraw,
    font: ImageFont.FreeTypeFont,
    max_width: int,
) -> list[str]:
    parts: list[str] = []
    current = ""

    for char in word:
        candidate = current + char

        if draw.textlength(candidate, font=font) <= max_width:
            current = candidate
        else:
            if current:
                parts.append(current)
            current = char

    if current:
        parts.append(current)

    return parts


def wrap_paragraph(
    paragraph: str,
    draw: ImageDraw.ImageDraw,
    font: ImageFont.FreeTypeFont,
    max_width: int,
) -> list[str]:
    if not paragraph.strip():
        return [""]

    words = paragraph.split()
    lines: list[str] = []
    current = ""

    for word in words:
        candidate = word if not current else f"{current} {word}"

        if draw.textlength(candidate, font=font) <= max_width:
            current = candidate
            continue

        if current:
            lines.append(current)
            current = ""

        if draw.textlength(word, font=font) <= max_width:
            current = word
            continue

        parts = split_long_word(
            word=word,
            draw=draw,
            font=font,
            max_width=max_width,
        )

        if parts:
            lines.extend(parts[:-1])
            current = parts[-1]

    if current:
        lines.append(current)

    return lines


def wrap_text(
    text: str,
    font: ImageFont.FreeTypeFont,
    max_width: int,
) -> list[str]:
    measuring_image = Image.new("RGB", (10, 10), "white")
    draw = ImageDraw.Draw(measuring_image)

    all_lines: list[str] = []

    for paragraph in normalize_text(text).split("\n"):
        all_lines.extend(
            wrap_paragraph(
                paragraph=paragraph,
                draw=draw,
                font=font,
                max_width=max_width,
            )
        )

    return all_lines


def paginate_lines(
    lines: list[str],
    font_size: int,
    line_spacing: int,
    randomness: int,
) -> list[list[str]]:
    normal_height = font_size + line_spacing + 8 + randomness + 2
    blank_height = normal_height // 2 + 10
    available_height = PAGE_HEIGHT - TOP_MARGIN - BOTTOM_MARGIN

    pages: list[list[str]] = []
    current_page: list[str] = []
    used_height = 0

    for line in lines:
        line_height = blank_height if not line.strip() else normal_height

        if current_page and used_height + line_height > available_height:
            pages.append(current_page)
            current_page = []
            used_height = 0

        current_page.append(line)
        used_height += line_height

    if current_page or not pages:
        pages.append(current_page)

    return pages


# ============================================================
# 手写渲染
# ============================================================
def varied_ink_color(
    ink_name: str,
    rng: random.Random,
) -> tuple[int, int, int]:
    if ink_name == "蓝色":
        base = (36, 58, 115)
        variation = 10
    else:
        base = (32, 32, 32)
        variation = 8

    return (
        clamp(base[0] + rng.randint(-variation, variation)),
        clamp(base[1] + rng.randint(-variation, variation)),
        clamp(base[2] + rng.randint(-variation, variation)),
    )


def create_pressure_mask(
    text_mask: Image.Image,
    rng: random.Random,
) -> Image.Image:
    width, height = text_mask.size
    small_width = max(2, width // 30)
    small_height = max(2, height // 12)

    pressure_pixels = bytes(
        rng.randint(218, 255)
        for _ in range(small_width * small_height)
    )

    pressure = Image.frombytes(
        "L",
        (small_width, small_height),
        pressure_pixels,
    ).resize(
        (width, height),
        Image.Resampling.BILINEAR,
    )

    return ImageChops.multiply(text_mask, pressure)


def draw_handwritten_line(
    page: Image.Image,
    text: str,
    x: int,
    y: int,
    settings: RenderSettings,
    rng: random.Random,
) -> None:
    scale = settings.quality_scale

    high_font = load_font(
        settings.font_path,
        settings.font_size * scale,
    )

    measuring_image = Image.new("L", (1, 1), 0)
    measuring_draw = ImageDraw.Draw(measuring_image)

    bbox = measuring_draw.textbbox(
        (0, 0),
        text,
        font=high_font,
    )

    padding = 14 * scale
    text_width = max(1, bbox[2] - bbox[0])
    text_height = max(1, bbox[3] - bbox[1])

    mask = Image.new(
        "L",
        (
            text_width + padding * 2,
            text_height + padding * 2,
        ),
        0,
    )

    mask_draw = ImageDraw.Draw(mask)

    mask_draw.text(
        (
            padding - bbox[0],
            padding - bbox[1],
        ),
        text,
        font=high_font,
        fill=255,
    )

    mask = mask.filter(
        ImageFilter.GaussianBlur(radius=0.12 * scale)
    )

    pressure_mask = create_pressure_mask(mask, rng)
    ink_color = varied_ink_color(settings.ink_name, rng)

    line_image = Image.new(
        "RGBA",
        mask.size,
        (
            ink_color[0],
            ink_color[1],
            ink_color[2],
            0,
        ),
    )

    line_image.putalpha(pressure_mask)

    width_change = rng.uniform(
        0.994 - settings.randomness * 0.0015,
        1.006 + settings.randomness * 0.0015,
    )

    height_change = rng.uniform(
        0.997 - settings.randomness * 0.0008,
        1.003 + settings.randomness * 0.0008,
    )

    new_width = max(
        1,
        int(line_image.width / scale * width_change),
    )

    new_height = max(
        1,
        int(line_image.height / scale * height_change),
    )

    line_image = line_image.resize(
        (new_width, new_height),
        Image.Resampling.LANCZOS,
    )

    angle = rng.uniform(
        -0.12 - settings.randomness * 0.025,
        0.12 + settings.randomness * 0.025,
    )

    line_image = line_image.rotate(
        angle,
        expand=True,
        resample=Image.Resampling.BICUBIC,
    )

    paste_x = max(0, x + rng.randint(-1, 1))
    paste_y = max(0, y + rng.randint(-1, 1))

    max_allowed_width = PAGE_WIDTH - RIGHT_MARGIN - paste_x

    if line_image.width > max_allowed_width and max_allowed_width > 10:
        ratio = max_allowed_width / line_image.width

        line_image = line_image.resize(
            (
                max_allowed_width,
                max(1, int(line_image.height * ratio)),
            ),
            Image.Resampling.LANCZOS,
        )

    page.alpha_composite(
        line_image,
        (paste_x, paste_y),
    )


def render_single_page(
    page_lines: list[str],
    settings: RenderSettings,
    page_number: int,
) -> Image.Image:
    page_seed = settings.seed + page_number * 100_003
    rng = random.Random(page_seed)

    image = create_paper_background(
        settings.paper_type,
        rng,
    )

    line_height = settings.font_size + settings.line_spacing + 8
    y = TOP_MARGIN
    previous_was_blank = True

    for line in page_lines:
        if not line.strip():
            y += line_height // 2 + 10
            previous_was_blank = True
            continue

        paragraph_indent = (
            rng.randint(18, 42)
            if previous_was_blank
            else 0
        )

        x = (
            LEFT_MARGIN
            + paragraph_indent
            + rng.randint(
                -settings.randomness * 2,
                settings.randomness * 3,
            )
        )

        y_offset = rng.randint(
            -settings.randomness,
            settings.randomness + 1,
        )

        draw_handwritten_line(
            page=image,
            text=line,
            x=x,
            y=y + y_offset,
            settings=settings,
            rng=rng,
        )

        y += line_height + rng.randint(
            -settings.randomness,
            settings.randomness + 2,
        )

        previous_was_blank = False

    return image.convert("RGB")


def render_document(
    text: str,
    settings: RenderSettings,
) -> list[Image.Image]:
    normal_font = load_font(
        settings.font_path,
        settings.font_size,
    )

    max_text_width = (
        PAGE_WIDTH
        - LEFT_MARGIN
        - RIGHT_MARGIN
        - 90
    )

    lines = wrap_text(
        text=text,
        font=normal_font,
        max_width=max_text_width,
    )

    pages_lines = paginate_lines(
        lines=lines,
        font_size=settings.font_size,
        line_spacing=settings.line_spacing,
        randomness=settings.randomness,
    )

    return [
        render_single_page(
            page_lines=page_lines,
            settings=settings,
            page_number=index,
        )
        for index, page_lines in enumerate(pages_lines)
    ]


# ============================================================
# 导出
# ============================================================
def image_to_png_bytes(image: Image.Image) -> bytes:
    buffer = BytesIO()

    image.save(
        buffer,
        format="PNG",
        optimize=True,
    )

    return buffer.getvalue()


def pages_to_zip_bytes(pages: list[Image.Image]) -> bytes:
    buffer = BytesIO()

    with zipfile.ZipFile(
        buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archive:
        for index, page in enumerate(pages, start=1):
            archive.writestr(
                f"page_{index:03d}.png",
                image_to_png_bytes(page),
            )

    return buffer.getvalue()


def pages_to_pdf_bytes(pages: list[Image.Image]) -> bytes:
    buffer = BytesIO()
    rgb_pages = [page.convert("RGB") for page in pages]

    rgb_pages[0].save(
        buffer,
        format="PDF",
        save_all=True,
        append_images=rgb_pages[1:],
        resolution=150.0,
    )

    return buffer.getvalue()


def save_result_to_session(pages: list[Image.Image]) -> None:
    st.session_state["generated_result"] = {
        "preview_pages": pages[:5],
        "page_count": len(pages),
        "first_png": image_to_png_bytes(pages[0]),
        "zip_bytes": pages_to_zip_bytes(pages),
        "pdf_bytes": pages_to_pdf_bytes(pages),
    }


# ============================================================
# Streamlit 界面
# ============================================================
st.set_page_config(
    page_title="俄语手写图片生成器",
    page_icon="✍️",
    layout="wide",
)

st.title("✍️ 俄语手写图片生成器")
st.caption("支持俄语文字、TXT、DOCX 和文字型 PDF。")

font_files = get_font_files()

if not font_files:
    st.error(
        "没有找到字体。请把支持俄语的 .ttf 或 .otf 手写字体，"
        "复制到项目的 fonts 文件夹后刷新页面。"
    )

    st.code(
        "项目文件夹\n"
        "├── app.py\n"
        "└── fonts\n"
        "    └── 你的俄语手写字体.ttf"
    )

    st.stop()


with st.sidebar:
    st.header("页面设置")

    selected_font_name = st.selectbox(
        "手写字体",
        options=list(font_files.keys()),
    )

    font_size = st.slider(
        "字体大小",
        min_value=35,
        max_value=90,
        value=58,
    )

    line_spacing = st.slider(
        "行距",
        min_value=5,
        max_value=55,
        value=22,
    )

    paper_type = st.selectbox(
        "纸张",
        ["横线纸", "白纸", "方格纸"],
    )

    ink_name = st.selectbox(
        "墨水",
        ["蓝色", "黑色"],
    )

    randomness = st.slider(
        "自然随机程度",
        min_value=0,
        max_value=6,
        value=3,
        help="建议选择 2～4。数值过高可能显得杂乱。",
    )

    seed = st.number_input(
        "随机种子",
        min_value=0,
        max_value=999_999,
        value=12_345,
        step=1,
        help="修改数字可以生成另一种随机效果。",
    )


st.subheader("1. 上传文档或输入文字")

uploaded_file = st.file_uploader(
    "上传 TXT、DOCX 或文字型 PDF",
    type=SUPPORTED_FILE_TYPES,
)

if "editor_text" not in st.session_state:
    st.session_state["editor_text"] = (
        "Привет! Это мой первый текст.\n\n"
        "Сегодня я создаю приложение, которое превращает русский текст "
        "в рукописное изображение."
    )


if uploaded_file is not None:
    file_identity = (
        uploaded_file.name,
        uploaded_file.size,
    )

    if st.session_state.get("last_uploaded_file") != file_identity:
        try:
            extracted_text = extract_uploaded_file(uploaded_file)
            extracted_text = normalize_text(extracted_text)

            if extracted_text:
                st.session_state["editor_text"] = extracted_text
                st.session_state["last_uploaded_file"] = file_identity
                st.success("文档文字提取成功，可以在下面继续修改。")
            else:
                st.warning(
                    "没有从文件中提取到文字。这个 PDF 可能是扫描版图片，"
                    "当前版本暂未加入 OCR。"
                )

        except Exception as error:
            st.error(f"读取文件失败：{error}")


text = st.text_area(
    "俄语文字",
    key="editor_text",
    height=320,
)


st.subheader("2. 生成手写文档")

generate_clicked = st.button(
    "生成手写文档",
    type="primary",
)

if generate_clicked:
    if not text.strip():
        st.warning("请先输入或上传俄语文字。")
    else:
        settings = RenderSettings(
            font_path=str(font_files[selected_font_name]),
            font_size=font_size,
            line_spacing=line_spacing,
            paper_type=paper_type,
            ink_name=ink_name,
            randomness=randomness,
            seed=int(seed),
        )

        try:
            with st.spinner("正在生成手写页面……"):
                pages = render_document(
                    text=text,
                    settings=settings,
                )

                save_result_to_session(pages)

            st.success(
                f"生成成功，共 {len(pages)} 页。"
            )

        except OSError:
            st.error(
                "字体无法打开。请确认字体文件有效，并且支持西里尔字母。"
            )

        except MemoryError:
            st.error(
                "文字太多，电脑内存不足。请减少文字数量后分批生成。"
            )

        except Exception as error:
            st.exception(error)


result = st.session_state.get("generated_result")

if result:
    st.subheader("3. 预览")

    preview_pages = result["preview_pages"]

    for index, page in enumerate(preview_pages, start=1):
        st.image(
            page,
            caption=f"第 {index} 页",
            use_container_width=True,
        )

    if result["page_count"] > len(preview_pages):
        st.info(
            f"网页只预览前 {len(preview_pages)} 页。"
            f"完整的 {result['page_count']} 页可以下载为 ZIP 或 PDF。"
        )

    st.subheader("4. 下载")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            "下载第一页 PNG",
            data=result["first_png"],
            file_name="russian_handwriting_page_001.png",
            mime="image/png",
            use_container_width=True,
        )

    with col2:
        st.download_button(
            "下载全部 PNG（ZIP）",
            data=result["zip_bytes"],
            file_name="russian_handwriting_pages.zip",
            mime="application/zip",
            use_container_width=True,
        )

    with col3:
        st.download_button(
            "下载全部页面 PDF",
            data=result["pdf_bytes"],
            file_name="russian_handwriting_document.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
