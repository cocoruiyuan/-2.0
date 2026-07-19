from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import random
import zipfile

import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from PIL import (
    Image,
    ImageChops,
    ImageDraw,
    ImageFilter,
    ImageFont,
)


# ============================================================
# 页面设置
# ============================================================
PAGE_WIDTH = 1240
PAGE_HEIGHT = 1754

LEFT_MARGIN = 110
RIGHT_MARGIN = 80
TOP_MARGIN = 125
BOTTOM_MARGIN = 100

BASE_DIR = Path(__file__).resolve().parent

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
    slant: float
    word_spacing: int
    quality_scale: int = 3

    @property
    def rule_spacing(self) -> int:
        """
        横线间距与文字行距保持一致。
        """
        return max(
            self.font_size + 8,
            self.font_size + self.line_spacing,
        )


# ============================================================
# 通用工具
# ============================================================
def clamp(value: int) -> int:
    return max(0, min(255, value))


def normalize_text(text: str) -> str:
    return (
        text.replace("\r\n", "\n")
        .replace("\r", "\n")
        .strip()
    )


def get_font_files() -> dict[str, Path]:
    """
    搜索整个项目中的 TTF 和 OTF 字体。

    字体既可以放在 fonts 文件夹里，
    也可以暂时与 app.py 并列。
    """
    discovered: list[Path] = []

    for pattern in ("*.ttf", "*.otf", "*.TTF", "*.OTF"):
        discovered.extend(BASE_DIR.rglob(pattern))

    unique_files = {
        path.resolve()
        for path in discovered
        if path.is_file()
        and ".venv" not in path.parts
        and "site-packages" not in path.parts
    }

    files = sorted(
        unique_files,
        key=lambda path: str(path).lower(),
    )

    return {
        str(path.relative_to(BASE_DIR)): path
        for path in files
    }


@st.cache_resource(show_spinner=False)
def load_font(
    font_path: str,
    font_size: int,
) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(
        font_path,
        font_size,
    )


# ============================================================
# 文档文字提取
# ============================================================
def extract_text_from_txt(file_bytes: bytes) -> str:
    encodings = [
        "utf-8-sig",
        "utf-8",
        "cp1251",
        "windows-1251",
    ]

    for encoding in encodings:
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue

    raise ValueError(
        "无法识别 TXT 文件编码。请将文件另存为 UTF-8。"
    )


def extract_text_from_docx(file_bytes: bytes) -> str:
    document = Document(BytesIO(file_bytes))
    result: list[str] = []

    # 读取普通段落
    for paragraph in document.paragraphs:
        result.append(paragraph.text)

    # 读取表格
    for table in document.tables:
        result.append("")

        for row in table.rows:
            cells = [
                cell.text.strip()
                for cell in row.cells
            ]
            result.append(" | ".join(cells))

    return "\n".join(result)


def extract_text_from_pdf(file_bytes: bytes) -> str:
    """
    只提取文字型 PDF。
    扫描版 PDF 暂时没有 OCR。
    """
    result: list[str] = []

    with fitz.open(
        stream=file_bytes,
        filetype="pdf",
    ) as document:
        for page_index, page in enumerate(document):
            page_text = page.get_text("text").strip()

            if page_text:
                result.append(page_text)

            if page_index < len(document) - 1:
                result.append("")

    return "\n".join(result)


def extract_uploaded_file(uploaded_file) -> str:
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
# 作业本纸张
# ============================================================
def add_light_paper_texture(
    image: Image.Image,
    rng: random.Random,
) -> None:
    """
    添加非常轻微的纸张颗粒。
    参考图的纸比较干净，所以纹理不能太强。
    """
    draw = ImageDraw.Draw(image)

    for _ in range(1800):
        x = rng.randint(0, PAGE_WIDTH - 1)
        y = rng.randint(0, PAGE_HEIGHT - 1)
        shade = rng.randint(-3, 3)

        draw.point(
            (x, y),
            fill=(
                clamp(247 + shade),
                clamp(244 + shade),
                clamp(235 + shade),
            ),
        )

    for _ in range(55):
        x1 = rng.randint(0, PAGE_WIDTH - 1)
        y1 = rng.randint(0, PAGE_HEIGHT - 1)
        x2 = x1 + rng.randint(-13, 13)
        y2 = y1 + rng.randint(-5, 5)

        draw.line(
            (x1, y1, x2, y2),
            fill=(240, 237, 229),
            width=1,
        )


def create_paper_background(
    settings: RenderSettings,
    rng: random.Random,
) -> Image.Image:
    """
    创建接近参考图片的米白色俄语作业本纸张。
    """
    image = Image.new(
        "RGB",
        (PAGE_WIDTH, PAGE_HEIGHT),
        (247, 244, 235),
    )

    add_light_paper_texture(image, rng)
    draw = ImageDraw.Draw(image)

    if settings.paper_type == "横线纸":
        # 顶部双横线
        draw.line(
            (55, 65, PAGE_WIDTH - 55, 65),
            fill=(174, 184, 210),
            width=2,
        )
        draw.line(
            (55, 78, PAGE_WIDTH - 55, 78),
            fill=(192, 200, 220),
            width=1,
        )

        # 作业本横线
        for y in range(
            TOP_MARGIN,
            PAGE_HEIGHT - BOTTOM_MARGIN,
            settings.rule_spacing,
        ):
            draw.line(
                (55, y, PAGE_WIDTH - 55, y),
                fill=(167, 184, 218),
                width=2,
            )

        # 左侧红色竖线
        red_line_x = LEFT_MARGIN + 48

        draw.line(
            (
                red_line_x,
                70,
                red_line_x,
                PAGE_HEIGHT - 70,
            ),
            fill=(204, 137, 140),
            width=2,
        )

    elif settings.paper_type == "方格纸":
        spacing = settings.rule_spacing

        for x in range(55, PAGE_WIDTH - 55, spacing):
            draw.line(
                (x, 55, x, PAGE_HEIGHT - 55),
                fill=(205, 214, 230),
                width=1,
            )

        for y in range(55, PAGE_HEIGHT - 55, spacing):
            draw.line(
                (55, y, PAGE_WIDTH - 55, y),
                fill=(205, 214, 230),
                width=1,
            )

    return image.convert("RGBA")


# ============================================================
# 自动换行和分页
# ============================================================
def text_width(
    text: str,
    font: ImageFont.FreeTypeFont,
) -> float:
    measuring_image = Image.new(
        "RGB",
        (10, 10),
        "white",
    )
    draw = ImageDraw.Draw(measuring_image)
    return draw.textlength(text, font=font)


def split_long_word(
    word: str,
    font: ImageFont.FreeTypeFont,
    max_width: int,
) -> list[str]:
    parts: list[str] = []
    current = ""

    for char in word:
        candidate = current + char

        if text_width(candidate, font) <= max_width:
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
    font: ImageFont.FreeTypeFont,
    max_width: int,
) -> list[str]:
    if not paragraph.strip():
        return [""]

    words = paragraph.split()
    lines: list[str] = []
    current = ""

    for word in words:
        candidate = (
            word
            if not current
            else f"{current} {word}"
        )

        if text_width(candidate, font) <= max_width:
            current = candidate
            continue

        if current:
            lines.append(current)
            current = ""

        if text_width(word, font) <= max_width:
            current = word
            continue

        parts = split_long_word(
            word=word,
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
    all_lines: list[str] = []

    for paragraph in normalize_text(text).split("\n"):
        all_lines.extend(
            wrap_paragraph(
                paragraph=paragraph,
                font=font,
                max_width=max_width,
            )
        )

    return all_lines


def paginate_lines(
    lines: list[str],
    settings: RenderSettings,
) -> list[list[str]]:
    available_height = (
        PAGE_HEIGHT
        - TOP_MARGIN
        - BOTTOM_MARGIN
    )

    maximum_rows = max(
        1,
        available_height // settings.rule_spacing,
    )

    pages: list[list[str]] = []
    current_page: list[str] = []

    for line in lines:
        if len(current_page) >= maximum_rows:
            pages.append(current_page)
            current_page = []

        current_page.append(line)

    if current_page or not pages:
        pages.append(current_page)

    return pages


# ============================================================
# 手写字渲染
# ============================================================
def varied_ink_color(
    ink_name: str,
    rng: random.Random,
) -> tuple[int, int, int]:
    """
    参考图墨水比较稳定，所以颜色波动很小。
    """
    if ink_name == "蓝色":
        base = (42, 62, 112)
        variation = 4
    else:
        base = (39, 39, 38)
        variation = 3

    return (
        clamp(base[0] + rng.randint(-variation, variation)),
        clamp(base[1] + rng.randint(-variation, variation)),
        clamp(base[2] + rng.randint(-variation, variation)),
    )


def create_pressure_mask(
    text_mask: Image.Image,
    rng: random.Random,
) -> Image.Image:
    """
    模拟圆珠笔书写压力，但保持变化很小。
    """
    width, height = text_mask.size

    small_width = max(2, width // 35)
    small_height = max(2, height // 14)

    pressure_pixels = bytes(
        rng.randint(238, 255)
        for _ in range(
            small_width * small_height
        )
    )

    pressure = Image.frombytes(
        "L",
        (small_width, small_height),
        pressure_pixels,
    ).resize(
        (width, height),
        Image.Resampling.BILINEAR,
    )

    return ImageChops.multiply(
        text_mask,
        pressure,
    )


def shear_right(
    image: Image.Image,
    shear: float,
) -> Image.Image:
    """
    让字形向右倾斜。

    数值建议在 0.06～0.16 之间。
    """
    if shear <= 0:
        return image

    extra_width = int(
        abs(shear) * image.height
    )

    new_width = image.width + extra_width

    # 让字的上部向右倾斜
    offset = -shear * image.height

    return image.transform(
        (new_width, image.height),
        Image.Transform.AFFINE,
        (
            1,
            shear,
            offset,
            0,
            1,
            0,
        ),
        resample=Image.Resampling.BICUBIC,
    )


def render_word_image(
    word: str,
    settings: RenderSettings,
    rng: random.Random,
) -> tuple[Image.Image, int]:
    """
    单独渲染一个单词。

    单词内部仍然保留字体连笔，
    不同单词之间有细微大小和基线变化。
    """
    scale = settings.quality_scale

    high_font = load_font(
        settings.font_path,
        settings.font_size * scale,
    )

    ascent, descent = high_font.getmetrics()
    padding = 12 * scale

    measuring_image = Image.new(
        "L",
        (1, 1),
        0,
    )
    measuring_draw = ImageDraw.Draw(
        measuring_image
    )

    bbox = measuring_draw.textbbox(
        (0, 0),
        word,
        font=high_font,
        anchor="ls",
    )

    text_width_value = max(
        1,
        bbox[2] - bbox[0],
    )

    canvas_width = (
        text_width_value
        + padding * 2
        + int(settings.slant * (ascent + descent))
    )

    canvas_height = (
        ascent
        + descent
        + padding * 2
    )

    baseline_y = padding + ascent

    mask = Image.new(
        "L",
        (
            max(1, canvas_width),
            max(1, canvas_height),
        ),
        0,
    )

    mask_draw = ImageDraw.Draw(mask)

    mask_draw.text(
        (
            padding - bbox[0],
            baseline_y,
        ),
        word,
        font=high_font,
        fill=255,
        anchor="ls",
    )

    # 只做极轻的边缘柔化
    mask = mask.filter(
        ImageFilter.GaussianBlur(
            radius=0.045 * scale
        )
    )

    pressure_mask = create_pressure_mask(
        mask,
        rng,
    )

    ink_color = varied_ink_color(
        settings.ink_name,
        rng,
    )

    word_image = Image.new(
        "RGBA",
        mask.size,
        (
            ink_color[0],
            ink_color[1],
            ink_color[2],
            0,
        ),
    )

    word_image.putalpha(pressure_mask)

    # 很小的宽高变化
    width_change = rng.uniform(
        0.997 - settings.randomness * 0.0007,
        1.003 + settings.randomness * 0.0007,
    )

    height_change = rng.uniform(
        0.999 - settings.randomness * 0.0004,
        1.001 + settings.randomness * 0.0004,
    )

    resized_width = max(
        1,
        int(
            word_image.width
            / scale
            * width_change
        ),
    )

    resized_height = max(
        1,
        int(
            word_image.height
            / scale
            * height_change
        ),
    )

    baseline_after_resize = int(
        baseline_y
        / scale
        * height_change
    )

    word_image = word_image.resize(
        (
            resized_width,
            resized_height,
        ),
        Image.Resampling.LANCZOS,
    )

    word_slant = max(
        0.0,
        settings.slant
        + rng.uniform(-0.012, 0.012),
    )

    word_image = shear_right(
        word_image,
        word_slant,
    )

    return (
        word_image,
        baseline_after_resize,
    )


def draw_handwritten_line(
    page: Image.Image,
    text: str,
    baseline_y: int,
    x_start: int,
    settings: RenderSettings,
    rng: random.Random,
) -> None:
    """
    按单词绘制整行。

    这样比整行一次性拉伸更自然，
    又不会破坏单词内部的俄语连笔。
    """
    words = text.split()

    if not words:
        return

    normal_font = load_font(
        settings.font_path,
        settings.font_size,
    )

    base_space = int(
        text_width(" ", normal_font)
    )

    current_x = x_start

    for word_index, word in enumerate(words):
        word_image, word_baseline = (
            render_word_image(
                word=word,
                settings=settings,
                rng=rng,
            )
        )

        word_y_shift = rng.randint(
            -max(1, settings.randomness),
            max(1, settings.randomness),
        )

        paste_y = (
            baseline_y
            - word_baseline
            + word_y_shift
        )

        max_width = (
            PAGE_WIDTH
            - RIGHT_MARGIN
            - current_x
        )

        if max_width <= 4:
            break

        if word_image.width > max_width:
            ratio = max_width / word_image.width

            word_image = word_image.resize(
                (
                    max_width,
                    max(
                        1,
                        int(
                            word_image.height
                            * ratio
                        ),
                    ),
                ),
                Image.Resampling.LANCZOS,
            )

        page.alpha_composite(
            word_image,
            (
                max(0, current_x),
                max(0, paste_y),
            ),
        )

        current_x += (
            word_image.width
            + base_space
            + settings.word_spacing
            + rng.randint(-1, 2)
        )

        if word_index == len(words) - 1:
            break


def render_single_page(
    page_lines: list[str],
    settings: RenderSettings,
    page_number: int,
) -> Image.Image:
    page_seed = (
        settings.seed
        + page_number * 100_003
    )

    rng = random.Random(page_seed)

    image = create_paper_background(
        settings=settings,
        rng=rng,
    )

    red_line_x = LEFT_MARGIN + 48
    text_start_x = red_line_x + 20

    previous_was_blank = True

    for row_index, line in enumerate(page_lines):
        rule_y = (
            TOP_MARGIN
            + row_index * settings.rule_spacing
        )

        if not line.strip():
            previous_was_blank = True
            continue

        paragraph_indent = (
            rng.randint(8, 18)
            if previous_was_blank
            else 0
        )

        x_start = (
            text_start_x
            + paragraph_indent
            + rng.randint(-1, 2)
        )

        # 让字自然坐在横线上
        baseline_y = (
            rule_y
            - 2
            + rng.randint(-1, 1)
        )

        draw_handwritten_line(
            page=image,
            text=line,
            baseline_y=baseline_y,
            x_start=x_start,
            settings=settings,
            rng=rng,
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

    red_line_x = LEFT_MARGIN + 48
    text_start_x = red_line_x + 20

    max_text_width = (
        PAGE_WIDTH
        - RIGHT_MARGIN
        - text_start_x
        - 28
    )

    lines = wrap_text(
        text=text,
        font=normal_font,
        max_width=max_text_width,
    )

    pages_lines = paginate_lines(
        lines=lines,
        settings=settings,
    )

    return [
        render_single_page(
            page_lines=page_lines,
            settings=settings,
            page_number=index,
        )
        for index, page_lines in enumerate(
            pages_lines
        )
    ]


# ============================================================
# 导出
# ============================================================
def image_to_png_bytes(
    image: Image.Image,
) -> bytes:
    buffer = BytesIO()

    image.save(
        buffer,
        format="PNG",
        optimize=True,
    )

    return buffer.getvalue()


def pages_to_zip_bytes(
    pages: list[Image.Image],
) -> bytes:
    buffer = BytesIO()

    with zipfile.ZipFile(
        buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archive:
        for index, page in enumerate(
            pages,
            start=1,
        ):
            archive.writestr(
                f"page_{index:03d}.png",
                image_to_png_bytes(page),
            )

    return buffer.getvalue()


def pages_to_pdf_bytes(
    pages: list[Image.Image],
) -> bytes:
    buffer = BytesIO()

    rgb_pages = [
        page.convert("RGB")
        for page in pages
    ]

    rgb_pages[0].save(
        buffer,
        format="PDF",
        save_all=True,
        append_images=rgb_pages[1:],
        resolution=150.0,
    )

    return buffer.getvalue()


def save_result_to_session(
    pages: list[Image.Image],
) -> None:
    st.session_state["generated_result"] = {
        "preview_pages": pages[:5],
        "page_count": len(pages),
        "first_png": image_to_png_bytes(
            pages[0]
        ),
        "zip_bytes": pages_to_zip_bytes(
            pages
        ),
        "pdf_bytes": pages_to_pdf_bytes(
            pages
        ),
    }


# ============================================================
# Streamlit 页面
# ============================================================
st.set_page_config(
    page_title="俄语手写图片生成器",
    page_icon="✍️",
    layout="wide",
)

st.title("✍️ 俄语手写图片生成器")

st.caption(
    "作业本连笔风格：支持俄语文字、TXT、DOCX 和文字型 PDF。"
)

font_files = get_font_files()

if not font_files:
    st.error(
        "没有找到字体。请上传一个支持俄语的 "
        ".ttf 或 .otf 手写字体。"
    )

    st.code(
        "推荐结构：\n"
        "app.py\n"
        "fonts/\n"
        "    BadScript-Regular.ttf"
    )

    st.stop()


with st.sidebar:
    st.header("手写设置")

    selected_font_name = st.selectbox(
        "手写字体",
        options=list(font_files.keys()),
        help=(
            "想接近参考图，推荐使用 Bad Script，"
            "其次是 Marck Script。"
        ),
    )

    font_size = st.slider(
        "字体大小",
        min_value=36,
        max_value=72,
        value=48,
    )

    line_spacing = st.slider(
        "横线间距",
        min_value=4,
        max_value=24,
        value=10,
    )

    paper_type = st.selectbox(
        "纸张",
        [
            "横线纸",
            "白纸",
            "方格纸",
        ],
    )

    ink_name = st.selectbox(
        "墨水",
        [
            "黑色",
            "蓝色",
        ],
    )

    randomness = st.slider(
        "自然随机程度",
        min_value=0,
        max_value=4,
        value=1,
        help=(
            "参考图的字比较整齐，"
            "建议保持在 1～2。"
        ),
    )

    slant = st.slider(
        "右倾程度",
        min_value=0.00,
        max_value=0.20,
        value=0.10,
        step=0.01,
        help=(
            "Bad Script 本身已经有倾斜，"
            "可以调到 0.05～0.10。"
        ),
    )

    word_spacing = st.slider(
        "单词间距",
        min_value=-4,
        max_value=12,
        value=0,
    )

    seed = st.number_input(
        "随机种子",
        min_value=0,
        max_value=999_999,
        value=12_345,
        step=1,
        help=(
            "修改这个数字可以生成另一种"
            "轻微不同的手写排列。"
        ),
    )


st.subheader("1. 上传文档或输入文字")

uploaded_file = st.file_uploader(
    "上传 TXT、DOCX 或文字型 PDF",
    type=SUPPORTED_FILE_TYPES,
)

if "editor_text" not in st.session_state:
    st.session_state["editor_text"] = (
        "Это моя семья.\n"
        "Это моя семья.\n"
        "Зовут меня Мин. Я ученик сельской школы.\n"
        "Вот мои родители. Это мой папа и моя мама.\n"
        "Мы любим нашу маленькую семью."
    )


if uploaded_file is not None:
    file_identity = (
        uploaded_file.name,
        uploaded_file.size,
    )

    if (
        st.session_state.get(
            "last_uploaded_file"
        )
        != file_identity
    ):
        try:
            extracted_text = (
                extract_uploaded_file(
                    uploaded_file
                )
            )

            extracted_text = normalize_text(
                extracted_text
            )

            if extracted_text:
                st.session_state[
                    "editor_text"
                ] = extracted_text

                st.session_state[
                    "last_uploaded_file"
                ] = file_identity

                st.success(
                    "文字提取成功，"
                    "可以在下面继续修改。"
                )
            else:
                st.warning(
                    "没有提取到文字。这个 PDF "
                    "可能是扫描版图片，"
                    "当前版本暂未加入 OCR。"
                )

        except Exception as error:
            st.error(
                f"读取文件失败：{error}"
            )


text = st.text_area(
    "俄语文字",
    key="editor_text",
    height=330,
)


st.subheader("2. 生成手写文档")

generate_clicked = st.button(
    "生成手写文档",
    type="primary",
    use_container_width=True,
)

if generate_clicked:
    if not text.strip():
        st.warning(
            "请先输入或上传俄语文字。"
        )
    else:
        settings = RenderSettings(
            font_path=str(
                font_files[
                    selected_font_name
                ]
            ),
            font_size=font_size,
            line_spacing=line_spacing,
            paper_type=paper_type,
            ink_name=ink_name,
            randomness=randomness,
            seed=int(seed),
            slant=float(slant),
            word_spacing=word_spacing,
        )

        try:
            with st.spinner(
                "正在生成作业本手写页面……"
            ):
                pages = render_document(
                    text=text,
                    settings=settings,
                )

                save_result_to_session(
                    pages
                )

            st.success(
                f"生成成功，共 {len(pages)} 页。"
            )

        except OSError:
            st.error(
                "字体无法打开。请确认字体文件"
                "有效，并且支持西里尔字母。"
            )

        except MemoryError:
            st.error(
                "文字太多，电脑内存不足。"
                "请减少文字后分批生成。"
            )

        except Exception as error:
            st.exception(error)


result = st.session_state.get(
    "generated_result"
)

if result:
    st.subheader("3. 预览")

    preview_pages = result[
        "preview_pages"
    ]

    for index, page in enumerate(
        preview_pages,
        start=1,
    ):
        st.image(
            page,
            caption=f"第 {index} 页",
            use_container_width=True,
        )

    if (
        result["page_count"]
        > len(preview_pages)
    ):
        st.info(
            f"网页只预览前 "
            f"{len(preview_pages)} 页。"
            f"全部 {result['page_count']} 页"
            f"可以下载为 ZIP 或 PDF。"
        )

    st.subheader("4. 下载")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            "下载第一页 PNG",
            data=result["first_png"],
            file_name=(
                "russian_handwriting_"
                "page_001.png"
            ),
            mime="image/png",
            use_container_width=True,
        )

    with col2:
        st.download_button(
            "下载全部 PNG（ZIP）",
            data=result["zip_bytes"],
            file_name=(
                "russian_handwriting_"
                "pages.zip"
            ),
            mime="application/zip",
            use_container_width=True,
        )

    with col3:
        st.download_button(
            "下载全部页面 PDF",
            data=result["pdf_bytes"],
            file_name=(
                "russian_handwriting_"
                "document.pdf"
            ),
            mime="application/pdf",
            use_container_width=True,
        )
