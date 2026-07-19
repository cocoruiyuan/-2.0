from __future__ import annotations

from dataclasses import dataclass, replace
from io import BytesIO
from pathlib import Path
import math
import random
import zipfile

import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from PIL import (
    Image,
    ImageChops,
    ImageDraw,
    ImageEnhance,
    ImageFilter,
    ImageFont,
    ImageOps,
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
    paper_age: float
    photo_effect: str
    ink_name: str
    pen_style: str
    texture_strength: float
    randomness: int
    seed: int
    slant: float
    word_spacing: int
    connection_strength: float
    baseline_wave: float
    flourish_level: int
    correction_level: int
    teacher_marks: int
    header_enabled: bool
    header_date: str
    header_lesson: str
    show_page_number: bool
    quality_scale: int = 3

    @property
    def rule_spacing(self) -> int:
        return max(
            self.font_size + 8,
            self.font_size + self.line_spacing,
        )



PRESETS: dict[str, dict[str, object]] = {
    "参考图：蓝色圆珠笔": {
        "font_size": 49,
        "line_spacing": 10,
        "paper_type": "俄语作业本",
        "paper_age": 0.12,
        "photo_effect": "无",
        "ink_name": "蓝色",
        "pen_style": "圆珠笔",
        "texture_strength": 0.38,
        "randomness": 1,
        "slant": 0.10,
        "connection_strength": 0.82,
        "word_spacing": -1,
        "baseline_wave": 0.8,
        "flourish_level": 1,
        "correction_level": 1,
        "teacher_marks": 0,
    },
    "整齐钢笔作业": {
        "font_size": 48,
        "line_spacing": 10,
        "paper_type": "俄语作业本",
        "paper_age": 0.05,
        "photo_effect": "无",
        "ink_name": "蓝色",
        "pen_style": "钢笔",
        "texture_strength": 0.24,
        "randomness": 1,
        "slant": 0.11,
        "connection_strength": 0.86,
        "word_spacing": -1,
        "baseline_wave": 0.5,
        "flourish_level": 2,
        "correction_level": 0,
        "teacher_marks": 0,
    },
    "黑色中性笔作业": {
        "font_size": 48,
        "line_spacing": 10,
        "paper_type": "普通横线本",
        "paper_age": 0.08,
        "photo_effect": "无",
        "ink_name": "黑色",
        "pen_style": "中性笔",
        "texture_strength": 0.22,
        "randomness": 1,
        "slant": 0.08,
        "connection_strength": 0.76,
        "word_spacing": 0,
        "baseline_wave": 0.7,
        "flourish_level": 1,
        "correction_level": 1,
        "teacher_marks": 0,
    },
    "铅笔课堂草稿": {
        "font_size": 47,
        "line_spacing": 9,
        "paper_type": "方格本",
        "paper_age": 0.18,
        "photo_effect": "扫描件",
        "ink_name": "黑色",
        "pen_style": "铅笔",
        "texture_strength": 0.68,
        "randomness": 2,
        "slant": 0.06,
        "connection_strength": 0.64,
        "word_spacing": 0,
        "baseline_wave": 1.4,
        "flourish_level": 0,
        "correction_level": 2,
        "teacher_marks": 0,
    },
    "凌乱学生课堂笔记": {
        "font_size": 50,
        "line_spacing": 10,
        "paper_type": "普通横线本",
        "paper_age": 0.22,
        "photo_effect": "手机拍照",
        "ink_name": "蓝色",
        "pen_style": "圆珠笔",
        "texture_strength": 0.48,
        "randomness": 3,
        "slant": 0.12,
        "connection_strength": 0.78,
        "word_spacing": -2,
        "baseline_wave": 2.0,
        "flourish_level": 2,
        "correction_level": 2,
        "teacher_marks": 0,
    },
    "老师批改后的作业": {
        "font_size": 48,
        "line_spacing": 10,
        "paper_type": "俄语作业本",
        "paper_age": 0.14,
        "photo_effect": "手机拍照",
        "ink_name": "蓝色",
        "pen_style": "圆珠笔",
        "texture_strength": 0.38,
        "randomness": 1,
        "slant": 0.10,
        "connection_strength": 0.82,
        "word_spacing": -1,
        "baseline_wave": 0.8,
        "flourish_level": 1,
        "correction_level": 1,
        "teacher_marks": 2,
    },
}


def option_index(
    options: list,
    value,
) -> int:
    try:
        return options.index(value)
    except ValueError:
        return 0


# ============================================================
# 通用工具
# ============================================================
def clamp(value: int) -> int:
    return max(0, min(255, value))


def normalize_text(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()


def get_font_files() -> dict[str, Path]:
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
    return ImageFont.truetype(font_path, font_size)


def draw_text_with_cursive_features(
    draw: ImageDraw.ImageDraw,
    position: tuple[int | float, int | float],
    text: str,
    font: ImageFont.FreeTypeFont,
    fill,
    anchor: str | None = None,
) -> None:
    """
    尝试启用字体中的连字和上下文替换功能。

    如果服务器上的 Pillow 没有 RAQM 支持，
    会自动退回普通绘制，不会导致程序崩溃。
    """
    kwargs = {
        "font": font,
        "fill": fill,
    }

    if anchor is not None:
        kwargs["anchor"] = anchor

    try:
        draw.text(
            position,
            text,
            features=["calt", "liga", "clig"],
            **kwargs,
        )
    except Exception:
        draw.text(
            position,
            text,
            **kwargs,
        )


def get_pen_profile(
    settings: RenderSettings,
) -> dict[str, float | tuple[int, int, int]]:
    """
    不同笔迹工具的基础参数。
    """
    if settings.pen_style == "钢笔":
        if settings.ink_name == "蓝色":
            base_color = (34, 57, 116)
        else:
            base_color = (28, 28, 30)

        return {
            "base_color": base_color,
            "variation": 4,
            "pressure_low": 236,
            "pressure_high": 255,
            "alpha_low": 235,
            "alpha_high": 250,
            "blur": 0.030,
            "grain": 0.10 + settings.texture_strength * 0.12,
            "skipiness": 0.02,
            "feather": 0.12 + settings.texture_strength * 0.10,
        }

    if settings.pen_style == "圆珠笔":
        if settings.ink_name == "蓝色":
            base_color = (48, 76, 140)
        else:
            base_color = (40, 40, 42)

        return {
            "base_color": base_color,
            "variation": 5,
            "pressure_low": 228,
            "pressure_high": 255,
            "alpha_low": 224,
            "alpha_high": 246,
            "blur": 0.028,
            "grain": 0.14 + settings.texture_strength * 0.12,
            "skipiness": 0.06 + settings.texture_strength * 0.05,
            "feather": 0.06,
        }

    if settings.pen_style == "铅笔":
        return {
            "base_color": (88, 88, 92),
            "variation": 8,
            "pressure_low": 190,
            "pressure_high": 245,
            "alpha_low": 160,
            "alpha_high": 220,
            "blur": 0.020,
            "grain": 0.34 + settings.texture_strength * 0.26,
            "skipiness": 0.16 + settings.texture_strength * 0.12,
            "feather": 0.02,
        }

    # 默认：中性笔
    if settings.ink_name == "蓝色":
        base_color = (36, 61, 122)
    else:
        base_color = (35, 35, 36)

    return {
        "base_color": base_color,
        "variation": 3,
        "pressure_low": 238,
        "pressure_high": 255,
        "alpha_low": 236,
        "alpha_high": 250,
        "blur": 0.026,
        "grain": 0.08 + settings.texture_strength * 0.08,
        "skipiness": 0.02,
        "feather": 0.04,
    }


def varied_ink_color(
    settings: RenderSettings,
    rng: random.Random,
) -> tuple[int, int, int]:
    profile = get_pen_profile(settings)
    base = profile["base_color"]
    variation = int(profile["variation"])

    return (
        clamp(base[0] + rng.randint(-variation, variation)),
        clamp(base[1] + rng.randint(-variation, variation)),
        clamp(base[2] + rng.randint(-variation, variation)),
    )


def create_pressure_mask(
    text_mask: Image.Image,
    settings: RenderSettings,
    rng: random.Random,
) -> Image.Image:
    """
    生成基础笔压透明度。
    """
    profile = get_pen_profile(settings)

    width, height = text_mask.size
    small_width = max(2, width // 35)
    small_height = max(2, height // 14)

    pressure_pixels = bytes(
        rng.randint(
            int(profile["pressure_low"]),
            int(profile["pressure_high"]),
        )
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


def create_stroke_texture_mask(
    size: tuple[int, int],
    settings: RenderSettings,
    rng: random.Random,
) -> Image.Image:
    """
    为笔迹添加纹理：
    - 钢笔：比较顺滑，只带轻微墨水不均
    - 圆珠笔：有轻微断墨和粗糙感
    - 中性笔：较稳定
    - 铅笔：明显颗粒感和石墨摩擦痕迹
    """
    width, height = size
    profile = get_pen_profile(settings)
    grain = float(profile["grain"])
    skipiness = float(profile["skipiness"])

    small_width = max(2, width // 10)
    small_height = max(2, height // 6)

    pixels = bytearray()

    for _ in range(small_width * small_height):
        base_value = 255 - int(rng.random() * 255 * grain * 0.55)
        if rng.random() < skipiness:
            base_value -= rng.randint(12, 55)
        pixels.append(clamp(base_value))

    texture = Image.frombytes(
        "L",
        (small_width, small_height),
        bytes(pixels),
    ).resize(
        (width, height),
        Image.Resampling.BILINEAR,
    )

    if settings.pen_style == "铅笔":
        # 再叠加一层水平方向的石墨纹理
        graphite = Image.new("L", (width, height), 255)
        gdraw = ImageDraw.Draw(graphite)

        line_count = max(6, height // 4)
        for _ in range(line_count):
            y = rng.randint(0, height - 1)
            gdraw.line(
                (
                    0,
                    y,
                    width,
                    y + rng.randint(-1, 1),
                ),
                fill=clamp(232 - rng.randint(0, 42)),
                width=1,
            )

        texture = ImageChops.multiply(texture, graphite)

    return texture


def apply_pen_texture(
    text_mask: Image.Image,
    settings: RenderSettings,
    rng: random.Random,
) -> tuple[Image.Image, tuple[int, int, int], int]:
    """
    根据书写工具把基础文字 mask 变成更像真实笔迹的 alpha。
    """
    profile = get_pen_profile(settings)

    pressure_mask = create_pressure_mask(
        text_mask,
        settings,
        rng,
    )

    texture_mask = create_stroke_texture_mask(
        pressure_mask.size,
        settings,
        rng,
    )

    alpha = ImageChops.multiply(
        pressure_mask,
        texture_mask,
    )

    if settings.pen_style == "钢笔" and float(profile["feather"]) > 0:
        feather = alpha.filter(
            ImageFilter.GaussianBlur(radius=1.0)
        ).point(
            lambda pixel: int(pixel * float(profile["feather"]))
        )
        alpha = ImageChops.lighter(alpha, feather)

    color = varied_ink_color(settings, rng)
    line_alpha = rng.randint(
        int(profile["alpha_low"]),
        int(profile["alpha_high"]),
    )

    return alpha, color, line_alpha


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

    raise ValueError("无法识别 TXT 文件编码。请将文件另存为 UTF-8。")


def extract_text_from_docx(file_bytes: bytes) -> str:
    document = Document(BytesIO(file_bytes))
    result: list[str] = []

    for paragraph in document.paragraphs:
        result.append(paragraph.text)

    for table in document.tables:
        result.append("")
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            result.append(" | ".join(cells))

    return "\n".join(result)


def extract_text_from_pdf(file_bytes: bytes) -> str:
    result: list[str] = []

    with fitz.open(stream=file_bytes, filetype="pdf") as document:
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
    settings: RenderSettings,
    rng: random.Random,
) -> None:
    """
    添加纸纤维、轻微旧化和不均匀底色。
    """
    draw = ImageDraw.Draw(image)
    age = settings.paper_age

    grain_count = 1500 + int(age * 4200)

    for _ in range(grain_count):
        x = rng.randint(0, PAGE_WIDTH - 1)
        y = rng.randint(0, PAGE_HEIGHT - 1)
        spread = 3 + int(age * 8)
        shade = rng.randint(-spread, spread)

        draw.point(
            (x, y),
            fill=(
                clamp(248 + shade),
                clamp(246 + shade),
                clamp(239 + shade - int(age * 4)),
            ),
        )

    fiber_count = 40 + int(age * 110)

    for _ in range(fiber_count):
        x1 = rng.randint(0, PAGE_WIDTH - 1)
        y1 = rng.randint(0, PAGE_HEIGHT - 1)
        x2 = x1 + rng.randint(-18, 18)
        y2 = y1 + rng.randint(-5, 5)

        draw.line(
            (x1, y1, x2, y2),
            fill=(
                240 - int(age * 10),
                237 - int(age * 9),
                230 - int(age * 7),
            ),
            width=1,
        )

    if age > 0.15:
        stain_layer = Image.new(
            "RGBA",
            image.size,
            (0, 0, 0, 0),
        )
        stain_draw = ImageDraw.Draw(stain_layer)

        stain_count = 1 + int(age * 6)

        for _ in range(stain_count):
            cx = rng.randint(80, PAGE_WIDTH - 80)
            cy = rng.randint(80, PAGE_HEIGHT - 80)
            rx = rng.randint(30, 110)
            ry = rng.randint(20, 80)

            stain_draw.ellipse(
                (cx - rx, cy - ry, cx + rx, cy + ry),
                fill=(146, 116, 75, rng.randint(3, 10)),
            )

        stain_layer = stain_layer.filter(
            ImageFilter.GaussianBlur(radius=22)
        )
        image.alpha_composite(stain_layer)


def create_paper_background(
    settings: RenderSettings,
    rng: random.Random,
) -> Image.Image:
    """
    生成学生作业本纸张。
    """
    warmth = int(settings.paper_age * 12)

    image = Image.new(
        "RGBA",
        (PAGE_WIDTH, PAGE_HEIGHT),
        (
            249 - warmth,
            247 - warmth,
            241 - int(warmth * 0.7),
            255,
        ),
    )

    add_light_paper_texture(
        image,
        settings,
        rng,
    )

    draw = ImageDraw.Draw(image)

    ruled_types = {
        "俄语作业本",
        "普通横线本",
    }

    if settings.paper_type in ruled_types:
        if settings.paper_type == "俄语作业本":
            # 顶部双线和左侧字段，接近俄语学生作业本。
            draw.line(
                (48, 62, PAGE_WIDTH - 48, 62),
                fill=(169, 180, 210, 220),
                width=2,
            )
            draw.line(
                (48, 78, PAGE_WIDTH - 48, 78),
                fill=(190, 198, 219, 190),
                width=1,
            )

            draw.text(
                (57, 31),
                "Дата",
                fill=(132, 143, 168, 205),
            )
            draw.line(
                (92, 48, 225, 48),
                fill=(155, 165, 192, 200),
                width=1,
            )

            draw.text(
                (PAGE_WIDTH - 285, 31),
                "Классная работа",
                fill=(132, 143, 168, 190),
            )

        for y in range(
            TOP_MARGIN,
            PAGE_HEIGHT - BOTTOM_MARGIN,
            settings.rule_spacing,
        ):
            line_color = (
                (164, 181, 216, 215)
                if settings.paper_type == "俄语作业本"
                else (177, 191, 220, 205)
            )

            draw.line(
                (50, y, PAGE_WIDTH - 50, y),
                fill=line_color,
                width=2,
            )

        red_line_x = LEFT_MARGIN + 48

        draw.line(
            (
                red_line_x,
                68,
                red_line_x,
                PAGE_HEIGHT - 68,
            ),
            fill=(202, 132, 137, 205),
            width=2,
        )

    elif settings.paper_type == "方格本":
        spacing = settings.rule_spacing

        for x in range(
            50,
            PAGE_WIDTH - 50,
            spacing,
        ):
            draw.line(
                (x, 50, x, PAGE_HEIGHT - 50),
                fill=(202, 213, 231, 175),
                width=1,
            )

        for y in range(
            50,
            PAGE_HEIGHT - 50,
            spacing,
        ):
            draw.line(
                (50, y, PAGE_WIDTH - 50, y),
                fill=(202, 213, 231, 175),
                width=1,
            )

    return image


def get_text_start_x(
    settings: RenderSettings,
) -> int:
    if settings.paper_type in {
        "俄语作业本",
        "普通横线本",
    }:
        return LEFT_MARGIN + 48 + 20

    return LEFT_MARGIN


def draw_notebook_header(
    page: Image.Image,
    settings: RenderSettings,
    rng: random.Random,
) -> None:
    """
    在页面顶部绘制日期和课题。
    """
    if not settings.header_enabled:
        return

    header_settings = replace(
        settings,
        font_size=max(30, settings.font_size - 10),
        texture_strength=min(
            1.0,
            settings.texture_strength + 0.04,
        ),
        randomness=min(
            2,
            settings.randomness,
        ),
        baseline_wave=0.25,
        flourish_level=0,
        correction_level=0,
        teacher_marks=0,
        header_enabled=False,
        show_page_number=False,
    )

    baseline_y = 92

    if settings.header_date.strip():
        draw_handwritten_line(
            page=page,
            text=settings.header_date.strip(),
            baseline_y=baseline_y,
            x_start=72,
            settings=header_settings,
            rng=rng,
        )

    if settings.header_lesson.strip():
        lesson_x = max(
            360,
            PAGE_WIDTH // 2 - 90,
        )

        draw_handwritten_line(
            page=page,
            text=settings.header_lesson.strip(),
            baseline_y=baseline_y,
            x_start=lesson_x,
            settings=header_settings,
            rng=rng,
        )


def draw_page_number(
    page: Image.Image,
    page_number: int,
    settings: RenderSettings,
    rng: random.Random,
) -> None:
    if not settings.show_page_number:
        return

    number_settings = replace(
        settings,
        font_size=28,
        randomness=1,
        baseline_wave=0.0,
        flourish_level=0,
        correction_level=0,
        teacher_marks=0,
        header_enabled=False,
        show_page_number=False,
    )

    draw_handwritten_line(
        page=page,
        text=str(page_number),
        baseline_y=PAGE_HEIGHT - 48,
        x_start=PAGE_WIDTH - 105,
        settings=number_settings,
        rng=rng,
    )


def apply_page_finish(
    image: Image.Image,
    settings: RenderSettings,
    rng: random.Random,
) -> Image.Image:
    """
    添加扫描件或手机拍照外观。
    """
    page = image.convert("RGB")

    if settings.photo_effect == "无":
        return page

    if settings.photo_effect == "扫描件":
        page = ImageOps.autocontrast(
            page,
            cutoff=0.4,
        )

        page = ImageEnhance.Contrast(
            page
        ).enhance(1.035)

        page = ImageEnhance.Sharpness(
            page
        ).enhance(1.08)

        vignette = Image.new(
            "L",
            page.size,
            255,
        )
        vignette_draw = ImageDraw.Draw(vignette)

        for inset in range(0, 65, 5):
            alpha = clamp(
                224 + inset // 2
            )

            vignette_draw.rectangle(
                (
                    inset,
                    inset,
                    PAGE_WIDTH - inset - 1,
                    PAGE_HEIGHT - inset - 1,
                ),
                outline=alpha,
                width=5,
            )

        shadow = Image.new(
            "RGB",
            page.size,
            (226, 226, 226),
        )

        return Image.composite(
            page,
            shadow,
            vignette,
        )

    # 手机拍照：桌面背景、页面阴影和轻微旋转。
    canvas_width = 1420
    canvas_height = 1940

    table = Image.new(
        "RGB",
        (canvas_width, canvas_height),
        (48, 47, 45),
    )

    table_draw = ImageDraw.Draw(table)

    for y in range(0, canvas_height, 7):
        shade = rng.randint(-5, 6)
        table_draw.line(
            (0, y, canvas_width, y),
            fill=(
                clamp(48 + shade),
                clamp(47 + shade),
                clamp(45 + shade),
            ),
            width=1,
        )

    angle = rng.uniform(-0.55, 0.55)

    page_rgba = page.convert("RGBA").rotate(
        angle,
        expand=True,
        resample=Image.Resampling.BICUBIC,
        fillcolor=(0, 0, 0, 0),
    )

    shadow_alpha = page_rgba.getchannel("A").filter(
        ImageFilter.GaussianBlur(radius=16)
    )

    shadow_layer = Image.new(
        "RGBA",
        page_rgba.size,
        (0, 0, 0, 0),
    )
    shadow_layer.putalpha(
        shadow_alpha.point(
            lambda pixel: int(pixel * 0.40)
        )
    )

    paste_x = (
        canvas_width - page_rgba.width
    ) // 2
    paste_y = (
        canvas_height - page_rgba.height
    ) // 2

    table_rgba = table.convert("RGBA")

    table_rgba.alpha_composite(
        shadow_layer,
        (paste_x + 15, paste_y + 18),
    )

    table_rgba.alpha_composite(
        page_rgba,
        (paste_x, paste_y),
    )

    return table_rgba.convert("RGB")


# ============================================================
# 自动换行和分页
# ============================================================
def text_width(
    text: str,
    font: ImageFont.FreeTypeFont,
) -> float:
    measuring_image = Image.new("RGB", (10, 10), "white")
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
        candidate = word if not current else f"{current} {word}"

        if text_width(candidate, font) <= max_width:
            current = candidate
            continue

        if current:
            lines.append(current)
            current = ""

        if text_width(word, font) <= max_width:
            current = word
            continue

        parts = split_long_word(word, font, max_width)
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
    available_height = PAGE_HEIGHT - TOP_MARGIN - BOTTOM_MARGIN
    maximum_rows = max(1, available_height // settings.rule_spacing)

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
def shear_right(
    image: Image.Image,
    shear: float,
) -> Image.Image:
    if shear <= 0:
        return image

    extra_width = int(abs(shear) * image.height)
    new_width = image.width + extra_width
    offset = -shear * image.height

    return image.transform(
        (new_width, image.height),
        Image.Transform.AFFINE,
        (1, shear, offset, 0, 1, 0),
        resample=Image.Resampling.BICUBIC,
    )


def sample_quadratic_curve(
    start: tuple[float, float],
    control: tuple[float, float],
    end: tuple[float, float],
    steps: int = 18,
) -> list[tuple[int, int]]:
    """
    生成二次贝塞尔曲线上的点。
    """
    points: list[tuple[int, int]] = []

    for index in range(steps + 1):
        t = index / steps
        one_minus_t = 1.0 - t

        x = (
            one_minus_t * one_minus_t * start[0]
            + 2 * one_minus_t * t * control[0]
            + t * t * end[0]
        )

        y = (
            one_minus_t * one_minus_t * start[1]
            + 2 * one_minus_t * t * control[1]
            + t * t * end[1]
        )

        points.append((round(x), round(y)))

    return points


def draw_connector_curve(
    page: Image.Image,
    start_x: int,
    start_y: int,
    end_x: int,
    end_y: int,
    color: tuple[int, int, int],
    alpha: int,
    strength: float,
    rng: random.Random,
) -> None:
    """
    用弧线连接相邻单词，让整行更像连续书写。
    """
    gap = end_x - start_x

    if gap <= 0 or gap > 44:
        return

    lift = rng.uniform(0.5, 3.5) * max(0.2, strength)
    control_x = start_x + gap * rng.uniform(0.42, 0.58)
    control_y = min(start_y, end_y) - lift

    points = sample_quadratic_curve(
        (start_x, start_y),
        (control_x, control_y),
        (end_x, end_y),
        steps=max(10, gap),
    )

    draw = ImageDraw.Draw(page)

    draw.line(
        points,
        fill=(color[0], color[1], color[2], alpha),
        width=1,
        joint="curve",
    )

    if strength > 0.72 and rng.random() < 0.35:
        second_points = [
            (x, y + 1)
            for x, y in points[:-2]
        ]

        draw.line(
            second_points,
            fill=(
                color[0],
                color[1],
                color[2],
                max(55, alpha - 115),
            ),
            width=1,
            joint="curve",
        )


def draw_entry_stroke(
    page: Image.Image,
    x: int,
    baseline_y: int,
    color: tuple[int, int, int],
    strength: float,
    rng: random.Random,
) -> None:
    """
    在一行开头偶尔添加轻微入笔。
    """
    if strength < 0.45 or rng.random() > 0.42:
        return

    length = rng.randint(7, 15)
    start = (x - length, baseline_y + rng.randint(1, 4))
    control = (x - length // 2, baseline_y - rng.randint(0, 3))
    end = (x + 1, baseline_y)

    points = sample_quadratic_curve(
        start,
        control,
        end,
        steps=14,
    )

    ImageDraw.Draw(page).line(
        points,
        fill=(color[0], color[1], color[2], 175),
        width=1,
        joint="curve",
    )


def draw_end_flourish(
    page: Image.Image,
    start_x: int,
    baseline_y: int,
    color: tuple[int, int, int],
    level: int,
    line_text: str,
    rng: random.Random,
) -> None:
    """
    在句尾或行尾添加自然收笔弧线。
    """
    if level <= 0 or start_x >= PAGE_WIDTH - RIGHT_MARGIN - 8:
        return

    base_probability = {
        1: 0.20,
        2: 0.43,
        3: 0.70,
    }.get(level, 0.20)

    if line_text.rstrip().endswith((".", "!", "?", "…", ":", ";")):
        base_probability += 0.18

    if rng.random() > min(0.92, base_probability):
        return

    available = PAGE_WIDTH - RIGHT_MARGIN - start_x
    length = min(
        available,
        rng.randint(12 + level * 5, 23 + level * 11),
    )

    if length < 7:
        return

    end_x = start_x + length
    end_y = baseline_y + rng.randint(-3, 3)

    control_x = start_x + length * rng.uniform(0.38, 0.64)
    control_y = baseline_y + rng.randint(-7, 3)

    points = sample_quadratic_curve(
        (start_x, baseline_y),
        (control_x, control_y),
        (end_x, end_y),
        steps=22,
    )

    draw = ImageDraw.Draw(page)
    draw.line(
        points,
        fill=(color[0], color[1], color[2], rng.randint(155, 215)),
        width=1,
        joint="curve",
    )

    if level >= 3 and rng.random() < 0.32:
        loop_radius = rng.randint(4, 8)
        draw.arc(
            (
                end_x - loop_radius,
                end_y - loop_radius,
                end_x + loop_radius,
                end_y + loop_radius,
            ),
            start=rng.randint(145, 190),
            end=rng.randint(315, 355),
            fill=(color[0], color[1], color[2], 145),
            width=1,
        )


def draw_strikethrough(
    page: Image.Image,
    left: int,
    top: int,
    width: int,
    height: int,
    color: tuple[int, int, int],
    rng: random.Random,
) -> None:
    """
    绘制一到两道手写删除线。
    """
    draw = ImageDraw.Draw(page)

    line_count = 1 if rng.random() < 0.43 else 2

    for index in range(line_count):
        start_y = top + int(
            height * rng.uniform(0.38, 0.66)
        ) + index * rng.randint(1, 3)

        end_y = start_y + rng.randint(-4, 4)

        middle_x = left + width // 2
        middle_y = (
            (start_y + end_y) // 2
            + rng.randint(-3, 3)
        )

        points = sample_quadratic_curve(
            (left - 3, start_y),
            (middle_x, middle_y),
            (left + width + 4, end_y),
            steps=max(14, width // 5),
        )

        draw.line(
            points,
            fill=(
                color[0],
                color[1],
                color[2],
                rng.randint(190, 235),
            ),
            width=1 if index else 2,
            joint="curve",
        )


def draw_scribble_correction(
    page: Image.Image,
    left: int,
    top: int,
    width: int,
    height: int,
    color: tuple[int, int, int],
    rng: random.Random,
) -> None:
    """
    在词上画出快速来回涂抹。
    """
    draw = ImageDraw.Draw(page)
    stroke_count = rng.randint(3, 6)

    for stroke_index in range(stroke_count):
        points: list[tuple[int, int]] = []
        segments = rng.randint(4, 7)

        for segment in range(segments + 1):
            x = left + int(width * segment / segments)
            x += rng.randint(-3, 3)

            relative_y = (
                0.30
                if segment % 2 == 0
                else 0.68
            )

            y = (
                top
                + int(height * relative_y)
                + rng.randint(-4, 4)
                + stroke_index
                - stroke_count // 2
            )

            points.append((x, y))

        draw.line(
            points,
            fill=(
                color[0],
                color[1],
                color[2],
                rng.randint(125, 205),
            ),
            width=rng.choice([1, 1, 2]),
            joint="curve",
        )


def draw_caret_mark(
    page: Image.Image,
    center_x: int,
    baseline_y: int,
    color: tuple[int, int, int],
    rng: random.Random,
) -> None:
    """
    绘制学生常用的插入符号。
    """
    size = rng.randint(6, 10)
    bottom_y = baseline_y + rng.randint(7, 11)
    top_y = bottom_y - size

    ImageDraw.Draw(page).line(
        (
            center_x - size,
            bottom_y,
            center_x,
            top_y,
            center_x + size,
            bottom_y,
        ),
        fill=(color[0], color[1], color[2], 205),
        width=1,
        joint="curve",
    )


def draw_rewrite_above(
    page: Image.Image,
    rewrite_text: str,
    left: int,
    top: int,
    settings: RenderSettings,
    rng: random.Random,
) -> None:
    """
    用同一套笔迹在原词上方重写。
    """
    rewrite_settings = replace(
        settings,
        font_size=max(24, settings.font_size - 9),
        randomness=max(0, settings.randomness - 1),
        correction_level=0,
        teacher_marks=0,
        flourish_level=0,
        baseline_wave=0.0,
        connection_strength=min(
            0.75,
            settings.connection_strength,
        ),
    )

    rewrite_image, _, _, _ = render_word_image(
        rewrite_text,
        rewrite_settings,
        rng,
    )

    rewrite_x = left + rng.randint(0, 9)
    rewrite_y = (
        top
        - int(settings.font_size * 0.48)
        + rng.randint(-3, 2)
    )

    page.alpha_composite(
        rewrite_image,
        (
            max(0, rewrite_x),
            max(0, rewrite_y),
        ),
    )


def maybe_draw_correction(
    page: Image.Image,
    word: str,
    left: int,
    top: int,
    width: int,
    height: int,
    baseline_y: int,
    settings: RenderSettings,
    rng: random.Random,
) -> None:
    """
    随机生成四种涂改：
    1. 删除线并在上方重写
    2. 来回涂抹
    3. 插入符号并在上方补写
    4. 重叠描写后划掉
    """
    if settings.correction_level <= 0:
        return

    clean_word = word.strip(
        ".,!?;:…—-()[]{}«»\"'"
    )

    if len(clean_word) < 3:
        return

    trigger_denominator = {
        1: 24,
        2: 13,
        3: 8,
    }.get(settings.correction_level, 13)

    if rng.randint(1, trigger_denominator) != 1:
        return

    color = varied_ink_color(
        settings,
        rng,
    )

    style = rng.choices(
        ["rewrite", "scribble", "caret", "overwrite"],
        weights=[38, 27, 20, 15],
        k=1,
    )[0]

    if style == "rewrite":
        draw_strikethrough(
            page,
            left,
            top,
            width,
            height,
            color,
            rng,
        )

        draw_rewrite_above(
            page,
            clean_word,
            left,
            top,
            settings,
            rng,
        )

    elif style == "scribble":
        draw_scribble_correction(
            page,
            left,
            top,
            width,
            height,
            color,
            rng,
        )

    elif style == "caret":
        draw_caret_mark(
            page,
            left + width // 2,
            baseline_y,
            color,
            rng,
        )

        short_rewrite = (
            clean_word
            if len(clean_word) <= 7
            else clean_word[: rng.randint(3, 6)]
        )

        draw_rewrite_above(
            page,
            short_rewrite,
            left + width // 3,
            top,
            settings,
            rng,
        )

    else:
        duplicate_settings = replace(
            settings,
            correction_level=0,
            teacher_marks=0,
            flourish_level=0,
            baseline_wave=0.0,
        )

        duplicate, _, _, _ = render_word_image(
            clean_word,
            duplicate_settings,
            rng,
        )

        duplicate_alpha = duplicate.getchannel("A").point(
            lambda pixel: int(pixel * 0.47)
        )
        duplicate.putalpha(duplicate_alpha)

        page.alpha_composite(
            duplicate,
            (
                max(0, left + rng.randint(-2, 3)),
                max(0, top + rng.randint(-2, 2)),
            ),
        )

        draw_strikethrough(
            page,
            left,
            top,
            width,
            height,
            color,
            rng,
        )


def render_word_image(
    word: str,
    settings: RenderSettings,
    rng: random.Random,
) -> tuple[
    Image.Image,
    int,
    tuple[int, int, int],
    int,
]:
    """
    用当前笔迹工具渲染单个词，主要用于涂改后的上方重写。
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

    try:
        bbox = measuring_draw.textbbox(
            (0, 0),
            word,
            font=high_font,
            anchor="ls",
            features=["calt", "liga", "clig", "kern"],
        )
    except Exception:
        bbox = measuring_draw.textbbox(
            (0, 0),
            word,
            font=high_font,
            anchor="ls",
        )

    content_width = max(
        1,
        bbox[2] - bbox[0],
    )

    canvas_width = (
        content_width
        + padding * 2
        + int(
            settings.slant
            * (ascent + descent)
        )
    )

    canvas_height = (
        ascent
        + descent
        + padding * 2
    )

    baseline_high = padding + ascent

    mask = Image.new(
        "L",
        (
            max(1, canvas_width),
            max(1, canvas_height),
        ),
        0,
    )

    draw = ImageDraw.Draw(mask)

    draw_text_with_cursive_features(
        draw,
        (
            padding - bbox[0],
            baseline_high,
        ),
        word,
        high_font,
        255,
        anchor="ls",
    )

    profile = get_pen_profile(settings)

    mask = mask.filter(
        ImageFilter.GaussianBlur(
            radius=float(
                profile["blur"]
            )
            * scale
        )
    )

    (
        alpha_mask,
        ink_color,
        word_alpha,
    ) = apply_pen_texture(
        mask,
        settings,
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

    word_image.putalpha(
        alpha_mask
    )

    width_change = rng.uniform(
        0.996
        - settings.randomness * 0.0008,
        1.004
        + settings.randomness * 0.0008,
    )

    height_change = rng.uniform(
        0.998
        - settings.randomness * 0.0005,
        1.002
        + settings.randomness * 0.0005,
    )

    target_width = max(
        1,
        int(
            word_image.width
            / scale
            * width_change
        ),
    )

    target_height = max(
        1,
        int(
            word_image.height
            / scale
            * height_change
        ),
    )

    baseline_after_resize = int(
        baseline_high
        / scale
        * height_change
    )

    word_image = word_image.resize(
        (
            target_width,
            target_height,
        ),
        Image.Resampling.LANCZOS,
    )

    word_image = shear_right(
        word_image,
        max(
            0.0,
            settings.slant
            + rng.uniform(-0.008, 0.008),
        ),
    )

    final_alpha = word_image.getchannel(
        "A"
    ).point(
        lambda pixel: int(
            pixel
            * word_alpha
            / 255
        )
    )

    word_image.putalpha(
        final_alpha
    )

    return (
        word_image,
        baseline_after_resize,
        ink_color,
        word_alpha,
    )


def warp_cursive_line(
    image: Image.Image,
    amplitude: float,
    rng: random.Random,
) -> Image.Image:
    """
    对整行文字进行非常轻微的纵向曲线变形。

    与逐词上下移动不同，这种方式不会切断单词内部的连笔，
    更接近参考图片中连续、平稳的俄语草写。
    """
    if amplitude <= 0.05 or image.width < 20:
        return image

    width, height = image.size
    segment_count = max(
        8,
        min(36, width // 45),
    )

    phase = rng.uniform(0.0, math.tau)
    frequency = rng.uniform(0.75, 1.25)
    drift = rng.uniform(-0.18, 0.18)

    boundaries = [
        round(index * width / segment_count)
        for index in range(segment_count + 1)
    ]

    offsets: list[float] = []

    for index in range(segment_count + 1):
        progress = index / segment_count

        sine_offset = (
            math.sin(
                phase
                + progress
                * math.tau
                * frequency
            )
            * amplitude
        )

        drift_offset = (
            drift
            * amplitude
            * (progress - 0.5)
        )

        offsets.append(
            sine_offset + drift_offset
        )

    mesh = []

    for index in range(segment_count):
        x0 = boundaries[index]
        x1 = boundaries[index + 1]

        if x1 <= x0:
            continue

        left_offset = offsets[index]
        right_offset = offsets[index + 1]

        mesh.append(
            (
                (x0, 0, x1, height),
                (
                    x0,
                    -left_offset,
                    x0,
                    height - left_offset,
                    x1,
                    height - right_offset,
                    x1,
                    -right_offset,
                ),
            )
        )

    return image.transform(
        image.size,
        Image.Transform.MESH,
        mesh,
        resample=Image.Resampling.BICUBIC,
    )


def render_cursive_line_image(
    text: str,
    settings: RenderSettings,
    rng: random.Random,
) -> tuple[
    Image.Image,
    int,
    tuple[int, int, int],
    int,
]:
    """
    把整行文字一次性绘制。

    这一版重点：
    1. 去掉字符间人工桥接，避免出现奇怪虚线；
    2. 保留字体自己的连字与连写；
    3. 增加钢笔 / 圆珠笔 / 中性笔 / 铅笔的纹理差异。
    """
    scale = settings.quality_scale

    high_font = load_font(
        settings.font_path,
        settings.font_size * scale,
    )

    ascent, descent = high_font.getmetrics()
    padding = 20 * scale
    baseline_high = padding + ascent

    measuring_image = Image.new(
        "L",
        (1, 1),
        0,
    )
    measuring_draw = ImageDraw.Draw(measuring_image)

    try:
        bbox = measuring_draw.textbbox(
            (0, 0),
            text,
            font=high_font,
            anchor="ls",
            features=["calt", "liga", "clig", "kern"],
        )
    except Exception:
        bbox = measuring_draw.textbbox(
            (0, 0),
            text,
            font=high_font,
            anchor="ls",
        )

    content_width = max(1, bbox[2] - bbox[0])

    canvas_width = (
        content_width
        + padding * 2
        + int(settings.slant * (ascent + descent))
    )

    canvas_height = ascent + descent + padding * 2

    mask = Image.new(
        "L",
        (max(1, canvas_width), max(1, canvas_height)),
        0,
    )

    mask_draw = ImageDraw.Draw(mask)

    draw_text_with_cursive_features(
        mask_draw,
        (padding - bbox[0], baseline_high),
        text,
        high_font,
        255,
        anchor="ls",
    )

    profile = get_pen_profile(settings)

    # 只做极轻微柔化，不再使用字符间桥接，避免虚线或怪异连接。
    blur_radius = float(profile["blur"]) * scale
    mask = mask.filter(
        ImageFilter.GaussianBlur(radius=blur_radius)
    )

    alpha_mask, ink_color, line_alpha = apply_pen_texture(
        mask,
        settings,
        rng,
    )

    line_image = Image.new(
        "RGBA",
        mask.size,
        (ink_color[0], ink_color[1], ink_color[2], 0),
    )
    line_image.putalpha(alpha_mask)

    space_count = text.count(" ")

    width_factor = (
        1.0
        - settings.connection_strength * 0.012
        + rng.uniform(-0.002, 0.002)
    )

    target_width = int(
        line_image.width / scale * width_factor
        + settings.word_spacing * space_count
    )

    height_factor = rng.uniform(
        0.998 - settings.randomness * 0.0004,
        1.002 + settings.randomness * 0.0004,
    )

    target_height = int(
        line_image.height / scale * height_factor
    )

    target_width = max(1, target_width)
    target_height = max(1, target_height)

    baseline_after_resize = int(
        baseline_high / scale * height_factor
    )

    line_image = line_image.resize(
        (target_width, target_height),
        Image.Resampling.LANCZOS,
    )

    effective_slant = max(
        0.0,
        settings.slant + settings.connection_strength * 0.008 + rng.uniform(-0.005, 0.005),
    )
    line_image = shear_right(line_image, effective_slant)

    line_image = warp_cursive_line(
        line_image,
        amplitude=(settings.baseline_wave * 0.55),
        rng=rng,
    )

    rotation_angle = rng.uniform(
        -0.030 - settings.randomness * 0.015,
        0.030 + settings.randomness * 0.015,
    )
    old_height = line_image.height

    line_image = line_image.rotate(
        rotation_angle,
        expand=True,
        resample=Image.Resampling.BICUBIC,
    )

    baseline_after_resize += (
        line_image.height - old_height
    ) // 2

    # 总体透明度微调
    alpha_final = line_image.getchannel("A").point(
        lambda pixel: min(255, max(0, int(pixel * line_alpha / 255)))
    )
    line_image.putalpha(alpha_final)

    visible_bbox = line_image.getchannel("A").getbbox()
    if visible_bbox is not None:
        baseline_after_resize -= visible_bbox[1]
        line_image = line_image.crop(visible_bbox)

    return line_image, baseline_after_resize, ink_color, line_alpha


def approximate_word_positions(
    text: str,
    x_start: int,
    rendered_width: int,
    font: ImageFont.FreeTypeFont,
) -> list[tuple[str, int, int]]:
    """
    估算整行渲染后各单词的位置，用于添加涂改痕迹。

    这里只影响涂改位置，不影响主体文字连笔。
    """
    raw_width = max(
        1.0,
        text_width(text, font),
    )

    width_scale = (
        rendered_width / raw_width
    )

    positions: list[
        tuple[str, int, int]
    ] = []

    search_from = 0

    for word in text.split():
        word_index = text.find(
            word,
            search_from,
        )

        if word_index < 0:
            continue

        prefix = text[:word_index]

        word_left = (
            x_start
            + round(
                text_width(
                    prefix,
                    font,
                )
                * width_scale
            )
        )

        word_width = max(
            8,
            round(
                text_width(
                    word,
                    font,
                )
                * width_scale
            ),
        )

        positions.append(
            (
                word,
                word_left,
                word_width,
            )
        )

        search_from = (
            word_index + len(word)
        )

    return positions


def draw_handwritten_line(
    page: Image.Image,
    text: str,
    baseline_y: int,
    x_start: int,
    settings: RenderSettings,
    rng: random.Random,
) -> int:
    """
    整行一次性绘制，主要模拟参考图中的俄语连笔。

    单词内部由字体真正连接；
    单词之间保留自然空格，不再额外画生硬的连接线。
    """
    if not text.strip():
        return x_start

    (
        line_image,
        line_baseline,
        ink_color,
        _,
    ) = render_cursive_line_image(
        text=text,
        settings=settings,
        rng=rng,
    )

    paste_y = (
        baseline_y
        - line_baseline
    )

    max_allowed_width = (
        PAGE_WIDTH
        - RIGHT_MARGIN
        - x_start
    )

    if (
        line_image.width
        > max_allowed_width
        and max_allowed_width > 10
    ):
        resize_ratio = (
            max_allowed_width
            / line_image.width
        )

        old_height = (
            line_image.height
        )

        line_image = line_image.resize(
            (
                max_allowed_width,
                max(
                    1,
                    round(
                        old_height
                        * resize_ratio
                    ),
                ),
            ),
            Image.Resampling.LANCZOS,
        )

        line_baseline = round(
            line_baseline
            * resize_ratio
        )

        paste_y = (
            baseline_y
            - line_baseline
        )

    draw_entry_stroke(
        page=page,
        x=x_start,
        baseline_y=baseline_y,
        color=ink_color,
        strength=(
            settings.connection_strength
        ),
        rng=rng,
    )

    page.alpha_composite(
        line_image,
        (
            max(0, x_start),
            max(0, paste_y),
        ),
    )

    # 主体整行已经完成，再按估算位置添加少量涂改。
    normal_font = load_font(
        settings.font_path,
        settings.font_size,
    )

    word_positions = (
        approximate_word_positions(
            text=text,
            x_start=x_start,
            rendered_width=line_image.width,
            font=normal_font,
        )
    )

    for (
        word,
        word_left,
        word_width,
    ) in word_positions:
        maybe_draw_correction(
            page=page,
            word=word,
            left=word_left,
            top=max(0, paste_y),
            width=word_width,
            height=line_image.height,
            baseline_y=baseline_y,
            settings=settings,
            rng=rng,
        )

    line_end_x = (
        x_start + line_image.width
    )

    draw_end_flourish(
        page=page,
        start_x=min(
            line_end_x - 2,
            PAGE_WIDTH
            - RIGHT_MARGIN
            - 2,
        ),
        baseline_y=baseline_y,
        color=ink_color,
        level=settings.flourish_level,
        line_text=text,
        rng=rng,
    )

    return line_end_x


def draw_teacher_feedback(
    page: Image.Image,
    line_baselines: list[int],
    text_start_x: int,
    settings: RenderSettings,
    rng: random.Random,
) -> None:
    """
    可选的老师红笔批改痕迹：
    对号、短下划线和圆圈。
    """
    if settings.teacher_marks <= 0 or not line_baselines:
        return

    mark_count = min(
        len(line_baselines),
        {
            1: rng.randint(1, 2),
            2: rng.randint(2, 4),
            3: rng.randint(4, 7),
        }.get(settings.teacher_marks, 2),
    )

    selected_baselines = rng.sample(
        line_baselines,
        k=mark_count,
    )

    draw = ImageDraw.Draw(page)
    red = (185, 46, 52, 210)

    for baseline in selected_baselines:
        style = rng.choice(
            ["check", "underline", "circle"]
        )

        if style == "check":
            x = PAGE_WIDTH - RIGHT_MARGIN + rng.randint(-14, 9)
            y = baseline - rng.randint(13, 22)

            draw.line(
                (
                    x - 8,
                    y + 8,
                    x - 2,
                    y + 14,
                    x + 12,
                    y - 4,
                ),
                fill=red,
                width=2,
                joint="curve",
            )

        elif style == "underline":
            start_x = rng.randint(
                text_start_x + 40,
                max(
                    text_start_x + 45,
                    PAGE_WIDTH - RIGHT_MARGIN - 240,
                ),
            )

            length = rng.randint(75, 180)
            y = baseline + rng.randint(3, 7)

            points = sample_quadratic_curve(
                (start_x, y),
                (
                    start_x + length // 2,
                    y + rng.randint(-3, 3),
                ),
                (
                    min(
                        PAGE_WIDTH - RIGHT_MARGIN,
                        start_x + length,
                    ),
                    y + rng.randint(-2, 2),
                ),
                steps=24,
            )

            draw.line(
                points,
                fill=red,
                width=2,
                joint="curve",
            )

        else:
            center_x = PAGE_WIDTH - RIGHT_MARGIN - rng.randint(35, 90)
            center_y = baseline - rng.randint(10, 20)
            radius_x = rng.randint(12, 20)
            radius_y = rng.randint(8, 15)

            draw.ellipse(
                (
                    center_x - radius_x,
                    center_y - radius_y,
                    center_x + radius_x,
                    center_y + radius_y,
                ),
                outline=red,
                width=2,
            )


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

    text_start_x = get_text_start_x(
        settings
    )

    draw_notebook_header(
        page=image,
        settings=settings,
        rng=rng,
    )

    previous_was_blank = True
    written_baselines: list[int] = []

    for row_index, line in enumerate(
        page_lines
    ):
        rule_y = (
            TOP_MARGIN
            + row_index
            * settings.rule_spacing
        )

        if not line.strip():
            previous_was_blank = True
            continue

        paragraph_indent = (
            rng.randint(8, 20)
            if previous_was_blank
            else 0
        )

        x_start = (
            text_start_x
            + paragraph_indent
            + rng.randint(
                -max(1, settings.randomness),
                max(2, settings.randomness + 1),
            )
        )

        baseline_y = (
            rule_y
            - 2
            + rng.randint(
                -max(1, settings.randomness // 2),
                max(1, settings.randomness // 2),
            )
        )

        draw_handwritten_line(
            page=image,
            text=line,
            baseline_y=baseline_y,
            x_start=x_start,
            settings=settings,
            rng=rng,
        )

        written_baselines.append(
            baseline_y
        )

        previous_was_blank = False

    draw_teacher_feedback(
        page=image,
        line_baselines=written_baselines,
        text_start_x=text_start_x,
        settings=settings,
        rng=rng,
    )

    draw_page_number(
        page=image,
        page_number=page_number + 1,
        settings=settings,
        rng=rng,
    )

    return apply_page_finish(
        image,
        settings,
        rng,
    )


def render_document(
    text: str,
    settings: RenderSettings,
) -> list[Image.Image]:
    normal_font = load_font(settings.font_path, settings.font_size)

    text_start_x = get_text_start_x(
        settings
    )

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

    pages_lines = paginate_lines(lines=lines, settings=settings)

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
    image.save(buffer, format="PNG", optimize=True)
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
# Streamlit 页面
# ============================================================
st.set_page_config(
    page_title="俄语手写图片生成器",
    page_icon="✍️",
    layout="wide",
)

st.title("✍️ 俄语学生作业本生成器")
st.caption("学生作业本终极版：预设、俄语本子页眉、真实笔迹纹理、涂改、批改、扫描和手机拍照外观。")

font_files = get_font_files()

if not font_files:
    st.error("没有找到字体。请上传一个支持俄语的 .ttf 或 .otf 手写字体。")
    st.code(
        "推荐结构：\n"
        "app.py\n"
        "fonts/\n"
        "    BadScript-Regular.ttf"
    )
    st.stop()


with st.sidebar:
    st.header("学生作业本设置")

    preset_name = st.selectbox(
        "快速预设",
        options=list(PRESETS.keys()),
        help="先选择最接近的效果，再微调下面参数。",
    )

    preset = PRESETS[preset_name]
    preset_key = preset_name.replace(" ", "_")

    selected_font_name = st.selectbox(
        "手写字体",
        options=list(font_files.keys()),
        help=(
            "俄语连笔推荐 Marck Script 或 Bad Script。"
            "真正的字母连接形状主要由字体决定。"
        ),
        key=f"font_{preset_key}",
    )

    font_size = st.slider(
        "字体大小",
        min_value=36,
        max_value=72,
        value=int(preset["font_size"]),
        key=f"font_size_{preset_key}",
    )

    line_spacing = st.slider(
        "横线间距",
        min_value=4,
        max_value=24,
        value=int(preset["line_spacing"]),
        key=f"line_spacing_{preset_key}",
    )

    paper_options = [
        "俄语作业本",
        "普通横线本",
        "方格本",
        "白纸",
    ]

    paper_type = st.selectbox(
        "本子类型",
        paper_options,
        index=option_index(
            paper_options,
            preset["paper_type"],
        ),
        key=f"paper_{preset_key}",
    )

    paper_age = st.slider(
        "纸张旧化",
        min_value=0.0,
        max_value=1.0,
        value=float(preset["paper_age"]),
        step=0.05,
        key=f"paper_age_{preset_key}",
    )

    photo_options = [
        "无",
        "扫描件",
        "手机拍照",
    ]

    photo_effect = st.selectbox(
        "导出外观",
        photo_options,
        index=option_index(
            photo_options,
            preset["photo_effect"],
        ),
        key=f"photo_{preset_key}",
    )

    ink_options = [
        "黑色",
        "蓝色",
    ]

    ink_name = st.selectbox(
        "墨水颜色",
        ink_options,
        index=option_index(
            ink_options,
            preset["ink_name"],
        ),
        key=f"ink_{preset_key}",
    )

    pen_options = [
        "中性笔",
        "钢笔",
        "圆珠笔",
        "铅笔",
    ]

    pen_style = st.selectbox(
        "书写工具",
        pen_options,
        index=option_index(
            pen_options,
            preset["pen_style"],
        ),
        key=f"pen_{preset_key}",
    )

    texture_strength = st.slider(
        "笔迹纹理",
        min_value=0.0,
        max_value=1.0,
        value=float(
            preset["texture_strength"]
        ),
        step=0.05,
        help=(
            "铅笔建议 0.55～0.80；"
            "圆珠笔建议 0.30～0.50；"
            "钢笔建议 0.15～0.35。"
        ),
        key=f"texture_{preset_key}",
    )

    randomness = st.slider(
        "自然随机程度",
        min_value=0,
        max_value=4,
        value=int(preset["randomness"]),
        help="整齐作业建议 1，课堂笔记建议 2～3。",
        key=f"random_{preset_key}",
    )

    slant = st.slider(
        "右倾程度",
        min_value=0.00,
        max_value=0.20,
        value=float(preset["slant"]),
        step=0.01,
        key=f"slant_{preset_key}",
    )

    connection_strength = st.slider(
        "连笔流畅度",
        min_value=0.0,
        max_value=1.0,
        value=float(
            preset["connection_strength"]
        ),
        step=0.05,
        help=(
            "不会人工画字符连接线，只调整字体连字、"
            "整体压缩和连写感觉。"
        ),
        key=f"connection_{preset_key}",
    )

    word_spacing = st.slider(
        "单词间距",
        min_value=-8,
        max_value=12,
        value=int(preset["word_spacing"]),
        help="单词应保持可辨认，通常使用 -2～1。",
        key=f"word_spacing_{preset_key}",
    )

    baseline_wave = st.slider(
        "基线起伏",
        min_value=0.0,
        max_value=5.0,
        value=float(
            preset["baseline_wave"]
        ),
        step=0.2,
        key=f"baseline_{preset_key}",
    )

    flourish_level = st.select_slider(
        "句尾收笔",
        options=[0, 1, 2, 3],
        value=int(
            preset["flourish_level"]
        ),
        key=f"flourish_{preset_key}",
    )

    correction_level = st.select_slider(
        "学生涂改痕迹",
        options=[0, 1, 2, 3],
        value=int(
            preset["correction_level"]
        ),
        help=(
            "包括删除线、快速涂抹、插入符号、"
            "上方重写和重叠描写。"
        ),
        key=f"correction_{preset_key}",
    )

    teacher_marks = st.select_slider(
        "老师红笔批改",
        options=[0, 1, 2, 3],
        value=int(
            preset["teacher_marks"]
        ),
        key=f"teacher_{preset_key}",
    )

    st.divider()
    st.subheader("页眉和页码")

    header_enabled = st.checkbox(
        "显示手写页眉",
        value=True,
        key=f"header_enabled_{preset_key}",
    )

    header_date = st.text_input(
        "日期",
        value="15.01",
        key=f"header_date_{preset_key}",
    )

    header_lesson = st.text_input(
        "课题",
        value="Урок русского языка",
        key=f"header_lesson_{preset_key}",
    )

    show_page_number = st.checkbox(
        "显示手写页码",
        value=False,
        key=f"page_number_{preset_key}",
    )

    seed = st.number_input(
        "随机种子",
        min_value=0,
        max_value=999_999,
        value=12345,
        step=1,
        key=f"seed_{preset_key}",
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
    file_identity = (uploaded_file.name, uploaded_file.size)

    if st.session_state.get("last_uploaded_file") != file_identity:
        try:
            extracted_text = extract_uploaded_file(uploaded_file)
            extracted_text = normalize_text(extracted_text)

            if extracted_text:
                st.session_state["editor_text"] = extracted_text
                st.session_state["last_uploaded_file"] = file_identity
                st.success("文字提取成功，可以在下面继续修改。")
            else:
                st.warning("没有提取到文字。这个 PDF 可能是扫描版图片，当前版本暂未加入 OCR。")

        except Exception as error:
            st.error(f"读取文件失败：{error}")


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
        st.warning("请先输入或上传俄语文字。")
    else:
        settings = RenderSettings(
            font_path=str(
                font_files[selected_font_name]
            ),
            font_size=font_size,
            line_spacing=line_spacing,
            paper_type=paper_type,
            paper_age=float(paper_age),
            photo_effect=photo_effect,
            ink_name=ink_name,
            pen_style=pen_style,
            texture_strength=float(
                texture_strength
            ),
            randomness=randomness,
            seed=int(seed),
            slant=float(slant),
            word_spacing=word_spacing,
            connection_strength=float(
                connection_strength
            ),
            baseline_wave=float(
                baseline_wave
            ),
            flourish_level=int(
                flourish_level
            ),
            correction_level=int(
                correction_level
            ),
            teacher_marks=int(
                teacher_marks
            ),
            header_enabled=bool(
                header_enabled
            ),
            header_date=header_date,
            header_lesson=header_lesson,
            show_page_number=bool(
                show_page_number
            ),
        )

        try:
            with st.spinner("正在生成作业本手写页面……"):
                pages = render_document(text=text, settings=settings)
                save_result_to_session(pages)

            st.success(f"生成成功，共 {len(pages)} 页。")

        except OSError:
            st.error("字体无法打开。请确认字体文件有效，并且支持西里尔字母。")

        except MemoryError:
            st.error("文字太多，电脑内存不足。请减少文字后分批生成。")

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
            f"全部 {result['page_count']} 页可以下载为 ZIP 或 PDF。"
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
