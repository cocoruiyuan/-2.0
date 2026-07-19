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
    connection_strength: float
    baseline_wave: float
    flourish_level: int
    correction_level: int
    teacher_marks: int
    quality_scale: int = 3

    @property
    def rule_spacing(self) -> int:
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
    rng: random.Random,
) -> None:
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
    image = Image.new(
        "RGB",
        (PAGE_WIDTH, PAGE_HEIGHT),
        (247, 244, 235),
    )

    add_light_paper_texture(image, rng)
    draw = ImageDraw.Draw(image)

    if settings.paper_type == "横线纸":
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

        for y in range(TOP_MARGIN, PAGE_HEIGHT - BOTTOM_MARGIN, settings.rule_spacing):
            draw.line(
                (55, y, PAGE_WIDTH - 55, y),
                fill=(167, 184, 218),
                width=2,
            )

        red_line_x = LEFT_MARGIN + 48
        draw.line(
            (red_line_x, 70, red_line_x, PAGE_HEIGHT - 70),
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
def varied_ink_color(
    ink_name: str,
    rng: random.Random,
) -> tuple[int, int, int]:
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
    width, height = text_mask.size
    small_width = max(2, width // 35)
    small_height = max(2, height // 14)

    pressure_pixels = bytes(
        rng.randint(236, 255)
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
        settings.ink_name,
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
            features=["calt", "liga", "clig"],
        )
    except Exception:
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

    draw_text_with_cursive_features(
        mask_draw,
        (
            padding - bbox[0],
            baseline_y,
        ),
        word,
        high_font,
        255,
        anchor="ls",
    )

    mask = mask.filter(
        ImageFilter.GaussianBlur(
            radius=0.042 * scale
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

    word_alpha = rng.randint(226, 249)

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

    width_change = rng.uniform(
        0.9955
        - settings.randomness * 0.0009,
        1.0045
        + settings.randomness * 0.0009,
    )

    height_change = rng.uniform(
        0.9975
        - settings.randomness * 0.0006,
        1.0025
        + settings.randomness * 0.0006,
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
        + settings.connection_strength * 0.018
        + rng.uniform(-0.012, 0.012),
    )

    word_image = shear_right(
        word_image,
        word_slant,
    )

    if settings.randomness > 0:
        rotation_angle = rng.uniform(
            -0.10 - settings.randomness * 0.055,
            0.10 + settings.randomness * 0.055,
        )

        old_height = word_image.height

        word_image = word_image.rotate(
            rotation_angle,
            expand=True,
            resample=Image.Resampling.BICUBIC,
        )

        baseline_after_resize += (
            word_image.height - old_height
        ) // 2

    alpha_mask = word_image.getchannel(
        "A"
    ).point(
        lambda pixel: min(
            255,
            max(
                0,
                int(
                    pixel
                    * (word_alpha / 255)
                ),
            ),
        )
    )

    word_image.putalpha(alpha_mask)

    return (
        word_image,
        baseline_after_resize,
        ink_color,
        word_alpha,
    )


def draw_handwritten_line(
    page: Image.Image,
    text: str,
    baseline_y: int,
    x_start: int,
    settings: RenderSettings,
    rng: random.Random,
) -> int:
    """
    返回该行结束位置，供句尾收笔使用。
    """
    words = text.split()

    if not words:
        return x_start

    normal_font = load_font(
        settings.font_path,
        settings.font_size,
    )

    base_space = int(
        text_width(" ", normal_font)
    )

    current_x = x_start
    previous_end_y: int | None = None
    previous_tail_x: int | None = None
    last_ink_color = varied_ink_color(
        settings.ink_name,
        rng,
    )

    phase = rng.uniform(0.0, math.tau)
    slow_drift = rng.uniform(-0.16, 0.16)

    for word_index, word in enumerate(words):
        (
            word_image,
            word_baseline,
            ink_color,
            word_alpha,
        ) = render_word_image(
            word=word,
            settings=settings,
            rng=rng,
        )

        wave_shift = (
            math.sin(
                phase
                + word_index * 0.82
            )
            * settings.baseline_wave
        )

        drift_shift = (
            slow_drift
            * word_index
            * min(
                1.0,
                settings.baseline_wave / 2.0,
            )
        )

        random_shift = rng.randint(
            -max(1, settings.randomness),
            max(1, settings.randomness),
        )

        word_baseline_y = round(
            baseline_y
            + wave_shift
            + drift_shift
            + random_shift
        )

        paste_y = (
            word_baseline_y
            - word_baseline
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

        if word_index == 0:
            draw_entry_stroke(
                page,
                current_x,
                word_baseline_y,
                ink_color,
                settings.connection_strength,
                rng,
            )

        if (
            settings.connection_strength > 0
            and previous_tail_x is not None
            and previous_end_y is not None
        ):
            connection_probability = min(
                0.96,
                0.23
                + settings.connection_strength * 0.69,
            )

            if rng.random() < connection_probability:
                draw_connector_curve(
                    page=page,
                    start_x=previous_tail_x,
                    start_y=previous_end_y,
                    end_x=current_x + 2,
                    end_y=word_baseline_y,
                    color=ink_color,
                    alpha=max(
                        110,
                        word_alpha - 55,
                    ),
                    strength=(
                        settings.connection_strength
                    ),
                    rng=rng,
                )

        page.alpha_composite(
            word_image,
            (
                max(0, current_x),
                max(0, paste_y),
            ),
        )

        maybe_draw_correction(
            page=page,
            word=word,
            left=max(0, current_x),
            top=max(0, paste_y),
            width=word_image.width,
            height=word_image.height,
            baseline_y=word_baseline_y,
            settings=settings,
            rng=rng,
        )

        previous_tail_x = (
            current_x
            + max(
                2,
                word_image.width
                - rng.randint(5, 13),
            )
        )

        previous_end_y = (
            word_baseline_y
            + rng.randint(-1, 2)
        )

        last_ink_color = ink_color

        current_x += (
            word_image.width
            + base_space
            + settings.word_spacing
            - int(
                settings.connection_strength
                * 3.2
            )
            + rng.randint(-1, 2)
        )

    draw_end_flourish(
        page=page,
        start_x=min(
            current_x - max(2, base_space // 2),
            PAGE_WIDTH - RIGHT_MARGIN - 2,
        ),
        baseline_y=(
            previous_end_y
            if previous_end_y is not None
            else baseline_y
        ),
        color=last_ink_color,
        level=settings.flourish_level,
        line_text=text,
        rng=rng,
    )

    return current_x


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

    red_line_x = LEFT_MARGIN + 48
    text_start_x = red_line_x + 20

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
            rng.randint(8, 18)
            if previous_was_blank
            else 0
        )

        x_start = (
            text_start_x
            + paragraph_indent
            + rng.randint(-1, 2)
        )

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

    return image.convert("RGB")


def render_document(
    text: str,
    settings: RenderSettings,
) -> list[Image.Image]:
    normal_font = load_font(settings.font_path, settings.font_size)

    red_line_x = LEFT_MARGIN + 48
    text_start_x = red_line_x + 20

    max_text_width = PAGE_WIDTH - RIGHT_MARGIN - text_start_x - 28

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

st.title("✍️ 俄语手写图片生成器")
st.caption("进一步增强版：字体连字、自然基线、句尾收笔、多种涂改和可选红笔批改。")

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
    st.header("手写设置")

    selected_font_name = st.selectbox(
        "手写字体",
        options=list(font_files.keys()),
        help="想接近参考图，推荐使用 Bad Script，较适合增强连笔。",
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
        ["横线纸", "白纸", "方格纸"],
    )

    ink_name = st.selectbox(
        "墨水",
        ["黑色", "蓝色"],
    )

    randomness = st.slider(
        "自然随机程度",
        min_value=0,
        max_value=4,
        value=1,
        help="参考图比较整齐，建议保持在 1～2。",
    )

    slant = st.slider(
        "右倾程度",
        min_value=0.00,
        max_value=0.20,
        value=0.10,
        step=0.01,
    )

    connection_strength = st.slider(
        "连笔增强",
        min_value=0.0,
        max_value=1.0,
        value=0.62,
        step=0.05,
        help=(
            "启用字体连字特性，并在相邻单词之间"
            "增加自然弧形连接。推荐 0.50～0.75。"
        ),
    )

    word_spacing = st.slider(
        "单词间距",
        min_value=-8,
        max_value=12,
        value=-2,
        help="想更像连续书写，可以设置为 -1～-3。",
    )

    baseline_wave = st.slider(
        "基线起伏",
        min_value=0.0,
        max_value=5.0,
        value=1.2,
        step=0.2,
        help="让同一行文字轻微上下起伏。建议 0.8～1.8。",
    )

    flourish_level = st.select_slider(
        "句尾收笔",
        options=[0, 1, 2, 3],
        value=2,
        help="0=关闭，1=少量，2=自然，3=较明显。",
    )

    correction_level = st.select_slider(
        "涂改痕迹",
        options=[0, 1, 2, 3],
        value=1,
        help=(
            "包含删除线、快速涂抹、插入符号、"
            "上方重写和重叠描写。"
        ),
    )

    teacher_marks = st.select_slider(
        "老师红笔批改",
        options=[0, 1, 2, 3],
        value=0,
        help="0=关闭，1=少量，2=中等，3=较多。",
    )

    seed = st.number_input(
        "随机种子",
        min_value=0,
        max_value=999_999,
        value=12345,
        step=1,
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
            font_path=str(font_files[selected_font_name]),
            font_size=font_size,
            line_spacing=line_spacing,
            paper_type=paper_type,
            ink_name=ink_name,
            randomness=randomness,
            seed=int(seed),
            slant=float(slant),
            word_spacing=word_spacing,
            connection_strength=float(connection_strength),
            baseline_wave=float(baseline_wave),
            flourish_level=int(flourish_level),
            correction_level=int(correction_level),
            teacher_marks=int(teacher_marks),
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
