from __future__ import annotations

import math
import re
import struct
import zlib
import xml.etree.ElementTree as ET
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "src"
OUTPUT = ROOT / "build" / "palette12"
CUSTOM_UI_NS = "http://schemas.microsoft.com/office/2009/07/customui"
NS = f"{{{CUSTOM_UI_NS}}}"
SHADE_NAMES = (
    "lightest",
    "very_light",
    "light",
    "soft",
    "medium",
    "rich",
    "deep",
    "dark",
    "darkest",
)

# Eight hues named with the most familiar English color words: the seven
# chromatic basic color terms except brown, plus teal, the most common name
# for blue-green in the xkcd color survey. Each OKLCH hue sits within 10
# degrees of the survey's typical color for its name, placed to keep the
# closest pair of swatches in different groups as far apart as possible
# (within 2% of the optimum), preferring the typical hue where that costs
# nothing. Yellow stays at 100 rather than its typical 110, which turns
# olive in the mid and dark shades. None marks the neutral group.
PALETTE = (
    ("Red", 20),
    ("Orange", 59),
    ("Yellow", 100),
    ("Green", 143),
    ("Teal", 184),
    ("Blue", 263),
    ("Purple", 307),
    ("Pink", 343),
    ("Gray", None),
)

# Shades use an even OKLCH lightness ladder and a chroma target per shade,
# capped at what sRGB can show for the hue. 0.88-0.36 rather than 0.94-0.30
# leaves enough chroma at both ends that the lightest and darkest shades of
# neighboring hues stay distinguishable. Gray runs from white to black on its
# own even sRGB ladder instead (see shade_rgb), since it is the one group
# where readers expect the literal extremes and an even visual spread.
LIGHTNESS_TOP = 0.88
LIGHTNESS_BOTTOM = 0.36
CHROMA_TARGETS = (0.085, 0.115, 0.145, 0.165, 0.175, 0.170, 0.155, 0.135, 0.105)


def linear_to_srgb(value: float) -> int:
    value = min(max(value, 0.0), 1.0)
    encoded = 12.92 * value if value <= 0.0031308 else 1.055 * value ** (1 / 2.4) - 0.055
    return round(encoded * 255)


def oklch_to_linear_rgb(lightness: float, chroma: float, hue: float) -> tuple[float, float, float]:
    a = chroma * math.cos(math.radians(hue))
    b = chroma * math.sin(math.radians(hue))
    l = (lightness + 0.3963377774 * a + 0.2158037573 * b) ** 3
    m = (lightness - 0.1055613458 * a - 0.0638541728 * b) ** 3
    s = (lightness - 0.0894841775 * a - 1.2914855480 * b) ** 3
    return (
        4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s,
        -1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s,
        -0.0041960863 * l - 0.7034186147 * m + 1.7076147010 * s,
    )


def max_chroma(lightness: float, hue: float) -> float:
    low, high = 0.0, 0.5
    for _ in range(40):
        middle = (low + high) / 2
        if all(0 <= value <= 1 for value in oklch_to_linear_rgb(lightness, middle, hue)):
            low = middle
        else:
            high = middle
    return low


def shade_rgb(shade_index: int, hue: float | None) -> tuple[int, int, int]:
    if hue is None:
        # OKLab lightness steps compress heavily near black once gamma-encoded
        # (the two darkest shades become nearly indistinguishable), so gray
        # steps evenly through the sRGB byte range instead.
        value = round(255 - (shade_index - 1) * 255 / (len(SHADE_NAMES) - 1))
        return (value, value, value)
    step = (LIGHTNESS_TOP - LIGHTNESS_BOTTOM) / (len(SHADE_NAMES) - 1)
    lightness = LIGHTNESS_TOP - step * (shade_index - 1)
    chroma = min(CHROMA_TARGETS[shade_index - 1], max_chroma(lightness, hue))
    return tuple(linear_to_srgb(value) for value in oklch_to_linear_rgb(lightness, chroma, hue))


def png_rgb(width: int, height: int, rgb: tuple[int, int, int], border: bool = False) -> bytes:
    if border:
        black_row = b"\x00" + bytes((0, 0, 0)) * width
        middle_row = b"\x00" + bytes((0, 0, 0)) + bytes(rgb) * (width - 2) + bytes((0, 0, 0))
        raw = black_row + middle_row * (height - 2) + black_row
    else:
        row = b"\x00" + bytes(rgb) * width
        raw = row * height

    def chunk(kind: bytes, payload: bytes) -> bytes:
        return (
            struct.pack(">I", len(payload))
            + kind
            + payload
            + struct.pack(">I", zlib.crc32(kind + payload) & 0xFFFFFFFF)
        )

    return b"\x89PNG\r\n\x1a\n" + chunk(
        b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    ) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")


def slug(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", value.lower()).strip("_")


def palette_group(color_index: int, color_name: str, hue: float | None, icon_dir: Path) -> ET.Element:
    # One Ribbon group per color, so Office draws a separator between colors
    # and shows the name beneath the 3x3 grid of shades.
    group = ET.Element(NS + "group", {"id": f"grpPalette{color_index:02d}", "label": color_name})
    grid = ET.SubElement(group, NS + "box", {"id": f"palette{color_index:02d}", "boxStyle": "vertical"})
    rows = [
        ET.SubElement(
            grid,
            NS + "box",
            {"id": f"palette{color_index:02d}_row{row:02d}", "boxStyle": "horizontal"},
        )
        for row in range(1, 4)
    ]
    for shade_index, shade_name in enumerate(SHADE_NAMES, start=1):
        rgb = shade_rgb(shade_index, hue)
        image_id = f"{slug(color_name)}_{shade_name}"
        # The lightest gray is pure white, which is invisible against the
        # ribbon background without an outline.
        border = hue is None and rgb == (255, 255, 255)
        (icon_dir / f"{image_id}.png").write_bytes(png_rgb(16, 16, rgb, border=border))
        ET.SubElement(
            rows[(shade_index - 1) // 3],
            NS + "toggleButton",
            {
                "id": f"btn{color_index:02d}_{shade_index:02d}",
                "image": image_id,
                "showLabel": "false",
                "tag": ",".join(str(channel) for channel in rgb),
                "screentip": f"{color_name} — {shade_name.replace('_', ' ').title()}",
                "supertip": "#{:02X}{:02X}{:02X}".format(*rgb)
                + f" | RGB {rgb[0]}, {rgb[1]}, {rgb[2]}"
                + f" | Color {color_index:02d} | Shade {shade_index:02d}",
                "getPressed": "PaletteHighlighter_GetPressed",
                "onAction": "PaletteHighlighter_Color",
            },
        )
    return group


def main() -> None:
    ET.register_namespace("", CUSTOM_UI_NS)
    root = ET.parse(SOURCE / "ribbonx.xml").getroot()
    tabs = root.find(NS + "ribbon/" + NS + "tabs")
    tab = tabs.find(NS + "tab") if tabs is not None else None
    if tab is None:
        raise ValueError("RibbonX source does not contain the Highlighter tab")

    OUTPUT.mkdir(parents=True, exist_ok=True)
    icon_dir = OUTPUT / "icons"
    icon_dir.mkdir(parents=True, exist_ok=True)
    for old_icon in icon_dir.glob("*.png"):
        old_icon.unlink()

    for insert_index, (color_name, hue) in enumerate(PALETTE):
        tab.insert(insert_index, palette_group(insert_index + 1, color_name, hue, icon_dir))

    xml_path = OUTPUT / "ribbonx.xml"
    xml_path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))
    print(f"palette={xml_path}")
    print(f"groups={len(PALETTE)}")
    print(f"shades_per_color={len(SHADE_NAMES)}")
    print(f"icons={len(list(icon_dir.glob('*.png')))}")


if __name__ == "__main__":
    main()
