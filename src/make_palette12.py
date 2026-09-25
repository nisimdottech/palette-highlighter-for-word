from __future__ import annotations

import copy
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

# The eighteen source families sit 20 degrees apart in OKLCH hue. Every other
# family gives nine hues evenly spaced 40 degrees apart, so no two neighbors
# are closer than any other pair.
SOURCE_GROUPS = (
    "grp01",  # Yellow, 95
    "grp03",  # Yellow Green, 135
    "grp05",  # Emerald, 175
    "grp07",  # Cyan, 215
    "grp09",  # Blue, 255
    "grp11",  # Violet, 295
    "grp13",  # Magenta, 335
    "grp15",  # Red, 15
    "grp17",  # Orange, 55
)


def png_rgb(width: int, height: int, rgb: tuple[int, int, int]) -> bytes:
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


def main() -> None:
    ET.register_namespace("", CUSTOM_UI_NS)
    root = ET.parse(SOURCE / "ribbonx.xml").getroot()
    tabs = root.find(NS + "ribbon/" + NS + "tabs")
    tab = tabs.find(NS + "tab") if tabs is not None else None
    if tab is None:
        raise ValueError("RibbonX source does not contain the Highlighter tab")

    groups = {group.attrib.get("id"): group for group in tab.findall(NS + "group")}
    palette_groups = []

    OUTPUT.mkdir(parents=True, exist_ok=True)
    icon_dir = OUTPUT / "icons"
    icon_dir.mkdir(parents=True, exist_ok=True)
    for old_icon in icon_dir.glob("*.png"):
        old_icon.unlink()

    for new_color_index, source_group_id in enumerate(SOURCE_GROUPS, start=1):
        source_group = groups[source_group_id]
        grid = source_group.find(NS + "box")
        if grid is None:
            raise ValueError(f"Missing grid in {source_group_id}")

        # Keep one Ribbon group per hue so Office draws the original visual
        # separator between adjacent colors.
        palette = ET.Element(NS + "group", {"id": f"grpPalette{new_color_index:02d}"})
        palette_box = ET.SubElement(
            palette,
            NS + "box",
            {"id": f"palette{new_color_index:02d}", "boxStyle": "vertical"},
        )
        palette_rows = [
            ET.SubElement(
                palette_box,
                NS + "box",
                {
                    "id": f"palette{new_color_index:02d}_row{row:02d}",
                    "boxStyle": "horizontal",
                },
            )
            for row in range(1, 4)
        ]

        color_name = ""
        rows = grid.findall(NS + "box")
        for row_index, row in enumerate(rows, start=1):
            for column_index, button in enumerate(row.findall(NS + "toggleButton"), start=1):
                source_shade_index = (row_index - 1) * 3 + column_index
                shade_index = source_shade_index
                button = copy.deepcopy(button)
                if not color_name:
                    color_name = button.attrib["screentip"].split(" — ", 1)[0]
                rgb = tuple(int(part) for part in button.attrib["tag"].split(","))
                image_id = f"{slug(color_name)}_{SHADE_NAMES[shade_index - 1]}"
                button.attrib["id"] = f"btn{new_color_index:02d}_{shade_index:02d}"
                button.attrib["image"] = image_id
                button.attrib["supertip"] = button.attrib["supertip"].replace(
                    re.search(r"Color \d+", button.attrib["supertip"]).group(0),
                    f"Color {new_color_index:02d}",
                )
                button.attrib["supertip"] = re.sub(
                    r"Shade \d+",
                    f"Shade {shade_index:02d}",
                    button.attrib["supertip"],
                )
                (icon_dir / f"{image_id}.png").write_bytes(png_rgb(16, 16, rgb))
                palette_rows[(shade_index - 1) // 3].append(button)

        palette_groups.append(palette)

    for group in list(tab.findall(NS + "group")):
        if group.attrib.get("id") != "grpUtility":
            tab.remove(group)
    for insert_index, palette in enumerate(palette_groups):
        tab.insert(insert_index, palette)

    xml_path = OUTPUT / "ribbonx.xml"
    xml_path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))
    print(f"palette={xml_path}")
    print(f"groups={len(SOURCE_GROUPS)}")
    print(f"shades_per_color={len(SHADE_NAMES)}")
    print(f"icons={len(list(icon_dir.glob('*.png')))}")


if __name__ == "__main__":
    main()
