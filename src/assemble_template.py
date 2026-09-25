from __future__ import annotations

import subprocess
from pathlib import Path
from zipfile import ZipFile

from make_palette12 import main as make_palette12


ROOT = Path(__file__).resolve().parents[1]
BUILD = ROOT / "build"
PROJECT_SOURCE = ROOT / "src"
BASE_TEMPLATE = BUILD / "base_template.dotm"
DOTM = BUILD / "PaletteHighlighterForWord.dotm"


def ribbon_editor() -> Path:
    candidates = (
        Path(r"C:\Program Files\Office RibbonX Editor\OfficeRibbonXEditor.CommandLine.exe"),
        Path(r"C:\Program Files (x86)\Office RibbonX Editor\OfficeRibbonXEditor.CommandLine.exe"),
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise FileNotFoundError("Office RibbonX Editor command-line tool was not found")


def main() -> None:
    if not BASE_TEMPLATE.exists():
        raise FileNotFoundError(f"Missing {BASE_TEMPLATE}; run src\\embed_vba.ps1 first")

    make_palette12()
    palette_source = BUILD / "palette12"
    ribbon_xml = palette_source / "ribbonx.xml"
    icons = sorted((palette_source / "icons").glob("*.png"))
    if not ribbon_xml.is_file() or not icons:
        raise FileNotFoundError("RibbonX XML or icon sources are missing")

    dotm = DOTM
    if dotm.exists():
        dotm.unlink()
    subprocess.run(
        [
            str(ribbon_editor()),
            "insert",
            "--type",
            "14",
            "--xml",
            str(ribbon_xml),
            "--icons",
            str(palette_source / "icons"),
            "--output",
            str(dotm),
            str(BASE_TEMPLATE),
        ],
        check=True,
        cwd=ROOT,
    )

    with ZipFile(dotm, "r") as package:
        names = set(package.namelist())
        if package.testzip() is not None:
            raise ValueError("The generated template contains a corrupt ZIP member")
        if package.read("customUI/customUI14.xml") != ribbon_xml.read_bytes():
            raise ValueError("Replanned RibbonX XML was not embedded byte-for-byte")
        embedded_icons = [name for name in names if name.startswith("customUI/images/")]
        if len(embedded_icons) != len(icons):
            raise ValueError(f"Expected {len(icons)} icons, found {len(embedded_icons)}")

    print(f"template={dotm}")
    print(f"icons={len(icons)}")
    print(f"size={dotm.stat().st_size}")


if __name__ == "__main__":
    main()
