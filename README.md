# Palette Highlighter for Word

A Microsoft Word add-in that adds a **Highlighter** tab to the ribbon with 81
highlight colors: 8 hues plus gray, each in 9 shades from light to dark.

Word's built-in highlighter offers only a handful of saturated colors. This
add-in applies any palette color as text shading and automatically switches the
text to black or white, whichever has the higher contrast, so every shade stays
readable.

## Features

- **8 familiar hues plus gray:** Red, Orange, Yellow, Green, Teal, Blue,
  Purple, Pink, and Gray. The names are the most widely known English color
  words, and each hue is placed near the typical color for its name while
  keeping the swatches of different colors as far apart as possible. Each
  color's name appears under its swatches.
- **9 shades per color,** from Lightest to Darkest, at matching lightness across
  the 8 hues. The lightness range is chosen so that even the lightest and
  darkest shades of neighboring colors stay distinguishable. Gray uses its own
  ladder running all the way from white to black.
- **Last used color:** the swatch you used last stays pressed in the palette,
  so you can see which color you picked. The choice is remembered across Word
  sessions.
- **Remove:** clears the shading and restores the automatic text color.
- Hover over a swatch to see its name, hex code, and RGB value.

## Install

1. Close Microsoft Word.
2. Run `PaletteHighlighterForWord.exe`.
3. Start Word and open the **Highlighter** tab.

The installer copies `PaletteHighlighterForWord.dotm` into Word's STARTUP folder
(`%APPDATA%\Microsoft\Word\STARTUP`), so Word loads it automatically. It also
removes a previous `CustomHighlighter.dotm` installation (the add-in's former
name).

Any template the installer replaces is backed up to
`%LOCALAPPDATA%\PaletteHighlighterForWord\backups`, which keeps the three most
recent backups. Backups that earlier versions left in the STARTUP folder are
moved there.

To uninstall, use **Settings → Apps → Installed apps → Palette Highlighter for
Word**. If the installed template has been modified since installation, the
uninstaller leaves it in place.

## Notes

- Colors are applied as character shading (**Font → Shading**), not as Word's
  highlight. Word's "Find highlight" and highlight-based filters don't see them.
- Select some text before choosing a color; nothing is applied to an empty
  selection.

## Building from source

Requirements:

- Windows with desktop Microsoft Word
- [uv](https://docs.astral.sh/uv/)
- [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor/releases),
  installed in its default location under `Program Files`

Build:

```powershell
src\build_installer.bat
```

The output is `dist\PaletteHighlighterForWord.exe`.

The build has three steps:

1. `src\embed_vba.ps1` copies `src\base_template.dotm` into `build\` and uses
   a hidden Word instance to replace its macro module with `src\VBA.bas`. Word
   only allows this when **Trust access to the VBA project object model** is
   on, so the script enables that setting for the current user during the run
   and restores the previous value afterwards.
2. `src\assemble_template.py` generates the palette ribbon and icons
   from `src\ribbonx.xml` (via `src\make_palette12.py`) and inserts them into
   the template with the Office RibbonX Editor command-line tool.
3. PyInstaller bundles the template and `src\installer.py` into a single
   executable.

## Project layout

| Path | Purpose |
| --- | --- |
| `src/VBA.bas` | Ribbon callbacks: apply, remove, last used color |
| `src/ribbonx.xml` | Ribbon tab and the Remove Highlight button |
| `src/make_palette12.py` | Defines the 9 colors, computes their shades in OKLCH, and generates the color groups and icons |
| `src/base_template.dotm` | Word-authored template that holds the VBA project |
| `src/embed_vba.ps1` | Embeds `VBA.bas` into the template via Word |
| `src/assemble_template.py` | Builds the final `.dotm` |
| `src/installer.py` | Install and uninstall logic of the executable |
| `src/build_installer.bat` | Full build |
| `src/installer.iss` | Inno Setup reference script (not used by the build) |

## License

Licensed under the [Apache License, Version 2.0](LICENSE).
