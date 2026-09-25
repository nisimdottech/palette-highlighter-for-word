; Inno Setup reference script. The current build uses the uv/PyInstaller
; fallback in build_installer.bat because Inno Setup is not installed here.
[Setup]
AppId={{D213C2CE-85E4-4A8B-B624-3D7F77F52AC1}
AppName=Palette Highlighter for Word
AppVersion=1.0.0
AppPublisher=Nisim Levi
DefaultDirName={userappdata}\Microsoft\Word\STARTUP
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=PaletteHighlighterForWord_Inno
Compression=lzma2
SolidCompression=yes
Uninstallable=yes
WizardStyle=modern

[Files]
Source: "..\build\PaletteHighlighterForWord.dotm"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; Flags: ignoreversion

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\PaletteHighlighterForWord.dotm"

[Messages]
FinishedLabel=Palette Highlighter for Word has been installed. Restart Microsoft Word to load it.
