Option Explicit

' The last used swatch is kept in the per-user VBA settings store so it
' survives Word restarts.
Private Const SETTINGS_APP As String = "PaletteHighlighterForWord"
Private Const SETTINGS_SECTION As String = "LastUsed"

Private highlighterRibbon As IRibbonUI

' Word resolves Ribbon callbacks by name across every loaded project,
' including Normal.dotm, so all callbacks carry a prefix unique to this add-in.

Public Sub PaletteHighlighter_Load(ribbon As IRibbonUI)
    Set highlighterRibbon = ribbon
End Sub

' Generic Ribbon callback for all 81 swatches.
' Each RibbonX toggle button stores only its RGB triplet in control.Tag.
Public Sub PaletteHighlighter_Color(control As IRibbonControl, pressed As Boolean)
    If ApplyTag(control.Tag) Then
        SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Id", control.Id
    End If

    ' Clicking a toggle flips its pressed state, so always resync the Ribbon.
    If Not highlighterRibbon Is Nothing Then highlighterRibbon.Invalidate
End Sub

Public Sub PaletteHighlighter_GetPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (control.Id = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Id", ""))
End Sub

Public Sub PaletteHighlighter_Remove(control As IRibbonControl)
    RemoveCustomHighlight
End Sub

Private Function ApplyTag(ByVal tag As String) As Boolean
    Dim r As Long, g As Long, b As Long

    If Selection.Type = wdSelectionIP Then
        MsgBox "Select some text first.", vbInformation, "Palette Highlighter"
        Exit Function
    End If

    If Not ParseTag(tag, r, g, b) Then Exit Function

    ApplyCustomHighlight r, g, b
    ApplyTag = True
End Function

Private Function ParseTag(ByVal tag As String, ByRef r As Long, ByRef g As Long, ByRef b As Long) As Boolean
    Dim parts() As String

    parts = Split(tag, ",")
    If UBound(parts) <> 2 Then Exit Function

    r = CLng(parts(0))
    g = CLng(parts(1))
    b = CLng(parts(2))
    ParseTag = True
End Function

Private Sub ApplyCustomHighlight(ByVal r As Long, ByVal g As Long, ByVal b As Long)
    Dim foregroundColor As Long

    foregroundColor = BestForegroundColor(r, g, b)

    With Selection.Font.Shading
        .Texture = wdTextureNone
        .ForegroundPatternColor = wdColorAutomatic
        .BackgroundPatternColor = RGB(r, g, b)
    End With

    Selection.Font.Color = foregroundColor
End Sub

Public Sub RemoveCustomHighlight()
    If Selection.Type = wdSelectionIP Then
        MsgBox "Select some text first.", vbInformation, "Palette Highlighter"
        Exit Sub
    End If

    With Selection.Font.Shading
        .Texture = wdTextureNone
        .ForegroundPatternColor = wdColorAutomatic
        .BackgroundPatternColor = wdColorAutomatic
    End With

    Selection.Font.Color = wdColorAutomatic
End Sub

Private Function LinearSRGB(ByVal channel As Long) As Double
    Dim s As Double
    s = channel / 255#

    If s <= 0.04045 Then
        LinearSRGB = s / 12.92
    Else
        LinearSRGB = ((s + 0.055) / 1.055) ^ 2.4
    End If
End Function

Private Function RelativeLuminance(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Double
    RelativeLuminance = _
        0.2126 * LinearSRGB(r) + _
        0.7152 * LinearSRGB(g) + _
        0.0722 * LinearSRGB(b)
End Function

Private Function ContrastWithBlack(ByVal luminance As Double) As Double
    ContrastWithBlack = (luminance + 0.05) / 0.05
End Function

Private Function ContrastWithWhite(ByVal luminance As Double) As Double
    ContrastWithWhite = 1.05 / (luminance + 0.05)
End Function

Private Function BestForegroundColor(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    Dim luminance As Double
    Dim blackContrast As Double
    Dim whiteContrast As Double

    luminance = RelativeLuminance(r, g, b)
    blackContrast = ContrastWithBlack(luminance)
    whiteContrast = ContrastWithWhite(luminance)

    If blackContrast >= whiteContrast Then
        BestForegroundColor = wdColorBlack
    Else
        BestForegroundColor = wdColorWhite
    End If
End Function
