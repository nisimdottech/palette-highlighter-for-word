Option Explicit

' The last used swatch is kept in the per-user VBA settings store so it
' survives Word restarts.
Private Const SETTINGS_APP As String = "PaletteHighlighterForWord"
Private Const SETTINGS_SECTION As String = "LastUsed"
Private Const SWATCH_SIZE As Long = 32

Private highlighterRibbon As IRibbonUI

Public Sub Ribbon_Load(ribbon As IRibbonUI)
    Set highlighterRibbon = ribbon
End Sub

' Generic Ribbon callback for all 81 swatches.
' Each RibbonX toggle button stores only its RGB triplet in control.Tag.
Public Sub Ribbon_Color(control As IRibbonControl, pressed As Boolean)
    If ApplyTag(control.Tag) Then
        SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Id", control.Id
        SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Tag", control.Tag
    End If

    ' Clicking a toggle flips its pressed state, so always resync the Ribbon.
    If Not highlighterRibbon Is Nothing Then highlighterRibbon.Invalidate
End Sub

Public Sub Ribbon_GetPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (control.Id = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Id", ""))
End Sub

Public Sub Ribbon_ApplyLast(control As IRibbonControl)
    ApplyTag GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Tag", "")
End Sub

Public Sub Ribbon_GetLastEnabled(control As IRibbonControl, ByRef returnedVal)
    Dim r As Long, g As Long, b As Long

    returnedVal = ParseTag(GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Tag", ""), r, g, b)
End Sub

Public Sub Ribbon_GetLastImage(control As IRibbonControl, ByRef returnedVal)
    Dim r As Long, g As Long, b As Long

    If Not ParseTag(GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Tag", ""), r, g, b) Then
        r = 255
        g = 255
        b = 255
    End If

    Set returnedVal = SwatchPicture(r, g, b)
End Sub

Public Sub Ribbon_Remove(control As IRibbonControl)
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

' Ribbon images must be IPictureDisp objects, and LoadPicture only reads
' files, so the swatch is written as a temporary 24-bit BMP with a gray border.
Private Function SwatchPicture(ByVal r As Long, ByVal g As Long, ByVal b As Long) As IPictureDisp
    Const HEADER_SIZE As Long = 54
    Dim data() As Byte
    Dim imageSize As Long
    Dim x As Long, y As Long, offset As Long
    Dim path As String
    Dim fileNumber As Integer

    imageSize = SWATCH_SIZE * SWATCH_SIZE * 3
    ReDim data(0 To HEADER_SIZE + imageSize - 1)

    data(0) = Asc("B")
    data(1) = Asc("M")
    PutLong data, 2, HEADER_SIZE + imageSize
    PutLong data, 10, HEADER_SIZE
    PutLong data, 14, 40
    PutLong data, 18, SWATCH_SIZE
    PutLong data, 22, SWATCH_SIZE
    data(26) = 1
    data(28) = 24
    PutLong data, 34, imageSize

    offset = HEADER_SIZE
    For y = 0 To SWATCH_SIZE - 1
        For x = 0 To SWATCH_SIZE - 1
            If x = 0 Or y = 0 Or x = SWATCH_SIZE - 1 Or y = SWATCH_SIZE - 1 Then
                data(offset) = 128
                data(offset + 1) = 128
                data(offset + 2) = 128
            Else
                data(offset) = b
                data(offset + 1) = g
                data(offset + 2) = r
            End If
            offset = offset + 3
        Next x
    Next y

    path = Environ$("TEMP") & "\PaletteHighlighterForWord_last.bmp"
    fileNumber = FreeFile
    Open path For Binary Access Write As #fileNumber
    Put #fileNumber, 1, data
    Close #fileNumber

    Set SwatchPicture = LoadPicture(path)
    Kill path
End Function

Private Sub PutLong(ByRef data() As Byte, ByVal offset As Long, ByVal value As Long)
    data(offset) = value And &HFF&
    data(offset + 1) = (value \ &H100&) And &HFF&
    data(offset + 2) = (value \ &H10000) And &HFF&
    data(offset + 3) = (value \ &H1000000) And &HFF&
End Sub

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
