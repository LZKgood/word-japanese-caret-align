Attribute VB_Name = "CaretDownAlignedModule"
Option Explicit

' ============================================================
' CaretDownAlignedFinal
' Vertical caret alignment for Japanese intensive-listening notes
'
' Shortcut: Alt+Down
'
' First press: creates a full-width-space grid line below the
' Japanese sentence and places the caret under the current word.
' Later presses: return to the same grid line at the new word's
' column. If there is text to the right, the macro enables
' Word's Overtype mode so existing annotations are not pushed away.
'
' Notes:
' - Delete old indent-style annotation lines created by pre-1.0
'   versions before using this version.
' - The annotation line contains full-width spaces, which are the
'   horizontal anchors. They are invisible unless formatting marks
'   are shown.
' - Normal single-column main text only. Tables, text boxes,
'   columns, and vertical writing are not supported.
' ============================================================
Sub CaretDownAlignedFinal()
    Dim sel As Object
    Dim doc As Object
    Dim curP As Object
    Dim nextP As Object
    Dim x As Single
    Dim srcSize As Single
    Dim srcName As String
    Dim curIdx As Long
    Dim i As Long

    Set sel = Selection
    Set doc = ActiveDocument

    ' Start in insert mode to avoid accidental overwriting in the
    ' Japanese sentence. The mode is set again after positioning.
    doc.Application.Options.Overtype = False

    If sel.Type <> 1 Then Beep: Exit Sub
    If sel.StoryType <> 1 Then Beep: Exit Sub
    If sel.Information(12) <> 0 Then Beep: Exit Sub

    ' Horizontal position of the caret relative to the text boundary,
    ' in points.
    x = sel.Information(7)
    If x < 0 Then Beep: Exit Sub

    ' Use the source font size so that the full-width-space grid has
    ' the same column width as the Japanese text.
    srcSize = sel.Font.Size
    If srcSize <= 0 Or srcSize > 200 Then srcSize = 10.5
    srcName = sel.Font.Name
    If srcName = "" Then srcName = "MS Mincho"

    Set curP = sel.Paragraphs(1)

    ' Word's Paragraph object has no Index property. Locate the
    ' paragraph's position in the Paragraphs collection by Range.Start.
    curIdx = 0
    For i = 1 To doc.Paragraphs.Count
        If doc.Paragraphs(i).Range.Start = curP.Range.Start Then
            curIdx = i
            Exit For
        End If
    Next i

    If curIdx = 0 Then
        Beep
        Exit Sub
    End If

    Set nextP = Nothing
    If curIdx < doc.Paragraphs.Count Then
        Set nextP = doc.Paragraphs(curIdx + 1)
    End If

    If nextP Is Nothing Then
        ' No following paragraph: add one after the current paragraph.
        curP.Range.InsertParagraphAfter
        Set nextP = doc.Paragraphs(curIdx + 1)
    ElseIf Not IsBlankPara(nextP) And Not IsGridText(nextP) Then
        ' The following paragraph is ordinary text. Do not modify it;
        ' insert a dedicated annotation line between the two paragraphs.
        If nextP.Range.ParagraphFormat.FirstLineIndent > srcSize * 1.5 Then
            MsgBox "An old indent-style annotation line was detected. Delete that line, then press Alt+Down again from the Japanese line.", vbInformation
            Exit Sub
        End If

        curP.Range.InsertParagraphAfter
        Set nextP = doc.Paragraphs(curIdx + 1)
    End If

    If IsBlankPara(nextP) Then
        MakeGridLine doc, nextP, srcName, srcSize
        MoveCaretToX doc, nextP, x, sel
        doc.Application.Options.Overtype = False
    ElseIf IsGridText(nextP) Then
        MoveCaretToX doc, nextP, x, sel
        SetOvertypeForGrid doc, nextP, sel
    Else
        MoveCaretToX doc, nextP, x, sel
    End If

    If doc.Application.Options.Overtype Then
        doc.Application.StatusBar = "CaretDownAligned: Overtype is ON (press Insert to return to insert mode)."
    Else
        doc.Application.StatusBar = "CaretDownAligned: caret aligned."
    End If
End Sub

' Build a full-width-space grid line.
Private Sub MakeGridLine(doc As Object, p As Object, srcName As String, srcSize As Single)
    Dim textWidth As Single
    Dim charWidth As Single
    Dim n As Long
    Dim spaces As String

    With p.Range.ParagraphFormat
        .Alignment = 0
        .LeftIndent = 0
        .FirstLineIndent = 0
        .RightIndent = 0
    End With

    If srcName <> "" Then p.Range.Font.Name = srcName
    p.Range.Font.Size = srcSize

    textWidth = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin
    charWidth = srcSize
    If textWidth <= 0 Or charWidth <= 0 Then Exit Sub

    n = Int(textWidth / charWidth) - 1
    If n < 1 Then n = 1
    If n > 200 Then n = 200

    spaces = String(n, ChrW(&H3000))
    p.Range.InsertBefore spaces

    If srcName <> "" Then p.Range.Font.Name = srcName
    p.Range.Font.Size = srcSize
End Sub

' Move the caret to the character boundary closest to x.
Private Sub MoveCaretToX(doc As Object, p As Object, x As Single, sel As Object)
    Dim i As Long
    Dim maxChars As Long
    Dim bestOffset As Long
    Dim bestDiff As Single
    Dim curX As Single
    Dim diff As Single
    Dim pStart As Long

    On Error GoTo Done

    p.Range.Select
    sel.Collapse 1
    pStart = p.Range.Start
    bestOffset = 0
    bestDiff = Abs(sel.Information(7) - x)
    maxChars = Len(p.Range.Text) - 1
    If maxChars < 1 Then GoTo Done

    For i = 1 To maxChars
        sel.MoveRight 1, 1
        If sel.Paragraphs(1).Range.Start <> pStart Then Exit For
        curX = sel.Information(7)
        diff = Abs(curX - x)
        If diff < bestDiff Then
            bestDiff = diff
            bestOffset = i
        End If
        If curX >= x Then Exit For
    Next i

Done:
    On Error GoTo 0
    p.Range.Select
    sel.Collapse 1
    If bestOffset > 0 Then sel.MoveRight 1, bestOffset
End Sub

' If non-space text exists to the right of the caret, enable Overtype
' so that typing replaces padding spaces instead of shifting text.
Private Sub SetOvertypeForGrid(doc As Object, p As Object, sel As Object)
    Dim tailText As String
    Dim r As Object

    On Error GoTo OffMode

    If sel.Start > p.Range.End - 1 Then GoTo OffMode

    Set r = doc.Range(sel.Start, p.Range.End - 1)
    tailText = r.Text
    tailText = Replace(tailText, ChrW(&H3000), "")
    tailText = Replace(tailText, " ", "")
    tailText = Replace(tailText, Chr(9), "")
    tailText = Replace(tailText, Chr(13), "")
    tailText = Replace(tailText, Chr(10), "")

    If Len(tailText) > 0 Then
        doc.Application.Options.Overtype = True
    Else
        doc.Application.Options.Overtype = False
    End If
    Exit Sub

OffMode:
    doc.Application.Options.Overtype = False
End Sub

Private Function IsBlankPara(p As Object) As Boolean
    Dim t As String
    On Error Resume Next
    t = p.Range.Text
    t = Replace(t, Chr(13), "")
    t = Replace(t, Chr(10), "")
    t = Replace(t, Chr(7), "")
    IsBlankPara = (Len(t) = 0)
End Function

' True when the paragraph is an annotation grid line created by this
' macro. Grid lines contain many full-width spaces.
Private Function IsGridText(p As Object) As Boolean
    Dim t As String
    Dim i As Long
    Dim n As Long
    On Error Resume Next
    t = p.Range.Text
    For i = 1 To Len(t)
        If AscW(Mid$(t, i, 1)) = &H3000 Then n = n + 1
    Next i
    IsGridText = (n >= 3)
End Function
