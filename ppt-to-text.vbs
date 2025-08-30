' Usage:
'   cscript //nologo ppt_to_text_v2.vbs "C:\path\slides.pptx" ["C:\path\output.txt"]

Option Explicit

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim args : Set args = WScript.Arguments
If args.Count < 1 Then
  WScript.Echo "Usage: cscript //nologo ppt_to_text_v2.vbs <input.pptx> [output.txt]"
  WScript.Quit 1
End If

Dim inPath : inPath = fso.GetAbsolutePathName(args(0))
If Not fso.FileExists(inPath) Then
  WScript.Echo "Input file not found: " & inPath
  WScript.Quit 1
End If

Dim outPath
If args.Count >= 2 Then
  outPath = fso.GetAbsolutePathName(args(1))
Else
  outPath = fso.BuildPath(fso.GetParentFolderName(inPath), fso.GetBaseName(inPath) & ".txt")
End If

Dim ppApp, pres
On Error Resume Next
Set ppApp = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Or (ppApp Is Nothing) Then
  WScript.Echo "Could not start PowerPoint. Is it installed?"
  WScript.Quit 1
End If
On Error GoTo 0

ppApp.Visible = True
' Try to keep prompts quiet (1 = ppAlertsNone)
On Error Resume Next
ppApp.DisplayAlerts = 1
On Error GoTo 0

' --- Robust open (handles Protected View / OneDrive quirks) ---
Set pres = OpenPresentation(ppApp, inPath)
If pres Is Nothing Then
  WScript.Echo "Failed to open presentation: " & inPath & " (PowerPoint/Protected View or OneDrive issue)"
  ppApp.Quit
  WScript.Quit 1
End If

Dim txt : txt = "File: " & inPath & vbCrLf & String(80, "-") & vbCrLf

Dim i
For i = 1 To pres.Slides.Count
  Dim slide : Set slide = pres.Slides(i)
  txt = txt & vbCrLf & "Slide " & i & vbCrLf & String(60, "=") & vbCrLf

  ' Title (if present)
  Dim titleText : titleText = ""
  On Error Resume Next
  If Not slide.Shapes.Title Is Nothing Then
    If HasText(slide.Shapes.Title) Then titleText = Trim(GetText(slide.Shapes.Title))
  End If
  On Error GoTo 0
  If Len(titleText) > 0 Then
    txt = txt & "[Title]" & vbCrLf & titleText & vbCrLf & vbCrLf
  End If

  ' Body / shapes
  Dim bodyText : bodyText = ExtractSlideText(slide)
  If Len(Trim(bodyText)) > 0 Then
    txt = txt & "[Slide Content]" & vbCrLf & bodyText & vbCrLf
  End If

  ' Notes
  Dim notesText : notesText = ExtractNotesText(slide)
  If Len(Trim(notesText)) > 0 Then
    txt = txt & "[Notes]" & vbCrLf & notesText & vbCrLf
  End If
Next

On Error Resume Next
pres.Close
ppApp.Quit
On Error GoTo 0

If WriteUtf8(outPath, txt) Then
  WScript.Echo "Wrote outline to: " & outPath
Else
  WScript.Echo "Failed to write output to: " & outPath
  WScript.Quit 1
End If

' ================= Helpers =================

Function OpenPresentation(app, path)
  Dim p
  On Error Resume Next

  ' 1) Normal open WITH a window (helps with Protected View)
  Err.Clear
  Set p = app.Presentations.Open(path, False, False, True) ' ReadOnly=False, Untitled=False, WithWindow=True
  If Err.Number = 0 Then
    Set OpenPresentation = p
    Exit Function
  End If

  ' 2) Try Protected View → Edit
  Dim pvw
  Err.Clear
  Set pvw = app.ProtectedViewWindows.Open(path, "")
  If Err.Number = 0 And Not pvw Is Nothing Then
    Err.Clear
    pvw.Edit ' exit protected view
    If Err.Number = 0 Then
      ' Find the now-opened presentation
      Dim j
      For j = 1 To app.Presentations.Count
        If LCase(app.Presentations(j).Name) = LCase(fso.GetFileName(path)) Then
          Set OpenPresentation = app.Presentations(j)
          Exit Function
        End If
      Next
    End If
  End If

  Set OpenPresentation = Nothing
End Function

Function ExtractSlideText(sl)
  Dim itemsY(), itemsX(), itemsText(), itemsCount
  ReDim itemsY(0) : ReDim itemsX(0) : ReDim itemsText(0)
  itemsCount = 0

  Dim j
  For j = 1 To sl.Shapes.Count
    CollectShape sl.Shapes(j), 0, 0, itemsY, itemsX, itemsText, itemsCount
  Next

  SortItems itemsY, itemsX, itemsText, itemsCount

  Dim buff : buff = ""
  Dim i
  For i = 0 To itemsCount - 1
    buff = buff & itemsText(i)
  Next
  ExtractSlideText = buff
End Function

Sub CollectShape(shp, offX, offY, itemsY, itemsX, itemsText, ByRef itemsCount)
  On Error Resume Next

  ' Groups: recurse with absolute offset
  If shp.Type = 6 Then ' msoGroup
    Dim k
    For k = 1 To shp.GroupItems.Count
      CollectShape shp.GroupItems(k), offX + shp.Left, offY + shp.Top, itemsY, itemsX, itemsText, itemsCount
    Next
    Exit Sub
  End If

  ' Tables: each cell becomes an item (row-major)
  If shp.HasTable Then
    Dim r, c, cellShape, t
    For r = 1 To shp.Table.Rows.Count
      For c = 1 To shp.Table.Columns.Count
        Set cellShape = shp.Table.Cell(r, c).Shape
        If HasText(cellShape) Then
          t = Trim(GetText(cellShape))
          If Len(t) > 0 Then
            AddItem itemsY, itemsX, itemsText, itemsCount, _
                    offY + shp.Top + r * 0.01, offX + shp.Left + c * 0.01, _
                    "- " & t & vbCrLf
          End If
        End If
      Next
    Next
  End If

  ' SmartArt: indent by node level (best-effort order)
  If shp.HasSmartArt Then
    Dim node, level, txt, seq
    seq = 0
    For Each node In shp.SmartArt.AllNodes
      If Not node Is Nothing Then
        If Not node.TextFrame2 Is Nothing Then
          txt = Trim(node.TextFrame2.TextRange.Text)
          If Len(txt) > 0 Then
            On Error Resume Next
            level = node.Level
            On Error GoTo 0
            If level < 1 Then level = 1
            AddItem itemsY, itemsX, itemsText, itemsCount, _
                    offY + shp.Top + seq * 0.001, offX + shp.Left, _
                    String((level-1)*2, " ") & "- " & txt & vbCrLf
            seq = seq + 1
          End If
        End If
      End If
    Next
  End If

  ' Normal text frames / placeholders: paragraph-by-paragraph with indent
  If HasText(shp) Then
    Dim tr, pc, p, para, s, lvl
    Set tr = shp.TextFrame.TextRange
    pc = tr.Paragraphs.Count
    For p = 1 To pc
      Set para = tr.Paragraphs(p)
      s = Trim(para.Text)
      If Len(s) > 0 Then
        On Error Resume Next
        lvl = para.ParagraphFormat.IndentLevel
        On Error GoTo 0
        If lvl < 1 Then lvl = 1
        AddItem itemsY, itemsX, itemsText, itemsCount, _
                offY + shp.Top + p * 0.001, offX + shp.Left, _
                String((lvl-1)*2, " ") & "- " & s & vbCrLf
      End If
    Next
  End If

  On Error GoTo 0
End Sub

Sub AddItem(ByRef itemsY, ByRef itemsX, ByRef itemsText, ByRef itemsCount, y, x, text)
  If itemsCount = 0 Then
    itemsY(0) = CDbl(y)
    itemsX(0) = CDbl(x)
    itemsText(0) = CStr(text)
    itemsCount = 1
  Else
    ReDim Preserve itemsY(itemsCount)
    ReDim Preserve itemsX(itemsCount)
    ReDim Preserve itemsText(itemsCount)
    itemsY(itemsCount) = CDbl(y)
    itemsX(itemsCount) = CDbl(x)
    itemsText(itemsCount) = CStr(text)
    itemsCount = itemsCount + 1
  End If
End Sub

Sub SortItems(ByRef itemsY, ByRef itemsX, ByRef itemsText, ByRef itemsCount)
  Dim i, j, ty, tx, tt
  For i = 0 To itemsCount - 2
    For j = i + 1 To itemsCount - 1
      If (itemsY(j) < itemsY(i)) Or _
         (itemsY(j) = itemsY(i) And itemsX(j) < itemsX(i)) Then
        ' swap y
        ty = itemsY(i) : itemsY(i) = itemsY(j) : itemsY(j) = ty
        ' swap x
        tx = itemsX(i) : itemsX(i) = itemsX(j) : itemsX(j) = tx
        ' swap text
        tt = itemsText(i) : itemsText(i) = itemsText(j) : itemsText(j) = tt
      End If
    Next
  Next
End Sub


Function HasText(shp)
  On Error Resume Next
  HasText = False
  If shp.HasTextFrame Then
    If shp.TextFrame.HasText Then HasText = True
  End If
  On Error GoTo 0
End Function

Function GetText(shp)
  On Error Resume Next
  GetText = ""
  If shp.HasTextFrame Then GetText = shp.TextFrame.TextRange.Text
  On Error GoTo 0
End Function

Function ExtractNotesText(sl)
  On Error Resume Next
  Dim buff : buff = ""
  Dim ns : Set ns = sl.NotesPage
  If Not ns Is Nothing Then
    Dim i
    For i = 1 To ns.Shapes.Count
      If HasText(ns.Shapes(i)) Then
        Dim t : t = Trim(GetText(ns.Shapes(i)))
        If Len(t) > 0 Then buff = buff & t & vbCrLf
      End If
    Next
  End If
  On Error GoTo 0
  ExtractNotesText = buff
End Function

Function WriteUtf8(path, content)
  On Error Resume Next
  Dim stm : Set stm = CreateObject("ADODB.Stream")
  stm.Type = 2                ' text
  stm.Charset = "utf-8"
  stm.Open
  stm.WriteText content
  stm.SaveToFile path, 2      ' adSaveCreateOverWrite
  stm.Close
  WriteUtf8 = (Err.Number = 0)
  On Error GoTo 0
End Function
