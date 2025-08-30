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
  Dim buff : buff = ""
  Dim j
  For j = 1 To sl.Shapes.Count
    buff = buff & ExtractShapeText(sl.Shapes(j))
  Next
  ExtractSlideText = buff
End Function

Function ExtractShapeText(shp)
  Dim buff : buff = ""
  On Error Resume Next

  ' Groups
  If shp.Type = 6 Then ' msoGroup
    Dim k
    For k = 1 To shp.GroupItems.Count
      buff = buff & ExtractShapeText(shp.GroupItems(k))
    Next
    ExtractShapeText = buff
    Exit Function
  End If

  ' Tables
  If shp.HasTable Then
    Dim r, c, cellText
    For r = 1 To shp.Table.Rows.Count
      For c = 1 To shp.Table.Columns.Count
        If HasText(shp.Table.Cell(r, c).Shape) Then
          cellText = Trim(GetText(shp.Table.Cell(r, c).Shape))
          If Len(cellText) > 0 Then buff = buff & "- " & cellText & vbCrLf
        End If
      Next
    Next
  End If

  ' SmartArt
  Err.Clear
  If shp.HasSmartArt Then
    Dim node
    For Each node In shp.SmartArt.AllNodes
      On Error Resume Next
      If Not node Is Nothing Then
        If Not node.TextFrame2 Is Nothing Then
          If Len(Trim(node.TextFrame2.TextRange.Text)) > 0 Then
            buff = buff & "- " & Trim(node.TextFrame2.TextRange.Text) & vbCrLf
          End If
        End If
      End If
    Next
  End If

  ' Normal text frames / placeholders
  If HasText(shp) Then
    Dim t : t = Trim(GetText(shp))
    If Len(t) > 0 Then buff = buff & "- " & t & vbCrLf
  End If

  On Error GoTo 0
  ExtractShapeText = buff
End Function

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
