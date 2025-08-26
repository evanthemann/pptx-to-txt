' run_ppt_to_text_prompt.vbs
' WScript GUI wrapper:
' - prompts user to pick a PPT/PPTX (Open File dialog via PowerShell; InputBox fallback)
' - builds <SameName>.txt in the same folder
' - runs: cscript //nologo ppt_to_text.vbs "<input>" "<output>"
' Requires: PowerPoint installed; keep OneDrive files "Always keep on this device" if needed.

Option Explicit

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh  : Set sh  = CreateObject("WScript.Shell")

' Locate converter (same folder as this wrapper)
Dim wrapperDir : wrapperDir = fso.GetParentFolderName(WScript.ScriptFullName)
Dim converter  : converter  = fso.BuildPath(wrapperDir, "ppt_to_text.vbs")

If Not fso.FileExists(converter) Then
  converter = InputBox("Can't find 'ppt_to_text.vbs' next to this wrapper." & vbCrLf & _
                       "Paste the full path to your working converter script:", "Locate converter")
  converter = TrimQuotes(Trim(converter))
  If Len(converter) = 0 Or Not fso.FileExists(converter) Then
    MsgBox "Converter script not found. Exiting.", vbExclamation, "ppt_to_text"
    WScript.Quit 1
  End If
End If

' Ask user for a PPT/PPTX
Dim inPath : inPath = ChoosePptFile()
inPath = TrimQuotes(Trim(inPath))
If Len(inPath) = 0 Then
  MsgBox "No file selected. Exiting.", vbExclamation, "ppt_to_text"
  WScript.Quit 1
End If

On Error Resume Next
inPath = fso.GetAbsolutePathName(inPath)
On Error GoTo 0

If Not fso.FileExists(inPath) Then
  MsgBox "File not found:" & vbCrLf & inPath, vbExclamation, "ppt_to_text"
  WScript.Quit 1
End If

Dim ext : ext = LCase(fso.GetExtensionName(inPath))
If ext <> "pptx" And ext <> "ppt" And ext <> "pptm" Then
  MsgBox "Unsupported file type: ." & ext & vbCrLf & "Use a .pptx, .pptm, or .ppt file.", vbExclamation, "ppt_to_text"
  WScript.Quit 1
End If

' Build output .txt in the same folder
Dim outPath : outPath = fso.BuildPath(fso.GetParentFolderName(inPath), fso.GetBaseName(inPath) & ".txt")

Dim confirmMsg
confirmMsg = "Input : " & inPath & vbCrLf & _
             "Output: " & outPath & vbCrLf & vbCrLf & _
             "Proceed?"
If MsgBox(confirmMsg, vbQuestion + vbOKCancel, "ppt_to_text") = vbCancel Then WScript.Quit 0

' Run the converter via cscript (hidden console; wait)
Dim cmd : cmd = "cscript.exe //nologo " & Q(converter) & " " & Q(inPath) & " " & Q(outPath)
Dim rc  : rc = sh.Run(cmd, 0, True)

If rc = 0 And fso.FileExists(outPath) Then
  If MsgBox("Success!" & vbCrLf & "Wrote: " & outPath & vbCrLf & vbCrLf & "Open folder now?", vbInformation + vbYesNo, "ppt_to_text") = vbYes Then
    sh.Run "explorer.exe /select," & Q(outPath), 1, False
	
  End If
  WScript.Quit 0
Else
  MsgBox "The converter exited with code " & rc & "." & vbCrLf & _
         "If the file is in OneDrive, try 'Always keep on this device' or open once and click 'Enable Editing'.", _
         vbExclamation, "ppt_to_text"
  WScript.Quit rc
End If

' ===== Helpers =====
Function Q(s) : Q = """" & s & """" : End Function

Function TrimQuotes(s)
  If Len(s) >= 2 Then
    If (Left(s,1) = """" And Right(s,1) = """") Or (Left(s,1) = "'" And Right(s,1) = "'") Then
      TrimQuotes = Mid(s, 2, Len(s)-2)
      Exit Function
    End If
  End If
  TrimQuotes = s
End Function

Function ChoosePptFile()
  ' Try a real Open File dialog via PowerShell/.NET; fall back to InputBox.
  Dim psCmd, execPS, sel
  psCmd = "powershell -NoProfile -Command " & Q( _
      "Add-Type -AssemblyName System.Windows.Forms; " & _
      "$f = New-Object System.Windows.Forms.OpenFileDialog; " & _
      "$f.Filter = 'PowerPoint (*.ppt;*.pptx;*.pptm)|*.ppt;*.pptx;*.pptm|All files (*.*)|*.*'; " & _
      "$f.Multiselect = $false; " & _
      "if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {[Console]::Write($f.FileName)}" )
  On Error Resume Next
  Set execPS = sh.Exec(psCmd)
  If Err.Number = 0 Then
    Do While execPS.Status = 0
      WScript.Sleep 50
    Loop
    sel = ""
    If Not execPS.StdOut.AtEndOfStream Then sel = execPS.StdOut.ReadAll
  End If
  On Error GoTo 0

  sel = Trim(sel)
  If Len(sel) > 0 Then
    ChoosePptFile = sel
  Else
    ChoosePptFile = TrimQuotes(Trim(InputBox("Enter or paste the full path to a PowerPoint file (.pptx/.ppt/.pptm):", "Select PowerPoint")))
  End If
End Function
