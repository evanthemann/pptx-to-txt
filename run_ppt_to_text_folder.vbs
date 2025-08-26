' run_ppt_to_text_folder.vbs
' Batch GUI wrapper:
' - asks for a folder
' - loops .pptx/.pptm/.ppt files
' - calls: cscript //nologo ppt_to_text.vbs "<input>" "<samefolder>\<samebasename>.txt"
' - optional overwrite of existing .txt files
' - writes a summary log in the selected folder

Option Explicit

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim sh  : Set sh  = CreateObject("WScript.Shell")

' --- locate the converter next to this wrapper (or ask for it) ---
Dim wrapperDir : wrapperDir = fso.GetParentFolderName(WScript.ScriptFullName)
Dim converter  : converter  = fso.BuildPath(wrapperDir, "ppt_to_text.vbs")
If Not fso.FileExists(converter) Then
  converter = InputBox( _
    "Can't find 'ppt_to_text.vbs' next to this wrapper." & vbCrLf & _
    "Paste the full path to your working converter script:", _
    "Locate converter")
  converter = TrimQuotes(Trim(converter))
  If Len(converter) = 0 Or Not fso.FileExists(converter) Then
    MsgBox "Converter script not found. Exiting.", vbExclamation, "ppt_to_text batch"
    WScript.Quit 1
  End If
End If

' --- pick a folder ---
Dim folderPath : folderPath = ChooseFolder()
folderPath = TrimQuotes(Trim(folderPath))
If Len(folderPath) = 0 Then
  MsgBox "No folder selected. Exiting.", vbExclamation, "ppt_to_text batch"
  WScript.Quit 1
End If
If Not fso.FolderExists(folderPath) Then
  MsgBox "Folder not found:" & vbCrLf & folderPath, vbExclamation, "ppt_to_text batch"
  WScript.Quit 1
End If

' --- overwrite choice ---
Dim overwrite, ans
ans = MsgBox("Overwrite existing .txt files if they already exist?", vbQuestion + vbYesNoCancel, "ppt_to_text batch")
If ans = vbCancel Then WScript.Quit 0
overwrite = (ans = vbYes)

' --- collect files ---
Dim folder : Set folder = fso.GetFolder(folderPath)
Dim exts : exts = Array("pptx","pptm","ppt")
Dim files : Set files = CreateObject("Scripting.Dictionary")

Dim file, ext
For Each file In folder.Files
  ext = LCase(fso.GetExtensionName(file.Name))
  If IsInArray(ext, exts) Then
    files.Add file.Path, True
  End If
Next

If files.Count = 0 Then
  MsgBox "No .pptx/.pptm/.ppt files found in:" & vbCrLf & folderPath, vbInformation, "ppt_to_text batch"
  WScript.Quit 0
End If

' --- run batch ---
Dim total : total = files.Count
Dim processed : processed = 0
Dim succeeded : succeeded = 0
Dim failed : failed = 0
Dim skipped : skipped = 0

Dim results : results = "ppt_to_text batch" & vbCrLf & String(40, "-") & vbCrLf & _
                        "Folder: " & folderPath & vbCrLf & _
                        "Converter: " & converter & vbCrLf & vbCrLf

Dim inPath, outPath, rc
For Each inPath In files.Keys
  outPath = fso.BuildPath(fso.GetParentFolderName(inPath), fso.GetBaseName(inPath) & ".txt")

  If (Not overwrite) And fso.FileExists(outPath) Then
    results = results & "SKIP  - " & inPath & "  (exists: " & outPath & ")" & vbCrLf
    skipped = skipped + 1
  Else
    ' quick progress toast (non-blocking enough with 1s timeout)
    sh.Popup "Processing:" & vbCrLf & inPath, 1, "ppt_to_text batch", 64

    rc = sh.Run("cscript.exe //nologo " & Q(converter) & " " & Q(inPath) & " " & Q(outPath), 0, True)
    If rc = 0 And fso.FileExists(outPath) Then
      results = results & "OK    - " & inPath & "  →  " & outPath & vbCrLf
      succeeded = succeeded + 1
    Else
      results = results & "FAIL  - " & inPath & "  (exit " & rc & ")" & vbCrLf
      failed = failed + 1
    End If
  End If

  processed = processed + 1
Next

results = results & vbCrLf & "Done." & vbCrLf & _
          "Total: " & total & _
          " | OK: " & succeeded & _
          " | Skipped: " & skipped & _
          " | Failed: " & failed & vbCrLf

' --- write batch log in the folder ---
Dim logPath : logPath = fso.BuildPath(folderPath, "ppt_to_text_batch_log.txt")
Call WriteUtf8(logPath, results)

Dim summary
summary = "Processed " & total & " file(s)." & vbCrLf & _
          "OK: " & succeeded & "   Skipped: " & skipped & "   Failed: " & failed & vbCrLf & vbCrLf & _
          "Log: " & logPath & vbCrLf & vbCrLf & _
          "Open the folder now?"

If MsgBox(summary, vbInformation + vbYesNo, "ppt_to_text batch") = vbYes Then
  sh.Run "explorer.exe " & Q(folderPath), 1, False
End If

WScript.Quit 0

' ================= Helpers =================

Function IsInArray(val, arr)
  Dim i
  For i = 0 To UBound(arr)
    If LCase(val) = LCase(arr(i)) Then IsInArray = True : Exit Function
  Next
  IsInArray = False
End Function

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

Function ChooseFolder()
  ' Shell BrowseForFolder (no PowerShell required); fallback to InputBox
  On Error Resume Next
  Dim shell, objFolder
  Set shell = CreateObject("Shell.Application")
  Set objFolder = shell.BrowseForFolder(0, "Select folder containing PowerPoint files:", 0)
  If Err.Number = 0 And Not objFolder Is Nothing Then
    ChooseFolder = objFolder.Self.Path
    Exit Function
  End If
  On Error GoTo 0

  ChooseFolder = TrimQuotes(Trim(InputBox("Enter or paste the full path to a folder:", "Select folder")))
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
  On Error GoTo 0
End Function
