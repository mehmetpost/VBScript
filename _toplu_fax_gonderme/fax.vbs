Option Explicit
On Error Resume Next

Dim FileName, Find, ReplaceWith, FileContents

Dim v_fso , strDir, objFolder, colFiles, v_FileObj
Dim v_fn_liste, v_f_liste, v_fn_sablon, v_f_output
Dim v_fn_source, v_f_source, v_aranan, v_konum, v_satir, v_bulundu 

v_fn_liste = "_liste.txt"
v_fn_sablon = "_sablon.txt"

Set v_fso = CreateObject("Scripting.FileSystemObject")
Set v_f_liste = v_fso.OpenTextFile(v_fn_liste, 1, False)

If Err.Number <> 0 Then
  MsgBox v_fn_liste & " dosyasý bulunamadý!" & Err.Number
Else
  strDir = v_fso.GetParentFolderName(WScript.ScriptFullName)
  'dosyadaki her bir satiri
  do while v_f_liste.AtEndOfStream <> True
    Dim  tokens
    v_aranan = v_f_liste.ReadLine
    tokens = Split(v_aranan , ";")
    'For i=0 To UBound(tokens), tokens(i)

    FileContents = GetFile(v_fn_sablon)
    FileContents = replace(FileContents, "*fax*", tokens(2), 1, -1, 1)
    FileContents = replace(FileContents, "*unvan*", tokens(1), 1, -1, 1)
    WriteFile tokens(0) + ".txt", FileContents
  loop
  v_f_liste.Close
End If 


'Read text file
function GetFile(FileName)
  If FileName<>"" Then
    Dim FS, FileStream
    Set FS = CreateObject("Scripting.FileSystemObject")
      on error resume Next
      Set FileStream = FS.OpenTextFile(FileName)
      GetFile = FileStream.ReadAll
  End If
End Function

'Write string As a text file.
function WriteFile(FileName, Contents)
  Dim OutStream, FS
  on error resume Next
  Set FS = CreateObject("Scripting.FileSystemObject")
    Set OutStream = FS.OpenTextFile(FileName, 2, True)
    OutStream.Write Contents
End Function

 
