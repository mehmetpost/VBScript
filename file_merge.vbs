Dim Separator
Dim fso,strDir ,objFolder, colFiles, v_fn_output
Dim fsSourceStream
Dim fsResStream
Dim sSeparator
Dim i
    
On Error Resume Next

v_fn_output = "_result.txt"
Separator = "################################ DOSYA ADI : #FilePath# ###########################"
Set fso  = CreateObject("Scripting.FileSystemObject")
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
Set objFolder = fso.GetFolder(strDir)
Set colFiles = objFolder.Files
Set fsResStream = fso.CreateTextFile(v_fn_output, True)
    
    For Each v_FileObj in colFiles
      if v_FileObj.Name <> v_fn_output and v_FileObj.Name <> WScript.ScriptName then
        sSeparator = Replace(Separator, "#FilePath#", v_FileObj.Name)
        fsResStream.Write sSeparator & vbCrLf
        Set fsSourceStream = fso.OpenTextFile(v_FileObj.Name, 1, False)
        fsResStream.Write fsSourceStream.ReadAll & vbCrLf
      end if
    Next 
fsResStream.Close

