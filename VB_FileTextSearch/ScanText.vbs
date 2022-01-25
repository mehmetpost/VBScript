'Mehmet Yigit - 2004
'Bu script bulunugu klasordeki dosyalar icinde _liste.txt icinde satirlar halinde belirtilen kelimeleri arar
've arama sonuclarini bir csv dosya icine yazar.

Option Explicit
On Error Resume Next
Dim v_fso , strDir, objFolder, colFiles, v_FileObj
Dim v_fn_liste, v_f_liste, v_fn_output, v_f_output
Dim v_fn_source, v_f_source, v_aranan, v_konum, v_satir, v_bulundu 

v_fn_liste = "_liste.txt"
v_fn_output = "_result.csv"

Set v_fso = CreateObject("Scripting.FileSystemObject")
Set v_f_liste = v_fso.OpenTextFile(v_fn_liste, 1, False)

If Err.Number <> 0 Then
  MsgBox v_fn_liste & " dosyasý bulunamadý!"
Else
  strDir = v_fso.GetParentFolderName(WScript.ScriptFullName)
  Set objFolder = v_fso.GetFolder(strDir)
  Set colFiles = objFolder.Files

  Set v_f_output = v_fso.CreateTextFile(v_fn_output, True)
  'dosyadaki her bir satiri
  do while v_f_liste.AtEndOfStream <> True
    v_aranan = v_f_liste.ReadLine
    'calisma dizinindeki her bir dosyanin icini tara
    For Each v_FileObj in colFiles
      v_bulundu = false
      v_fn_source = v_FileObj.Name
      if (v_fn_source <> v_fn_liste) and (v_fn_source <> v_fn_output) then
        Set v_f_source = v_fso.OpenTextFile(v_fn_source, 1, False)
        do while v_f_source.AtEndOfStream <> True
          v_satir = v_f_source.ReadLine
          v_konum = instr(v_satir,v_aranan)
          if v_konum > 0 then
            v_bulundu = true
            exit do
          end if
        loop
        v_f_source.close
        if v_bulundu = true then
          v_f_output.WriteLine v_aranan & ";" & v_fn_source
        end if
      end if
    Next
  loop
  v_f_liste.Close
  v_f_output.Close
End If 


