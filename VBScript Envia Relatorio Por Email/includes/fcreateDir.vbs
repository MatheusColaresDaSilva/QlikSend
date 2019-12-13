Function inArray(arr, obj)
	On Error Resume Next
	Dim x: x = -1			
		
		For i = 0 To arr.Count-1
		  If arr.Item(i).Name&".txt" = obj Then
			  x = i
			  Exit For
		  End If
	  Next

	Call err_report()
	inArray = x

End Function

public Function createDir (ByVal document)

        strPasteName = Mid(document, InStrRev(document, "\") + 1)
	strPasteName = Replace(strPasteName,".qvw","-pst")
	strDir = sDirAcesso
	strDirQvw = strDir&strPasteName  

	Dim oFSO
	Set oFSO = CreateObject("Scripting.FileSystemObject")

'Criar pasta com nome da aplicação
	If Not oFSO.FolderExists(strDirQvw) Then
		oFSO.CreateFolder strDirQvw
	End If

	Set ri = document.GetDocReportInfo
    Set fldr = oFSO.GetFolder(strDirQvw) ''
    Set files = fldr.Files''

'Verifica se há algum arquivo que não há no Report e deleta-o se existir
	For each file in files     
   		If (inArray(ri, file.Name) = -1) Then
   		 oFSO.DeleteFile(strDirQvw&"\"&file.Name)
   		End If
    Next

'Criar arquivos com o nome dos Reports
  	For i = 0 to ri.Count-1 'for que vai de 0 até o número de relatório e salva e pdf cada um deles
	  set r = ri.Item(i)  'Pega o relatório
	  nameFile = strDirQvw&"\"&r.Name&".txt"

	  If Not oFSO.FileExists(nameFile) Then
	    Set objFile = oFSO.CreateTextFile(nameFile,True)
        objFile.Write ";"
	    objFile.Close		
	  End If
  	Next

End Function