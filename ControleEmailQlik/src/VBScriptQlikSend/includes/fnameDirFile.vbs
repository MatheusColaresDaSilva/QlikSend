Function nameDirFile
		
	Set objShell = CreateObject("Wscript.Shell")
	strPath = Wscript.ScriptFullName

	Dim strFileName
	strFileName = Mid(strPath, InStrRev(strPath, "\") + 1)
	strFileName = Replace(strFileName,"vbs","qvw")
	strDirQvw = sDirAplicacao
	strDirQvw = strDirQvw&strFileName    
	nameDirFile = strDirQvw
    
End Function