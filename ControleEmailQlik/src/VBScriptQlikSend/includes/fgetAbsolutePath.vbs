function getAbsolutePath(ByVal filePath)
  if Mid(filePath,2,1) = ":" OR Left(filePath,2) = "\\" then 'Absolute path in input parameter'
    getAbsolutePath = filePath
  else
    dim fso: set fso = CreateObject("Scripting.FileSystemObject")
    getAbsolutePath = fso.BuildPath(fso.GetParentFolderName(Wscript.ScriptFullName), filePath)
  end if
end function