Sub include(ByVal fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

sDirInclude = ".\includes\"
sDirAplicacao = "..\..\..\Aplicacao\"
sDirAcesso = ".\acessos\"
sDirSaveReport = "C:\Users\mcs.silva\Desktop\matheus\QlikSend\QlikSend\Reports\"


include sDirInclude&"header.vbs"


include sDirInclude&"footer.vbs"