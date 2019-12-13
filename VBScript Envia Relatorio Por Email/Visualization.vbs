Sub include(ByVal fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

sDirInclude = "includes\"
sDirAplicacao = "aplicacao\"
sDirAcesso = "acessos\"
sDirSaveReport = "C:\Users\mcs.silva\Desktop\matheus\SendEmailReportQlik\VBScript Envia Relatorio Por Email\reports\" 'here must be a absolutepath


include sDirInclude&"header.vbs"


qlik.doc.ClearAll true 'Limpa todos filtros
qlik.doc.Fields("Ano").Select qlik.doc.Evaluate("(Year(Today()))") 'Ano atual 

include sDirInclude&"footer.vbs"