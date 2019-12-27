function getAllEmails(ByVal document, ByRef emails, ByRef arrayAllEmail, ByRef arrayAllEmailDistinct, ByRef arrayEmailReport )

Set ri = document.GetDocReportInfo
 strPasteName = Mid(document, InStrRev(document, "\") + 1)
 strPasteName = Replace(strPasteName,".qvw","-pst")
 strDir = sDirAcesso
 strDirQvw = strDir&strPasteName

 for i = 0 to ri.Count-1
  set r = ri.Item(i)
  nameFile = strDirQvw&"\"&r.Name&".txt"
  Dim oFSO
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  
  If oFSO.FileExists(nameFile) Then
    Set objFile = oFSO.OpenTextFile(nameFile)
    Do Until objFile.AtEndOfStream
     arrayEmailReport= objFile.ReadLine
    Loop
    objFile.Close
  End IF
  
  arrayEmailReport = Replace(arrayEmailReport, chr(10), "") 'Retira quebra de linha
  emails = emails&arrayEmailReport
 next
 
 arrayAllEmail = Split(emails,";")  'Transforma em um Array de String

 call removerDuplicados (arrayAllEmail, arrayAllEmailDistinct)  'Remove Emails Duplicados

end function

'Função que remove email duplicado do subprograma  getAllEmails 
Function removerDuplicados(ByRef arrayAllEmail, ByRef arrayAllEmailDistinct)

Redim arrayAllEmailDistinct(-1) 'Torna o array secundario não nulo

For i = 0 To UBound(arrayAllEmail) 'para cada elemento do array de emails
  if UBound(Filter(arrayAllEmailDistinct, arrayAllEmail(i))) = -1 Then ' verifica se o email pertence ao array secundario
    Redim preserve arrayAllEmailDistinct(Ubound(arrayAllEmailDistinct) + 1) 'Se sim, cria nova posição no array 
    arrayAllEmailDistinct(Ubound(arrayAllEmailDistinct)) = arrayAllEmail(i) 'e adiciona o email
  End If
Next 

arrayAllEmail = arrayAllEmailDistinct ' Iguala o array principal ao secundario

End Function

'Subprograma que manda e-mail para cada email cadastrado
function sendEmail(ByVal document, ByVal sDir, ByVal strCurDate, ByVal NameQVW, ByVal srchRight)
  
  Dim emails
  Dim arrayAllEmail
  Dim arrayAllEmailDistinct
  Dim arrayEmailReport

  Call getAllEmails (document, emails, arrayAllEmail, arrayAllEmailDistinct, arrayEmailReport ) 'Cria uma lista com todos os emails que receberam os arquivos
  
 For each email in arrayAllEmail

     If InStrRev(email, "|") > 0 Then
     
      configSend = Mid(email, InStrRev(email, "|") + 1)
      email = Left(email, InStrRev(email, "|") - 1)
 
     End If

     'Variaveis para acesso ao diretório
     Set fso = CreateObject("Scripting.FileSystemObject")     
     Set fldr = fso.GetFolder(sDir)
     Set files = fldr.Files
    
     'Variaveis para criação de objeto de email
     Set objMsg = CreateObject("CDO.Message") 
     Set msgConf = CreateObject("CDO.Configuration") 
    
     'Configuração Serve SMTP
     msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
     msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 
     msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.santacasamaringa.com.br" 
     msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
     msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "ti.qlikview@santacasamaringa.com.br" 
     msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "indica@2020" 
     msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 0 
     msgConf.Fields.Update  
  
     objMsg.To = email
     objMsg.From = "ti.qlikview@santacasamaringa.com.br"
     objMsg.Subject = "Relatórios extraidos do Qlikview" 
     objMsg.HTMLBody = "Email Qlikview Indicadores"
     objMsg.Sender = "Business Intelligence" 
     
      
      Set ri = document.GetDocReportInfo       
      strPasteName = Mid(document, InStrRev(document, "\") + 1)
      strPasteName = Replace(strPasteName,".qvw","-pst")
      strDir = sDirAcesso
      strDirQvw = strDir&strPasteName 

      For i = 0 to ri.Count-1
        set r = ri.Item(i)
        nameFile = strDirQvw&"\"&r.Name&".txt"

        Dim oFSO
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        If oFSO.FileExists(nameFile) Then
          Set objFile = oFSO.OpenTextFile(nameFile)
          Do Until objFile.AtEndOfStream
            arrayEmailReport= objFile.ReadLine
          Loop
          objFile.Close

            arrayEmailReport = Replace(arrayEmailReport, chr(10), "") 'Retira quebra de linha
            arrayEmailReport = Replace(arrayEmailReport, "|", "") 'Retira quebra de linha
            arrayEmailReport = Replace(arrayEmailReport, "D", "") 'Retira quebra de linha
            arrayEmailReport = Replace(arrayEmailReport, "S", "") 'Retira quebra de linha
            arrayEmailReport = Replace(arrayEmailReport, "Q", "") 'Retira quebra de linha
            arrayEmailReport = Replace(arrayEmailReport, "M", "") 'Retira quebra de linha
            arrayEmailReport = Split(arrayEmailReport,";")
          
            If UBound(Filter(arrayEmailReport, email)) >= 0 Then 'Verifica se o email atual faz parte do escopo dos emails do relatório
            'Se sim, acha o arquivo gerado anteriormente com o nome corrensponde ao dia_app_relatorio.pdf em anexa no email
             For each file in files 
                If file.Name = strCurDate&"_"&NameQVW&"_"&r.Name&srchRight Then
                  objMsg.AddAttachment   sDir&file.Name
                End if
              Next
            End If

        End If
        
      Next
    
    Set objMsg.Configuration = msgConf    
      ' Enviar 
    On error resume next

    if configSend = "D" Then
      objMsg.Send
    End If

    if configSend = "S" and Weekday(Date,0) = 1 Then ' Se for domingo
      objMsg.Send
    End If

    if configSend = "Q" and FormatDateTime(Date,2) = FormatDateTime("16/" & Month(date) & "/" & Year(Date),2) Then
      objMsg.Send
    End If

    if configSend = "M" and FormatDateTime(Date,2) = FormatDateTime("04/" & Month(date) & "/" & Year(Date),2) Then
      objMsg.Send
    End If

    If Err <> 0 Then ' Se Email não existe
     On error goto 0
      'msgbox("Email Inválido")
    Else
     On error goto 0
      'msgbox("Email send ok")
    End if
      
    ' Limpar 
      Set objMsg = nothing
      Set msgConf = nothing
    
  Next
  
end Function