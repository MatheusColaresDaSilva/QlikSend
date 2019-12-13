public Function printSend (ByVal document, ByVal sDir, ByVal strCurDate, ByVal NameQVW, ByVal srchRight)
  Set ri = document.GetDocReportInfo

  for i = 0 to ri.Count-1 'for que vai de 0 até o número de relatório e salva e pdf cada um deles
      set r = ri.Item(i)  'Pega o relatório
      call printPDF (document ,sDir, strCurDate, NameQVW, r.Name,r.Id) 'Salva em pdf    
  next    
   
  call sendEmail (document ,sDir, strCurDate, NameQVW, srchRight)

End Function