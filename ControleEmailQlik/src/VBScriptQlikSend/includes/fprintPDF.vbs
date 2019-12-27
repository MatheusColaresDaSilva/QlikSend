function printPDF(ByVal document, ByVal saveDir, ByVal dateFile, ByVal nameApp, ByVal namePDF, ByVal qvObject)

      Set PDFCreator = CreateObject("PDFCreator.clsPDFCreator")
      PDFCreator.cStart ("/NoProcessingAtStartup")
      PDFCreator.cOption("UseAutosave") = 1 ' Permiti autosave
      PDFCreator.cOption("UseAutosaveDirectory") = 1 ' Especificar diretorio para salvar
      PDFCreator.cOption("AutosaveDirectory") =saveDir  ' Setar autosave e dizer diretório a ser usado
      PDFCreator.cOption("AutosaveFormat") = 0 '  0=PDF, 1=PNG, 2=JPG, 3=BMP, 4=PCX, 5=TIFF, 6=PS, 7= EPS, 8=ASCII
      PDFCreator.cOption("AutosaveFilename") = dateFile&"_"&nameApp &"_"&namePDF ' Nome do arquivo
      'Padrão arquivo: DATA_NOMEAPLICAÇÃO_NOMERELATORIO
      PDFCreator.cPrinterStop = False          'Falso para impressão
      document.PrintDocReport qvObject, "PDFCreator"    'Salvar pdf
      document.GetApplication.Sleep 5000    ' Paras a macro por 5seg pra dar tempo salva os pdf 
      PDFCreator.cClearCache()    'Limpar Cache
      PDFCreator.cClose() 'Fechar pdf

end Function