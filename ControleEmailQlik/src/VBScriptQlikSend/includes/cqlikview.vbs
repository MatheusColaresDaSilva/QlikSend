Class QlikView
  Private m_App
  Private m_Doc
  Private m_docName
  Private NameQVW 
  Private strCurDate
  Private srchRight
  Private sDir
 
  Private Sub Class_Initialize
    m_docName = ""
  End Sub

  Public Property Get app
    set app = m_App
  End Property

  Public Property Get doc
    set doc = m_Doc
  End Property

  Public Property Get docName
    docName = m_docName
  End Property

  public function setDocument(ByVal docName)
    m_docName = getAbsolutePath(docName)
  end function

  Public function Quit
      m_App.Quit
      Release
  end function

  Public function Release
    set m_shell = Nothing
    set m_Doc = Nothing
    set mApp = Nothing
  end function

  public Function open(ByVal docName, Byval sDirSaveReport)
    
    setDocument(docName)
    
    Dim oFSO
    Set oFSO = CreateObject("Scripting.FileSystemObject")
       
    If oFSO.FileExists(m_docName) Then
      
    set m_App  = CreateObject("QlikTech.QlikView")
    set m_Doc = app.OpenDoc(m_docName,"admin","indica@2017")
    NameQVW  = left(m_Doc.Name , len(m_Doc.Name )-4)
    strCurDate = Left(date(),2) & Mid(Date(),4,2) & Right(Date(),2)
    srchRight = ".pdf"
    sDir = sDirSaveReport
  
    End If

  end Function

  public Function Main
    
    call createDir (m_Doc)
    call printSend (m_Doc ,sDir, strCurDate, NameQVW, srchRight)
 
  end Function 

End Class