include sDirInclude&"fnameDirFile.vbs"
include sDirInclude&"fcreateDir.vbs"
include sDirInclude&"fgetAbsolutePath.vbs"
include sDirInclude&"fprintPDF.vbs"

include sDirInclude&"fsendEmail.vbs"
include sDirInclude&"fprintSend.vbs"
include sDirInclude&"cqlikview.vbs"

Dim qlik

set qlik = new QlikView

call qlik.open(nameDirFile,sDirSaveReport)