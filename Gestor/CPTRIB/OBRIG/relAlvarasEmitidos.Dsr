VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} relAlvarasEmitidos 
   Caption         =   "DAM EMITIDO"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19606
   SectionData     =   "relAlvarasEmitidos.dsx":0000
End
Attribute VB_Name = "relAlvarasEmitidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSaldo As Currency
Dim mSaldo1 As Currency
Dim mSaldo2 As Currency
Private Sub ActiveReport_ReportStart()

  ' lblEmpresa = strEMPRESA
  ' lblSlogan = strSLOGAN
   'mSaldo = 0
    lblEmpresa = Temp.PegaParametro(Bdados, "CLIENTE")
    lblSlogan = Temp.PegaParametro(Bdados, "SLOGAN")
 
   
End Sub

Private Sub PageFooter_Format()

    lblDate = Format$(Date, "dd/mm/yyyy")
    lblPage = Format(Me.pageNumber, "0000")

End Sub

