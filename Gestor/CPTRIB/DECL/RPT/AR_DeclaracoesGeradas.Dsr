VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_DeclaracoesGeradas 
   Caption         =   "Relatório de Declarações"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   Icon            =   "AR_DeclaracoesGeradas.dsx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   18071
   _ExtentY        =   11642
   SectionData     =   "AR_DeclaracoesGeradas.dsx":0442
End
Attribute VB_Name = "AR_DeclaracoesGeradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
    With AR_NotasFiscais
        .DataControl1.ConnectionString = DataControl1.ConnectionString
        .DataControl1.Source = "select * from vis_nota_fiscal where TNF_TDC_NUM_DECLARACAO=" & Link & " ORDER BY TNF_COD_OPERACAO"
        .Show 1
    End With
End Sub

Private Sub ActiveReport_NoData()
    Util.Avisa "Nenhum registro encontrado"
    Unload Me
End Sub

Private Sub ActiveReport_ReportStart()
    lblTitulo1 = Temp.PegaParametro(bDados, "ESTADO")
    lblTitulo2 = Temp.PegaParametro(bDados, "CLIENTE")
End Sub

Private Sub Detail_BeforePrint()
    txtCodigo.Hyperlink = txtCodigo.Text
End Sub

