VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_Declaracao 
   Caption         =   "Relatório de Entrega de Declaração"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "AR_Declaracao.dsx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "AR_Declaracao.dsx":0442
End
Attribute VB_Name = "AR_Declaracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_NoData()
    Util.Avisa "Nenhum registro encontrado"
    Unload Me
End Sub

Private Sub ActiveReport_ReportStart()
    lblTitulo1 = Temp.PegaParametro(Bdados, "ESTADO")
    lblTitulo2 = Temp.PegaParametro(Bdados, "CLIENTE")
End Sub

Private Sub GroupFooter1_Format()
    On Error Resume Next
    Set SubReport1.object = New AR_Apuracao_Imposto
    SubReport1.object.PreencherDados txtCodDeclaracao, Bdados
End Sub

Private Sub GroupHeader2_Format()
    If txtOperacao2 = "ENTRADAS" Then
        txtOperacao1 = "Notas Recebidas"
        lblCancelada.Visible = False
        chkCancelada.Visible = False
    Else
        txtOperacao1 = "Notas Emitidas"
        lblCancelada.Visible = True
        chkCancelada.Visible = True
    End If
End Sub
