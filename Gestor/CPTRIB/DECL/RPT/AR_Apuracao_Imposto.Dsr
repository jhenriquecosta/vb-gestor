VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_Apuracao_Imposto 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20452
   _ExtentY        =   12515
   SectionData     =   "AR_Apuracao_Imposto.dsx":0000
End
Attribute VB_Name = "AR_Apuracao_Imposto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub PreencherDados(CodDeclaracao As Long, bDados As VSDados)
    Dim rsApuracao As VSRecordset
    Dim Controle As DDActiveReports2.Field
    
    bDados.AbreTabela "select * from tab_detalhe_declaracao where TDD_TDC_NUM_DECLARACAO=" & CodDeclaracao & " order by TDD_TCD_COD_ITEM", rsApuracao, Estatico, SomenteLeitura
    With rsApuracao
        Do Until .EOF
            Set Controle = Me.Detail.Controls("txtItem" & CStr(!TDD_TCD_COD_ITEM))
            If Controle.OutputFormat <> "" Then
                Controle = Format(!TDD_VALOR_ITEM, Controle.OutputFormat)
            Else
                Controle = !TDD_VALOR_ITEM
            End If
            .MoveNext
        Loop
        .Fechar
    End With
End Sub

