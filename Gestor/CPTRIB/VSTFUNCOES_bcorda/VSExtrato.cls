VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSExtrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum TipoExtrato
    extNotificacao
    extExtratoPagamento
End Enum

Public Function GravarDocumentos(Tipo As TipoExtrato, CodPagamento As Double, lstDocs As Object) As String
    Dim Valores As String, Campos As String, Tabela As String
    Dim ListaDocs As String
    Dim Inserir As Boolean, I As Integer
    
    '1.
    Select Case Tipo
        Case extNotificacao
            Campos = "TPN_TNO_COD_NOTIFICACAO,TPN_TGT_COD_PAGAMENTO,TPN_SUB_VALOR,TPN_TIP_COD_IMPOSTO"
            Tabela = "TAB_PAGAMENTO_NOTIFICACAO"
        
        Case extExtratoPagamento
            Campos = "TPE_COD_PAGAMENTO_EXTRATO,TPE_TGT_COD_PAGAMENTO,TPE_SUB_VALOR,TPE_TIP_COD_IMPOSTO"
            Tabela = "TAB_PAGAMENTO_EXTRATO"
    End Select
    
    '2.
    ListaDocs = ""
    For I = 1 To lstDocs.ListItems.Count
        If Tipo = extExtratoPagamento Then
            If lstDocs.ListItems(I).Checked Then
                Inserir = True
            End If
        Else
            Inserir = True
        End If
        If Inserir Then
            Valores = Bdados.PreparaValor(CodPagamento, lstDocs.ListItems(I).Text, Bdados.Converte(lstDocs.ListItems(I).SubItems(5), TCDuplo), lstDocs.ListItems(I).SubItems(7))
            Bdados.InsereDados Tabela, Valores, Campos
            ListaDocs = ListaDocs & lstDocs.ListItems(I).Text & " (" & lstDocs.ListItems(I).SubItems(5) & ") - "
        End If
    Next
    
    '3.
    ListaDocs = Left(ListaDocs, Len(ListaDocs) - 3)
    GravarDocumentos = ListaDocs
End Function

Public Function GerarExtrato(Tipo As TipoExtrato, IM As String, IC As String, Vencimento As Date, Valor As Double) As Double
    Dim Valores As String, Campos As String, Tabela As String
    Dim Correlativo As TipoCorrelativo, Constante As String
    Dim Conta As New ContaCorrente
    Dim CodPagamento As String
    
    '1.
    Select Case Tipo
        Case extNotificacao
            Correlativo = EtsNotificacao
            Constante = Const_Notificacao
            
        Case extExtratoPagamento
            Correlativo = EtsExtratoPagamento
            Constante = Const_Extrato
    End Select
    
    '2.
    CodPagamento = Conta.GeraCodPagamento(CStr(Correlativo))
    Conta.GeraPagamento IM, IC, Constante, Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), CStr(Vencimento), Valor, 0, 0, CodPagamento, 0, 0, 0, , EtcCreditoTributario
    
    '3.
    GerarExtrato = CodPagamento
    Set Conta = Nothing
End Function
