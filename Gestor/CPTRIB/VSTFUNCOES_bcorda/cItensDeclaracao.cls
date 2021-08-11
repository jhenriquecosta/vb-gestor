VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cItensDeclaracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objColecao As Collection

Private Sub Class_Initialize()
    Set objColecao = New Collection
End Sub
Private Sub Class_Terminate()
    Set objColecao = Nothing
End Sub


Public Sub Adicionar(Item As cItemDeclaracao)
    objColecao.Add Item
End Sub


Public Property Get Item(vntIndexKey As Variant) As cItemDeclaracao
Attribute Item.VB_UserMemId = 0
  Set Item = objColecao(vntIndexKey)
End Property


Public Property Get Qtd() As Long
    Qtd = objColecao.Count
End Property

Public Sub Remover(vntIndexKey As Variant)
    objColecao.Remove vntIndexKey
End Sub

Public Sub Limpar()
    Set objColecao = New Collection
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = objColecao.[_NewEnum]
End Property

Friend Function Gravar(Declaracao As String) As Boolean
    Dim Campos As String, Valores As String
    Dim Item As cItemDeclaracao
    Dim ValorItem As String
    Dim I As Integer
    
    Gravar = True
    Bdados.DeletaDados "tab_detalhe_declaracao", "TDD_TDC_NUM_DECLARACAO=" & Declaracao
    Campos = "TDD_TDC_NUM_DECLARACAO, TDD_TCD_COD_ITEM, TDD_NUM_DOC, TDD_VALOR_ITEM"
    
    For Each Item In objColecao
        If IsNumeric(Item.Valor) Then
            ValorItem = Bdados.Converte(Item.Valor, TCDuplo)
        ElseIf IsDate(Item.Valor) Then
            ValorItem = Bdados.Converte(Item.Valor, TCDataHora)
        End If
        Valores = Bdados.PreparaValor(Declaracao, Item.Numero, Item.Documento, ValorItem)
        Gravar = Gravar And Bdados.InsereDados("tab_detalhe_declaracao", Valores, Campos)
    Next
End Function

Friend Function Total(Componente As Integer) As Double
    Dim Item As cItemDeclaracao
    Dim dblValor As Double
    
    dblValor = 0
    For Each Item In objColecao
        If Item.Numero = Componente Then
            dblValor = dblValor + CDbl(Item.Valor)
        End If
    Next
    Total = dblValor
End Function
