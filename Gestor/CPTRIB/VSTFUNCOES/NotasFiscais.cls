VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotasFiscais"
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


Public Sub Adicionar(Item As NotaFiscal)
    objColecao.Add Item
End Sub


Public Property Get Nota(vntIndexKey As Variant) As NotaFiscal
  Set Nota = objColecao(vntIndexKey)
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
    Set NewEnum = objColecao.[_NewEnum]
End Property

Friend Function Gravar(Declaracao As String) As Boolean
    Dim Campos As String, Valores As String
    Dim Nota As NotaFiscal
    Dim ValorItem As String
    Gravar = True
    Bdados.DeletaDados "TAB_NOTA_FISCAL", "TNF_TDC_NUM_DECLARACAO=" & Declaracao
    Campos = "TNF_TDC_NUM_DECLARACAO, TNF_INSCRICAO_OPERACAO, TNF_NUM_NOTA, TNF_COD_OPERACAO,TNS_DATA_NOTA," & _
            "TNS_VALOR_NOTA,TNS_ALIQUOTA,TNS_BASE_CALCULO,TNS_VALOR_MATERIAL_ICMS,TNS_VALOR_IMPOSTO_DEVIDO," & _
            "TNS_VALOR_IMPOSTO_RETIDO,TNS_NOTA_CANCELADA,TNS_AIDF,TNS_TCI_IM_TOMADOR"
    For Each Nota In objColecao
        Valores = Bdados.PreparaValor(Declaracao, Bdados.Converte(Nota.Destinatario, tctexto), Nota.Numero, _
            Nota.TipoOperacao, Bdados.Converte(Nota.Data, TCDataHora), Bdados.Converte(Nota.ValorTotal, TCDuplo), Bdados.Converte(Nota.Aliquota, TCDuplo), _
            Bdados.Converte(Nota.BaseCalculo, TCDuplo), Bdados.Converte(Nota.ValorMaterialICMS, TCDuplo), Bdados.Converte(Nota.ImpostoDevido, TCDuplo), _
            Bdados.Converte(Nota.ImpostoRetido, TCDuplo), Nota.Status, Nota.AIDF, Nota.Tomador)
        Bdados.InsereDados "TAB_NOTA_FISCAL", Valores, Campos
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

