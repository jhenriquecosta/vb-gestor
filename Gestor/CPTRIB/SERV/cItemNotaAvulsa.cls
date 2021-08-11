VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cItemNotaAvulsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_NumNota As String
Private m_DescricaoServico As String
Private m_Valor As Double
Private m_Aliquota As Double
Private m_Quantidade As Double
Private m_Unidade As String
Private m_Codigo As Integer
Public Property Get Codigo() As Integer
    Codigo = m_Codigo
End Property
Public Property Let Codigo(ByVal Value As Integer)
    m_Codigo = Value
End Property

Public Property Get Unidade() As String
    Unidade = m_Unidade
End Property

Public Property Let Unidade(ByVal Value As String)
    m_Unidade = Value
End Property

Public Property Get Quantidade() As Double
    Quantidade = m_Quantidade
End Property

Public Property Let Quantidade(ByVal Value As Double)
    m_Quantidade = Value
End Property

Public Property Get Valor() As Double
    Valor = m_Valor
End Property

Public Property Let Valor(ByVal Value As Double)
    m_Valor = Value
End Property


Public Property Get Aliquota() As Double
    Aliquota = m_Aliquota
End Property

Public Property Let Aliquota(ByVal Value As Double)
    m_Aliquota = Value
End Property

Public Property Get DescricaoServico() As String
    DescricaoServico = m_DescricaoServico
End Property

Public Property Let DescricaoServico(ByVal Value As String)
    m_DescricaoServico = Value
End Property

Public Property Get NumNota() As String
    NumNota = m_NumNota
End Property

Public Property Let NumNota(ByVal Value As String)
    m_NumNota = Value
End Property

Public Function Salvar() As Boolean
    Dim campos As String
    Dim valores As String
    campos = "tin_tna_numero_nota, tin_descricao_servico, tin_valor, tin_quantidade, tin_unidade,TIN_ALIQUOTA,tin_codigo"
    valores = Bdados.PreparaValor(m_NumNota, m_DescricaoServico, m_Valor, m_Quantidade, m_Unidade, m_Aliquota, m_Codigo)
    If Bdados.InsereDados("Tab_item_nota_avulsa", valores, campos) Then
        Salvar = True
    End If
End Function

Public Function Excluir(NumNota As String) As Boolean
    Bdados.DeletaDados "Tab_Item_Nota_Avulsa", "tin_tna_numero_nota ='" & NumNota & "'"
    
End Function

Public Function PreencherGrid(Grd As Object, NumNota As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "select tin_descricao_servico as Descrição," & _
            " tin_quantidade as Quantidade," & _
            " tin_unidade as Unidade, " & _
            " tin_valor As Valor_Unit, tin_quantidade * tin_valor as Valor_Total,TIN_ALIQUOTA AS Aliquita" & _
            " From tab_item_nota_avulsa" & _
            " Where tin_tna_numero_nota = '" & NumNota & "'"
    If Grd.Preencher(Bdados, Sql, 6000, 1100, 0, 1100, 1100) Then PreencherGrid = True
End Function
