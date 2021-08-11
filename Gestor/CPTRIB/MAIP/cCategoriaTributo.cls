VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCategoriaTributo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Codigo As Integer
Private m_contaContabil As String
Private m_Categoria As String
Private m_Tipo As cTipoTributo

Public Property Get Tipo() As cTipoTributo
    Set Tipo = m_Tipo
End Property

Public Property Get Categoria() As String
    Categoria = m_Categoria
End Property

Public Property Get contaContabil() As String
    contaContabil = m_contaContabil
End Property

Public Property Get Codigo() As Integer
    Codigo = m_Codigo
End Property

Private Sub Class_Initialize()
    Set m_Tipo = New cTipoTributo
End Sub

Private Sub Class_Terminate()
    Set m_Tipo = Nothing
End Sub

Public Function buscarProximo() As Integer
    Dim rs As VSRecordset
    Dim Sql As String
    
    Sql = "SELECT MAX(TIC_COD_CATEGORIA)+1 FROM TAB_IMPOSTO_CATEGORIA"
    If Bdados.AbreTabela(Sql, rs) Then
        m_Codigo = Nvl(rs(0), 1)
    Else
        m_Codigo = 1
    End If
    buscarProximo = m_Codigo
    Bdados.FechaTabela rs
End Function

Public Function preencherGrid(ByRef grd As Object, Optional Tipo As String) As Boolean
    Dim Sql As String
    
    Sql = "SELECT TIC_COD_CATEGORIA AS Codigo, TIC_NOME_CATEGORIA as Categoria, TIC_CONTA as Conta, TTT_GRUPO as Tipo" & _
            " FROM TAB_IMPOSTO_CATEGORIA, TAB_TIPO_TRIBUTO" & _
            " WHERE TIC_TTT_COD_TIPO=TTT_COD_TIPO"
    If Trim$(Tipo) <> "" Then Sql = Sql & " AND TTT_GRUPO='" & Tipo & "'"
    preencherGrid = grd.Preencher(Bdados, Sql)
End Function
Public Sub preencherObjeto(Codigo As Integer, Categoria As String, Conta As String, Tipo As String)
    If Codigo = 0 Then
        m_Codigo = Me.buscarProximo()
    Else
        m_Codigo = Codigo
    End If
    m_Categoria = Categoria
    m_contaContabil = Conta
'    m_Receita.Descricao = Receita
End Sub

