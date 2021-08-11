VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Tabela As String
Private m_Codigo As Integer
Private m_Descricao As String

Public Property Get Tabela() As String
    Tabela = m_Tabela
End Property

Public Property Let Tabela(ByVal Value As String)
    Static bInitialized As Boolean
    If bInitialized Then Exit Property
    bInitialized = True
    m_Tabela = Value
End Property

Public Property Let Descricao(ByVal Value As String)
    Buscar 0, Value
End Property
Public Property Get Descricao() As String
    Descricao = m_Descricao
End Property

Public Property Let Codigo(ByVal Value As Integer)
    Buscar Value, ""
End Property
Public Property Get Codigo() As Integer
    Codigo = m_Codigo
End Property


Private Function Buscar(Cod As Integer, Descr As String) As Boolean
    Dim rs As VSRecordset
    Dim sql As String
    
    sql = "SELECT * FROM TAB_GERAL " & _
            " WHERE TGE_TIPO=(SELECT TGE_TIPO " & _
                                " FROM TAB_GERAL " & _
                                " WHERE TGE_CODIGO=0 AND TGE_NOME='" & Me.Tabela & "')"
    If Cod > 0 Then
        sql = sql & " AND TGE_CODIGO=" & Cod
    Else
        If Descr <> "" Then
            sql = sql & " AND TGE_NOME='" & Descr & "'"
        Else
            Exit Function
        End If
    End If
    If Bdados.AbreTabela(sql, rs) Then
        m_Codigo = rs!TGE_CODIGO
        m_Descricao = rs!TGE_NOME
    End If
    Bdados.FechaTabela rs
End Function

Public Sub preencherCombo(ByRef cbo As Object)
    cbo.PreencherGeral Bdados, Me.Tabela
End Sub
