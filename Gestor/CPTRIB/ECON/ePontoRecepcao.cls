VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ePontoRecepcao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_CodPonto As Double
Private m_NomePonto As String

Public Property Get NomePonto() As String
    NomePonto = m_NomePonto
End Property

Public Property Let NomePonto(ByVal Value As String)
    m_NomePonto = Value
End Property

Public Property Get CodPonto() As Double
    CodPonto = m_CodPonto
End Property

Public Property Let CodPonto(ByVal Value As Double)
    m_CodPonto = Value
End Property

Public Function Buscar(CodPonto As String) As Boolean
    Dim sql As String
    Dim rs As VSRecordset
    sql = "SELECT TPR_COD_PONTO, TPR_DESCRICAO FROM TAB_PONTO_RECEPCAO WHERE TPR_COD_PONTO = '" & CodPonto & "'"
    If Bdados.AbreTabela(sql, rs) Then
        m_CodPonto = "" & rs!TPR_COD_PONTO
        m_NomePonto = "" & rs!TPR_DESCRICAO
        Buscar = True
    End If
End Function

Public Function Gravar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    campos = "TPR_COD_PONTO, TPR_DESCRICAO"
    valores = Bdados.PreparaValor(m_CodPonto, m_NomePonto)
    condicao = "TPR_COD_PONTO = '" & m_CodPonto & "'"
    If Bdados.GravaDados("TAB_PONTO_RECEPCAO", valores, campos, condicao) Then
        Gravar = True
    End If
End Function

Public Function Excluir(CodPonto As String) As Boolean
    If Bdados.DeletaDados("TAB_PONTO_RECEPCAO", "TPR_COD_PONTO = '" & m_CodPonto & "'") Then Excluir = True
End Function

Public Function PreencherCbo(Combo As Object) As Boolean
    Dim sql As String
    sql = "SELECT TPR_DESCRICAO, TPR_COD_PONTO FROM TAB_PONTO_RECEPCAO"
    Combo.Preencher Bdados, sql
End Function

Public Function PreencherGrd(Grid As Object, Optional CodPonto As String, Optional NomePonto As String) As Boolean
    Dim sql As String
    Dim sqlaux As String
    sql = "SELECT TPR_COD_PONTO as Código, TPR_DESCRICAO as Descrição FROM TAB_PONTO_RECEPCAO"
    If Trim(CodPonto) <> "" Then
        sqlaux = sqlaux & " and TPR_COD_PONTO = '" & CodPonto & "'"
    End If
    If Trim(NomePonto) <> "" Then
        sqlaux = sqlaux & " and TPR_DESCRICAO = '" & NomePonto & "'"
    End If
    If sqlaux <> "" Then sql = sql & " WHERE " & Right(sqlaux, Len(sqlaux) - 4)
    sql = sql & " order by TPR_DESCRICAO"
    If Grid.Preencher(Bdados, sql) Then PreencherGrd = True
End Function
