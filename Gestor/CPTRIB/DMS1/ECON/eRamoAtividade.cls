VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eRamoAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_CodRamo As Double
Private m_NomeRamo As String

Public Property Get NomeRamo() As String
    NomeRamo = m_NomeRamo
End Property

Public Property Let NomeRamo(ByVal Value As String)
    m_NomeRamo = Value
End Property

Public Property Get CodRamo() As Double
    CodRamo = m_CodRamo
End Property

Public Property Let CodRamo(ByVal Value As Double)
    m_CodRamo = Value
End Property

Public Function Buscar(CodRamo As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "SELECT TRA_COD_RAMO, TRA_NOME_RAMO FROM TAB_RAMO_ATIVIDADE WHERE TRA_COD_RAMO = '" & CodRamo & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        m_CodRamo = "" & Rs!TRA_COD_RAMO
        m_NomeRamo = "" & Rs!TRA_NOME_RAMO
        Buscar = True
    End If
End Function

Public Function Gravar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    campos = "TRA_COD_RAMO, TRA_NOME_RAMO"
    valores = Bdados.PreparaValor(m_CodRamo, m_NomeRamo)
    condicao = "TRA_COD_RAMO = '" & m_CodRamo & "'"
    If Bdados.GravaDados("TAB_RAMO_ATIVIDADE", valores, campos, condicao) Then
        Gravar = True
    End If
End Function

Public Function Excluir(CodRamo As String) As Boolean
    If Bdados.DeletaDados("TAB_RAMO_ATIVIDADE", "TRA_COD_RAMO = '" & m_CodRamo & "'") Then Excluir = True
End Function

Public Function PreencherCbo(Combo As Object) As Boolean
    Dim Sql As String
    Sql = "SELECT TRA_NOME_RAMO, TRA_COD_RAMO FROM TAB_RAMO_ATIVIDADE"
    Combo.Preencher Bdados, Sql
End Function

Public Function PreencherGrd(Grid As Object, Optional CodRamo As String, Optional NomeRamo As String) As Boolean
    Dim Sql As String
    Dim sqlaux As String
    Sql = "SELECT TRA_COD_RAMO as Código, TRA_NOME_RAMO as Descrição FROM TAB_RAMO_ATIVIDADE"
    If Trim(CodRamo) <> "" Then
        sqlaux = sqlaux & " and TRA_COD_RAMO = '" & CodRamo & "'"
    End If
    If Trim(NomeRamo) <> "" Then
        sqlaux = sqlaux & " and TRA_NOME_RAMO = '" & NomeRamo & "'"
    End If
    If sqlaux <> "" Then Sql = Sql & " WHERE " & Right(sqlaux, Len(sqlaux) - 4)
    Sql = Sql & " order by TRA_NOME_RAMO"
    If Grid.Preencher(Bdados, Sql) Then PreencherGrd = True
End Function
