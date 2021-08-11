VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDistritoSetor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Distrito As String
Private m_Setor As String
Private m_Quadra As String

Public Property Get Quadra() As String
    Quadra = m_Quadra
End Property

Public Property Let Quadra(ByVal Value As String)
    m_Quadra = Value
End Property

Public Property Get Setor() As String
    Setor = m_Setor
End Property

Public Property Let Setor(ByVal Value As String)
    m_Setor = Value
End Property

Public Property Get Distrito() As String
    Distrito = m_Distrito
End Property

Public Property Let Distrito(ByVal Value As String)
    m_Distrito = Value
End Property

Public Function Contem(Distrito As String, Setor As String) As Boolean
    On Error GoTo Trata
    Contem = Bdados.AbreTabela("SELECT * FROM TAB_DISTRITO_SETOR WHERE TDS_SETOR = '" & Setor & "' AND TDS_DISTRITO = '" & Distrito & "'")
    Exit Function
Trata:
    Util.Erro Err.Description
End Function

Public Function Gravar() As Boolean
    On Error GoTo Trata
    Gravar = Bdados.GravaDados("TAB_DISTRITO_SETOR", Bdados.PreparaValor(Bdados.Converte(m_Distrito, TCTexto), Bdados.Converte(m_Setor, TCTexto)), "TDS_DISTRITO,TDS_SETOR", "TDS_SETOR = '" & m_Setor & "' AND TDS_DISTRITO = '" & m_Distrito & "'")
    Exit Function
Trata:
    Util.Erro Err.Description
End Function

Public Function Excluir(Distrito As String, Setor As String) As Boolean
    On Error GoTo Trata
    Excluir = Bdados.DeletaDados("TAB_DISTRITO_SETOR", "TDS_SETOR = '" & Setor & "' AND TDS_DISTRITO = '" & Distrito & "'")
    Exit Function
Trata:
    Util.Erro Err.Description
End Function

Public Function PreencherGrd(grid As Object, Optional Distrito As String, Optional Setor As String) As Boolean
    Dim Sql As String
    On Error GoTo Trata
    Sql = "SELECT TDS_DISTRITO AS Distrito, TDS_SETOR AS Setor " & _
                "FROM TAB_DISTRITO_SETOR "
    Sql = Sql & IIf(Distrito <> "", "AND TDS_DISTRITO = '" & Distrito & "'", "")
    Sql = Sql & IIf(Setor <> "", "AND TDS_SETOR = '" & Setor & "'", "")
    Sql = Sql & " ORDER BY 1"
    grid.Preencher Bdados, Sql
    PreencherGrd = grid.ListItems.Count > 0
    Exit Function
Trata:
    Util.Erro Err.Description
End Function
