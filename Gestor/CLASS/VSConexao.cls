VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSConexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Conn As Connection
Public Dsn As String
Public User As String
Public Password As String
Public Catalog As String
Public FormatoBanco As TipoBanco
Public Enum TipoBanco
    Access = 0
    SQLServer = 1
    oracle = 2
    Interbase = 3
    Postgres = 4
End Enum

Public Property Get DBConnection() As Connection
    If bRegistrado Then Set DBConnection = Conn
End Property

Public Property Set DBConnection(ByVal NewDBConnection As Connection)
    If bRegistrado Then Set Conn = NewDBConnection
End Property

Friend Sub BeginTrans()
    If bRegistrado Then Conn.BeginTrans
End Sub

Friend Sub CommitTrans()
    If bRegistrado Then Conn.CommitTrans
End Sub

Friend Sub RollbackTrans()
    If bRegistrado Then Conn.RollbackTrans
End Sub

Friend Sub Abrir(ConnStr As String)
    If bRegistrado Then
        Conn.Open ConnStr
        Conn.CommandTimeout = 0
    End If
End Sub

Public Property Get Errors() As ADODB.Errors
    If bRegistrado Then Set Errors = Conn.Errors
End Property

Friend Sub Execute(Str As String)
    If bRegistrado Then Conn.Execute Str
End Sub

Public Property Get State() As Long
    If bRegistrado Then State = Conn.State
End Property

Friend Sub Fechar()
    If bRegistrado Then Conn.Close
End Sub

Public Property Get ConnectionString() As String
    If bRegistrado Then ConnectionString = Conn.ConnectionString
End Property

Private Sub Class_Initialize()
    ValidaComponente "CLASS"
    If bRegistrado Then
        Set Conn = New Connection
    End If
End Sub

Private Sub Class_Terminate()
    If bRegistrado Then Set Conn = Nothing
End Sub
