VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private RS As Recordset

Public Property Get Eof() As Boolean
    On Error GoTo Trata
    If bRegistrado Then Eof = RS.Eof
Trata:
End Property

Public Property Get Bof() As Boolean
    On Error GoTo Trata
    If bRegistrado Then Bof = RS.Bof
Trata:
End Property

Public Sub Fechar()
    If bRegistrado Then RS.Close
End Sub

Public Sub Move(NumRecords As Long)
    On Error GoTo Trata
    If bRegistrado Then RS.Move NumRecords
Trata:
End Sub

Public Sub MoveNext()
    On Error GoTo Trata
    If bRegistrado Then RS.MoveNext
Trata:
End Sub

Public Sub MoveFirst()
    On Error GoTo Trata
    If bRegistrado Then RS.MoveFirst
Trata:
End Sub

Public Sub MoveLast()
    On Error GoTo Trata
    If bRegistrado Then RS.MoveLast
Trata:
End Sub

Public Sub MovePrevious()
    On Error GoTo Trata
    If bRegistrado Then RS.MovePrevious
Trata:
End Sub

Public Sub Abrir(Source As String, ActiveConnection As VSConexao, Optional CursorType As TipoCursor = SomenteAvanco, Optional LockType As TipoTrava = SomenteLeitura, Optional Options As Long = -1)
    On Error GoTo Trata
    If bRegistrado Then RS.Open Source, ActiveConnection.DBConnection, CursorType, LockType, Options
Trata:
End Sub

Public Property Get RecordCount() As Long
    On Error GoTo Trata
    If bRegistrado Then RecordCount = RS.RecordCount
Trata:
End Property

Public Property Get AbsolutePosition() As Long
    On Error GoTo Trata
    If bRegistrado Then AbsolutePosition = RS.AbsolutePosition
Trata:
End Property

Public Property Get State() As Long
    On Error GoTo Trata
    If bRegistrado Then State = RS.State
Trata:
End Property

Private Sub Class_Initialize()
    ValidaComponente "CLASS"
    If bRegistrado Then Set RS = New Recordset
Trata:
End Sub

Public Property Get Fields() As adodb.Fields
Attribute Fields.VB_UserMemId = 0
    On Error GoTo Trata
    If bRegistrado Then Set Fields = RS.Fields
Trata:
End Property

Private Sub Class_Terminate()
    If bRegistrado Then Set RS = Nothing
End Sub
