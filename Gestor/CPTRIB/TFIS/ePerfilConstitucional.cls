VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ePerfilConstitucional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public procedimento As New eProcedimentosFiscais
Private a_CodPerfil As Integer
Private a_Perfil As String

Public Property Get CodPerfil() As String
    CodPerfil = a_CodPerfil
End Property

Public Property Let CodPerfil(ByVal Value As String)
    a_CodPerfil = Value
End Property

Public Property Get Perfil() As String
    Perfil = a_Perfil
End Property

Public Property Let Perfil(ByVal Value As String)
    a_Perfil = Value
End Property

'Public Function CarregaAtasNit()
'    Dim sql As String
'    Dim Condicao As String
'
'    sql = "Select * from Tab_Perfil_Constitucional"
'
'    If CodPerfil <> "" Then
'        If Condicao = "" Then
'            Condicao = " tpc_codigo_perfil = '" & CodPerfil & "'"
'        Else
'             Condicao = " AND tpc_codigo_perfil = '" & CodPerfil & "'"
'        End If
'
'End Function
