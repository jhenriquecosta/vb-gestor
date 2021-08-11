VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Aplicacoes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Municipio() As String
    Municipio = MUN
End Property

Public Property Let Municipio(S As String)
    MUN = S
    AplicacoesVTFuncoes.Municipio = S
End Property
 
Public Property Get Codigo_Municipio() As String
    Codigo_Municipio = CODMUN
End Property

Public Property Let Codigo_Municipio(S As String)
    CODMUN = S
    AplicacoesVTFuncoes.Codigo_Municipio = S
End Property

Public Property Get Usuario() As String
    Usuario = User
End Property

Public Property Let Usuario(ByVal vNewValue As String)
    User = vNewValue
    AplicacoesVTFuncoes.Usuario = vNewValue
End Property

Public Property Get Banco() As Object
    Set Banco = Bdados.Conexao
End Property

Public Property Set Banco(NovoBanco As Object)
    Set Bdados.Conexao = NovoBanco
    Set AplicacoesVTFuncoes.Banco = Bdados.Conexao
End Property

Private Sub Class_Initialize()
    Set Bdados = New VSDados
    Set Edita = New VSTexto
    Set Util = New VSUtil
    Set Instala = New VSInstala
    Set Temp = New VSTemp
End Sub


Public Function Abre_Aplicacao(Formulario As String, Modo As Byte, Cod_s As String, Sis As String, Desc_F As String) As Boolean
    Cod_sis = Cod_s: Sistema = Sis: Desc_Form = Desc_F:
    Select Case UCase(Formulario)
'--------------------------------------------------
        Case "TCIM101": TCIM101.Show Modo
        Case "TCIM102", "TDIG102": TCIM102.Show Modo
        Case "TCIM103", "TDIG103": TCIM103.Show Modo
        Case "TCIM104", "TDIG104": TCIM104.Show Modo
        Case "TCIM201": TCIM201.Show Modo
        Case "TCIM301": TCIM301.Show Modo
        Case "TCIM401": TCIM401.Show Modo
        Case "TCIM403", "TDIG403": TCIM403.Show Modo
        Case "TCIM404": TCIM404.Show Modo
'--------------------------------------------------
        Case Else
            Util.Informa Formulario & " : Acesso Negado."
    End Select
End Function