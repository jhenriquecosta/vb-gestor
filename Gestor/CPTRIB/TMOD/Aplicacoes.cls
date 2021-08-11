VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Aplicacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public Usuario As String
Public Sistema As String
Public BdSis As Object
Public BdSeg As Object
Public Edita As Object
Public Util As Object
Public Instala As Object
Public Seguranca As Object
Public Temp As Object
Public VSis As Object
Private Tipo As Long
Private Dsn As String
Private User As String
Private Pass As String
Private Cat As String
Private ArqBanco As String
Public Lotacoes As String

Public Property Get municipio() As String
    municipio = MUN
End Property

Public Property Let municipio(S As String)
    MUN = S
    AplicacoesVTFuncoes.municipio = S
End Property
 
Public Property Get Codigo_Municipio() As String
    Codigo_Municipio = Temp.PegaParametro(Bdados, "MUNICIPIO")
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
    On Error Resume Next
    Set Bdados.Conexao = NovoBanco
'    Set AplicacoesDecl.Banco = Bdados.Conexao
    Set AplicacoesVTFuncoes.Banco = Bdados.Conexao
    
End Property

Private Sub Class_Initialize()
    Set Bdados = New VSDados
    Set Edita = New VSTexto
    Set Util = New VSUtil
    Set Instala = New VSInstala
    Set Temp = New VSTemp
End Sub


Public Function Abre_Aplicacao(Formulario As String, modo As Byte, Cod_s As String, Sis As String, Desc_F As String, Optional Tag As String) As Boolean
    On Error GoTo Erro
    Cod_sis = Cod_s: Sistema = Sis: Desc_Form = Desc_F:
    Dim NomeFormulario  As Object
    
    Set NomeFormulario = Forms.Add(Formulario)
    If Trim(Tag) <> "" Then NomeFormulario.Tag = Tag
    NomeFormulario.Show modo
    Abre_Aplicacao = True
    Exit Function
Erro:

    If Err.Number = 424 Then
        Util.Avisa "Formulário " & Formulario & " não cadastrado no projeto."
        Abre_Aplicacao = False
        Exit Function
    Else
        Util.Avisa Err.Description
        Abre_Aplicacao = False
    End If
End Function
