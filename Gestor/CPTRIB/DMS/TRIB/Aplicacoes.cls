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
Private Objeto As Object

Public Property Get Municipio() As String
    Municipio = MUN
End Property

Public Property Let Municipio(S As String)
    MUN = S
End Property

Public Property Get Codigo_Municipio() As String
    Codigo_Municipio = CODMUN
End Property

Public Property Let Codigo_Municipio(S As String)
    CODMUN = S
End Property

Public Property Get Usuario() As String
    Usuario = User
End Property

Public Property Let Usuario(ByVal vNewValue As String)
    User = vNewValue
End Property

Public Property Get Banco() As Object
    Set Banco = Bdados.Conexao
End Property

Public Property Set Banco(NovoBanco As Object)
    Set Bdados.Conexao = NovoBanco
End Property

Public Function Abre_Aplicacao(Formulario As String, Modo As Byte, Cod_s As String, Sis As String, Desc_F As String) As Boolean
    Dim Classe As String
    Dim SQL As String
    Dim RS As Object
    Cod_sis = Cod_s: Sistema = Sis: Desc_Form = Desc_F:
    SQL = "select  tmt_classe from tab_modulo_tributario where tmt_modulo = '" & Mid(UCase(Formulario), 1, 4) & "'"
    If Bdados.AbreTabela(SQL, RS) Then
        Classe = RS!tmt_classe & ".Aplicacoes"
        Set Objeto = Nothing
        Set Objeto = CreateObject(Classe)
        Set Objeto.Banco = Bdados.Conexao
        Objeto.Codigo_Municipio = Aplicacoes.Codigo_Municipio
        Objeto.Municipio = Aplicacoes.Municipio
        Objeto.Usuario = Aplicacoes.Usuario
        Objeto.Abre_Aplicacao Formulario, Modo, Cod_s, Sis, Desc_F
    Else
        Classe = ""
        Util.AVISA "Classe do Módulo " & Mid(UCase(Formulario), 1, 4) & " não cadastrada."
    End If
End Function
