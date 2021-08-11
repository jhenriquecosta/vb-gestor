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

Public Property Get Banco() As VSClass.VSConexao
    Set Banco = Bdados.Conexao
End Property

Public Property Set Banco(NovoBanco As VSClass.VSConexao)
    Set Bdados.Conexao = NovoBanco
End Property

Private Sub Class_Initialize()
    Set Bdados = New VSDados
    Set Edita = New VSTexto
    Set Util = New VSUtil
    Set Instala = New VSInstala
    Set Temp = New VSTemp
End Sub
