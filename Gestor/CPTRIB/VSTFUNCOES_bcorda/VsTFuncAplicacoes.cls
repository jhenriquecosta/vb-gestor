VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VsTFuncAplicacoes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum TipoContrib
    TcoFisica = 0
    TcoJuridica = 1
End Enum

Public Property Get Municipio() As String
    Municipio = UCase(Temp.PegaParametro(Bdados, "FANTASIA"))
End Property

Public Property Let Municipio(S As String)
    MUN = S
End Property

 
Public Property Get Codigo_Municipio() As String
    Codigo_Municipio = Temp.PegaParametro(Bdados, "MUNICIPIO")
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
    Imposto.TaxaFixaImpressaoDAM = CDbl(Nvl(Trim(Temp.PegaParametro(Bdados, "TXTDAM")), 0)) / 100
End Property

Private Sub Class_Initialize()
    Set Bdados = New VSDados
    Set Edita = New VSTexto
    Set Util = New VSUtil
    Set Instala = New VSInstala
    Set Temp = New VSTemp
    
End Sub
'Public Function BuscaNoEconomico(TipoPessoa As TipoContrib, Optional Controle As Object) As String
Public Sub BuscaNoEconomico(TipoPessoa As TipoContrib, Optional Controle As Object)
    'TCIP101A.cboTipoContrib.ListIndex = TipoPessoa
    TCIP101A.Inicia TipoPessoa, Controle
    TCIP101A.Show 1
    
    'BuscaNoEconomico = TCIP101A.Tag
End Sub

Public Function BuscaNoImobiliario() As String
    TCIS101A.Show 1
    BuscaNoImobiliario = TCIS101A.Tag
End Function

Public Sub BuscaInscricao(TipoInscricao As TipoInsc, Optional Controle As Object, Optional Nome As Object)
    TCIP102A.Inicia TipoInscricao, Controle, Nome
    TCIP102A.Show 1
End Sub
