VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSComando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Comando As adodb.Command
Public Enum enuTipoComando
    cmdArquivo = 256
    cmdStoredProcedure = 4
    cmdTabela = 2
    cmdTabelaDireta = 512
    cmdTexto = 1
    cmdDesconhecido = 8
End Enum
Public Enum enuTipoDado
    tipArray = 8192
    tipBigInt = 20
    tipBinary = 128
    tipBoolean = 11
    tipBSTR = 8
    tipChapter = 136
    tipChar = 129
    tipCurrency = 6
    tipDate = 7
    tipDBDate = 133
    tipDBTime = 134
    tipDBTimeStamp = 135
    tipDecimal = 14
    tipDouble = 5
    tipEmpty = 0
    tipError = 10
    tipFileTime = 64
    tipGUID = 72
    tipDispatch = 9
    tipInteger = 3
    tipUnknown = 13
    tipLongVarBinary = 205
    tipLongVarChar = 201
    tipLongVarWChar = 203
    tipNumeric = 131
    tipPropVariant = 138
    tipSingle = 4
    tipSmallint = 2
    tipTinyInt = 16
    tipUnsignedBigInt = 21
    tipUnsignedInt = 19
    tipUnsignedSmallint = 18
    tipUnsignedTinyInt = 17
    tipUserDefined = 132
    tipVarBinary = 204
    tipVarChar = 200
    tipVariant = 12
    tipVarNumeric = 139
    tipVarWChar = 202
    tipWChar = 130
End Enum
Public Enum enuDirecaoParametro
    parEntrada = 1
    parEntradaSaida = 3
    parSaida = 2
    parValorRetorno = 4
    parDesconhecido = 0
End Enum

Private Sub Class_Initialize()
    Set Comando = New adodb.Command
End Sub
Private Sub Class_Terminate()
    Set Comando = Nothing
End Sub


Public Sub Texto(Bdados As Object, TSQL As String, Tipo As enuTipoComando)

    Set Comando.ActiveConnection = Bdados.Conexao.DBConnection
    Comando.CommandText = TSQL
    Comando.CommandType = Tipo
End Sub


Public Sub Executa()
    Comando.Execute
End Sub


Public Sub setarParametro(Nome As String, Tipo As enuTipoDado, Direcao As enuDirecaoParametro, Optional Tamanho As Long, Optional Valor)
    Dim Parametro As adodb.Parameter
    
    Set Parametro = Comando.CreateParameter(Nome, Tipo, Direcao, Tamanho, Valor)
    Comando.Parameters.Append Parametro
End Sub


Public Function Parametro(Indice) As adodb.Parameter
    Set Parametro = Comando.Parameters(Indice)
End Function
