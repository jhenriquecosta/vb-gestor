VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim vNumero As String
Dim vData As String
Dim vValorTotsl As Double
Dim vValorMaterial As Double
Dim vValorTotal As Double
Dim vValorMaterialICMS As Double
Dim vBaseCalculo As Double
Dim vImpostoDevido As Double
Dim vImpostoRetido As Double
Dim vAliquota As Double
Dim vTomador As String

Dim vTipoOperacao As eTipoNotaOperacao
Dim vStatus As eTipoStatusNota
Dim vDestinatario As String
Public Enum eTipoNotaOperacao
    etoEntrada = 1
    etoSaida = 2
End Enum
Public AIDF As String
Public DescricaoServico As String
Public Enum eTipoStatusNota
    etsValida = 0
    etsCancelada = 1
End Enum

Public Property Get Numero() As String
    Numero = vNumero
End Property

Public Property Let Numero(ByVal vNewValue As String)
    vNumero = vNewValue
End Property

Public Property Get Data() As String
    Data = vData
End Property

Public Property Let Data(ByVal vNewValue As String)
    vData = vNewValue
End Property

Public Property Get ValorTotal() As Double
    ValorTotal = vValorTotal
End Property

Public Property Let ValorTotal(ByVal vNewValue As Double)
    vValorTotal = vNewValue
End Property

Public Property Get ValorMaterialICMS() As Double
    ValorMaterialICMS = vValorMaterialICMS
End Property

Public Property Let ValorMaterialICMS(ByVal vNewValue As Double)
    vValorMaterialICMS = vNewValue
End Property

Public Property Get BaseCalculo() As Double
    BaseCalculo = vBaseCalculo
End Property

Public Property Let BaseCalculo(ByVal vNewValue As Double)
    vBaseCalculo = vNewValue
End Property

Public Property Get ImpostoDevido() As Double
    ImpostoDevido = vImpostoDevido
End Property

Public Property Let ImpostoDevido(ByVal vNewValue As Double)
    vImpostoDevido = vNewValue
End Property

Public Property Get ImpostoRetido() As Double
    ImpostoRetido = vImpostoRetido
End Property

Public Property Let ImpostoRetido(ByVal vNewValue As Double)
    vImpostoRetido = vNewValue
End Property

Public Property Get TipoOperacao() As eTipoNotaOperacao
    TipoOperacao = vTipoOperacao
End Property

Public Property Let TipoOperacao(ByVal vNewValue As eTipoNotaOperacao)
    vTipoOperacao = vNewValue
End Property

Public Property Get Status() As eTipoStatusNota
    Status = vStatus
End Property

Public Property Let Status(ByVal vNewValue As eTipoStatusNota)
    vStatus = vNewValue
End Property

Public Property Get Destinatario() As String
    Destinatario = vDestinatario
End Property

Public Property Let Destinatario(ByVal vNewValue As String)
    vDestinatario = vNewValue
End Property


Public Property Get Aliquota() As Double
    Aliquota = vAliquota
End Property

Public Property Let Aliquota(ByVal vNewValue As Double)
    vAliquota = vNewValue
End Property

Public Property Get Tomador() As String
    Tomador = vTomador
End Property

Public Property Let Tomador(ByVal vNewValue As String)
    vTomador = vNewValue
End Property
