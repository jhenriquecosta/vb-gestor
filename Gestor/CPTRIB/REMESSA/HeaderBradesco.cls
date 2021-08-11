VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HeaderBradesco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Versao As String
Private CodRemessa As String
Private CodBanco As String
Private DataGeracao  As String
Private DataPagamento  As String
Private NumArq  As String

Public Property Get Remessa() As Double
    Remessa = CodRemessa
End Property

Public Property Let Remessa(ByVal dRemessa As Double)
    CodRemessa = dRemessa
End Property

Public Property Get Banco() As Double
    Banco = CodBanco
End Property

Public Property Let Banco(ByVal dBanco As Double)
    CodBanco = dBanco
End Property

Public Property Get DataGeracaoArquivo() As String
    DataGeracaoArquivo = DataGeracao
End Property

Public Property Let DataGeracaoArquivo(ByVal sData As String)
    DataGeracao = sData
End Property

Public Property Get DataPagamentoDoc() As String
    DataPagamentoDoc = DataPagamento
End Property

Public Property Let DataPagamentoDoc(ByVal sData As String)
    DataPagamento = sData
End Property

Public Property Get NumeroArquivo() As Double
    NumeroArquivo = NumArq
End Property

Public Property Let NumeroArquivo(ByVal dArq As Double)
    NumArq = dArq
End Property

Public Property Get VersaoArquivo() As String
    VersaoArquivo = Versao
End Property

Public Property Let VersaoArquivo(ByVal sVersao As String)
    Versao = sVersao
End Property
