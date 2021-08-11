VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Detalhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ValorTarifa                 As Double
Private ValorRecebido            As Double
Private NumeroDocumento     As Double
Private AgenciaArrec                     As String
Private BancoArrec                    As String
Private AgenciaCred                     As String
Private ContaCred                        As String
Private CodPagamento           As String
Private TotalArquivos             As Double
Private ValorTotal                  As Double
Private VersaoBarra                  As Double
Private Inscricao As String
Private DataPag As String
Private CodTributo As String
Private Periodo As Double
Private Parcela As Integer
Public Property Get CodigoPagamento() As String
    CodigoPagamento = CodPagamento
End Property
Public Property Let CodigoPagamento(ByVal Value As String)
    CodPagamento = Value
'    Bdados.sp_GeraCorrelativo
End Property

Public Property Get Valor_Tarifa() As Double
    Valor_Tarifa = ValorTarifa
End Property

Public Property Let Valor_Tarifa(ByVal Value As Double)
    ValorTarifa = Value
End Property

Public Property Get Valor_Recebido() As Double
    Valor_Recebido = ValorRecebido
End Property

Public Property Let Valor_Recebido(ByVal Value As Double)
    ValorRecebido = Value
End Property

Public Property Get Versao_Barra() As Double
    Versao_Barra = VersaoBarra
End Property

Public Property Let Versao_Barra(ByVal Value As Double)
    VersaoBarra = Value
End Property

Public Property Get Documento() As Double
    Documento = NumeroDocumento
End Property

Public Property Let Documento(ByVal Value As Double)
    NumeroDocumento = Value
End Property

Public Property Get AgenciaArrecadacao() As Double
    AgenciaArrecadacao = AgenciaArrec
End Property

Public Property Get BancoArrecadacao() As String
    BancoArrecadacao = BancoArrec
End Property

Public Property Let BancoArrecadacao(ByVal Value As String)
    BancoArrec = Value
End Property

Public Property Get DataPagamento() As String
    DataPagamento = DataPag
End Property

Public Property Let DataPagamento(ByVal Value As String)
    DataPag = Value
End Property

Public Property Let AgenciaArrecadacao(ByVal Value As Double)
    AgenciaArrec = Value
End Property

Public Property Get AgenciaCredito() As Double
    AgenciaCredito = AgenciaCred
End Property

Public Property Let AgenciaCredito(ByVal Value As Double)
    AgenciaCred = Value
End Property

Public Property Get ContaCredito() As Double
    ContaCredito = ContaCred
End Property

Public Property Let ContaCredito(ByVal Value As Double)
    ContaCred = Value
End Property

Public Property Get obInscricao() As String
    obInscricao = Inscricao
End Property

Public Property Let obInscricao(ByVal Value As String)
    Inscricao = Value
End Property

Public Property Get obCodTributo() As String
    obCodTributo = CodTributo
End Property

Public Property Let obCodTributo(ByVal Value As String)
    CodTributo = Value
End Property

Public Property Get obPeriodo() As Double
    obPeriodo = Periodo
End Property

Public Property Let obPeriodo(ByVal Value As Double)
    Periodo = Value
End Property

Public Property Get obParcela() As Double
    obParcela = Parcela
End Property

Public Property Let obParcela(ByVal Value As Double)
    Parcela = Value
End Property
