VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TraillerBradesco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private TotalArquivos As Double
Private ValorTotal As Double

Public Property Get TotalDeArquivos() As Double
    TotalDeArquivos = TotalArquivos
End Property

Public Property Let TotalDeArquivos(ByVal dTotal As Double)
    TotalArquivos = dTotal
End Property

Public Property Get TotalRemessa() As Double
    TotalRemessa = ValorTotal
End Property

Public Property Let TotalRemessa(ByVal dValor As Double)
    ValorTotal = dValor
End Property
