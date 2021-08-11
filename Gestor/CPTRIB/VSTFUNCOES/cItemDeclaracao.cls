VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cItemDeclaracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private varNumero As Integer
Private varValor As String
Private varDocumento As String


Public Property Let Numero(value As Integer)
    varNumero = value
End Property
Public Property Get Numero() As Integer
    Numero = varNumero
End Property

Public Property Let Valor(value As String)
    varValor = value
End Property
Public Property Get Valor() As String
    Valor = varValor
End Property


Public Property Let Documento(value As String)
    varDocumento = value
End Property
Public Property Get Documento() As String
    Documento = varDocumento
End Property
