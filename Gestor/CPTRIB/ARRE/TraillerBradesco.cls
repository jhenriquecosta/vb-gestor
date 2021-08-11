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
Private mIdentificacaoRegistro As String
Private mNumeroSequencialRegistro As String
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
Public Property Get IdentificacaoRegistro() As String
     IdentificacaoRegistro = mIdentificacaoRegistro
End Property
Public Property Let IdentificacaoRegistro(ByVal Valor As String)
     mIdentificacaoRegistro = Valor
End Property
Public Property Get NumeroSequencialRegistro() As String
     NumeroSequencialRegistro = mNumeroSequencialRegistro
End Property
Public Property Let NumeroSequencialRegistro(ByVal Valor As String)
     mNumeroSequencialRegistro = Valor
End Property
Public Function gerarTrailerRemessa(diretorio As String, arquivo As String) As String
    Dim NumeroArquivo         As Integer
    Dim branco2_394 As String
    branco2_394 = repete(" ", 393)
    NumeroArquivo = FreeFile
    Open diretorio & arquivo For Append Access Write As #NumeroArquivo
    Print #NumeroArquivo, Me.IdentificacaoRegistro & branco2_394 & Me.NumeroSequencialRegistro
    Close NumeroArquivo
    gerarTrailerRemessa = arquivo
End Function
Private Function repete(caractere As String, x As Integer) As String
    Dim resultado As String
    Dim i As Integer
    resultado = ""
    For i = 1 To x
         resultado = resultado & caractere
    Next i
    repete = resultado
End Function

