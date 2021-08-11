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
Private mIdentificacaoRegistro As String
Private mIdentificaoArquivoRemessa As String
Private mLiteralRemessa As String
Private mCodigoServico  As String
Private mLiteralServico  As String
Private mCodigoEmpresa  As String
Private mNomeEmpresa  As String
Private mNumeroCamaraCompensacao  As String 'BRADESCO=237
Private mNomeBanco As String 'BRADESCO
Private mDataGravacaoArquivo As String 'DDMMAA
Private mIdentificacaoSistema As String 'MX
Private mNumeroSequencialRemessa As String 'Auto Numerico
Private mNumeroSequencialRegistroUmAUm As String
Public Property Get NomeEmpresa() As String
    NomeEmpresa = mNomeEmpresa
End Property
Public Property Let NomeEmpresa(ByVal Valor As String)
    mNomeEmpresa = Valor
End Property
Public Property Get LiteralServico() As String
    LiteralServico = mLiteralServico
End Property
Public Property Let LiteralServico(ByVal Valor As String)
    mLiteralServico = Valor
End Property
Public Property Get CodigoServico() As String
    CodigoServico = mCodigoServico
End Property
Public Property Let CodigoServico(ByVal Valor As String)
    mCodigoServico = Valor
End Property
Public Property Get NumeroSequencialRegistroUmAUm() As String
    NumeroSequencialRegistroUmAUm = mNumeroSequencialRegistroUmAUm
End Property
Public Property Let NumeroSequencialRegistroUmAUm(ByVal Valor As String)
    mNumeroSequencialRegistroUmAUm = Valor
End Property
Public Property Get NumeroSequencialRemessa() As String
    NumeroSequencialRemessa = mNumeroSequencialRemessa
End Property
Public Property Let NumeroSequencialRemessa(ByVal Valor As String)
    mNumeroSequencialRemessa = Valor
End Property
Public Property Get IdentificacaoSistema() As String
    IdentificacaoSistema = mIdentificacaoSistema
End Property
Public Property Let IdentificacaoSistema(ByVal Valor As String)
    mIdentificacaoSistema = Valor
End Property
Public Property Get DataGravacaoArquivo() As String
    DataGravacaoArquivo = mDataGravacaoArquivo
End Property
Public Property Let DataGravacaoArquivo(ByVal Valor As String)
    mDataGravacaoArquivo = Valor
End Property
Public Property Get NomeBanco() As String
    NomeBanco = mNomeBanco
End Property
Public Property Let NomeBanco(ByVal Valor As String)
    mNomeBanco = Valor
End Property
Public Property Get NumeroCamaraCompensacao() As String
    NumeroCamaraCompensacao = mNumeroCamaraCompensacao
End Property
Public Property Let NumeroCamaraCompensacao(ByVal Valor As String)
    mNumeroCamaraCompensacao = Valor
End Property
Public Property Get CodigoEmpresa() As String
    CodigoEmpresa = mCodigoEmpresa
End Property

Public Property Let CodigoEmpresa(ByVal Valor As String)
    mCodigoEmpresa = Valor
End Property
Public Property Get IdentificacaoRegistro() As String
    IdentificacaoRegistro = mIdentificacaoRegistro
End Property

Public Property Let IdentificacaoRegistro(ByVal Valor As String)
    mIdentificacaoRegistro = Valor
End Property

Public Property Get LiteralRemessa() As String
    LiteralRemessa = mLiteralRemessa
End Property

Public Property Let LiteralRemessa(ByVal Valor As String)
    mLiteralRemessa = Valor
End Property

Public Property Get IdentificaoArquivoRemessa() As String
    IdentificaoArquivoRemessa = mIdentificaoArquivoRemessa
End Property

Public Property Let IdentificaoArquivoRemessa(ByVal Valor As String)
    mIdentificaoArquivoRemessa = Valor
End Property
Public Function gerarHeaderRemessa(diretorio As String, arquivo As String) As String
    Dim NumeroArquivo         As Integer
    Dim branco101_108 As String, branco118_394 As String
    branco101_108 = repete(" ", 8)
    branco118_394 = repete(" ", 277)
    NumeroArquivo = FreeFile
    Open diretorio & arquivo For Append Access Write As #NumeroArquivo
    Print #NumeroArquivo, Me.IdentificacaoRegistro & Me.IdentificaoArquivoRemessa & Me.LiteralRemessa & Me.CodigoServico _
    & Me.LiteralServico & Me.CodigoEmpresa & Me.NomeEmpresa & Me.NumeroCamaraCompensacao & Me.NomeBanco & Me.DataGravacaoArquivo _
    & branco101_108 & Me.IdentificacaoSistema & Me.NumeroSequencialRemessa & branco118_394 & Me.NumeroSequencialRegistroUmAUm
    Close NumeroArquivo
    gerarHeaderRemessa = arquivo
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



