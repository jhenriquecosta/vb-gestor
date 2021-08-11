VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Propriedades do HEADER
Private h_Versao                                                            As String
Private h_Bancos                                                           As String
Private h_PrefixoAgencia                                               As String
Private h_DigitoVerificadorAgencia                               As String
Private h_CodigoCedente                                               As String
Private h_DigitoVerificadorCodigoCedente                   As String
Private h_NomeEmpresa                                               As String
Private h_Banco                                                             As String
Private h_DataGravacao                                                As String
Private h_SequencialRemessa                                        As String
Private h_SequencialRegistro                                         As String

'Propriedades do DETALHE
Private d_TipoIncricao As String
Private d_InscricaoEmpresa As String
Private d_PrefixoAgencia As String
Private d_DigitoVerificadorCedente As String
Private d_NumeroConvenio As String



'Propriedades do HEADER
Public Property Get Versao() As String
    Versao = h_Versao
End Property

Public Property Let Versao(ByVal Valor As String)
    h_Banco = Valor
End Property

Public Property Get Bancos() As String
    Bancos = h_Bancos
End Property

Public Property Let Bancos(ByVal Valor As String)
    h_Bancos = Valor
End Property

Public Property Get PrefixoDaAgencia() As String
    PrefixoDaAgencia = h_PrefixoAgencia
End Property

Public Property Let PrefixoDaAgencia(ByVal Valor As String)
    h_PrefixoAgencia = Valor
End Property

Public Property Get DigitoVerificadorAgencia() As String
    DigitoVerificadorAgencia = h_DigitoVerificadorAgencia
End Property

Public Property Let DigitoVerificadorAgencia(ByVal Valor As String)
    h_DigitoVerificadorAgencia = Valor
End Property

Public Property Get CodigoCedente() As String
    CodigoCedente = h_CodigoCedente
End Property

Public Property Let CodigoCedente(ByVal Valor As String)
    h_CodigoCedente = Valor
End Property

Public Property Get DigitoVerificadorCodigoCedente() As String
    DigitoVerificadorCodigoCedente = h_DigitoVerificadorCodigoCedente
End Property

Public Property Let DigitoVerificadorCodigoCedente(ByVal Valor As String)
    h_DigitoVerificadorCodigoCedente = Valor
End Property

Public Property Get NomeEmpresa() As String
    NomeEmpresa = h_NomeEmpresa
End Property

Public Property Let NomeEmpresa(ByVal Valor As String)
    h_NomeEmpresa = Valor
End Property

Public Property Get Banco() As String
    Banco = h_Banco
End Property

Public Property Let DataGravacao(ByVal Valor As String)
    h_DataGravacao = Valor
End Property

Public Property Get DataGravacao() As String
    DataGravacao = h_Versao
End Property

Public Property Let SequencialRemessa(ByVal Valor As String)
    h_SequencialRemessa = Valor
End Property

Public Property Get SequencialRemessa() As String
    SequencialRemessa = h_SequencialRemessa
End Property

Public Property Let SequencialRegistro(ByVal Valor As String)
    h_SequencialRegistro = Valor
End Property

Public Property Get SequencialRegistro() As String
    SequencialRegistro = h_SequencialRegistro
End Property
