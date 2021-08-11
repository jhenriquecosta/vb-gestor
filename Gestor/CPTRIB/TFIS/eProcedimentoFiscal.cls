VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eProcedimentoFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum TipoAcaoFiscal
    AcaoPlantao = 1
    AcaoEstimativa = 2
    AcaoArbitramento = 3
    AcaoHomologacao = 4
End Enum

Public Enum TipoProcedimento
    tpProcedimentoTIAF = 1
    tpProcedimentoTREF = 2
    tpProcedimentoAITI = 3
    tpProcedimentoTEDI = 4
    tpProcedimentoTIFI = 5
    tpProcedimentoREFI = 6
    tpProcedimentoTEAF = 7
    tpProcedimentoAPRE = 8
    tpProcedimentoATA_RIT = 9
    tpProcedimentoATA_RIF = 10
    tpProcedimentoATA_NIT = 11
    tpProcedimentoTI = 12
End Enum

Public Autoridade As New eFuncionarios

Private p_NumeroProcedimento As String
Private p_NumeroFiscalizacao As String
Private p_Contribuinte As String
Private p_RepresentantePassivoNome As String
Private p_RepresentantePassivoCPF As String
Private p_DocumentosOriginarios  As String
Private p_DocumentosDecorrentes  As String
Private p_FundamentacaoLegal As String
Private p_TipoAcaoFiscal As Integer
Private p_LavraturaLocal As String
Private p_LavraturaData As String
Private p_LavraturaHora As String
Private p_DescricaoMotivos As String
Private p_TipoProcedimento As Integer

Public Property Get NumeroProcedimento() As String
    NumeroProcedimento = p_NumeroProcedimento
End Property

Public Property Let NumeroProcedimento(ByVal Value As String)
    p_NumeroProcedimento = Value
End Property

Public Property Get NumeroFiscalizacao() As String
    NumeroFiscalizacao = p_NumeroFiscalizacao
End Property

Public Property Let NumeroFiscalizacao(ByVal Value As String)
    p_NumeroFiscalizacao = Value
End Property

Public Property Get Contribuinte() As String
    Contribuinte = p_Contribuinte
End Property

Public Property Let Contribuinte(ByVal Value As String)
    p_Contribuinte = Value
End Property

Public Property Get RepresentantePassivoNome() As String
    RepresentantePassivoNome = p_RepresentantePassivoNome
End Property

Public Property Let RepresentantePassivoNome(ByVal Value As String)
    p_RepresentantePassivoNome = Value
End Property

Public Property Get RepresentantePassivoCPF() As String
    RepresentantePassivoCPF = p_RepresentantePassivoCPF
End Property

Public Property Let RepresentantePassivoCPF(ByVal Value As String)
    p_RepresentantePassivoCPF = Value
End Property

Public Property Get DocumentosOriginarios() As String
    DocumentosOriginarios = p_DocumentosOriginarios
End Property

Public Property Let DocumentosOriginarios(ByVal Value As String)
    p_DocumentosOriginarios = Value
End Property

Public Property Get DocumentosDecorrentes() As String
    DocumentosDecorrentes = p_DocumentosDecorrentes
End Property

Public Property Let DocumentosDecorrentes(ByVal Value As String)
    p_DocumentosDecorrentes = Value
End Property

Public Property Get FundamentacaoLegal() As String
    FundamentacaoLegal = p_FundamentacaoLegal
End Property

Public Property Let FundamentacaoLegal(ByVal Value As String)
    p_FundamentacaoLegal = Value
End Property

Public Property Get TipoAcaoFiscal() As String
    TipoAcaoFiscal = p_TipoAcaoFiscal
End Property

Public Property Let TipoAcaoFiscal(ByVal Value As String)
    p_TipoAcaoFiscal = Value
End Property

Public Property Get LavraturaLocal() As String
    LavraturaLocal = p_LavraturaLocal
End Property

Public Property Let LavraturaLocal(ByVal Value As String)
    p_LavraturaLocal = Value
End Property

Public Property Get LavraturaData() As String
    LavraturaData = p_LavraturaData
End Property

Public Property Let LavraturaData(ByVal Value As String)
    p_LavraturaData = Value
End Property

Public Property Get LavraturaHora() As String
    LavraturaHora = p_LavraturaHora
End Property

Public Property Let LavraturaHora(ByVal Value As String)
    p_LavraturaHora = Value
End Property

Public Property Get DescricaoMotivos() As String
    DescricaoMotivos = p_DescricaoMotivos
End Property

Public Property Let DescricaoMotivos(ByVal Value As String)
    p_DescricaoMotivos = Value
End Property

Public Property Get TipoProcedimento() As String
    TipoProcedimento = p_TipoProcedimento
End Property

Public Property Let TipoProcedimento(ByVal Value As String)
    p_TipoProcedimento = Value
End Property

