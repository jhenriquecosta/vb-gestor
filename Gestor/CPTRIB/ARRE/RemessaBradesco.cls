VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemessaBradesco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mIdentificacaoRegistro As String
Private mAgenciaDebito As String
Private mDigitoAgenciaDebito As String
Private mRazaoContaCorrente As String
Private mContaCorrente As String
Private mDigitoContaCorrente As String
Private mIdentificacaoEmpresaCedenteNoBanco As String
Private mNumeroControleParticipante As String
Private mCodigoBancoCamaraCompensacao As String
Private mIdentificacaoTituloBanco As String
Private mDigitoAutoConferenciaNossoNumero As String
Private mDescontoBonificacaoDia As String
Private mCondicaoParaEmissaoPapeladaCobranca As String
Private mEmitePapeletaDebitoAutomatico As String 'S/N
Private mIdentificacaoOperacaoBanco As String
Private mIndicadorRateioCredito As String
Private mEnderecamentoAvisoDebAutoCC As String
Private mIndicacaoOcorrencia As String
Private mNumeroDocumento As String
Private mDataVencimentoTitulo As String 'DDMMAA
Private mValorTitulo As String
Private mBancoEncarregadoCobranca As String
Private mAgenciaDepositaria As String
Private mEspecieTitulo As String
Private mIdentificacao As String
Private mDataEmissaoTitulo As String 'DDMMAA
Private mPrimeiraInstrucao As String
Private mSegundaInstrucao As String
Private mValorCobradoDiaAtraso As String
Private mDataLimiteConcessaoDesconto As String
Private mValorDesconto As String
Private mValorIOF As String
Private mValorAbatimento As String
Private mIdentificacaoTipoInscricaoSacado As String
Private mNumeroInscricaoSacado As String
Private mNomeSacado As String
Private mCampoMulta As String
Private mEnderecoCompleto As String
Private mPrimeiraMensagem As String
Private mCep As String
Private mSufixoCEP As String
Private mSacadorAvalistaSegundaMensagem As String
Private mNumeroSequencialRegistro As String
Private mPercentualMulta As String
Private mFiller02 As String
Private mFiller05 As String
Private mFiller10 As String
Private mFiller36 As String
Private mDebitoAutomatico As String
Private mInstrucao1 As String
Private mInstrucao2 As String
Private mBairro As String


Private mMunicipio As String
Private mUF As String

Public Property Get UF() As String
     UF = mUF
End Property
Public Property Let UF(ByVal Valor As String)
     mUF = Valor
End Property

Public Property Get Municipio() As String
     Municipio = mMunicipio
End Property
Public Property Let Municipio(ByVal Valor As String)
     mMunicipio = Valor
End Property



Public Property Get Bairro() As String
     Bairro = mBairro
End Property
Public Property Let Bairro(ByVal Valor As String)
     mBairro = Valor
End Property

Public Property Get Instrucao1() As String
     Instrucao1 = mInstrucao1
End Property
Public Property Let Instrucao1(ByVal Valor As String)
     mInstrucao1 = Valor
End Property


Public Property Get Instrucao2() As String
     Instrucao2 = mInstrucao2
End Property
Public Property Let Instrucao2(ByVal Valor As String)
     mInstrucao2 = Valor
End Property

Public Property Get DebitoAutomatico() As String
     DebitoAutomatico = mDebitoAutomatico
End Property
Public Property Let DebitoAutomatico(ByVal Valor As String)
     mDebitoAutomatico = Valor
End Property

Public Property Get Filler02() As String
     Filler02 = mFiller02
End Property
Public Property Let Filler02(ByVal Valor As String)
     mFiller02 = Valor
End Property

Public Property Get Filler36() As String
     Filler36 = mFiller36
End Property
Public Property Let Filler36(ByVal Valor As String)
     mFiller36 = Valor
End Property

Public Property Get Filler10() As String
     Filler10 = mFiller10
End Property
Public Property Let Filler10(ByVal Valor As String)
     mFiller10 = Valor
End Property

Public Property Get Filler05() As String
     Filler05 = mFiller05
End Property
Public Property Let Filler05(ByVal Valor As String)
     mFiller05 = Valor
End Property

Public Property Get PercentualMulta() As String
     PercentualMulta = mPercentualMulta
End Property
Public Property Let PercentualMulta(ByVal Valor As String)
     mPercentualMulta = Valor
End Property
Public Property Get CampoMulta() As String
     CampoMulta = mCampoMulta
End Property
Public Property Let CampoMulta(ByVal Valor As String)
     mCampoMulta = Valor
End Property
Public Property Get IdentificacaoRegistro() As String
     IdentificacaoRegistro = mIdentificacaoRegistro
End Property
Public Property Let IdentificacaoRegistro(ByVal Valor As String)
     mIdentificacaoRegistro = Valor
End Property
Public Property Get AgenciaDebito() As String
     AgenciaDebito = mAgenciaDebito
End Property
Public Property Let AgenciaDebito(ByVal Valor As String)
     mAgenciaDebito = Valor
End Property
Public Property Get DigitoAgenciaDebito() As String
     DigitoAgenciaDebito = mDigitoAgenciaDebito
End Property
Public Property Let DigitoAgenciaDebito(ByVal Valor As String)
     mDigitoAgenciaDebito = Valor
End Property
Public Property Get RazaoContaCorrente() As String
     RazaoContaCorrente = mRazaoContaCorrente
End Property
Public Property Let RazaoContaCorrente(ByVal Valor As String)
     mRazaoContaCorrente = Valor
End Property
Public Property Get ContaCorrente() As String
     ContaCorrente = mContaCorrente
End Property
Public Property Let ContaCorrente(ByVal Valor As String)
     mContaCorrente = Valor
End Property
Public Property Get DigitoContaCorrente() As String
     DigitoContaCorrente = mDigitoContaCorrente
End Property
Public Property Let DigitoContaCorrente(ByVal Valor As String)
     mDigitoContaCorrente = Valor
End Property
Public Property Get IdentificacaoEmpresaCedenteNoBanco() As String
     IdentificacaoEmpresaCedenteNoBanco = mIdentificacaoEmpresaCedenteNoBanco
End Property
Public Property Let IdentificacaoEmpresaCedenteNoBanco(ByVal Valor As String)
     mIdentificacaoEmpresaCedenteNoBanco = Valor
End Property
Public Property Get NumeroControleParticipante() As String
     NumeroControleParticipante = mNumeroControleParticipante
End Property
Public Property Let NumeroControleParticipante(ByVal Valor As String)
     mNumeroControleParticipante = Valor
End Property
Public Property Get CodigoBancoCamaraCompensacao() As String
     CodigoBancoCamaraCompensacao = mCodigoBancoCamaraCompensacao
End Property
Public Property Let CodigoBancoCamaraCompensacao(ByVal Valor As String)
     mCodigoBancoCamaraCompensacao = Valor
End Property
Public Property Get IdentificacaoTituloBanco() As String
     IdentificacaoTituloBanco = mIdentificacaoTituloBanco
End Property
Public Property Let IdentificacaoTituloBanco(ByVal Valor As String)
     mIdentificacaoTituloBanco = Valor
End Property
Public Property Get DigitoAutoConferenciaNossoNumero() As String
     DigitoAutoConferenciaNossoNumero = mDigitoAutoConferenciaNossoNumero
End Property
Public Property Let DigitoAutoConferenciaNossoNumero(ByVal Valor As String)
     mDigitoAutoConferenciaNossoNumero = Valor
End Property
Public Property Get DescontoBonificacaoDia() As String
     DescontoBonificacaoDia = mDescontoBonificacaoDia
End Property
Public Property Let DescontoBonificacaoDia(ByVal Valor As String)
     mDescontoBonificacaoDia = Valor
End Property
Public Property Get CondicaoParaEmissaoPapeladaCobranca() As String
     CondicaoParaEmissaoPapeladaCobranca = mCondicaoParaEmissaoPapeladaCobranca
End Property
Public Property Let CondicaoParaEmissaoPapeladaCobranca(ByVal Valor As String)
     mCondicaoParaEmissaoPapeladaCobranca = Valor
End Property
Public Property Get EmitePapeletaDebitoAutomatico() As String
     EmitePapeletaDebitoAutomatico = mEmitePapeletaDebitoAutomatico
End Property
Public Property Let EmitePapeletaDebitoAutomatico(ByVal Valor As String)
     mEmitePapeletaDebitoAutomatico = Valor
End Property
Public Property Get IdentificacaoOperacaoBanco() As String
     IdentificacaoOperacaoBanco = mIdentificacaoOperacaoBanco
End Property
Public Property Let IdentificacaoOperacaoBanco(ByVal Valor As String)
     mIdentificacaoOperacaoBanco = Valor
End Property
Public Property Get IndicadorRateioCredito() As String
     IndicadorRateioCredito = mIndicadorRateioCredito
End Property
Public Property Let IndicadorRateioCredito(ByVal Valor As String)
     mIndicadorRateioCredito = Valor
End Property
Public Property Get EnderecamentoAvisoDebAutoCC() As String
     EnderecamentoAvisoDebAutoCC = mEnderecamentoAvisoDebAutoCC
End Property
Public Property Let EnderecamentoAvisoDebAutoCC(ByVal Valor As String)
     mEnderecamentoAvisoDebAutoCC = Valor
End Property
Public Property Get IndicacaoOcorrencia() As String
     IndicacaoOcorrencia = mIndicacaoOcorrencia
End Property
Public Property Let IndicacaoOcorrencia(ByVal Valor As String)
     mIndicacaoOcorrencia = Valor
End Property
Public Property Get NumeroDocumento() As String
     NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal Valor As String)
     mNumeroDocumento = Valor
End Property
Public Property Get DataVencimentoTitulo() As String
     DataVencimentoTitulo = mDataVencimentoTitulo
End Property
Public Property Let DataVencimentoTitulo(ByVal Valor As String)
     mDataVencimentoTitulo = Valor
End Property
Public Property Get ValorTitulo() As String
     ValorTitulo = mValorTitulo
End Property
Public Property Let ValorTitulo(ByVal Valor As String)
     mValorTitulo = Valor
End Property
Public Property Get BancoEncarregadoCobranca() As String
     BancoEncarregadoCobranca = mBancoEncarregadoCobranca
End Property
Public Property Let BancoEncarregadoCobranca(ByVal Valor As String)
     mBancoEncarregadoCobranca = Valor
End Property
Public Property Get AgenciaDepositaria() As String
     AgenciaDepositaria = mAgenciaDepositaria
End Property
Public Property Let AgenciaDepositaria(ByVal Valor As String)
     mAgenciaDepositaria = Valor
End Property
Public Property Get EspecieTitulo() As String
     EspecieTitulo = mEspecieTitulo
End Property
Public Property Let EspecieTitulo(ByVal Valor As String)
     mEspecieTitulo = Valor
End Property
Public Property Get Identificacao() As String
     Identificacao = mIdentificacao
End Property
Public Property Let Identificacao(ByVal Valor As String)
     mIdentificacao = Valor
End Property
Public Property Get DataEmissaoTitulo() As String
     DataEmissaoTitulo = mDataEmissaoTitulo
End Property
Public Property Let DataEmissaoTitulo(ByVal Valor As String)
     mDataEmissaoTitulo = Valor
End Property
Public Property Get PrimeiraInstrucao() As String
     PrimeiraInstrucao = mPrimeiraInstrucao
End Property
Public Property Let PrimeiraInstrucao(ByVal Valor As String)
     mPrimeiraInstrucao = Valor
End Property
Public Property Get SegundaInstrucao() As String
     SegundaInstrucao = mSegundaInstrucao
End Property
Public Property Let SegundaInstrucao(ByVal Valor As String)
     mSegundaInstrucao = Valor
End Property
Public Property Get ValorCobradoDiaAtraso() As String
     ValorCobradoDiaAtraso = mValorCobradoDiaAtraso
End Property
Public Property Let ValorCobradoDiaAtraso(ByVal Valor As String)
     mValorCobradoDiaAtraso = Valor
End Property
Public Property Get DataLimiteConcessaoDesconto() As String
     DataLimiteConcessaoDesconto = mDataLimiteConcessaoDesconto
End Property
Public Property Let DataLimiteConcessaoDesconto(ByVal Valor As String)
     mDataLimiteConcessaoDesconto = Valor
End Property
Public Property Get ValorDesconto() As String
     ValorDesconto = mValorDesconto
End Property
Public Property Let ValorDesconto(ByVal Valor As String)
     mValorDesconto = Valor
End Property
Public Property Get ValorIOF() As String
     ValorIOF = mValorIOF
End Property
Public Property Let ValorIOF(ByVal Valor As String)
     mValorIOF = Valor
End Property
Public Property Get ValorAbatimento() As String
     ValorAbatimento = mValorAbatimento
End Property
Public Property Let ValorAbatimento(ByVal Valor As String)
     mValorAbatimento = Valor
End Property
Public Property Get IdentificacaoTipoInscricaoSacado() As String
     IdentificacaoTipoInscricaoSacado = mIdentificacaoTipoInscricaoSacado
End Property
Public Property Let IdentificacaoTipoInscricaoSacado(ByVal Valor As String)
     mIdentificacaoTipoInscricaoSacado = Valor
End Property
Public Property Get NumeroInscricaoSacado() As String
     NumeroInscricaoSacado = mNumeroInscricaoSacado
End Property
Public Property Let NumeroInscricaoSacado(ByVal Valor As String)
     mNumeroInscricaoSacado = Valor
End Property
Public Property Get NomeSacado() As String
     NomeSacado = mNomeSacado
End Property
Public Property Let NomeSacado(ByVal Valor As String)
     mNomeSacado = Valor
End Property
Public Property Get EnderecoCompleto() As String
     EnderecoCompleto = mEnderecoCompleto
End Property
Public Property Let EnderecoCompleto(ByVal Valor As String)
     mEnderecoCompleto = Valor
End Property
Public Property Get PrimeiraMensagem() As String
     PrimeiraMensagem = mPrimeiraMensagem
End Property
Public Property Let PrimeiraMensagem(ByVal Valor As String)
     mPrimeiraMensagem = Valor
End Property
Public Property Get Cep() As String
     Cep = mCep
End Property
Public Property Let Cep(ByVal Valor As String)
     mCep = Valor
End Property
Public Property Get SufixoCEP() As String
     SufixoCEP = mSufixoCEP
End Property
Public Property Let SufixoCEP(ByVal Valor As String)
     mSufixoCEP = Valor
End Property
Public Property Get SacadorAvalistaSegundaMensagem() As String
     SacadorAvalistaSegundaMensagem = mSacadorAvalistaSegundaMensagem
End Property
Public Property Let SacadorAvalistaSegundaMensagem(ByVal Valor As String)
     mSacadorAvalistaSegundaMensagem = Valor
End Property
Public Property Get NumeroSequencialRegistro() As String
     NumeroSequencialRegistro = mNumeroSequencialRegistro
End Property
Public Property Let NumeroSequencialRegistro(ByVal Valor As String)
     mNumeroSequencialRegistro = Valor
End Property
Public Function gerarDetalheRemessa(diretorio As String, arquivo As String) As String
    Dim NumeroArquivo         As Integer
    Dim branco95_104 As String, branco107_108 As String
    branco95_104 = repete(" ", 10)
    branco107_108 = repete(" ", 2)
    NumeroArquivo = FreeFile
    Open diretorio & arquivo For Append Access Write As #NumeroArquivo
    Print #NumeroArquivo, Me.IdentificacaoRegistro & Me.AgenciaDebito & Me.DigitoAgenciaDebito & Me.RazaoContaCorrente _
    & Me.ContaCorrente & Me.DigitoContaCorrente & Me.IdentificacaoEmpresaCedenteNoBanco & Me.NumeroControleParticipante _
    & Me.CodigoBancoCamaraCompensacao & Me.CampoMulta & Me.PercentualMulta & Me.IdentificacaoTituloBanco & Me.DigitoAutoConferenciaNossoNumero _
    & Me.DescontoBonificacaoDia & Me.CondicaoParaEmissaoPapeladaCobranca & Me.EmitePapeletaDebitoAutomatico & branco95_104 _
    & Me.IndicadorRateioCredito & Me.EnderecamentoAvisoDebAutoCC & branco107_108 & Me.IndicacaoOcorrencia & Me.NumeroDocumento _
    & Me.DataVencimentoTitulo & Me.ValorTitulo & Me.BancoEncarregadoCobranca & Me.AgenciaDepositaria & Me.EspecieTitulo _
    & Me.Identificacao & Me.DataEmissaoTitulo & Me.PrimeiraInstrucao & Me.SegundaInstrucao & Me.ValorCobradoDiaAtraso & Me.DataLimiteConcessaoDesconto _
    & Me.ValorDesconto & Me.ValorIOF & Me.ValorAbatimento & Me.IdentificacaoTipoInscricaoSacado & Me.NumeroInscricaoSacado & Me.NomeSacado _
    & Me.EnderecoCompleto & Me.PrimeiraMensagem & Me.Cep & Me.SufixoCEP & Me.SacadorAvalistaSegundaMensagem & Me.NumeroSequencialRegistro
    Close NumeroArquivo
    gerarDetalheRemessa = arquivo
End Function

Public Function gerarDetalheRemessaBasa(diretorio As String, arquivo As String) As String
    Dim NumeroArquivo         As Integer
    Dim branco95_104 As String, branco107_108 As String
    branco95_104 = repete(" ", 10)
    branco107_108 = repete(" ", 2)
    NumeroArquivo = FreeFile
    Open diretorio & arquivo For Append Access Write As #NumeroArquivo
    Print #NumeroArquivo, Me.IdentificacaoRegistro & Me.Filler02 & Me.IdentificacaoEmpresaCedenteNoBanco & Me.NumeroControleParticipante & Me.Filler05 & Me.IdentificacaoTituloBanco & Me.DescontoBonificacaoDia _
            & Me.CondicaoParaEmissaoPapeladaCobranca & Me.DebitoAutomatico & Me.Filler10 & Me.IndicacaoOcorrencia & Me.NumeroDocumento & Me.DataVencimentoTitulo & Me.ValorTitulo & Me.BancoEncarregadoCobranca _
            & Me.AgenciaDepositaria & Me.EspecieTitulo & Me.Identificacao & Me.DataEmissaoTitulo & Me.Instrucao1 & Me.Instrucao2 & Me.ValorCobradoDiaAtraso & Me.DataLimiteConcessaoDesconto & Me.ValorDesconto & Me.ValorIOF _
            & Me.ValorAbatimento & Me.IdentificacaoTipoInscricaoSacado & Me.NumeroInscricaoSacado & Me.NomeSacado & Me.EnderecoCompleto & Me.Bairro & Me.Cep & Me.SufixoCEP & Me.Municipio & Me.UF & Me.Filler36 & Me.NumeroSequencialRegistro
    Close NumeroArquivo
    gerarDetalheRemessaBasa = arquivo
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
Public Function gerarDigitoConferencia82(carteira As String) As String
    'carteira formato  (00)
    Dim x As Integer, resultado As Integer, digito As String
    Dim nossoNumero As String, numero As String
    nossoNumero = Format(carteira, "00") & Me.IdentificacaoTituloBanco
    numero = 0
    resultado = 0
    For x = 1 To Len(nossoNumero)
        numero = Mid(nossoNumero, x, 1)
        numero = Format(numero * numeroParaMultiplicacao(x), "00")
        resultado = resultado + numero
    Next x
    digito = resultado Mod 11 ' de acordo com o documento
    If digito = 1 Then
        digito = "P"
    ElseIf digito = 0 Then
        digito = 0
    Else
        digito = 11 - digito
    End If
    
    gerarDigitoConferencia82 = digito
End Function
Private Function numeroParaMultiplicacao(Posicao As Integer) As Integer
        Dim numero As Integer
        If Posicao = 1 Or Posicao = 7 Or Posicao = 13 Then
            numero = 2
        ElseIf Posicao = 2 Or Posicao = 8 Then
            numero = 7
        ElseIf Posicao = 3 Or Posicao = 9 Then
            numero = 6
        ElseIf Posicao = 4 Or Posicao = 10 Then
            numero = 5
        ElseIf Posicao = 5 Or Posicao = 11 Then
            numero = 4
        ElseIf Posicao = 6 Or Posicao = 12 Then
            numero = 3
        End If
        numeroParaMultiplicacao = numero
End Function


