VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Remessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Propriedades do HEADER
Private h_Versao                                                              As String
Private h_Bancos                                                              As String
Private h_PrefixoAgencia                                                   As String
Private h_DV_PrefixoAgencia                                             As String

Private h_CodigoCedente                                                  As String
Private h_DigitoVerificadorCodigoCedente                          As String
Private h_NomeEmpresa                                                   As String
Private h_NumeroConvenente                                            As String
Private h_Banco                                                               As String
 
Private h_DataGravacao                                                   As String
Private h_SequencialRemessa                                           As String
Private Const h_SequencialRegistro = "000001"
Private hSequencialCabecalho As String
Private Const h_DadosPadrao = "01REMESSA01COBRANCA       "
'Propriedades do DETALHE
Private d_TipoInscricaoCedente                                         As String
Private d_InscricaoCedente                                               As String
Private d_PrefixoAgencia                                                   As String
Private d_CodigoCedente                                                  As String
Private d_DV_CodigoCedente                                            As String
Private d_NumeroConvenio                                               As String
Private d_NumeroControleParticipante                                As String
Private d_NossoNumero                                                     As String
Private d_SeuNumero                                                        As String
Private d_DigitoVerificadorNossoNumero                             As String
Private d_NumeroPrestacao                                               As String
Private d_IndicativoSacador                                               As String
Private d_PrefixoTitulo                                                       As String
Private d_VariacaoCaretira                                                 As String
Private d_ContaCaucao                                                      As String
Private d_CodigoResponsabilidade                                      As String
Private d_DVCodigoResponsabilidade                                  As String
Private d_NumeroBordero                                                  As String
Private d_NumeroCarteira                                                  As String
Private d_Carteira                                                             As String
Private d_Comando                                                           As String
Private d_DataVencimento                                                 As String
Private d_ValorTitulo                                                         As String
Private d_NumeroBanco                                                     As String
Private d_PrefixoAgenciaCobradora                                    As String
Private d_DV_Pref_Agencia_Cobradora                               As String
Private d_EspecieTitulo                                                     As String
Private d_Aceite                                                                As String
Private d_DataEmissao                                                     As String
Private d_PrimeiraInstrucaoCodificada                                As String
Private d_SegundaInstrucaoCodificada                               As String
Private d_JurosMoraDia                                                     As String
Private d_DataLimiteConcessaoDesconto                            As String
Private d_ValorDesconto                                                    As String
Private d_CampoEspecial_1                                               As String
Private d_ValorAbatimento                                                 As String
Private d_TipoInscricaoSacado                                           As String
Private d_CPF_CNPJ_Sacado                                              As String
Private d_NomeSacado                                                      As String
Private d_EnderecoSacado                                                 As String
Private d_CepSacado                                                         As String
Private d_CidadeSacado                                                     As String
Private d_UFSacado                                                           As String
Private d_Observacoes                                                       As String
Private d_DiasParaProtesto                                                 As String
Private d_Correlatico As String
Private CodBarra As New CodigoDeBarra

Public Enum etaTIpoArquivo
    etaRemessa = 1
    etaRetorno = 2
End Enum
Public Property Get SequencialRegistro() As String
    SequencialRegistro = d_Correlatico
End Property

Public Property Let SequencialRegistro(ByVal Valor As String)
  d_Correlatico = Format(Valor, String(6, "0"))
End Property
Public Property Get NumeroConvenente() As String
    NumeroConvenente = h_NumeroConvenente
End Property
Public Property Let NumeroConvenente(ByVal Valor As String)
  h_NumeroConvenente = Format(Valor, String(6, "0"))
End Property
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
Public Property Get DigitoVerificadorPrefixoAgencia() As String
    DigitoVerificadorPrefixoAgencia = h_DV_PrefixoAgencia
End Property
Public Property Let DigitoVerificadorPrefixoAgencia(ByVal Valor As String)
    h_DV_PrefixoAgencia = Valor
End Property
Public Property Get CodigoCedente() As String
    CodigoCedente = h_CodigoCedente
End Property
Public Property Let CodigoCedente(ByVal Valor As String)
    h_CodigoCedente = Format(Valor, String(8, "0"))
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
    h_NomeEmpresa = UCase(Left(Valor, 30))
    h_NomeEmpresa = h_NomeEmpresa & String(30 - Len(h_NomeEmpresa), " ")
End Property
Public Property Get Banco() As String
    Banco = h_Banco
End Property
Public Property Let DataGravacao(ByVal Valor As String)
    h_DataGravacao = Format(Day(Valor), "00") & Format(Month(Valor), "00") & Right(Valor, 2)
End Property
Public Property Get DataGravacao() As String
    DataGravacao = h_DataGravacao
End Property
Public Sub SequencialRemessa(Bdados As VSDados)
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Correlativo As String
    Dim NomeImposto As String
        
    
        
    Sql = "UPDATE Tab_Num_Correlativo set tnc_sequencia =  tnc_sequencia + 1 where tnc_tipo_operacao = 17"
    
    
    Bdados.Executa Sql
     Sql = "SELECT tnc_sequencia"
     Sql = Sql & " from tab_num_correlativo "
     Sql = Sql & " where tnc_tipo_operacao = 17"
    
    If Bdados.AbreTabela(Sql, rs) Then
        h_SequencialRemessa = Trim(Correlativo) & Format(IIf(IsNull(rs(0)), 1, rs(0)), "000000")
    End If
    Bdados.FechaTabela rs
End Sub
Public Property Let TipoInscricaoCedente(ByVal Valor As String)
    d_TipoInscricaoCedente = Format(Valor, String(2, "0"))
End Property
Public Property Get TipoInscricaoCedente() As String
    TipoInscricaoCedente = d_TipoInscricaoCedente
End Property
Public Property Let InscricaoCedente(ByVal Valor As String)
    d_InscricaoCedente = Format(TiraTudo(Valor), String(14, "0"))
End Property
Public Property Get InscricaoCedente() As String
    InscricaoCedente = d_InscricaoCedente
End Property
Public Property Let PrefixoAgencia(ByVal Valor As String)
    h_PrefixoAgencia = Format(Valor, String(4, "0"))
End Property
Public Property Get PrefixoAgencia() As String
    PrefixoAgencia = h_PrefixoAgencia
End Property
Public Property Let NumeroConvenio(ByVal Valor As String)
    d_NumeroConvenio = Format(Valor, String(6, "0"))
End Property
Public Property Get NumeroConvenio() As String
    NumeroConvenio = d_NumeroConvenio
End Property
Public Property Let NumeroControleParticipante(ByVal Valor As String)
    d_NumeroControleParticipante = Valor
    d_NumeroControleParticipante = NumeroControleParticipante & String(25 - Len(d_NumeroControleParticipante), "0")
End Property
Public Property Get NumeroControleParticipante() As String
    NumeroControleParticipante = d_NumeroControleParticipante
End Property
Public Property Let SeuNumero(ByVal Valor As String)
    d_SeuNumero = Format(Valor, String(10, "0"))
End Property
Public Property Get SeuNumero() As String
    SeuNumero = d_SeuNumero
End Property
Public Property Let NossoNumero(ByVal Valor As String)
    If d_NumeroCarteira = 11 Or d_NumeroCarteira = 31 Or d_NumeroCarteira = 51 Then
        d_NossoNumero = 0
    Else
        d_NossoNumero = d_NumeroConvenio & Valor
    End If
    Call DigitoVerificadorNossoNumero
End Property
Public Property Get NossoNumero() As String
    NossoNumero = d_NossoNumero
End Property
Public Sub DigitoVerificadorNossoNumero()
    If d_NumeroCarteira = 11 Or d_NumeroCarteira = 31 Or d_NumeroCarteira = 51 Then
        d_DigitoVerificadorNossoNumero = 0
    Else
'        d_DigitoVerificadorNossoNumero = CodBarra.CalculaDV(d_NossoNumero, etmModulo11)
        d_DigitoVerificadorNossoNumero = CodBarra.Calculo_DV11(d_NossoNumero)
    End If
End Sub
Public Property Let NumeroPrestacao(ByVal Valor As String)
    d_NumeroPrestacao = Format(Valor, String(2, "0"))
End Property
Public Property Get NumeroPrestacao() As String
    NumeroPrestacao = d_NumeroPrestacao
End Property
Public Property Let IndicativoSacador(ByVal Valor As String)
    If Valor = "" Then
        d_IndicativoSacador = " "
    Else
        d_IndicativoSacador = Valor
    End If
End Property
Public Property Get IndicativoSacador() As String
    IndicativoSacador = d_IndicativoSacador
End Property
Public Property Let PrefixoTitulo(ByVal Valor As String)
    If Val(Valor) = 31 Or Val(Valor) = 51 Then
        d_PrefixoTitulo = "SD"
    ElseIf Val(Valor) = 12 Then
        d_PrefixoTitulo = "AIU"
    Else
        d_PrefixoTitulo = "AI"
    End If
End Property
Public Property Get PrefixoTitulo() As String
    PrefixoTitulo = d_PrefixoTitulo
End Property
Public Property Let VariacaoCaretira(ByVal Valor As String)
    If Len(Valor) = 3 Then
       d_VariacaoCaretira = Valor
    ElseIf Len(Valor) = 2 Then
        d_VariacaoCaretira = " " & Valor
    ElseIf Len(Valor) = 1 Then
        d_VariacaoCaretira = "  " & Valor
    End If
End Property
Public Property Get VariacaoCaretira() As String
    VariacaoCaretira = d_VariacaoCaretira
End Property
Public Property Let ContaCaucao(ByVal Valor As String)
    d_ContaCaucao = Valor
End Property
Public Property Get ContaCaucao() As String
    ContaCaucao = d_ContaCaucao
End Property
Public Property Let CodigoResponsabilidade(ByVal Valor As String)
    d_CodigoResponsabilidade = Format(Valor, String(5, "0"))
End Property
Public Property Get CodigoResponsabilidade() As String
    CodigoResponsabilidade = d_CodigoResponsabilidade
End Property
Public Property Let DVCodigoResponsabilidade(ByVal Valor As String)
    d_DVCodigoResponsabilidade = Valor
End Property
Public Property Get DVCodigoResponsabilidade() As String
    DVCodigoResponsabilidade = d_DVCodigoResponsabilidade
End Property
Public Property Let NumeroBordero(ByVal Valor As String)
    d_NumeroBordero = Format(Valor, String(7, "0"))
End Property
Public Property Get NumeroBordero() As String
    NumeroBordero = d_NumeroBordero
End Property
Public Property Let NumeroCarteira(ByVal Valor As String)
    d_NumeroCarteira = Format(Valor, String(2, "0"))
End Property
Public Property Get NumeroCarteira() As String
    NumeroCarteira = d_NumeroCarteira
End Property
Public Property Let Carteira(ByVal Valor As String)
    If Trim(Valor) <> "" Then
        d_Carteira = Format(Valor, String(5, "0"))
    Else
        d_Carteira = String(5, " ")
    End If
End Property
Public Property Get Carteira() As String
    Carteira = d_Carteira
End Property
Public Property Let Comando(ByVal Valor As String)
    d_Comando = Format(Valor, String(2, "0"))
    If Format(Valor, "00") = "99" Then
        d_Comando = "00"
    End If
End Property
Public Property Get Comando() As String
    Comando = d_Comando
End Property
Public Property Let DataVencimento(ByVal Valor As String)
    d_DataVencimento = Format(Day(Valor), "00") & Format(Month(Valor), "00") & Right(Valor, 2)
End Property
Public Property Get DataVencimento() As String
    DataVencimento = d_DataVencimento
End Property
Public Property Let ValorTitulo(ByVal Valor As String)
    Dim Pos As Integer
    Valor = Format(Valor, Const_Monetario)
    Pos = PosPic(Valor, ",")
    d_ValorTitulo = Format(TiraTudo(Left(Valor, Pos - 1)), String(11, "0")) & Right(Valor, 2)
End Property
Public Property Get ValorTitulo() As String
    ValorTitulo = d_ValorTitulo
End Property
Public Property Let NumeroBanco(ByVal Valor As String)
    d_NumeroBanco = Format(Valor, String(3, "0"))
End Property
Public Property Get NumeroBanco() As String
    NumeroBanco = d_NumeroBanco
End Property
Public Property Let PrefixoAgenciaCobradora(ByVal Valor As String)
    d_PrefixoAgenciaCobradora = Format(Valor, String(4, "0"))
End Property
Public Property Get PrefixoAgenciaCobradora() As String
    PrefixoAgenciaCobradora = d_PrefixoAgenciaCobradora
End Property
Public Property Let DV_Pref_Agencia_Cobradora(ByVal Valor As String)
    d_DV_Pref_Agencia_Cobradora = Valor
End Property
Public Property Get DV_Pref_Agencia_Cobradora() As String
    DV_Pref_Agencia_Cobradora = d_DV_Pref_Agencia_Cobradora
End Property
Public Property Let EspecieTitulo(ByVal Valor As String)
    d_EspecieTitulo = Format(Valor, String(2, "0"))
End Property
Public Property Get EspecieTitulo() As String
    EspecieTitulo = d_EspecieTitulo
End Property
Public Property Let Aceite(ByVal Valor As String)
    d_Aceite = UCase(Left(Valor, 1))
End Property
Public Property Get Aceite() As String
    Aceite = d_Aceite
End Property
Public Property Let DataEmissao(ByVal Valor As String)
    d_DataEmissao = Format(Day(Valor), "00") & Format(Month(Valor), "00") & Right(Valor, 2)
End Property
Public Property Get DataEmissao() As String
    DataEmissao = d_DataEmissao
End Property
Public Property Let PrimeiraInstrucaoCodificada(ByVal Valor As String)
    d_PrimeiraInstrucaoCodificada = Format(Valor, String(2, "0"))
End Property
Public Property Get PrimeiraInstrucaoCodificada() As String
    PrimeiraInstrucaoCodificada = d_PrimeiraInstrucaoCodificada
End Property
Public Property Let SegundaInstrucaoCodificada(ByVal Valor As String)
    d_SegundaInstrucaoCodificada = Format(Valor, String(2, "0"))
End Property
Public Property Get SegundaInstrucaoCodificada() As String
    SegundaInstrucaoCodificada = d_SegundaInstrucaoCodificada
End Property
Public Property Let JurosMoraDia(ByVal Valor As String)
    Dim Pos As Integer
    Valor = Format(Valor, Const_Monetario)
    Pos = PosPic(Valor, ",")
    d_JurosMoraDia = Format(TiraTudo(Left(Valor, Pos - 1)), String(11, "0")) & Right(Valor, 2)
End Property
Public Property Get JurosMoraDia() As String
    JurosMoraDia = d_JurosMoraDia
End Property
Public Property Let DataLimiteConcessaoDesconto(ByVal Valor As String)
    d_DataLimiteConcessaoDesconto = Format(Day(Valor), "00") & Format(Month(Valor), "00") & Right(Valor, 2)
End Property
Public Property Get DataLimiteConcessaoDesconto() As String
    DataLimiteConcessaoDesconto = d_DataLimiteConcessaoDesconto
End Property
Public Property Let ValorDesconto(ByVal Valor As String)
    Dim Pos As Integer
    Valor = Format(Valor, Const_Monetario)
    Pos = PosPic(Valor, ",")
    d_ValorDesconto = Format(TiraTudo(Left(Valor, Pos - 1)), String(11, "0")) & Right(Valor, 2)
End Property
Public Property Get ValorDesconto() As String
    ValorDesconto = d_ValorDesconto
End Property
Public Property Let CampoEspecial_1(ByVal Valor As String)
    d_CampoEspecial_1 = Valor
End Property
Public Property Get CampoEspecial_1() As String
    CampoEspecial_1 = d_CampoEspecial_1
End Property
Public Property Let ValorAbatimento(ByVal Valor As String)
    Dim Pos As Integer
    Pos = PosPic(Valor, ",")
    d_ValorAbatimento = Format(TiraTudo(Left(Valor, Pos - 1)), String(11, "0")) & Right(Valor, 2)
End Property
Public Property Get ValorAbatimento() As String
    ValorAbatimento = d_ValorAbatimento
End Property
Public Property Let TipoInscricaoSacado(ByVal Valor As String)
    d_TipoInscricaoSacado = Valor
End Property
Public Property Get TipoInscricaoSacado() As String
    TipoInscricaoSacado = d_TipoInscricaoSacado
End Property
Public Property Let CPF_CNPJ_Sacado(ByVal Valor As String)
    d_CPF_CNPJ_Sacado = Format(TiraTudo(Valor), String(14, "0"))
End Property
Public Property Get CPF_CNPJ_Sacado() As String
    CPF_CNPJ_Sacado = d_CPF_CNPJ_Sacado
End Property
Public Property Let NomeSacado(ByVal Valor As String)
    d_NomeSacado = Valor & String(37 - Len(Trim(Left(Valor, 37))), " ")
    d_NomeSacado = Left(d_NomeSacado, 37)
End Property
Public Property Get NomeSacado() As String
    NomeSacado = d_NomeSacado
End Property
Public Property Let EnderecoSacado(ByVal Valor As String)
    d_EnderecoSacado = Valor & String(37 - Len(Left(Valor, 37)), " ")
    d_EnderecoSacado = Left(d_EnderecoSacado, 37)
End Property
Public Property Get EnderecoSacado() As String
    EnderecoSacado = d_EnderecoSacado
End Property
Public Property Let CepSacado(ByVal Valor As String)
    If Valor = "" Then
        Valor = "0"
    End If
    d_CepSacado = Format(TiraTudo(Valor), String(8, "0"))
End Property
Public Property Get CepSacado() As String
    CepSacado = d_CepSacado
End Property
Public Property Let CidadeSacado(ByVal Valor As String)
    d_CidadeSacado = Valor & String(15 - Len(Valor), " ")
End Property
Public Property Get CidadeSacado() As String
    CidadeSacado = d_CidadeSacado
End Property
Public Property Let UFSacado(ByVal Valor As String)
    d_UFSacado = Valor
End Property
Public Property Get UFSacado() As String
    UFSacado = d_UFSacado
End Property
Public Property Let Observacoes(ByVal Valor As String)
    d_Observacoes = Valor & String(40 - Len(Trim(Valor)), " ")
End Property
Public Property Get Observacoes() As String
    Observacoes = d_Observacoes
End Property
Public Property Let DiasParaProtesto(ByVal Valor As String)
    If Valor <> "" Then
        d_DiasParaProtesto = Valor
    Else
        d_DiasParaProtesto = "  "
    End If
End Property
Public Function Gera_TRAILER(Arquivo As String)
    Dim NumeroArquivo         As Integer
    Dim Dados                      As String
    NumeroArquivo = FreeFile
    
    'Abro o arquivo como leitura para pegar o último registro...
    Open PegaConfiguracaoEscola(TEC_CAMINHO_REMESSA) & "\" & Arquivo For Input As #NumeroArquivo
    Do Until EOF(Numero)
            Line Input #Numero, Dados
    Loop
    Close NumeroArquivo
    'Gero o Trailer
    Open PegaConfiguracaoEscola(TEC_CAMINHO_REMESSA) & "\" & Arquivo For Append Access Write As #NumeroArquivo
    Print #NumeroArquivo, 9 & String(393, " ") & Format(Right(Dados, 6) + 1, String(6, "0"))
    Close NumeroArquivo
End Function
Public Function Gera_DETALHE(Arquivo As String)
    Dim NumeroArquivo         As Integer
    NumeroArquivo = FreeFile
    Open PegaConfiguracaoEscola(TEC_CAMINHO_REMESSA) & "\" & Arquivo For Append Access Write As #NumeroArquivo
    Print #NumeroArquivo, 1 & d_TipoInscricaoCedente & d_InscricaoCedente & h_PrefixoAgencia _
    ; h_DV_PrefixoAgencia & h_CodigoCedente & h_DigitoVerificadorCodigoCedente; d_NumeroConvenio _
    ; d_NumeroControleParticipante & d_NossoNumero & d_DigitoVerificadorNossoNumero _
    ; d_NumeroPrestacao & "00   " & d_IndicativoSacador & d_PrefixoTitulo & d_VariacaoCaretira _
    ; d_ContaCaucao & d_CodigoResponsabilidade & d_DVCodigoResponsabilidade _
    ; d_NumeroBordero & d_Carteira & d_NumeroCarteira _
    ; d_Comando & d_SeuNumero & d_DataVencimento & d_ValorTitulo & d_NumeroBanco _
    ; d_PrefixoAgenciaCobradora & d_DV_Pref_Agencia_Cobradora & d_EspecieTitulo _
    ; d_Aceite & d_DataEmissao & d_PrimeiraInstrucaoCodificada & d_SegundaInstrucaoCodificada & d_JurosMoraDia _
    ; d_DataLimiteConcessaoDesconto & d_ValorDesconto & d_CampoEspecial_1 & d_ValorAbatimento _
    ; d_TipoInscricaoSacado & d_CPF_CNPJ_Sacado & d_NomeSacado & "   " & d_EnderecoSacado & String(15, " ") & _
    d_CepSacado & d_CidadeSacado & d_UFSacado & d_Observacoes & d_DiasParaProtesto & " " & d_Correlatico
    Close NumeroArquivo
End Function
Public Function Gera_HEADER() As String
    Dim Arquivo         As Integer
    Dim NomeArquio   As String
    Arquivo = FreeFile
    NomeArquio = Format(Day(Date), "00") & Format(Month(Date), "00") & Year(Date) & " Hora" & Hour(Time) & "_" & Second(Time) & " CBR641.rem"
    Open PegaConfiguracaoEscola(TEC_CAMINHO_REMESSA) & NomeArquio For Output Access Write As #Arquivo
    Print #Arquivo, h_DadosPadrao & h_PrefixoAgencia & h_DV_PrefixoAgencia & h_CodigoCedente & h_DigitoVerificadorCodigoCedente & h_NumeroConvenente & NomeEmpresa & "001BANCO DO BRASIL" & h_DataGravacao & Format(h_SequencialRemessa, "0000000") & String(276, " ") & String(2, " ") & String(3, " ") & String(6, " ") & h_SequencialRegistro
    Close Arquivo
    Gera_HEADER = NomeArquio
End Function
Private Sub Class_Initialize()
    Set CodBarra = New CodigoDeBarra
End Sub
Private Sub Class_Terminate()
    Set CodBarra = Nothing
End Sub
