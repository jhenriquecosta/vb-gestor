VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  
 Option Explicit

Dim Imposto As New VSImposto
Dim Ecom As New VsContribuinte
Public Enum TipoDestinoImpressao
    tdiTela = 1
    tdiImpressora = 2
End Enum
Public Enum Tipo_Dam
    etdBranco = 1
    etdNormal = 2
End Enum

Public Enum Geracao_Taxa
    IPM = 1
    CIAP = 2
End Enum

'Private Function DescontoIptu(Parcela As Integer) As Double
'    Dim sql As String
'    Dim rs As VSRecordset
'    sql = "Select tge_nome from tab_geral where tge_tipo = 713 and tge_codigo =" & _
'        IIf(Parcela = 0, 1, 2)
'    If Bdados.AbreTabela(sql, rs) Then
'        DescontoIptu = rs(0)
'    End If
'    Bdados.FechaTabela rs
'End Function

Public Function GeraValoresImovel(InscCadastral As String, InscricaoMunicipal As String, Periodo As Integer, DtVenc As String) As String
    Dim RsComp As VSRecordset
    Dim RsTab As VSRecordset
    Dim RsDET As VSRecordset
    Dim RsAux As VSRecordset
    Dim RsEst As VSRecordset
    Dim Sql As String
    Dim Valor As Double
    Dim Unidade As Byte
    Dim Base As Double
    Dim Valor_Fator As Double
    Dim RsFat As VSRecordset
    Dim Valores As String
    Dim Campos As String
    
    Sql = "SELECT tdi_tco_cod_componente,tdi_valor_item FROM Tab_Detalhe_Imovel" & _
        " WHERE tdi_tim_ic = '" & InscCadastral & "'"
    If Bdados.AbreTabela(Sql, RsDET) Then
        RsDET.MoveFirst
        Do
            Sql = "SELECT tco_valor, tco_unid_moneta, tco_base_calculo,tco_componente_fator_calculo " & _
                " FROM tab_componente_avancado WHERE tco_cod_componente = " & RsDET(0) & _
                " and tco_tmu_cod_municipio =" & AplicacoesVTFuncoes.Codigo_Municipio
            If Bdados.AbreTabela(Sql, RsComp) Then
                'Dados dos Detalhes do Imovel
                Valor = IIf(IsNull(RsComp(0)), 0, RsComp(0))
                Unidade = IIf(IsNull(RsComp(1)), 0, RsComp(1))
                Base = IIf(IsNull(RsComp(2)), 0, RsComp(2))
                'Busco valor do Fator p/ Base de Calculo
                If RsComp(3) <> 0 And Not IsNull(RsComp(3)) Then
                    Sql = "SELECT tdi_valor_item FROM Tab_Detalhe_Imovel" & _
                        " WHERE tdi_tim_ic = '" & InscCadastral & "' and tdi_tco_cod_componente =" & RsComp(3)
                    If Bdados.AbreTabela(Sql, RsFat) Then
                        Valor_Fator = RsFat(0)
                    End If
                    Bdados.FechaTabela RsFat
                Else
                    Valor_Fator = Valor
                End If
                If Valor = 0 Then
                'Possui tabela estimada
                    Sql = "SELECT tce_valor FROM Tab_Componente_Estimado " & _
                    " WHERE tce_tco_cod_componente = " & RsDET(0) & " and " & _
                    " tce_limite = (select min(tce_limite) FROM Tab_Componente_Estimado " & _
                    " WHERE tce_tco_cod_componente =" & RsDET(0) & ") and " & _
                    " tce_limite >=" & Valor_Fator
                    If Bdados.AbreTabela(Sql, RsEst) Then
                        'Dados dos Valores Estimados
                        Valor = RsEst(0)
                    End If
                    Bdados.FechaTabela RsEst
                End If
                Select Case RsComp(1) 'Moeda -> 0 - Real, 1 - Ufir, 2 - Ufm, 3 - Perc
                    Case 0
                        Valor = Valor
                    Case 1, 2
                        Sql = "Select tmo_valor FROM Tab_Monetaria " & _
                            " WHERE tmo_unidade = 1 and tmo_data = (Select max(tmo_data) " & _
                            " FROM Tab_Monetaria)"
                        If Bdados.AbreTabela(Sql, RsAux) Then
                            Valor = Valor * CDbl(RsAux(0))
                            If RsComp(1) = "2" Then
                                Sql = "Select tmo_valor FROM Tab_Monetaria " & _
                                " WHERE tmo_unidade = 2 and tmo_data = (Select max(tmo_data) " & _
                                " FROM Tab_Monetaria)"
                                If Bdados.AbreTabela(Sql, RsAux) Then
                                    Valor = Valor * CDbl(RsAux(0))
                                Else
                                    Call Util.Avisa("Não existe valor de UFM cadastrado.")
                                End If
                            End If
                        Else
                            Call Util.Avisa("Não existe valor de UFIR cadastrado.")
                        End If
                        Bdados.FechaTabela RsAux
                    Case 3
                        Valor = (Valor_Fator * Base) * (Valor / 100)
                End Select
            End If
            'Gravo valores encontrados
            Valores = Bdados.PreparaValor(InscCadastral, RsDET(0), Bdados.Converte(Valor, TCDuplo))
            Campos = "tdi_tim_ic,tdi_tco_cod_componente,tdi_valor_calc"
            If Not Bdados.AbreTabela("Select tdi_tim_ic from Tab_Detalhe_Imovel where tdi_tim_ic='" & InscCadastral & "' and  tdi_tco_cod_componente =" & RsDET(0), RsAux) Then
                Call Bdados.InsereDados("Tab_Detalhe_Imovel", Valores)
            Else
                Call Bdados.AtualizaDados("Tab_Detalhe_Imovel", Valores, Campos, "tdi_tim_ic='" & InscCadastral & "' and  tdi_tco_cod_componente =" & RsDET(0))
            End If
            Bdados.FechaTabela RsAux
            RsDET.MoveNext
        Loop While Not RsDET.EOF
    End If
    Bdados.FechaTabela RsAux
    Bdados.FechaTabela RsComp
    Bdados.FechaTabela RsDET
    Bdados.FechaTabela RsFat
    Bdados.FechaTabela RsTab
End Function

Public Function PagouDarm(NumDam As String) As Boolean
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim SqlParc As String
    Sql = "SELECT tdr_tgt_cod_pagamento from Tab_Darm_Recebido where tdr_tgt_cod_pagamento = " & NumDam
    If Bdados.AbreTabela(Sql, Rs) Then
        PagouDarm = True
    Else
        SqlParc = "Select TPA_NUM_PARCELAMENTO from tab_parcelamento,TAB_GERACAO_TRIBUTO where TPA_TCI_IM= tgt_im" & _
            " and TPA_TIP_COD_IMPOSTO = tgt_tip_cod_imposto and TPA_PERIODO=" & _
            " tgt_periodo and TPA_TIM_IC= tgt_tim_ic and TPA_STATUS_PARCELAMENTO < 2 " & _
            " AND tgt_cod_pagamento=" & NumDam
        PagouDarm = Bdados.AbreTabela(SqlParc, Rs)
    End If
    Bdados.FechaTabela Rs
End Function

Public Sub ImprimeDam(Report As VSRelatorio, CodPagamento As String, InscMunicipal As String, RazaoSocial As String, _
CgcCpf As String, EnderecoContribuinte As String, InscCadastral As String, EnderecoImovel As String, CodTributo As String, _
Sigla As String, NomeImposto As String, Exercicio As String, Parcela As String, Conta As Integer, DataVencimento As String, _
BaseCalculo As String, ValorTributo As String, ValorMulta As String, ValorJuros As String, ByVal TaxaServico As Double, _
Desconto As String, AtividadeEconomica As String, Observacao As String, Optional Picture As Object, Optional NotaInicial As String, _
Optional NotaFinal As String, Optional MaterialICMS As String, Optional ValorMetro As Double, Optional TaxaParcela As Double, _
Optional AreaTotal As Double, Optional AreaConstruida As Double, Optional ValorTerreno As Double, Optional ValorEdificacao As Double, Optional Zona As Integer, Optional Destino As TipoDestinoImpressao = tdiTela, Optional Tipo As Tipo_Dam = etdNormal, Optional StringTaxa As String, Optional VencimentoNormal As String, Optional Correcao As Double, Optional StatusDivida As TipoStatusObrigacao, Optional DescontosAtualizacao As Double)
   On Error GoTo Trata
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim MaxCotas As Byte
    Dim CodBarra As New CodigoDeBarra
    Dim a As Byte
    Dim DataDesconto As String
    Dim CgcPref As String
    Dim LinhaDigitavel As String
    Dim InscricaoSP As String
    Dim ObsTributo As String
    Dim TipoIptu As Integer
    Dim Iptu As New VSIptu
    Dim Txea As Double
    Dim Tsd As Double
    Dim ValorCompleto As Double
    Dim TXDAM As Double
    Dim TxP As Double
    Dim NumParcelamento As String
    Dim Obrig As New Obrigacao
    Dim FatorAlvara As Double
    Dim fantasia As String
    Dim rc As VSRecordset
    Dim robss As VSRecordset
    If Len(Observacao) = 0 Then
        If Bdados.AbreTabela("SELECT TOC_OBS FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO=" & CodPagamento, robss) Then
            If Not IsNull(robss(0)) Then
                Observacao = robss(0)
            End If
        End If
    End If
    TipoIptu = Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0)
    InscMunicipal = Trim(InscMunicipal)
    InscCadastral = Trim(InscCadastral)
    DataDesconto = Temp.PegaParametro(Bdados, "DATA DESCONTO")
    CodTributo = Trim(CodTributo)
    If Trim(CodTributo) = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) And CInt(Nvl(Parcela, 0)) = 0 Then
        Txea = Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXEA"), ".", ","), 0)
        TxP = Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXP"), ".", ","), 0)
        Sql = "select tci_fantasia,tae_nome,tci_fator_alvara from tab_contribuinte,tab_atividade_economica  where tci_im = '" & InscMunicipal & "' and tci_tae_cae = tae_cae "
        If Bdados.AbreTabela(Sql, Rs) Then
            AtividadeEconomica = Trim(Rs!tae_nome) & " - " & Trim(Rs!tci_fantasia)
            FatorAlvara = Rs!tci_fator_alvara
            If Temp.PegaParametro(Bdados, "TXP ESTIMADA") = "SIM" Then
                If FatorAlvara > CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXP ESTIMADA LIMITE"), ".", ","), 0)) Then
                    TxP = TxP * CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXP ESTIMADA LIMITE VALOR"), ".", ","), 0))
                End If
            End If
        End If
    End If
    If Trim(CodTributo) = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) And CInt(Nvl(Parcela, 0)) = 0 Then
        Tsd = Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TSD"), ".", ","), 0)
    End If
    'VERIFICA SE O TRIBUTO SERÁ IMPRESSO COM TAXA
    Sql = "SELECT TPI_GERA_TAXA_IMPRESSAO "
    Sql = Sql & "  FROM TAB_PARAMETRO_IMPOSTO "
    Sql = Sql & " WHERE TPI_TIP_COD_IMPOSTO = '" & CodTributo & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
            If CInt(Nvl("" & Rs!TPI_GERA_TAXA_IMPRESSAO, 1)) = 1 Then
                TXDAM = TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXDAM"), 0), ".", ",")
            End If
        Else
            If CInt(Nvl("" & Rs!TPI_GERA_TAXA_IMPRESSAO, 1)) = 1 And CDbl(ValorTributo) > 0 Then
                TXDAM = TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXDAM"), 0), ".", ",")
            Else
'                TaxaServico = 0  CUIDADO COM ESSA TAXA
            End If
        End If
    End If
    TXDAM = Format(TXDAM, Const_Monetario)
    Bdados.FechaTabela Rs
    
    If CodTributo <> Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) Then
        Desconto = ValorTributo * CDbl(Nvl(Desconto, 0)) / 100
    Else
        If Verifica_Geracao_Taxas(CodPagamento) = IPM Then
            Desconto = Format((ValorTributo - TXDAM) * CDbl(Nvl(Desconto, 0)) / 100, Const_Monetario)
        Else
            Desconto = Format((ValorTributo) * CDbl(Nvl(Desconto, 0)) / 100, Const_Monetario)
        End If
    End If
    Desconto = Desconto + DescontosAtualizacao
    If TaxaServico = 0 Then
        TaxaServico = TXDAM + Tsd
    Else
        If Verifica_Geracao_Taxas(CodPagamento) = IPM Then
            TaxaServico = TaxaServico + TXDAM
        Else
            TaxaServico = TaxaServico + Tsd
        End If
    End If
    TaxaServico = TaxaServico + Txea + TxP
    If Trim(CodTributo) = Trim(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))) Then
        StringTaxa = " [TSD - R$" & Format(Tsd, Const_Monetario) & "]"
    End If
    With Report
        ' gerar o arquivo bmp da barra antes de chamar o rpt pois somente assim irá sair com o bmp correto
        If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
            If Not .DefinirArquivo(Bdados, App.Path & "\TDAM.rpt") Then Exit Sub
        Else
            If Not .DefinirArquivo(Bdados, App.Path & "\TDAM_.rpt") Then Exit Sub
        End If
        InscricaoSP = IIf(Trim(InscCadastral) <> "", Trim(InscCadastral), Trim(InscMunicipal))
        InscricaoSP = TiraPic(TiraPic(TiraPic(InscricaoSP, "."), "-"), "/")
        If Len(InscricaoSP) = 10 Then
            InscricaoSP = Left(InscMunicipal, 2) & "." & Mid(InscMunicipal, 3, 3) & "." & Mid(InscMunicipal, 6, 3) & Right(InscMunicipal, 3)
            .Formulas "endereco", IIf(Trim(EnderecoContribuinte) <> "", EnderecoContribuinte, EnderecoImovel)
        Else
            If Len(InscricaoSP) = 9 Or Len(InscricaoSP) = 15 Then
                InscricaoSP = IIf(Len(InscCadastral) = 15, Left(InscCadastral, 2) & "." & Mid(InscCadastral, 3, 2) & "." & Mid(InscCadastral, 5, 4) & "." & Mid(InscCadastral, 9, 4) & "." & Right(InscCadastral, 3), InscCadastral)
                .Formulas "endereco", IIf(Trim(EnderecoImovel) <> "", EnderecoImovel, EnderecoContribuinte)
            Else
                If Len(InscricaoSP) = 18 Then
                    InscricaoSP = Left(InscCadastral, 1) & "." & Mid(InscCadastral, 2, 4) & "." & Mid(InscCadastral, 6, 3) & "." & Mid(InscCadastral, 9, 2) & "." & Mid(InscCadastral, 11, 4) & "." & Right(InscCadastral, 4)
                    .Formulas "endereco", IIf(Trim(EnderecoImovel) <> "", EnderecoImovel, EnderecoContribuinte)
                Else
                    Dim Objeto As Object
                    .Formulas "endereco", IIf(Trim(EnderecoContribuinte) <> "", EnderecoContribuinte, EnderecoImovel)
                End If
            End If
        End If
        If Sigla = Imposto.NomeTributo(ttr_IPTU) Then
            'DADOS DO SINFAZ
            If Verifica_Geracao_Taxas(CodPagamento) = IPM Then
                If Parcela = 0 Then ValorTributo = ValorTributo - TaxaServico
            End If
        End If
        ValorCompleto = CDbl(Nvl(ValorTributo, 0)) + CDbl(Nvl(ValorMulta, 0)) + CDbl(Nvl(ValorJuros, 0)) + TaxaServico + Correcao - CDbl(Nvl(Desconto, 0))
        If Trim(VencimentoNormal) <> "" Then
            .Formulas "VencimentoNormal", VencimentoNormal
        Else
            .Formulas "VencimentoNormal", Imposto.BuscaDataVencimento(CodTributo, Exercicio)
        End If
        'If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            
            If CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) Then
                .Formulas "INSCRICAO_REDUZIDA", InscCadastral
                .Formulas "INSCRICAO_CADASTRAL", Iptu.Busca_IC_Auxiliar(InscCadastral)
                'BCP
                Dim e As String
                Dim o As String
                Dim rdt As VSRecordset
                If Bdados.AbreTabela("SELECT TIPOLOGRADOURO, LOGRADOURO,TBA_NOME,TIM_NUMERO,tim_cgc_cpf_ocupante, tim_ocupante FROM VIS_IMOVEL WHERE TIM_IC='" & Trim(InscCadastral) & "'", rdt) Then
                     e = IIf(IsNull(rdt("TIPOLOGRADOURO")), "", rdt("TIPOLOGRADOURO"))
                    e = e & " " & IIf(IsNull(rdt("LOGRADOURO")), "", rdt("LOGRADOURO"))
                    e = e & "," & IIf(IsNull(rdt("TIM_NUMERO")), "", rdt("TIM_NUMERO"))
                    .Formulas "ENDERECO_IMOVEL", e
                    .Formulas "LOTEAMENTO", IIf(IsNull(rdt("TBA_NOME")), "", rdt("TBA_NOME"))
                    o = IIf(IsNull(rdt("tim_cgc_cpf_ocupante")), "", rdt("tim_cgc_cpf_ocupante"))
                    o = o & "-" & IIf(IsNull(rdt("tim_ocupante")), "", rdt("tim_ocupante"))
                    .Formulas "LOCATARIO", o
                End If
                If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 Then
                    .Formulas "Edificio", "CODÓ - MA"
                End If
                'FIM BCP
                .Formulas "PERIDO", Exercicio
            End If
        'End If
        If Not Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
            .Formulas "Inscricao", InscricaoSP
            .Formulas "Prefeitura", Temp.PegaParametro(Bdados, "CLIENTE")
            .Formulas "CodigoCedente", Temp.PegaParametro(Bdados, "AGENCIA/CEDENTE")
            .Formulas "Carteira", Temp.PegaParametro(Bdados, "CARTEIRA")
            .Formulas "VT_NOME_DAM", Temp.PegaParametro(Bdados, "NOME DAM")
        End If
        
        .Formulas "AgenciaConvenio", Temp.PegaParametro(Bdados, "AGENCIA/CEDENTE")
        .Formulas "MensagemPadrao", Temp.PegaParametro(Bdados, "MENSAGEM PADRAO")
        .Formulas "PERIDO", IIf(Len(Exercicio) = 4, Exercicio, Left(Exercicio, 4) & Right(Exercicio, 2))
        .Formulas "PREFEITURA", Temp.PegaParametro(Bdados, "CLIENTE")
        .Formulas "Nome", InscricaoSP & "-" & RazaoSocial
        .Formulas "cpf/cnpj", IIf(Len(Trim(Edita.TiraTudo(CgcCpf))) >= 11, CgcCpf, "")
        .Formulas "CPF_NOME", IIf(Len(Trim(Edita.TiraTudo(CgcCpf))) >= 11, CgcCpf, "") & " " & InscricaoSP & "-" & RazaoSocial
        
        If Bdados.AbreTabela("select tci_fantasia from tab_contribuinte where tci_im='" & Replace(InscricaoSP, ".", "") & "'", rc) Then
            fantasia = rc(0)
        Else
            fantasia = ""
        End If
        .Formulas "fantasia", fantasia
        
        'BCP
        Dim cb As CodigoDeBarra
        Set cb = New CodigoDeBarra
        Dim carteira As String
        Dim nossoNumero As Long
        carteira = Format(Temp.PegaParametro(Bdados, "CONVENIO CARTEIRA"), "00")
        'PARA GARANTIR A CADA REIMPRESSAO VALIDAR O NOSSO NUMERO BASA
        If Bdados.AbreTabela("SELECT TOC_NOSSO_NUMERO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO=" & CodPagamento) Then
            If IsNull(Bdados.Tabela("TOC_NOSSO_NUMERO")) Or Bdados.Tabela("TOC_NOSSO_NUMERO") = "" Then
                If Temp.PegaParametro(Bdados, "BANCO ARRECADACAO") = 3 Then 'BASA
                    nossoNumero = Imposto.GeraNumCorrelativo(1, 199)  'BASA NOSSO NUMERO
                Else
                    nossoNumero = CodPagamento
                End If
                Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_NOSSO_NUMERO=" & nossoNumero & " WHERE TOC_COD_OBRIGACAO=" & CodPagamento)
            Else
                nossoNumero = Bdados.Tabela("TOC_NOSSO_NUMERO")
            End If
            
        End If
        
        CodPagamento = Format(CodPagamento, "00000000000")
        .Formulas "NossoNumero", carteira & " / " & nossoNumero
        .Formulas "NoDocumento", CodPagamento
        'FIM BCP
        
        .Formulas "datavencimento", DataVencimento
        '.Formulas "VT_VENCIMENTO_NORMAL", VencimentoNormal
        
        If StatusDivida = etsCreditoDividaAtiva Then
            Dim SiglaDAT As String
            SiglaDAT = Imposto.NomeTributo(ttr_DATIVA)
            .Formulas "CodigoTributo", Imposto.BuscaCodImposto(SiglaDAT) & "-" & SiglaDAT
        Else
            If UCase(AplicacoesVTFuncoes.Municipio) = "PETROLINA" Then
                .Formulas "CodigoTributo", Trim(Imposto.PegaSubConta(CodTributo)) & "-" & Trim(Sigla)
            Else
                .Formulas "CodigoTributo", Trim(CodTributo) & "-" & Trim(Sigla)
            End If
        End If
        'If CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) Or CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN)) Then
            .Formulas "ATIVIDADE", Imposto.BuscaNomeCAE(Ecom.BuscaAtividadeContribuinte(InscMunicipal))
        'End If
        If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
           'Se for um dam normal pego o Valor da TXDAM.
            If Tipo = etdNormal Then
                If StringTaxa = "" Then
                    StringTaxa = "[ TXDAM - " & TXDAM & " ]"
                Else
                    StringTaxa = StringTaxa & ", [TXDAM - R$" & Format(TXDAM, Const_Monetario) & "]"
                End If
            End If
        End If
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            .Formulas "Atividade", AtividadeEconomica
        End If
        If Tipo = etdNormal Then
            .Formulas "ValorTributo", Format((CDbl(Nvl(ValorTributo, 0))), Const_Monetario)
            .Formulas "ValorMulta ", Format(ValorMulta, Const_Monetario)
            .Formulas "ValorJuros ", Format(ValorJuros, Const_Monetario)
            .Formulas "Correcao ", Format(0, Const_Monetario)
        Else
            .Formulas "ValorTributo", " "
            .Formulas "ValorMulta ", " "
            .Formulas "ValorJuros ", " "
            .Formulas "Correcao ", " "
        End If
        
        Bdados.FechaTabela Rs
        If Trim(CodTributo) = Trim(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA))) Then
            If Verifica_Geracao_Taxas(CodPagamento) = CIAP Then
                StringTaxa = StringTaxa & ", [ TXEA - " & TrocaPic(Temp.PegaParametro(Bdados, "TXEA"), ".", ",") & "]" & _
                        ", [ TXP - " & Format(TxP, Const_Monetario) & "]"
                
            End If
        End If
        
        If Sigla <> Imposto.NomeTributo(ttr_IPTU) Then
            .Formulas "Mensagem1", UCase(Trim(Observacao))
        Else
'******************COLETA DE DADOS******************
        Sql = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & InscCadastral & _
            "' and (tdi_tco_cod_componente IN (110,108,113))"
        If Bdados.AbreTabela(Sql, Rs) Then
            Rs.MoveFirst
            Do While Not Rs.EOF
                If TipoIptu = 1 Then
                    If Rs(0) = 108 Then
                        AreaTotal = Nvl("" & Rs(1), 0)
                    ElseIf Rs(0) = 113 Then
                        AreaConstruida = Nvl(Rs(1), 0)
                    End If
                Else 'mudei
                    If Rs(0) = 108 Then
                        AreaTotal = Nvl("" & Rs(1), 0)
                    ElseIf Rs(0) = 113 Then
                        AreaConstruida = Nvl(Rs(1), 0)
                    End If
                End If
                Rs.MoveNext
            Loop
        End If
        Bdados.FechaTabela Rs
        If TipoIptu <> 1 Then
            Sql = "select tvl_valor from tab_valor_terreno where tvl_tlg_cod_logradouro IN (" & _
                " select tim_tlg_cod_logradouro from tab_imovel where tim_ic='" & InscCadastral & "')"
            If Bdados.AbreTabela(Sql, Rs) Then
                If Not Rs.EOF Then
                    ValorMetro = Format(TrocaPic(Rs(0), ".", ","), Const_Monetario)
                End If
            Else
                Sql = "select TTC_valor from TAB_TRECHO where ttc_tlg_cod_logradouro IN ( select tim_tlg_cod_logradouro from tab_imovel where tim_ic='" & InscCadastral & "')"
                If Bdados.AbreTabela(Sql, Rs) Then
                    If Not Rs.EOF Then
                        ValorMetro = Format(TrocaPic(Rs(0), ".", ","), Const_Monetario)
                    End If
                End If
            End If
        End If
                
        Sql = "select TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC," & _
            " tim_valor ,TIM_SITUACAO_LOTE   from TAB_IMOVEL where tim_ic='" & InscCadastral & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            ValorTerreno = Nvl("" & Rs!TIM_VALOR_TERRENO, 0)
            ValorEdificacao = Nvl("" & Rs!TIM_VALOR_EDIFIC, 0)
            BaseCalculo = "" & Rs!tim_valor
        End If
'******************FIM DE COLETA DE DADOS******************
            If TipoIptu = 1 Then
                
                .Formulas "Mensagem1", Trim(Observacao)
                If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
                    Observacao = "Área Total do Imóvel: " & Format(AreaTotal, Const_Monetario) & "m2     -       Área Total Construída: " & Format(AreaConstruida, Const_Monetario) & "m2"
                End If
                .Formulas "Mensagem2", Trim(Observacao)
            Else
                'Valor m2 Terreno: R$" & Format(ValorMetro, Const_Monetario) & " -
                .Formulas "Mensagem1", "Valor m2 Terreno: R$" & Format(ValorMetro, Const_Monetario) & " - Valor Terreno: R$" & Format(ValorTerreno, Const_Monetario) & " - Valor Edificacão: R$" & Format(ValorEdificacao, Const_Monetario) & " - Valor Venal: R$" & Format(BaseCalculo, Const_Monetario)   '& " - TSU: R$" & Format(TaxaServico, Const_Monetario) & " - Valor IPTU: R$" & Format(MaxCotas * (CDbl(Nvl(ValorTributo, 0)) * (100 + Nvl(Desconto, 0)) / 100) - TaxaParcela, Const_Monetario) & _
                IIf(Not IsNumeric(Parcela) And Nvl(Desconto, 0) > 0, Space(32) & "DESCONTO DE " & Desconto & "% EM COTA ÚNICA", "")
                
                .Formulas "Mensagem3", StringTaxa
            End If
        End If
        If Trim(NotaInicial) = "" And Trim(NotaFinal) = "" Then
            If Bdados.AbreTabela("SELECT TDD_VALOR_ITEM FROM TAB_DETALHE_DECLARACAO WHERE TDD_TCD_COD_ITEM IN (1,2) AND " & _
                            " TDD_TDC_NUM_DECLARACAO= (" & _
                            "SELECT TDC_NUM_DECLARACAO FROM TAB_DECLARACAO_CONTRIBUINTE WHERE TDC_TCI_IM ='" & _
                            InscMunicipal & "' AND TDC_PERIODO =" & Exercicio & " AND TDC_TIP_COD_IMPOSTO ='" & CodTributo & _
                            "')  ORDER BY TDD_TCD_COD_ITEM ASC") Then
                Bdados.Tabela.MoveFirst
                NotaInicial = Format(Bdados.Tabela(0) & "", "000")
                Bdados.Tabela.MoveNext
                If Not Bdados.Tabela.EOF Then NotaFinal = Format(Bdados.Tabela(0) & "", "000")
                Bdados.FechaTabela
            Else
                If Bdados.AbreTabela("SELECT tna_numero_nota FROM tab_nota_avulsa WHERE tna_cod_pagamento =" & Nvl(CodPagamento, 0)) Then
                    NotaInicial = Format(Bdados.Tabela(0) & "", "000")
                    NotaFinal = Format(Bdados.Tabela(0) & "", "000")
                End If
            End If
            Bdados.FechaTabela
        End If
        
        If Trim(NotaInicial) <> "" And Trim(NotaFinal) <> "" Then
            If Tipo = etdNormal Then
                .Formulas "Mensagem3", "Nº da Nota Inicial: " & NotaInicial & " - Nº da Nota Final: " & NotaFinal
            Else
                .Formulas "Mensagem3", "Nº da Nota Inicial ............... Nº da Nota Final ..............."
            End If
        ElseIf Sigla = Imposto.NomeTributo(ttr_IPTU) Then
            If Not (AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE") Then
                If TipoIptu = 1 Then
                    .Formulas "Mensagem3", "Valor Terreno: R$" & Format(ValorTerreno, Const_Monetario) & "     -       Valor Edificacao: R$" & Format(ValorEdificacao, Const_Monetario)
                ElseIf TipoIptu <> 1 Then
                    .Formulas "Mensagem2 ", "Área Total do Imóvel: " & Format(AreaTotal, Const_Monetario) & "m2     -       Área Total Construída: " & Format(AreaConstruida, Const_Monetario) & "m2"
                End If
            Else 'ADICONADO EM PETROLINA PARA IDENTIFICACAO DE DAT
                If StatusDivida = etsCreditoDividaAtiva Then
                    .Formulas "Mensagem3", "DÍVIDA ATIVA TRIBUTÁRIA - " & Sigla
                End If
            End If
            If AplicacoesVTFuncoes.Municipio = "COLINAS" Then
                Sql = "Select tim_ic_anterior from tab_imovel where tim_ic ='" & InscricaoSP & "'"
                If Bdados.AbreTabela(Sql, Rs) Then
                    If Len(Trim(Rs!TIM_IC_ANTERIOR)) = 14 Then
                        .Formulas "endereco", Left(Trim(Rs!TIM_IC_ANTERIOR), 2) & "." & Mid(Trim(Rs!TIM_IC_ANTERIOR), 3, 2) & "." & _
                        Mid(Trim(Rs!TIM_IC_ANTERIOR), 5, 3) & "." & Mid(Trim(Rs!TIM_IC_ANTERIOR), 8, 4) & "." & Right(Trim(Rs!TIM_IC_ANTERIOR), 3) & " - " & EnderecoImovel
                    Else
                        .Formulas "endereco", Trim(Rs!TIM_IC_ANTERIOR) & " - " & EnderecoImovel
                    End If
                End If
            End If
        End If
        If Desconto = 0 Then
            Dim r As VSRecordset
            If Bdados.AbreTabela("SELECT TCC_DESCONTO_CONCEDIDO FROM TAB_CONTA_CONTRIBUINTE WHERE TCC_CODIGO_CONTA=" & CLng(CodPagamento), r) Then
                If Not IsNull(r(0)) Then
                    Desconto = CCur(r(0))
                    ValorCompleto = ValorCompleto - Desconto
                End If
            End If
        End If
        If Tipo = etdNormal Then
            .Formulas "DEDUCAO", Format(CDbl(Nvl(Desconto, 0)), Const_Monetario) 'Format(CDbl(CDbl(Nvl(ValorTributo, 0)) + TaxaServico) * (Nvl(Desconto, 0) / 100), Const_Monetario)
        Else
            .Formulas "DEDUCAO", " "
        End If
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
        End If
        CgcPref = Temp.PegaParametro(Bdados, "CGC CLIENTE")
'        .Formulas "CgcPrefeitura", "CNPJ " & Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
       'ESTA DATA EMISSÃO DAM
       .Formulas "EMISSAO", Date 'Imposto.BuscaDataGeracaoDam(CodPagamento)
        If Nvl(MaterialICMS, 0) > 0 Then
            .Formulas "Mensagem3", "Valor de material sujeito ao ICMS: R$" & Format(MaterialICMS, Const_Monetario)
        End If
        If Sigla = Imposto.NomeTributo(ttr_IPTU) Then
            If TipoIptu = 1 And Conta = 1 Then
                .Formulas "Mensagem3 ", " Valor Venal: R$" & Format(BaseCalculo, Const_Monetario) & Space(5) & " TSU: R$" & Format(TaxaServico, Const_Monetario) & " - Valor IPTU: R$" & Format(ValorTributo, Const_Monetario)
            End If
        End If
        InscricaoSP = Edita.TiraTudo(InscricaoSP)
        
        If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
              If Tipo = etdNormal Then
                LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(InscricaoSP, CodTributo, ValorCompleto, Exercicio, Picture, DataVencimento, Nvl(Parcela, 0), CodPagamento)
              ElseIf Tipo = etdBranco Then
                LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(InscricaoSP, CodTributo, 0, Exercicio, Picture, DataVencimento, Nvl(Parcela, 0), CodPagamento)
              End If
        Else
            If Tipo = etdNormal Then
                LinhaDigitavel = CodBarra.CriaLinhaDigitavel(InscricaoSP, CodTributo, ValorCompleto, Exercicio, DataVencimento, Nvl(Parcela, 0), CodPagamento)
            ElseIf Tipo = etdBranco Then
                LinhaDigitavel = CodBarra.CriaLinhaDigitavel(InscricaoSP, CodTributo, 0, Exercicio, DataVencimento, Nvl(Parcela, 0), CodPagamento)
            End If
        End If
        .Formulas "LinhaDigitavel", LinhaDigitavel
        .Formulas "VT_LinhaBarra", CodBarra.LinhaBarraGerada
        .Formulas "VT_USUARIO", AplicacoesVTFuncoes.Usuario '(RECOLOCAR PARA ITAPECURU)
        .Formulas "Correcao", Format(Correcao, Const_Monetario)
        If Parcela <> "0" And Parcela <> "" And Parcela <> "1" Then
            Sql = "SELECT count(*) as Conta,TCO_TOC_COD_OBRIGACAO FROM TAB_COTAS_OBRIGACAO WHERE TCO_TOC_COD_OBRIGACAO =(" & _
                    " SELECT TCO_TOC_COD_OBRIGACAO FROM TAB_COTAS_OBRIGACAO where TCO_COD_OBRIGACAo_PARCELA = " & CodPagamento & ") group by TCO_TOC_COD_OBRIGACAO"
            If Bdados.AbreTabela(Sql, Rs) Then
                Parcela = Parcela & "/" & "" & Rs!Conta
                NumParcelamento = "" & Rs!TCO_TOC_COD_OBRIGACAO
            Else
                Sql = "SELECT count(*) AS Conta,TCP_TPA_COD_PARCELAMENTO FROM TAB_COTAS_PARCELAMENTO WHERE TCP_TPA_COD_PARCELAMENTO =(" & _
                    " SELECT TCP_TPA_COD_PARCELAMENTO FROM TAB_COTAS_PARCELAMENTO where TCP_NUM_COTA = " & CodPagamento & ") group by TCP_TPA_COD_PARCELAMENTO"
                If Bdados.AbreTabela(Sql, Rs) Then
                    Parcela = Parcela & "/" & "" & Rs!Conta
                    NumParcelamento = "" & Rs!TCP_TPA_COD_PARCELAMENTO
                End If
            End If
        End If
        'ANEXOS
'        If (AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE") Or Txea > 0 Then
'            .Formulas "Mensagem2", StringTaxa
'        End If
        If Tipo = etdNormal Then
            .Formulas "TaxaExpediente ", Format(TaxaServico, Const_Monetario)
        Else
            .Formulas "TaxaExpediente ", Format(TrocaPic(Temp.PegaParametro(Bdados, "TXDAM"), ".", ","), Const_Monetario)
        End If
        If Tipo = etdNormal Then
            If Verifica_Geracao_Taxas(CodPagamento) = CIAP Then
                .Formulas "ValorTotal ", Format(ValorCompleto, Const_Monetario)
            Else
                .Formulas "ValorTotal ", Format(ValorCompleto, Const_Monetario)
            End If
        Else
            .Formulas "ValorTotal ", " "
        End If
        
        If StatusDivida = etsCreditoDividaAtiva Then
            .Formulas "Mensagem3", "DÍVIDA ATIVA TRIBUTÁRIA - " & Sigla
        ElseIf Len(Trim(NumParcelamento)) > 0 Then
            .Formulas "Mensagem3", "PARCELAMENTO No. - " & NumParcelamento
        End If
        If Tipo = etdNormal Then
            .Formulas "Parcela", IIf(Parcela = "0", "UNICA", Parcela)
        ElseIf Tipo = etdBranco Then
            .Formulas "Parcela", Month(DataVencimento) - 1
        End If
        .titulo = "Documento de Arrecadação Municipal - DAM"
        'REMOVER DEPOIS
        '.Formulas "Emissao", "16/06/2011"
        '.Formulas "DataVencimento", "16/06/2011"
'        .CopiasDetalhes = 3
        .Arvore = False
        'BCP
        Dim RsDET As VSRecordset
        If Bdados.AbreTabela("SELECT TOC_STATUS_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO = " & CLng(CodPagamento), RsDET) Then
            If RsDET(0) = 8 Then 'CANCELADO
                .Formulas "STATUS", "BOLETO CANCELADO"
            End If
        End If
        Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_REMESSA = null WHERE TOC_REMESSA=0 AND TOC_COD_OBRIGACAO = " & CLng(CodPagamento))
        'FIM
        If Destino = tdiTela Then
            .Visualizar
        Else
            .Imprimir
        End If
        DoEvents
    End With
    Set Rpt = Nothing
    Exit Sub
    
Trata:
    If Err.Number = 20515 Or Err.Number = 3265 Then
        Report.Formulas "Mensagem1 ", ""
        Resume
    ElseIf Err.Number = -2147189582 Then
        Resume Next
        Exit Sub
    End If
    Avisa "O DAM não será impresso. Informe o erro a seguir ao suporte."
    Avisa Err.Number & " - " & Err.Description
    Exit Sub
    Resume
End Sub

Public Sub ImprimeDamBarra(ByRef Rpt As Object, InscricaoSP As String, CodTributo As String, ValorTotal As Double, Exercicio As String, Picture As Object, DataVencimento As String, Parcela As Integer, CodPagamento As String)
    Dim LinhaDigitavel  As String
    Dim CodBarra As New CodigoDeBarra
    InscricaoSP = Edita.TiraTudo(InscricaoSP)
    
    
    If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
        LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(InscricaoSP, CodTributo, ValorTotal, Exercicio, Picture, DataVencimento, Parcela, CodPagamento)
    Else
        LinhaDigitavel = CodBarra.CriaLinhaDigitavel(InscricaoSP, CodTributo, ValorTotal, Exercicio, DataVencimento, Parcela, CodPagamento)
    End If
    Rpt.Formulas "LinhaDigitavel", LinhaDigitavel
    Rpt.Formulas "VT_LinhaBarra", CodBarra.LinhaBarraGerada
'
'    If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
'            .Formulas "VT_LinhaBarra", CodBarra.LinhaBarraGerada
'        End If
End Sub

Public Function Verifica_Geracao_Taxas(Obrigacao As String) As Geracao_Taxa
    Dim Sql As String
    Sql = "Select toc_situacao_Taxa from tab_obrigacao_contribuinte where toc_cod_obrigacao   = '" & Obrigacao & "'"
    If Bdados.AbreTabela(Sql) Then
        If Bdados.Tabela(0) = 1 Then
            Verifica_Geracao_Taxas = IPM
        ElseIf Bdados.Tabela(0) = 2 Then
            Verifica_Geracao_Taxas = CIAP
        Else
'            Util.Avisa "Taxas não atualizadas."
        End If
    End If
End Function






Public Sub ImprimeDamITBI(CodPagamento As String, InscricaoMunicipal As String, CgcAdquirente As String, NomeCedente As String, _
                            EnderecoCedente As String, DataVencimento As String, NomeAdquirente As String, EnderecoAdquirente As String, _
                            InscricaoCadastral As String, EnderecoImovel As String, Periodo As String, ValorItbi As Double, _
                            Multa As Double, Juros As Double, BaseCalculo As Double, AliquotaProprio As Double, AliquotaFinanciado As Double, _
                            TaxaServico As Double, Observacao As String, Especie As String, PicBarra As Object, Optional ValorFinanciado As Double, Optional ValorAvista As Double, Optional ValorAvaliacao As Double, Optional Processo As String, _
                            Optional NomeTributo As String, Optional CodImposto As String)
    On Error GoTo Trata
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim MaxCotas As Byte
    Dim Cobranca As New VSCobranca
    Dim CodBarra As New CodigoDeBarra
    Dim LinhaDigitavel As String
    Dim MSG As String
    Dim a As Byte
    
    Dim CgcPref As String
    
    With Rpt
        If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
            'BCP ANTES If Not .DefinirArquivo(Bdados, App.Path + "\TDAM_ITBI_Barra_TITULO.rpt") Then Exit Sub
            If Not .DefinirArquivo(Bdados, App.Path + "\TDAMBarra_TITULO.rpt") Then Exit Sub
        Else
            'BCP ANTES If Not .DefinirArquivo(Bdados, App.Path + "\TDAM_ITBI_Barra.rpt") Then Exit Sub
            If Not .DefinirArquivo(Bdados, App.Path + "\TDAMBarra_TITULO.rpt") Then Exit Sub
        End If
        MSG = Entrada("Observações", "Mensagem")
        If MSG <> "" Then
          .Formulas "Mensagem1", MSG
        End If
        'BCP
        Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_REMESSA = null WHERE TOC_REMESSA=0 AND TOC_COD_OBRIGACAO = " & CLng(CodPagamento))
        
        Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_OBSERVACAO='" & MSG & "' WHERE TOC_COD_OBRIGACAO=" & CodPagamento)
        
        'FIM BCP
        
        Sql = "SELECT TPI_GERA_TAXA_IMPRESSAO FROM TAB_PARAMETRO_IMPOSTO WHERE TPI_TIP_COD_IMPOSTO = '" & CodImposto & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            If CInt(Nvl("" & Rs!TPI_GERA_TAXA_IMPRESSAO, 1)) = 1 Then TaxaServico = TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXDAM"), 0), ".", ",")
        End If
        If UCase(AplicacoesVTFuncoes.Municipio) = "PETROLINA" Then
            .Formulas "Mensagem2", "[ TxDam ] - " & Temp.PegaParametro(Bdados, "TXTDAM")
        End If
        
        .Formulas "VT_NOME_DAM", Temp.PegaParametro(Bdados, "NOME DAM")
        .Formulas "VALOR_AVALIACAO", Format(CStr(ValorAvaliacao), "###,###,###,##0.00")
        .Formulas "BASE_CALCULO2", Format(CStr(ValorFinanciado), "###,###,###,##0.00")
        .Formulas "ALIQUOTA1", Format(CStr(AliquotaProprio), "###,###,###,##0.00") & " %"
        .Formulas "VALORAVISTA", Format(CDbl(AliquotaProprio) * CDbl(ValorAvista) / 100, "###,###,###,##0.00")
        .Formulas "VALORfinanciado", Format(CDbl(AliquotaFinanciado) * CDbl(ValorFinanciado) / 100, "###,###,###,##0.00")
        .Formulas "Parcela", "UNICA"
        .Formulas "Deducao", "0,00"
        .Formulas "Correcao", "0,00"
        .Formulas "PROCESSO", Processo
        .Formulas "VencimentoNormal", DataVencimento
        .Formulas "PERIDO", Periodo
        .Formulas "ALIQUOTA2", Format(CStr(AliquotaFinanciado), "###,###,###,##0.00") & " %"
        'BCP
        Dim cb As CodigoDeBarra
        Set cb = New CodigoDeBarra
        Dim carteira As String
        carteira = Format(Temp.PegaParametro(Bdados, "CONVENIO CARTEIRA"), "00")
        CodPagamento = Format(CodPagamento, "00000000000")
        .Formulas "NossoNumero", carteira & " / " & CodPagamento & " - " & cb.gerarDigitoConferencia(carteira, CodPagamento) ' Format(CodPagamento, "00000000000000000")
        'FIM BCP
        .Formulas "InscMunicipal", InscricaoMunicipal
        
        .Formulas "noDocumento", CodPagamento
        CgcPref = Temp.PegaParametro(Bdados, "CGC CLIENTE")
        .Formulas "CgcPrefeitura", "CNPJ " & Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
        .Formulas "CPF/CNPJ", CgcAdquirente
        .Formulas "nome", NomeAdquirente
        .Formulas "documento ", CStr(CodPagamento)
        .Formulas "endereco", Edita.TiraPic(EnderecoAdquirente, "'")
        .Formulas "datavencimento", DataVencimento
        .Formulas "NOMECEDENTE", Edita.TiraPic(NomeCedente, "'")
        .Formulas "ENDERECOCEDENTE", Edita.TiraPic(EnderecoCedente, "'")
        .Formulas "localizacao", InscricaoCadastral & " - " & EnderecoImovel
        .Formulas "exercicio", Periodo
        .Formulas "ValorTributo", Format(ValorItbi, Const_Monetario)
        .Formulas "CPF_NOME", CgcAdquirente & " " & NomeAdquirente
        Dim RsDET As VSRecordset
        'BCP
        If Bdados.AbreTabela("SELECT * FROM VIS_TRANSFERENCIA_IMOVEL WHERE TTI_TOC_COD_OBRIGACAO = " & CodPagamento, RsDET) Then
            .Formulas "INSCRICAO_REDUZIDA", RsDET("TTI_TIM_IC")
            .Formulas "ENDERECO_IMOVEL", RsDET("TTL_NOME") & " " & RsDET("TLG_NOME") & ", " & RsDET("TIM_NUMERO")
            .Formulas "LOTEAMENTO", RsDET("TBA_NOME")
            .Formulas "EDIFICIO", RsDET("TMU_NOME")
            '.Formulas "INSCRICAO_REDUZIDA", inscricaoImovel
            '.Formulas "ENDERECO_IMOVEL", EnderecoImovel
            '.Formulas "LOTEAMENTO", bairroImovel
            '.Formulas "EDIFICIO", cidadeImovel
        End If
        If Bdados.AbreTabela("SELECT TOC_STATUS_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO = " & CodPagamento, RsDET) Then
            If RsDET(0) = 8 Then 'CANCELADO
                .Formulas "STATUS", "BOLETO CANCELADO"
            End If
        End If
        'FIM BCP
        If UCase(AplicacoesVTFuncoes.Municipio) = "PETROLINA" Then
            .Formulas "CodigoTributo", Imposto.PegaSubConta(CodImposto)
        Else
            .Formulas "CodigoTributo", CodImposto
        End If
        .Formulas "ValorMulta", Format(Multa, Const_Monetario)
        .Formulas "ValorJuros", Format(Juros, Const_Monetario)
        
        .Formulas "TaxaExpediente", Format(TaxaServico, Const_Monetario)
        
        .Formulas "ValorTotal", Format(ValorItbi + Multa + Juros + TaxaServico, Const_Monetario)
'         .Formulas "CodigoTributo", Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITBI)) & " - ITBI"
        .Formulas "NUM_NOTAS", "Valor Venal do Imóvel(R$): " & _
           Format(BaseCalculo, Const_Monetario) & " - [Aliquotas(%): Prop= " & Format(AliquotaProprio, Const_Monetario) & _
           " ; Financ.=" & Format(AliquotaFinanciado, Const_Monetario) & "]"
        .Formulas "BASECALCULO", Format(ValorAvista, Const_Monetario)
        .Formulas "VT_ESPECIE", Especie
          Dim PosTraco As Byte
        .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
        .Formulas "EMISSAO", Imposto.BuscaDataGeracaoDam(CodPagamento)
        .Formulas "Imposto", IIf(NomeTributo = "", "IMPOSTO SOBRE TRANSMISSÃO DE BENS INTERVIVOS", NomeTributo)
'         .Formulas "LinhaDigitavel", Cobranca.GeraCodBarra(CodPagamento, CodImposto, CDbl(TotalImposto) + CDbl(TaxaServico), PicBarra, Right(Exercicio, 4) & Left(Exercicio, 2), Data_Vencimento, 0, 1)
         If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
          LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(InscricaoMunicipal, Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITBI)), ValorItbi + Multa + Juros + TaxaServico, Right(Periodo, 4) & Left(Periodo, 2), PicBarra, DataVencimento, 0, CStr(CodPagamento))
        Else
          LinhaDigitavel = CodBarra.CriaLinhaDigitavel(InscricaoMunicipal, Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITBI)), ValorItbi + Multa + Juros + TaxaServico, Right(Periodo, 4) & Left(Periodo, 2), DataVencimento, 0, CStr(CodPagamento))
        End If
        .Formulas "LinhaDigitavel", LinhaDigitavel
        .Formulas "VT_LinhaBarra", CodBarra.LinhaBarraGerada
        .Formulas "OBSERVACAO", Edita.TiraPic(Trim(Observacao), "'")
        .Formulas "VT_USUARIO", AplicacoesVTFuncoes.Usuario
     '.Connect = Bdados.BDSistema.Connect
      .titulo = "Documento de Arrecadação Municipal - DAM"
      .Arvore = False
      .Visualizar
    End With
    Exit Sub
Trata:
    If Err.Number = 20515 Or Err.Number = 3265 Then
        Rpt.Formulas "OBSERVACAO", ""
        Resume
    End If
    Avisa "O DAM não será gerado."
    Avisa Err.Description
    Exit Sub
    Resume
End Sub

Public Function LiberaImpressaoDam(Status As Integer) As Boolean
    LiberaImpressaoDam = True
    
    If Status = 3 Or Status = 4 Or Status = 6 Then
        Avisa "DAM não pode ser impresso. Operacão não autorizada para status atual do crédito."
        Screen.MousePointer = 0
        LiberaImpressaoDam = False
    End If
End Function


'***** EM 14/04/2011 ***********
Public Sub ImprimeDamIPTU(Report As VSRelatorio, CodPagamento As String, InscMunicipal As String, RazaoSocial As String, _
CgcCpf As String, EnderecoContribuinte As String, InscCadastral As String, EnderecoImovel As String, CodTributo As String, _
Sigla As String, NomeImposto As String, Exercicio As String, Parcela As String, Conta As Integer, DataVencimento As String, _
BaseCalculo As String, ValorTributo As String, ValorMulta As String, ValorJuros As String, ByVal TaxaServico As Double, _
Desconto As String, AtividadeEconomica As String, Observacao As String, Optional Picture As Object, Optional NotaInicial As String, _
Optional NotaFinal As String, Optional MaterialICMS As String, Optional ValorMetro As Double, Optional TaxaParcela As Double, _
Optional AreaTotal As Double, Optional AreaConstruida As Double, Optional ValorTerreno As Double, Optional ValorEdificacao As Double, Optional Zona As Integer, Optional Destino As TipoDestinoImpressao = tdiTela, Optional Tipo As Tipo_Dam = etdNormal, Optional StringTaxa As String, Optional VencimentoNormal As String, Optional Correcao As Double, Optional StatusDivida As TipoStatusObrigacao, Optional DescontosAtualizacao As Double)
   On Error GoTo Trata
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim MaxCotas As Byte
    Dim CodBarra As New CodigoDeBarra
    Dim a As Byte
    Dim DataDesconto As String
    Dim CgcPref As String
    Dim LinhaDigitavel As String
    Dim InscricaoSP As String
    Dim ObsTributo As String
    Dim TipoIptu As Integer
    Dim Iptu As New VSIptu
    Dim Txea As Double
    Dim Tsd As Double
    Dim ValorCompleto As Double
    Dim TXDAM As Double
    Dim TxP As Double
    Dim NumParcelamento As String
    Dim Obrig As New Obrigacao
    Dim FatorAlvara As Double
    
    'BCP
        Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_REMESSA = null WHERE TOC_REMESSA=0 AND TOC_COD_OBRIGACAO = " & CLng(CodPagamento))
      Dim cb As CodigoDeBarra
        Set cb = New CodigoDeBarra
        Dim carteira As String
        carteira = Format(Temp.PegaParametro(Bdados, "CONVENIO CARTEIRA"), "00")
        CodPagamento = Format(CodPagamento, "00000000000")
        
    'FIM BCP
    
    TipoIptu = Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0)
    InscMunicipal = Trim(InscMunicipal)
    InscCadastral = Trim(InscCadastral)
    DataDesconto = Temp.PegaParametro(Bdados, "DATA DESCONTO")
    CodTributo = Trim(CodTributo)
    If Trim(CodTributo) = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) And CInt(Nvl(Parcela, 0)) = 0 Then
        Txea = Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXEA"), ".", ","), 0)
        TxP = Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXP"), ".", ","), 0)
        Sql = "select tci_fantasia,tae_nome,tci_fator_alvara from tab_contribuinte,tab_atividade_economica  where tci_im = '" & InscMunicipal & "' and tci_tae_cae = tae_cae "
        If Bdados.AbreTabela(Sql, Rs) Then
            AtividadeEconomica = Trim(Rs!tae_nome) & " - " & Trim(Rs!tci_fantasia)
            FatorAlvara = Rs!tci_fator_alvara
            If Temp.PegaParametro(Bdados, "TXP ESTIMADA") = "SIM" Then
                If FatorAlvara > CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXP ESTIMADA LIMITE"), ".", ","), 0)) Then
                    TxP = TxP * CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXP ESTIMADA LIMITE VALOR"), ".", ","), 0))
                End If
            End If
        End If
    End If
    If Trim(CodTributo) = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) And CInt(Nvl(Parcela, 0)) = 0 Then
        Tsd = Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TSD"), ".", ","), 0)
    End If
    'VERIFICA SE O TRIBUTO SERÁ IMPRESSO COM TAXA
    Sql = "SELECT TPI_GERA_TAXA_IMPRESSAO "
    Sql = Sql & "  FROM TAB_PARAMETRO_IMPOSTO "
    Sql = Sql & " WHERE TPI_TIP_COD_IMPOSTO = '" & CodTributo & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
            If CInt(Nvl("" & Rs!TPI_GERA_TAXA_IMPRESSAO, 1)) = 1 Then
                TXDAM = TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXDAM"), 0), ".", ",")
            End If
        Else
            If CInt(Nvl("" & Rs!TPI_GERA_TAXA_IMPRESSAO, 1)) = 1 And CDbl(ValorTributo) > 0 Then
                TXDAM = TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXDAM"), 0), ".", ",")
            Else
'                TaxaServico = 0  CUIDADO COM ESSA TAXA
            End If
        End If
    End If
    TXDAM = Format(TXDAM, Const_Monetario)
    Bdados.FechaTabela Rs
    
    If CodTributo <> Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) Then
        Desconto = ValorTributo * CDbl(Nvl(Desconto, 0)) / 100
    Else
        If Verifica_Geracao_Taxas(CodPagamento) = IPM Then
            Desconto = Format((ValorTributo - TXDAM) * CDbl(Nvl(Desconto, 0)) / 100, Const_Monetario)
        Else
            Desconto = Format((ValorTributo) * CDbl(Nvl(Desconto, 0)) / 100, Const_Monetario)
        End If
    End If
    Desconto = Desconto + DescontosAtualizacao
    If TaxaServico = 0 Then
        TaxaServico = TXDAM + Tsd
    Else
        If Verifica_Geracao_Taxas(CodPagamento) = IPM Then
            TaxaServico = TaxaServico + TXDAM
        Else
            TaxaServico = TaxaServico + Tsd
        End If
    End If
    TaxaServico = TaxaServico + Txea + TxP
    If Trim(CodTributo) = Trim(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))) Then
        StringTaxa = " [TSD - R$" & Format(Tsd, Const_Monetario) & "]"
    End If
    Dim relatorio As String
    relatorio = "\TDAMBARRA_TITULO_IPTU.rpt"
    
    With Report
        ' gerar o arquivo bmp da barra antes de chamar o rpt pois somente assim irá sair com o bmp correto
        If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
            If Not .DefinirArquivo(Bdados, App.Path & relatorio) Then Exit Sub '"\TdamBarra_TITULO.rpt"
        Else
            If Not .DefinirArquivo(Bdados, App.Path & relatorio) Then Exit Sub '"\TdamBarra_BB.rpt"
        End If
        InscricaoSP = IIf(Trim(InscCadastral) <> "", Trim(InscCadastral), Trim(InscMunicipal))
        InscricaoSP = TiraPic(TiraPic(TiraPic(InscricaoSP, "."), "-"), "/")
        If Len(InscricaoSP) = 10 Then
            InscricaoSP = Left(InscMunicipal, 2) & "." & Mid(InscMunicipal, 3, 3) & "." & Mid(InscMunicipal, 6, 3) & Right(InscMunicipal, 3)
            .Formulas "endereco", IIf(Trim(EnderecoContribuinte) <> "", EnderecoContribuinte, EnderecoImovel)
        Else
            If Len(InscricaoSP) = 9 Or Len(InscricaoSP) = 15 Then
                InscricaoSP = IIf(Len(InscCadastral) = 15, Left(InscCadastral, 2) & "." & Mid(InscCadastral, 3, 2) & "." & Mid(InscCadastral, 5, 4) & "." & Mid(InscCadastral, 9, 4) & "." & Right(InscCadastral, 3), InscCadastral)
                .Formulas "endereco", IIf(Trim(EnderecoImovel) <> "", EnderecoImovel, EnderecoContribuinte)
            Else
                If Len(InscricaoSP) = 18 Then
                    InscricaoSP = Left(InscCadastral, 1) & "." & Mid(InscCadastral, 2, 4) & "." & Mid(InscCadastral, 6, 3) & "." & Mid(InscCadastral, 9, 2) & "." & Mid(InscCadastral, 11, 4) & "." & Right(InscCadastral, 4)
                    .Formulas "endereco", IIf(Trim(EnderecoImovel) <> "", EnderecoImovel, EnderecoContribuinte)
                Else
                    Dim Objeto As Object
                    .Formulas "endereco", IIf(Trim(EnderecoContribuinte) <> "", EnderecoContribuinte, EnderecoImovel)
                End If
            End If
        End If
        If Sigla = Imposto.NomeTributo(ttr_IPTU) Then
            'DADOS DO SINFAZ
            If Verifica_Geracao_Taxas(CodPagamento) = IPM Then
                If Parcela = 0 Then ValorTributo = ValorTributo - TaxaServico
            End If
        End If
        ValorCompleto = CDbl(Nvl(ValorTributo, 0)) + CDbl(Nvl(ValorMulta, 0)) + CDbl(Nvl(ValorJuros, 0)) + TaxaServico + Correcao - CDbl(Nvl(Desconto, 0))
        If Trim(VencimentoNormal) <> "" Then
            .Formulas "VencimentoNormal", VencimentoNormal
        Else
            .Formulas "VencimentoNormal", Imposto.BuscaDataVencimento(CodTributo, Exercicio)
        End If
        'If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            
            If CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) Then
                .Formulas "INSCRICAO_REDUZIDA", InscCadastral
                .Formulas "INSCRICAO_CADASTRAL", Iptu.Busca_IC_Auxiliar(InscCadastral)
                .Formulas "ENDERECO_IMOVEL", Iptu.Busca_Endereco_Imovel(InscCadastral)
                .Formulas "LOTEAMENTO", Iptu.Busca_Loteamento(InscCadastral)
                .Formulas "Edificio", Iptu.Busca_Edificio(InscCadastral)
                .Formulas "PERIDO", Exercicio
            End If
        'End If
        If Not Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
            .Formulas "Inscricao", InscricaoSP
            .Formulas "Prefeitura", Temp.PegaParametro(Bdados, "CLIENTE")
            .Formulas "CodigoCedente", Temp.PegaParametro(Bdados, "AGENCIA/CEDENTE")
            .Formulas "Carteira", Temp.PegaParametro(Bdados, "CARTEIRA")
            .Formulas "VT_NOME_DAM", Temp.PegaParametro(Bdados, "NOME DAM")
        End If
        .Formulas "MensagemPadrao", Temp.PegaParametro(Bdados, "MENSAGEM PADRAO")
        .Formulas "PERIDO", IIf(Len(Exercicio) = 4, Exercicio, Left(Exercicio, 4) & Right(Exercicio, 2))
        .Formulas "PREFEITURA", Temp.PegaParametro(Bdados, "CLIENTE")
        .Formulas "Nome", InscricaoSP & "-" & RazaoSocial
        .Formulas "cpf/cnpj", IIf(Len(Trim(Edita.TiraTudo(CgcCpf))) >= 11, CgcCpf, "")
        .Formulas "CPF_NOME", IIf(Len(Trim(Edita.TiraTudo(CgcCpf))) >= 11, CgcCpf, "") & " " & InscricaoSP & "-" & RazaoSocial
        'BCP
        .Formulas "NossoNumero", carteira & " / " & CodPagamento & " - " & cb.gerarDigitoConferencia(carteira, CodPagamento) ' Format(CodPagamento, "00000000000000000")
        'FIM BCP
        .Formulas "NoDocumento", CodPagamento
        
        'CAMPOS PARA IPTU
        Dim rsDadosImovel As VSRecordset
        Dim sqlDadosImovel As String, Ocupacao As String, Destinacao As String, tipoEdificacao As String
        Dim I As Integer
        sqlDadosImovel = "select tdi_tgc_cod_grupo, tco_descricao_componente from vis_detalhe_imovel where tdi_tgc_cod_grupo in ('26','39','77') and tdi_tim_ic='" & InscCadastral & "'"
        
        If Bdados.AbreTabela(sqlDadosImovel, rsDadosImovel) Then
            I = rsDadosImovel("tdi_tgc_cod_grupo")
            Do While Not rsDadosImovel.EOF
                Select Case I
                Case 26
                    Ocupacao = rsDadosImovel("tco_descricao_componente")
                Case 39
                    tipoEdificacao = rsDadosImovel("tco_descricao_componente")
                Case 77
                    Destinacao = rsDadosImovel("tco_descricao_componente")
                End Select
                rsDadosImovel.MoveNext
            Loop
        End If
        .Formulas "OCUPACAO_IMOVEL", Ocupacao
        .Formulas "UTILIZACAO_IMOVEL", Destinacao
        .Formulas "TIPO_EDIFICACAO", tipoEdificacao
        .Formulas "ENDERECO_IMOVEL", EnderecoImovel
        'FIM DADOS DO IMOVEL IPTU
        .Formulas "datavencimento", DataVencimento
        '.Formulas "VT_VENCIMENTO_NORMAL", VencimentoNormal
        
        If StatusDivida = etsCreditoDividaAtiva Then
            Dim SiglaDAT As String
            SiglaDAT = Imposto.NomeTributo(ttr_DATIVA)
            .Formulas "CodigoTributo", Imposto.BuscaCodImposto(SiglaDAT) & "-" & SiglaDAT
        Else
            If UCase(AplicacoesVTFuncoes.Municipio) = "PETROLINA" Then
                .Formulas "CodigoTributo", Trim(Imposto.PegaSubConta(CodTributo)) & "-" & Trim(Sigla)
            Else
                .Formulas "CodigoTributo", Trim(CodTributo) & "-" & Trim(Sigla)
            End If
        End If
        'If CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) Or CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN)) Then
            .Formulas "ATIVIDADE", Imposto.BuscaNomeCAE(Ecom.BuscaAtividadeContribuinte(InscMunicipal))
        'End If
        If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
           'Se for um dam normal pego o Valor da TXDAM.
            If Tipo = etdNormal Then
                If StringTaxa = "" Then
                    StringTaxa = "[ TXDAM - " & TXDAM & " ]"
                Else
                    StringTaxa = StringTaxa & ", [TXDAM - R$" & Format(TXDAM, Const_Monetario) & "]"
                End If
            End If
        End If
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            .Formulas "Atividade", AtividadeEconomica
        End If
        If Tipo = etdNormal Then
            .Formulas "ValorTributo", Format((CDbl(Nvl(ValorTributo, 0))), Const_Monetario)
            .Formulas "ValorMulta ", Format(ValorMulta, Const_Monetario)
            .Formulas "ValorJuros ", Format(ValorJuros, Const_Monetario)
        Else
            .Formulas "ValorTributo", " "
            .Formulas "ValorMulta ", " "
            .Formulas "ValorJuros ", " "
        End If
        
        Bdados.FechaTabela Rs
        If Trim(CodTributo) = Trim(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA))) Then
            If Verifica_Geracao_Taxas(CodPagamento) = CIAP Then
                StringTaxa = StringTaxa & ", [ TXEA - " & TrocaPic(Temp.PegaParametro(Bdados, "TXEA"), ".", ",") & "]" & _
                        ", [ TXP - " & Format(TxP, Const_Monetario) & "]"
                
            End If
        End If
        
        If Sigla <> Imposto.NomeTributo(ttr_IPTU) Then
            .Formulas "Mensagem1", UCase(Trim(Observacao))
        Else
'******************COLETA DE DADOS******************
        Sql = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & InscCadastral & _
            "' and (tdi_tco_cod_componente IN (110,108,113))"
        If Bdados.AbreTabela(Sql, Rs) Then
            Rs.MoveFirst
            Do While Not Rs.EOF
                If TipoIptu = 1 Then
                    If Rs(0) = 108 Then
                        AreaTotal = Nvl("" & Rs(1), 0)
                    ElseIf Rs(0) = 113 Then
                        AreaConstruida = Nvl(Rs(1), 0)
                    End If
                Else 'mudei
                    If Rs(0) = 108 Then
                        AreaTotal = Nvl("" & Rs(1), 0)
                    ElseIf Rs(0) = 113 Then
                        AreaConstruida = Nvl(Rs(1), 0)
                    End If
                End If
                Rs.MoveNext
            Loop
        End If
        Bdados.FechaTabela Rs
        If TipoIptu <> 1 Then
            Sql = "select tvl_valor from tab_valor_terreno where tvl_tlg_cod_logradouro IN (" & _
                " select tim_tlg_cod_logradouro from tab_imovel where tim_ic='" & InscCadastral & "')"
            If Bdados.AbreTabela(Sql, Rs) Then
                If Not Rs.EOF Then
                    ValorMetro = Format(TrocaPic(Rs(0), ".", ","), Const_Monetario)
                End If
            Else
                Sql = "select TTC_valor from TAB_TRECHO where ttc_tlg_cod_logradouro IN ( select tim_tlg_cod_logradouro from tab_imovel where tim_ic='" & InscCadastral & "')"
                If Bdados.AbreTabela(Sql, Rs) Then
                    If Not Rs.EOF Then
                        ValorMetro = Format(TrocaPic(Rs(0), ".", ","), Const_Monetario)
                    End If
                End If
            End If
        End If
                
        Sql = "select TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC," & _
            " tim_valor ,TIM_SITUACAO_LOTE   from TAB_IMOVEL where tim_ic='" & InscCadastral & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            ValorTerreno = Nvl("" & Rs!TIM_VALOR_TERRENO, 0)
            ValorEdificacao = Nvl("" & Rs!TIM_VALOR_EDIFIC, 0)
            BaseCalculo = "" & Rs!tim_valor
        End If
'******************FIM DE COLETA DE DADOS******************
            If TipoIptu = 1 Then
                
                .Formulas "Mensagem1", Trim(Observacao)
                If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
                    Observacao = "Área Total do Imóvel: " & Format(AreaTotal, Const_Monetario) & "m2     -       Área Total Construída: " & Format(AreaConstruida, Const_Monetario) & "m2"
                End If
                .Formulas "Mensagem2", Trim(Observacao)
            Else
                'Valor m2 Terreno: R$" & Format(ValorMetro, Const_Monetario) & " -
                .Formulas "Mensagem1", "Valor m2 Terreno: R$" & Format(ValorMetro, Const_Monetario) & " - Valor Terreno: R$" & Format(ValorTerreno, Const_Monetario) & " - Valor Edificacão: R$" & Format(ValorEdificacao, Const_Monetario) & " - Valor Venal: R$" & Format(BaseCalculo, Const_Monetario)   '& " - TSU: R$" & Format(TaxaServico, Const_Monetario) & " - Valor IPTU: R$" & Format(MaxCotas * (CDbl(Nvl(ValorTributo, 0)) * (100 + Nvl(Desconto, 0)) / 100) - TaxaParcela, Const_Monetario) & _
                IIf(Not IsNumeric(Parcela) And Nvl(Desconto, 0) > 0, Space(32) & "DESCONTO DE " & Desconto & "% EM COTA ÚNICA", "")
                
                .Formulas "Mensagem3", StringTaxa
            End If
        End If
        If Trim(NotaInicial) = "" And Trim(NotaFinal) = "" Then
            If Bdados.AbreTabela("SELECT TDD_VALOR_ITEM FROM TAB_DETALHE_DECLARACAO WHERE TDD_TCD_COD_ITEM IN (1,2) AND " & _
                            " TDD_TDC_NUM_DECLARACAO= (" & _
                            "SELECT TDC_NUM_DECLARACAO FROM TAB_DECLARACAO_CONTRIBUINTE WHERE TDC_TCI_IM ='" & _
                            InscMunicipal & "' AND TDC_PERIODO =" & Exercicio & " AND TDC_TIP_COD_IMPOSTO ='" & CodTributo & _
                            "')  ORDER BY TDD_TCD_COD_ITEM ASC") Then
                Bdados.Tabela.MoveFirst
                NotaInicial = Format(Bdados.Tabela(0) & "", "000")
                Bdados.Tabela.MoveNext
                If Not Bdados.Tabela.EOF Then NotaFinal = Format(Bdados.Tabela(0) & "", "000")
                Bdados.FechaTabela
            Else
                If Bdados.AbreTabela("SELECT tna_numero_nota FROM tab_nota_avulsa WHERE tna_cod_pagamento =" & Nvl(CodPagamento, 0)) Then
                    NotaInicial = Format(Bdados.Tabela(0) & "", "000")
                    NotaFinal = Format(Bdados.Tabela(0) & "", "000")
                End If
            End If
            Bdados.FechaTabela
        End If
        
        If Trim(NotaInicial) <> "" And Trim(NotaFinal) <> "" Then
            If Tipo = etdNormal Then
                .Formulas "Mensagem3", "Nº da Nota Inicial: " & NotaInicial & " - Nº da Nota Final: " & NotaFinal
            Else
                .Formulas "Mensagem3", "Nº da Nota Inicial ............... Nº da Nota Final ..............."
            End If
        ElseIf Sigla = Imposto.NomeTributo(ttr_IPTU) Then
            If Not (AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE") Then
                If TipoIptu = 1 Then
                    .Formulas "Mensagem3", "Valor Terreno: R$" & Format(ValorTerreno, Const_Monetario) & "     -       Valor Edificacao: R$" & Format(ValorEdificacao, Const_Monetario)
                ElseIf TipoIptu <> 1 Then
                    .Formulas "Mensagem2 ", "Área Total do Imóvel: " & Format(AreaTotal, Const_Monetario) & "m2     -       Área Total Construída: " & Format(AreaConstruida, Const_Monetario) & "m2"
                End If
            Else 'ADICONADO EM PETROLINA PARA IDENTIFICACAO DE DAT
                If StatusDivida = etsCreditoDividaAtiva Then
                    .Formulas "Mensagem3", "DÍVIDA ATIVA TRIBUTÁRIA - " & Sigla
                End If
            End If
            If AplicacoesVTFuncoes.Municipio = "COLINAS" Then
                Sql = "Select tim_ic_anterior from tab_imovel where tim_ic ='" & InscricaoSP & "'"
                If Bdados.AbreTabela(Sql, Rs) Then
                    If Len(Trim(Rs!TIM_IC_ANTERIOR)) = 14 Then
                        .Formulas "endereco", Left(Trim(Rs!TIM_IC_ANTERIOR), 2) & "." & Mid(Trim(Rs!TIM_IC_ANTERIOR), 3, 2) & "." & _
                        Mid(Trim(Rs!TIM_IC_ANTERIOR), 5, 3) & "." & Mid(Trim(Rs!TIM_IC_ANTERIOR), 8, 4) & "." & Right(Trim(Rs!TIM_IC_ANTERIOR), 3) & " - " & EnderecoImovel
                    Else
                        .Formulas "endereco", Trim(Rs!TIM_IC_ANTERIOR) & " - " & EnderecoImovel
                    End If
                End If
            End If
        End If
        If Tipo = etdNormal Then
            .Formulas "DEDUCAO", Format(CDbl(Nvl(Desconto, 0)), Const_Monetario) 'Format(CDbl(CDbl(Nvl(ValorTributo, 0)) + TaxaServico) * (Nvl(Desconto, 0) / 100), Const_Monetario)
        Else
            .Formulas "DEDUCAO", " "
        End If
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
        End If
        CgcPref = Temp.PegaParametro(Bdados, "CGC CLIENTE")
'        .Formulas "CgcPrefeitura", "CNPJ " & Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
       .Formulas "EMISSAO", Date 'Imposto.BuscaDataGeracaoDam(CodPagamento)"
        If Nvl(MaterialICMS, 0) > 0 Then
            .Formulas "Mensagem3", "Valor de material sujeito ao ICMS: R$" & Format(MaterialICMS, Const_Monetario)
        End If
        If Sigla = Imposto.NomeTributo(ttr_IPTU) Then
            If TipoIptu = 1 And Conta = 1 Then
                .Formulas "Mensagem3 ", " Valor Venal: R$" & Format(BaseCalculo, Const_Monetario) & Space(5) & " TSU: R$" & Format(TaxaServico, Const_Monetario) & " - Valor IPTU: R$" & Format(ValorTributo, Const_Monetario)
            End If
        End If
        InscricaoSP = Edita.TiraTudo(InscricaoSP)
        
        If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
              If Tipo = etdNormal Then
                LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(InscricaoSP, CodTributo, ValorCompleto, Exercicio, Picture, DataVencimento, Nvl(Parcela, 0), CodPagamento)
              ElseIf Tipo = etdBranco Then
                LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(InscricaoSP, CodTributo, 0, Exercicio, Picture, DataVencimento, Nvl(Parcela, 0), CodPagamento)
              End If
        Else
            If Tipo = etdNormal Then
                LinhaDigitavel = CodBarra.CriaLinhaDigitavel(InscricaoSP, CodTributo, ValorCompleto, Exercicio, DataVencimento, Nvl(Parcela, 0), CodPagamento)
            ElseIf Tipo = etdBranco Then
                LinhaDigitavel = CodBarra.CriaLinhaDigitavel(InscricaoSP, CodTributo, 0, Exercicio, DataVencimento, Nvl(Parcela, 0), CodPagamento)
            End If
        End If
        .Formulas "LinhaDigitavel", LinhaDigitavel
        .Formulas "VT_LinhaBarra", CodBarra.LinhaBarraGerada
        .Formulas "VT_USUARIO", AplicacoesVTFuncoes.Usuario '(RECOLOCAR PARA ITAPECURU)
        .Formulas "Correcao", Format(Correcao, Const_Monetario)
        If Parcela <> "0" And Parcela <> "" And Parcela <> "1" Then
            Sql = "SELECT count(*) as Conta,TCO_TOC_COD_OBRIGACAO FROM TAB_COTAS_OBRIGACAO WHERE TCO_TOC_COD_OBRIGACAO =(" & _
                    " SELECT TCO_TOC_COD_OBRIGACAO FROM TAB_COTAS_OBRIGACAO where TCO_COD_OBRIGACAo_PARCELA = " & CodPagamento & ") group by TCO_TOC_COD_OBRIGACAO"
            If Bdados.AbreTabela(Sql, Rs) Then
                Parcela = Parcela & "/" & "" & Rs!Conta
                NumParcelamento = "" & Rs!TCO_TOC_COD_OBRIGACAO
            Else
                Sql = "SELECT count(*) AS Conta,TCP_TPA_COD_PARCELAMENTO FROM TAB_COTAS_PARCELAMENTO WHERE TCP_TPA_COD_PARCELAMENTO =(" & _
                    " SELECT TCP_TPA_COD_PARCELAMENTO FROM TAB_COTAS_PARCELAMENTO where TCP_NUM_COTA = " & CodPagamento & ") group by TCP_TPA_COD_PARCELAMENTO"
                If Bdados.AbreTabela(Sql, Rs) Then
                    Parcela = Parcela & "/" & "" & Rs!Conta
                    NumParcelamento = "" & Rs!TCP_TPA_COD_PARCELAMENTO
                End If
            End If
        End If
        'ANEXOS
'        If (AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE") Or Txea > 0 Then
'            .Formulas "Mensagem2", StringTaxa
'        End If
        If Tipo = etdNormal Then
            .Formulas "TaxaExpediente ", Format(TaxaServico, Const_Monetario)
        Else
            .Formulas "TaxaExpediente ", Format(TrocaPic(Temp.PegaParametro(Bdados, "TXDAM"), ".", ","), Const_Monetario)
        End If
        If Tipo = etdNormal Then
            If Verifica_Geracao_Taxas(CodPagamento) = CIAP Then
                .Formulas "ValorTotal ", Format(ValorCompleto, Const_Monetario)
            Else
                .Formulas "ValorTotal ", Format(ValorCompleto, Const_Monetario)
            End If
        Else
            .Formulas "ValorTotal ", " "
        End If
        
        If StatusDivida = etsCreditoDividaAtiva Then
            .Formulas "Mensagem3", "DÍVIDA ATIVA TRIBUTÁRIA - " & Sigla
        ElseIf Len(Trim(NumParcelamento)) > 0 Then
            .Formulas "Mensagem3", "PARCELAMENTO No. - " & NumParcelamento
        End If
        If Tipo = etdNormal Then
            .Formulas "Parcela", IIf(Parcela = "0", "UNICA", Parcela)
        ElseIf Tipo = etdBranco Then
            .Formulas "Parcela", Month(DataVencimento) - 1
        End If
        
        '.Formulas "VT_CodigoBarra", Picture
        .titulo = "Documento de Arrecadação Municipal - DAM"
'        .CopiasDetalhes = 3
        .Arvore = False
        .Formulas "calculoTerreno", "Desconto de 30% no pagamento à vista até o VENCIMENTO"
        .Formulas "formaParcelamento", "Parcelamento disponível com base na escala de valores abaixo:"
        .Formulas "fp1", "Até R$ 100 (2X) - de R$ 101 até R$ 200 (3X) -  de R$ 201 até R$ 500 (4X)"
        .Formulas "fp2", "de R$ 501 até R$ 1.000 (5X)  - acima de R$ 1.000 negociação com o secretário de tributos"
        
        If Destino = tdiTela Then
            .Visualizar
        Else
           .Imprimir
        End If
        DoEvents
    End With
    Set Rpt = Nothing
    Exit Sub
    
Trata:
    If Err.Number = 20515 Or Err.Number = 3265 Then
        Report.Formulas "Mensagem1 ", ""
        Resume
    ElseIf Err.Number = -2147189582 Then
        Resume Next
        Exit Sub
    End If
    Avisa "O DAM não será impresso. Informe o erro a seguir ao suporte."
    Avisa Err.Number & " - " & Err.Description
    Exit Sub
    Resume
End Sub
 

