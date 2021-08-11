VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Obrigacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private JurosObrigacao As Double
Private MultaObrigacao As Double
Private Calculo As New CalculoObrigacao
Public obContribuinte As String
Public obInscMunicipal As String
Public obInscCadastral As String
Public obPeriodo As String
Public obCodImposto As String
Public obStatusObrigacao As TipoStatusObrigacao
Public obDataVencimento As String
Public obTipoInscricao As Byte
Public obCodigoObrigacao As String
Public obDocumentoOrigem As String
Public obValorObrigacao As Double
Public obValorMulta As Double
Public obValorJuros As Double
Public obValorDesconto As Double
Public obNumParcelamento As Double
Public obParcela As Integer
Public NumeroIncidenciaPeriodo As Integer
Private GravandoTaxasAnexadas As Boolean


Private TaxaInclusa() As New TaxasInclusas
Public GeraExtratoUnificado As Boolean
Public Enum TipoStatusObrigacao
    etsCreditoNaoLancado = 1
    etsCreditoOriginalAberto = 2
    etsCreditoPago = 3
    etsCreditoParcelado = 4
    etsCreditoDividaAtiva = 5
    etsCreditoIsento = 6
    etsCreditoCancelado = 8
    etsSemMovimento = 7
    etsCreditoParceladoOficio = 9
    etsAutuado = 10
    stsParcelamentoCancelado = 11
    etsCreditoPagoEmAuto = 12
End Enum
Public Enum Tipo_Log_Obrigacao
    Gravacao = 1
    Alteracao = 2
    Exclusao = 3
End Enum
Public Enum TipoTransacaoOb
    ettAvista = 1
    ettParcelada = 2
    ettDividaAtiva = 3
End Enum

Public Enum TipoStatusListagem
    etlTodos = 1
    etlNaoPagos = 2
    etlPagos = 3
    etlParcelaveis = 4
    etlNaoPagosNaoVencidos = 5
    etlNaoPagosVencidos = 6
    etlNaoPagosLancados = 7
End Enum


Public Enum etcTipoComboTributo
    etcTributario = 1
    etcNaoTributario = 2
    etcTodos = 3
End Enum

Public Enum TipoSubstituicaoObrigacao
    etsNaoSubstitui = 1
    etsSubstitui = 2
    etsCriaNova = 3
End Enum
Enum TipoInscricaoObrigacao
    etiImovel = 1
    etiContribuinte = 2
End Enum

Public Function BuscaSujeitoPassivoObrigacao(Inscricao As String, Optional ByRef RazaoSocial As Object, Optional ByRef Endereco As Object, Optional ByRef Documento As String, Optional tipoInscricao As TipoInscricaoObrigacao) As String
    On Error Resume Next
    Dim Sql As String
    Dim Rs As VSRecordset
    RazaoSocial = ""
    If Not Endereco Is Nothing Then Endereco = ""
    Documento = ""
    If Len(Trim(Inscricao)) = 10 And IsNumeric(Inscricao) Then
        Inscricao = Imposto.FormataInscricao(Trim(Inscricao), InscContrib)
    End If
    If Inscricao = "" Then Exit Function
    If tipoInscricao <> etiImovel Then
        Sql = "SELECT TCI_IM AS VIN_INSCRICAO ,TCI_NOME AS VIN_RAZAO, ltrim(rtrim(tci_logradouro)) as Logr , ltrim(rtrim(tci_nome_logradouro)) as NomeLogr " & _
            ", ltrim(rtrim(tci_numero)) as numero, ltrim(rtrim(tci_bairro)) as bairro,  ltrim(rtrim(tci_cidade)) as cidade, ltrim(rtrim(tci_UF)) AS uf " & _
            ", tci_cgc_cpf as VIN_DOCUMENTO, tci_optante_simples  from TAB_CONTRIBUINTE  WHERE TCI_IM = '" & Inscricao & "'"
    Else
       ' Sql = "SELECT     tim_ic AS VIN_INSCRICAO, tim_tci_im AS VIN_DOCUMENTO, tci_nome AS VIN_RAZAO," & _
        '    "ltrim(rtrim(TTL_NOME)) as Logr, ltrim(rtrim(tlg_nome)) as NomeLogr, ltrim(rtrim(TIM_NUMERO)) as numero, ltrim(rtrim(TBA_NOME)) AS bairro, " & _
         '   " ltrim(rtrim(TLO_DESCRICAO)) as Loteamento,ltrim(rtrim(TIM_LOTE)) as lote,ltrim(rtrim(TIM_QUADRA)) as quadra,TIM_IC_AUXILIAR as icaux,tim_bloco as Bloco,tim_apto as Apto,tim_sala_loja as sala,ted_descricao as Edificio FROM " & _
          '  " VIS_IMOVEL WHERE TIM_IC ='" & Inscricao & "'"
        'BCP
        Sql = "SELECT     tim_ic AS VIN_INSCRICAO, tim_tci_im AS VIN_DOCUMENTO, tci_nome AS VIN_RAZAO," & _
            "ltrim(rtrim(TIPOLOGRADOURO)) as Logr, ltrim(rtrim(LOGRADOURO)) as NomeLogr, ltrim(rtrim(TIM_NUMERO)) as numero, ltrim(rtrim(TBA_NOME)) AS bairro, " & _
            " ltrim(rtrim(TLO_DESCRICAO)) as Loteamento,ltrim(rtrim(TIM_LOTE)) as lote,ltrim(rtrim(TIM_QUADRA)) as quadra,TIM_IC_AUXILIAR as icaux,tim_bloco as Bloco,tim_apto as Apto,tim_sala_loja as sala,ted_descricao as Edificio FROM " & _
            " VIS_IMOVEL WHERE TIM_IC ='" & Inscricao & "'"
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        If AplicacoesVTFuncoes.Municipio = "ITAPECURU MIRIM" Then
            If Nvl("" & Rs!TCI_OPTANTE_SIMPLES, 0) = 2 Then
                Avisa "EMPRESA OPTANTE DO SIMPLES!"
            End If
        End If
        RazaoSocial = Rs.Fields("VIN_RAZAO")
        If tipoInscricao <> etiImovel Then
            If Not Endereco Is Nothing Then Endereco = "" & Trim(Rs!Logr) & " " & Trim(Rs!NomeLogr) & " " & Trim(Rs!Numero) & " " & Trim(Rs!Bairro) & " " & Trim(Rs!Cidade) & " " & Trim(Rs!UF)
        Else
            If Not Endereco Is Nothing Then
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                    Endereco = "" & Trim(Rs.Fields("icaux")) & " - " & Rs!Logr & " - " & Rs!NomeLogr & " - " & Trim(Rs!Numero) & " - " & Trim(Rs!Loteamento) & IIf(Trim(Rs!Lote) <> "", " Lote - " & Trim(Rs!Lote), "") & IIf(Trim(Rs!Quadra) <> "", " Quadra - " & Trim(Rs!Quadra), "") & " " & Trim(Rs!Bairro)
                Else
                    Endereco = Rs!Logr & " - " & Rs!NomeLogr & " - " & Trim(Rs!Numero) & " - " & Trim(Rs!Loteamento) & IIf(Trim(Rs!Lote) <> "", " Lote - " & Trim(Rs!Lote), "") & IIf(Trim(Rs!Quadra) <> "", " Quadra - " & Trim(Rs!Quadra), "") & " " & Trim(Rs!Bairro)
                End If
            End If
            If Not Endereco Is Nothing Then Endereco = Endereco & " " & IIf(Trim(Rs!Edificio) <> " ", " Ed. " & Trim(Rs!Edificio), "") & " " & IIf(Trim(Rs!Bloco) <> " ", " Bl. " & Trim(Rs!Bloco), "") & " " & IIf(Trim(Rs!Apto) <> " ", " Ap. " & Trim(Rs!Apto), "") & " " & IIf(Trim(Rs!sala) <> " ", " Sl " & Trim(Rs!sala), "")
        End If
        Documento = "" & Rs!VIN_DOCUMENTO
        BuscaSujeitoPassivoObrigacao = Trim(Rs!VIN_INSCRICAO & "")
    End If
    
End Function


Public Function CalculaObrigacao(ByVal Inscricao As String, ByVal tipoInscricao As TipoInsc, ByVal CodTributo As String, _
                ByVal SiglaTributo As String, ByVal Periodo As String, ByRef DataVence As String, Optional ByRef TipoLancamento As TipoLanc, Optional MultiplicadorTaxa As Double, Optional DocOrigem As String, Optional ValorTaxaInclusaGerada As Double) As Double
        Dim Taxas As Double
        Dim CodPagamento As Double
        
        Dim Conta As New ContaCorrente
        Dim ValorObrigacao As Double
        Dim Sql As String
        Dim rsTaxas As VSRecordset
        
        Inscricao = Trim(Inscricao)
        Select Case SiglaTributo
            Case Imposto.NomeTributo(ttr_ALVARA)
                ValorObrigacao = Calculo.Alvara(Inscricao, Periodo, CodTributo, DataVence, DocOrigem)
            Case Imposto.NomeTributo(ttr_IPTU)
                If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
                    ValorObrigacao = Calculo.IptuSinfaz(Inscricao, Periodo, CodTributo, DataVence)
                ElseIf AplicacoesVTFuncoes.Municipio = "BARRA MANSA" Then
                    ValorObrigacao = Calculo.IptuBarra(Inscricao, Periodo, CodTributo, DataVence)
                Else
                    ValorObrigacao = Calculo.Iptu(Inscricao, Periodo, CodTributo, DataVence, , Taxas)
                End If
                
                
            Case Imposto.NomeTributo(ttr_ISSQN), Imposto.NomeTributo(ttr_ISSQNSUBST), Imposto.NomeTributo(ttr_ISSQNCOMP), Imposto.NomeTributo(ttr_ISSQNEST), Imposto.NomeTributo(ttr_ISSQNRET)
                If SiglaTributo = Imposto.NomeTributo(ttr_ISSQNEST) Then
                    ValorObrigacao = Calculo.GeraIssEstimativo(Inscricao, Periodo)
                Else
                    ValorObrigacao = Calculo.Issqn(Inscricao, Periodo, CodTributo, DataVence)
                End If
            Case Imposto.NomeTributo(ttr_ISSFIXO)
                If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
                    ValorObrigacao = Nvl(Calculo.Calcula_IssFixo_Barra_Mansa(Inscricao), 0)
                    If ValorObrigacao <> 0 Then
                        'PEGO O DESCONTO DA DATA DE FORMATURA
                        Dim DESCONTO_FORMATURA As Double
                        Dim DESCONTO_FUNCIONARIO_SUS As Double
                        
                        DESCONTO_FUNCIONARIO_SUS = Calculo.Pega_Desconto_Funcionario_SUS(Inscricao)
                        If DESCONTO_FUNCIONARIO_SUS <> 0 Then
                            'APLICO O DESCONTO PELO SUS
                            ValorObrigacao = ValorObrigacao - (DESCONTO_FUNCIONARIO_SUS * ValorObrigacao / 100)
                        End If
                        DESCONTO_FORMATURA = Calculo.Pega_Desconto_Formatura(Inscricao, Periodo)
                        If DESCONTO_FORMATURA <> 0 Then
                            'APLICO O DESCONTO PELA DATA DE FORMATURA
                            ValorObrigacao = ValorObrigacao - (DESCONTO_FORMATURA * ValorObrigacao / 100)
                        End If
                    End If
                Else
                    ValorObrigacao = Calculo.IssqnFixo(Inscricao, Periodo, CodTributo, DataVence)
                End If
            Case Imposto.NomeTributo(ttr_PUBLICIDADE)
                ValorObrigacao = Calculo.PublicidadeAnuncio(Inscricao, CodTributo, DocOrigem, DataVence)
            Case Imposto.NomeTributo(ttr_TFA)
                ValorObrigacao = Calculo.Calcula_Taxas(Inscricao, CodTributo, Periodo)
            Case Imposto.NomeTributo(ttr_TFOP)
                ValorObrigacao = Calculo.Calcula_Taxas(Inscricao, CodTributo, Periodo)
            Case Imposto.NomeTributo(ttr_TFAF)
                ValorObrigacao = Calculo.Calcula_Taxas(Inscricao, CodTributo, Periodo)
            Case Imposto.NomeTributo(ttr_TFV)
                ValorObrigacao = Calculo.Calcula_Taxas(Inscricao, CodTributo, Periodo)
            Case Imposto.NomeTributo(ttr_tfs)
                ValorObrigacao = Calculo.Calcula_Taxas(Inscricao, CodTributo, Periodo)

            Case Else
                ValorObrigacao = Calculo.TaxaFixa(Inscricao, CodTributo, Right(DataVence, 4), MultiplicadorTaxa)
        End Select
        Dim ValImpostoMinimo As Double
        ValImpostoMinimo = Imposto.BuscaValorMinimoImposto(CodTributo, Periodo)
        ValorObrigacao = IIf(ValImpostoMinimo > ValorObrigacao, ValImpostoMinimo, ValorObrigacao)
        CalculaObrigacao = Format(ValorObrigacao, Const_Monetario)
        ValorTaxaInclusaGerada = Taxas
        TipoLancamento = Calculo.TipoLancamento
End Function

Public Function CarregaListaObrigacao(Lista As Object, Optional Inscricao As String, Optional CodTributo As String, Optional Periodo As String, Optional TipoListagem As TipoStatusListagem = etlTodos, Optional NumParcelamento As String, Optional CarregaObrigImoveis As Boolean) As Boolean
    Dim Sql As String
    
    If Temp.PegaParametro(Bdados, "VERIFICA NOTIFICACAO") = "SIM" Then
        If Not LiberaContribuinteNotificado(Inscricao) Then
            Lista.ListItems.Clear
            Exit Function
        End If
    End If
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "select toc_inscricao as Contribuinte, " & _
                "tip_sigla_imposto as [Cod Receita] , toc_periodo as Periodo, toc_data_vencimento as Vencimento, " & _
                "toc_valor_obrigacao as Original, toc_valor_juros as Juros, toc_valor_multa as Multa, " & _
                " tge_nome as Status,toc_tip_cod_imposto,toc_cod_obrigacao AS Obrigacao " & _
                "from tab_obrigacao_contribuinte,tab_imposto,vis_status_obrigacao where  toc_status_obrigacao  = tge_codigo and  toc_tip_cod_imposto = tip_cod_imposto "
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "select toc_inscricao as Contribuinte, " & _
                "tip_sigla_imposto as Cod_Receita , toc_periodo as Periodo, toc_data_vencimento as Vencimento, " & _
                "toc_valor_obrigacao as Original, toc_valor_juros as Juros, toc_valor_multa as Multa, " & _
                " tge_nome as Status,toc_tip_cod_imposto,toc_cod_obrigacao AS Obrigacao " & _
                "from tab_obrigacao_contribuinte,tab_imposto,vis_status_obrigacao where  toc_status_obrigacao  = tge_codigo and  toc_tip_cod_imposto = tip_cod_imposto "
    End If
    'VIS_STATUS_OBRIGACAO
    If TipoListagem = etlParcelaveis Then
        Sql = Sql & " and toc_status_obrigacao not in(" & Const_NaoParcelaveis & ") "
        Sql = Sql & " and toc_data_vencimento < " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
    End If
    If TipoListagem = etlNaoPagosNaoVencidos Then
        Sql = Sql & " and toc_status_obrigacao in (" & Const_NaoPagos & ")"
        Sql = Sql & " and toc_data_vencimento > " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
    End If
    If TipoListagem = etlNaoPagosVencidos Then
        Sql = Sql & " and toc_status_obrigacao in (" & Const_NaoPagos & ")"
        Sql = Sql & " and toc_data_vencimento < " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
    End If
    If TipoListagem = etlPagos Then Sql = Sql & " and toc_status_obrigacao = " & etsCreditoPago
    If TipoListagem = etlNaoPagos Then Sql = Sql & " and toc_status_obrigacao in (" & Const_NaoPagos & ")"
    If Trim(CodTributo) <> "" Then Sql = Sql & " AND toc_tip_cod_imposto ='" & CodTributo & "'"
    If Trim(Inscricao) <> "" Then
        Sql = Sql & " and toc_inscricao = '" & Inscricao & "'"
    End If
    If Trim(NumParcelamento) <> "" Then Sql = Sql & " and TOC_TPA_COD_PARCELAMENTO = " & NumParcelamento
    CarregaListaObrigacao = Lista.Preencher(Bdados, Sql, 1400, 1200, 800, 1200, 900, 900, 900, 2000, 0, 0)
End Function

Public Function CarregaListaObrigacaoVencida(Lista As Object, Optional Inscricao As String, Optional CodTributo As String, Optional Periodo As String, Optional TipoListagem As TipoStatusListagem = etlTodos, Optional NumParcelamento As String) As Boolean
    Dim Sql As String
    If Not LiberaContribuinteNotificado(Inscricao) Then
        Lista.ListItems.Clear
        Exit Function
    End If
    Sql = "select tcc_codigo_conta AS Obrigacao,tcc_inscricao as Contribuinte, " & _
            "tip_sigla_imposto as Cod_Receita , tcc_periodo as Periodo, tcc_data_vencimento as Vencimento, " & _
            "tcc_imposto_original as Original, tcc_juros_atual as Juros, tcc_multa_atual as Multa, " & _
            "tcc_saldo_atual as Total,tcc_tip_cod_imposto " & _
            "from tab_conta_contribuinte,tab_imposto where tcc_tip_cod_imposto = tip_cod_imposto and tcc_data_vencimento < " & Bdados.Converte(Date, TCDataHora)
    If TipoListagem = etlParcelaveis Then Sql = Sql & " and tcc_status_conta not in(" & Const_NaoParcelaveis & ") "
    If TipoListagem = etlPagos Then Sql = Sql & " and tcc_status_conta = " & etsCreditoPago
    If TipoListagem = etlNaoPagos Then Sql = Sql & " and tcc_status_conta in (" & Const_NaoPagos & ")"
    If Trim(CodTributo) <> "" Then Sql = Sql & " AND tcc_tip_cod_imposto ='" & CodTributo & "'"
    If Trim(Inscricao) <> "" Then Sql = Sql & " and tcc_inscricao = '" & Inscricao & "'"
    If Trim(NumParcelamento) <> "" Then Sql = Sql & " and TCC_TPA_COD_PARCELAMENTO = " & NumParcelamento
    CarregaListaObrigacaoVencida = Lista.Preencher(Bdados, Sql, 1200, 1600, 1200, 800, 1200, 900, 900, 900, 900, 0)
End Function

Public Function CarregaListaObrigacaoAtualizada(Lista As Object, Optional Inscricao As String, Optional _
                        CodTributo As String, Optional Periodo As String, Optional TipoListagem As _
                        TipoStatusListagem = etlTodos, Optional NumParcelamento As String, Optional Modo As _
                        TipoInscricaoObrigacao, Optional TrazerSubdividas As Boolean = True) As Boolean
    Dim Sql As String
    Dim SqlAux As String
    Dim Rs As VSRecordset
    If UCase(Temp.PegaParametro(Bdados, "VERIFICA NOTIFICACAO")) = "SIM" Then
        If Not LiberaContribuinteNotificado(Inscricao) Then
            Lista.ListItems.Clear
            Exit Function
        End If
    End If
    Sql = "select tcc_codigo_conta AS Obrigacao,tcc_inscricao as Contribuinte, "
    Sql = Sql & " tip_sigla_imposto as Cod_Receita , tcc_periodo as Periodo, tcc_data_vencimento as Vencimento, "
    Sql = Sql & " tcc_imposto_original as Original, tcc_correcao_monetaria AS Atualizacão ,"
    Sql = Sql & " tcc_juros_atual as Juros, tcc_multa_atual as Multa,TCC_DESCONTO_CONCEDIDO as Desconto, "
    Sql = Sql & " CASE "
    Sql = Sql & " WHEN tcc_imposto_original > 0  THEN tcc_imposto_original + tcc_correcao_monetaria + tcc_juros_atual + tcc_multa_atual - TCC_DESCONTO_CONCEDIDO"
    Sql = Sql & " WHEN tcc_imposto_original = 0 then   tcc_imposto_original + tcc_correcao_monetaria + tcc_juros_atual + tcc_multa_atual"
    Sql = Sql & " END  As Total,"
    Sql = Sql & " tcc_tip_cod_imposto"
    Sql = Sql & " from tab_conta_contribuinte,tab_imposto "
    Sql = Sql & " where  tcc_tip_cod_imposto = tip_cod_imposto  "
    
    If TipoListagem = etlParcelaveis Then
        Sql = Sql & " and tcc_status_conta not in(" & Const_NaoParcelaveis & ") "
    End If
    If TipoListagem = etlPagos Then
        Sql = Sql & " and tcc_status_conta = " & etsCreditoPago
    End If
    
    If Trim(CodTributo) <> "" Then
        Sql = Sql & " AND tcc_tip_cod_imposto ='" & CodTributo & "'"
    End If
    If Trim(Inscricao) <> "" Then
        
        If Modo = etiContribuinte Then
            Sql = Sql & " and ((tcc_inscricao = '" & Inscricao & "' and TCC_TIPO_INSCRICAO = " & Modo & ")"
            If TrazerSubdividas Then
                Dim Inscricoes As String
                SqlAux = "select tim_ic from tab_imovel where  tim_tci_im = '" & Inscricao & "'"
                If Bdados.AbreTabela(SqlAux, Rs) Then
                    Rs.MoveFirst
                    Inscricoes = ""
                    Do While Not Rs.EOF
                        Inscricoes = Inscricoes & "'" & Rs!tim_ic & "',"
                        Rs.MoveNext
                    Loop
                    Inscricoes = Left(Inscricoes, Len(Inscricoes) - 1)
                    Sql = Sql & " or (tcc_inscricao in (" & Inscricoes & ") AND TCC_TIPO_INSCRICAO = 1)"
                End If
            End If
        Else
            Sql = Sql & " and (TCC_TIPO_INSCRICAO = " & Modo
            Sql = Sql & " and tcc_inscricao = '" & Inscricao & "'"
        End If
        Sql = Sql & ")"
    End If
    
    If Trim(NumParcelamento) <> "" Then
        Sql = Sql & " and TCC_TPA_COD_PARCELAMENTO = " & NumParcelamento
    End If
    If Trim(Periodo) <> "" Then
        Sql = Sql & " and TCC_PERIODO = " & Periodo
    End If
    If TipoListagem = etlNaoPagos Then
        Sql = Sql & " and tcc_status_conta in (" & Const_NaoPagos & ") and tcc_imposto_original + tcc_correcao_monetaria + tcc_juros_atual + tcc_multa_atual > 0"
    End If
    If TipoListagem = etlNaoPagosVencidos Then
        Sql = Sql & " and tcc_status_conta in (" & Const_NaoPagos & ") and tcc_imposto_original + tcc_correcao_monetaria + tcc_juros_atual + tcc_multa_atual > 0 and tcc_data_vencimento < " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
    End If
    Dim dataVencimentto As String
    
    
    Sql = Sql & " and tcc_status_conta not in  (6,8,10) "
    Sql = Sql & " order by tcc_inscricao,tcc_periodo asc"
     CarregaListaObrigacaoAtualizada = Lista.Preencher(Bdados, Sql, 1200, 1600, 1200, 800, 1200, 900, 1200, 900, 900, 900, 900, 0)
    
End Function

Public Function CarregaListaObrigacaoNaoVencida(Lista As Object, Optional Inscricao As String, Optional CodTributo As String, Optional Periodo As String, Optional TipoListagem As TipoStatusListagem = etlTodos, Optional NumParcelamento As String) As Boolean
    Dim Sql As String
    If Not LiberaContribuinteNotificado(Inscricao) Then
        Lista.ListItems.Clear
        Exit Function
    End If
    Sql = "select tcc_codigo_conta AS Obrigacao,tcc_inscricao as Contribuinte, " & _
            "tip_sigla_imposto as Cod_Receita , tcc_periodo as Periodo, tcc_data_vencimento as Vencimento, " & _
            "tcc_imposto_original as Original, tcc_juros_atual as Juros, tcc_multa_atual as Multa, " & _
            "tcc_saldo_atual as Total,tcc_tip_cod_imposto " & _
            "from tab_conta_contribuinte,tab_imposto where tcc_tip_cod_imposto = tip_cod_imposto and tcc_data_vencimento > " & Bdados.Converte(Date, TCDataHora)
    If TipoListagem = etlParcelaveis Then Sql = Sql & " and tcc_status_conta not in(" & Const_NaoParcelaveis & ") "
    If TipoListagem = etlPagos Then Sql = Sql & " and tcc_status_conta = " & etsCreditoPago
    If TipoListagem = etlNaoPagos Then Sql = Sql & " and tcc_status_conta in (" & Const_NaoPagos & ")"
    If Trim(CodTributo) <> "" Then Sql = Sql & " AND tcc_tip_cod_imposto ='" & CodTributo & "'"
    If Trim(Inscricao) <> "" Then Sql = Sql & " and tcc_inscricao = '" & Inscricao & "'"
    If Trim(NumParcelamento) <> "" Then Sql = Sql & " and TCC_TPA_COD_PARCELAMENTO = " & NumParcelamento
    CarregaListaObrigacaoNaoVencida = Lista.Preencher(Bdados, Sql, 1200, 1600, 1200, 800, 1200, 900, 900, 900, 900, 0)
End Function
Public Function GeraDAMObrigacao(CodObrigacao As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Conta As New ContaCorrente
    Dim CodPagamento As String
    'Gera qualquer DOC diferente de IPTU (DOCS de IPTU são mais complexos, e são gerados em lugar especifico [VsIptu.cls])
    Sql = "Select tab_obrigacao_contribuinte.*,TAB_IMPOSTO.TIP_COD_CORRELATIVO from tab_obrigacao_contribuinte,TAB_IMPOSTO where " & _
        "TOC_COD_OBRIGACAO =" & CodObrigacao & " and TOC_TIPO_INSCRICAO = 2 AND TOC_TIP_COD_IMPOSTO=tip_cod_imposto"
    If Bdados.AbreTabela(Sql, Rs) Then
        CodPagamento = Conta.Correlativo("TRIB", Rs!TIP_COD_CORRELATIVO)
        Conta.GeraPagamento Trim(Rs!TOC_INSCRICAO), "", Rs!TOC_TIP_COD_IMPOSTO, Rs!TOC_PERIODO, Rs!TOC_DATA_VENCIMENTO, Rs!TOC_VALOR_OBRIGACAO, Rs!TOC_VALOR_MULTA, Rs!TOC_VALOR_JUROS, CDbl(CodPagamento), 0, 0, 0
        CodObrigacao = CodPagamento
    End If
End Function

Public Function IdentificaInscricao(Optional Im As String, Optional IC As String) As String
    If Len(Trim(Edita.TiraPic(Edita.TiraPic(Edita.TiraPic(Im, "."), "-"), "/"))) = 10 Then
        IdentificaInscricao = Im
    Else
        IdentificaInscricao = IC
    End If
End Function


Public Function MostraObrigacaoGerada(Lista As Object, Tributo As String, Optional Contribuinte As String, Optional Restricao As Integer, _
                Optional Status As TipoStatusObrigacao = -1, Optional PeriodoInicial As String, _
                Optional PeriodoFinal As String, Optional ExercicioInicial As String, Optional ExercicioFinal As String, Optional DocOrigem As String, Optional Imovel As String, Optional TipoListagem As TipoStatusListagem = etlTodos, Optional TrazerSubdividas As Boolean = True, Optional CodObrigacao As String, Optional NumParcelamento As String, Optional SubCondicao As String) As Boolean
    Dim Sql As String
    Dim SqlAux As String
    Dim Rs As VSRecordset
    Dim Inscricoes As String
    Dim CampoPeriodo As String
    Contribuinte = Trim(Contribuinte)
    Dim Senha As String
'**********NOTIFICADO*************
    If Temp.PegaParametro(Bdados, "VERIFICA NOTIFICACAO") = "SIM" Then
        If Not LiberaContribuinteNotificado(Contribuinte) Then
            Lista.ListItems.Clear
            MostraObrigacaoGerada = False
            Exit Function
        End If
    End If
    If Bdados.Conexao.FormatoBanco = SQLServer Then CampoPeriodo = " TOC_PERIODO AS PERIODO " 'CampoPeriodo = "CASE LEN(TOC_PERIODO) WHEN 4 THEN CONVERT(VARCHAR(4),TOC_PERIODO) WHEN 6 THEN SUBSTRING(CONVERT(VARCHAR(6),TOC_PERIODO),5,2) + '/' + SUBSTRING(CONVERT(VARCHAR(6),TOC_PERIODO),1,4) END AS PERIODO"
    If Bdados.Conexao.FormatoBanco = oracle Then CampoPeriodo = " TOC_PERIODO AS PERIODO "
    ExercicioInicial = IIf(Len(ExercicioInicial) = 4, ExercicioInicial, Right(ExercicioInicial, 4) & Left(ExercicioInicial, 2))
    ExercicioFinal = IIf(Len(ExercicioFinal) = 4, ExercicioFinal, Right(ExercicioFinal, 4) & Left(ExercicioFinal, 2))
    Screen.MousePointer = 11
    Sql = "SELECT TOC_COD_OBRIGACAO as Documento,TOC_INSCRICAO as INSCRICAO,TOC_PERIODO ,TIP_SIGLA_IMPOSTO AS TRIBUTO," & CampoPeriodo & " ," & _
        " TOC_DATA_VENCIMENTO AS VENCTO,TOC_VALOR_OBRIGACAO AS VALOR ," & _
        " TOC_VALOR_JUROS AS JUROS,TOC_VALOR_MULTA as MULTA,TGE_NOME as SITUACAO,TOC_TOTAL_TAXA_INCLUSA as TAXA," & _
        " TOC_TIP_COD_IMPOSTO,TOC_PARCELA AS PARCELA,TOC_OBSERVACAO AS OBSERVACAO,TOC_NUM_DOC_ORIGEM AS DOC_ORIGEM,TOC_STATUS_OBRIGACAO,TOC_TPA_COD_PARCELAMENTO AS " & _
        " Parcelamento,TOC_CORRECAO_MONETARIA as Correcao,TOC_DESCONTO AS Desconto,TOC_DATA_GERACAO AS Geracao,TOC_OBS AS Observação FROM TAB_OBRIGACAO_CONTRIBUINTE," & _
        "TAB_IMPOSTO,VIS_STATUS_OBRIGACAO "
        
    Sql = Sql & " WHERE TOC_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO AND " & _
            " TOC_STATUS_OBRIGACAO = TGE_CODIGO "
    If Trim(Tributo) <> "" Then Sql = Sql & " AND TOC_TIP_COD_IMPOSTO ='" & Tributo & _
                "' "
    If Trim(ExercicioInicial) <> "" And Trim(ExercicioFinal) <> "" Then
        Sql = Sql & " AND TOC_PERIODO BETWEEN " & ExercicioInicial & " AND " & ExercicioFinal
    End If
    If Trim(PeriodoInicial) <> "" And Trim(PeriodoFinal) <> "" Then
        Sql = Sql & " AND (TOC_DATA_GERACAO >= " & Bdados.Converte(PeriodoInicial, TCDataHora) & " AND TOC_DATA_GERACAO  <= " & Bdados.Converte(PeriodoFinal, TCDataHora) & ")"
    End If
    If Trim(CodObrigacao) <> "" Then
        Sql = Sql & " AND TOC_COD_OBRIGACAO = " & CodObrigacao
    End If
    If Status > 0 Then Sql = Sql & " AND TOC_STATUS_OBRIGACAO=" & Status
    If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
        If Val(DocOrigem) > 0 Then Sql = Sql & " AND TOC_NUM_DOC_ORIGEM=" & DocOrigem
    End If
    If Restricao = 1 Then Sql = Sql & " AND TOC_STATUS_OBRIGACAO not in (" & etsCreditoPago & "," & etsCreditoIsento & ")"
    If Restricao = 2 Then Sql = Sql & " AND TOC_STATUS_OBRIGACAO = " & etsCreditoPago
'    If Restricao = 3 Then Sql = Sql & " AND (TOC_STATUS_OBRIGACAO NOT IN (" & etsCreditoIsento & "," & etsSemMovimento & "," & etsCreditoParcelado & "," & etsCreditoParceladoOficio & "," & etsCreditoPago & " )) AND toc_data_vencimento < " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
'    If Tributo <> Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) Then
        If Trim(Contribuinte) <> "" Then
            If VBA.InStr(1, Contribuinte, "-") > 0 Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then  'INSCRICAO MUNICIPAL
                If TrazerSubdividas Then
                    SqlAux = "select tim_ic from tab_imovel where  tim_tci_im = '" & Contribuinte & "'"
                    If Bdados.AbreTabela(SqlAux, Rs) Then
                        Rs.MoveFirst
                        Inscricoes = ""
                        Do While Not Rs.EOF
                            Inscricoes = Inscricoes & "'" & Trim(Rs!tim_ic) & "',"
                            Rs.MoveNext
                        Loop
                        Inscricoes = Left(Inscricoes, Len(Inscricoes) - 1)
                        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Or True Then
                            Sql = Sql & " AND ((TOC_INSCRICAO = '" & Trim(Contribuinte) & "' AND TOC_TIPO_INSCRICAO = 2) or (toc_inscricao in (  " & Trim(Inscricoes) & ") AND TOC_TIPO_INSCRICAO = 1))"
                        Else
                            Sql = Sql & " AND ((TOC_INSCRICAO = '" & Trim(Contribuinte) & "' or toc_inscricao in (  " & Inscricoes & "))"
                        End If
                    Else
                        Sql = Sql & " AND TOC_INSCRICAO = '" & Trim(Contribuinte) & "' AND TOC_TIPO_INSCRICAO = 2"
                    End If
                Else
                    Sql = Sql & " AND TOC_INSCRICAO = '" & Trim(Contribuinte) & "' AND TOC_TIPO_INSCRICAO = 2"
                End If
            End If
        ElseIf Trim(Imovel) <> "" Then   ' INSCRICAO CADASTRAL(imovel)
                Sql = Sql & " AND (TOC_INSCRICAO = '" & Trim(Imovel) & "' AND TOC_TIPO_INSCRICAO = 1)"
        End If
    If TipoListagem = etlPagos Then Sql = Sql & " and toc_status_obrigacao = " & etsCreditoPago
    If TipoListagem = etlNaoPagos Then Sql = Sql & " and toc_status_obrigacao in (" & Const_NaoPagos & ")"
    If TipoListagem = etlNaoPagosNaoVencidos Then
        Sql = Sql & " and toc_status_obrigacao in (" & Const_NaoPagos & ")"
        Sql = Sql & " and toc_data_vencimento > " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
    End If
    If TipoListagem = etlNaoPagosLancados Then Sql = Sql & " and toc_status_obrigacao in (" & Const_NaoPagosTodos & ")"
    If TipoListagem = etlNaoPagosVencidos Then
        Sql = Sql & " and toc_status_obrigacao in (" & Const_NaoPagos & ")"
        Sql = Sql & " and toc_data_vencimento < " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
    End If
    If TipoListagem = 8 Then
        Sql = Sql & " and toc_status_obrigacao in (2,5,10)"
    End If
    If NumParcelamento <> "" Then
        Sql = Sql & " AND TOC_TPA_COD_PARCELAMENTO = '" & NumParcelamento & "'"
    End If
    If SubCondicao <> "" Then
        Sql = Sql & " " & SubCondicao
    End If
    Sql = Sql & " ORDER BY PERIODO"
    '--DEVIDO A SELIC
        Dim Conta As ContaCorrente
        Dim rsOb As VSRecordset
    MostraObrigacaoGerada = Lista.Preencher(Bdados, Sql, 1200, 1500, 0, 1100, 1100, 1100, 1150, 0, 0, 2100, 1000, 0, 1000, 1300, 0, 0, 0, 0, 0, 1300, 0)
      If Lista.ListItems.Count > 0 Then Lista.Mensagem = "Total Lancado: R$" & Format(Lista.Colunas(7).Soma, Const_Monetario)
    Screen.MousePointer = 0
End Function

Public Function PossuiObrigacaoNoExercicio(Inscricao As String, Exercicio As String) As Boolean
    'Funcao para checagem se contribuinte possui obrigacao do tributo no periodo
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim CondMesInicial As String
    Dim CondMesFinal As String
    PossuiObrigacaoNoExercicio = False
    If Len(Trim(Exercicio)) = 6 Then 'Periodo (AAAAMM)
        CondMesInicial = "right(convert(char(10),TPO_PERIODO_INICIAL,103),4) + substring(convert(char(10),TPO_PERIODO_INICIAL,103),4,2)"
        CondMesFinal = "right(convert(char(10),TPO_PERIODO_FINAL,103),4) + substring(convert(char(10),TPO_PERIODO_FINAL,103),4,2)"
    Else 'Perido(AAAA)
        CondMesInicial = "right(convert(char(10),TPO_PERIODO_INICIAL,103),4)"
        CondMesFinal = "right(convert(char(10),TPO_PERIODO_FINAL,103),4) "
    End If
    Sql = "Select TPO_INSCRICAO FROM TAB_PERIODO_OBRIGACAO WHERE " & _
            " TPO_INSCRICAO = '" & Inscricao & "' AND " & CondMesInicial & " <= " & Exercicio & _
            " AND  (" & CondMesFinal & " >= " & Exercicio & " OR " & CondMesFinal & " IS NULL)"
    PossuiObrigacaoNoExercicio = Bdados.AbreTabela(Sql, Rs)
    Bdados.FechaTabela Rs
End Function

Public Sub PreencheComboTributo(Combo As Object, Optional SomenteComObrigacao As Boolean = True, Optional TipoTributo As etcTipoComboTributo)
    Dim Sql As String
     Sql = "Select  tip_cod_imposto,tip_sigla_imposto " & Bdados.Concatena & "  ' # ' " & Bdados.Concatena & " tip_nome_imposto "
     Sql = Sql & " From TAB_IMPOSTO "
    If SomenteComObrigacao Then
        Sql = Sql & " where tip_cod_imposto in (select tpi_tip_cod_imposto from tab_parametro_imposto where " & _
                " TPI_GERA_OBRIGACAO = 1) "
    End If
    If TipoTributo = etcTributario Then
        Sql = Sql & " AND TIP_NATUREZA = 1"
    ElseIf TipoTributo = etcNaoTributario Then
        Sql = Sql & " AND TIP_NATUREZA = 2"
    End If
    Sql = Sql & " order by tip_sigla_imposto asc"
    Combo.Preencher Bdados, Sql, 1
End Sub

Public Function SelecionaContribuinte(Contribuinte As String, tipoInscricao As TipoInsc) As VSRecordset
    Dim Sql As String
    
    On Error GoTo Trata
    Select Case tipoInscricao
        Case 1 'IC Obrigatoria
            Sql = "Select tim_ic from Tab_Imovel "
            If Trim(Contribuinte) <> "" Then Sql = Sql & " where tim_ic = '" & Contribuinte & "'"
        Case 0, 2 'IC nao Obrigatoria
            Sql = "Select tci_im from Tab_Contribuinte "
            If Trim(Contribuinte) <> "" Then Sql = Sql & " where tci_im = '" & Contribuinte & "'"Rem or tci_cgc_cpf = '" & Contribuinte & "'" ' AND tci_tipo_recolhimento_iss IN (1,2)"
    End Select
    
    If Not Bdados.AbreTabela(Sql, SelecionaContribuinte) Then
        Sql = "Select VIN_INSCRICAO from VIS_INSCRICAO where VIN_INSCRICAO = '" & Contribuinte & "'"
        If tipoInscricao = 1 And Len(Trim(Contribuinte)) > 0 Then
            Sql = "SELECT VIN_INSCRICAO FROM  VIS_INSCRICAO where "
            If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
                If CInt(Right(Contribuinte, 3)) <> 200 Then
                    Sql = Sql & "  VIN_INSCRICAO ='" & Contribuinte & "'"
                Else
                    Sql = Sql & "  VIN_INSCRICAO > '" & Contribuinte & "' and  VIN_INSCRICAO  <= '" & Left(Contribuinte, 12) & "300'"
                End If
            Else
                Sql = Sql & "  VIN_INSCRICAO ='" & Contribuinte & "'"
            End If
            Sql = Sql & " order by VIN_INSCRICAO ASC"
            
'            Sql = "Select VIN_INSCRICAO from VIS_INSCRICAO where VIN_INSCRICAO like '" & Contribuinte & "%'"
            'feito para a geracao por grupos, setores etc...
        End If
        Call Bdados.AbreTabela(Sql, SelecionaContribuinte)
   End If
   Exit Function
Trata:
    Avisa Err.Description
    Exit Function
    Resume
End Function

Public Function CriaObrigacao(ByVal Tributo As String, ByVal PeriodoInicial As String, ByVal PeriodoFinal As String, _
                Optional Contribuinte As String, Optional ValorObrigacao As Double = 0, Optional Status As TipoStatusObrigacao = -1, _
                Optional SubstituirExistente As TipoSubstituicaoObrigacao = etsSubstitui, Optional Vencimento As String, Optional Progresso As Object, _
                Optional CodigoObrigacao As String, Optional PorGrupo As Byte, Optional ByRef UltGerado As Object, Optional Incidencia As Integer, _
                Optional MultiplicadorTaxa As Double, Optional ByVal DocOrigem As String, Optional Parcela As Integer, Optional Imovel As String, _
                Optional Modo As TipoInscricaoObrigacao, Optional Remessa As Boolean) As String
                
    On Error GoTo Trata
    Dim Sql As String
    Dim RsContrib As VSRecordset
    Dim RsTrib As VSRecordset
    Dim DataVence As String
    Dim Periodo As Double
    Dim TipoLancamento As TipoLanc
    Dim Regs As Double
    Dim ValorFixo As Double
    Dim InscGerado As String
    Dim Tipo As Integer
    Dim Taxas As Double
    Dim Ano As String
    
    NumeroIncidenciaPeriodo = Incidencia
    CriaObrigacao = 0
    Contribuinte = Trim(Contribuinte)
    PeriodoInicial = IIf(Len(PeriodoInicial) = 4, PeriodoInicial, Right(PeriodoInicial, 4) & Left(PeriodoInicial, 2))
    PeriodoFinal = IIf(Len(PeriodoFinal) = 4, PeriodoFinal, Right(PeriodoFinal, 4) & Left(PeriodoFinal, 2))
    
    ValorFixo = ValorObrigacao
    If Trim(CodigoObrigacao) <> "" Then
        Sql = "Select TOC_COD_OBRIGACAO from TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO = " & CodigoObrigacao
        If Bdados.AbreTabela(Sql, RsTrib) Then
            CriaObrigacao = CodigoObrigacao
            If SubstituirExistente = etsNaoSubstitui Then Exit Function
        End If
    End If
    Sql = "SELECT tpi_tip_cod_imposto,tip_sigla_imposto,TPI_PERIODO_CALCULA,TPI_PERIODO_DECLARA, tpi_tipo_ic from " & _
            " TAB_IMPOSTO,TAB_PARAMETRO_IMPOSTO WHERE tpi_tip_cod_imposto = tip_cod_imposto " & _
            " AND tpi_tip_cod_imposto ='" & Tributo & "'"
    Ano = Imposto.BuscaAnoImposto(Tributo, Right(Periodo, 4))
    If Trim(Ano) <> "" Then
        Sql = Sql & " AND tpi_ano_imposto ='" & Ano & "'"
    Else
        Sql = Sql & " AND (tpi_ano_imposto ='' or tpi_ano_imposto is null) "
    End If
    If Bdados.AbreTabela(Sql, RsTrib) Then
        If Modo = 0 Then
            Modo = RsTrib.Fields("tpi_tipo_ic")
        End If
        If Not Progresso Is Nothing Then Progresso.Min = 0
        RsTrib.MoveFirst
        Do While Not RsTrib.EOF
            ' problema do parcelamento é que tem que identificar o registro que está sendo gravado e guardar a inscrição certa
            ' senão a linha abaixo de selecionar contribuinte vai dar erro.
            If Contribuinte = "" Then Contribuinte = Imovel
            Set RsContrib = SelecionaContribuinte(Contribuinte, Modo)
'            If PorGrupo = 1 Then
'                Set RsContrib = SelecionaContribuinte(Contribuinte, 3)
'            Else
'                Set RsContrib = SelecionaContribuinte(Contribuinte, Nvl("" & RsTrib!tpi_tipo_ic, 2))
'            End If
            RsContrib.MoveFirst
            PeriodoInicial = Val(PeriodoInicial)
            PeriodoFinal = Val(PeriodoFinal)
            For Periodo = CDbl(PeriodoInicial) To CDbl(PeriodoFinal)
                If Trim(Vencimento) = "" Then
                    DataVence = Imposto.BuscaDataVencimento("" & RsTrib!tpi_tip_cod_imposto, CStr(IIf("" & RsTrib!TPI_PERIODO_DECLARA = 1 Or (Len(CStr(Periodo)) = 4), Left(Periodo, 4), Right(CStr(Periodo), 2) & Left(CStr(Periodo), 4))))
                    If Trim(DataVence) = "" Then
                        CriaObrigacao = "0"
                        Exit Function
                    End If
                Else
                    DataVence = Vencimento
                End If
                If Not Progresso Is Nothing Then Progresso.Max = RsContrib.Fields.Count
                If Not Progresso Is Nothing Then Progresso.Visible = True
                Regs = 0
                ValorObrigacao = 0
                
                Do While Not RsContrib.EOF
                    If PossuiObrigacaoNoExercicio(Trim(RsContrib(0)), CStr(Periodo)) Then
                        If ValorFixo = 0 Then
                            ValorObrigacao = CalculaObrigacao(Trim(RsContrib(0)), Nvl("" & RsTrib!tpi_tipo_ic, 2), "" & RsTrib!tpi_tip_cod_imposto, _
                                    "" & RsTrib!TIP_SIGLA_IMPOSTO, Periodo, DataVence, TipoLancamento, , DocOrigem, Taxas)
                        Else
                            ValorObrigacao = ValorFixo
                        End If
                        
                        
                        'REMARCADO POR HENRIQUE PELO FATO DA LINHA SER RELEVANTE!
                  '      If (AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE") And RsTrib!tpi_tip_cod_imposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) Then
                  '          DataVence = Imposto.BuscaDataVencimento("" & RsTrib!tpi_tip_cod_imposto, CStr(IIf("" & RsTrib!TPI_PERIODO_DECLARA = 1 Or (Len(CStr(Periodo)) = 4), Left(Periodo, 4), Right(CStr(Periodo), 2) & Left(CStr(Periodo), 4))))
                  '      End If
                        
                        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                            Calculo.GravaObrigacao = True
                        Else
                            Calculo.GravaObrigacao = IIf(ValorObrigacao > 0, True, False)
                        End If
                        'ver dam 0 no parametro
                        'CHECO SE NA DEFINIÇAO DE TRIBUTOS PODE SER GERADO COM O VALOR ZERO
                        If ValorObrigacao = 0 Then
                            If Not Valida_Geracao_Triuto(Tributo, CDbl(Periodo)) Then Exit Function
                        End If
                        ValorObrigacao = ValorObrigacao
                        
                        If Not IsDate(Format(DataVence, "dd/mm/yyyy")) Then
                            DataVence = Format(Date, "dd/mm/yyyy")
                        End If
                        CriaObrigacao = GravaObrigacao(Trim(RsContrib(0)), Modo, "" & RsTrib!tpi_tip_cod_imposto, Periodo, ValorObrigacao, _
                                     DataVence, TipoLancamento, IIf(Status > 0, Status, IIf(TipoLancamento = etlDeclaracao And ValorObrigacao = 0, 1, 2)), _
                                     Progresso, SubstituirExistente, CodigoObrigacao, NumeroIncidenciaPeriodo, DocOrigem, Parcela, Taxas, Remessa)
                        
                        If Trim(CriaObrigacao) <> "" Then Conta.CriaContaContribuinte (CriaObrigacao), , , Modo
'                        If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
'                            If Tributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) Or Tributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_tfs)) Then
'                                If CriaObrigacao <> "" Then
'                                    'ALTERO A SITUAÇÃO DO ALVARA PARA IMPRESSO...
'                                    If Tributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) Then
'                                        Bdados.GravaDados "TAB_MOVIMENTO_CONTRIBUINTE", Bdados.PreparaValor(Bdados.Converte("1", tctexto)), "TMC_LIBERACAO_ALVARA", "TMC_CODIGO_MOVIMENTACAO  = '" & Processo & "'"
'                                    ElseIf Tributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_tfs)) Then
'                                        Bdados.GravaDados "TAB_MOVIMENTO_CONTRIBUINTE", Bdados.PreparaValor(Bdados.Converte("1", tctexto)), "TMC_LIBERACAO_SANITARIA", "TMC_CODIGO_MOVIMENTACAO  = '" & Processo & "'"
'                                    End If
'                                    Processo = ""
'                                End If
'                            End If
'                        End If
                                                                   
                    End If
                    Regs = Regs + 1
                    InscGerado = RsContrib(0)
                    If Not UltGerado Is Nothing Then UltGerado = "Registro nº " & Regs & "  -  " & InscGerado
                    RsContrib.MoveNext
                    If Not Progresso Is Nothing Then Progresso.Value = Regs
                    DoEvents
                Loop
                RsContrib.MoveFirst
PROXIMO_ALVARA:
            Next
            RsTrib.MoveNext
        Loop
    Else
'        Util.Avisa "Erro ao gerar obrigação, verifique configurações na Definição de Tributos."
    End If
    Bdados.FechaTabela RsContrib
    Bdados.FechaTabela RsTrib
    obCodigoObrigacao = CriaObrigacao
    Exit Function
Trata:
    If Err.Number = 91 Or Err.Number = -2147467259 Then
        Resume Next
'        Resume
    Else
        Avisa Err.Description
        Exit Function
        Resume
    End If
End Function

Public Function TrocaSitObrigacao(CodObrigacao As String, Situacao As TipoStatusObrigacao) As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim TabelaDAT As String
    
    TrocaSitObrigacao = Bdados.Executa("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_STATUS_OBRIGACAO =" & IIf(Situacao < 1, etsCreditoOriginalAberto, Situacao) & _
                        ", TOC_STATUS_ANTERIOR_OBRIGACAO = TOC_STATUS_OBRIGACAO WHERE TOC_COD_OBRIGACAO= " & CodObrigacao)
    Conta.TrocaStatusConta CodObrigacao, IIf(Situacao < 1, etsCreditoOriginalAberto, Situacao)
    
'    Campos = "TOC_STATUS_OBRIGACAO" & IIf(Situacao < 1, ",", "")
'    Valores = Bdados.PreparaValor(IIf(Situacao < 1, etsCreditoOriginalAberto, Situacao))
'    If Situacao < 1 Then
'        Valores = Valores & Bdados.PreparaValor(etsCreditoOriginalAberto)
'    End If
'    TrocaSitObrigacao = Bdados.AtualizaDados("TAB_OBRIGACAO_CONTRIBUINTE", Valores, Campos, "TOC_COD_OBRIGACAO= " & CodObrigacao)
End Function

Public Function GravaPeriodoObrigacao(Inscricao As String, PeriodoInicial As String, _
                 Optional PeriodoFinal As String, Optional CodTributo As String)
    Dim Valores As String
    Dim Campos As String
    
    Campos = "TPO_INSCRICAO,TPO_PERIODO_INICIAL,TPO_PERIODO_FINAL"
    Valores = Bdados.PreparaValor(Bdados.Converte(Inscricao, tctexto), _
                        Bdados.Converte(PeriodoInicial, TCDataHora))
    If Trim(PeriodoFinal) <> "" Then
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(PeriodoFinal, TCDataHora))
    Else
        Valores = Valores & Bdados.PreparaValor("Null")
    End If
    GravaPeriodoObrigacao = Bdados.GravaDados("TAB_PERIODO_OBRIGACAO", Valores, Campos, _
            "TPO_INSCRICAO= '" & Inscricao & "' and TPO_PERIODO_INICIAL = " & _
            Bdados.Converte(PeriodoInicial, TCDataHora)) '& " and TPO_PERIODO_FINAL = " & _
            Bdados.Converte(PeriodoFinal, TCDataHora))
End Function

Public Function GravaObrigacao(Inscricao As String, tipoInscricao As TipoInsc, CodTributo As String, _
                Periodo As Double, ValorObrigacao As Double, DataVence As String, Optional TipoLancamento As TipoLanc, _
                Optional StatusObrigacao As TipoStatusObrigacao = etsCreditoOriginalAberto, Optional Progresso As Object, _
                Optional SubstituirExistente As TipoSubstituicaoObrigacao = etsSubstitui, Optional CodigoObrigacao As String, _
                Optional Incidencia As Integer, Optional DocOrigem As String, Optional Parcela As Integer, _
                Optional TaxaInclusa As Double, Optional Remessa As Boolean) As Double
                
    Dim Vencimento As String, Multa As Double
    Dim Juros As Double
    
    Dim Sql As String
    Dim RsImp As VSRecordset
    Dim Campos As String
    Dim Valores As String
    Dim CodObrigacao As String
    Dim Condicao As String
    Dim Txea As String
    Dim Tsd As String
    Dim TXDAM As String
    Dim Rs As VSRecordset
    Dim TxP As Double
    'BCP
    Dim obsObrig, basaNossoNumero As String
    Dim vtx As Double
    obsObrig = ""
    vtx = 0
    If CodTributo = 11130501 And Left(DocOrigem, 2) = 65 Then ' SO QUANDO FOR NOTA AVULSA
        obsObrig = "DOC. REF. AO TRIBUTO ISSQN DA NOTA FISCAL " & DocOrigem '1179=CODO
        Remessa = True 'AUTOMATICAMENTE O DAM VAI PARA A REMESSA
    End If
    If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 And Left(DocOrigem, 2) = 65 Then ' SO QUANDO FOR NOTA AVULSA Then 'CODO
        
        If CodTributo = 11130501 Then 'ISSQN
            TaxaInclusa = Temp.PegaParametro(Bdados, "TXTDAM")
            
        End If
    End If
    
    'FIM BCP

    On Error GoTo Trata
    Sql = "SELECT TPI_GERA_TAXA_IMPRESSAO FROM TAB_PARAMETRO_IMPOSTO WHERE TPI_TIP_COD_IMPOSTO = '" & CodTributo & "'"
  '  If Bdados.AbreTabela(Sql, rs) Then
   '     If CInt(Nvl("" & rs!TPI_GERA_TAXA_IMPRESSAO, 1)) = 1 Then TaxaInclusa = Imposto.TaxaFixaImpressaoDAM
    'End If
   ' Bdados.FechaTabela rs
    'Gambi
    
'    TaxaInclusa = 0
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        If CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) Then
            Txea = TrocaPic(Temp.PegaParametro(Bdados, "txea"), ".", ",")
            TxP = Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXP"), ".", ","), 0)
            TaxaInclusa = CDbl(Nvl(Txea, 0)) + TxP
        End If
        Tsd = TrocaPic(Temp.PegaParametro(Bdados, "TSD"), ".", ",")
        TXDAM = TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
        TaxaInclusa = TaxaInclusa + CDbl(Nvl(TXDAM, 0)) '+ CDbl(Tsd)
    End If
    'STATUS OBRIGACAO
    '1: Credito nao lancado
    '2: Credito Original em aberto
    '3: Credito pago
    '4: Credito em parcelamento
    '5: Credito em Divida Ativa
    '6. Credito Isento
    '7. Credito Cancelado
    DocOrigem = Edita.TiraTudo(DocOrigem)
    Campos = "TOC_COD_OBRIGACAO,TOC_INSCRICAO,TOC_TIP_COD_IMPOSTO,TOC_PERIODO,TOC_TIPO_INSCRICAO,TOC_DATA_VENCIMENTO," & _
                "TOC_DATA_GERACAO,TOC_VALOR_OBRIGACAO,TOC_VALOR_MULTA,TOC_VALOR_JUROS,TOC_STATUS_OBRIGACAO," & _
                "TOC_STATUS_ANTERIOR_OBRIGACAO,TOC_TIPO_LANCAMENTO,TOC_Indice,TOC_TOTAL_TAXA_INCLUSA," & _
                "TOC_NUM_DOC_ORIGEM,TOC_PARCELA,TOC_SITUACAO_TAXA,TOC_OBSERVACAO,TOC_DESCONTO,TOC_REMESSA,TOC_NOSSO_NUMERO" 'CODIGO DA REMESSA COM VALOR PADRAO 1, PARA QUANDO DE FATO PAGAR SE TORNAR NULO E IR NA CONSULTA DE BOLETOS PARA A REMESSA
    If Calculo.GravaObrigacao Then
        If CodigoObrigacao <> "" And CodigoObrigacao <> "0" Then
            CodObrigacao = CodigoObrigacao
            GravaObrigacao = CodObrigacao
            Sql = "Select TOC_COD_OBRIGACAO from TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO = " & CodObrigacao
            If Bdados.AbreTabela(Sql, RsImp) Then
                If SubstituirExistente = etsNaoSubstitui Then Exit Function
            End If
        Else
            Campos = "TOC_COD_OBRIGACAO,TOC_INSCRICAO,TOC_TIP_COD_IMPOSTO,TOC_PERIODO,TOC_TIPO_INSCRICAO," & _
                    "TOC_DATA_VENCIMENTO,TOC_DATA_GERACAO,TOC_VALOR_OBRIGACAO," & _
                    "TOC_VALOR_MULTA,TOC_VALOR_JUROS,TOC_STATUS_OBRIGACAO,TOC_STATUS_ANTERIOR_OBRIGACAO,TOC_TIPO_LANCAMENTO,Toc_indice,TOC_TOTAL_TAXA_INCLUSA,TOC_NUM_DOC_ORIGEM,TOC_PARCELA,TOC_SITUACAO_TAXA,TOC_OBSERVACAO,TOC_DESCONTO,TOC_REMESSA,TOC_NOSSO_NUMERO"
            
            Sql = "Select TOC_COD_OBRIGACAO,toc_indice from TAB_OBRIGACAO_CONTRIBUINTE where TOC_INSCRICAO ='" & Trim(Inscricao) & "'"
            Sql = Sql & " and TOC_PERIODO =" & Periodo
            Sql = Sql & " and  TOC_TIP_COD_IMPOSTO ='" & CodTributo & "' and TOC_PARCELA = " & Parcela
            If Trim(DocOrigem) = "" Or CStr(Nvl(Trim(DocOrigem), 0)) = "0" Then
                If Bdados.Conexao.FormatoBanco = SQLServer Then
                    Sql = Sql & " AND (toc_num_doc_origem IS NULL OR toc_num_doc_origem =''  or toc_num_doc_origem ='0')"
                ElseIf Bdados.Conexao.FormatoBanco = oracle Then
                    Sql = Sql & " AND (toc_num_doc_origem IS NULL OR toc_num_doc_origem =0)"
                End If
            Else
                Sql = Sql & " AND toc_num_doc_origem = '" & DocOrigem & "'"
            End If
            Sql = Sql & " ORDER BY toc_indice DESC"
            If Not Bdados.AbreTabela(Sql, RsImp) Or SubstituirExistente = etsCriaNova Then
                If CodigoObrigacao <> "" And CStr(CodObrigacao) <> "0" Then
                    CodObrigacao = CodigoObrigacao
                    GravaObrigacao = CodObrigacao
                    If Not RsImp!TOC_INDICE = "" Or Not IsNull(RsImp!TOC_INDICE) Then
                        NumeroIncidenciaPeriodo = Nvl("" & RsImp!TOC_INDICE, 0)
                    End If
                Else
                    CodObrigacao = Conta.GeraCodPagamento(CodTributo)
                    'CodObrigacao = Bdados.Correlativo("TRIB", Const_Obrig, "OBRIGACAO TRIBUTARIA")
                    If Temp.PegaParametro(Bdados, "BANCO ARRECADACAO") = 3 And Remessa = True Then 'Banco BASA quando gera remessa na tela
                        basaNossoNumero = Conta.GeraBasaNossoNumero()
                    Else 'DEMAIS BANCOS NÃO POSSUEM CONTROLE DE BOLETOS
                        
                        basaNossoNumero = CodObrigacao
                    End If
                    
                    If RsImp.EOF Then
                        NumeroIncidenciaPeriodo = 0
                    Else
                        If SubstituirExistente = etsCriaNova Then NumeroIncidenciaPeriodo = Nvl("" & RsImp!TOC_INDICE, 0) + 1
                    End If
                End If
            Else
                CodObrigacao = "" & RsImp!TOC_COD_OBRIGACAO
                GravaObrigacao = CodObrigacao
                If SubstituirExistente = etsCriaNova Then
                    CodObrigacao = Conta.GeraCodPagamento(CodTributo)
                    NumeroIncidenciaPeriodo = Nvl("" & RsImp!TOC_INDICE, 0) + 1
                End If
                If SubstituirExistente = etsNaoSubstitui Then Exit Function
            End If
        End If
        GravaObrigacao = CodObrigacao
        
        Sql = "SELECT tpi_tip_cod_imposto,tpi_ano_imposto,tpi_dias_pagar,tpi_periodic_calculo," & _
            "tpi_tipo_contribuinte,tpi_valor_min_multa,tpi_valor_max_multa,tpi_valor_juros," & _
            "tpi_aliquota,tpi_valor_taxa_fixa,tpi_juros_captalizados from TAB_PARAMETRO_IMPOSTO "
        If Bdados.AbreTabela(Sql, RsImp) Then
'            If Parcela = 0 Then
'                condicao = "TOC_COD_OBRIGACAO = " & CodObrigacao
'            Else
                Condicao = "TOC_INSCRICAO ='" & Trim(Inscricao) & "' and TOC_PERIODO =" & Periodo & " and  TOC_TIP_COD_IMPOSTO ='" & CodTributo & "' and TOC_PARCELA = " & Parcela
'            End If
            'ValorObrigacao = ValorObrigacao + TaxaInclusa
            Valores = Bdados.PreparaValor(CodObrigacao, Bdados.Converte(Trim(Inscricao), tctexto), Bdados.Converte(CodTributo, tctexto), Bdados.Converte(Periodo, tctexto), Bdados.Converte(tipoInscricao, tctexto), _
                    Bdados.Converte(DataVence, TCDataHora), Bdados.Converte(Date, _
                    TCDataHora), Bdados.Converte(Format(ValorObrigacao, "###,###,###,##0.00"), TCMonetario), Bdados.Converte(MultaObrigacao, TCDuplo), _
                    Bdados.Converte(JurosObrigacao, TCDuplo), StatusObrigacao, StatusObrigacao, TipoLancamento, NumeroIncidenciaPeriodo, TaxaInclusa, Bdados.Converte(Nvl(DocOrigem, 0), tctexto), Parcela, 2, obsObrig, 0, 1, basaNossoNumero)
            Bdados.GravaDados "TAB_OBRIGACAO_CONTRIBUINTE", Valores, Campos, "TOC_COD_OBRIGACAO = " & CodObrigacao 'Condicao
'            Bdados.GravaDados "TAB_OBRIGACAO_CONTRIBUINTE", Valores, Campos, _
                    "TOC_INSCRICAO ='" & Trim(Inscricao) & "' AND TOC_TIP_COD_IMPOSTO ='" & CodTributo & _
                    "' AND TOC_PERIODO =" & Periodo
            GravaObrigacao = CodObrigacao
            'Resume
        End If
        
        'Gera DAM da Obrigacao(Não gera Dam para Credito nao lancado)
'        If TipoLancamento <> etlDeclaracao Then
'            If ValorObrigacao > 0 Then GeraDAMObrigacao CodObrigacao
'        End If
        'If Not (TipoLancamento = etlDeclaracao And ValorObrigacao = 0) Then GeraDAMObrigacao CodObrigacao
        
        'GRAVA OS DETALHES DA OBRIGACAO(OBRIGAÇÕES ACESSORIAS)
        If Calculo.GravaObrigacaoAcessoria Then
            If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
                GravaObrigacaoAcessoria CodObrigacao, CodTributo, Imposto.TaxaFixaImpressaoDAM
                GravaObrigacaoAcessoria CodObrigacao, CodTributo, Imposto.TaxaFixaImpressaoDAM
            Else
                If Calculo.TaxaUnificada Then
                    GravaObrigacaoAcessoria CodObrigacao, CodTributo, ValorObrigacao - Imposto.ValorTaxas
                    GravaObrigacaoAcessoria CodObrigacao, Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_TSU)), Imposto.ValorTaxas
                Else
                    GravaObrigacaoAcessoria CodObrigacao, Calculo.DadosIptu.CodImpostoConservacao, Calculo.DadosIptu.ValorTCVL
                    GravaObrigacaoAcessoria CodObrigacao, Calculo.DadosIptu.CodImpostoLimpeza, Calculo.DadosIptu.ValorTLP
                    GravaObrigacaoAcessoria CodObrigacao, Calculo.DadosIptu.CodImpostoLixo, Calculo.DadosIptu.ValorTCL
                    GravaObrigacaoAcessoria CodObrigacao, CodTributo, ValorObrigacao - (Calculo.DadosIptu.ValorTCL + Calculo.DadosIptu.ValorTCVL + Calculo.DadosIptu.ValorTLP)
                    Calculo.GravaObrigacaoAcessoria = False
                End If
            End If
        End If
    End If
    GravaObrigacao = Nvl(CodObrigacao, 0)
    Exit Function
Trata:
    Avisa "GravaObrigacao - " & Err.Description
    Exit Function
    Resume
End Function
Public Function AlteraObrigacaoTOBR201(CodObrigacao As String, DataVence As String, ValorObrigacao As String, _
                ValorMulta As String, ValorJuros As String, ValorTaxas As Double, ValorCorrecao As Double, _
                Status As TipoStatusObrigacao, Optional DescontoConcedido As Double, _
                Optional Periodo As String, Optional Justificativa As String, Optional CodParcelamento As String) As Boolean
                
    Dim Sql As String
    Dim Campos As String
    Dim Valores As String
    
    If IsNull(DescontoConcedido) Then
        DescontoConcedido = 0
    End If
    
    Periodo = IIf(Len(Trim(Periodo)) = 6, Right(Periodo, 4) & Left(Periodo, 2), Periodo)
    Campos = "TOC_DATA_VENCIMENTO,TOC_VALOR_OBRIGACAO," & _
            "TOC_VALOR_MULTA,TOC_VALOR_JUROS,TOC_TOTAL_TAXA_INCLUSA,TOC_STATUS_OBRIGACAO," & _
            " TOC_USUARIO_ALTERACAO,TOC_JUSTIFICATIVA_ALTERACAO,TOC_DESCONTO,TOC_DATA_ALTERACAO,TOC_OBSERVACAO"
    
    Justificativa = "'" & Justificativa & "'"
    Valores = Bdados.PreparaValor(Bdados.Converte(DataVence, TCDataHora), Bdados.Converte(Nvl(ValorObrigacao, 0), TCDuplo), _
                Bdados.Converte(Nvl(ValorMulta, 0), TCDuplo), Bdados.Converte(Nvl(ValorJuros, 0), TCDuplo), _
                Bdados.Converte(ValorTaxas, TCDuplo), Status, AplicacoesVTFuncoes.Usuario, Justificativa, _
                Bdados.Converte(Format(DescontoConcedido, Const_Monetario), TCDuplo), Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora), Justificativa)
    If DescontoConcedido > 0 Then
        Campos = Campos & ",TOC_USUARIO_DESCONTO,TOC_DATA_DESCONTO"
        Valores = Valores & Bdados.PreparaValor(AplicacoesVTFuncoes.Usuario, Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora))
    End If
    If Trim(Periodo) <> "" Then
        Campos = Campos & ",TOC_PERIODO"
        Valores = Valores & Bdados.PreparaValor(Periodo)
    End If
    If Trim(CodParcelamento) <> "" Then
        Campos = Campos & ",TOC_TPA_COD_PARCELAMENTO"
        Valores = Valores & Bdados.PreparaValor(CodParcelamento)
    End If
    
    AlteraObrigacaoTOBR201 = Bdados.AtualizaDados("TAB_OBRIGACAO_CONTRIBUINTE", Valores, Campos, "TOC_COD_OBRIGACAO =" & CodObrigacao)
    If AlteraObrigacaoTOBR201 Then
       AlteraObrigacaoTOBR201 = Conta.CriaContaContribuinte(CodObrigacao)
       'BCP
       AlteraObrigacaoTOBR201 = Conta.CriaContaContribuinteBcp(CodObrigacao, ValorObrigacao, ValorJuros, ValorMulta, ValorCorrecao, DescontoConcedido)
        'FIM BCP
    End If

'    TrocaSitObrigacao CodObrigacao, Status
'    Campos = "TCC_TPA_COD_PARCELAMENTO"
'    Valores = Bdados.PreparaValor(0)
'    Bdados.AtualizaDados "TAB_CONTA_CONTRIBUINTE", Valores, Campos, "TCC_CODIGO_CONTA =" & CodObrigacao
'    GeraDAMObrigacao CodObrigacao
    'GRAVA OS DETALHES DA OBRIGACAO(OBRIGAÇÕES ACESSORIAS)
'    If Calculo.GravaObrigacaoAcessoria Then
'        GravaObrigacaoAcessoria CodObrigacao, Calculo.DadosIptu.CodImpostoConservacao, Calculo.DadosIptu.ValorTCVL
'        GravaObrigacaoAcessoria CodObrigacao, Calculo.DadosIptu.CodImpostoLimpeza, Calculo.DadosIptu.ValorTLP
'        GravaObrigacaoAcessoria CodObrigacao, Calculo.DadosIptu.CodImpostoLixo, Calculo.DadosIptu.ValorTCL
'        GravaObrigacaoAcessoria CodObrigacao, CodTributo, ValorObrigacao - (Calculo.DadosIptu.ValorTCL + Calculo.DadosIptu.ValorTCVL + Calculo.DadosIptu.ValorTLP)
'    End If
End Function

Public Function AlteraObrigacao(CodObrigacao As String, DataVence As String, ValorObrigacao As String, _
                ValorMulta As String, ValorJuros As String, ValorTaxas As Double, _
                Status As TipoStatusObrigacao, Optional DescontoConcedido As Double, _
                Optional Periodo As String, Optional Justificativa As String, Optional CodParcelamento As String) As Boolean
                
    Dim Sql As String
    Dim Campos As String
    Dim Valores As String
    
    If IsNull(DescontoConcedido) Then
        DescontoConcedido = 0
    End If
    
    Periodo = IIf(Len(Trim(Periodo)) = 6, Right(Periodo, 4) & Left(Periodo, 2), Periodo)
    Campos = "TOC_DATA_VENCIMENTO,TOC_VALOR_OBRIGACAO," & _
            "TOC_VALOR_MULTA,TOC_VALOR_JUROS,TOC_TOTAL_TAXA_INCLUSA,TOC_STATUS_OBRIGACAO," & _
            " TOC_USUARIO_ALTERACAO,TOC_JUSTIFICATIVA_ALTERACAO,TOC_DESCONTO,TOC_DATA_ALTERACAO,TOC_OBSERVACAO"
    
    Justificativa = "'" & Justificativa & "'"
    Valores = Bdados.PreparaValor(Bdados.Converte(DataVence, TCDataHora), Bdados.Converte(Nvl(ValorObrigacao, 0), TCDuplo), _
                Bdados.Converte(Nvl(ValorMulta, 0), TCDuplo), Bdados.Converte(Nvl(ValorJuros, 0), TCDuplo), _
                Bdados.Converte(0, TCDuplo), Status, AplicacoesVTFuncoes.Usuario, Justificativa, _
                Bdados.Converte(Format(DescontoConcedido, Const_Monetario), TCDuplo), Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora), Justificativa)
    If DescontoConcedido > 0 Then
        Campos = Campos & ",TOC_USUARIO_DESCONTO,TOC_DATA_DESCONTO"
        Valores = Valores & Bdados.PreparaValor(AplicacoesVTFuncoes.Usuario, Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora))
    End If
    If Trim(Periodo) <> "" Then
        Campos = Campos & ",TOC_PERIODO"
        Valores = Valores & Bdados.PreparaValor(Periodo)
    End If
    If Trim(CodParcelamento) <> "" Then
        Campos = Campos & ",TOC_TPA_COD_PARCELAMENTO"
        Valores = Valores & Bdados.PreparaValor(CodParcelamento)
    End If
    
    AlteraObrigacao = Bdados.AtualizaDados("TAB_OBRIGACAO_CONTRIBUINTE", Valores, Campos, "TOC_COD_OBRIGACAO =" & CodObrigacao)
    If AlteraObrigacao Then
       AlteraObrigacao = Conta.CriaContaContribuinte(CodObrigacao)
       'BCP
       AlteraObrigacao = Conta.CriaContaContribuinteBcp(CodObrigacao, ValorObrigacao, ValorJuros, ValorMulta, ValorTaxas, DescontoConcedido)
        'FIM BCP
    End If

'    TrocaSitObrigacao CodObrigacao, Status
'    Campos = "TCC_TPA_COD_PARCELAMENTO"
'    Valores = Bdados.PreparaValor(0)
'    Bdados.AtualizaDados "TAB_CONTA_CONTRIBUINTE", Valores, Campos, "TCC_CODIGO_CONTA =" & CodObrigacao
'    GeraDAMObrigacao CodObrigacao
    'GRAVA OS DETALHES DA OBRIGACAO(OBRIGAÇÕES ACESSORIAS)
'    If Calculo.GravaObrigacaoAcessoria Then
'        GravaObrigacaoAcessoria CodObrigacao, Calculo.DadosIptu.CodImpostoConservacao, Calculo.DadosIptu.ValorTCVL
'        GravaObrigacaoAcessoria CodObrigacao, Calculo.DadosIptu.CodImpostoLimpeza, Calculo.DadosIptu.ValorTLP
'        GravaObrigacaoAcessoria CodObrigacao, Calculo.DadosIptu.CodImpostoLixo, Calculo.DadosIptu.ValorTCL
'        GravaObrigacaoAcessoria CodObrigacao, CodTributo, ValorObrigacao - (Calculo.DadosIptu.ValorTCL + Calculo.DadosIptu.ValorTCVL + Calculo.DadosIptu.ValorTLP)
'    End If
End Function

Public Function EliminaObrigacao(CodObrigacao As String) As Boolean
                
    Dim Sql As String
    Dim Campos As String
    Dim Valores As String
    
    
    If Bdados.DeletaDados("TAB_OBRIGACAO_CONTRIBUINTE", "TOC_COD_OBRIGACAO ='" & CodObrigacao & "'") Then
        Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "TCC_CODIGO_CONTA =" & CodObrigacao
        EliminaObrigacao = True
    End If

End Function

Public Function EliminaPeriodoObrigacao(Inscricao As String, PeriodoInicial As String) As Boolean
    EliminaPeriodoObrigacao = Bdados.DeletaDados("TAB_PERIODO_OBRIGACAO", "TPO_INSCRICAO ='" & Inscricao & "' AND TPO_PERIODO_INICIAL = " & Bdados.Converte(PeriodoInicial, TCDataHora))
End Function

Private Sub GravaObrigacaoAcessoria(ByVal CodObrigacao As String, ByVal CodTributo As String, ByVal ValorObrigacao As Double)
    Dim Valores As String
    Dim Campos As String
    
    If ValorObrigacao > 0 Then
        Campos = "TDO_TOC_COD_OBRIGACAO,TDO_TIP_COD_IMPOSTO,TDO_VALOR_TRIBUTO"
        Valores = Bdados.PreparaValor(CodObrigacao, CodTributo, Bdados.Converte(ValorObrigacao, TCDuplo))
        Bdados.GravaDados "TAB_DETALHE_OBRIGACAO", Valores, Campos, "TDO_TOC_COD_OBRIGACAO =" & CodObrigacao & _
                        " and TDO_TIP_COD_IMPOSTO ='" & CodTributo & "'"
    End If
End Sub

Public Function BuscaDetalheObrigacao(CodObrigacao As String, Optional Modo As TipoInscricaoObrigacao, Optional ByVal TipoTrans As TipoTransacaoOb = ettAvista) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsAux As VSRecordset
    
    'Funcao que recupera os valores da obrigacao.
        
    BuscaDetalheObrigacao = False
    obContribuinte = ""
    Sql = "Select  * from TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO =" & CodObrigacao
    'If Modo > 0 Then Sql = Sql & " and toc_tipo_inscricao = '" & Modo & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaDetalheObrigacao = True
        obCodigoObrigacao = CodObrigacao
        obDataVencimento = Trim("" & Rs!TOC_DATA_VENCIMENTO)
        obCodImposto = Trim("" & Rs!TOC_TIP_COD_IMPOSTO)
        obPeriodo = Trim("" & Rs!TOC_PERIODO)
        obStatusObrigacao = Nvl("" & Rs!TOC_STATUS_OBRIGACAO, 0)
        obValorObrigacao = Format(Rs!TOC_VALOR_OBRIGACAO, Const_Monetario)
        obValorMulta = Format(Nvl("" & Rs!TOC_VALOR_MULTA, 0), Const_Monetario)
        obValorJuros = Format(Nvl("" & Rs!TOC_VALOR_JUROS, 0), Const_Monetario)
        obValorDesconto = Format(Nvl("" & Rs!TOC_DESCONTO, 0), Const_Monetario)
        
        obTipoInscricao = Nvl("" & Rs!TOC_TIPO_INSCRICAO, 0)
        obNumParcelamento = Nvl("" & Rs!TOC_TPA_COD_PARCELAMENTO, 0)
        obDocumentoOrigem = Nvl("" & Rs!TOC_NUM_DOC_ORIGEM, 0)
        obParcela = Nvl("" & Rs!TOC_PARCELA, 0)
        If Nvl("" & Rs!TOC_TIPO_INSCRICAO, 0) = 1 Then 'TRIBUTO DE IMOVEL
            obInscCadastral = Trim("" & Rs!TOC_INSCRICAO)
            obContribuinte = Trim("" & Rs!TOC_INSCRICAO)
            Sql = "Select tim_tci_im from tab_imovel where tim_ic = '" & Rs!TOC_INSCRICAO & "'" 'BUSCA PROPRIETÁRIO
            If Bdados.AbreTabela(Sql, RsAux) Then
                obInscMunicipal = "" & RsAux!tim_tci_im
            Else
                obInscMunicipal = "0"
            End If
            Bdados.FechaTabela RsAux
        Else
            If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                obContribuinte = Trim("" & Rs!TOC_INSCRICAO)
                obInscMunicipal = Trim("" & Rs!TOC_INSCRICAO)
            Else
                If Len(Trim("" & Trim(Rs!TOC_INSCRICAO))) <> 11 And Left(Right(Trim("" & Trim(Rs!TOC_INSCRICAO)), 3), 1) <> "-" Then
                    Sql = "Select tim_tci_im from tab_imovel where tim_ic = '" & Trim(Rs!TOC_INSCRICAO) & "'" 'BUSCA PROPRIETÁRIO
                    If Bdados.AbreTabela(Sql, RsAux) Then
                        obInscMunicipal = "" & RsAux!tim_tci_im
                    End If
                    Bdados.FechaTabela RsAux
                Else
                    obInscMunicipal = Trim("" & Rs!TOC_INSCRICAO)
                End If
                obContribuinte = Trim("" & Rs!TOC_INSCRICAO)
            End If
        End If
        If obValorDesconto = 0 Then CalculaDesconto TipoTrans
    End If
    
End Function

Public Function BuscaCodObrigacao(Inscricao As String, Periodo As String, CodImposto As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsAux As VSRecordset
        
    obContribuinte = ""
    Sql = "Select TOC_COD_OBRIGACAO from TAB_OBRIGACAO_CONTRIBUINTE WHERE " & _
        " TOC_INSCRICAO ='" & Inscricao & "' AND TOC_PERIODO =" & Periodo & " AND TOC_TIP_COD_IMPOSTO ='" & CodImposto & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaCodObrigacao = "" & Rs!TOC_COD_OBRIGACAO
    End If
End Function

Public Function CarregaPeriodosObrigacao(Lista As Object, Optional Inscricao As String, _
                Optional PeriodoInicial As String, Optional PeriodoFinal As String, _
                Optional CodTributo As String) As Boolean
    Dim Sql As String
    Screen.MousePointer = 11
    CarregaPeriodosObrigacao = False
    Sql = "Select  TPO_INSCRICAO AS Inscricao, " & _
        "TPO_PERIODO_INICIAL as Inicio,TPO_PERIODO_FINAL as Fim  from TAB_PERIODO_OBRIGACAO" & _
        " WHERE TPO_INSCRICAO <> ''"
    If Trim(CodTributo) <> "" Then Sql = Sql & " AND TPO_TIP_COD_IMPOSTO ='" & CodTributo & "'"
    If Trim(Inscricao) <> "" Then Sql = Sql & " and TPO_INSCRICAO = '" & Inscricao & "'"
    If Trim(PeriodoInicial) <> "" Then Sql = Sql & " and TPO_PERIODO_INICIAL >= " & Bdados.Converte(PeriodoInicial, TCDataHora)
    If Trim(PeriodoFinal) <> "" Then Sql = Sql & " and TPO_PERIODO_FINAL >= " & Bdados.Converte(PeriodoFinal, TCDataHora)

    CarregaPeriodosObrigacao = Lista.Preencher(Bdados, Sql, 2000, 1400, 1200, 1200)
    Screen.MousePointer = 0
End Function

'Private Sub DefineTaxasInclusas()
'    Dim i As Byte
'    Dim Sql As String
'    Dim Rs As VSRecordset
'    ReDim TaxaInclusa(1 To 1) As New TaxasInclusas
'    If Imposto.TaxaFixaImpressaoDAM > 0 And i = 1 Then
'        TaxaInclusa(i).Codigo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_TEMISSAODAM))
'        TaxaInclusa(i).Valor = Imposto.TaxaFixaImpressaoDAM
'    End If
'    If Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQNFIXO)) Then
'        ReDim Preserve TaxaInclusa(1 To 4) As New TaxasInclusas
'        'ALVARA
'            TaxaInclusa(i).Codigo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA))
'            TaxaInclusa(i).Valor = Imposto.TaxaFixaImpressaoDAM
'
'        'EMISSAO DE ALVARA
'            Sql = "SELECT TCI_ALVARA_LIBERADO FROM TAB_CONTRIBUINTE WHER TCI_IM = '" & obContribuinte & "'"
'            If Bdados.AbreTabela(Sql, Rs) Then
'                If Nvl("" & Rs!TCI_ALVARA_LIBERADO, 0) = 0 And obTipoInscricao = 1 Then
'                    TaxaInclusa(i).Codigo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_TEMISSAOALVARA))
'                    TaxaInclusa(i).Valor = Imposto.TaxaFixaImpressaoDAM
'                    Bdados.AtualizaDados "TAB_CONTRIBUINTE", Bdados.PreparaValor(1), "TCI_ALVARA_LIBERADO", "TCI_IM = '" & obContribuinte & "'"
'                End If
'            End If
'        'PUBLICIDADE
'            TaxaInclusa(i).Codigo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_TEMISSAOALVARA))
'            TaxaInclusa(i).Valor = Imposto.TaxaFixaImpressaoDAM
'    End If
'
'End Sub

Public Sub GBObrig(Contador As Object)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsAux As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim CodObrig As String
    Dim Inscricao As String
    Dim TipoInsc As Byte

    Dim TipoObrig As Integer
    Dim CodTributo As String
    Dim Cont As Double
    Contador = 0
    Sql = "Select tgt_im,tgt_tim_ic,tgt_tip_cod_imposto,tgt_periodo,tgt_data_vencimento,tgt_data_geracao," & _
                "tgt_valor_tributo,tgt_valor_juros,tgt_valor_multa,tgt_tip_cod_imposto_original,TGT_TAXA_EXPEDIENTE from tab_geracao_tributo "
    'Sql = Sql & " where len(tgt_inscricao) =15"
    Sql = Sql & " order by tgt_inscricao"
     If Bdados.AbreTabela(Sql, Rs, SomenteAvanco) Then
        Rs.MoveFirst
        Campos = "TOC_COD_OBRIGACAO,TOC_INSCRICAO,TOC_TIP_COD_IMPOSTO,TOC_PERIODO,TOC_TIPO_INSCRICAO," & _
                "TOC_DATA_VENCIMENTO,TOC_DATA_GERACAO,TOC_VALOR_OBRIGACAO," & _
                "TOC_VALOR_MULTA,TOC_VALOR_JUROS,TOC_STATUS_OBRIGACAO,TOC_TIPO_LANCAMENTO,TOC_TOTAL_TAXA_INCLUSA"
'            Campos = "TDO_TOC_COD_OBRIGACAO,TDO_TIP_COD_IMPOSTO,TDO_VALOR_TRIBUTO"

        Do
            If Trim("" & Rs!tgt_tip_cod_imposto) <> "11120200" Then
                TipoInsc = 2
                Inscricao = "" & Rs!tgt_im
            Else
                TipoInsc = 1
                Inscricao = "" & Trim(Rs!tgt_tim_ic)
            End If
'            If Rs!tgt_tip_cod_imposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_DATIVA)) Then
'                TipoObrig = 5
'                CodTributo = IIf(Trim("" & Rs!tgt_tip_cod_imposto_original) = "", Rs!tgt_tip_cod_imposto, "" & Rs!tgt_tip_cod_imposto_original)
'            Else
                TipoObrig = 2
                CodTributo = "" & Rs!tgt_tip_cod_imposto
'            End If
            Sql = "Select TOC_COD_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_INSCRICAO in ('" & Inscricao & _
                "','" & Rs!tgt_im & "') AND  TOC_PERIODO =" & Rs!tgt_periodo & " and TOC_TIPO_INSCRICAO = " & TipoInsc 'AND TOC_TIP_COD_IMPOSTO ='" & Rs!tgt_tip_cod_imposto & "'
            If Not Bdados.AbreTabela(Sql, RsAux, SomenteAvanco) Then
            '        CodObrig = RsAux!TOC_COD_OBRIGACAO
                    'CodObrig = Bdados.Correlativo("TRIB", Const_Obrig, "OBRIGACAO TRIBUTARIA")
            'Else
                    CodObrig = Conta.GeraCodPagamento(Const_Obrig)

                    Valores = Bdados.PreparaValor(CodObrig, Bdados.Converte(Inscricao, tctexto), CodTributo, Rs!tgt_periodo, TipoInsc, Bdados.FormataValorCampo(Rs!tgt_data_vencimento), _
                                "" & Rs!tgt_data_geracao, Rs!tgt_valor_tributo + CDbl(Nvl("" & Rs!tgt_taxa_expediente, 0)), Nvl("" & Rs!tgt_Valor_multa, 0), Nvl("" & Rs!tgt_valor_juros, 0), TipoObrig, 1, CDbl(Nvl("" & Rs!tgt_taxa_expediente, 0)))
                    'Valores = Bdados.PreparaValor(CodObrig, CodTributo, Rs!tgt_valor_tributo)
                    Bdados.GravaDados "TAB_OBRIGACAO_CONTRIBUINTE", Valores, Campos, "TOC_COD_OBRIGACAO=" & CodObrig
            End If
            Cont = Cont + 1
            Debug.Print Cont
            Rs.MoveNext
            DoEvents
        Loop While Not Rs.EOF
     End If
     Screen.MousePointer = 0
     Avisa "Processo finalizado."
End Sub


Public Sub GBDarm()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsAux As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim CodObrig As String
    Dim Inscricao As String
    Dim TipoInsc As Byte

    Dim TipoObrig As Integer
    Dim CodTributo As String
    Screen.MousePointer = 11
    Sql = "Select TDR_INSCRICAO,tdr_tim_ic,TDR_TGT_COD_PAGAMENTO,tdr_im FROM TAB_DARM_RECEBIDO  where tdr_inscricao is null"
    Sql = Sql & " order by TDR_INSCRICAO"
     If Bdados.AbreTabela(Sql, Rs, SomenteAvanco) Then
        Rs.MoveFirst
        Campos = "TDR_INSCRICAO"

        Do
            If Trim("" & Rs!TDR_TIM_IC) = "" Then
                Inscricao = Trim("" & Rs!tdr_im)
            Else
                Inscricao = "" & Trim(Rs!TDR_TIM_IC)
            End If
            Valores = Bdados.PreparaValor(Bdados.Converte(Inscricao, tctexto))
            Bdados.AtualizaDados "TAB_DARM_RECEBIDO", Valores, Campos, " TDR_TGT_COD_PAGAMENTO=" & Rs!TDR_TGT_COD_PAGAMENTO
            Rs.MoveNext
            DoEvents
        Loop While Not Rs.EOF
     End If
     Screen.MousePointer = 0
     Avisa "Processo finalizado."
End Sub

Public Sub GBDam()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsAux As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim CodObrig As String
    Dim Inscricao As String
    Dim TipoInsc As Byte

    Dim TipoObrig As Integer
    Dim CodTributo As String
    Screen.MousePointer = 11
    Sql = "Select tgt_im,tgt_tim_ic,TGT_COD_PAGAMENTO FROM TAB_GERACAO_TRIBUTO "
    Sql = Sql & " order by tgt_im"
     If Bdados.AbreTabela(Sql, Rs, SomenteAvanco) Then
        Rs.MoveFirst
        Campos = "TGT_INSCRICAO"
        Do
            If Trim("" & Rs!tgt_tim_ic) = "" Then
                Inscricao = Trim("" & Rs!tgt_im)
            Else
                Inscricao = "" & Trim(Rs!tgt_tim_ic)
            End If
            Valores = Bdados.PreparaValor(Bdados.Converte(Inscricao, tctexto))
            Bdados.AtualizaDados "TAB_GERACAO_TRIBUTO", Valores, Campos, " TGT_COD_PAGAMENTO=" & Rs!TGT_COD_PAGAMENTO
            Rs.MoveNext
            DoEvents
        Loop While Not Rs.EOF
     End If
     Screen.MousePointer = 0
     Avisa "Processo finalizado."
End Sub

Private Sub CalculaDesconto(Optional TipoTransacao As TipoTransacaoOb = ettAvista)
    On Error Resume Next
    Dim PercDesconto As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DataVencimento As String
    Dim Parcela As Integer
    Dim DescAteVencimento As Integer
    
    Sql = "Select TOC_DESCONTO,TOC_DATA_VENCIMENTO,TOC_PARCELA FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO = " & obCodigoObrigacao
    If Bdados.AbreTabela(Sql, Rs) Then
        obValorDesconto = Nvl("" & Rs!TOC_DESCONTO, 0)
        DataVencimento = "" & Rs!TOC_DATA_VENCIMENTO
        Parcela = "" & Rs!TOC_PARCELA
        If obValorDesconto > 0 Then
            Exit Sub
        End If
    End If
    If ((AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE") And obPeriodo = 2004 And obCodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))) And obPeriodo = 2004 Then
        If Parcela = 0 Then
            Sql = "Select TOC_COD_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_INSCRICAO ='" & obContribuinte & "' AND TOC_STATUS_OBRIGACAO not in (3,6,7,8) AND TOC_TIP_COD_IMPOSTO ='" & obCodImposto & "' AND TOC_PERIODO <> 2004"
            If Not Bdados.AbreTabela(Sql, Rs) Then
                obValorDesconto = 50
                Exit Sub
            End If
        End If
    End If
    Sql = "Select TPI_DESCONTO_TRIBUTO_PARC,TPI_DESCONTO_TRIBUTO,TPI_DESCONTO_ATE_VENCIMENTO,TPI_PARAMETRO_REFIS from Tab_Parametro_Imposto " & _
        " where tpi_tip_cod_imposto='" & obCodImposto & "'"
    Sql = Sql & " and tpi_ano_imposto= '" & Imposto.BuscaAnoImposto(obCodImposto, Left(obPeriodo, 4)) & "'"
   If Bdados.AbreTabela(Sql, Rs) Then
        DescAteVencimento = Nvl("" & Rs!TPI_DESCONTO_ATE_VENCIMENTO, 2)
        PercDesconto = Nvl("" & Rs!TPI_DESCONTO_TRIBUTO, 0)
        If Nvl("" & Rs!TPI_PARAMETRO_REFIS, 1) = 2 Then
            Sql = "select TPI_DESCONTO_TRIBUTO_PARC,TPI_DESCONTO_JUROS_PARC,TPI_DESCONTO_MULTA_PARC,TPI_DESCONTO_ATUALIZACAO_PARC "
            Sql = Sql & " from TAB_PARAMETRO_IMPOSTO_REFIS where TPI_TIP_COD_IMPOSTO ='" & obCodImposto & "'"
            Sql = Sql & " and tpi_ano_imposto= '" & Imposto.BuscaAnoImposto(obCodImposto, Left(obPeriodo, 4)) & "'"
            Sql = Sql & " and " & Parcela & " BETWEEN TPI_LIMITE_INFERIOR AND TPI_LIMITE_SUPERIOR "
            If Not Bdados.AbreTabela(Sql, Rs) Then
                Exit Sub
            Else
                PercDesconto = Nvl("" & Rs!TPI_DESCONTO_TRIBUTO_PARC, 0)
            End If
        Else
'            ValJurosDescVista = Nvl("" & Rs!TPI_DESCONTO_JUROS, 0)
        End If
    End If
    If Not Rs.EOF Then
        If TipoTransacao = ettAvista And Parcela = 0 Then
            If DescAteVencimento = 1 Then 'DESCONTO ATE VENCIMENTO
                If DateDiff("D", Date, DataVencimento) >= 0 Then
                    obValorDesconto = PercDesconto 'Format(obValorObrigacao * CDbl(Nvl("" & Rs!tpi_desconto, 0) / 100), Const_Monetario)
                End If
            Else 'INDISCRIMINADO
                obValorDesconto = PercDesconto 'Format(obValorObrigacao * CDbl(Nvl("" & Rs!tpi_desconto, 0) / 100), Const_Monetario)
            End If
        Else
            obValorDesconto = Nvl("" & Rs!TPI_DESCONTO_TRIBUTO_PARC, 0)
        End If
    End If
End Sub
Public Function Grava_Log_Obrigacao(Obrigacao As String, Tipo As Tipo_Log_Obrigacao, Optional Usuario As String, Optional Motivo As String) As Boolean
'    Dim Sql As String
'    Dim Rs As VSRecordset
'    Dim Valores As String
'    Dim Campos As String
'    Sql = "select TOC_COD_OBRIGACAO,TOC_INSCRICAO,TOC_TIP_COD_IMPOSTO,"
'    Sql = Sql & " TOC_PERIODO,TOC_INDICE,TOC_TIPO_INSCRICAO,"
'    Sql = Sql & " TOC_NUM_DOC_ORIGEM,TOC_TPA_COD_PARCELAMENTO,TOC_DATA_VENCIMENTO,"
'    Sql = Sql & " TOC_DATA_GERACAO,TOC_VALOR_OBRIGACAO,TOC_VALOR_MULTA,"
'    Sql = Sql & " TOC_VALOR_JUROS,TOC_CORRECAO_MONETARIA,TOC_TIPO_LANCAMENTO,"
'    Sql = Sql & " TOC_STATUS_OBRIGACAO,TOC_STATUS_ANTERIOR_OBRIGACAO,"
'    Sql = Sql & " TOC_TOTAL_TAXA_INCLUSA,TOC_OBSERVACAO,TOC_PARCELA,"
'    Sql = Sql & " TOC_VALOR_PAGO_PARCELADO, TOC_VALOR_OBRIGACAO_ORIGINAL,"
'    Sql = Sql & " TOC_SITUACAO_TAXA,TOC_DESCONTO,TOC_USUARIO_DESCONTO,"
'    Sql = Sql & " TOC_DATA_DESCONTO,TOC_JUSTIFICATIVA_ALTERACAO"
'    Sql = Sql & " from tab_obrigacao_contribuinte where TOC_COD_OBRIGACAO = " & Obrigacao
'    If Bdados.AbreTabela(Sql, Rs) Then
'        Campos = "TOC_COD_OBRIGACAO,TOC_INSCRICAO,TOC_TIP_COD_IMPOSTO,TOC_ACAO,TOC_PERIODO,"
'        Campos = Campos & " TOC_INDICE,TOC_TIPO_INSCRICAO,TOC_NUM_DOC_ORIGEM,TOC_TPA_COD_PARCELAMENTO,"
'        Campos = Campos & " TOC_DATA_VENCIMENTO,TOC_DATA_GERACAO,TOC_VALOR_OBRIGACAO,TOC_VALOR_MULTA,"
'        Campos = Campos & " TOC_VALOR_JUROS,TOC_CORRECAO_MONETARIA,TOC_TIPO_LANCAMENTO,TOC_STATUS_OBRIGACAO,"
'        Campos = Campos & " TOC_STATUS_ANTERIOR_OBRIGACAO,TOC_TOTAL_TAXA_INCLUSA,TOC_OBSERVACAO,"
'        Campos = Campos & " TOC_PARCELA,TOC_VALOR_PAGO_PARCELADO,TOC_VALOR_OBRIGACAO_ORIGINAL,TOC_SITUACAO_TAXA,"
'        Campos = Campos & " TOC_DESCONTO , TOC_USUARIO_DESCONTO, TOC_DATA_DESCONTO, TOC_JUSTIFICATIVA_ALTERACAO"
'
'
'        Valores = Bdados.PreparaValor(Rs.Fields("TOC_COD_OBRIGACAO"), Trim(Rs.Fields("TOC_INSCRICAO")), "" & Rs.Fields("TOC_TIP_COD_IMPOSTO"), _
'        Tipo, Rs.Fields("TOC_PERIODO"), _
'        "" & Rs.Fields("TOC_INDICE"), "" & Rs.Fields("TOC_TIPO_INSCRICAO"), _
'        "" & Rs.Fields("TOC_NUM_DOC_ORIGEM"), Nvl("" & Rs.Fields("TOC_TPA_COD_PARCELAMENTO"), 0), _
'        "" & Rs.Fields("TOC_DATA_VENCIMENTO"), "" & Rs.Fields("TOC_DATA_GERACAO"), _
'        "" & Rs.Fields("TOC_VALOR_OBRIGACAO"), "" & Rs.Fields("TOC_VALOR_MULTA"), _
'        "" & Rs.Fields("TOC_VALOR_JUROS"), Nvl("" & Rs.Fields("TOC_CORRECAO_MONETARIA"), 0), _
'        "" & Rs.Fields("TOC_TIPO_LANCAMENTO"), "" & Rs.Fields("TOC_STATUS_OBRIGACAO"), _
'        "" & Rs.Fields("TOC_STATUS_ANTERIOR_OBRIGACAO"), _
'        "" & Rs.Fields("TOC_TOTAL_TAXA_INCLUSA"), Motivo, _
'        "" & Rs.Fields("TOC_PARCELA"), Nvl("" & Rs.Fields("TOC_VALOR_PAGO_PARCELADO"), 0), _
'        "" & Rs.Fields("TOC_VALOR_OBRIGACAO_ORIGINAL"), _
'        "" & Rs.Fields("TOC_SITUACAO_TAXA"), Nvl("" & Rs.Fields("TOC_DESCONTO"), 0), _
'        Usuario, Nvl("" & Rs.Fields("TOC_DATA_DESCONTO"), 0), _
'        Nvl("" & Rs.Fields("TOC_JUSTIFICATIVA_ALTERACAO"), 0))
'        Grava_Log_Obrigacao = Bdados.InsereDados("tab_log_obrigacao", Valores, Campos)
        
    'End If
    Grava_Log_Obrigacao = True
End Function

Public Function SenhaValidaGerente(Senha As String) As Boolean
    On Error GoTo Trata
    Dim Rs As Object
    Dim Sql As String
    Dim Seguranca As New VSSeguranca
    Sql = "SELECT TUS_SENHA FROM TAB_USUARIO WHERE TUS_COD_USUARIO = '" & Temp.PegaParametro(Bdados, "GERENTE LIBERACAO") & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Rs(0) = Seguranca.Criptografa(Senha) Then
            SenhaValidaGerente = True
            Exit Function
        End If
    End If
    Bdados.FechaTabela Rs
    SenhaValidaGerente = False
    
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Function

Public Function LiberaContribuinteNotificado(Contribuinte As String) As Boolean
    Dim Sql As String
    
    Dim Formulario As Object
    
    Set Formulario = Forms.Add("TOBR409")
    LiberaContribuinteNotificado = True
    If Bdados.AbreTabela("SELECT MAX(TNT_ENTREGA) as data,COUNT(*) as Total FROM TAB_NOTIFICACAO " & _
            " WHERE TNT_INSCRICAO = '" & Contribuinte & "'AND NOT TNT_ENTREGA IS NULL and TNT_TIPO = 1") Then
        If IsNull(Bdados.Tabela(0)) Then
            If Not Bdados.AbreTabela("select * from tab_divida_ativa where tda_tci_im = '" & Contribuinte & "'") Then
                Exit Function
            End If
        End If
        Formulario.Show 1
        If Trim(SenhaLiberacao) = "" Then
            LiberaContribuinteNotificado = False
        Else
            If Not SenhaValidaGerente(SenhaLiberacao) Then
                Avisa "Senha inválida."
                LiberaContribuinteNotificado = False
            Else
                LiberaContribuinteNotificado = True
            End If
        End If
    End If
End Function
Public Function Valida_Geracao_Triuto(Tributo As String, Ano As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    
    
    'LEGENDA
    'SIM - SE A FUNÇÃO RETORNAR TRUE(SIM), ENTÃO SERÁ GERADA UMA OBRIGAÇÃO COM O VALOR ZERO...
    'NÃO - SE A FUNÇÃO RETORNAR FALSE(NAO), ENTÃO NÃO SERÁ GERADA UMA OBRIGAÇÃO COM O VALOR ZERO...
    
    Sql = "Select TPI_GERAR_OBRIGACAO_ZERADA "
    Sql = Sql & "  from tab_parametro_imposto"
    Sql = Sql & " where TPI_TIP_COD_IMPOSTO = '" & Tributo & "'"
    Sql = Sql & "  AND TPI_ANO_IMPOSTO = '" & Left(Ano, 4) & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Rs.Fields("TPI_GERAR_OBRIGACAO_ZERADA") = 1 Then 'SIM
            Valida_Geracao_Triuto = True
        ElseIf Rs.Fields("TPI_GERAR_OBRIGACAO_ZERADA") = 2 Then 'NÃO
            Valida_Geracao_Triuto = False
        End If
    Else
        'SE NÃO TIVER NADA DEFINIDO NA TAB_PARAMETRO_INPOSTO, ENTÃO COLOCO COMO NÃO
        Valida_Geracao_Triuto = False 'NÃO
    End If
End Function





