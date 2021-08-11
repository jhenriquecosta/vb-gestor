VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private VsImp As New VSImposto

Public Enum TipoItem
    EtdSaldoAtual = 0
    EtdSaldoDevedor
    EtdSaldoCredor
    EtdMulta
    EtdMultaAcumulada
    EtdJuros
    EtdImpostoOriginal
    EtdNaoTributo
    EtdDesconto
    EtdImpostoAtual
    EtdTotalPago
    
End Enum
Public Enum TipoTransacao
    ettAvista = 1
    ettParcelada = 2
    ettDividaAtiva = 3
End Enum

Public Enum Tipo_Geracao
    etgDAM = 1
    etgIPTU = 2
End Enum

Public Enum TipoOperacaoConta
    EtoDebito = 0
    EtoCredito = 1
    EtoGeracao = 2
End Enum

Public Enum TipoConta
    EtcCreditoTributario = 1
    EtcAutoInfracao
    EtcParcelamento
    EtcDividaAtiva
    EtcDossie
End Enum

Public Enum TipoContaTransacao
    EttPagamento = 1
    EttCreditoAnterior = 2
    EttCreditoParcelamento = 9
    EttDebitoMesSequente = 52
    EttDebitoConta = 53
    EttMovimentoDebito = 54
    EttDebitoMulta = 55
    EttDebitoJuros = 56
End Enum

Public Enum TipoCorrelativo
    EtsFinanciamento = 55
    EtsCotaParcelamento = 57
    EtsIptu = 73
    EtsDossieAnterior = 72
    EtsIss = 81
    EtsIssRetido = 82
    EtsAlvara = 91
    EtsTaxas = 93
    EtsDividaAtiva = 66
    EtsExtratoPagamento = 99
    EtsNotificacao = 97
    etsEconomico = 98
    EtsProtocolo = 53
End Enum
Public Enum TipoStatusConta
    etaAberta = 2
    etaFechada = 3
    etaEmParcelamento = 4
    etaDividaAtiva = 5
    
End Enum

Public Enum Tipo_Inscricao
    Fisica = 1
    Juridica = 2
End Enum
Public Enum TipoDebitar
    EtdeImposto = 2
    EtdeMulta
    EtdeJuros
    EtdeCorrecao
    EtdeDescontos
End Enum

Private CorrecaoConta As Double
Private Obrig As New Obrigacao
Private JurosConta As Double
Private MultaConta As Double
Public Sub ApagaContasAnteriores(Contribuinte As String, Modo As TipoInscricaoObrigacao, Optional CodObrigacao As String, Optional CodImposto As String)
        Dim Condicao As String
        Dim Sql As String
        If Trim(Contribuinte) <> "" Then
            Condicao = " TCC_INSCRICAO = '" & Contribuinte & "'"
            Condicao = Condicao & " and tcc_tipo_inscricao = '" & Modo & "'"
            If Trim(CodImposto) <> "" Then Condicao = Condicao & " and tcc_tip_cod_imposto ='" & CodImposto & "'"
        End If
        If Trim(CodObrigacao) <> "" Then Condicao = "tcc_codigo_conta =" & CodObrigacao
        Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", Condicao
End Sub

Public Sub ExecutaAtualizacao(Contribuinte As String, Optional Modo As TipoInscricaoObrigacao, _
            Optional TrazSubDividas As Boolean = True, Optional CodObrigacao As String, _
            Optional TipoTransacao As TipoTransacao = ettAvista, Optional DataMovimento As String, _
            Optional DescontoIndividualTributo As Double = 0, Optional DescontoIndividualAtualizacao As Double = 0, _
            Optional DescontoIndividualMulta As Double = 0, Optional DescontoIndividualJuros As Double = 0, _
            Optional CodImposto As String, Optional Status As TipoStatusListagem = etlNaoPagos, _
            Optional PeriodoInicial As String, Optional PeriodoFinal As String)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Inscricoes As String
    Dim SqlAux As String
    Dim CondicaoContaCorrente As String
    
'    Exit Sub
    'Motivo :
    'Sub rotina estava altualizando o status das obrigações pra crédito em aberto
    Screen.MousePointer = 11
   
'    If Trim(Contribuinte) = "" Then Exit Sub
    Sql = "SELECT TOC_COD_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE "
    If Status = etlTodos Then
        Sql = Sql & " WHERE 1 = 1 "
    Else
        If Status = etlParcelaveis Then
            Sql = Sql & " where toc_status_obrigacao not in(" & Const_NaoParcelaveis & ") "
            Sql = Sql & " and toc_data_vencimento < " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
        End If
        If Status = etlNaoPagosNaoVencidos Then
            Sql = Sql & " where toc_status_obrigacao in (" & Const_NaoPagos & ")"
            Sql = Sql & " and toc_data_vencimento > " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
        End If
        If Status = etlNaoPagosVencidos Then
            Sql = Sql & " where toc_status_obrigacao in (" & Const_NaoPagos & ")"
            Sql = Sql & " and toc_data_vencimento < " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
        End If
        If Status = etlPagos Then Sql = Sql & " where toc_status_obrigacao = " & etsCreditoPago
        If Status = etlNaoPagos Then Sql = Sql & " where toc_status_obrigacao in (" & Const_NaoPagos & ")"
    End If
    If Modo = etiContribuinte Then
        If Trim(Contribuinte) <> "" Then
            If TrazSubDividas Then
                SqlAux = "select tim_ic from tab_imovel where  tim_tci_im = '" & Contribuinte & "'"
                If Bdados.AbreTabela(SqlAux, Rs) Then
                    Rs.MoveFirst
                    Inscricoes = ""
                    Do While Not Rs.EOF
                        Inscricoes = Inscricoes & "'" & Trim(Rs!tim_ic) & "',"
                        Rs.MoveNext
                    Loop
                    Inscricoes = Left(Inscricoes, Len(Inscricoes) - 1)
                    Sql = Sql & " AND (TOC_INSCRICAO = '" & Trim(Contribuinte) & "' and toc_tipo_inscricao = 2) or (toc_inscricao in (  " & Inscricoes & ") and toc_tipo_inscricao =1)"
                Else
                    Sql = Sql & " AND (TOC_INSCRICAO ='" & Trim(Contribuinte) & "' AND toc_tipo_inscricao = 2)"
                End If
                If Trim(CodImposto) <> "" Then Sql = Sql & " AND TOC_TIP_COD_IMPOSTO ='" & CodImposto & "'"
            Else
                Sql = Sql & " AND TOC_INSCRICAO ='" & Trim(Contribuinte) & "'"
                If Trim(CodImposto) <> "" Then Sql = Sql & " AND TOC_TIP_COD_IMPOSTO ='" & CodImposto & "'"
            End If
        End If
    Else
        If CodObrigacao <> "" Then
            Sql = Sql & " AND TOC_COD_OBRIGACAO ='" & CodObrigacao & "'"
            If Trim(CodImposto) <> "" Then Sql = Sql & " AND TOC_TIP_COD_IMPOSTO ='" & CodImposto & "'"
        ElseIf Trim(Contribuinte) <> "" Then
            Sql = Sql & " AND TOC_INSCRICAO ='" & Contribuinte & "'"
            Sql = Sql & " and toc_tipo_inscricao = " & Modo
        End If
        If Trim(CodImposto) <> "" Then Sql = Sql & " AND TOC_TIP_COD_IMPOSTO ='" & CodImposto & "'"
    End If
    If Trim(PeriodoInicial) <> "" And Trim(PeriodoFinal) <> "" Then
        Sql = Sql & " and toc_periodo between " & PeriodoInicial & " and " & PeriodoFinal
    End If
    If Trim(Contribuinte) <> "" Then
        CondicaoContaCorrente = "TCC_INSCRICAO = '" & Contribuinte & "'"
        If Trim(CodImposto) <> "" Then
            CondicaoContaCorrente = CondicaoContaCorrente & " AND TCC_TIP_COD_IMPOSTO = '" & CodImposto & "'"
        End If
        Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", CondicaoContaCorrente
    End If
    Sql = Sql & " ORDER BY TOC_INSCRICAO,TOC_TIP_COD_IMPOSTO,TOC_PERIODO"
    If Bdados.AbreTabela(Sql, Rs, Estatico, SomenteLeitura) Then
'        CodObrigacao = Rs!TOC_COD_OBRIGACAO
'        Conta.ApagaContasAnteriores Contribuinte, Modo, CodObrigacao, CodImposto
'        Conta.PendenciarPagamentos Contribuinte, Modo, CodObrigacao, CodImposto
        Rs.MoveFirst
        Do While Not Rs.EOF
            Obrig.BuscaDetalheObrigacao Rs!TOC_COD_OBRIGACAO, Modo, TipoTransacao
            If Conta.CriaContaContribuinte(Nvl("" & Rs!TOC_COD_OBRIGACAO, 0), , Obrig, Modo) Then
                Conta.MovimentaContaContribuinte Nvl("" & Rs!TOC_COD_OBRIGACAO, 0), DataMovimento, Obrig, _
                    TipoTransacao, DescontoIndividualTributo, DescontoIndividualAtualizacao, DescontoIndividualMulta, _
                    DescontoIndividualJuros
                Conta.BaixaPagamentos Obrig.obContribuinte, Obrig.obPeriodo, Obrig.obCodImposto, 0, DataMovimento, "" & Rs!TOC_COD_OBRIGACAO, Obrig, Modo, Obrig.obParcela
            End If
            DoEvents
            Rs.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
End Sub

Public Function CalculaValoresJuros(ByVal CodConta As String, ByVal DataPagamento As String, _
        ByVal DataVencimento As String, Optional Obrig As Obrigacao, Optional TipoTransacao As _
        TipoTransacao = ettAvista, Optional ResultadoDesconto As Double) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DiasPassados As Double
    Dim ValJuros As Double
    Dim ValJurosDescVista As Double
    Dim ValJurosDescParc As Double
    Dim MultaMaxima As Double
    Dim Condicao As String
    Dim Valores As String
    Dim DebitoRestante  As Double
    Dim Pag As String
    Dim Venc As String
    Dim Cotas As Integer
    Dim ValorCota As Double
    Dim MesesPassados As Double
    Dim ClsImposto As New VSImposto
    Dim RsAux As VSRecordset
       
    'Busca no Banco
    Dim Periodo As Double
    Dim Imposto As String
    Dim Tipo As TipoConta
    Dim Parcela As Double
    Dim TipoJuros As Integer
    
    On Error GoTo Trata
    CalculaValoresJuros = True
    
    Sql = "Select tcc_periodo, tcc_tip_cod_imposto,tcc_parcela,tcc_num_conta,tcc_saldo_atual,tcc_imposto_original,tcc_correcao_monetaria from tab_conta_contribuinte where tcc_codigo_conta =" & CodConta
    If Bdados.AbreTabela(Sql, RsAux) Then
        Periodo = RsAux!tcc_periodo
        Imposto = RsAux!tcc_tip_cod_imposto
        Parcela = RsAux!tcc_parcela
        Tipo = RsAux!tcc_num_conta
    Else
        ResultadoDesconto = 0
        Exit Function
    End If
    Sql = "SELECT tpi_valor_juros,tpi_dias_pagar,tpi_juros_captalizados,TPI_DESCONTO_JUROS," & _
        "TPI_DESCONTO_JUROS_PARC,TPI_PARAMETRO_REFIS  FROM tab_parametro_imposto" & _
        " where tpi_tip_cod_imposto ='" & Imposto & "'"
    Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Periodo, 4)) & "'"
    If Bdados.AbreTabela(Sql & Condicao, Rs) Then
        TipoJuros = Nvl("" & Rs!tpi_juros_captalizados, 0)
        ValJuros = Nvl("" & Rs!tpi_valor_juros, 0)
        If Nvl("" & Rs!TPI_PARAMETRO_REFIS, 1) = 2 Then
            Sql = "select TPI_DESCONTO_TRIBUTO_PARC,TPI_DESCONTO_JUROS_PARC,TPI_DESCONTO_MULTA_PARC,TPI_DESCONTO_ATUALIZACAO_PARC "
            Sql = Sql & " from TAB_PARAMETRO_IMPOSTO_REFIS where TPI_TIP_COD_IMPOSTO ='" & Imposto & "'"
            Sql = Sql & " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Periodo, 4)) & "'"
            Sql = Sql & " and " & Parcela & " BETWEEN TPI_LIMITE_INFERIOR AND TPI_LIMITE_SUPERIOR "
            If Not Bdados.AbreTabela(Sql, Rs) Then
                Exit Function
            End If
        Else
            ValJurosDescVista = Nvl("" & Rs!TPI_DESCONTO_JUROS, 0)
        End If
    End If
    If Not Rs.EOF Then
        Pag = IIf(Trim(DataPagamento) = "", Format(Date, "dd/mm/yyyy"), DataPagamento)
        Pag = Right(Pag, 4) & Mid(Pag, 4, 2) & Left(Pag, 2)
        Venc = Right(DataVencimento, 4) & Mid(DataVencimento, 4, 2) & Left(DataVencimento, 2)
        If Pag > Venc Then
            DebitoRestante = Format(Nvl("" & RsAux!tcc_saldo_atual, 0), Const_Monetario)
            'If DebitoRestante > 0 Then 'SQz - Balsas 07032003
            If StatusConta(CodConta) <> etaFechada Then
                If IsDate(DataVencimento) And IsDate(DataPagamento) Then
                    DiasPassados = DateDiff("d", DataVencimento, IIf(Trim(DataPagamento) = "", Format(Date, "dd/mm/yyyy"), DataPagamento))
                Else
                    Exit Function
                End If
                If DiasPassados > 0 Then
                    MesesPassados = DateDiff("m", DataVencimento, IIf(Trim(DataPagamento) = "", Format(Date, "dd/mm/yyyy"), DataPagamento))
                    If MesesPassados > 0 Or Day(DataVencimento) < Day(Nvl(DataPagamento, DataVencimento)) Then
                        If Day(DataVencimento) < Day(Nvl(DataPagamento, DataVencimento)) Then
                            MesesPassados = MesesPassados + 1
                        End If
                    End If
'                    ValJuros = Nvl("" & Rs!tpi_valor_juros, 0)
                    If Parcela = 0 Then ValJurosDescVista = Nvl("" & Rs!TPI_DESCONTO_JUROS_PARC, 0)
                    ValJurosDescParc = Nvl("" & Rs!TPI_DESCONTO_JUROS_PARC, 0)
                    
                    If TipoJuros = 1 Then 'JUROS SIMPLES
                        ValJuros = (ValJuros * MesesPassados) / 100
                    Else 'JUROS CAPITALIZADOS
                        ValJuros = ValJuros / 100
                        ValJuros = (CDbl(1 + ValJuros) ^ MesesPassados)
                        ValJuros = Format(ValJuros - 1, "0.000000000")
                    End If
                End If
                DebitoRestante = Format(CDbl(Nvl("" & RsAux!tcc_imposto_original, 0)) + CDbl(Nvl(RsAux!tcc_correcao_monetaria, 0)), Const_Monetario)
                ValJuros = Format(ValJuros * DebitoRestante, Const_Monetario)
                If ResultadoDesconto = 0 Then
                    If TipoTransacao <> ettDividaAtiva Then
                        If TipoTransacao = ettAvista Then
                            ResultadoDesconto = ValJuros * (ValJurosDescVista / 100)
                        Else
                            ResultadoDesconto = ValJuros * (ValJurosDescParc / 100)
                        End If
                    Else
                        ResultadoDesconto = 0
                    End If
                Else
                    ResultadoDesconto = Format(ValJuros * (ResultadoDesconto / 100), Const_Monetario)
                End If
                JurosConta = ValJuros
                ResultadoDesconto = Format(ResultadoDesconto, Const_Monetario)
                Call AtualizaContaContribuinte(CodConta, CDate(DataPagamento), ValJuros, Tipo, EtoDebito, EtdeJuros)
            End If
        Else
            ResultadoDesconto = 0
        End If
    End If
    Bdados.FechaTabela Rs
    Exit Function
Trata:
        Avisa Err.Description
        CalculaValoresJuros = False
        Exit Function
        Resume
End Function

Public Function GeraCodPagamento(Imposto As String, Optional Ano_A As Integer = -1) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Correlativo As String
    Dim NomeImposto As String
    If Len(Imposto) > 2 Then
        Sql = "Select tip_cod_correlativo,tip_sigla_imposto from tab_imposto where tip_cod_imposto = '" & Imposto & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            Correlativo = Trim("" & Rs(0))
            NomeImposto = Rs!TIP_SIGLA_IMPOSTO
        End If
        Rs.Fechar
    Else
        Correlativo = Imposto
    End If
    Sql = "UPDATE Tab_Num_Correlativo set tnc_sequencia =  tnc_sequencia + 1 "
    If Trim(Correlativo) = "" Then
        Sql = Sql + " Where tnc_descricao = '" + Imposto + "'"
    Else
        Sql = Sql + " where tnc_tipo_operacao = " & IIf(Len(Correlativo) = 2, Correlativo, 0)
    End If
    If Ano_A >= 0 Then Sql = Sql + " and tnc_ano =" & Ano_A
    Bdados.Executa Sql
     Sql = "SELECT tnc_sequencia from tab_num_correlativo where " _
        & " tnc_tipo_operacao = " & IIf(Len(Correlativo) = 2, Correlativo, 0)
    Sql = "SELECT tnc_sequencia from tab_num_correlativo "
    If Trim(Correlativo) = "" Then
        Sql = Sql + " Where tnc_descricao = '" + Imposto + "'"
    Else
        Sql = Sql + " where tnc_tipo_operacao = " & IIf(Len(Correlativo) = 2, Correlativo, 0)
    End If
    If Ano_A >= 0 Then Sql = Sql + " and tnc_ano =" & Ano_A
    If Bdados.AbreTabela(Sql, Rs) Then
        GeraCodPagamento = Trim(Correlativo) & Format(IIf(IsNull(Rs(0)), 1, Rs(0)), "000000")
    Else
        Sql = "INSERT INTO Tab_Num_Correlativo(tnc_ano,tnc_tipo_operacao,tnc_sequencia,tnc_descricao) VALUES( " & Right(Date, 1) & "," & IIf(Len(Correlativo) = 2, Correlativo, 0) & ",0,'" & NomeImposto & "') "
        Bdados.Executa Sql
        GeraCodPagamento = GeraCodPagamento(Imposto)
    End If
    Bdados.FechaTabela Rs
'    Dim Procedure As VSClass.VSComando
'
'    Set Procedure = New VSComando
'        Procedure.Texto Bdados, "sp_gera_codigo_pagamento", cmdStoredProcedure
'        Procedure.setarParametro "CodImposto", tipChar, parEntrada, 15, Imposto
'        Procedure.setarParametro "NumCorrelativo", tipInteger, parEntradaSaida, , 0
'        Procedure.Executa
'        GeraCodPagamento = Procedure.Parametro("NumCorrelativo").Value
'    Set Procedure = Nothing
End Function

Public Sub DefineIsentos(Imposto As String, Periodo As Long)
    Dim ValorMinimo As Double
    Dim Sql As String
    Dim RsIsento As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Dim ClsImposto As New VSImposto
    Sql = "SELECT tpi_valor_min_imposto FROM tab_parametro_imposto WHERE " & _
        " tpi_tip_cod_imposto='" & Imposto & "'"
    Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Periodo, 4)) & "'"
    If Bdados.AbreTabela(Sql & Condicao, RsIsento) Then
        ValorMinimo = RsIsento(0)
    Else
        Bdados.FechaTabela RsIsento
        Exit Sub
    End If
    Bdados.FechaTabela RsIsento
    Sql = "SELECT tgt_im,tgt_tim_ic FROM Tab_Geracao_Tributo " & _
        " WHERE tgt_valor_tributo < " & ValorMinimo & " and tgt_tip_cod_imposto='" & Imposto & _
        "' and tgt_periodo=" & Periodo & " and tgt_parcela =0"
    If Bdados.AbreTabela(Sql, RsIsento) Then
        RsIsento.MoveFirst
        Campos = "TIS_IM,TIS_TIM_IC,TIS_COD_OBRIG,TIS_TPI_PERIODO"
        Call Bdados.DeletaDados("TAB_ISENTO", "TIS_COD_OBRIG ='" & Imposto & "' and TIS_TPI_PERIODO =" & Periodo)
        Do While Not RsIsento.EOF
            Valores = Bdados.PreparaValor(RsIsento(0), RsIsento(1), Imposto, Periodo)
            Call Bdados.GravaDados("TAB_ISENTO", Valores, Campos, "TIS_IM='" & RsIsento(0) & "' and TIS_TIM_IC='" & RsIsento(1) & "' and TIS_COD_OBRIG ='" & Imposto & "' and TIS_TPI_PERIODO =" & Periodo)
            Call Bdados.DeletaDados("Tab_Geracao_Tributo", "tgt_im='" & RsIsento(0) & "' and tgt_tim_ic='" & RsIsento(1) & "' and tgt_tip_cod_imposto ='" & Imposto & "' and tgt_periodo =" & Periodo)
            Call Bdados.DeletaDados("Tab_Conta_Contribuinte", "tcc_im='" & RsIsento(0) & "' and tcc_tim_ic='" & RsIsento(1) & "' and tcc_tip_cod_imposto ='" & Imposto & "' and tcc_periodo =" & Periodo)
            RsIsento.MoveNext
            DoEvents
        Loop
    End If
    Bdados.FechaTabela RsIsento
End Sub

Public Function BaixaPagamentos(Contribuinte As String, Periodo As String, CodTributo As String, ValorPagamento As Double, DataPagamento As String, Optional CodigoPagamento As String, Optional Obrig As Obrigacao, Optional Modo As TipoInscricaoObrigacao, Optional Parcela As Integer) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim SaldoCredorAnterior  As Double
    Dim ContaCredora As Double
    Dim CodAux As Double
    
    On Error GoTo Trata
    Dim NumObrig As String
      
    Contribuinte = Trim(Contribuinte)
    If Trim(CodigoPagamento) = "" Then
            Sql = "SELECT TDR_INSCRICAO,tdr_data_pagamento,tdr_valor_real_pago,tdr_data_vencimento, " & _
                "TDR_TGT_COD_PAGAMENTO,TDR_TGT_INCIDENCIA ,TDR_PARCELA,TDR_SIT_PAGO,TDR_DOC_ORIGEM from Tab_Darm_Recebido WHERE TDR_INSCRICAO = '" & Contribuinte & "' AND TDR_PERIODO =" & _
                Periodo & " AND TDR_TIP_COD_IMPOSTO = '" & CodTributo & "' AND TDR_SIT_PAGO = 0 AND TDR_PARCELA = " & Parcela
            If Not Bdados.AbreTabela(Sql, Rs, Estatico) Then Exit Function
    Else
        Sql = "SELECT TDR_SIT_PAGO, TDR_TGT_COD_PAGAMENTO,TDR_INSCRICAO,tdr_data_pagamento,tdr_valor_real_pago,tdr_data_vencimento, TDR_TGT_COD_PAGAMENTO ,TDR_PARCELA,TDR_SIT_PAGO,TDR_DOC_ORIGEM from Tab_Darm_Recebido WHERE tdr_tgt_cod_pagamento =" & CodigoPagamento
        If Not Bdados.AbreTabela(Sql, Rs, Estatico) Then
            Sql = "SELECT TDR_INSCRICAO,tdr_data_pagamento,tdr_valor_real_pago,tdr_data_vencimento, " & _
                    "TDR_TGT_COD_PAGAMENTO,TDR_TGT_INCIDENCIA,TDR_PARCELA,TDR_SIT_PAGO,TDR_DOC_ORIGEM from Tab_Darm_Recebido WHERE TDR_INSCRICAO = '" & Contribuinte & "' AND TDR_PERIODO =" & _
                    Periodo & " AND TDR_TIP_COD_IMPOSTO = '" & CodTributo & "' AND TDR_SIT_PAGO = 0 AND TDR_PARCELA = " & Obrig.obParcela
            If Not Bdados.AbreTabela(Sql, Rs, Estatico) Then
                Exit Function
            End If
        End If
    End If
    Parcela = Nvl("" & Rs!TDR_PARCELA, 0)
'    If Not rs.EOF And Parcela = 0 Then
    If Not Rs.EOF Then
            
        NumObrig = Obrig.CriaObrigacao(CodTributo, IIf(Len(Periodo) = 4, Periodo, Right(Periodo, 2) & Left(Periodo, 4)), IIf(Len(Periodo) = 4, Periodo, Right(Periodo, 2) & Left(Periodo, 4)), Contribuinte, Rs!tdr_valor_real_pago, etsCreditoPago, etsNaoSubstitui, "" & Rs!tdr_data_vencimento, , CodigoPagamento, , , , , , CInt(Nvl("" & Rs!TDR_PARCELA, 0)))
        If CDbl(NumObrig) > 0 Then
            If CriaContaContribuinte(CStr("" & Rs!TDR_TGT_COD_PAGAMENTO), False, Obrig) Then
                MovimentaContaContribuinte "" & Rs!TDR_TGT_COD_PAGAMENTO, DataPagamento, Obrig
            End If
        End If
        NumObrig = "" & Rs!TDR_TGT_COD_PAGAMENTO
        Do
            If CInt(Nvl("" & Rs!tdr_sit_pago, 0)) = 0 Then
                
                NumObrig = "" & Rs!TDR_TGT_COD_PAGAMENTO
                NumObrig = Obrig.CriaObrigacao(CodTributo, IIf(Len(Periodo) = 4, Periodo, Right(Periodo, 2) & Left(Periodo, 4)), IIf(Len(Periodo) = 4, Periodo, Right(Periodo, 2) & Left(Periodo, 4)), Contribuinte, Rs!tdr_valor_real_pago, etsCreditoPago, etsNaoSubstitui, "" & Rs!tdr_data_vencimento, , NumObrig, , , , , "" & Rs!TDR_DOC_ORIGEM, CInt(Nvl("" & Rs!TDR_PARCELA, 0)))
                If CDbl(NumObrig) > 0 Then
                    Call CriaContaContribuinte(NumObrig, False, Obrig)
                    Call AtualizaContaContribuinte(NumObrig, Nvl("" & Rs!tdr_data_pagamento, Date), Nvl("" & Rs!tdr_valor_real_pago, 0) + SaldoCredorAnterior, EtoCredito)
                    Call Bdados.AtualizaDados("TAB_DARM_RECEBIDO", Bdados.PreparaValor(1, NumObrig), _
                        "TDR_SIT_PAGO,tdr_tgt_cod_pagamento", "TDR_TGT_COD_PAGAMENTO = " & Rs!TDR_TGT_COD_PAGAMENTO)
                End If
            End If
            If Parcela > 0 Then VerificaFechamentoParcelamento CodigoPagamento
            Obrig.TrocaSitObrigacao NumObrig, etsCreditoPago
            Rs.MoveNext
        Loop While Not Rs.EOF
        BaixaPagamentos = True
    Else
        If Parcela > 0 Then
            Call Bdados.AtualizaDados("TAB_COTAS_OBRIGACAO", Bdados.PreparaValor(etsCreditoPago), "TCO_STATUS_OBRIGACAO_PARCELA", "TCO_COD_OBRIGACAO_PARCELA =" & CodigoPagamento)
            Call Bdados.AtualizaDados("TAB_COTAS_PARCELAMENTO", Bdados.PreparaValor(etsCreditoPago), "TCP_STATUS_OBRIGACAO_PARCELA", "TCP_NUM_COTA =" & CodigoPagamento)
            
'                    Campos = "TCP_TPA_COD_PARCELAMENTO,TCP_NUM_PARCELA,TCP_DATA_VENCIMENTO,TCP_DATA_GERACAO," & _
'                    "TCP_VALOR_PARCELA,TCP_VALOR_JUROS,TCP_STATUS_OBRIGACAO_PARCELA,TCP_NUMERO_PARCELAS," & _
'                    "TCP_NUM_COTA,TCP_TIP_COD_IMPOSTO,TCP_INSCRICAO,tcp_periodo,tcp_tipo_inscricao"
                
            
            
            VerificaFechamentoParcelamento CodigoPagamento
            BaixaPagamentos = True
        Else
            BaixaPagamentos = False
        End If
    End If
    Bdados.FechaTabela Rs
    Exit Function
Trata:
    Exit Function
    Avisa Err.Description
    Exit Function
    Resume
End Function

Public Sub InsereRemisso(Contribuinte As String, Imposto As String, Periodo As Double, DataMovimento As Date, Conta As TipoConta, Imovel As String, CodConta As Double)
    Dim Valores As String
    Dim Campos As String
    Dim Rs As VSRecordset
    Dim Sql As String
    Campos = "TRE_IM,TRE_TIP_COD_IMPOSTO,TRE_TPI_PERIODO,TRE_DATA_REMISSAO,TRE_IC,TRE_CONTA,TRE_TCC_COD_CONTA"
    Valores = Bdados.PreparaValor(Contribuinte, Imposto, Periodo, DataMovimento, Imovel, Conta, CodConta)
    Sql = "select TRE_IM from TAB_REMISSO where TRE_IM='" & Contribuinte & "' and TRE_TIP_COD_IMPOSTO='" & Imposto & "' and TRE_TPI_PERIODO=" & Periodo
    If Not Bdados.AbreTabela(Sql, Rs) Then
        Call Bdados.GravaDados("TAB_REMISSO", Valores, Campos, "TRE_IM='" & Contribuinte & "' and TRE_TIP_COD_IMPOSTO='" & Imposto & "' and TRE_TPI_PERIODO=" & Periodo & " AND TRE_CONTA =" & Conta & " AND TRE_IC='" & Imovel & "'")
    End If
    Bdados.FechaTabela Rs
End Sub

'Public Sub AtualizaContaContribuinte(CodConta As String, DataMovimento As Date, ValorAtualizado As Double, Operacao As TipoOperacaoConta, Optional TDebito As TipoDebitar, Optional MultaAcumulada As Double)
'    Dim Valores As String
'    Dim Campos As String
'    Dim Condicao As String
'    Dim ImpostoAtual As Double
'    Dim MultaAtual As Double
'    Dim JuroAtual As Double
'    Dim SaldoAtual As Double
'    Dim CreditoAtual As Double
'    Dim Sql As String
'    Dim Rs As VSRecordset
'    Dim Campo As String
'    Dim DebitoAtual As Double
'    Dim ValorPago As Double
'    Dim Desconto As Double
'    Dim NaoTributadas As Double
'    Dim TotalPagoAcumulado As Double
'    Dim CodImposto As String
'    Dim AnoImposto As String
'    Dim rsParametroImposto As VSRecordset
'    Dim SqlParametroImposto As String
'    Dim CampoPagouAMais As String
'    Dim CampoPagouAMenos As String
'
'    On Error GoTo trata
'    Condicao = "tcc_codigo_conta =" & CodConta
'    If Operacao = EtoCredito Then
'        Sql = "SELECT * FROM tab_Conta_contribuinte WHERE " & Condicao
'        If Bdados.AbreTabela(Sql, Rs) Then
'            ValorPago = ValorAtualizado
'            ImpostoAtual = Rs!tcc_imposto_atual
'            JuroAtual = Rs!tcc_juros_atual
'            MultaAtual = Rs!tcc_multa_atual
'            SaldoAtual = Rs!tcc_saldo_atual
'            CodImposto = Rs!tcc_tip_cod_imposto
'            AnoImposto = Left(CStr(Rs!tcc_periodo), 4)
'
'            If Rs!tcc_multa_atual <= ValorAtualizado Then 'Abate 1º a multa
'                MultaAtual = 0
'                ValorAtualizado = Format(ValorAtualizado - Rs!tcc_multa_atual, Const_Monetario)
'                If Rs!tcc_juros_atual <= ValorAtualizado Then 'Tenta abater juros tambem
'                    JuroAtual = 0
'                    ValorAtualizado = Format(ValorAtualizado - Rs!tcc_juros_atual, Const_Monetario)
'                    If Rs!tcc_imposto_atual <= ValorAtualizado Then 'Tenta abater imposto tambem (Zera débito)
'                        ImpostoAtual = 0
'                        ValorAtualizado = Format(ValorAtualizado - Rs!tcc_imposto_atual, Const_Monetario)
'
''                         VOLTAR ESSA LINHA ABAIXO APOS CRIAR PARAMETRO
'                         Obrig.TrocaSitObrigacao CodConta, etsCreditoPago
'
'                    Else
'                        ImpostoAtual = Rs!tcc_imposto_atual - ValorAtualizado
'                        ValorAtualizado = 0
'                    End If
'                Else 'Abate só juros e multa
'                    JuroAtual = Rs!tcc_juros_atual - ValorAtualizado
'                    ValorAtualizado = 0
'                End If
'            Else 'Abate só juros
'                MultaAtual = Rs!tcc_multa_atual - ValorAtualizado
'                ValorAtualizado = 0
'            End If
'        End If
''       RETIRAR ESSA LINHA ABAIXO APOS CRIAR PARAMETRO
'
'        Obrig.TrocaSitObrigacao CodConta, etsCreditoPago
'        Sql = "SELECT tcc_parcela,TCC_DESCONTO_CONCEDIDO,TCC_NAO_TRIBUTADA,tcc_tdr_valor_real_pago FROM tab_Conta_contribuinte WHERE " & Condicao
'        If Bdados.AbreTabela(Sql, Rs) Then
'            SqlParametroImposto = "SELECT TPI_RECEITA_A_MAIS, TPI_RECEITA_A_MENOS FROM TAB_PARAMETRO_IMPOSTO" _
'                & " WHERE tpi_tip_cod_imposto = '" & CodImposto & "' AND tpi_ano_imposto = " & Val(AnoImposto)
'            Bdados.AbreTabela SqlParametroImposto, rsParametroImposto
'
'            SaldoAtual = Format(ImpostoAtual + MultaAtual + JuroAtual, Const_Monetario)
'            'PAGOU A MENOS
'            If rsParametroImposto.EOF = False Then
'                If IsNull(rsParametroImposto!TPI_RECEITA_A_MENOS) = False Then
'                    If rsParametroImposto!TPI_RECEITA_A_MENOS = 2 Then 'SALDO DEVEDOR
'                        SaldoAtual = Format(CDbl(Nvl("" & Rs!TCC_DESCONTO_CONCEDIDO, 0)) + SaldoAtual, Const_Monetario)
'                    Else
'                        Desconto = Format(CDbl(Nvl("" & Rs!TCC_DESCONTO_CONCEDIDO, 0)) + SaldoAtual, Const_Monetario)
'                    End If
'                Else
'                    Desconto = Format(CDbl(Nvl("" & Rs!TCC_DESCONTO_CONCEDIDO, 0)) + SaldoAtual, Const_Monetario)
'                End If
'            Else
'                Desconto = Format(CDbl(Nvl("" & Rs!TCC_DESCONTO_CONCEDIDO, 0)) + SaldoAtual, Const_Monetario)
'            End If
'
'            If ValorAtualizado > 0 Then 'PAGOU A MAIS
'                'PAGOU A MAIS
'                If rsParametroImposto.EOF = False Then
'                    If IsNull(rsParametroImposto!TPI_RECEITA_A_MENOS) = False Then
'                        If rsParametroImposto!TPI_RECEITA_A_MAIS = 2 Then 'SALDO CREDOR
'                            CreditoAtual = Format(CDbl(Nvl("" & Rs!TCC_NAO_TRIBUTADA, 0)) + ValorAtualizado, Const_Monetario)
'                        Else
'                            NaoTributadas = Format(CDbl(Nvl("" & Rs!TCC_NAO_TRIBUTADA, 0)) + ValorAtualizado, Const_Monetario)
'                        End If
'                    Else
'                        NaoTributadas = Format(CDbl(Nvl("" & Rs!TCC_NAO_TRIBUTADA, 0)) + ValorAtualizado, Const_Monetario)
'                    End If
'                Else
'                    NaoTributadas = Format(CDbl(Nvl("" & Rs!TCC_NAO_TRIBUTADA, 0)) + ValorAtualizado, Const_Monetario)
'                End If
'            End If
'
'            TotalPagoAcumulado = Format(ValorPago + CDbl(Nvl("" & Rs!tcc_tdr_valor_real_pago, 0)), Const_Monetario)
'            Campos = "tcc_imposto_atual,tcc_juros_atual,tcc_multa_atual,tcc_saldo_atual," _
'            & " tcc_credito_atual,tcc_tipo_ultima_transacao,TCC_DESCONTO_CONCEDIDO,TCC_NAO_TRIBUTADA,tcc_tdr_valor_real_pago"
'
'            Valores = Bdados.PreparaValor(Bdados.Converte(Format(ImpostoAtual, Const_Monetario), TCDuplo), Bdados.Converte(Format(JuroAtual, Const_Monetario), TCDuplo), Bdados.Converte(Format(MultaAtual, Const_Monetario), TCDuplo), Bdados.Converte(Format(SaldoAtual, Const_Monetario), TCDuplo), _
'            Bdados.Converte(CreditoAtual, TCDuplo), Operacao, Bdados.Converte(Desconto, TCDuplo), Bdados.Converte(NaoTributadas, TCDuplo), Bdados.Converte(TotalPagoAcumulado, TCDuplo))
'            If SaldoAtual <= 0 Then
'                Campos = Campos & ",tcc_status_conta"
'                'VERIFICAR AQUI SE É UMA CONTA PARCELADA
'                Valores = Valores & Bdados.PreparaValor(etaFechada)
'            End If
'        End If
'    Else 'debito
'        Campos = IIf(TDebito = EtdeImposto, "tcc_imposto_atual", IIf(TDebito = EtdeMulta, "tcc_multa_atual,tcc_multa_acumulada", "tcc_juros_atual"))
'        DebitoAtual = Format(ValorAtualizado, Const_Monetario)
'        Sql = "SELECT tcc_parcela,tcc_imposto_original FROM tab_Conta_contribuinte WHERE " & Condicao
'        If Bdados.AbreTabela(Sql, Rs) Then
'            Campos = Campos & ",tcc_saldo_atual,tcc_credito_atual,tcc_tipo_ultima_transacao"
'            SaldoAtual = Format(Rs!tcc_imposto_original + JurosConta + MultaConta, Const_Monetario)
'            Valores = Bdados.PreparaValor(Bdados.Converte(Format(DebitoAtual, Const_Monetario), TCDuplo))
'            If TDebito = EtdeMulta Then
'                Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Format(MultaAcumulada, Const_Monetario), TCDuplo))
'            End If
'            Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Format(SaldoAtual, Const_Monetario), TCDuplo), Bdados.Converte(Format(CreditoAtual, Const_Monetario), TCDuplo), Operacao)
'        End If
'    End If
'    Valores = Valores & Bdados.PreparaValor(Bdados.Converte(DataMovimento, TCDataHora))
'    Campos = Campos & IIf(Trim(Campos) = "", "tcc_data_movimento", ",tcc_data_movimento")
'    Bdados.AtualizaDados "Tab_Conta_Contribuinte", Valores, Campos, Condicao
'    Bdados.FechaTabela Rs
'    Exit Sub
'trata:
'    Avisa Err.Description
'    Exit Sub
'    Resume
'End Sub

Private Sub InsereNaDividaAtiva(InscMunicipal As String, Imposto As String, Periodo As String, ValorTotalDivida As Double, ValorOriginalDivida As Double, CodContaOriginal As Double)
    Dim Valores As String
    Dim Campos As String
    Dim Rs As VSRecordset
    Dim Sql As String
    Sql = "select tda_total_divida, tda_valor_origial from TAB_DIVIDA_ATIVA WHERE " & _
    " tda_tci_im='" & InscMunicipal & "' and tda_tip_cod_imposto='" & _
    Imposto & "' and tda_periodo=" & Periodo
    If Bdados.AbreTabela(Sql, Rs) Then
        ValorTotalDivida = ValorTotalDivida + Rs!tda_total_divida
        ValorOriginalDivida = ValorOriginalDivida + Rs!tda_valor_origial
    End If
    Campos = "tda_tci_im,tda_tip_cod_imposto,tda_periodo,tda_total_divida,tda_valor_origial,tda_tus_cod_usuario"
    Valores = Bdados.PreparaValor(InscMunicipal, Imposto, Periodo, Bdados.Converte(ValorTotalDivida, TCDuplo), _
                ValorOriginalDivida, AplicacoesVTFuncoes.Usuario)
    If Rs.EOF Then
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Date, TCDataHora))
        Campos = Campos & ",tda_data_inscricao"
    End If
    Bdados.FechaTabela Rs
    Call Bdados.GravaDados("TAB_DIVIDA_ATIVA", Valores, Campos, _
        "tda_tci_im='" & InscMunicipal & "' and tda_tip_cod_imposto='" & Imposto & "' and tda_periodo=" & Periodo)
    Bdados.FechaTabela Rs
End Sub

Private Function BuscaDataPago(CodPagamento As String) As String
    Dim Rs As VSRecordset
    Dim Sql As String
    Sql = "SELECT tdr_data_pagamento from tab_darm_recebido where " & _
        " TDR_TGT_COD_PAGAMENTO =" & CodPagamento
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaDataPago = "" & Rs(0)
    End If
    Bdados.FechaTabela Rs
End Function

Private Function BuscaDataVence(CodigoConta As String) As String
    Dim Rs As VSRecordset
    Dim Sql As String
    Sql = "SELECT tcc_data_vencimento from Tab_Conta_Contribuinte where " & _
        " tcc_codigo_conta=" & CodigoConta
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaDataVence = Rs(0)
    End If
    Bdados.FechaTabela Rs
End Function

'Public Function MovimentaContaContribuinte(ByVal CodigoConta As String, _
'                                            Optional DataMovimentacao As String, Optional Obrig As Obrigacao) As Boolean
'    Dim Sql As String
'    Dim Rs As VSRecordset
'    Dim DataVencimento  As String
'    Dim DataMovimento  As String
'    Dim Venc As String
'    Dim Pag As String
'    Dim MovimentaConta As Boolean
'
'    On Error GoTo trata
'    If Not Obrig Is Nothing Then
'        If Obrig.obStatusObrigacao = 3 Then Exit Function
'    Else
'        If StatusConta(CodigoConta) = etaFechada Then
'            Exit Function
'        End If
'    End If
'    DataVencimento = ""
'    DataMovimento = ""
'    If Obrig Is Nothing Then
'        DataVencimento = BuscaDataVence(CodigoConta)
'    Else
'        DataVencimento = Obrig.obDataVencimento
'    End If
'    If Trim(DataMovimentacao) = "" Then
'        DataMovimento = BuscaDataPago(CodigoConta)
'    End If
'    Venc = Right(DataVencimento, 4) & Mid(DataVencimento, 4, 2) & Left(DataVencimento, 2)
'    DataMovimento = IIf(Trim(DataMovimento) = "", Date, DataMovimento)
'    Pag = Right(DataMovimento, 4) & Mid(DataMovimento, 4, 2) & Left(DataMovimento, 2)
'    If Trim(Venc) <> "" Then
'        Call CalculaValoresJuros(CodigoConta, DataMovimento, DataVencimento, Obrig)
'        Call CalculaValoresMulta(CodigoConta, DataMovimento, DataVencimento, Obrig)
'    Else
'        CriaContaContribuinte CodigoConta, , Obrig
' '       MovimentaContaContribuinte CodigoConta, DataMovimentacao
'    End If
'    JurosConta = 0
'    MultaConta = 0
'    Exit Function
'trata:
'    Avisa Err.Description
'    Exit Function
'    Resume
'End Function

Public Function EfetuaContaTransacoes(Contribuinte As String, Tributo As String, Periodo As String, Vencimento As String, ValorTransacao As Double, Optional DataMovimento As String, Optional TipoTransacao As TipoOperacaoConta = EtoCredito) As Boolean
'    Dim sql As String
'    Dim rs As VSRecordset
'    Dim NumObrig As Double
'    If TipoTransacao = EtoCredito Then
'        NumObrig = Obrig.CriaObrigacao(Tributo, Periodo, Periodo, Contribuinte, ValorTransacao, False)
'    End If
'    CriaContaContribuinte CStr(NumObrig)
'    MovimentaContaContribuinte NumObrig, DataMovimento
'
'    If MovimentaContaContribuinte Then
'        sql = "SELECT TOC_INSCRICAO,TOC_COD_OBRIGACAO,TOC_TIP_COD_IMPOSTO,TOC_PERIODO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_INSCRICAO ='" & CodigoConta & "'"
'        If Bdados.AbreTabela(sql, rs) Then
'            BaixaPagamentos CodigoConta, rs!TOC_INSCRICAO, rs!TOC_PERIODO, rs!TOC_TIP_COD_IMPOSTO
'        End If
'        Bdados.FechaTabela rs
'    End If
End Function
Private Sub RetiraDaDividaAtiva(Contribuinte As String, Imposto As String, Periodo As Long)
    Dim Condicao As String
    Dim Rs As VSRecordset
    
    Condicao = "tda_tci_im='" & Contribuinte & "' and tda_tip_cod_imposto='" & Imposto & "' and tda_periodo=" & Periodo
    Call Bdados.DeletaDados("TAB_DIVIDA_ATIVA", Condicao)
End Sub

Private Function CalculaValoresMulta(ByVal CodConta As String, ByVal DataPagamento As String, ByVal DataVencimento As _
                String, Optional Obrig As Obrigacao, Optional TipoTransacao As TipoTransacao = ettAvista, _
                Optional ResultadoDesconto As Double) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DiasPassados As Integer
    Dim MultaMinima As Double
    Dim MultaMaxima As Double
    Dim Juros As Double
    Dim ValorMultaCalculada As Double
    Dim ValMultaDescVista As Double
    Dim ValMultaDescParc As Double
    Dim InscMunicipal As String
    Dim InscEstadual As String
    Dim Condicao As String
    Dim Valores As String
    Dim DebitoOriginal As Double
    Dim MultaAcumulada As Double
    Dim Pago   As String
    Dim Venc As String
    Dim RsAux As VSRecordset
    Dim ClsImposto As New VSImposto
    
    'Busca no Banco
    Dim Periodo As Long
    Dim Imposto As String
    Dim Tipo As TipoConta
    Dim Parcela As Integer
    Dim IncrementoMensal As Double
    On Error GoTo Trata
    Sql = "Select tcc_periodo, tcc_tip_cod_imposto,tcc_parcela,tcc_num_conta," & _
        " tcc_multa_acumulada,tcc_saldo_atual,tcc_imposto_original,tcc_correcao_monetaria from tab_conta_contribuinte " & _
        " where tcc_codigo_conta =" & CodConta
    If Bdados.AbreTabela(Sql, RsAux) Then
        Periodo = RsAux!tcc_periodo
        Imposto = RsAux!tcc_tip_cod_imposto
        Parcela = RsAux!tcc_parcela
        Tipo = RsAux!tcc_num_conta
    Else
        ResultadoDesconto = 0
        Exit Function
    End If
    Sql = "SELECT tpi_valor_min_multa,tpi_valor_max_multa,TPI_DESCONTO_MULTA,TPI_DESCONTO_MULTA_PARC,tpi_incremento_mensal_multa,TPI_PARAMETRO_REFIS FROM tab_parametro_imposto " & _
        " where tpi_tip_cod_imposto ='" & Imposto & "'"
    Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Periodo, 4)) & "'"
    If Bdados.AbreTabela(Sql & Condicao, Rs) Then
        MultaMinima = Rs(0)
        MultaMaxima = Rs(1)
        ValMultaDescVista = Nvl("" & Rs!TPI_DESCONTO_MULTA, 0)
        ValMultaDescParc = Nvl("" & Rs!TPI_DESCONTO_MULTA_PARC, 0)
        IncrementoMensal = Nvl("" & Rs!tpi_incremento_mensal_multa, 0)
        If Nvl("" & Rs!TPI_PARAMETRO_REFIS, 1) = 2 Then
            Sql = "select TPI_DESCONTO_TRIBUTO_PARC,TPI_DESCONTO_JUROS_PARC,TPI_DESCONTO_MULTA_PARC,TPI_DESCONTO_ATUALIZACAO_PARC "
            Sql = Sql & " from TAB_PARAMETRO_IMPOSTO_REFIS where TPI_TIP_COD_IMPOSTO ='" & Imposto & "'"
            Sql = Sql & " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Periodo, 4)) & "'"
            Sql = Sql & " and " & Parcela & " BETWEEN TPI_LIMITE_INFERIOR AND TPI_LIMITE_SUPERIOR "
            If Not Bdados.AbreTabela(Sql, Rs) Then
                Exit Function
            Else
                If Parcela = 0 Then ValMultaDescVista = Nvl("" & Rs!TPI_DESCONTO_MULTA_PARC, 0)
                ValMultaDescParc = Nvl("" & Rs!TPI_DESCONTO_MULTA_PARC, 0)
            End If
        End If
    End If
    If Not Rs.EOF Then
        
        If MultaMinima = 0 Then
            ResultadoDesconto = 0
            Bdados.FechaTabela Rs
            Exit Function
        End If
    Else
        ResultadoDesconto = 0
        Bdados.FechaTabela Rs
        Exit Function
    End If
    Pago = IIf(Trim(DataPagamento) = "", Format(Date, "dd/mm/yyyy"), DataPagamento)
    Pago = Right(Pago, 4) & Mid(Pago, 4, 2) & Left(Pago, 2)
    Venc = Right(DataVencimento, 4) & Mid(DataVencimento, 4, 2) & Left(DataVencimento, 2)
    If Pago > Venc Then
        MultaAcumulada = Nvl("" & RsAux!tcc_multa_acumulada, 0)
'        MultaAcumulada = 0
        If MultaAcumulada < MultaMaxima Then
            DebitoOriginal = Format(RsAux!tcc_saldo_atual, Const_Monetario) 'FABIANO
            If DebitoOriginal > 0 Then
                If IsDate(DataVencimento) And IsDate(DataPagamento) Then
                    DiasPassados = DateDiff("m", DataVencimento, IIf(Trim(DataPagamento) = "", Format(Date, "dd/mm/yyyy"), DataPagamento))
                Else
                    Exit Function
                End If
                If DiasPassados > 0 Or Day(DataVencimento) < Day(Nvl(DataPagamento, DataVencimento)) Then
                    If Day(DataVencimento) < Day(Nvl(DataPagamento, DataVencimento)) Then
                        DiasPassados = DiasPassados + 1
                    End If
                    ValorMultaCalculada = MultaMinima + (IncrementoMensal * IIf(DiasPassados = 0, 0, (DiasPassados - 1)))
                    
                    If ValorMultaCalculada + MultaAcumulada <= MultaMaxima Then 'Ainda posso calcular
                        MultaAcumulada = MultaAcumulada + ValorMultaCalculada
                        ValorMultaCalculada = (ValorMultaCalculada / 100)
                    Else
                        If MultaAcumulada = 0 Then
                            ValorMultaCalculada = (MultaAcumulada + (MultaMaxima - MultaAcumulada)) / 100
                        Else
                            ValorMultaCalculada = (MultaMaxima - MultaAcumulada) / 100
                            MultaAcumulada = MultaAcumulada + (MultaMaxima - MultaAcumulada)
                        End If
                    End If
                End If
            Else
                ResultadoDesconto = 0
                CalculaValoresMulta = False
                Exit Function
            End If
        Else
            ResultadoDesconto = 0
            CalculaValoresMulta = False
            Exit Function
        End If
    Else
        ResultadoDesconto = 0
    End If
    DebitoOriginal = CDbl(Nvl("" & RsAux!tcc_imposto_original, 0)) + CDbl(Nvl("" & RsAux!tcc_correcao_monetaria, 0))
    
    'Nova tabela multa: Como nao ha um loop, basta trocar o valor da aliquota
    '(Problema: calcular parte do periodo de uma maneira e o restante da outra)
    Dim NovaMulta As Double
    
    NovaMulta = BuscarMultaTabelada(DataVencimento, DataPagamento)
    If NovaMulta > 0 Then ValorMultaCalculada = NovaMulta / 100
    
    ValorMultaCalculada = ValorMultaCalculada * DebitoOriginal
    If ResultadoDesconto = 0 Then
        If TipoTransacao <> ettDividaAtiva Then
            If TipoTransacao = ettAvista Then
                ResultadoDesconto = ValorMultaCalculada * (ValMultaDescVista / 100)
            Else
                ResultadoDesconto = ValorMultaCalculada * (ValMultaDescParc / 100)
            End If
        Else
            ResultadoDesconto = 0
        End If
    Else
        ResultadoDesconto = ValorMultaCalculada * (ResultadoDesconto / 100)
    End If
    ResultadoDesconto = Format(ResultadoDesconto, Const_Monetario)
    ValorMultaCalculada = Format(ValorMultaCalculada, Const_Monetario)
    Condicao = "tgt_cod_pagamento =" & CodConta
    MultaConta = ValorMultaCalculada
    Valores = Bdados.PreparaValor(Bdados.Converte(Format(ValorMultaCalculada, Const_Monetario), TCDuplo))
    If IsDate(DataPagamento) Then
        Call AtualizaContaContribuinte(CodConta, CDate(DataPagamento), ValorMultaCalculada, EtoDebito, EtdeMulta, MultaAcumulada)
    End If
    CalculaValoresMulta = True
    Bdados.FechaTabela Rs
    Exit Function
Trata:
        Avisa Err.Description
        Exit Function
        Resume
End Function
Private Sub AtualizaTransacao(CodigoConta As Double, Valor As Double, Data As String, Tipo As TipoContaTransacao)
    Dim Campos As String
    Dim Valores As String
    Campos = "tct_tcc_codigo_conta,tct_tipo_transacao,tct_valor_transacao,tct_data_transacao,tct_data_processamento"
    Valor = Format(Valor, Const_Monetario)
    Valores = Bdados.PreparaValor(CodigoConta, Tipo, Bdados.Converte(Valor, TCDuplo), Bdados.Converte(Data, TCDataHora), Bdados.Converte(Date, TCDataHora))
    Call Bdados.InsereDados("TAB_CONTA_TRANSACAO", Valores, Campos)
End Sub

Public Function CriaParcelamento(Contribuinte As String, Imposto As String, Cotas As Byte, _
            DtVencimento As String, DividaTotal As Double, ListaDeObrigacoes As Object, _
            Optional Perc_Entrada As Double, Optional Modo As TipoInscricaoObrigacao, _
            Optional OBS As String) As Double
    Dim NumParc As Double
    Dim ValorCotas As Double
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DataPagoCota As String
    Dim ValorEntrada As Double
    Dim I As Long
    Dim CodCota As Double
    Dim ValorJuros As Double
    Dim RsAux As VSRecordset
    Dim ContaParcelada  As Double
    Dim MesesPassados As Double
    Dim PeriodoInicial As String, PeriodoFinal As String
    Dim CodImposto As String
    Dim Tributo As New VSImposto
    Dim Ano As String
    Dim Multa As Double
    Dim Natureza As Integer
    
    CriaParcelamento = True
    NumParc = GeraCodPagamento(EtsFinanciamento)
    'Busca Multa,Juros Parcelamento
    Natureza = Nvl(BuscaNaturezaContribuinte(Contribuinte), 1)
    Natureza = IIf(Natureza = 0, 1, Natureza)
    Sql = " SELECT tpp_valor_juros,tpp_valor_minimo_entrada,tpp_reducao,tpp_valor_min_cota," & _
        " tpp_valor_min_base_calc,tpp_max_cotas FROM tab_parametro_parcelamento where 1 = 1 "
        Sql = Sql & " and  tpp_pessoa = " & Natureza
    If Bdados.AbreTabela(Sql, Rs) Then
        If Rs!tpp_max_cotas < 2 Then
            CriaParcelamento = False
            Exit Function
        End If
        ValorJuros = Rs(0) / 100
    Else
        CriaParcelamento = False
        Exit Function
    End If
    'Bdados.FechaTabela RS
    If Fix(Perc_Entrada) <> Fix(DividaTotal / Cotas) Then
        ValorCotas = Format(((DividaTotal - Perc_Entrada) / (Cotas - 1)), Const_Monetario)
    Else
        ValorCotas = Format((DividaTotal / Cotas), Const_Monetario)
    End If
    If Not ValidaParcelamento(DividaTotal, CDbl(Rs!tpp_valor_min_base_calc), CDbl(ValorCotas), CDbl(Rs!tpp_valor_min_cota)) Then
        CriaParcelamento = False
        Exit Function
    End If
    
    Obrig.BuscaDetalheObrigacao ListaDeObrigacoes.SelectedItem, Modo
    BuscaPeriodoParcelamento ListaDeObrigacoes, PeriodoInicial, PeriodoFinal
    ValorEntrada = Perc_Entrada '(DividaTotal * (Perc_Entrada / 100))
    Valores = Bdados.PreparaValor(NumParc, Bdados.Converte(Contribuinte, tctexto), _
    Obrig.obCodImposto, Bdados.Converte(Date, TCDataHora), Bdados.Converte(Format(DividaTotal, Const_Monetario), TCDuplo), Bdados.Converte(Format(DividaTotal, Const_Monetario), TCDuplo), _
    Cotas, AplicacoesVTFuncoes.Usuario, etaAberta, Bdados.Converte(DtVencimento, TCDataHora), PeriodoInicial, PeriodoFinal, Format(Date, "YYYYMM"), Modo, Bdados.Converte(OBS, tctexto), Bdados.Converte(Perc_Entrada, TCDuplo))

    'Cria parcelamento
    Campos = "tpa_num_parcelamento,tpa_inscricao,tpa_tip_cod_imposto," & _
            " TPA_DATA_FINANCIAMENTO,TPA_VALOR_TOTAL_DIVIDA,TPA_VALOR_PARCELADO," & _
            " TPA_NUM_COTAS,TPA_TUS_COD_USUARIO,TPA_STATUS_PARCELAMENTO,TPA_DATA_PRIMEIRA_COTA," & _
            " TPA_PERIODO_INICIAL,TPA_PERIODO_FINAL,TPA_PERIODO,TPA_TIPO_INSCRICAO,TPA_OBSERVACAO,TPA_VALOR_ENTRADA"
    Bdados.GravaDados "Tab_Parcelamento", Valores, Campos, "tpa_num_parcelamento=" & NumParc
    Dim ValorCota As Double
    
    For I = 1 To Cotas
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            CodCota = Conta.GeraCodPagamento("64") 'Conta.Correlativo("TRIB", "64", "COTAS PARCELAS")
        ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            CodCota = Conta.GeraCodPagamento(64)
        End If
        DataPagoCota = DateAdd("m", I - 1, DtVencimento)
        MesesPassados = DateDiff("m", Date, DataPagoCota)
        If I = 1 Then
            ValorCota = Perc_Entrada
        Else
            ValorCota = ValorCotas
        End If
        MesesPassados = IIf(MesesPassados < 0, ((MesesPassados * 2) * -1) + MesesPassados, MesesPassados)
        Campos = "TCP_TPA_COD_PARCELAMENTO,TCP_NUM_PARCELA,TCP_DATA_VENCIMENTO,TCP_DATA_GERACAO," & _
                "TCP_VALOR_PARCELA,TCP_VALOR_JUROS,TCP_STATUS_OBRIGACAO_PARCELA,TCP_NUMERO_PARCELAS," & _
                "TCP_NUM_COTA,TCP_TIP_COD_IMPOSTO,TCP_INSCRICAO,tcp_periodo,tcp_tipo_inscricao"
        Valores = Bdados.PreparaValor(NumParc, I, Bdados.Converte(DataPagoCota, TCDataHora), _
            Format(Date, "dd/mm/yyyy"), Bdados.Converte(Format(ValorCota, Const_Monetario), TCDuplo), Bdados.Converte(Format((ValorCota * ValorJuros * (MesesPassados)), Const_Monetario), TCDuplo), _
            etsCreditoOriginalAberto, Cotas, CodCota, Imposto, Bdados.Converte(Obrig.obContribuinte, tctexto), Month(Date) & Year(Date), Modo)
        Bdados.InsereDados "TAB_COTAS_PARCELAMENTO", Valores, Campos
        'GERO UMA OBRIGAÇÃO PARA CADA PARCELA DE ACORDO COM O NOVO PROCESSO DE PARCELAMENTO...
        Obrig.CriaObrigacao Obrig.obCodImposto, Month(Date) & Year(Date), Month(Date) & Year(Date), Obrig.obContribuinte, ValorCota, etsCreditoOriginalAberto, etsCriaNova, DataPagoCota, , CStr(CodCota), , , , , , CInt(I)
       'COLOCO O NUMERO DO PARCELAMENTO EM CADA COTA GERADA...
        Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(NumParc), "TOC_TPA_COD_PARCELAMENTO", "TOC_COD_OBRIGACAO=" & CodCota
        
'        CriaContaContribuinte CStr(CodCota)
'        MovimentaContaContribuinte CodCota, Date
    Next
    'TROCA STATUS OBRIGACOES
    For I = 1 To ListaDeObrigacoes.ListItems.Count
        If ListaDeObrigacoes.ListItems(I).Checked = True Then
            Obrig.TrocaSitObrigacao ListaDeObrigacoes.ListItems(I), etsCreditoParcelado
            Bdados.AtualizaDados "TAB_CONTA_CONTRIBUINTE", Bdados.PreparaValor(NumParc), "TCC_TPA_COD_PARCELAMENTO", "tcc_codigo_conta =" & ListaDeObrigacoes.ListItems(I)
            Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(NumParc), "TOC_TPA_COD_PARCELAMENTO", "TOC_COD_OBRIGACAO=" & ListaDeObrigacoes.ListItems(I)
            Bdados.Executa "UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_VALOR_OBRIGACAO_ORIGINAL =TOC_VALOR_OBRIGACAO WHERE  TOC_COD_OBRIGACAO = " & ListaDeObrigacoes.ListItems(I)
            Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(Bdados.Converte("0", TCMonetario)), "TOC_VALOR_OBRIGACAO", "TOC_COD_OBRIGACAO=" & ListaDeObrigacoes.ListItems(I)
        End If
    Next
    CriaParcelamento = NumParc
End Function

Private Function BuscaDadosParcelamento(ByVal Lista As Object, Index As Integer, Sub_Items As Integer) As String
    On Error GoTo Trata
    Dim I As Integer
    
    For I = Index To Lista.ListItems.Count
        If Lista.ListItems(Index).SubItems(Sub_Items) <> "" Then
            BuscaDadosParcelamento = Lista.ListItems(Index).SubItems(Sub_Items)
        End If
    Next
    
    
   '' For Each Item In Lista.ListItems
    ''    If Item.Checked Then
            ''If Item.SubItems(3) < Inicial Or Inicial = "" Then
      ''          Inicial = Item.SubItems(3)
            ''End If
        ''    If Item.SubItems(3) > Final Or Final = "" Then
              ''  Final = Item.SubItems(3)
          ''  End If
        ''End If
    ''Next
    Exit Function
Trata:
    Erro Err.Description
End Function

Private Sub BuscaPeriodoParcelamento(ByVal Lista As Object, ByRef Inicial As String, ByRef Final As String)
    On Error GoTo Trata
    Dim Item As ListItem
    Inicial = ""
    Final = ""
    For Each Item In Lista.ListItems
        If Item.Checked Then
            If Item.SubItems(3) < Inicial Or Inicial = "" Then
                Inicial = Item.SubItems(3)
            End If
            If Item.SubItems(3) > Final Or Final = "" Then
                Final = Item.SubItems(3)
            End If
        End If
    Next
    Exit Sub
Trata:
    Erro Err.Description
End Sub
Public Sub TrocaStatusConta(CodPagamento As String, Status As TipoStatusConta)
    Dim Condicao As String
    Dim Valores As String
        
    Call Bdados.Executa("UPDATE tab_Conta_contribuinte SET tcc_status_conta =" & Status & _
                        " WHERE tcc_codigo_conta = " & CodPagamento)
    
'    Valores = Bdados.PreparaValor(Status)
'    Condicao = "tcc_codigo_conta = " & CodPagamento
'    Call Bdados.AtualizaDados("tab_Conta_contribuinte", Valores, "tcc_status_conta", Condicao)
End Sub

Private Function StatusConta(CodigoConta As String) As TipoStatusConta
    Dim Condicao As String
    Dim Rs As VSRecordset
    On Error GoTo Trata
    Condicao = "tcc_codigo_conta=" & CodigoConta
    If Bdados.AbreTabela("Select tcc_status_conta from tab_Conta_Contribuinte where " & Condicao, Rs) Then
        StatusConta = Rs(0)
    End If
    Bdados.FechaTabela Rs
    Exit Function
Trata:
    Avisa Err.Description
    Exit Function
End Function

Public Function NumPagamento(Contribuinte As String, Periodo As Long, Imposto As String, Imovel As String, Parcela As Integer, Novo As Boolean, Incidencia As Integer) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Tabela As String
    Dim CodIPTU As New VSImposto
    
    Tabela = IIf(Parcela = 0, "tab_obrigacao_contribuinte", "Tab_Geracao_Tributo_Parcela")
    Sql = "Select toc_cod_obrigacao  from  " & Tabela
    Sql = Sql & " where toc_tip_cod_imposto='" & Imposto & "' and toc_periodo=" & Periodo

    If Imposto <> CodIPTU.BuscaCodImposto(CodIPTU.NomeTributo(1)) _
                And Imposto <> CodIPTU.BuscaCodImposto(CodIPTU.NomeTributo(20)) _
                And Imposto <> CodIPTU.BuscaCodImposto(CodIPTU.NomeTributo(21)) _
                And Imposto <> CodIPTU.BuscaCodImposto(CodIPTU.NomeTributo(22)) _
                And Imposto <> CodIPTU.BuscaCodImposto(CodIPTU.NomeTributo(23)) Then
        Sql = Sql & " and tgt_im='" & Contribuinte & "' and tgt_incidencia =" & Incidencia
    End If
    'If Trim(Imovel) <> "" Then Sql = Sql & " and tgt_tim_ic='" & Imovel & "' and tgt_parcela=" & Parcela
    Sql = Sql & " and toc_inscricao = '" & Contribuinte & "' and toc_parcela = " & Parcela
    If Not Bdados.AbreTabela(Sql, Rs) Then
        NumPagamento = GeraCodPagamento(Imposto)
        Novo = True
    Else
        If Rs(0) <> 0 Then
            NumPagamento = Rs(0)
        Else
            NumPagamento = GeraCodPagamento(Imposto)
        End If
        Novo = False
    End If
    Bdados.FechaTabela Rs
End Function


Public Sub CriaDividaAtiva(AnoCorrente As Integer, Optional Contribuinte As String)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DividaImposto As Double
    Dim JurosImposto As Double
    Dim MultaImposto As Double
    Dim Im As String
    Dim Imposto As String
    Dim Valores As String
    Dim ContaPendente As Boolean
    Dim CodPagamento As String
    'Apaga Divida Ativa atual
    Bdados.DeletaDados "Tab_Conta_Contribuinte", "tcc_num_conta = 4" & IIf(Trim(Contribuinte) <> "", " AND tcc_im ='" & Contribuinte & "'", "")
    Sql = "Select tcc_im,tcc_tim_ic,tcc_tip_cod_imposto,tcc_periodo,tcc_parcela," & _
        "tcc_num_conta,tcc_saldo_atual,tcc_juros_atual,tcc_multa_atual " & _
        " from tab_conta_contribuinte where tcc_saldo_atual > 0 and left(tcc_periodo,4) < " & AnoCorrente
    If Trim(Contribuinte) <> "" Then
        Sql = Sql & " and tcc_im ='" & Contribuinte & "'"
    End If
    Sql = Sql & " order by tcc_im,tcc_tip_cod_imposto"
    DividaImposto = 0
    JurosImposto = 0
    MultaImposto = 0
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Contribuinte = Rs!tcc_im
        Imposto = Rs!tcc_tip_cod_imposto
        Do While Not Rs.EOF
            If Contribuinte <> Rs!tcc_im Or Imposto <> Rs!tcc_tip_cod_imposto Then
                CodPagamento = GeraCodPagamento(EtsDividaAtiva)
                GeraPagamento Rs!tcc_im, "", Rs!tcc_tip_cod_imposto, AnoCorrente - 1, "31/12/" & AnoCorrente, DividaImposto, 0, 0, CodPagamento, 0, 0, 0
                CriaContaContribuinte CodPagamento
                DividaImposto = 0
                JurosImposto = 0
                MultaImposto = 0
                ContaPendente = False
            Else
                ContaPendente = True
                MovimentaContaContribuinte CodPagamento, Date
                DividaImposto = DividaImposto + CDbl(Nvl("" & Rs!tcc_saldo_atual, 0))
                JurosImposto = JurosImposto + CDbl(Nvl("" & Rs!tcc_juros_atual, 0))
                MultaImposto = MultaImposto + CDbl(Nvl("" & Rs!tcc_multa_atual, 0))
                TrocaStatusConta CodPagamento, etaFechada
            End If
            Contribuinte = Rs!tcc_im
            Imposto = Rs!tcc_tip_cod_imposto
            Rs.MoveNext
        Loop
        If ContaPendente Then
            InsereNaDividaAtiva Contribuinte, Imposto, AnoCorrente - 1, DividaImposto + MultaImposto + JurosImposto, DividaImposto, 0
            GeraPagamento Contribuinte, "", Imposto, AnoCorrente - 1, "31/12/" & AnoCorrente, DividaImposto, 0, 0, GeraCodPagamento(EtsDividaAtiva), 0, 0, 0
            CriaContaContribuinte CodPagamento
        End If
        Bdados.FechaTabela Rs
    End If
End Sub

Public Sub GeraContasCorrentes(Contribuinte As String)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Condicao As String
    
'    ApagaContasAnteriores Contribuinte
'    PendenciarPagamentos Contribuinte
    Sql = "Select tgt_cod_pagamento, tgt_im,tgt_tim_ic,tgt_tip_cod_imposto,tgt_periodo,tgt_data_vencimento," & _
        "tgt_valor_tributo,tgt_tpa_num_parcelamento,TGT_PARCELA,tgt_incidencia from tab_geracao_tributo "
    If Trim(Contribuinte) <> "" Then
        Sql = Sql & " where tgt_im='" & Contribuinte & "' order by tgt_tip_cod_imposto, tgt_parcela"
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            CriaContaContribuinte Rs!TGT_COD_PAGAMENTO
            MovimentaContaContribuinte Rs!TGT_COD_PAGAMENTO, Date
            Rs.MoveNext
            DoEvents
        Loop
    End If
    Bdados.FechaTabela Rs
End Sub

Public Sub PendenciarPagamentos(Contribuinte As String, Modo As TipoInscricaoObrigacao, Optional CodObrigacao As String, Optional CodImposto As String)
    Dim Condicao As String
        
    If Trim(Contribuinte) <> "" Then
        Condicao = "TDR_INSCRICAO='" & Contribuinte & "' "
        Condicao = Condicao & " and tdr_tipo_inscricao = " & Modo
        If Trim(CodImposto) <> "" Then Condicao = Condicao & " and tdr_tip_cod_imposto ='" & CodImposto & "'"
    End If
    If Trim(CodObrigacao) <> "" Then Condicao = " tdr_tgt_cod_pagamento =" & CodObrigacao
    Bdados.AtualizaDados "TAB_DARM_RECEBIDO", Bdados.PreparaValor(0), "tdr_sit_pago", Condicao
    
End Sub
Public Sub GeraPagamento(InscricaoMunicipal As String, InscricaoCadastral As String, Imposto As String, Periodo As Long, DataVencePagamento As String, ValorPagamento As Double, ValorMulta As Double, ValorJuros As Double, NumeroPagamento As String, NumParcelamento As Double, Parcela As Integer, Taxas As Double, Optional CodTaxa As String, Optional Conta As TipoConta, Optional Ativo As Integer = 0, Optional Incidencia As Integer = 0, Optional NumPagamentoOriginal As String = 0, Optional Tipo As Tipo_Geracao = etgDAM, Optional Tipo_DOC As Tipo_Dam = etdNormal)
    Dim Campos As String
    Dim P As Object
    Dim Valores As String
    Dim ClsImposto As New VSImposto
    Dim Cobranca As New VSCobranca
    Dim CodPagamentoTaxa As Double
    Dim LinhaDigitavel As String
    Dim Tabela As String
    Dim Sql As String, Rs As VSRecordset
    Dim CodBarra As New CodigoDeBarra
    
    
    Tabela = "TAB_GERACAO_TRIBUTO"
    Campos = "tgt_im,tgt_tip_cod_imposto,tgt_periodo,tgt_data_vencimento,tgt_valor_tributo,tgt_valor_multa,tgt_valor_juros,tgt_cod_pagamento,tgt_tpa_num_parcelamento,tgt_parcela,tgt_data_geracao,tgt_taxa_expediente,TGT_ATIVO,TGT_INCIDENCIA,TGT_COD_PAGAMENTO_VINCULADO,TGT_COD_PAGAMENTO_ORIGINAL, TGT_INSCRICAO,tgt_tipo_inscricao,tgt_tus_cod_usuario"
    
    
    'Tabela = IIf(Parcela = 0, "TAB_GERACAO_TRIBUTO", "TAB_GERACAO_TRIBUTO_PARCELA")
    'CRIA A LINHA CRIPTOGRAFADA PARA PODER INSERIR NO TDAM_BOLETO.RPT
    Valores = ""
    If Parcela = 0 Then
        Dim Dados As String
        If AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
            Dados = InscricaoMunicipal
        Else
            Dados = IIf(Len(Trim(InscricaoMunicipal)) = 11, Trim(InscricaoMunicipal), "")
        End If
        Valores = Bdados.PreparaValor(Bdados.Converte(Dados, tctexto), Imposto, Periodo, _
        Bdados.Converte(Format(DataVencePagamento, "DD/MM/YYYY"), TCDataHora), Bdados.Converte(Format(ValorPagamento, Const_Monetario), TCDuplo), Bdados.Converte(Format(ValorMulta, Const_Monetario), TCDuplo), Bdados.Converte(Format(ValorJuros, Const_Monetario), TCDuplo), NumeroPagamento, NumParcelamento, _
        Parcela, Bdados.Converte(Date, TCDataHora), Bdados.Converte(Format(Taxas, Const_Monetario), TCMonetario), Ativo, Incidencia, IIf(Val(NumPagamentoOriginal) = 0, NumeroPagamento, NumPagamentoOriginal), NumeroPagamento, Bdados.Converte(Trim(IIf(Trim(InscricaoCadastral) <> "", InscricaoCadastral, InscricaoMunicipal)), tctexto), IIf(Trim(InscricaoCadastral) <> "", 1, 2), AplicacoesVTFuncoes.Usuario)
        
'==========================================================
        
        If Trim(InscricaoCadastral) <> "" Then
            Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Trim(InscricaoCadastral), tctexto))
            Campos = Campos & ",tgt_tim_ic"
        End If
    Else
        Valores = Bdados.PreparaValor(NumParcelamento, NumeroPagamento, Parcela, Bdados.Converte(Date, TCDataHora), Bdados.Converte(Format(DataVencePagamento, "DD/MM/YYYY"), TCDataHora), Bdados.Converte(Format(ValorMulta, Const_Monetario), TCDuplo), Bdados.Converte(Format(ValorMulta, Const_Monetario), TCDuplo), Bdados.Converte(Format(ValorJuros, Const_Monetario), TCDuplo), etsCreditoOriginalAberto, Parcela)

        'Imposto, Periodo, _
        'Bdados.Converte(Format(DataVencePagamento, "DD/MM/YYYY"), TCDataHora), Bdados.Converte(Format(ValorPagamento, Const_Monetario), TCDuplo), Bdados.Converte(Format(ValorMulta, Const_Monetario), TCDuplo), Bdados.Converte(Format(ValorJuros, Const_Monetario), TCDuplo), NumeroPagamento, NumParcelamento, _
        'Parcela, , Bdados.Converte(Format(Taxas, Const_Monetario), TCMonetario), Ativo, Incidencia, IIf(Val(NumPagamentoOriginal) = 0, NumeroPagamento, NumPagamentoOriginal), NumeroPagamento, Bdados.Converte(Trim(IIf(InscricaoCadastral <> "", InscricaoCadastral, InscricaoMunicipal)), tctexto))
    End If
    'Dim Picture As Object
    Dim a As Object
    Dim b As String
    'If Tipo = etgIPTU Then
    '    b = CodBarra.CriaLinhaDigitavel(InscricaoCadastral, Imposto, ValorPagamento + Taxas, CStr(Periodo), DataVencePagamento, Nvl(CStr(Parcela), 0))
'    ElseIf Tipo = etgDAM Then
        If AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
            b = CodBarra.CriaLinhaDigitavelCBR(TiraPic(InscricaoMunicipal, "-"), Imposto, ValorPagamento + Taxas, CStr(Periodo), P, DataVencePagamento, Nvl(CStr(Parcela), 0), NumeroPagamento)
        Else
            b = CodBarra.CriaLinhaDigitavel(TiraPic(InscricaoMunicipal, "-"), Imposto, ValorPagamento + Taxas, CStr(Periodo), DataVencePagamento, Nvl(CStr(Parcela), 0), NumeroPagamento)
        End If
'    End If
    'b = CodBarra.CriaLinhaDigitavel(CodPagamentoTaxa, Imposto, ValorPagamento + Taxas, CStr(Periodo), DataVencePagamento, Nvl(CStr(Parcela), 0))
    LinhaDigitavel = CodBarra.LinhaBarraGerada
    
    If Parcela = 0 Then
        If Edita.PosPic(LinhaDigitavel, "'") = 0 Then
            Campos = Campos & " ,tgt_linha_barcode,tgt_linha"
            If Bdados.Conexao.FormatoBanco = SQLServer Then
                Valores = Valores & Bdados.PreparaValor("'" & LinhaDigitavel & "'", b)
            ElseIf Bdados.Conexao.FormatoBanco = oracle Then
                Valores = Valores & Bdados.PreparaValor(Bdados.Converte(LinhaDigitavel, tctexto), b)
            End If
            
        End If
        Campos = Campos & " , tgt_tipo "
        Valores = Valores & Bdados.PreparaValor(Tipo_DOC)
        If Bdados.GravaDados(Tabela, Valores, Campos, "tgt_cod_pagamento =" & Bdados.Converte(NumeroPagamento, tctexto)) = False Then
            Util.Avisa "ERRO"
        End If
    Else
        If Edita.PosPic(LinhaDigitavel, "'") = 0 Then
            Campos = Campos & " ,TCO_BARRA,TCO_LINHA"
            Valores = Valores & Bdados.PreparaValor("'" & LinhaDigitavel & "'", b)
        End If
        Bdados.GravaDados Tabela, Valores, Campos, "TCO_COD_OBRIGACAO_PARCELA = " & Bdados.Converte(NumParcelamento, tctexto)
    End If
    
    
    If Len(CodTaxa) > 0 And Taxas > 0 Then
      'Grava na Tabela de Taxas
        CodPagamentoTaxa = NumPagamento(InscricaoMunicipal, CLng(Periodo), CodTaxa, Trim(InscricaoCadastral), Parcela, False, 0)
'            CodPagamentoTaxa = GeraCodPagamento(CodTaxa)
        Campos = "TTD_TGT_COD_PAGAMENTO,TTD_TGT_COD_PAGAMENTO_TAXA,TDD_TIP_COD_IMPOSTO"
        Valores = Bdados.PreparaValor(NumeroPagamento, CodPagamentoTaxa, CodTaxa)
        Bdados.InsereDados "TAB_TAXA_DAM", Valores, Campos
        'Cria o registro da taxa na Tab_Geracao_Tributo
        
        Campos = "tgt_im,tgt_tim_ic,tgt_tip_cod_imposto,tgt_periodo,tgt_data_vencimento,tgt_valor_tributo,tgt_cod_pagamento,tgt_parcela,tgt_data_geracao,tgt_ativo,TGT_COD_PAGAMENTO_VINCULADO,TGT_COD_PAGAMENTO_ORIGINAL"
        Valores = Bdados.PreparaValor(InscricaoMunicipal, Bdados.Converte(Trim(InscricaoCadastral), tctexto), CodTaxa, Periodo, _
        Bdados.Converte(DataVencePagamento, TCDataHora), Bdados.Converte(Taxas, TCDuplo), CodPagamentoTaxa, Parcela, Bdados.Converte(Date, TCDataHora), 0, NumeroPagamento, CodPagamentoTaxa)
        Bdados.GravaDados Tabela, Valores, Campos, "tgt_cod_pagamento=" & CodPagamentoTaxa
    End If
End Sub
Public Function CalculaValoresJurosAvulsos(Imposto As String, Periodo As Long, Tipo As TipoConta, _
                ByVal DataPagamento As String, ByVal DataVencimento As String, ValorDivida As Double, _
                Optional TipoTransacao As TipoTransacao = ettAvista, Optional ResultadoDesconto As Double) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DiasPassados As Double
    Dim ValJuros As Double
    Dim ValJurosDescVista As Double
    Dim ValJurosDescParc As Double
    Dim MultaMaxima As Double
    Dim Condicao As String
    Dim Valores As String
    Dim DebitoRestante  As Double
    Dim Pag As String
    Dim Venc As String
    Dim Cotas As Byte
    Dim ValorCota As Double
    Dim MesesPassados As Double
    Dim ClsImposto As New VSImposto
    Dim TipoJuros As Integer
    Dim Parcela As Integer
    
    On Error GoTo Trata

    Sql = "SELECT tpi_valor_juros,tpi_dias_pagar,tpi_juros_captalizados,TPI_DESCONTO_JUROS," & _
        "TPI_DESCONTO_JUROS_PARC,TPI_PARAMETRO_REFIS  FROM tab_parametro_imposto" & _
        " where tpi_tip_cod_imposto ='" & Imposto & "'"
    Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Periodo, 4)) & "'"
    If Bdados.AbreTabela(Sql & Condicao, Rs) Then
        TipoJuros = Nvl("" & Rs!tpi_juros_captalizados, 0)
        ValJuros = Nvl("" & Rs!tpi_valor_juros, 0)
        If Nvl("" & Rs!TPI_PARAMETRO_REFIS, 1) = 2 Then
            Sql = "select TPI_DESCONTO_TRIBUTO_PARC,TPI_DESCONTO_JUROS_PARC,TPI_DESCONTO_MULTA_PARC,TPI_DESCONTO_ATUALIZACAO_PARC "
            Sql = Sql & " from TAB_PARAMETRO_IMPOSTO_REFIS where TPI_TIP_COD_IMPOSTO ='" & Imposto & "'"
            Sql = Sql & " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Periodo, 4)) & "'"
            Sql = Sql & " and " & Parcela & " BETWEEN TPI_LIMITE_INFERIOR AND TPI_LIMITE_SUPERIOR "
            If Not Bdados.AbreTabela(Sql, Rs) Then
                Exit Function
            End If
        Else
            ValJurosDescVista = Nvl("" & Rs!TPI_DESCONTO_JUROS, 0)
        End If
    End If
    If Not Rs.EOF Then
        Pag = IIf(Trim(DataPagamento) = "", Format(Date, "dd/mm/yyyy"), DataPagamento)
        Pag = Right(Pag, 4) & Mid(Pag, 4, 2) & Left(Pag, 2)
        Venc = Right(DataVencimento, 4) & Mid(DataVencimento, 4, 2) & Left(DataVencimento, 2)
        If Pag > Venc Then
            DebitoRestante = ValorDivida
            If DebitoRestante > 0 Then
                DiasPassados = DateDiff("d", IIf(IsDate(DataVencimento), DataVencimento, Date), IIf(Trim(DataPagamento) = "", Format(Date, "dd/mm/yyyy"), DataPagamento))
                If DiasPassados > 0 Then
                    MesesPassados = DateDiff("m", IIf(IsDate(DataVencimento), DataVencimento, Date), IIf(Trim(DataPagamento) = "", Format(Date, "dd/mm/yyyy"), DataPagamento))
                    If Day(DataVencimento) < Day(Nvl(DataPagamento, IIf(IsDate(DataVencimento), DataVencimento, Date))) Then
                        MesesPassados = MesesPassados + 1
                    End If
'                    ValJuros = Rs(0)
'                    ValJurosDescVista = Nvl("" & Rs!TPI_DESCONTO_JUROS, 0)
                    If Parcela = 0 Then ValJurosDescVista = Nvl("" & Rs!TPI_DESCONTO_JUROS_PARC, 0)
                    ValJurosDescParc = Nvl("" & Rs!TPI_DESCONTO_JUROS_PARC, 0)
                    If TipoJuros = 1 Then
                        ValJuros = (ValJuros * MesesPassados) / 100 'Calculo Original(Colocar para ter juros simples)
                    Else
                        ValJuros = ValJuros / 100 'Modificado pedido Silmar(Juros capitalizado)
                        'FORMULA MATEMATICA POR PEDRO SILMAR BÖSING
                        ValJuros = (CDbl(1 + ValJuros) ^ MesesPassados)
                        ValJuros = Format(ValJuros - 1, "0.000000000")
                    End If
                    ValJuros = Format(ValJuros * DebitoRestante, Const_Monetario)
                    If TipoTransacao <> ettDividaAtiva Then
                        If TipoTransacao = ettAvista Then
                            ResultadoDesconto = ValJuros * (ValJurosDescVista / 100)
                        Else
                            ResultadoDesconto = ValJuros * (ValJurosDescParc / 100)
                        End If
                    Else
                        ResultadoDesconto = 0
                    End If
                    ResultadoDesconto = Format(ResultadoDesconto, Const_Monetario)
                    'FIM DE FORMULA
                    CalculaValoresJurosAvulsos = ValJuros
                End If
            End If
        End If
    End If
    Bdados.FechaTabela Rs
    Exit Function
Trata:
        Avisa Err.Description
        Exit Function
        Resume
End Function

Private Function BuscarMultaTabelada(Vencimento As String, Pagamento As String) As Double
    Dim Sql As String, Rs As VSRecordset
    Dim Dias As Integer
    On Error Resume Next
    Dias = DateDiff("d", Vencimento, IIf(Trim(Pagamento) = "", Format(Date, "dd/mm/yyyy"), Pagamento))
    Sql = "SELECT TMI_ALIQUOTA FROM TAB_MULTA_IMPOSTO WHERE TMI_VALIDADE<=" & Bdados.FormataValorCampo(Vencimento) & _
            " AND (" & _
                    " (TMI_DIA_INICIO<=" & Dias & " AND TMI_DIA_FIM>=" & Dias & ")" & _
                    " OR (TMI_DIA_INICIO=0 AND TMI_DIA_FIM>=" & Dias & ")" & _
                    " OR (TMI_DIA_INICIO<=" & Dias & " AND TMI_DIA_FIM=0)" & _
                ")"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscarMultaTabelada = Rs!TMI_ALIQUOTA
    End If
    Bdados.FechaTabela Rs
End Function

Public Function CalculaValoresMultaAvulsos(Imposto As String, Periodo As Long, Tipo As TipoConta, _
            ByVal DataPagamento As String, ByVal DataVencimento As String, ValorDivida As Double, _
            Optional TipoTransacao As TipoTransacao = ettAvista, Optional ResultadoDesconto As Double) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DiasPassados As Integer
    Dim MultaMinima As Double
    Dim MultaMaxima As Double
    Dim ValMultaDescVista As Double
    Dim ValMultaDescParc As Double
    Dim ValorMultaCalculada As Double
    Dim InscMunicipal As String
    Dim InscEstadual As String
    Dim Condicao As String
    Dim Valores As String
    Dim DebitoOriginal As Double
    Dim MultaAcumulada As Double
    Dim Pago   As String
    Dim Venc As String
    Dim RsAux As VSRecordset
    Dim ClsImposto As New VSImposto
    Dim IncrementoMensal As Double
    Dim Parcela As Double
    
    On Error GoTo Trata
    
    
    Sql = "SELECT tpi_valor_min_multa,tpi_valor_max_multa,TPI_DESCONTO_MULTA,TPI_DESCONTO_MULTA_PARC," & _
        " tpi_incremento_mensal_multa,TPI_PARAMETRO_REFIS  FROM tab_parametro_imposto " & _
        " where tpi_tip_cod_imposto ='" & Imposto & "'"
    Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Periodo, 4)) & "'"
        If Bdados.AbreTabela(Sql & Condicao, Rs) Then
        MultaMinima = Rs(0)
        MultaMaxima = Rs(1)
        ValMultaDescVista = Nvl("" & Rs!TPI_DESCONTO_MULTA, 0)
        ValMultaDescParc = Nvl("" & Rs!TPI_DESCONTO_MULTA_PARC, 0)
        IncrementoMensal = Nvl("" & Rs!tpi_incremento_mensal_multa, 0)
        If Nvl("" & Rs!TPI_PARAMETRO_REFIS, 1) = 2 Then
            Sql = "select TPI_DESCONTO_TRIBUTO_PARC,TPI_DESCONTO_JUROS_PARC,TPI_DESCONTO_MULTA_PARC,TPI_DESCONTO_ATUALIZACAO_PARC "
            Sql = Sql & " from TAB_PARAMETRO_IMPOSTO_REFIS where TPI_TIP_COD_IMPOSTO ='" & Imposto & "'"
            Sql = Sql & " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Periodo, 4)) & "'"
            Sql = Sql & " and " & Parcela & " BETWEEN TPI_LIMITE_INFERIOR AND TPI_LIMITE_SUPERIOR "
            If Not Bdados.AbreTabela(Sql, Rs) Then
                Exit Function
            Else
                If Parcela = 0 Then ValMultaDescVista = Nvl("" & Rs!TPI_DESCONTO_MULTA_PARC, 0)
                ValMultaDescParc = Nvl("" & Rs!TPI_DESCONTO_MULTA_PARC, 0)
            End If
        End If
    End If
    If Not Rs.EOF Then
        If MultaMinima = 0 Then
            Bdados.FechaTabela Rs
            Exit Function
        End If
    Else
        Bdados.FechaTabela Rs
        Exit Function
    End If
    Pago = IIf(Trim(DataPagamento) = "", Format(Date, "dd/mm/yyyy"), DataPagamento)
    Pago = Right(Pago, 4) & Mid(Pago, 4, 2) & Left(Pago, 2)
    Venc = Right(DataVencimento, 4) & Mid(DataVencimento, 4, 2) & Left(DataVencimento, 2)
    If Pago > Venc Then
        MultaAcumulada = 0
        If MultaAcumulada < MultaMaxima Then
            DebitoOriginal = ValorDivida
            If DebitoOriginal > 0 Then
                DiasPassados = DateDiff("m", CDate(DataVencimento), IIf(Trim(DataPagamento) = "", Format(Date, "dd/mm/yyyy"), DataPagamento))
                If DiasPassados > 0 Or Day(DataVencimento) < Day(Nvl(DataPagamento, DataVencimento)) Then
                    If Day(DataVencimento) < Day(Nvl(DataPagamento, DataVencimento)) Then
                        DiasPassados = DiasPassados + 1
                    End If
                    
                    'ValorMultaCalculada = MultaMinima * DiasPassados
                    ValorMultaCalculada = MultaMinima + (IncrementoMensal * IIf(DiasPassados = 0, 0, (DiasPassados - 1)))
                    If ValorMultaCalculada <= MultaMaxima Then 'Ainda posso calcular
                        MultaAcumulada = MultaAcumulada + ValorMultaCalculada
                        ValorMultaCalculada = (ValorMultaCalculada / 100)
                    Else
                        If MultaAcumulada = 0 Then
                            ValorMultaCalculada = (MultaAcumulada + (MultaMaxima - MultaAcumulada)) / 100
                        Else
                            ValorMultaCalculada = (MultaMaxima - MultaAcumulada) / 100
                            MultaAcumulada = MultaAcumulada + (MultaMaxima - MultaAcumulada)
                        End If
                    End If
                    
                    'Nova tabela multa: Como nao ha um loop, basta trocar o valor da aliquota
                    '(Problema: calcular parte do periodo de uma maneira e o restante da outra)
                    Dim NovaMulta As Double
                    
                    NovaMulta = BuscarMultaTabelada(DataVencimento, DataPagamento)
                    If NovaMulta > 0 Then ValorMultaCalculada = NovaMulta / 100
                    
                    ValorMultaCalculada = ValorMultaCalculada * DebitoOriginal
                    If TipoTransacao <> ettDividaAtiva Then
                        If TipoTransacao = ettAvista Then
                            ResultadoDesconto = ValorMultaCalculada * (ValMultaDescVista / 100)
                        Else
                            ResultadoDesconto = ValorMultaCalculada * (ValMultaDescParc / 100)
                        End If
                    Else
                        ResultadoDesconto = 0
                    End If
                    CalculaValoresMultaAvulsos = Format(ValorMultaCalculada, Const_Monetario)
                    ResultadoDesconto = Format(ResultadoDesconto, Const_Monetario)
                End If
            Else
                CalculaValoresMultaAvulsos = 0
                Exit Function
            End If
        Else
            CalculaValoresMultaAvulsos = 0
            Exit Function
        End If
    End If
    Bdados.FechaTabela Rs
    Exit Function
Trata:
        Avisa Err.Description
        Exit Function
        Resume
End Function

Public Sub CalculaTributoProporcional(ByRef ImpostoProporcional As Double, ByRef TaxaProporcional As Double, ByVal ValorOriginalTotal As Double, ByVal ValorOriginalTaxa As Double, ByVal ValorPago As Double)
    
    If ValorOriginalTotal > 0 Then
        TaxaProporcional = ValorOriginalTaxa / ValorOriginalTotal
    Else
        TaxaProporcional = 0
    End If
    ImpostoProporcional = 1 - TaxaProporcional
    TaxaProporcional = Format(TaxaProporcional * IIf(ValorPago > 0, ValorPago, ValorOriginalTaxa), Const_Monetario)
    ImpostoProporcional = Format(ImpostoProporcional * IIf(ValorPago > 0, ValorPago, ValorOriginalTotal - ValorOriginalTaxa), Const_Monetario)
End Sub


Public Function ValidaParcelamento(ValorParcelado As Double, ValorMinBaseCalc As Double, ValorCotas As Double, ValorMinCota As Double) As Boolean
    ValidaParcelamento = True
    If ValorParcelado < ValorMinBaseCalc Then
        ValidaParcelamento = False
        Exit Function
    End If
    If ValorCotas < ValorMinCota Then
        ValidaParcelamento = False
        Exit Function
    End If
End Function


Private Sub VerificaFechamentoParcelamento(CodPagamentoParcela As String)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim CodParcelada As String
    Dim SaldoAtual As Double
    Dim CodParcelamento As Double
    
    Bdados.AtualizaDados "TAB_COTAS_PARCELAMENTO", Bdados.PreparaValor(etsCreditoPago, etsCreditoParcelado), "TCP_STATUS_OBRIGACAO_PARCELA", "tcp_num_cota =" & CodPagamentoParcela
    Bdados.AtualizaDados "TAB_COTAS_OBRIGACAO", Bdados.PreparaValor(etsCreditoPago, etsCreditoParcelado), "TCO_STATUS_OBRIGACAO_PARCELA", "TCO_COD_OBRIGACAO_PARCELA =" & CodPagamentoParcela
    Sql = "SELECT TOC_TPA_COD_PARCELAMENTO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO =" & CodPagamentoParcela
    If Bdados.AbreTabela(Sql, Rs) Then
        CodParcelamento = Nvl("" & Rs!TOC_TPA_COD_PARCELAMENTO, 0)
        If CodParcelamento = 0 Then Exit Sub
        Sql = "Select COUNT(*)  as Status   from TAB_COTAS_PARCELAMENTO WHERE " & _
                " TCP_TPA_COD_PARCELAMENTO= " & CodParcelamento & " AND TCP_STATUS_OBRIGACAO_PARCELA <> 3"
        If Bdados.AbreTabela(Sql, Rs) Then
            If Rs!Status = 0 Then
                Bdados.AtualizaDados "TAB_PARCELAMENTO", Bdados.PreparaValor(2), "TPA_STATUS_PARCELAMENTO", "TPA_NUM_PARCELAMENTO =" & CodParcelamento
            End If
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Public Function BuscaDam(Documento As String) As VSRecordset
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Condicao As String
    
    
    If Len(Documento) > 12 Then
        Dim Inscricao As String
        Dim Periodo As String
        Dim CodTributo As String
        If CInt(Mid(Documento, 6, 1)) = 1 Then  'Tipo Inscricao
            'IM
            Inscricao = Mid(Documento, 7, 13)
            Inscricao = Imposto.FormataInscricao(Left(Inscricao, 10), InscContrib)
            'Periodo
            Periodo = Mid(Documento, 20, 6)
            If CDbl(Right(Periodo, 2)) = 0 Then Periodo = Left(Periodo, 4)
        Else
            If AplicacoesVTFuncoes.Municipio = "GRAJAÚ" Then
                'IC
                Inscricao = Right(Documento, 14)
                'Periodo
                Periodo = Left(Documento, 4)
                CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
            Else
                'IC
                Inscricao = Mid(Documento, 7, 15)
                'Periodo
                Periodo = Mid(Documento, 22, 4)
                
                'CORRELATIVO
                If CDbl(Mid(Documento, 2, 2)) <> 11 Then 'BLOCO TEMPORARIO DEVIDO CORR. DO IPTU- TIRAR EM 30/09/2004
                    CodTributo = Imposto.BuscaCodImposto("", Mid(Documento, 2, 2))
                Else
                    CodTributo = Imposto.BuscaCodImposto("", "34")
                End If
            End If
        End If
        
    
        Sql = "Select toc_inscricao as Inscricao,toc_tip_cod_imposto as Imposto, toc_periodo as Periodo," & _
                "toc_data_vencimento as Vencimento, TOC_PARCELA as Parcela, TOC_VALOR_OBRIGACAO as ValorTributo," & _
                "toc_valor_juros as Juros, toc_valor_multa as Multa, TOC_TOTAL_TAXA_INCLUSA as Taxa,toc_cod_obrigacao as Documento " & _
                ", toc_tipo_inscricao as TipoInscricao,TOC_PARCELA as Parcela,0 as Sinfaz from tab_obrigacao_contribuinte where toc_inscricao ='" & _
                Inscricao & "' and toc_tip_cod_imposto ='" & CodTributo & _
                "' and toc_periodo = " & Periodo
        If Not Bdados.AbreTabela(Sql, Rs) Then
            Sql = "Select tgt_data_geracao,TGT_tim_ic, tgt_tip_cod_imposto_original, tgt_inscricao as Inscricao,tgt_tip_cod_imposto as Imposto, tgt_periodo as Periodo," & _
                "tgt_data_vencimento as Vencimento, tgt_parcela as Parcela, tgt_valor_tributo as ValorTributo," & _
                "tgt_valor_juros as Juros, tgt_valor_multa as Multa, tgt_taxa_expediente as Taxa, tgt_cod_pagamento as Documento,tgt_tipo " & _
                ", tgt_tipo_inscricao as TipoInscricao, 0 as Parcela,0 as Sinfaz from tab_geracao_tributo where tgt_inscricao ='" & _
                Inscricao & "' and tgt_tip_cod_imposto ='" & CodTributo & _
                "' and tgt_periodo = " & Periodo
            Call Bdados.AbreTabela(Sql, Rs)
            'BUSCA PARCELAS
        End If
    Else
        Sql = "Select toc_inscricao as Inscricao,toc_tip_cod_imposto as Imposto, toc_periodo as Periodo," & _
                "toc_data_vencimento as Vencimento, TOC_PARCELA as Parcela, TOC_VALOR_OBRIGACAO as ValorTributo," & _
                "toc_valor_juros as Juros, toc_valor_multa as Multa, TOC_TOTAL_TAXA_INCLUSA as Taxa,toc_cod_obrigacao as Documento " & _
                ", toc_tipo_inscricao as TipoInscricao,TOC_PARCELA as Parcela,0 as Sinfaz from tab_obrigacao_contribuinte where toc_cod_obrigacao=" & Documento
        
        If Not Bdados.AbreTabela(Sql, Rs) Then
            Sql = "Select tgt_data_geracao,TGT_tim_ic, tgt_tip_cod_imposto_original, tgt_inscricao as Inscricao,tgt_tip_cod_imposto as Imposto, tgt_periodo as Periodo," & _
                "tgt_data_vencimento as Vencimento, tgt_parcela as Parcela, tgt_valor_tributo as ValorTributo," & _
                "tgt_valor_juros as Juros, tgt_valor_multa as Multa, tgt_taxa_expediente as Taxa, tgt_cod_pagamento as Documento,tgt_tipo " & _
                ", tgt_tipo_inscricao as TipoInscricao, 0 as Parcela,0 as Sinfaz from tab_geracao_tributo where tgt_cod_pagamento=" & Documento
            If Not Bdados.AbreTabela(Sql, Rs) Then
                Sql = "Select tco_inscricao as Inscricao,tco_tip_cod_imposto as Imposto, tco_periodo as Periodo," & _
                "tco_data_vencimento as Vencimento, TCo_NUM_PARCELA as Parcela, TCo_VALOR_PARCELA as ValorTributo," & _
                "TCo_VALOR_JUROS as Juros, tco_valor_multa as Multa, 0 as Taxa,TCO_COD_OBRIGACAO_PARCELA as Documento " & _
                ", tco_tipo_inscricao as TipoInscricao,0 as Sinfaz from tab_cotas_obrigacao where TCO_COD_OBRIGACAO_PARCELA=" & Documento
                If Not Bdados.AbreTabela(Sql, Rs) Then
                    Sql = "Select tcp_inscricao as Inscricao,tcp_tip_cod_imposto as Imposto, tcp_periodo as Periodo," & _
                    "tcp_data_vencimento as Vencimento, TCP_NUM_PARCELA as Parcela, TCP_VALOR_PARCELA as ValorTributo," & _
                    "TCP_VALOR_JUROS as Juros, tcp_valor_multa as Multa, 0 as Taxa,tcp_num_cota as Documento " & _
                    ", tcp_tipo_inscricao as TipoInscricao,0 as Sinfaz from tab_cotas_parcelamento where tcp_num_cota=" & Documento
                    If Not Bdados.AbreTabela(Sql, Rs) Then
                        If Len(Documento) > 2 Then
                            Sql = "SELECT TCS_INSCRICAO AS INSCRICAO,TCS_PERIODO AS PERIODO,TCS_TIP_COD_IMPOSTO AS IMPOSTO, TCS_COD_MODULO AS " & _
                                " MODULO,1 as Sinfaz FROM TAB_COD_BAIXA_SINFAZ WHERE TCS_SEQUENCIAL = " & CDbl(Left(Documento, Len(Documento) - 2))
                        End If
                        Call Bdados.AbreTabela(Sql, Rs)
                    End If
                End If
            End If
        End If
    End If
    Set BuscaDam = Rs
End Function


Public Sub AtualizaMultaJuros(Inscricao As String, Data As Date)
    Dim Sql As String
    Dim Rs As VSRecordset
    
    On Error GoTo Trata
    Sql = "SELECT * FROM TAB_CONTA_CONTRIBUINTE WHERE TCC_INSCRICAO='" & Inscricao & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        Do While Not Rs.EOF
            CalculaValoresJuros Rs!tcc_codigo_conta, Data, Rs!TCC_DATA_VENCIMENTO
            CalculaValoresMulta Rs!tcc_codigo_conta, Data, Rs!TCC_DATA_VENCIMENTO
            Rs.MoveNext
        Loop
    End If
    Bdados.FechaTabela Rs
    
    Exit Sub
Trata:
    Avisa Err.Description
End Sub
Public Function BuscaNaturezaContribuinte(Inscicao As String) As String
    Dim Sql As String
    'Função que checa se o valor da variaval incricao corresponde a uma incricao municipal ou cadastral...
    Sql = "Select tci_tnj_cod_natureza  from tab_contribuinte where tci_im = '" & Inscicao & "'"
    If Bdados.AbreTabela(Sql) Then
        BuscaNaturezaContribuinte = "" & Bdados.Tabela(0)
    End If
End Function

'===================================================Andre Domingo===========================================================================================


Public Sub AtualizaContaContribuinte(CodConta As String, DataMovimento As Date, ValorAtualizado As Double, _
            Operacao As TipoOperacaoConta, Optional TDebito As TipoDebitar, Optional MultaAcumulada As Double)
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Dim ImpostoAtual As Double
    Dim MultaAtual As Double
    Dim JuroAtual As Double
    Dim CorrecaoAtual As Double
    Dim SaldoAtual As Double
    Dim CreditoAtual As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Campo As String
    Dim DebitoAtual As Double
    Dim ValorPago As Double
    Dim Desconto As Double
    Dim NaoTributadas As Double
    Dim TotalPagoAcumulado As Double
    Dim CodImposto As String
    Dim AnoImposto As String
    Dim rsParametroImposto As VSRecordset
    Dim SqlParametroImposto As String
    Dim CampoPagouAMais As String
    Dim CampoPagouAMenos As String

    On Error GoTo Trata
    Condicao = "tcc_codigo_conta =" & CodConta
    If Operacao = EtoCredito Then
        Sql = "SELECT * FROM tab_Conta_contribuinte WHERE " & Condicao
        If Bdados.AbreTabela(Sql, Rs) Then
            ValorPago = ValorAtualizado
            ImpostoAtual = Rs!tcc_imposto_atual
            JuroAtual = Rs!tcc_juros_atual
            MultaAtual = Rs!tcc_multa_atual
            CorrecaoAtual = Nvl("" & Rs!tcc_correcao_monetaria, 0)
            SaldoAtual = Rs!tcc_saldo_atual
            CodImposto = Rs!tcc_tip_cod_imposto
            AnoImposto = Left(CStr(Rs!tcc_periodo), 4)

            If Rs!tcc_multa_atual <= ValorAtualizado Then 'Abate 1º a multa
                MultaAtual = 0
                ValorAtualizado = Format(ValorAtualizado - Rs!tcc_multa_atual, Const_Monetario)
                If Rs!tcc_juros_atual <= ValorAtualizado Then 'Tenta abater juros tambem
                    JuroAtual = 0
                    ValorAtualizado = Format(ValorAtualizado - Rs!tcc_juros_atual, Const_Monetario)
                    If CDbl(Nvl("" & Rs!tcc_correcao_monetaria, 0)) <= ValorAtualizado Then 'Tenta abater correcao
                        CorrecaoAtual = 0
                        If Rs!tcc_imposto_atual <= ValorAtualizado Then 'Tenta abater imposto tambem (Zera débito)
                            ImpostoAtual = 0
                            ValorAtualizado = Format(ValorAtualizado - Rs!tcc_imposto_atual, Const_Monetario)
    '                         VOLTAR ESSA LINHA ABAIXO APOS CRIAR PARAMETRO
                             Obrig.TrocaSitObrigacao CodConta, etsCreditoPago

                        Else
                            ImpostoAtual = Rs!tcc_imposto_atual - ValorAtualizado
                            ValorAtualizado = 0
                        End If
                    Else 'Abate so Correcao
                        CorrecaoAtual = Format(ValorAtualizado - CDbl(Nvl("" & Rs!tcc_correcao_monetaria, 0)), Const_Monetario)
                        ValorAtualizado = 0
                    End If
                Else 'Abate só juros e multa
                    JuroAtual = Rs!tcc_juros_atual - ValorAtualizado
                    ValorAtualizado = 0
                End If
            Else 'Abate só juros
                MultaAtual = Rs!tcc_multa_atual - ValorAtualizado
                ValorAtualizado = 0
            End If
        End If
'       RETIRAR ESSA LINHA ABAIXO APOS CRIAR PARAMETRO

        Obrig.TrocaSitObrigacao CodConta, etsCreditoPago
        Sql = "SELECT tcc_parcela,TCC_DESCONTO_CONCEDIDO,TCC_NAO_TRIBUTADA,tcc_tdr_valor_real_pago FROM tab_Conta_contribuinte WHERE " & Condicao
        If Bdados.AbreTabela(Sql, Rs) Then
            SqlParametroImposto = "SELECT TPI_RECEITA_A_MAIS, TPI_RECEITA_A_MENOS FROM TAB_PARAMETRO_IMPOSTO" _
                & " WHERE tpi_tip_cod_imposto = '" & CodImposto & "' AND tpi_ano_imposto = " & Val(AnoImposto)
            Bdados.AbreTabela SqlParametroImposto, rsParametroImposto

            SaldoAtual = Format(ImpostoAtual + MultaAtual + JuroAtual + CorrecaoAtual, Const_Monetario)
            'PAGOU A MENOS
            If rsParametroImposto.EOF = False Then
                If IsNull(rsParametroImposto!TPI_RECEITA_A_MENOS) = False Then
                    If rsParametroImposto!TPI_RECEITA_A_MENOS = 2 Then 'SALDO DEVEDOR
                        SaldoAtual = Format(CDbl(Nvl("" & Rs!TCC_DESCONTO_CONCEDIDO, 0)) + SaldoAtual, Const_Monetario)
                    Else
                        Desconto = Format(CDbl(Nvl("" & Rs!TCC_DESCONTO_CONCEDIDO, 0)) + SaldoAtual, Const_Monetario)
                    End If
                Else
                    Desconto = Format(CDbl(Nvl("" & Rs!TCC_DESCONTO_CONCEDIDO, 0)) + SaldoAtual, Const_Monetario)
                End If
            Else
                Desconto = Format(CDbl(Nvl("" & Rs!TCC_DESCONTO_CONCEDIDO, 0)) + SaldoAtual, Const_Monetario)
            End If

            If ValorAtualizado > 0 Then 'PAGOU A MAIS
                'PAGOU A MAIS
                If rsParametroImposto.EOF = False Then
                    If IsNull(rsParametroImposto!TPI_RECEITA_A_MENOS) = False Then
                        If rsParametroImposto!TPI_RECEITA_A_MAIS = 2 Then 'SALDO CREDOR
                            CreditoAtual = Format(CDbl(Nvl("" & Rs!TCC_NAO_TRIBUTADA, 0)) + ValorAtualizado, Const_Monetario)
                        Else
                            NaoTributadas = Format(CDbl(Nvl("" & Rs!TCC_NAO_TRIBUTADA, 0)) + ValorAtualizado, Const_Monetario)
                        End If
                    Else
                        NaoTributadas = Format(CDbl(Nvl("" & Rs!TCC_NAO_TRIBUTADA, 0)) + ValorAtualizado, Const_Monetario)
                    End If
                Else
                    NaoTributadas = Format(CDbl(Nvl("" & Rs!TCC_NAO_TRIBUTADA, 0)) + ValorAtualizado, Const_Monetario)
                End If
            End If

            TotalPagoAcumulado = Format(ValorPago + CDbl(Nvl("" & Rs!tcc_tdr_valor_real_pago, 0)), Const_Monetario)
            Campos = "tcc_imposto_atual,tcc_juros_atual,tcc_multa_atual,tcc_saldo_atual," _
            & " tcc_credito_atual,tcc_tipo_ultima_transacao,TCC_DESCONTO_CONCEDIDO,TCC_NAO_TRIBUTADA,tcc_tdr_valor_real_pago"

            Valores = Bdados.PreparaValor(Bdados.Converte(Format(ImpostoAtual, Const_Monetario), TCDuplo), Bdados.Converte(Format(JuroAtual, Const_Monetario), TCDuplo), Bdados.Converte(Format(MultaAtual, Const_Monetario), TCDuplo), Bdados.Converte(Format(SaldoAtual, Const_Monetario), TCDuplo), _
            Bdados.Converte(CreditoAtual, TCDuplo), Operacao, Bdados.Converte(Desconto, TCDuplo), Bdados.Converte(NaoTributadas, TCDuplo), Bdados.Converte(TotalPagoAcumulado, TCDuplo))
            If SaldoAtual <= 0 Then
                Campos = Campos & ",tcc_status_conta"
                'VERIFICAR AQUI SE É UMA CONTA PARCELADA
                Valores = Valores & Bdados.PreparaValor(etaFechada)
            End If
        End If
    Else 'debito
        
        Campos = IIf(TDebito = EtdeImposto, "tcc_imposto_atual", IIf(TDebito = EtdeMulta, _
            "tcc_multa_atual", "tcc_juros_atual"))
        If TDebito = EtdeCorrecao Then Campos = "tcc_correcao_monetaria"
        If TDebito = EtdeDescontos Then
            Campos = "TCC_DESCONTO_CONCEDIDO"
            Desconto = ValorAtualizado
        End If
        Valores = Bdados.PreparaValor(Bdados.Converte(Format(ValorAtualizado, Const_Monetario), TCDuplo))
        
        DebitoAtual = Format(ValorAtualizado, Const_Monetario)
        Sql = "SELECT tcc_parcela,tcc_imposto_original FROM tab_Conta_contribuinte WHERE " & Condicao
        If Bdados.AbreTabela(Sql, Rs) Then
            Campos = Campos & ",tcc_saldo_atual,tcc_credito_atual,tcc_tipo_ultima_transacao"
            SaldoAtual = Format(IIf(Rs!tcc_imposto_original + JurosConta + MultaConta + CorrecaoConta - Desconto < 0, 0, Rs!tcc_imposto_original + JurosConta + MultaConta + CorrecaoConta - Desconto), Const_Monetario)
'            Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Format(DebitoAtual, Const_Monetario), TCDuplo))
            If TDebito = EtdeMulta Then
                Campos = Campos & ",tcc_multa_acumulada"
                Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Format(MultaAcumulada, Const_Monetario), TCDuplo))
            End If
            Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Format(SaldoAtual, Const_Monetario), TCDuplo), Bdados.Converte(Format(CreditoAtual, Const_Monetario), TCDuplo), Operacao)
        End If
    End If
    Valores = Valores & Bdados.PreparaValor(Bdados.Converte(DataMovimento, TCDataHora))
    Campos = Campos & IIf(Trim(Campos) = "", "tcc_data_movimento", ",tcc_data_movimento")
    Bdados.AtualizaDados "Tab_Conta_Contribuinte", Valores, Campos, Condicao
    Bdados.FechaTabela Rs
    Exit Sub
Trata:
    Avisa Err.Description
    Exit Sub
    Resume
End Sub

Public Function MovimentaContaContribuinte(ByVal CodigoConta As String, _
                Optional DataMovimentacao As String, Optional Obrig As Obrigacao, _
                Optional TipoTransacao As TipoTransacao = ettAvista, _
                Optional DescontoIndividualTributo As Double = 0, Optional DescontoIndividualAtualizacao As Double = 0, _
                Optional DescontoIndividualMulta As Double = 0, Optional DescontoIndividualJuros As Double = 0) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DataVencimento  As String
    Dim DataMovimento  As String
    Dim Venc As String
    Dim Pag As String
    Dim MovimentaConta As Boolean
    Dim DescontoCorrecao As Double
    Dim DescontoJuros As Double
    Dim DescontoMulta As Double
    Dim DescontoOriginal As Double
    
    On Error GoTo Trata
    If Not Obrig Is Nothing Then
        If Obrig.obStatusObrigacao = 3 Then Exit Function
    Else
        If StatusConta(CodigoConta) = etaFechada Then
            Exit Function
        End If
    End If
    DataVencimento = ""
    DataMovimento = ""
    If Obrig Is Nothing Then
        DataVencimento = BuscaDataVence(CodigoConta)
    Else
        DataVencimento = Obrig.obDataVencimento
    End If
    If Trim(DataMovimentacao) = "" Then
        DataMovimento = BuscaDataPago(CodigoConta)
    End If
    Venc = Right(DataVencimento, 4) & Mid(DataVencimento, 4, 2) & Left(DataVencimento, 2)
    DataMovimento = IIf(Trim(DataMovimento) = "", Date, DataMovimento)
    Pag = Right(DataMovimento, 4) & Mid(DataMovimento, 4, 2) & Left(DataMovimento, 2)
    If Trim(Venc) <> "" Then
        DescontoCorrecao = DescontoIndividualAtualizacao
        DescontoJuros = DescontoIndividualJuros
        DescontoMulta = DescontoIndividualMulta
        
        CalculaValoresCorrecao Obrig, DataMovimentacao, TipoTransacao, DescontoCorrecao
        
        Call CalculaValoresJuros(CodigoConta, DataMovimentacao, DataVencimento, Obrig, TipoTransacao, DescontoJuros)
        Call CalculaValoresMulta(CodigoConta, DataMovimentacao, DataVencimento, Obrig, TipoTransacao, DescontoMulta)
        CalculaValoresDescontos Obrig, DescontoCorrecao + DescontoJuros + DescontoMulta, DataMovimentacao, Obrig.obCodigoObrigacao, DescontoIndividualTributo, TipoTransacao
    Else
        CriaContaContribuinte CodigoConta, , Obrig
 '       MovimentaContaContribuinte CodigoConta, DataMovimentacao
    End If
    JurosConta = 0
    MultaConta = 0
    CorrecaoConta = 0
    Exit Function
Trata:
    Avisa Err.Description
    Exit Function
    Resume
End Function

Public Function CalculaValoresDescontos(Obrig As Obrigacao, DescontosAcumuladosIncrementos As Double, _
                Optional DataPagamento As String, Optional CodObrigacao As String, Optional DescontoIndividualTributo As Double = 0, Optional TipoTrans As TipoTransacao = ettAvista) As Boolean
    Dim Desconto As Double
    Dim PercDesc As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    
    If CodObrigacao <> "" Then Obrig.BuscaDetalheObrigacao CodObrigacao, , TipoTrans
        
    If DescontoIndividualTributo = 0 Then
        Sql = "select TPI_DESCONTO_TRIBUTO_PARC "
        Sql = Sql & " from TAB_PARAMETRO_IMPOSTO_REFIS where TPI_TIP_COD_IMPOSTO ='" & Obrig.obCodImposto & "'"
        Sql = Sql & " and tpi_ano_imposto= '" & Imposto.BuscaAnoImposto(Obrig.obCodImposto, Left(Obrig.obPeriodo, 4)) & "'"
        Sql = Sql & " and " & Obrig.obParcela & " BETWEEN TPI_LIMITE_INFERIOR AND TPI_LIMITE_SUPERIOR "
        If Bdados.AbreTabela(Sql, Rs) Then
            Obrig.obValorDesconto = Rs!TPI_DESCONTO_TRIBUTO_PARC
        End If
        Desconto = Format(Obrig.obValorObrigacao * (Obrig.obValorDesconto / 100) + DescontosAcumuladosIncrementos, Const_Monetario)
    Else
        Desconto = Format(Obrig.obValorObrigacao * (DescontoIndividualTributo / 100) + DescontosAcumuladosIncrementos, Const_Monetario)
    End If
    If IsDate(DataPagamento) Then
        If Obrig.obValorObrigacao > 0 Then
            Call AtualizaContaContribuinte(Obrig.obCodigoObrigacao, CDate(DataPagamento), Desconto, EtoDebito, EtdeDescontos)
        ElseIf Obrig.obValorObrigacao = 0 Then
            Call AtualizaContaContribuinte(Obrig.obCodigoObrigacao, CDate(DataPagamento), 0, EtoDebito, EtdeDescontos)
        End If
    End If
End Function

Public Function CriaContaContribuinte(ByRef CodigoObrigacao As String, Optional SubstituiExistente As Boolean = True, Optional Obrig As Obrigacao, Optional Modo As TipoInscricaoObrigacao) As Boolean
    'FUNCAO OK - OBRIGACAO
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim Taxas As Double
    
    Dim Rs As VSRecordset
    On Error GoTo Trata
    CriaContaContribuinte = False
    If Len(CodigoObrigacao) = 0 Then Exit Function
    Sql = "Select tcc_codigo_conta from tab_conta_contribuinte where tcc_codigo_conta =" & CodigoObrigacao
    If Bdados.AbreTabela(Sql, Rs) Then
        Bdados.FechaTabela Rs
        If Not SubstituiExistente Then
            CriaContaContribuinte = False
            Exit Function
        End If
    End If
    
    If Obrig Is Nothing Then
        Set Obrig = New Obrigacao
        Obrig.BuscaDetalheObrigacao CodigoObrigacao, Modo
    End If
    If Obrig.obCodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) Then
        Taxas = CDbl(Edita.TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXEA"), 0), ".", ",")) + CDbl(Edita.TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXP"), 0), ".", ","))
    Else
        Taxas = 0
    End If
    Campos = "TCC_INSCRICAO,tcc_im,tcc_num_conta, tcc_codigo_conta,tcc_tip_cod_imposto," & _
                "tcc_periodo,tcc_imposto_original,tcc_juros_original,tcc_multa_original," & _
                "tcc_data_vencimento,tcc_data_movimento,tcc_imposto_atual,tcc_multa_atual," & _
                "tcc_juros_atual,tcc_correcao_monetaria,tcc_saldo_atual,tcc_credito_atual,tcc_status_conta," & _
                "tcc_tipo_ultima_transacao,tcc_multa_acumulada," & _
                "tcc_parcela,tcc_saldo_devedor,tcc_tdr_valor_real_pago,Tcc_TPA_COD_PARCELAMENTO,TCC_TIPO_INSCRICAO"
'    If Obrig.obStatusObrigacao = etsCreditoNaoLancado Then Exit Function
'    If Obrig.obValorObrigacao = 0 Then
 '       Util.Avisa "Oi"
  '  End If
    
    If Trim(Obrig.obContribuinte) = "" Then Exit Function
    Valores = Bdados.PreparaValor(Bdados.Converte(Trim(Obrig.obContribuinte), tctexto), Bdados.Converte(Trim(Obrig.obInscMunicipal), tctexto), _
                        Obrig.obStatusObrigacao, CodigoObrigacao, _
                        Obrig.obCodImposto, Obrig.obPeriodo, Bdados.Converte(Format(Obrig.obValorObrigacao + Taxas, Const_Monetario), TCDuplo), _
                        Bdados.Converte(Format(Obrig.obValorMulta, Const_Monetario), TCDuplo), Bdados.Converte(Format(Obrig.obValorJuros, _
                        Const_Monetario), TCDuplo), Bdados.Converte(Obrig.obDataVencimento, TCDataHora), Bdados.Converte(Date, TCDataHora), _
                        Bdados.Converte(Obrig.obValorObrigacao, TCDuplo), Bdados.Converte(Obrig.obValorMulta, TCDuplo), _
                        Bdados.Converte(Obrig.obValorJuros, TCDuplo), 0, Bdados.Converte(Format(Obrig.obValorObrigacao + _
                        Obrig.obValorMulta + Obrig.obValorJuros + Taxas, Const_Monetario), TCDuplo), 0, Obrig.obStatusObrigacao, 2, 0, Obrig.obParcela, 0, 0, Obrig.obNumParcelamento, Obrig.obTipoInscricao)


    CriaContaContribuinte = Bdados.GravaDados("Tab_Conta_Contribuinte", Valores, Campos, "tcc_codigo_conta =" & Obrig.obCodigoObrigacao)

    Exit Function
Trata:
    Avisa Err.Description
    Exit Function
    Resume
End Function

Public Function CalculaValoresCorrecaoRelacional(Obrig As Obrigacao, Optional DataPagamento As String, _
            Optional TipoTransacao As TipoTransacao = ettAvista, Optional ResultadoDesconto As Double) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DiasPassados As Double
    Dim CorrecaoMonetaria As Double
    Dim Condicao As String
    Dim Valores As String
    Dim Pag As String
    Dim Venc As String
    Dim Cotas As Integer
    Dim ValorCota As Double
    Dim ClsImposto As New VSImposto
    Dim RsAux As VSRecordset
    Dim ValorDescCorrecaoAvista As Double
    Dim ValorDescCorrecaoParc As Double
    Dim IndiceEpocaLancamento As Double
    Dim IndiceAtual As Double
    Dim AutorizaCorrecao As Integer
    
    On Error GoTo Trata
    CalculaValoresCorrecaoRelacional = True
    If Trim(DataPagamento) = "" Then DataPagamento = Format(Date, "dd/mm/yyyy")
    Sql = "select tpi_autorizacao_correcao,TPI_DESCONTO_ATUALIZACAO,TPI_DESCONTO_ATUALIZACAO_PARC,TPI_PARAMETRO_REFIS from tab_parametro_imposto where tpi_tip_cod_imposto ='" & Obrig.obCodImposto & "'"
    Sql = Sql & " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Obrig.obCodImposto, Left(Obrig.obPeriodo, 4)) & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        AutorizaCorrecao = CDbl(Nvl("" & Rs!tpi_autorizacao_correcao, 0))
        If Nvl("" & Rs!TPI_PARAMETRO_REFIS, 1) = 2 Then
            Sql = "select TPI_DESCONTO_TRIBUTO_PARC,TPI_DESCONTO_JUROS_PARC,TPI_DESCONTO_MULTA_PARC,TPI_DESCONTO_ATUALIZACAO_PARC "
            Sql = Sql & " from TAB_PARAMETRO_IMPOSTO_REFIS where TPI_TIP_COD_IMPOSTO ='" & Obrig.obCodImposto & "'"
            Sql = Sql & " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Obrig.obCodImposto, Left(Obrig.obPeriodo, 4)) & "'"
            If Not Bdados.AbreTabela(Sql, Rs) Then
                Exit Function
            End If
        End If
    End If
    If Not Rs.EOF Then
        If AutorizaCorrecao <> 1 And TipoTransacao <> ettDividaAtiva Then
            ResultadoDesconto = 0
            Exit Function
        End If
        ValorDescCorrecaoAvista = CDbl(Nvl("" & Rs!TPI_DESCONTO_ATUALIZACAO, 0))
        ValorDescCorrecaoParc = CDbl(Nvl("" & Rs!TPI_DESCONTO_ATUALIZACAO_PARC, 0))
    Else
        ResultadoDesconto = 0
        Exit Function
    End If
    Sql = "Select tmo_valor from TAB_MONETARIA where TMO_DATA =(SELECT MAX(TMO_DATA) FROM " & _
            "TAB_MONETARIA WHERE TMO_DATA <=" & Bdados.Converte(Obrig.obDataVencimento, TCDataHora) & ")"
    If Bdados.AbreTabela(Sql, RsAux) Then
        IndiceEpocaLancamento = RsAux!tmo_valor
        Sql = "Select tmo_valor from TAB_MONETARIA where TMO_DATA =(SELECT MAX(TMO_DATA) FROM " & _
            "TAB_MONETARIA WHERE TMO_DATA <=" & Bdados.Converte(DataPagamento, TCDataHora) & ")"
        If Bdados.AbreTabela(Sql, RsAux) Then
            IndiceAtual = RsAux!tmo_valor
        Else
            CalculaValoresCorrecaoRelacional = False
            ResultadoDesconto = 0
            Exit Function
        End If
    Else
        ResultadoDesconto = 0
        CalculaValoresCorrecaoRelacional = False
        Exit Function
    End If
    
    CorrecaoMonetaria = ((IndiceAtual / IndiceEpocaLancamento) - 1)
    CorrecaoMonetaria = CorrecaoMonetaria * Obrig.obValorObrigacao
    If ResultadoDesconto = 0 Then
        If TipoTransacao <> ettDividaAtiva Then
            If TipoTransacao = ettAvista Then
                ResultadoDesconto = CorrecaoMonetaria * (ValorDescCorrecaoAvista / 100)
            Else
                ResultadoDesconto = CorrecaoMonetaria * (ValorDescCorrecaoParc / 100)
            End If
        Else
            ResultadoDesconto = 0
        End If
    Else
        ResultadoDesconto = CorrecaoMonetaria * (ResultadoDesconto / 100)
    End If
    CorrecaoConta = Format(CorrecaoMonetaria, Const_Monetario)
    Call AtualizaContaContribuinte(Obrig.obCodigoObrigacao, CDate(DataPagamento), CorrecaoMonetaria, EtoDebito, EtdeCorrecao)
    Bdados.FechaTabela Rs
    Exit Function
Trata:
        Avisa Err.Description
        CalculaValoresCorrecaoRelacional = False
        Exit Function
        Resume
End Function

Public Function CalculaValoresCorrecao(Obrig As Obrigacao, Optional DataPagamento As String, _
            Optional TipoTransacao As TipoTransacao = ettAvista, Optional ResultadoDesconto As Double) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DiasPassados As Double
    Dim CorrecaoMonetaria As Double
    Dim Condicao As String
    Dim Valores As String
    Dim Pag As String
    Dim Venc As String
    Dim Cotas As Integer
    Dim ValorCota As Double
    Dim ClsImposto As New VSImposto
    Dim RsAux As VSRecordset
    Dim ValorDescCorrecaoAvista As Double
    Dim ValorDescCorrecaoParc As Double
    Dim IndiceEpocaLancamento As Double
    Dim IndiceAtual As Double
    Dim AutorizaCorrecao As Integer
    
    On Error GoTo Trata
    CalculaValoresCorrecao = True
    If Trim(DataPagamento) = "" Then DataPagamento = Format(Date, "dd/mm/yyyy")
    Sql = "select tpi_autorizacao_correcao,TPI_DESCONTO_ATUALIZACAO,TPI_DESCONTO_ATUALIZACAO_PARC,TPI_PARAMETRO_REFIS from tab_parametro_imposto where tpi_tip_cod_imposto ='" & Obrig.obCodImposto & "'"
    Sql = Sql & " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Obrig.obCodImposto, Left(Obrig.obPeriodo, 4)) & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        AutorizaCorrecao = CDbl(Nvl("" & Rs!tpi_autorizacao_correcao, 0))
        ValorDescCorrecaoAvista = CDbl(Nvl("" & Rs!TPI_DESCONTO_ATUALIZACAO, 0))
        ValorDescCorrecaoParc = CDbl(Nvl("" & Rs!TPI_DESCONTO_ATUALIZACAO_PARC, 0))
        If Nvl("" & Rs!TPI_PARAMETRO_REFIS, 1) = 2 Then
            Sql = "select TPI_DESCONTO_TRIBUTO_PARC,TPI_DESCONTO_JUROS_PARC,TPI_DESCONTO_MULTA_PARC,TPI_DESCONTO_ATUALIZACAO_PARC "
            Sql = Sql & " from TAB_PARAMETRO_IMPOSTO_REFIS where TPI_TIP_COD_IMPOSTO ='" & Obrig.obCodImposto & "'"
            Sql = Sql & " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Obrig.obCodImposto, Left(Obrig.obPeriodo, 4)) & "'"
            Sql = Sql & " and " & Obrig.obParcela & " BETWEEN TPI_LIMITE_INFERIOR AND TPI_LIMITE_SUPERIOR "
            If Not Bdados.AbreTabela(Sql, Rs) Then
                Exit Function
            Else
                If Obrig.obParcela = 0 Then ValorDescCorrecaoAvista = CDbl(Nvl("" & Rs!TPI_DESCONTO_ATUALIZACAO_PARC, 0))
                ValorDescCorrecaoParc = CDbl(Nvl("" & Rs!TPI_DESCONTO_ATUALIZACAO_PARC, 0))
            End If
        End If
    End If
    If Not Rs.EOF Then
        If AutorizaCorrecao <> 1 And TipoTransacao <> ettDividaAtiva Then
            If AutorizaCorrecao = 3 Then
                CalculaValoresCorrecaoRelacional Obrig, DataPagamento, _
                    TipoTransacao, ResultadoDesconto
            Else
                ResultadoDesconto = 0
                Exit Function
            End If
        End If
        
    Else
        ResultadoDesconto = 0
        Exit Function
    End If
    Sql = "Select tmo_valor from TAB_MONETARIA where YEAR(TMO_DATA) = " & Year(Obrig.obDataVencimento) & _
            " and MONTH(TMO_DATA) = " & Month(Obrig.obDataVencimento) & " order by tmo_data asc"
    If Bdados.AbreTabela(Sql, RsAux) Then
        IndiceEpocaLancamento = RsAux!tmo_valor
'        Sql = "Select tmo_valor from TAB_MONETARIA where YEAR(TMO_DATA) = " & Year(DataPagamento)
'        Sql = Sql & " and MONTH(TMO_DATA) = " & Month(DataPagamento) & " order by tmo_data asc"
        
        Sql = "Select tmo_valor from TAB_MONETARIA where TMO_DATA = (Select MAX(TMO_DATA) from TAB_MONETARIA where TMO_DATA < = CONVERT(SMALLDATETIME,'30-08-2007',103))"
        If Bdados.AbreTabela(Sql, RsAux) Then
            IndiceAtual = RsAux!tmo_valor
        End If
        CorrecaoMonetaria = Obrig.obValorObrigacao * (IndiceEpocaLancamento / IndiceAtual) 'SÓ SERVE COM IPCA
        If CorrecaoMonetaria > 0 Then CorrecaoMonetaria = CorrecaoMonetaria - Obrig.obValorObrigacao
    Else
        ResultadoDesconto = 0
        CalculaValoresCorrecao = False
        Exit Function
    End If
    If ResultadoDesconto = 0 Then
        If TipoTransacao <> ettDividaAtiva Then
            If TipoTransacao = ettAvista Then
                ResultadoDesconto = CorrecaoMonetaria * (ValorDescCorrecaoAvista / 100)
            Else
                ResultadoDesconto = CorrecaoMonetaria * (ValorDescCorrecaoParc / 100)
            End If
        Else
            ResultadoDesconto = 0
        End If
    Else
        ResultadoDesconto = CorrecaoMonetaria * (ResultadoDesconto / 100)
    End If
    CorrecaoConta = Format(CorrecaoMonetaria, Const_Monetario)
    Call AtualizaContaContribuinte(Obrig.obCodigoObrigacao, CDate(DataPagamento), CorrecaoMonetaria, EtoDebito, EtdeCorrecao)
    Bdados.FechaTabela Rs
    Exit Function
Trata:
        Avisa Err.Description
        CalculaValoresCorrecao = False
        Exit Function
        Resume
        
End Function


Public Function CalculaValoresCorrecaoAvulso(CodImposto As String, Periodo As String, DataVencimento As String, _
            DataPagamento As String, ValorOriginal As Double, Optional TipoTransacao As _
            TipoTransacao = ettAvista, Optional ResultadoDesconto As Double) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DiasPassados As Double
    Dim CorrecaoMonetaria As Double
    Dim Condicao As String
    Dim Valores As String
    Dim ValorDescCorrecaoAvista As Double
    Dim ValorDescCorrecaoParc As Double
    Dim Pag As String
    Dim Venc As String
    Dim Cotas As Integer
    Dim ValorCota As Double
    Dim ClsImposto As New VSImposto
    Dim RsAux As VSRecordset

    Dim IndiceEpocaLancamento As Double
    Dim IndiceAtual As Double
    On Error GoTo Trata
    
    Sql = "select tpi_autorizacao_correcao,TPI_DESCONTO_ATUALIZACAO,TPI_DESCONTO_ATUALIZACAO_PARC  from tab_parametro_imposto where tpi_tip_cod_imposto ='" & CodImposto & "'"
    Sql = Sql & " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(CodImposto, Left(Periodo, 4)) & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If CDbl(Nvl("" & Rs!tpi_autorizacao_correcao, 0)) <> 1 And TipoTransacao <> ettDividaAtiva Then
            If CDbl(Nvl("" & Rs!tpi_autorizacao_correcao, 0)) = 3 Then
                CalculaValoresCorrecaoAvulsoRelacional CodImposto, Periodo, DataVencimento, _
                    DataPagamento, ValorOriginal, TipoTransacao, ResultadoDesconto
            Else
                Exit Function
            End If
        End If
        ValorDescCorrecaoAvista = CDbl(Nvl("" & Rs!TPI_DESCONTO_ATUALIZACAO, 0))
        ValorDescCorrecaoParc = CDbl(Nvl("" & Rs!TPI_DESCONTO_ATUALIZACAO_PARC, 0))
    Else
        Exit Function
    End If
    Obrig.obDataVencimento = IIf(Trim(Obrig.obDataVencimento) = "", Date, Obrig.obDataVencimento)
    Sql = "Select tmo_valor from TAB_MONETARIA where YEAR(TMO_DATA) = " & Year(DataVencimento) & _
            " and MONTH(TMO_DATA) = " & Month(DataVencimento) & " order by tmo_data asc"
    If Bdados.AbreTabela(Sql, RsAux) Then
        IndiceEpocaLancamento = RsAux!tmo_valor
        'Sql = "Select tmo_valor from TAB_MONETARIA where YEAR(TMO_DATA) = " & Year(DataPagamento)
        'Sql = Sql & " and MONTH(TMO_DATA) = " & Month(DataPagamento) & " order by tmo_data asc"
        Sql = "Select tmo_valor from TAB_MONETARIA where TMO_DATA = (Select MAX(TMO_DATA) from TAB_MONETARIA where TMO_DATA < = CONVERT(SMALLDATETIME,'30-08-2007',103))"
        If Bdados.AbreTabela(Sql, RsAux) Then
            IndiceAtual = RsAux!tmo_valor
        End If
        CorrecaoMonetaria = ValorOriginal * (IndiceEpocaLancamento / IndiceAtual)
        If CorrecaoMonetaria > 0 Then CorrecaoMonetaria = CorrecaoMonetaria - ValorOriginal
    Else
        ResultadoDesconto = 0
        CalculaValoresCorrecaoAvulso = False
        Exit Function
    End If
    If TipoTransacao <> ettDividaAtiva Then
        If TipoTransacao = ettAvista Then
            ResultadoDesconto = CorrecaoMonetaria * (ValorDescCorrecaoAvista / 100)
        Else
            ResultadoDesconto = CorrecaoMonetaria * (ValorDescCorrecaoParc / 100)
        End If
    Else
        ResultadoDesconto = 0
    End If
    CalculaValoresCorrecaoAvulso = Format(CorrecaoMonetaria, Const_Monetario)
    Bdados.FechaTabela Rs
    Exit Function
Trata:
        Avisa Err.Description
        Exit Function
        Resume
End Function

Public Function CalculaValoresCorrecaoAvulsoRelacional(CodImposto As String, Periodo As String, DataVencimento As String, _
            DataPagamento As String, ValorOriginal As Double, Optional TipoTransacao As _
            TipoTransacao = ettAvista, Optional ResultadoDesconto As Double) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim DiasPassados As Double
    Dim CorrecaoMonetaria As Double
    Dim Condicao As String
    Dim Valores As String
    Dim ValorDescCorrecaoAvista As Double
    Dim ValorDescCorrecaoParc As Double
    Dim Pag As String
    Dim Venc As String
    Dim Cotas As Integer
    Dim ValorCota As Double
    Dim ClsImposto As New VSImposto
    Dim RsAux As VSRecordset

    Dim IndiceEpocaLancamento As Double
    Dim IndiceAtual As Double
    On Error GoTo Trata
    
    Sql = "select tpi_autorizacao_correcao,TPI_DESCONTO_ATUALIZACAO,TPI_DESCONTO_ATUALIZACAO_PARC  from tab_parametro_imposto where tpi_tip_cod_imposto ='" & CodImposto & "'"
    Sql = Sql & " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(CodImposto, Left(Periodo, 4)) & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If CDbl(Nvl("" & Rs!tpi_autorizacao_correcao, 0)) <> 1 And TipoTransacao <> ettDividaAtiva Then Exit Function
        ValorDescCorrecaoAvista = CDbl(Nvl("" & Rs!TPI_DESCONTO_ATUALIZACAO, 0))
        ValorDescCorrecaoParc = CDbl(Nvl("" & Rs!TPI_DESCONTO_ATUALIZACAO_PARC, 0))
    Else
        Exit Function
    End If
    Sql = "Select tmo_valor from TAB_MONETARIA where TMO_DATA =(SELECT MAX(TMO_DATA) FROM " & _
            "TAB_MONETARIA WHERE TMO_DATA <=" & Bdados.Converte(DataVencimento, TCDataHora) & ")"
    If Bdados.AbreTabela(Sql, RsAux) Then
        IndiceEpocaLancamento = RsAux!tmo_valor
        'Sql = "Select tmo_valor from TAB_MONETARIA where TMO_DATA =(SELECT MAX(TMO_DATA) FROM " & _
            "TAB_MONETARIA WHERE TMO_DATA <=" & Bdados.Converte(DataPagamento, TCDataHora) & ")"
        Sql = "Select tmo_valor from TAB_MONETARIA where TMO_DATA = (Select MAX(TMO_DATA) from TAB_MONETARIA where TMO_DATA < = CONVERT(SMALLDATETIME,'30-08-2007',103))"
        If Bdados.AbreTabela(Sql, RsAux) Then
            IndiceAtual = RsAux!tmo_valor
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    CorrecaoMonetaria = ((IndiceEpocaLancamento / IndiceAtual) - 1)
    CorrecaoMonetaria = CorrecaoMonetaria * ValorOriginal
    If TipoTransacao <> ettDividaAtiva Then
        If TipoTransacao = ettAvista Then
            ResultadoDesconto = CorrecaoMonetaria * (ValorDescCorrecaoAvista / 100)
        Else
            ResultadoDesconto = CorrecaoMonetaria * (ValorDescCorrecaoParc / 100)
        End If
    Else
        ResultadoDesconto = 0
    End If
    CalculaValoresCorrecaoAvulsoRelacional = Format(CorrecaoMonetaria, Const_Monetario)
    Bdados.FechaTabela Rs
    Exit Function
Trata:
        Avisa Err.Description
        Exit Function
        Resume
End Function


Public Function Correlativo(Sistema As String, Operador As Integer, Optional Descricao As String, _
                            Optional Incrementar As Boolean = True, _
                            Optional Referencia As String, Optional Incremento As Integer = 1, _
                            Optional Inicio As Integer = 1, _
                            Optional Mascara As String = "999999") As String
    Dim Comando As VSComando
    
    Set Comando = New VSComando
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Comando.Texto Bdados, "cpseg.dbo.sp_num_correlativo", cmdStoredProcedure
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Comando.Texto Bdados, "SP_NUM_CORRELATIVO", cmdStoredProcedure
    End If
    Comando.setarParametro "NomeDoSistema", tipChar, parEntrada, 4, Sistema
    Comando.setarParametro "NumCorrelativo", tipInteger, parEntradaSaida, , Operador
    Comando.setarParametro "ReferenciaDoOperador", tipLongVarChar, parEntrada, 20, Referencia
    Comando.setarParametro "SeqDoOperador", tipInteger, parEntrada, , Inicio
    Comando.setarParametro "IncDaSequencia", tipInteger, parEntrada, , Incremento
    Comando.setarParametro "AutoInc", tipInteger, parEntrada, , Abs(Incrementar)
    Comando.setarParametro "DisplayFormat", tipChar, parEntrada, 6, Mascara
    Comando.setarParametro "DescrDoOperador", tipVarChar, parEntrada, 50, Descricao
    Comando.Executa
    Correlativo = Comando.Parametro("NumCorrelativo").Value
End Function
