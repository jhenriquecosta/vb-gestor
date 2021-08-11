VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arrecadacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Function GravaDados() As Boolean
    On Error GoTo trata
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim PeriodoImposto As String
    Dim Conta As New ContaCorrente
    Dim Rs As VSRecordset
    Dim RSPago As VSRecordset
    Dim SequenciaLote  As Double
    Dim ValorImpostos As Double, ValorTaxas As Double
    PeriodoImposto = txtPeriodo
    
    Conta.PovoaParcelas txtDAM, False
    GravaDados = True
    Campos = "tdr_im,tdr_tip_cod_imposto,tdr_data_vencimento,tdr_periodo," & _
            "tdr_data_pagamento,tdr_valor_original,tdr_juros,tdr_multa," & _
            "tdr_valor_total,tdr_data_entrada,tdr_tus_cod_usuario,tdr_tgt_cod_pagamento,tdr_valor_real_pago,TDR_TLP_COD_LOTE,tdr_valor_real_juros,tdr_valor_real_multa,TDR_SEQUENCIA_DAM_LOTE,tdr_taxa_expidiente,tdr_sit_pago,TDR_TGT_COD_PAGAMENTO_VINCULADO"
    
    If Trim(txtCodImposto) = Const_Extrato Then
        Sql = " SELECT TPE_TGT_COD_PAGAMENTO,TPE_SUB_VALOR from TAB_PAGAMENTO_EXTRATO where TPE_COD_PAGAMENTO_EXTRATO =" & txtDAM
        If Bdados.AbreTabela(Sql, Rs) Then
            Rs.MoveFirst
            Do
                Sql = "Select tgt_cod_pagamento,tgt_im,tgt_tip_cod_imposto,tgt_data_vencimento,tgt_periodo," & _
                " tgt_valor_tributo,tgt_tim_ic,tgt_parcela,tgt_taxa_expediente,TGT_COD_PAGAMENTO_VINCULADO from tab_geracao_tributo where tgt_cod_pagamento = " & Rs(0) & " or tgt_cod_pagamento_original = " & Rs(0)
                If Bdados.AbreTabela(Sql, RSPago) Then
                    Valores = Bdados.PreparaValor(RSPago!TGT_im, RSPago!tgt_tip_cod_imposto, _
                    Bdados.Converte(RSPago!tgt_data_vencimento, TCDataHora), RSPago!TGT_periodo, Bdados.Converte(TxtDtPagamento, TCDataHora), RSPago!TGT_VALOR_TRIBUTO, _
                     0, 0, Bdados.Converte(RSPago!TGT_VALOR_TRIBUTO, TCDuplo), Bdados.Converte(Date, TCDataHora), Aplicacoes.Usuario, RSPago!tgt_cod_Pagamento, Rs!TPE_SUB_VALOR, txtNumLote, 0, 0, Bdados.Converte(txtSeq, TCInteiro), 0, 0, RSPago!TGT_COD_PAGAMENTO_VINCULADO)
                    If Bdados.InsereDados("Tab_Darm_Recebido", Valores, Campos) Then
                        Conta.MovimentaContaContribuinte RSPago!tgt_cod_Pagamento, RSPago!TGT_im, RSPago!TGT_tim_ic, RSPago!tgt_tip_cod_imposto, IIf(Len(RSPago!TGT_periodo) = 4, RSPago!TGT_periodo, Right(RSPago!TGT_periodo, 4) & Left(RSPago!TGT_periodo, 2)), IIf(CInt(RSPago!TGT_PARCELA) = 0, 1, 3), Nvl(RSPago!TGT_PARCELA, 0), TxtDtPagamento
'                        BaixaTaxasVinculadas RS!TPE_TGT_COD_PAGAMENTO, nvl("" & RSPago!tgt_taxa_expediente, 0)
                    Else
                        GravaDados = False
                        Exit Function
                    End If
                End If
                Rs.MoveNext
            Loop While Not Rs.EOF
        End If
        Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "TGT_COD_PAGAMENTO =" & txtDAM
        Exit Function
    End If
    Conta.CalculaTributoProporcional ValorImpostos, ValorTaxas, CDbl(CDbl(txtValorOriginal) + CDbl(Nvl(txtTaxa, 0)) + CDbl(Nvl(txtValorPagoJuros, 0)) + CDbl(Nvl(txtValorPagoMulta, 0))), Nvl(txtTaxa, 0), txtValorPago
    'Conta.CalculaTributoProporcional ValorImpostos, ValorTaxas, CDbl(CDbl(txtValorOriginal) + CDbl(Nvl(txtTaxa, 0))), Nvl(txtTaxa, 0), txtValorPago
    'MUDOU IMPOSTO OU TAXA
    If (ValorOriginalImpostoLancado <> CDbl(txtValorOriginal)) Or (ValorOriginalTaxaLancado <> CDbl(txtTaxa)) Then
        Bdados.GravaDados "tab_geracao_tributo", Bdados.PreparaValor(txtValorOriginal, txtTaxa), "tgt_valor_tributo,tgt_taxa_expediente", "tgt_cod_pagamento=" & txtDAM
        Bdados.DeletaDados "tab_conta_contribuinte", "tcc_codigo_conta = " & txtDAM
        Conta.CriaContaContribuinte txtDAM, txtIM, txtInscCadastral, txtCodImposto, IIf(Len(PeriodoImposto) = 4, PeriodoImposto, Right(PeriodoImposto, 4) & Left(PeriodoImposto, 2)), IIf(CInt(txtParcela) = 0, 1, 3)
    End If
    'Valores = Bdados.PreparaValor(IIf(Trim(txtIM) = "", Const_ImAvulso, txtIM), Bdados.Converte(txtCodImposto, TCTexto), Bdados.Converte(txtDtVencimento, TCDataHora), _
    IIf(Len(PeriodoImposto) = 4, PeriodoImposto, Right(PeriodoImposto, 4) & Left(PeriodoImposto, 2)), Bdados.Converte(TxtDtPagamento, TCDataHora), Bdados.Converte(txtValorOriginal, TCDuplo), Bdados.Converte(txtValorJuro, TCDuplo), Bdados.Converte(txtValorMulta, TCDuplo), _
    Bdados.Converte(CDbl(Nvl(txtValorOriginal, 0)) + CDbl(Nvl(txtTaxa, 0)), TCDuplo), Bdados.Converte(Date, TCDataHora), Aplicacoes.Usuario, txtDAM, Bdados.Converte(txtValorPago, TCDuplo), Bdados.Converte(txtNumLote, TCDuplo), Bdados.Converte(txtValorPagoJuros, TCDuplo), Bdados.Converte(txtValorPagoMulta, TCDuplo), Bdados.Converte(Nvl(txtSeq, 0), TCDuplo), Bdados.Converte(ValorTaxas, TCDuplo), 0, txtDAM)
    Valores = Bdados.PreparaValor(IIf(Trim(txtIM) = "", Const_ImAvulso, txtIM), Bdados.Converte(txtCodImposto, TCTexto), Bdados.Converte(txtDtVencimento, TCDataHora), _
    IIf(Len(PeriodoImposto) = 4, PeriodoImposto, Right(PeriodoImposto, 4) & Left(PeriodoImposto, 2)), Bdados.Converte(TxtDtPagamento, TCDataHora), Bdados.Converte(txtValorOriginal, TCDuplo), Bdados.Converte(txtValorJuro, TCDuplo), Bdados.Converte(txtValorMulta, TCDuplo), _
    Bdados.Converte(CDbl(Nvl(txtValorPago, 0)), TCDuplo), Bdados.Converte(Date, TCDataHora), Aplicacoes.Usuario, txtDAM, Bdados.Converte(ValorImpostos, TCDuplo), Bdados.Converte(txtNumLote, TCDuplo), Bdados.Converte(txtValorPagoJuros, TCDuplo), Bdados.Converte(txtValorPagoMulta, TCDuplo), Bdados.Converte(Nvl(txtSeq, 0), TCDuplo), Bdados.Converte(ValorTaxas, TCDuplo), 0, txtDAM)
    If Trim(txtInscCadastral) <> "" Then
        Campos = Campos & ",tdr_tim_ic"
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(txtInscCadastral, TCTexto))
    End If
    If Bdados.InsereDados("Tab_Darm_Recebido", Valores, Campos) Then
        Conta.MovimentaContaContribuinte txtDAM, txtIM, txtInscCadastral, txtCodImposto, IIf(Len(PeriodoImposto) = 4, PeriodoImposto, Right(PeriodoImposto, 4) & Left(PeriodoImposto, 2)), IIf(CInt(txtParcela) = 0, 1, 3), Nvl(txtParcela, 0), TxtDtPagamento
    Else
        GravaDados = False
        Exit Function
    End If
    'BAIXA AS TAXAS ANEXADAS AOS DAMs
    BaixaTaxasVinculadas txtDAM, ValorTaxas
    'FIM
    Conta.PovoaParcelas txtDAM, True
    Exit Function
trata:
    GravaDados = False
    Informa Err.Description
End Function



