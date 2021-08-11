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
Option Explicit
Public Enum ResultadoValidacaoDAM
    ervDocExtrato = 1
    ervDocNotificacao = 2
    ervDocOK = 3
    ervDocInexistente = 4
    ervDocPago = 5
    ervErro = 6
End Enum

Public Inscricao As String
Public TipoInscricao As Integer
Public CodImposto As String
Public Periodo As String
Public DtVencimento As String
Public DtPagamento As String
Public ValorOriginal  As Double
Public Juros As Double
Public Multa As Double
Public Parcela As Integer
Public DAM As String

Private Sub BaixaTaxasVinculadas(NumDam As Double, ValorTaxaModificada As Double, ValorAjustadoTaxa As Double, _
                DtPagamento As String, sNumLote As String, sSequencia As String)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Campos As String
    Dim Valores As String
    Dim Conta As New ContaCorrente
    Dim PercAtual As Double
    'VALORES ORIGINAIS
    Dim ValorOriginalImpostoLancado As Double, ValorOriginalTaxaLancado As Double
    Dim ValorOriginalJuros As Double, ValorOriginalMulta As Double
    Dim sPeriodo As String, sCodImposto As String, Parcela As Integer
    Dim Im As String, IC As String
    Dim DtVencimento As String
    
    If ValorOriginalTaxaLancado = 0 Then Exit Sub
    PercAtual = (ValorAjustadoTaxa / ValorOriginalTaxaLancado)
    If ValorAjustadoTaxa > 0 Then
        Sql = "Select TGT_COD_PAGAMENTO,TGT_TIP_COD_IMPOSTO,TGT_VALOR_TRIBUTO from tab_geracao_tributo where TGT_COD_PAGAMENTO_VINCULADO=" & NumDam & " AND TGT_COD_PAGAMENTO <> " & NumDam
        If Bdados.AbreTabela(Sql, Rs) Then
            Rs.MoveFirst
            If (ValorOriginalTaxaLancado <> CDbl(ValorTaxaModificada)) Then
                Do
                    Bdados.GravaDados "tab_geracao_tributo", Bdados.PreparaValor(Format(PercAtual * Rs!tgt_valor_tributo, Const_Monetario), 0), "tgt_valor_tributo,tgt_taxa_expediente", "tgt_cod_pagamento=" & Rs!TGT_COD_PAGAMENTO
                    Bdados.DeletaDados "tab_conta_contribuinte", "tcc_codigo_conta = " & Rs!TGT_COD_PAGAMENTO
                    Conta.CriaContaContribuinte Rs!TGT_COD_PAGAMENTO
                    Rs.MoveNext
                Loop While Not Rs.EOF
                Rs.MoveFirst
            End If
            'PEGA INFORMAÇÕES ORIGINAIS DO PAGAMENTO
            Sql = "Select toc_cod_obrigacao,TOC_INSCRICAO,TOC_tip_cod_imposto,TOC_data_vencimento,TOC_Periodo,TOC_Valor_multa," & _
                " TOC_valor_juros,TOC_valor_OBRIGACAO from TAB_OBRIGACAO_CONTRIBUINTE where TOC_COD_OBRIGACAO = " & NumDam
            If Bdados.AbreTabela(Sql, Rs) Then
                 Rs.MoveFirst
                Do
                    ValorOriginalImpostoLancado = Rs!TOC_VALOR_OBRIGACAO: ValorOriginalTaxaLancado = 0
                    ValorOriginalJuros = 0: ValorOriginalMulta = 0
                    sPeriodo = Rs!TOC_PERIODO: sCodImposto = Rs!TOC_TIP_COD_IMPOSTO:
                    Parcela = 0: DtVencimento = Rs!TOC_DATA_VENCIMENTO
                    Im = Rs!TOC_INSCRICAO: IC = Rs!TOC_INSCRICAO
                    
                    Valores = Bdados.PreparaValor(Im, sCodImposto, Bdados.Converte(DtVencimento, TCDataHora), _
                    IIf(Len(sPeriodo) = 4, sPeriodo, Right(sPeriodo, 4) & Left(sPeriodo, 2)), Bdados.Converte(DtPagamento, TCDataHora), Bdados.Converte(Format(Rs!tgt_valor_tributo, Const_Monetario), TCDuplo), 0, 0, _
                    Bdados.Converte(Format(ValorOriginalImpostoLancado, Const_Monetario), TCDuplo), Bdados.Converte(Date, TCDataHora), AplicacoesVTFuncoes.Usuario, , Bdados.Converte(Format(Rs!TOC_VALOR_OBRIGACAO * PercAtual, Const_Monetario), TCDuplo), Trim(sNumLote), 0, 0, Bdados.Converte(Nvl(sSequencia, 0), TCDuplo), 0, 0, NumDam)
                    Campos = "TDR_INSCRICAO,tdr_tip_cod_imposto,tdr_data_vencimento,tdr_sPeriodo," & _
                        "tdr_data_pagamento,tdr_valor_original,tdr_juros,tdr_multa," & _
                        "tdr_valor_total,tdr_data_entrada,tdr_tus_cod_usuario,tdr_tgt_cod_pagamento,tdr_valor_real_pago,TDR_TLP_COD_LOTE,tdr_valor_real_juros,tdr_valor_real_multa,TDR_SEQUENCIA_DAM_LOTE,tdr_taxa_expidiente,tdr_sit_pago,TDR_TGT_COD_PAGAMENTO_VINCULADO"
                    If Trim(IC) <> "" Then
                        Campos = Campos & ",tdr_tim_ic"
                        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Trim(IC), tctexto))
                    End If
                    Call Bdados.InsereDados("Tab_Darm_Recebido", Valores, Campos)
                    Conta.MovimentaContaContribuinte Rs!TGT_COD_PAGAMENTO
                    Rs.MoveNext
                Loop While Not Rs.EOF
            End If
            Rs.Fechar
        End If
    End If
End Sub

Public Function GravaPagamento(ByVal Contribuinte As String, TipoInscricao As Integer, ByVal CodTributo As String, ByVal Periodo As String, _
        ByVal DtVencimento As String, ByVal DtPagamento As String, ByVal ValorTributo As Double, ByVal ValorTaxas As Double, _
        ByVal ValorDeducao As Double, ByVal ValorJuros As Double, ByVal Correcao As Double, ByVal ValorMulta As Double, ByVal sNumLote As String, _
        ByVal sSequencia As String, ByVal Parcela As Integer, Optional CodObrigacao As String, Optional SubstituiAntiga As TipoSubstituicaoObrigacao = etsNaoSubstitui, Optional ValorPago As Double = -1, Optional Origem As String, Optional Status As TipoStatusObrigacao) As Boolean
                     
    On Error GoTo Trata
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim Conta As New ContaCorrente
    Dim Rs As VSRecordset
    Dim RSPago As VSRecordset
    Dim SequenciaLote  As Double
    Dim ValorImpostoAjustado As Double, ValorTaxasAjustado As Double
    Dim Obrig As New Obrigacao
    Dim CodigoObrigacao As String
    Dim DAMEncontrado As Boolean
    Dim ValorTotalPago As Double
    Dim Substitui As Integer
    Dim Contador As Integer
    Dim RsTot As VSRecordset
    GravaPagamento = True
    Contribuinte = Trim(Contribuinte)
    If CodTributo = Const_Notificacao Or CodTributo = Const_Extrato Then
        Sql = "SELECT TPE_TGT_COD_PAGAMENTO,TPE_SUB_VALOR FROM TAB_PAGAMENTO_EXTRATO WHERE TPE_COD_PAGAMENTO_EXTRATO = " & CodObrigacao
        If Bdados.AbreTabela(Sql, Rs) Then
            Rs.MoveFirst
            Contador = 0
            Sql = "SELECT SUM(TPE_SUB_VALOR) AS TOTAL FROM TAB_PAGAMENTO_EXTRATO WHERE TPE_COD_PAGAMENTO_EXTRATO = " & CodObrigacao
            Call Bdados.AbreTabela(Sql, RsTot)
            Do
                Contador = Contador + 1
                Obrig.BuscaDetalheObrigacao Rs!TPE_TGT_COD_PAGAMENTO
                If Obrig.obContribuinte <> "" Then
                    'CUIDADO COM TAXA ABAIXO DE R$1,00
                    If Fix(ValorTributo) > Fix(RsTot!Total) And Contador = 1 Then
                        GravaPagamento Obrig.obContribuinte, Obrig.obTipoInscricao, Obrig.obCodImposto, Obrig.obPeriodo, Obrig.obDataVencimento, DtPagamento, Rs!TPE_SUB_VALOR, 0, 0, 0, 0, 0, sNumLote, sSequencia, Parcela, Obrig.obCodigoObrigacao, , CDbl(Rs!TPE_SUB_VALOR + CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0)))
                    Else
                        GravaPagamento Obrig.obContribuinte, Obrig.obTipoInscricao, Obrig.obCodImposto, Obrig.obPeriodo, Obrig.obDataVencimento, DtPagamento, Rs!TPE_SUB_VALOR, 0, 0, 0, 0, 0, sNumLote, sSequencia, Parcela, Obrig.obCodigoObrigacao, , Rs!TPE_SUB_VALOR
                    End If
                Else
                    Avisa "Documento No. " & Rs!TPE_TGT_COD_PAGAMENTO & " do extrato No. " & CodObrigacao & " nao encontrado. Verifique informacões do registro na consulta de extrato."
                End If
                Rs.MoveNext
            Loop While Not Rs.EOF
        End If
    Else
        If CodTributo = Const_AutoInfracao Then
            Sql = "SELECT TAO_TOC_COD_OBRIGACAO,TAO_TOC_VALOR_TOTAL FROM TAB_AUTO_OBRIGACAO WHERE TAO_TAI_AUTO= " & CodObrigacao
            If Bdados.AbreTabela(Sql, Rs) Then
                Rs.MoveFirst
                Do While Not Rs.EOF
                    Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(etsCreditoPagoEmAuto), "TOC_STATUS_OBRIGACAO", "TOC_COD_OBRIGACAO=" & Rs!TAO_TOC_COD_OBRIGACAO
                    Rs.MoveNext
                Loop
            End If
        End If
        If Left(Periodo, 2) > 12 Then
            Periodo = IIf(Len(Periodo) = 4, Periodo, Right(Periodo, 2) & Left(Periodo, 4))
        End If
        If ValorPago = -1 Then
            ValorTotalPago = (ValorTributo + ValorTaxas + ValorJuros + ValorMulta) - ValorDeducao
        Else
            ValorTotalPago = ValorPago
        End If
        Substitui = IIf(CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) Or CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) Or CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN)) Or CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) Or CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQNSUBST)), etsNaoSubstitui, etsCriaNova)
        CodigoObrigacao = Obrig.CriaObrigacao(CodTributo, Periodo, Periodo, Contribuinte, _
                              ValorPago + ValorTaxas, Status, CInt(Substitui), DtVencimento, , CodObrigacao, , , , , Origem, Parcela, Contribuinte, CInt(TipoInscricao))

        Periodo = IIf(Len(Periodo) = 4, Periodo, Right(Periodo, 4) & Left(Periodo, 2))
        Campos = "TDR_INSCRICAO, tdr_tip_cod_imposto," _
                & " tdr_periodo, tdr_data_vencimento," _
                & " tdr_data_pagamento, tdr_valor_original," _
                & " tdr_juros, tdr_multa,tdr_corrercao," _
                & " tdr_valor_total, tdr_valor_real_pago," _
                & " tdr_valor_real_juros, tdr_valor_real_multa," _
                & " TDR_TLP_COD_LOTE, TDR_SEQUENCIA_DAM_LOTE," _
                & " tdr_sit_pago, tdr_data_entrada," _
                & " tdr_tus_cod_usuario, tdr_tgt_cod_pagamento," _
                & " tdr_parcela,tdr_tipo_inscricao"
    
        Valores = Bdados.PreparaValor(Bdados.Converte(Contribuinte, tctexto), CodTributo, _
            Periodo, Bdados.Converte(DtVencimento, TCDataHora), _
            Bdados.Converte(DtPagamento, TCDataHora), Bdados.Converte(ValorTributo, TCDuplo), _
            Bdados.Converte(Juros, TCDuplo), Bdados.Converte(Multa, TCDuplo), Bdados.Converte(Correcao, TCDuplo), _
            Bdados.Converte(ValorPago + ValorJuros + ValorMulta - ValorDeducao, TCDuplo), Bdados.Converte(ValorPago, TCDuplo), _
            Bdados.Converte(ValorJuros, TCDuplo), Bdados.Converte(ValorMulta, TCDuplo), _
            Bdados.Converte(sNumLote, TCDuplo), Bdados.Converte(Nvl(sSequencia, 0), TCDuplo), _
            0, Bdados.Converte(Date, TCDataHora), _
            AplicacoesVTFuncoes.Usuario, CodigoObrigacao, _
            Parcela, TipoInscricao)
        If CDbl(Nvl(Trim(CodigoObrigacao), 0)) = 0 Then
            GravaPagamento = False
            Exit Function
        End If
        If Bdados.GravaDados("Tab_Darm_Recebido", Valores, Campos, "tdr_tgt_cod_pagamento =" & CodigoObrigacao) Then
            Conta.BaixaPagamentos Contribuinte, Periodo, CodTributo, CDbl(ValorTotalPago), DtPagamento, _
            CodigoObrigacao, Obrig, CInt(TipoInscricao), Parcela
        Else
            GravaPagamento = False
            Exit Function
        End If
    End If
    Exit Function
Trata:
    GravaPagamento = False
    Informa Err.Description
End Function

Public Function AbreLote(Banco As Integer, Agencia As String, Conta As String, _
                ValorLote As Double, DataArrecadacao As String, DataRecepcao As String) As Double
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Condicao As String
    Dim NumLote As Double
    Dim cConta As New ContaCorrente
'    If Bdados.AbreTabela("Select max(TLP_COD_LOTE) from TAB_LOTE_PAGAMENTO", Rs) Then
'        NumLote = CDbl(Nvl("" & Rs(0), 0)) + 1
'    Else
'        NumLote = 1
'    End If
'    Bdados.FechaTabela Rs
    NumLote = cConta.GeraCodPagamento("LOTE") 'LOTE
    Campos = "TLP_COD_LOTE,TLP_TAR_COD_AGENTE,TLP_NUM_SUCURSAL,TLP_NUM_CONTA,TLP_VALOR_ARRECADADO," & _
            "TLP_DATA_ABERTURA,TLP_DATA_ARRECADACAO,TLP_DATA_RECEPCAO,TLP_SITUACAO_LOTE," & _
            "TLP_TUS_COD_USUARIO"
    Valores = Bdados.PreparaValor(NumLote, Banco, Agencia, Conta, _
            Bdados.Converte(ValorLote, TCDuplo), Date, DataArrecadacao, DataRecepcao, 0, _
            AplicacoesVTFuncoes.Usuario)
    If Bdados.InsereDados("TAB_LOTE_PAGAMENTO", Valores, Campos) Then
        AbreLote = NumLote
    End If
End Function
Public Sub LogOcorrencia(CodPagamento As String, Valor As Double, DataArrecadacao As String, Situacao As ResultadoValidacaoDAM, Banco As String, Agencia As String)
    Bdados.GravaDados "TAB_LOG_OCORRENCIA", _
            Bdados.PreparaValor(CodPagamento, Bdados.Converte(Valor, TCMonetario), _
                                DataArrecadacao, Situacao, Banco, Agencia), _
            "TLA_TGT_COD_PAGAMENTO,TLA_VALOR,TLA_DATA_ARRECADACAO,TLA_OCORRENCIA," & _
            "TLA_TAR_COD_AGENTE,TLA_AGENCIA", "TLA_TGT_COD_PAGAMENTO=" & CodPagamento
End Sub
Public Function ValidaDocumento(ByVal CodPagamento As String, ByRef Valor As Double, ByRef DataArrecadacao As String, _
            ByVal Banco As String, Agencia As String, Optional ByRef ValorOriginalImposto, Optional NomeArquivo As String, Optional CodLote As String) As ResultadoValidacaoDAM
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsAux As VSRecordset
    Dim Valida As ResultadoValidacaoDAM
    Dim SistemaOrigem As Byte '1 - GESTOR 2 - OUTRO SISTEMA(SINFAZ)
    Dim NossoNumero As String
    Valida = ervDocInexistente
    If CodPagamento = 0 Then Exit Function
    'RETIRAR NA PROXIMA VIAGEM A BALSAS(APROX. 14-SET-2004)
    DAM = ""
    DtVencimento = ""
    DtPagamento = ""
    ValorOriginal = 0
    Inscricao = ""
    Set Rs = Conta.BuscaDam(CodPagamento)
    If Len(CodPagamento) > 12 Then
        NossoNumero = CodPagamento
    Else
        NossoNumero = CDbl(CodPagamento)
    End If
    If Not Rs.EOF Then
        Sql = "Select tdr_tgt_cod_pagamento from tab_darm_recebido where tdr_tgt_cod_pagamento =" & CodPagamento & " and tdr_tlp_cod_lote <> " & CodLote
        If Bdados.AbreTabela(Sql, RsAux) Then
            Valida = ervDocPago
            ValidaDocumento = ervDocPago
            Exit Function
        End If
        Valida = ervDocOK
        If Rs!Sinfaz = 0 Then
            Inscricao = Trim("" & Rs!Inscricao)
            TipoInscricao = "" & Rs!TipoInscricao
            CodImposto = Trim("" & Rs!Imposto)
            Periodo = "" & Rs!Periodo
            If Len(Periodo) = 5 Then
                Periodo = Mid(Periodo, 2) & Format(Left(Periodo, 1), "00")
            End If
            DtVencimento = Nvl("" & Rs!Vencimento, 0)
            DtPagamento = "" & DataArrecadacao
            Parcela = Nvl("" & Rs!Parcela, 0)
            ValorOriginal = Nvl("" & Rs!ValorTributo, 0)
            Juros = Nvl("" & Rs!Juros, 0)
            Multa = Nvl("" & Rs!Multa, 0)
            DAM = Rs!Documento
            SistemaOrigem = 1
            
        Else 'PARA PETROLINA(SOLUCAO TEMPORARIA PARA CONCORRENCIA DE DAM(SINFAZ/GESTOR))
            Inscricao = Trim("" & Rs!Inscricao)
            TipoInscricao = "" & Rs!Modulo
            CodImposto = Trim("" & Rs!Imposto)
            
            If TipoInscricao = 1 Then
                Periodo = "" & Rs!Periodo
                Parcela = CInt(Right(CodPagamento, 2))
            Else
                Periodo = "" & Rs!Periodo & Right(CodPagamento, 2)
                Parcela = 0
            End If
            SistemaOrigem = 2
            Sql = "SELECT TOC_INSCRICAO AS Inscricao,TOC_COD_OBRIGACAO AS Documento, TOC_DATA_VENCIMENTO AS Vencimento," & _
                " TOC_VALOR_OBRIGACAO AS ValorTributo FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE " & _
                " TOC_TIP_COD_IMPOSTO ='" & CodImposto & "' AND (TOC_INSCRICAO LIKE '" & Inscricao & "_' OR TOC_INSCRICAO = '" & Inscricao & "') " & _
                " AND TOC_PERIODO = " & Periodo & " AND TOC_TIPO_INSCRICAO =" & TipoInscricao & _
                " AND TOC_PARCELA = " & Parcela
            If Bdados.AbreTabela(Sql, Rs) Then
                DAM = Rs!Documento
                CodPagamento = DAM
                DtVencimento = Nvl("" & Rs!Vencimento, 0)
                DtPagamento = "" & DataArrecadacao
                ValorOriginal = Nvl("" & Rs!ValorTributo, 0)
                Inscricao = "" & Rs!Inscricao
            Else
                DtVencimento = Imposto.BuscaDataVencimento(CodImposto, IIf(Len(Periodo) = 4, Periodo, Right(Periodo, 2) & Left(Periodo, 4)))
                DtPagamento = "" & DataArrecadacao
                DAM = Conta.GeraCodPagamento(CodImposto)
                CodPagamento = DAM
                If TipoInscricao = 2 Then
                    Sql = "SELECT TCI_IM AS INSCRICAO FROM TAB_CONTRIBUINTE WHERE TCI_IM LIKE '" & Inscricao & "_' OR TCI_IM = '" & Inscricao & "'"
                    If Bdados.AbreTabela(Sql, RsAux) Then
                        Inscricao = "" & RsAux!Inscricao
                    End If
                End If
            End If
        End If
    End If
    
    ValidaDocumento = Valida
    
    If CodImposto = Const_Extrato Then Valida = ervDocExtrato
    If CodImposto = Const_Notificacao Then Valida = ervDocNotificacao
    Bdados.GravaDados "TAB_LOG_OCORRENCIA", _
        Bdados.PreparaValor(CodPagamento, Bdados.Converte(Valor, TCMonetario), DataArrecadacao, Valida, _
        Banco, Agencia, NomeArquivo, Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora), _
        AplicacoesVTFuncoes.Usuario, CodLote, Periodo, CodImposto, SistemaOrigem, NossoNumero), _
        "TLA_TGT_COD_PAGAMENTO,TLA_VALOR,TLA_DATA_ARRECADACAO,TLA_OCORRENCIA," & _
        "TLA_TAR_COD_AGENTE,TLA_AGENCIA,TLA_ARQUIVO,TLA_DATA_RECEPCAO,TLA_TUS_COD_USUARIO,TLA_TLP_COD_LOTE,TLA_TGT_PERIODO,TLA_TIP_COD_IMPOSTO,TLA_SISTEMA_ORIGEM,TLA_NOSSO_NUMERO", "TLA_TGT_COD_PAGAMENTO=" & CodPagamento
End Function

Public Function FechaLote(CodLote As Double) As Boolean
    FechaLote = Bdados.AtualizaDados("TAB_LOTE_PAGAMENTO", Bdados.PreparaValor(1), "TLP_SITUACAO_LOTE", "TLP_COD_LOTE = " & CodLote)
End Function

Public Function GravaPagamento_Cotas_Obrigacao(ByVal Contribuinte As String, ByVal CodTributo As String, ByVal Periodo As String, _
        ByVal DtVencimento As String, ByVal DtPagamento As String, ByVal ValorTributo As Double, ByVal ValorTaxas As Double, _
        ByVal ValorDeducao As Double, ByVal ValorJuros As Double, ByVal ValorMulta As Double, ByVal sNumLote As String, _
        ByVal sSequencia As String, ByVal Parcela As Integer, Optional CodObrigacao As String, Optional SubstituiAntiga As TipoSubstituicaoObrigacao = etsNaoSubstitui, Optional ValorPago As Double = -1, Optional Origem As String, Optional Status As TipoStatusObrigacao) As Boolean
                     
    On Error GoTo Trata
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim Conta As New ContaCorrente
    Dim Rs As VSRecordset
    Dim RSPago As VSRecordset
    Dim SequenciaLote  As Double
    Dim ValorImpostoAjustado As Double, ValorTaxasAjustado As Double
    Dim Obrig As New Obrigacao
    Dim CodigoObrigacao As String
    Dim DAMEncontrado As Boolean
    Dim ValorTotalPago As Double
    ValorTotalPago = ValorPago
    GravaPagamento_Cotas_Obrigacao = True
    Contribuinte = Trim(Contribuinte)
    Periodo = IIf(Len(Periodo) = 4, Periodo, Right(Periodo, 4) & Left(Periodo, 2))
    If CodObrigacao = "" Then
        Sql = "Select TCO_TOC_COD_OBRIGACAO as CodPagamento from TAB_COTAS_OBRIGACAO where tco_inscricao ='" & _
                Contribuinte & "' and tco_tip_cod_imposto ='" & CodTributo & _
                "' and tco_periodo = " & Periodo & " and tco_num_parcela =" & Parcela
        If Not Bdados.AbreTabela(Sql, Rs) Then
            Sql = "Select TCP_NUM_COTA as CodPagamento from TAB_COTAS_PARCELAMENTO where tcp_inscricao ='" & _
                Contribuinte & "' and tcp_tip_cod_imposto ='" & CodTributo & _
                "' and tcp_periodo = " & Periodo & " and tcp_num_parcela =" & Parcela
            If Not Bdados.AbreTabela(Sql, Rs) Then
                Exit Function
            End If
        End If
        CodObrigacao = "" & Rs!CodPagamento
    End If
    
    
    Campos = "TDR_INSCRICAO, tdr_tip_cod_imposto," _
            & " tdr_periodo, tdr_data_vencimento," _
            & " tdr_data_pagamento, tdr_valor_original," _
            & " tdr_juros, tdr_multa," _
            & " tdr_valor_total, tdr_valor_real_pago," _
            & " tdr_valor_real_juros, tdr_valor_real_multa," _
            & " TDR_TLP_COD_LOTE, TDR_SEQUENCIA_DAM_LOTE," _
            & " tdr_sit_pago, tdr_data_entrada," _
            & " tdr_tus_cod_usuario, tdr_tgt_cod_pagamento," _
            & " tdr_parcela"

    Valores = Bdados.PreparaValor(Bdados.Converte(Contribuinte, tctexto), CodTributo, _
        Periodo, Bdados.Converte(DtVencimento, TCDataHora), _
        Bdados.Converte(DtPagamento, TCDataHora), Bdados.Converte(ValorTributo, TCDuplo), _
        Bdados.Converte(ValorJuros, TCDuplo), Bdados.Converte(ValorMulta, TCDuplo), _
        Bdados.Converte(ValorTributo + ValorJuros + ValorMulta - ValorDeducao, TCDuplo), Bdados.Converte(ValorTotalPago, TCDuplo), _
        Bdados.Converte(ValorJuros, TCDuplo), Bdados.Converte(ValorMulta, TCDuplo), _
        Bdados.Converte(sNumLote, TCDuplo), Bdados.Converte(Nvl(sSequencia, 0), TCDuplo), _
        0, Bdados.Converte(Date, TCDataHora), _
        AplicacoesVTFuncoes.Usuario, CodObrigacao, _
        Parcela)
    If CDbl(Nvl(Trim(CodObrigacao), 0)) = 0 Then
        GravaPagamento_Cotas_Obrigacao = False
        Exit Function
    End If
    
    If Bdados.InsereDados("Tab_Darm_Recebido", Valores, Campos) Then
        Conta.BaixaPagamentos Contribuinte, Periodo, _
        CodTributo, CDbl(ValorTotalPago), DtPagamento, CodObrigacao, Obrig, , Parcela
    Else
        GravaPagamento_Cotas_Obrigacao = False
        Exit Function
    End If

    Exit Function
Trata:
    GravaPagamento_Cotas_Obrigacao = False
    Informa Err.Description
    Exit Function
    Resume
End Function

