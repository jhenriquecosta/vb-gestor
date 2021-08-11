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
Private Const Const_Obrig As String = "63"
Private Const Const_NaoParcelaveis As String = "3,4,6,7"
Private Const Const_NaoPagos As String = "2,4,5"
Public obContribuinte As String
Public obInscMunicipal As String
Public obInscCadastral As String
Public obPeriodo As String
Public obCodImposto As String
Public obStatusObrigacao As Byte
Public obDataVencimento As String
Public obTipoInscricao As Byte
Public obCodigoObrigacao As String
Public obValorObrigacao As Double
Public obValorMulta As Double
Public obValorJuros As Double
Public obNumParcelamento As Double

Public Enum TipoStatusObrigacao
    etsCreditoNaoLancado = 1
    etsCreditoOriginalAberto = 2
    etsCreditoPago = 3
    etsCreditoParcelado = 4
    etsCreditoDividaAtiva = 5
    etsCreditoIsento = 6
    etsCreditoCancelado = 7
End Enum

Public Enum TipoStatusListagem
    etlTodos = 1
    etlNaoPagos = 2
    etlPagos = 3
    etlParcelaveis = 4
End Enum

Public Function BuscaSujeitoPassivoObrigacao(Inscricao As String, Optional ByRef RazaoSocial As Object, Optional ByRef Endereco As Object, Optional ByRef Documento As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    RazaoSocial = ""
    Endereco = ""
    Documento = ""
    If Len(Trim(Inscricao)) = 10 And IsNumeric(Inscricao) Then
        Inscricao = Imposto.FormataInscricao(Trim(Inscricao), InscContrib)
    End If
'    If Len(Trim(Inscricao2)) = 10 Then
'        Inscricao2 = Imposto.FormataInscricao(Trim(Inscricao2), InscContrib)
'    End If
    If Inscricao = "" Then Exit Function

    Sql = "select VIN_INSCRICAO, VIN_RAZAO,VIN_ENDERECO, VIN_DOCUMENTO from VIS_INSCRICAO " & _
    " where VIN_INSCRICAO  = '" & Inscricao & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        RazaoSocial = Rs!VIN_Razao & ""
        Endereco = Rs!VIN_Endereco & ""
        Documento = Rs!VIN_DOCUMENTO & ""
        BuscaSujeitoPassivoObrigacao = Trim(Rs!VIN_INSCRICAO & "")
    End If
End Function


Public Function CalculaObrigacao(ByVal Inscricao As String, ByVal TipoInscricao As TipoInsc, ByVal CodTributo As String, _
                ByVal SiglaTributo As String, ByVal Periodo As String, ByVal DataVence As String, Optional ByRef TipoLancamento As TipoLanc) As Double
        Dim Taxas As Double
        Dim CodPagamento As Double
        
        Dim Conta As New ContaCorrente
        Dim ValorObrigacao As Double
        Inscricao = Trim(Inscricao)
        Select Case SiglaTributo
            Case Imposto.NomeTributo(ttr_ALVARA)
                ValorObrigacao = Calculo.Alvara(Inscricao, Periodo, CodTributo, DataVence)
                
            Case Imposto.NomeTributo(ttr_IPTU)
                ValorObrigacao = Calculo.Iptu(Inscricao, Periodo, CodTributo, DataVence)
                
            Case Imposto.NomeTributo(ttr_ISSQN), Imposto.NomeTributo(ttr_ISSSUBST)
                ValorObrigacao = Calculo.Issqn(Inscricao, Periodo, CodTributo, DataVence)
        End Select
        
        JurosObrigacao = Format(Conta.CalculaValoresJurosAvulsos(CodTributo, CLng(Periodo), 1, Format(Date, "DD/MM/YYYY"), DataVence, ValorObrigacao), Const_Monetario)
        MultaObrigacao = Format(Conta.CalculaValoresMultaAvulsos(CodTributo, CLng(Periodo), 1, Format(Date, "DD/MM/YYYY"), DataVence, ValorObrigacao), Const_Monetario)
        CalculaObrigacao = Format(ValorObrigacao, Const_Monetario)
        TipoLancamento = Calculo.TipoLancamento
End Function

Public Function CarregaListaObrigacao(Lista As Object, Optional Inscricao As String, Optional CodTributo As String, Optional Periodo As String, Optional TipoListagem As TipoStatusListagem = etlTodos, Optional NumParcelamento As String) As Boolean
    Dim Sql As String
    Sql = "select tcc_codigo_conta AS Obrigacao,tcc_inscricao as Contribuinte, " & _
            "tip_sigla_imposto as [Cod Receita] , tcc_periodo as Periodo, tcc_data_vencimento as Vencimento, " & _
            "tcc_imposto_original as Original, tcc_juros_atual as Juros, tcc_multa_atual as Multa, " & _
            "tcc_saldo_atual as Total,tcc_tip_cod_imposto " & _
            "from tab_conta_contribuinte,tab_imposto where tcc_tip_cod_imposto = tip_cod_imposto "
    If TipoListagem = etlParcelaveis Then
        Sql = Sql & " and tcc_status_conta not in(" & Const_NaoParcelaveis & ") "
        Sql = Sql & " and tcc_data_vencimento < " & Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora)
    End If
    If TipoListagem = etlPagos Then Sql = Sql & " and tcc_status_conta = " & etsCreditoPago
    If TipoListagem = etlNaoPagos Then Sql = Sql & " and tcc_status_conta in (" & Const_NaoPagos & ")"
    If Trim(CodTributo) <> "" Then Sql = Sql & " AND tcc_tip_cod_imposto ='" & CodTributo & "'"
    If Trim(Inscricao) <> "" Then Sql = Sql & " and tcc_inscricao = '" & Inscricao & "'"
    If Trim(NumParcelamento) <> "" Then Sql = Sql & " and TCC_TPA_COD_PARCELAMENTO = " & NumParcelamento
    CarregaListaObrigacao = Lista.Preencher(Bdados, Sql, 0, 1600, 1200, 800, 1200, 900, 900, 900, 900, 0)
End Function

Public Function CarregaListaObrigacaoVencida(Lista As Object, Optional Inscricao As String, Optional CodTributo As String, Optional Periodo As String, Optional TipoListagem As TipoStatusListagem = etlTodos, Optional NumParcelamento As String) As Boolean
    Dim Sql As String
    Sql = "select tcc_codigo_conta AS Obrigacao,tcc_inscricao as Contribuinte, " & _
            "tip_sigla_imposto as [Cod Receita] , tcc_periodo as Periodo, tcc_data_vencimento as Vencimento, " & _
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

Public Function CarregaListaObrigacaoNaoVencida(Lista As Object, Optional Inscricao As String, Optional CodTributo As String, Optional Periodo As String, Optional TipoListagem As TipoStatusListagem = etlTodos, Optional NumParcelamento As String) As Boolean
    Dim Sql As String
    Sql = "select tcc_codigo_conta AS Obrigacao,tcc_inscricao as Contribuinte, " & _
            "tip_sigla_imposto as [Cod Receita] , tcc_periodo as Periodo, tcc_data_vencimento as Vencimento, " & _
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
        CodPagamento = Bdados.Correlativo("TRIB", Rs!TIP_COD_CORRELATIVO)
        Conta.GeraPagamento Trim(Rs!TOC_INSCRICAO), "", Rs!TOC_TIP_COD_IMPOSTO, Rs!TOC_periodo, Rs!TOC_data_vencimento, Rs!TOC_valor_OBRIGACAO, Rs!TOC_Valor_multa, Rs!TOC_valor_juros, CDbl(CodPagamento), 0, 0, 0
        CodObrigacao = CodPagamento
    End If
End Function

Public Function IdentificaInscricao(Optional IM As String, Optional IC As String) As String
    If Len(Trim(Edita.TiraPic(Edita.TiraPic(Edita.TiraPic(IM, "."), "-"), "/"))) = 10 Then
        IdentificaInscricao = IM
    Else
        IdentificaInscricao = IC
    End If
End Function


Public Function MostraObrigacaoGerada(Lista As Object, Tributo As String, Optional Contribuinte As String, Optional Restricao As Integer, _
                Optional Status As TipoStatusObrigacao = -1, Optional PeriodoInicial As String, _
                Optional PeriodoFinal As String, Optional ExercicioInicial As String, Optional ExercicioFinal As String) As Boolean
    Dim Sql As String
    ExercicioInicial = IIf(Len(ExercicioInicial) = 4, ExercicioInicial, Right(ExercicioInicial, 4) & Left(ExercicioInicial, 2))
    ExercicioFinal = IIf(Len(ExercicioFinal) = 4, ExercicioFinal, Right(ExercicioFinal, 4) & Left(ExercicioFinal, 2))
    'sql = "SELECT TOC_COD_OBRIGACAO,TOC_INSCRICAO as INSCRICAO,TIP_SIGLA_IMPOSTO AS TRIBUTO, TOC_PERIODO AS PERIODO,TOC_DATA_VENCIMENTO AS VENCIMENTO,TOC_VALOR_OBRIGACAO AS VALOR ," & _
            "TOC_VALOR_JUROS AS JUROS,TOC_VALOR_MULTA as MULTA,TOC_STATUS_OBRIGACAO as SIT,TOC_TOTAL_TAXA_INCLUSA,TOC_TIP_COD_IMPOSTO,tip_nome_imposto  FROM TAB_OBRIGACAO_CONTRIBUINTE,TAB_IMPOSTO "
    Sql = "SELECT TCC_CODIGO_CONTA,TCC_INSCRICAO as INSCRICAO,TIP_SIGLA_IMPOSTO AS TRIBUTO, TCC_PERIODO AS PERIODO,TCC_DATA_VENCIMENTO AS VENCIMENTO,TCC_IMPOSTO_ORIGINAL AS VALOR ," & _
            "tcc_juros_atual AS JUROS,tcc_multa_atual as MULTA,tcc_status_conta as SIT,TOC_TOTAL_TAXA_INCLUSA,TOC_TIP_COD_IMPOSTO,tip_nome_imposto  FROM TAB_CONTA_CONTRIBUINTE,TAB_OBRIGACAO_CONTRIBUINTE,TAB_IMPOSTO "
    Sql = Sql & " WHERE TCC_CODIGO_CONTA = TOC_COD_OBRIGACAO AND TCC_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
    If Trim(Tributo) <> "" Then Sql = Sql & " AND TOC_TIP_COD_IMPOSTO ='" & Tributo & "'"
        
    If Trim(PeriodoInicial) <> "" And Trim(PeriodoFinal) <> "" Then
        Sql = Sql & " AND TOC_PERIODO BETWEEN " & PeriodoInicial & " AND " & PeriodoFinal
    End If
    
'    If Trim(PeriodoInicial) <> "" And Trim(PeriodoFinal) <> "" Then
'        Sql = Sql & " AND TOC_DATA_GERACAO >= " & Bdados.Converte(PeriodoInicial, TCDataHora) & " and TOC_DATA_GERACAO <= " & Bdados.Converte(PeriodoFinal, TCDataHora)
'    End If
    
    If Status > 0 Then Sql = Sql & " AND TOC_STATUS_OBRIGACAO=" & Status
    If Restricao = 1 Then Sql = Sql & " AND TOC_STATUS_OBRIGACAO not in (" & etsCreditoPago & "," & etsCreditoIsento & ")"
    If Restricao = 2 Then Sql = Sql & " AND TOC_STATUS_OBRIGACAO = " & etsCreditoPago
    If Trim(Contribuinte) <> "" Then Sql = Sql & " AND TOC_INSCRICAO like  '" & Trim(Contribuinte) & "%'" 'feito p Maickel, para trazer todas as inscricoes de setores, quadaras de setoes etc.
    Sql = Sql & " ORDER BY TCC_PERIODO"
    MostraObrigacaoGerada = Lista.Preencher(Bdados, Sql, 0, 1600, 1100, 1100, 1150, 1000, 800, 800, 450, 0, 0, 0)
    If Lista.ListItems.Count > 0 Then Lista.Mensagem = "Total em Imposto: R$" & Format(Lista.Colunas(6).Soma, Const_Monetario) & _
        " - Total em Juros: R$" & Format(Lista.Colunas(7).Soma, Const_Monetario) & _
        " - Total em Multa: R$" & Format(Lista.Colunas(8).Soma, Const_Monetario) & _
        " - Total Geral: R$" & Format(Lista.Colunas(6).Soma + Lista.Colunas(7).Soma + Lista.Colunas(8).Soma, Const_Monetario)
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

Public Sub PreencheComboTributo(Combo As Object, Optional SomenteComObrigacao As Boolean = True)
    Dim Sql As String
     Sql = "Select  tip_cod_imposto,tip_sigla_imposto " & Bdados.Concatena & " ' # ' " & Bdados.Concatena & " tip_nome_imposto From TAB_IMPOSTO "
    If SomenteComObrigacao Then
        Sql = Sql & " where tip_cod_imposto in (select tpi_tip_cod_imposto from tab_parametro_imposto where " & _
                " TPI_GERA_OBRIGACAO = 1) "
    End If
    Sql = Sql & " order by tip_sigla_imposto asc"
    Combo.Preencher Bdados, Sql, 1
End Sub

Private Function SelecionaContribuinte(Contribuinte As String, TipoInscricao As TipoInsc) As VSRecordset
    Dim Sql As String
    Select Case TipoInscricao
        Case 1 'IC Obrigatoria
            Sql = "Select tim_ic from Tab_Imovel "
            If Trim(Contribuinte) <> "" Then Sql = Sql & " where tim_ic = '" & Contribuinte & "'"
        Case 2 'IC nao Obrigatoria
            Sql = "Select tci_im from Tab_Contribuinte "
            If Trim(Contribuinte) <> "" Then Sql = Sql & " where tci_im = '" & Contribuinte & "'" ' AND tci_tipo_recolhimento_iss IN (1,2)"
    End Select
    
    Sql = "Select VIN_INSCRICAO from VIS_INSCRICAO where VIN_INSCRICAO = '" & Contribuinte & "'"
    If TipoInscricao = 3 Then
        Sql = "Select VIN_INSCRICAO from VIS_INSCRICAO where VIN_INSCRICAO like '" & Contribuinte & "%'"
        'feito para a geracao por grupos, setores etc...
    End If
    Call Bdados.AbreTabela(Sql, SelecionaContribuinte)
End Function

Public Function CriaObrigacao(ByVal Tributo As String, ByVal PeriodoInicial As String, ByVal PeriodoFinal As String, _
                Optional Contribuinte As String, Optional ValorObrigacao As Double = 0, Optional Status As TipoStatusObrigacao = -1, Optional SubstituirExistente As Boolean = True, Optional Vencimento As String, Optional Progresso As Object, Optional CodigoObrigacao As String, Optional PorGrupo As Byte, Optional ByRef UltGerado As Object) As String
                
    On Error GoTo trata
    Dim Sql As String
    Dim RsContrib As VSRecordset
    Dim RsTrib As VSRecordset
    Dim DataVence As String
    Dim Periodo As Double
    Dim TipoLancamento As TipoLanc
    Dim Regs As Double
    Dim ValorFixo As Double
    
    CriaObrigacao = 0
    PeriodoInicial = IIf(Len(PeriodoInicial) = 4, PeriodoInicial, Right(PeriodoInicial, 4) & Left(PeriodoInicial, 2))
    PeriodoFinal = IIf(Len(PeriodoFinal) = 4, PeriodoFinal, Right(PeriodoFinal, 4) & Left(PeriodoFinal, 2))
    ValorFixo = ValorObrigacao
    
    Sql = "SELECT tpi_tip_cod_imposto,tip_sigla_imposto,TPI_PERIODIC_CALCULO,TPI_PERIODIC_DECLARA, tpi_tipo_ic from " & _
            " TAB_IMPOSTO,TAB_PARAMETRO_IMPOSTO WHERE tpi_tip_cod_imposto = tip_cod_imposto AND tpi_ano_imposto ='" & _
            Imposto.BuscaAnoImposto(Tributo, Right(Periodo, 4)) & "' AND tpi_tip_cod_imposto ='" & Tributo & "'"
    If Bdados.AbreTabela(Sql, RsTrib) Then
        If Not Progresso Is Nothing Then Progresso.Min = 0
        RsTrib.MoveFirst
        Do While Not RsTrib.EOF
            ' problema do parcelamento é que tem que identificar o registro que está sendo gravado e guardar a inscrição certa
            ' senão a linha abaixo de selecionar contribuinte vai dar erro.
            
            If PorGrupo = 1 Then
                Set RsContrib = SelecionaContribuinte(Contribuinte, 3)
            Else
                Set RsContrib = SelecionaContribuinte(Contribuinte, Nvl("" & RsTrib!tpi_tipo_ic, 2))
            End If
            RsContrib.MoveFirst
            For Periodo = CDbl(PeriodoInicial) To CDbl(PeriodoFinal)
                If Trim(Vencimento) = "" Then
                    DataVence = Imposto.BuscaDataVencimento(RsTrib!tpi_tip_cod_imposto, CStr(IIf(RsTrib!TPI_PERIODIC_DECLARA = 1, Left(Periodo, 4), Right(CStr(Periodo), 2) & Left(CStr(Periodo), 4))))
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
                            ValorObrigacao = CalculaObrigacao(Trim(RsContrib(0)), Nvl("" & RsTrib!tpi_tipo_ic, 2), RsTrib!tpi_tip_cod_imposto, _
                                    RsTrib!tip_sigla_imposto, Periodo, DataVence, TipoLancamento)
                        Else
                            ValorObrigacao = ValorFixo
                        End If
                        
                        If ValorObrigacao > 0 Then Calculo.GravaObrigacao = True
                        CriaObrigacao = GravaObrigacao(Trim(RsContrib(0)), Nvl("" & RsTrib!tpi_tipo_ic, 2), RsTrib!tpi_tip_cod_imposto, Periodo, ValorObrigacao, _
                                        DataVence, TipoLancamento, IIf(Status > 0, Status, IIf(TipoLancamento = etlDeclaracao And ValorObrigacao = 0, 1, 2)), Progresso, SubstituirExistente, CodigoObrigacao)
                    End If
                    Regs = Regs + 1
                    If Not UltGerado Is Nothing Then UltGerado = "Registro nº " & Regs & "  -  " & RsContrib(0)
                    RsContrib.MoveNext
                    If Not Progresso Is Nothing Then Progresso.value = Regs
                    DoEvents
                Loop
                RsContrib.MoveFirst
            Next
            RsTrib.MoveNext
        Loop
    End If
    Bdados.FechaTabela RsContrib
    Bdados.FechaTabela RsTrib
    obCodigoObrigacao = CriaObrigacao
    Exit Function
trata:
    If Err.Number = 91 Then
        Resume Next
    Else
        Avisa Err.Description
    End If
End Function

Public Sub TrocaSitObrigacao(CodObrigacao As String, Situacao As TipoStatusObrigacao)
    Dim Valores As String
    Dim Campos As String
    
    Campos = "TOC_STATUS_OBRIGACAO" & IIf(Situacao < 1, ",TOC_STATUS_ANTERIOR_OBRIGACAO", "")
    Valores = Bdados.PreparaValor(IIf(Situacao < 1, etsCreditoOriginalAberto, Situacao))
    If Situacao < 1 Then Valores = Valores & Bdados.PreparaValor(etsCreditoOriginalAberto)
    Call Bdados.AtualizaDados("TAB_OBRIGACAO_CONTRIBUINTE", Valores, Campos, "TOC_COD_OBRIGACAO= " & CodObrigacao)
    Conta.TrocaStatusConta CodObrigacao, IIf(Situacao < 1, etsCreditoOriginalAberto, Situacao)
End Sub

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

Public Function GravaObrigacao(Inscricao As String, TipoInscricao As TipoInsc, CodTributo As String, _
                Periodo As Double, ValorObrigacao As Double, DataVence As String, Optional TipoLancamento As TipoLanc, _
                Optional StatusObrigacao As TipoStatusObrigacao = etsCreditoOriginalAberto, Optional Progresso As Object, Optional SubstituirExistente As Boolean = True, Optional CodigoObrigacao As String) As Double
                
    Dim Vencimento As String, Multa As Double
    Dim Juros As Double
    
    Dim Sql As String
    Dim RsImp As VSRecordset
    Dim Campos As String
    Dim Valores As String
    Dim CodObrigacao As String
    
    'STATUS OBRIGACAO
    '1: Credito nao lancado
    '2: Credito Original em aberto
    '3: Credito pago
    '4: Credito em parcelamento
    '5: Credito em Divida Ativa
    '6. Credito Isento
    '7. Credito Cancelado
    Campos = "TOC_COD_OBRIGACAO,TOC_INSCRICAO,TOC_TIP_COD_IMPOSTO,TOC_PERIODO,TOC_TIPO_INSCRICAO,TOC_DATA_VENCIMENTO,TOC_DATA_GERACAO,TOC_VALOR_OBRIGACAO,TOC_VALOR_MULTA,TOC_VALOR_JUROS,TOC_STATUS_OBRIGACAO,TOC_STATUS_ANTERIOR_OBRIGACAO,TOC_TIPO_LANCAMENTO"
    If Calculo.GravaObrigacao Then
        If CodigoObrigacao <> "" Then
            CodObrigacao = CodigoObrigacao
            GravaObrigacao = CodObrigacao
            'If Not SubstituirExistente Then Exit Function
        Else
            Campos = "TOC_COD_OBRIGACAO,TOC_INSCRICAO,TOC_TIP_COD_IMPOSTO,TOC_PERIODO,TOC_TIPO_INSCRICAO," & _
                    "TOC_DATA_VENCIMENTO,TOC_DATA_GERACAO,TOC_VALOR_OBRIGACAO," & _
                    "TOC_VALOR_MULTA,TOC_VALOR_JUROS,TOC_STATUS_OBRIGACAO,TOC_STATUS_ANTERIOR_OBRIGACAO,TOC_TIPO_LANCAMENTO"
            Sql = "Select TOC_COD_OBRIGACAO from TAB_OBRIGACAO_CONTRIBUINTE where TOC_INSCRICAO ='" & Trim(Inscricao) & "'"
            Sql = Sql & " and TOC_PERIODO =" & Periodo
            Sql = Sql & " and  TOC_TIP_COD_IMPOSTO ='" & CodTributo & "'"
            If Not Bdados.AbreTabela(Sql, RsImp) Then
                'CodObrigacao = Bdados.Correlativo("TRIB", Const_Obrig, "OBRIGACAO TRIBUTARIA")
                CodObrigacao = Conta.GeraCodPagamento(CodTributo)
            Else
                CodObrigacao = RsImp!TOC_COD_OBRIGACAO
                GravaObrigacao = CodObrigacao
                If Not SubstituirExistente Then Exit Function
            End If
        End If
        GravaObrigacao = CodObrigacao
        
        Sql = "SELECT tpi_tip_cod_imposto,tpi_ano_imposto,tpi_dias_pagar,tpi_periodic_calculo," & _
            "tpi_tipo_contribuinte,tpi_valor_min_multa,tpi_valor_max_multa,tpi_valor_juros," & _
            "tpi_aliquota,tpi_valor_taxa_fixa,TPI_TIPO_JUROS from TAB_PARAMETRO_IMPOSTO "
        If Bdados.AbreTabela(Sql, RsImp) Then
            Valores = Bdados.PreparaValor(CodObrigacao, Bdados.Converte(Trim(Inscricao), tctexto), CodTributo, Periodo, TipoInscricao, _
                    Bdados.Converte(DataVence, TCDataHora), Bdados.Converte(Format(Date, "dd/mm/yyyy"), _
                    TCDataHora), Bdados.Converte(ValorObrigacao, TCDuplo), Bdados.Converte(MultaObrigacao, TCDuplo), _
                    Bdados.Converte(JurosObrigacao, TCDuplo), StatusObrigacao, StatusObrigacao, TipoLancamento)
            Bdados.GravaDados "TAB_OBRIGACAO_CONTRIBUINTE", Valores, Campos, _
                    "TOC_INSCRICAO ='" & Trim(Inscricao) & "' AND TOC_TIP_COD_IMPOSTO ='" & CodTributo & _
                    "' AND TOC_PERIODO =" & Periodo
            GravaObrigacao = CodObrigacao
            Conta.CriaContaContribuinte CodObrigacao
        End If
        
        'Gera DAM da Obrigacao(Não gera Dam para Credito nao lancado)
'        If TipoLancamento <> etlDeclaracao Then
'            If ValorObrigacao > 0 Then GeraDAMObrigacao CodObrigacao
'        End If
        'If Not (TipoLancamento = etlDeclaracao And ValorObrigacao = 0) Then GeraDAMObrigacao CodObrigacao
        
        'GRAVA OS DETALHES DA OBRIGACAO(OBRIGAÇÕES ACESSORIAS)
        If Calculo.GravaObrigacaoAcessoria Then
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
    GravaObrigacao = Nvl(CodObrigacao, 0)
End Function

Public Function AlteraObrigacao(CodObrigacao As String, DataVence As String, ValorObrigacao As String, ValorMulta As String, ValorJuros As String, ValorTaxas As Double, Status As TipoStatusObrigacao) As Boolean
                
    Dim Sql As String
    Dim Campos As String
    Dim Valores As String
    
    
    Campos = "TOC_DATA_VENCIMENTO,TOC_VALOR_OBRIGACAO," & _
            "TOC_VALOR_MULTA,TOC_VALOR_JUROS,TOC_TOTAL_TAXA_INCLUSA,TOC_STATUS_OBRIGACAO,TOC_TPA_COD_PARCELAMENTO"
    Valores = Bdados.PreparaValor(Bdados.Converte(DataVence, TCDataHora), Bdados.Converte(Nvl(ValorObrigacao, 0), TCDuplo), _
                Bdados.Converte(Nvl(ValorMulta, 0), TCDuplo), Bdados.Converte(Nvl(ValorJuros, 0), TCDuplo), Bdados.Converte(ValorTaxas, TCDuplo), Status, 0)
    AlteraObrigacao = Bdados.AtualizaDados("TAB_OBRIGACAO_CONTRIBUINTE", Valores, Campos, "TOC_COD_OBRIGACAO ='" & CodObrigacao & "'")
    If AlteraObrigacao Then
       AlteraObrigacao = Conta.CriaContaContribuinte(CodObrigacao)
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

Public Function BuscaDetalheObrigacao(CodObrigacao As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsAux As VSRecordset
        
    BuscaDetalheObrigacao = False
    obContribuinte = ""
    Sql = "Select  * from TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO =" & CodObrigacao
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaDetalheObrigacao = True
        obCodigoObrigacao = CodObrigacao
        obDataVencimento = Rs!TOC_data_vencimento
        obCodImposto = Rs!TOC_TIP_COD_IMPOSTO
        obPeriodo = Rs!TOC_periodo
        obStatusObrigacao = Rs!TOC_STATUS_OBRIGACAO
        obValorObrigacao = Format(Rs!TOC_valor_OBRIGACAO, Const_Monetario)
        obValorMulta = Format(Rs!TOC_Valor_multa, Const_Monetario)
        obValorJuros = Format(Rs!TOC_valor_juros, Const_Monetario)
        obTipoInscricao = Rs!TOC_TIPO_INSCRICAO
        obNumParcelamento = Nvl("" & Rs!TOC_TPA_COD_PARCELAMENTO, 0)
        If Rs!TOC_TIPO_INSCRICAO = 1 Then  'TRIBUTO DE IMOVEL
            obInscCadastral = Trim(Rs!TOC_INSCRICAO)
            obContribuinte = Trim(Rs!TOC_INSCRICAO)
            Sql = "Select tim_tci_im from tab_imovel where tim_ic = '" & Rs!TOC_INSCRICAO & "'" 'BUSCA PROPRIETÁRIO
            If Bdados.AbreTabela(Sql, RsAux) Then
                obInscMunicipal = RsAux!tim_tci_im
            Else
                obInscMunicipal = "0"
            End If
            Bdados.FechaTabela RsAux
        Else
            If Len(Trim(Rs!TOC_INSCRICAO)) <> 11 Then
                Sql = "Select tim_tci_im from tab_imovel where tim_ic = '" & Rs!TOC_INSCRICAO & "'" 'BUSCA PROPRIETÁRIO
                If Bdados.AbreTabela(Sql, RsAux) Then
                    obInscMunicipal = RsAux!tim_tci_im
                End If
                Bdados.FechaTabela RsAux
            Else
                obInscMunicipal = Trim(Rs!TOC_INSCRICAO)
            End If
            obContribuinte = Trim(Rs!TOC_INSCRICAO)
        End If
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
            If Trim("" & Rs!TGT_TIP_COD_IMPOSTO) <> "11120200" Then
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
                CodTributo = "" & Rs!TGT_TIP_COD_IMPOSTO
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
            Bdados.AtualizaDados "TAB_DARM_RECEBIDO", Valores, Campos, " TDR_TGT_COD_PAGAMENTO=" & Rs!tdr_tgt_cod_pagamento
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

