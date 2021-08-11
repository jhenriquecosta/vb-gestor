VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VsIptuBoaVista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Padroes
Public Iptu As New VSIptu
Private ProfundidadeMinima As Double
Private ProfundidadeMaxima As Double
Private TestadaReferencia As Double
Private ValorBasicoUnitario As Double
Private CustoBasicoUnitario As Double
Private ValorMetroTSU As Double
Private ValorMinimoBase As Double
Private ValorMinimoImposto As Double
Private ValorImposto As Double
Private GeraParcelas As Boolean
'Aliquotas
Private AliquotaTerreno As Double
Private AliquotaEdificacao As Double
Private AliquotaComercial As Double

Private AliquotaIndustria As Double
Private AliquotaPrestacao As Double
Private AliquotaOutros As Double
Private Pavimentos As Integer

'Variaveis Auxiliares
Private AreaEdificadaTotalCondominio As Double
Private AreaTerrenoTotalCondominio As Double
Private Lote As typImovel
Public AnoLancamento As Double

'Auxiliares
Private CodImposto As String
Private SiglaImposto As String
Private NomeImposto As String
Public VencimentoCotaUnica As String
Private VencimentoParcela() As String
Private DescontoParcela() As Double
Public RegistrosProcessados As Double
Private Parcelas As Byte
Private Bairro As String
Private Zona As Integer
Private Logr As String
Private Trecho As String

Private AreaTributavelTerreno As Double
Private SomatorioPontos As Double

Private Secao As String
Private Setor As Integer
Private AliquotaTipologia As Double
Private Tipologia As Integer, Localizacao As Integer, Parede As Integer
Private Cobertura As Integer, Revestimento As Integer, Esquadria As Integer
Private Estrutura As Integer, Padrao As Integer


Public GeraCotaUnica As Boolean
Private DadosCalculo As typDadosCalculo
Private Sub CalculaAreaTributavel()
    Dim Fracao As Double
    AreaTributavelTerreno = Lote.AreaLote
    If Lote.Apartamento Then
        If Lote.AreaConstruidaTotal > 0 Then Fracao = Lote.AreaConstruidaUnidade(1) / Lote.AreaLote
        Lote.AreaLote = Lote.AreaLote * Fracao
        AreaTributavelTerreno = Lote.AreaLote
    Else
        If Lote.AreaLote <= 200 And Lote.Profundidade > 20 Then
            AreaTributavelTerreno = Lote.TestadaLote * 20 + (Lote.AreaLote - Lote.TestadaLote * 20) * 0.5
        ElseIf Lote.AreaLote > 200 Then
            AreaTributavelTerreno = Lote.TestadaLote * 20 + (Lote.AreaLote - Lote.TestadaLote * 20) * 0.1
        End If
    End If
End Sub

Private Sub CalculaSomatorioPontos(Tipologia As Integer, Localizacao As Integer, Parede As Integer, _
                                Cobertura As Integer, Revestimento As Integer, Esquadria As Integer)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim MaxPontuacao As Double
    Dim PontosAcumulados As Double
    'TCS_MAX_PONTUACAO_TIPOLOGIA,TCS_VALOR
    SomatorioPontos = 0
    If (Esquadria = Localizacao = Cobertura = Parede) And Parede = 0 Then
        SomatorioPontos = 0.75
        Exit Sub
    End If
    Sql = "SELECT * from TAB_CUB_SINFAZ "
    Sql = Sql & " where TCS_TIPOLOGIA=" & Tipologia & " and ("
    If Esquadria <> 0 Then Sql = Sql & "  TCS_ESQUADRIA= " & Esquadria
    If Localizacao <> 0 Then Sql = Sql & " or TCS_LOCALIZACAO = " & Localizacao
    If Parede <> 0 Then Sql = Sql & " or TCS_PAREDE = " & Parede
    If Cobertura <> 0 Then Sql = Sql & " or TCS_COBERTURA = " & Cobertura
    If Cobertura <> 0 Then Sql = Sql & "  or TCS_REVESTIMENTO = " & Revestimento
    Sql = Sql & ")"
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        MaxPontuacao = Rs!TCS_MAX_PONTUACAO_TIPOLOGIA
        Do
            PontosAcumulados = PontosAcumulados + Rs!TCS_VALOR
            If CInt(Nvl("" & Rs!TCS_LOCALIZACAO, 0)) > 0 Then
                DadosCalculo.PontoLocalizacao = Rs!TCS_VALOR
            ElseIf CInt(Nvl("" & Rs!TCS_PAREDE, 0)) > 0 Then
                DadosCalculo.PontoParede = Rs!TCS_VALOR
            ElseIf CInt(Nvl("" & Rs!TCS_COBERTURA, 0)) > 0 Then
                DadosCalculo.PontoCobertura = Rs!TCS_VALOR
            ElseIf CInt(Nvl("" & Rs!TCS_REVESTIMENTO, 0)) > 0 Then
                DadosCalculo.PontoRevestimento = Rs!TCS_VALOR
            ElseIf CInt(Nvl("" & Rs!TCS_ESQUADRIA, 0)) > 0 Then
                DadosCalculo.PontoEsquadria = Rs!TCS_VALOR
            End If
            Rs.MoveNext
        Loop While Not Rs.EOF
    End If
    DadosCalculo.SomatorioPontos = PontosAcumulados
    DadosCalculo.MaximoPontos = MaxPontuacao
    SomatorioPontos = PontosAcumulados / MaxPontuacao
    DadosCalculo.RelacaoSomaMaximo = SomatorioPontos
End Sub

Private Sub CalculaValorVenalTerreno()
    
    
    If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
        CalculaAreaTributavel
    Else
        AreaTributavelTerreno = Lote.AreaLote
    End If
    If Lote.FatorTopografia = 0 Then Lote.FatorTopografia = 1
    If Lote.FatorPedologia = 0 Then Lote.FatorPedologia = 1
    If Lote.FatorSituacao = 0 Then Lote.FatorSituacao = 1
    
    ValorBasicoUnitario = CalculaVBU(Logr, Secao) 'Valor Basico Unitario
    If ValorBasicoUnitario = 0 Then ValorBasicoUnitario = 1
    Lote.ValorVenalTerreno = AreaTributavelTerreno * Lote.FatorTopografia * Lote.FatorPedologia * _
                ValorBasicoUnitario * Lote.FatorSituacao
                
    DadosCalculo.ValorM2Terreno = ValorBasicoUnitario
    DadosCalculo.AreaTributavelTerreno = AreaTributavelTerreno
    DadosCalculo.FatorTopografia = Lote.FatorTopografia
    DadosCalculo.FatorPedologia = Lote.FatorPedologia
    DadosCalculo.FatorSituacao = Lote.FatorSituacao
    DadosCalculo.ValorTerreno = Lote.ValorVenalTerreno
    DadosCalculo.AreaLote = Lote.AreaLote
    DadosCalculo.TestadaPrincipal = Lote.TestadaLote
End Sub

Private Sub CalculaValorVenalEdificacao(Unidade As Integer)
    On Error GoTo Trata
    ReDim Preserve Lote.ValorVenalEdificacao(1 To Unidade) As Double
    If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
        If Tipologia = 31 Or Tipologia = 50 Then
            CustoBasicoUnitario = CalculaCUB(Tipologia, 9, Setor)
        Else
            CustoBasicoUnitario = CalculaCUB(Tipologia, 39, Setor)
        End If
    Else
        CustoBasicoUnitario = CalculaCUB(Tipologia, 39, 0)
    End If
    If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
        CalculaSomatorioPontos Tipologia, Localizacao, Parede, Cobertura, Revestimento, Esquadria
    Else
        CustoBasicoUnitario = 1
    End If
    If Lote.FatorEstrutura = 0 Then Lote.FatorEstrutura = 1
    If Lote.FatorConservacao = 0 Then Lote.FatorConservacao = 1
    If Lote.FatorPadrao = 0 Then Lote.FatorPadrao = 1
    
    Lote.ValorVenalEdificacao(Unidade) = IIf(Lote.Apartamento, Lote.AreaConstruidaUnidade(Unidade), Lote.AreaConstruidaTotal) * _
        Lote.FatorEstrutura * Lote.FatorConservacao * CustoBasicoUnitario * Lote.FatorPadrao * SomatorioPontos
    Lote.ValorVenalEdificacoes = Lote.ValorVenalEdificacao(Unidade)
    
    DadosCalculo.ValorEdificacao = Lote.ValorVenalEdificacoes
    DadosCalculo.ValorM2Construcao = CustoBasicoUnitario
    DadosCalculo.AreaConstruida = Lote.AreaConstruidaTotal
    DadosCalculo.FatorEstrutura = Lote.FatorEstrutura
    DadosCalculo.FatorEstado = Lote.FatorConservacao
    DadosCalculo.FatorPadrao = Lote.FatorPadrao
    Exit Sub
    
Trata:
    Lote.ValorVenalEdificacao(Unidade) = 0
End Sub

Public Sub InicializaValores(Imovel As String)
    'Busca valores gerais para calculo do IPTU
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim I As Byte
    Sql = "Select tip_cod_imposto,tip_sigla_imposto, tip_nome_imposto from tab_imposto " & _
        " where tip_sigla_imposto = '" & Imposto.NomeTributo(ttr_IPTU) & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        CodImposto = Rs!tip_cod_imposto
        SiglaImposto = Rs!TIP_SIGLA_IMPOSTO
        NomeImposto = Rs!tip_nome_imposto
    End If
    Bdados.FechaTabela Rs
    'Busca Aliquotas de terreno, edificacao, comercio, industria, prestacao de servicos e outros
        
    VencimentoCotaUnica = TrocaPic(Iptu.BuscaParamentroIptu("VENCIMENTO COTAS", 1, CStr(AnoLancamento)), ".", ",")
'    CodImpostoITU = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU))
    Parcelas = 0
    Lote.Isento = False
    ReDim Preserve VencimentoParcela(0 To 12) As String
    ReDim Preserve DescontoParcela(0 To 12) As Double
'    Dim ParcelasAtivas As String
    Sql = "Select TPP_PARCELA,TPP_VENCIMENTO,TPP_DESCONTO from TAB_PARAMETRO_PARCELA_IPTU WHERE TPP_ANO=" & Me.AnoLancamento & " ORDER BY TPP_PARCELA ASC"
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do
            VencimentoParcela(Rs!TPP_PARCELA) = Rs!TPP_VENCIMENTO
            DescontoParcela(Rs!TPP_PARCELA) = Rs!TPP_DESCONTO
            Parcelas = Parcelas + 1
            Rs.MoveNext
        Loop While Not Rs.EOF
        VencimentoCotaUnica = VencimentoParcela(0)
        Parcelas = IIf(Parcelas > 1, Parcelas - 1, Parcelas)
    End If
    RegistrosProcessados = 0
End Sub

Public Function CalculaValorIptu(Optional GeraSomenteValorVenal As Boolean = False) As Double
    Dim Unidade As Integer
    Dim ValorTotalConstrucao As Double
    Dim LoteComercial As Boolean
    Dim AliquotaEspecifica As Double
    Dim Aliquota As Double
    Dim I As Byte
    Dim j As Byte
    On Error GoTo Trata
    Lote.ValorImpostoEdificacao = 0
    ValorTotalConstrucao = 0
    Lote.ValorImpostoTerreno = 0
    If Right(Lote.Inscricao, 3) = 200 Then Exit Function
    If AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Then
        Aliquota = AliquotaTipologia
    Else
        Aliquota = BuscaAliquota(AnoLancamento, IIf(Lote.AreaConstruidaTotal = 0, 2, 1), CInt(Lote.Destino(1)), Lote.AreaLote, Setor)
    End If
    DadosCalculo.Aliquota = Aliquota
    
    If Not Lote.Apartamento Then Lote.ValorImpostoTerreno = Lote.ValorVenalTerreno * (Aliquota / 100)
    Lote.ValorImpostoEdificacao = (Lote.ValorVenalEdificacoes * (Aliquota / 100))
    DadosCalculo.ValorImposto = Lote.ValorImpostoTerreno + Lote.ValorImpostoEdificacao
    
    If Lote.Cobranca = tcbNormal Or Lote.Cobranca = tcbIsentoImposto Then 'Pode calcular taxas
        CalculaTaxasServico
    End If
    'AREA DE GRAVACAO DA PARCELA UNICA
    GravaValorVenalImovel
    If GeraSomenteValorVenal Then
        DadosCalculo.ValorImposto = 0
        Exit Function
    End If
    GravaLog
    DoEvents
    
    If GeraCotaUnica Then 'Verifica Base de Calculo
        LoteComercial = False
        'ANALISO SE O LOTE É USADO COMERCIALMENTE(QUALQUER UNIDADE), SE SIM ESTE NAO POSSUI ISENCAO
        For I = 1 To UBound(Lote.Destino)
            If (Lote.Destino(I) <> tcoGleba) And (Lote.Destino(I) <> tcoResidencial) Then
                LoteComercial = True
                Exit For
            End If
        Next
        If (Lote.Cobranca < tcbIsentoTotal) Or Lote.ValorImpostoEdificacao = 0 Then
            'Se Valor Venal menor que base de calculo minima = isento
            'Restricoes: Tem que ser lote edificado e residencial
            If (((CDbl(Lote.ValorVenalTerreno) + CDbl(Lote.ValorVenalEdificacoes)) < ValorMinimoBase) Or _
                    ((CDbl(Lote.ValorImpostoEdificacao) + CDbl(Lote.ValorImpostoTerreno)) < ValorMinimoImposto)) And _
                    Lote.TipoLote = tloResidencial And Lote.ValorImpostoEdificacao > 0 And Not LoteComercial Then
                Iptu.GravaIsento CodImposto, 1, CDbl(Lote.ValorImpostoTerreno) + CDbl(Lote.ValorImpostoEdificacao), CDbl(AnoLancamento)
                Lote.ValorImpostoEdificacao = 0
                Lote.ValorImpostoTerreno = 0
            End If
            If Trim(VencimentoParcela(0)) <> "" Then GravaDadosIptu CDbl(AnoLancamento), VencimentoParcela(0), 0
        Else
            'É isento total ou imune
            Iptu.GravaIsento CodImposto, Lote.Cobranca, CDbl(Lote.ValorImpostoTerreno) + CDbl(Lote.ValorImpostoEdificacao), CDbl(AnoLancamento)
        End If
    End If
    'AREA DE CALCULO DOS VALORES DAS PARCELAS. VALOR TOTAL DO IMPOSTO DIVIDIDO PELA
    'QUANTIDADE DE PARCELAS.
    'FOI NECESSARIO APENAS DIVIDIR OS VALORES-BASE PELA QUANTID. DE PARCELAS E A FUNCAO
    'GravaDadosIptu FEZ O RESTO
    If CalculaValorIptu = 0 Then CalculaValorIptu = Lote.ValorImpostoEdificacao + Lote.ValorImpostoTerreno + Lote.ValorTaxas
    ValorImposto = CalculaValorIptu
    If (GeraParcelas And (Lote.Cobranca < tcbIsentoTotal)) And (Lote.ValorImpostoEdificacao + Lote.ValorImpostoTerreno <> 0) Then
        If Lote.Cobranca = tcbNormal Or Lote.Cobranca = tcbIsentoImposto Then
            Lote.ValorImpostoEdificacao = Lote.ValorImpostoEdificacao / Parcelas
            Lote.ValorImpostoTerreno = Lote.ValorImpostoTerreno / Parcelas
            Lote.ValorTaxas = Lote.ValorTaxas / Parcelas
            Lote.ValorTaxaLixo = Lote.ValorTaxaLixo / Parcelas
            Lote.ValorTaxaConservacao = Lote.ValorTaxaConservacao / Parcelas
            Lote.ValorTaxaLimpeza = Lote.ValorTaxaLimpeza / Parcelas
            For I = 1 To 12
                If Trim(VencimentoParcela(I)) <> "" Then
                    GravaDadosIptu CDbl(AnoLancamento), VencimentoParcela(I), CInt(I)
                Else
                    Exit For
                End If
            Next
        End If
    End If
 
    
    RegistrosProcessados = RegistrosProcessados + 1
    Lote.Isento = False
    Exit Function
Trata:
    Avisa Err.Description
    Exit Function
    Resume
End Function

Private Sub CalculaTaxasServico()
    Lote.ValorTaxas = 1
    DadosCalculo.ValorTaxas = Lote.ValorTaxas
End Sub

Private Sub GravaDadosIptu(Periodo As Double, Vencimento As String, Parcela As Integer)

End Sub

Private Function CalculaVBU(Logradouro As String, Secao As String) As Double
    'Atraves dos nome do logradouro e do trecho do imovel, encontra o valor
    'basico unitario na tabela TAB_TRECHO
    Dim Sql As String
    Dim Rs As VSRecordset
    
    If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
        If Trim(Secao) = "" Then Exit Function
    End If
    If Trim(Logradouro) = "" Then Exit Function
    Sql = "Select TTC_VALOR from TAB_TRECHO where TTC_TLG_COD_LOGRADOURO ='" & Logradouro & _
        "'"
        If Trim(Secao) <> "" Then
            Sql = Sql & " and TTC_SEQ_TRECHO =" & Left(Secao, IIf(Len(Secao) - 1 <= 0, 1, Len(Secao) - 1)) & " and TTC_COD_TRECHO ='" & Right(Secao, 1) & "'"
        End If
    If Bdados.AbreTabela(Sql, Rs) Then
        CalculaVBU = Nvl("" & Rs!TTC_VALOR, 0)
    End If
End Function


Private Sub BuscaLocalizacaoImovel(ByVal Imovel As String, ByRef Bairro As String, ByRef Logradouro As String, Secao As String, Setor As Integer)
    'Procedimento que retorna o cod. do bairro, cod. do logradouro e nome
    'do trecho em que o lote se encontra na tabela TAB_TRECHO
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsEndereco As VSRecordset
    Sql = "Select tim_tci_im,tim_ic_auxiliar,tim_numero,tim_tlg_cod_logradouro,TIM_TBA_COD_BAIRRO,tim_complemento,ttl_nome,tlg_nome,tba_nome,tim_secao from vis_imovel" & _
        " where tim_ic = '" & Imovel & "' "
    If Bdados.AbreTabela(Sql, Rs) Then
        Lote.Contribuinte = Trim("" & Rs!tim_tci_im)
        Bairro = Trim("" & Rs!tim_tba_cod_bairro)
        Logradouro = Trim("" & Rs!tim_tlg_cod_logradouro)
        Secao = Trim("" & Rs!tim_secao)
        If Len(Trim("" & Rs!tim_ic_auxiliar)) > 3 Then Setor = Mid(Trim("" & Rs!tim_ic_auxiliar), 3, 2)
        Lote.Endereco = "" & Rs!ttl_nome & " " & Rs!tlg_nome & " " & Rs!tim_numero & " " & Rs!tim_complemento & " " & Rs!tba_nome
     End If
End Sub

Public Function CarregaDetalheLote(Imovel As String) As Boolean
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Unidade As Integer
    Dim Aux As Integer
    Dim RsAux As VSRecordset
    Dim RsEng As VSRecordset
    Dim InscImobiliaria As String
    Dim Englobado As Boolean
    Static Cobs As Integer
'    ApagaParcelasDuplicadas Imovel
    'ZERA OS VETORES
    ReDim Lote.ValorVenalEdificacao(1 To 1) As Double
    ReDim Lote.AreaConstruidaUnidade(1 To 1) As Double
    ReDim Lote.FracaoIdealLoteComum(1 To 1) As Double
    ReDim Lote.Destino(1 To 1) As TipoDestino
    On Error GoTo Trata
    Bairro = ""
    Logr = ""
    Trecho = ""
    '-------------------------------------------------
    Imovel = Trim(Imovel)
    Lote.Inscricao = Imovel
    Aux = 0
    CarregaDetalheLote = True
    Englobado = True
    'Verifica se há algum pagamento deste periodo. Se existir, o usuario deve estornar este(s) pagamentos primeiro.
    Sql = " SELECT TDR_TGT_COD_PAGAMENTO FROM TAB_DARM_RECEBIDO WHERE (TDR_INSCRICAO ='" & Lote.Inscricao & _
        "' AND TDR_TIPO_INSCRICAO = 1) AND TDR_PERIODO =" & AnoLancamento & " AND TDR_TIP_COD_IMPOSTO ='" & CodImposto & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        CarregaDetalheLote = False
        Exit Function
    End If
    'APAGA AS ISENCOES DO LOTE
    If CInt(Right(Imovel, 3)) = 200 Or CInt(Right(Imovel, 3)) = 600 Then Exit Function 'SE FOR BT DO CONDOMINIO (200), NAO FAZ NADA
    
    Sql = "Select tdi_tgc_cod_grupo,tdi_valor_item,tdi_tco_cod_componente,tdi_tim_ic_unidade from tab_detalhe_imovel where "
    Sql = Sql & " tdi_tim_ic = '" & Imovel & "'" ' and  ((tdi_tgc_cod_grupo >= 1 and tdi_tgc_cod_grupo <= 25) " & _
    "  or  (tdi_tgc_cod_grupo >= 100 and tdi_tgc_cod_grupo <= 115)) "
    If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
        Sql = Sql & " and  ((tdi_tgc_cod_grupo >= 1 and tdi_tgc_cod_grupo <= 25) "
        Sql = Sql & "or  (tdi_tgc_cod_grupo >= 100 and tdi_tgc_cod_grupo <= 115))"
    End If
    Sql = Sql & "  order by  tdi_tim_ic_unidade,tdi_tgc_cod_grupo "
    If Bdados.AbreTabela(Sql, Rs) Then
        Lote.Inscricao = Imovel
        DadosCalculo.Inscricao = Imovel
        Sql = "Select tim_tipo_imovel,tim_ted_cod_edificio,tim_ic_auxiliar from tab_imovel where tim_ic ='" & Imovel & "'"
        If Bdados.AbreTabela(Sql, RsAux) Then
            Lote.TipoLote = RsAux!tim_tipo_imovel
            InscImobiliaria = RsAux!tim_ic_auxiliar
            Lote.Apartamento = IIf(CInt(Nvl("" & RsAux!tim_ted_cod_edificio, 0)), True, False)
        End If
        'pega sub unidades
        Sql = "Select  tim_ic from tab_imovel where tim_ic_auxiliar='" & InscImobiliaria & "'" 'Left(InscImobiliaria, Len(InscImobiliaria) - 5) & "20000'"
        If Bdados.AbreTabela(Sql, RsEng) Then
            InscImobiliaria = RsEng!tim_ic
        End If
        Sql = "SELECT TOC_COD_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_INSCRICAO ='" & InscImobiliaria & "' AND TOC_PERIODO =" & AnoLancamento & " AND TOC_TIP_COD_IMPOSTO ='" & CodImposto & "'"
        If Bdados.AbreTabela(Sql, RsEng) Then
              If Nvl("" & RsEng!TOC_COD_OBRIGACAO, 0) > 0 Then Englobado = False
        End If
        If Right(Imovel, 3) > 200 Then
            Lote.TipoLote = tloResidencial
        End If
        
        Rs.MoveFirst
        Aux = Rs!tdi_tim_ic_unidade
        Unidade = 0
        Do While Not Rs.EOF
            If Aux <> Rs!tdi_tim_ic_unidade And (Rs!tdi_tim_ic_unidade <> 0 And Rs!tdi_tim_ic_unidade <> 200) Then
                If (Aux <> 0 And Aux <> 200) Then
                    'Calcula Fatores
                    'Calcula Valores Venais
                    CalculaValorVenalEdificacao Unidade
                End If
                Unidade = Unidade + 1
                Aux = Rs!tdi_tim_ic_unidade
            End If
            If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
                Select Case Rs!tdi_tgc_cod_grupo
                    Case 1 'Ocupacao
                        Lote.Ocupacao = Rs!tdi_tco_cod_componente
                    Case 3 'Cod. Cobranca
                        Lote.Cobranca = Rs!tdi_tco_cod_componente
                    Case 6 'Situacao
                        CalculaFatorTabelado tftSituacao, Rs!tdi_tco_cod_componente, Unidade
                    Case 7 'Tografia
                        CalculaFatorTabelado tftTopografia, Rs!tdi_tco_cod_componente, Unidade
                    Case 8 'Pedologia
                        CalculaFatorTabelado tftPedologia, Rs!tdi_tco_cod_componente, Unidade
                    Case 9 'Tipologia
                        Tipologia = Rs!tdi_tco_cod_componente
                        CalculaFatorTabelado 9, Rs!tdi_tco_cod_componente, Unidade
                    Case 10 'Estrutura
                        Estrutura = Rs!tdi_valor_item
                        CalculaFatorTabelado tftEstrutura, Rs!tdi_tco_cod_componente, Unidade
                    Case 11 'Tipo de Construcao(Destino)
                        BuscaDestinoLote Rs!tdi_tco_cod_componente, IIf(Unidade = 0, 1, Unidade)
                    Case 12 'Padrao
                        CalculaFatorTabelado tftPadrao, Rs!tdi_tco_cod_componente, Unidade
                    Case 13 'Conservacao
                        CalculaFatorTabelado tftConservacao, Rs!tdi_tco_cod_componente, Unidade
                    Case 20 'Localizacao
                        Localizacao = Rs!tdi_tco_cod_componente
                    Case 22 'Paredes
                        Parede = Rs!tdi_tco_cod_componente
                    Case 23 'Cobertura
                        Cobertura = Rs!tdi_tco_cod_componente
                    Case 24 'Revestimento
                        Revestimento = Rs!tdi_tco_cod_componente
                    Case 25 'Esquadrias
                        Esquadria = Rs!tdi_tco_cod_componente
                    Case 100 'Testada
                        Lote.TestadaLote = Rs!tdi_valor_item
                    Case 108 'Area Lote
                        Lote.AreaLote = Rs!tdi_valor_item
                    Case 110
                        Pavimentos = Rs!tdi_valor_item
                    Case 111 'Ano da Construcao
                        Lote.AnoConstrucao = Rs!tdi_valor_item
                    Case 112 'Area Construida Unidade
                        Unidade = IIf(Unidade = 0, 1, Unidade)
                        ReDim Preserve Lote.AreaConstruidaUnidade(1 To Unidade) As Double
                        Lote.AreaConstruidaUnidade(Unidade) = Format(Rs!tdi_valor_item, Const_Monetario)
                        If Not Englobado Then Lote.AreaConstruidaTotal = Format(Rs!tdi_valor_item, Const_Monetario)
                    Case 113 'Area Construida Total
                        If Englobado Then Lote.AreaConstruidaTotal = Format(Rs!tdi_valor_item, Const_Monetario)
                    Case 114
                        SomatorioPontos = Rs!tdi_valor_item
                    Case 115
                        Lote.Profundidade = Format(Rs!tdi_valor_item, Const_Monetario)
                End Select
            Else
                Select Case Rs!tdi_tgc_cod_grupo
                    Case 26 'Ocupacao
                        Lote.Ocupacao = Rs!tdi_tco_cod_componente
                    Case 43 'Situacao
                        CalculaFatorTabelado Rs!tdi_tgc_cod_grupo, Rs!tdi_tco_cod_componente, Unidade
                    Case 44 'Tografia
                        CalculaFatorTabelado Rs!tdi_tgc_cod_grupo, Rs!tdi_tco_cod_componente, Unidade
                    Case 45 'Pedologia
                        CalculaFatorTabelado Rs!tdi_tgc_cod_grupo, Rs!tdi_tco_cod_componente, Unidade
                    Case 39 'Tipologia
                        Tipologia = Rs!tdi_tco_cod_componente
                        CalculaFatorTabelado Rs!tdi_tgc_cod_grupo, Rs!tdi_tco_cod_componente, Unidade
                    Case 78 'Estrutura
                        Estrutura = Rs!tdi_valor_item
                        CalculaFatorTabelado Rs!tdi_tgc_cod_grupo, Rs!tdi_tco_cod_componente, Unidade
                    Case 77 'Tipo de Construcao(Destino)
                        BuscaDestinoLote Rs!tdi_tco_cod_componente, IIf(Unidade = 0, 1, Unidade)
                    Case 81 'Padrao
                        CalculaFatorTabelado tftPadrao, Rs!tdi_tco_cod_componente, Unidade
                    Case 92 'Conservacao
                        CalculaFatorTabelado Rs!tdi_tgc_cod_grupo, Rs!tdi_tco_cod_componente, Unidade
                    Case 100 'Testada
                        Lote.TestadaLote = Rs!tdi_valor_item
                    Case 108 'Area Lote
                        Lote.AreaLote = Rs!tdi_valor_item
                    Case 110
                        Pavimentos = Rs!tdi_valor_item
                    Case 111 'Ano da Construcao
                        Lote.AnoConstrucao = Rs!tdi_valor_item
                    Case 112 'Area Construida Unidade
                        Unidade = IIf(Unidade = 0, 1, Unidade)
                        ReDim Preserve Lote.AreaConstruidaUnidade(1 To Unidade) As Double
                        Lote.AreaConstruidaUnidade(Unidade) = Format(Rs!tdi_valor_item, Const_Monetario)
                        If Not Englobado Then Lote.AreaConstruidaTotal = Format(Rs!tdi_valor_item, Const_Monetario)
                    Case 113 'Area Construida Total
                        If Englobado Then Lote.AreaConstruidaTotal = Format(Rs!tdi_valor_item, Const_Monetario)
                    Case 114
                        SomatorioPontos = Rs!tdi_valor_item
                    Case 115
                        Lote.Profundidade = Format(Rs!tdi_valor_item, Const_Monetario)
                End Select
            End If
            Rs.MoveNext
        Loop
    Else
        CarregaDetalheLote = False
        Exit Function
    End If
    '
    'Calcula Fatores para a ultima unidade construida
    BuscaLocalizacaoImovel Lote.Inscricao, Bairro, Logr, Secao, Setor
    If (Unidade <> 0 And Unidade <> 200) Then
        'Calcula Valores Venais
        CalculaValorVenalEdificacao Unidade
    End If
    'Fim Fatores
'    If Lote.TestadaLote > 0 Then
'        Lote.Profundidade = Format(Lote.AreaLote / Lote.TestadaLote, Const_Monetario)
'    Else
'        Lote.Profundidade = 0
'    End If
    DadosCalculo.Profundidade = Lote.Profundidade
    'calculo do valor do terreno não leva em consideração se é ou não gleba, trata tudo igualmente - ederson - balsas
    'If Lote.TipoLote = tcoGleba Then ' por ederson -> ' Or Lote.TipoLote = tloGleba Then
    '    Call CalculaValorVenalGleba
    'Else
    CalculaValorVenalTerreno
    Exit Function
Trata:
    Avisa Err.Description
    Exit Function
    Resume '
End Function


Private Function CalculaCUB(Tipologia As Integer, Grupo As Integer, Setor As Integer) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "select TCS_VALOR,TCS_VALOR_UFM from TAB_COMPONENTE_AVANCADO_SETOR where " & _
        " TCS_tco_cod_componente = " & Tipologia & " and TCS_tco_grupo =" & Grupo & _
        " and TCS_SETOR =" & Setor
    If Bdados.AbreTabela(Sql, Rs) Then
        CalculaCUB = Rs!TCS_VALOR
    End If
    
End Function
Private Sub CalculaFatorTabelado(Fator As TipoFatorTabelado, Opcao As Integer, Unidade As Integer)
    'Busca na Tab_Componente_Avancado valores tabelados para os componentes
    'Conservacao, Estrutura, Pedologia, Situacao, Topografia
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "Select tco_valor from tab_componente_avancado where tco_grupo = " & Fator & " and tco_cod_componente = " & Opcao
    If Bdados.AbreTabela(Sql, Rs) Then
        
        If (Fator = 6 And Opcao = 6) Then
            ReDim Preserve Lote.Destino(1 To IIf(Unidade = 0, 1, Unidade)) As TipoDestino
            Lote.Destino(IIf(Unidade = 0, 1, Unidade)) = tloGleba
            Lote.TipoLote = tloGleba
        End If
        If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Then
            Select Case Fator
                Case tftConservacao
                    Lote.FatorConservacao = Rs!tco_valor
                Case tftEstrutura
                    Lote.FatorEstrutura = Rs!tco_valor
                    Lote.VidaUtil = Rs!tco_valor
                Case tftPedologia
                    Lote.FatorPedologia = Rs!tco_valor
                Case tftSituacao
                    Lote.FatorSituacao = Rs!tco_valor
                Case tftTopografia
                    Lote.FatorTopografia = Rs!tco_valor
                Case tftPadrao
                    Lote.FatorPadrao = Rs!tco_valor
                Case 9
                    AliquotaTipologia = Rs!tco_valor
            End Select
        Else
            Select Case Fator
                Case 92
                    Lote.FatorConservacao = Nvl("" & Rs!tco_valor, 0)
                Case 78
                    Lote.FatorEstrutura = Nvl("" & Rs!tco_valor, 0)
                    Lote.VidaUtil = Nvl("" & Rs!tco_valor, 0)
                Case 45
                    Lote.FatorPedologia = Nvl("" & Rs!tco_valor, 0)
                Case 43
                    Lote.FatorSituacao = Nvl("" & Rs!tco_valor, 0)
                Case 44
                    Lote.FatorTopografia = Nvl("" & Rs!tco_valor, 0)
                Case 81
                    Lote.FatorPadrao = Nvl("" & Rs!tco_valor, 0)
                Case 39
                    AliquotaTipologia = Nvl("" & Rs!tco_valor, 0)
            End Select
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub BuscaDestinoLote(Tipo As Integer, Unidade As Integer)
    Dim Sql As String
    Dim Rs As VSRecordset
    'Tipos de Construcao = Residencial, Comercial, Terreno, Gleba
    ReDim Preserve Lote.Destino(1 To Unidade) As TipoDestino
    
    If Lote.Destino(Unidade) <> tcoGleba Then
        'A definicao de gleba foi concebida na funcao CalculaFatorTabelado(Fator = 6,Opcao = 6)
        'Caso o lote seja mesmo uma gleba, o tipo de construcao já está setado. Caso contrario é setado aqui
        If Right(Lote.Inscricao, 3) > 200 Then
            Sql = "Select tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic = '" & Lote.Inscricao & _
                "' and tdi_tgc_cod_grupo = 11"
        Else
            Sql = "Select tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic = '" & Lote.Inscricao & _
                "' and tdi_tim_ic_unidade = " & Unidade & " and tdi_tgc_cod_grupo = 11"
        End If
        If Bdados.AbreTabela(Sql, Rs) Then
            Lote.Destino(Unidade) = Rs!tdi_valor_item
        End If
    End If
End Sub

Private Function BuscaAliquota(Ano As Double, TipoImovel As TipoImovel, Optional Destinacao As Integer = 0, Optional Area As Double, Optional Setor As Integer) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
'
'    If TipoImovel = tim_TerrenoSemEdif Then
'        Sql = "SELECT TPA_ALIQUOTA FROM TAB_PARAMETRO_IPTU_ALIQUOTA WHERE TPA_ANO =" & Ano & " AND TPA_TIPO_IMOVEL = " & TipoImovel
'    Else
'
        Sql = "SELECT TPA_ALIQUOTA FROM TAB_PARAMETRO_IPTU_ALIQUOTA WHERE TPA_ANO >=" & Ano & " AND TPA_TIPO_IMOVEL = " & IIf(Lote.AreaConstruidaTotal = 0, 2, 1)
        If Lote.AreaConstruidaTotal = 0 Then
            If Lote.AreaLote <= 200 Then
                Sql = Sql & " and tpa_setor =" & IIf(Setor = 1 Or Setor = 2 Or Setor = 6 Or Setor = 7, Setor, 0)
            Else
                If TipoImovel = tim_Gleba Then
                    Sql = Sql & " AND TPA_DESTINACAO = 59"
                Else
                    Sql = Sql & " AND (TPA_LIM_INF_AREA < " & Bdados.Converte(Area, TCDuplo) & " AND (TPA_LIM_SUP_AREA > " & Bdados.Converte(Area, TCDuplo) & " OR TPA_LIM_SUP_AREA = 0))"
                End If
            End If
        Else
            Sql = Sql & " AND (TPA_DESTINACAO = " & IIf(Destinacao = 25 Or Destinacao = 33 Or Destinacao = 86, Destinacao, 0) & ")"
        End If
'        Sql = Sql & " AND (TPA_LIM_INF_AREA < " & Bdados.Converte(Area, TCDuplo) & " AND (TPA_LIM_SUP_AREA > " & Bdados.Converte(Area, TCDuplo) & " OR TPA_LIM_SUP_AREA = 0))"
'        If TipoImovel = 2 Then
'            Sql = Sql & " AND (TPA_SETOR = " & IIf(Setor = 1 Or Setor = 2 Or Setor = 6 Or Setor = 7, Setor, 0) & ")"
'        Else
'            Sql = Sql & " AND (TPA_SETOR = 0)"
'        End If
'        If TipoImovel = 2 Then
'            Sql = Sql & " AND (TPA_DESTINACAO = " & IIf(Destinacao = 59, Destinacao, 0) & ")"
'        Else
'            Sql = Sql & " AND (TPA_DESTINACAO = " & IIf(Destinacao = 25 Or Destinacao = 33 Or Destinacao = 86, Destinacao, 0) & ")"
'        End If
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaAliquota = Rs!TPA_ALIQUOTA
    End If
End Function


Private Sub GravaLog()
    Dim Campos As String
    Dim Valores As String
    
    Campos = "TLC_TIM_IC,TLC_ANO,TLC_AREA_TERRENO,TLC_AREA_TERRENO_TRIBUTAVEL,TLC_VALOR_M2_TERRENO," & _
            "TLC_VALOR_M2_CONSTRUCAO,TLC_FATOR_TOPOGRAIFA,TLC_FATOR_PEDOLOGIA,TLC_FATOR_SITUACAO," & _
            "TLC_FATOR_ESTADO,TLC_FATOR_ESTRUTURA ,TLC_FATOR_PADRAO,TLC_SOMATORIO_PONTOS,TLC_MAXIMO_PONTOS," & _
            "TLC_RELACAO_SOMA_MAXIMO,TLC_ALIQUOTA,TLC_VALOR_TERRENO,TLC_VALOR_EDIFICACAO,TLC_AREA_LOTE," & _
            "TLC_PROFUNDIDADE_MEDIA,TLC_AREA_EDIFICADA,TLC_PONTO_LOCALIZACAO,TLC_PONTO_PAREDE," & _
            "TLC_PONTO_COBERTURA,TLC_PONTO_REVESTIMENTO,TLC_PONTO_ESQUADRIA,TLC_VALOR_IMPOSTO," & _
            "TLC_VALOR_TAXAS,TLC_TESTADA_PRINCIPAL"
    Valores = Bdados.PreparaValor(DadosCalculo.Inscricao, DadosCalculo.Ano, DadosCalculo.AreaLote, DadosCalculo.AreaTributavelTerreno, _
                DadosCalculo.ValorM2Terreno, DadosCalculo.ValorM2Construcao, DadosCalculo.FatorTopografia, _
                DadosCalculo.FatorPedologia, DadosCalculo.FatorSituacao, DadosCalculo.FatorEstado, _
                DadosCalculo.FatorEstrutura, DadosCalculo.FatorPadrao, DadosCalculo.SomatorioPontos, _
                DadosCalculo.MaximoPontos, DadosCalculo.RelacaoSomaMaximo, DadosCalculo.Aliquota, _
                DadosCalculo.ValorTerreno, DadosCalculo.ValorEdificacao, DadosCalculo.AreaLote, _
                DadosCalculo.Profundidade, DadosCalculo.AreaConstruida, DadosCalculo.PontoLocalizacao, _
                DadosCalculo.PontoParede, DadosCalculo.PontoCobertura, DadosCalculo.PontoRevestimento, _
                DadosCalculo.PontoEsquadria, DadosCalculo.ValorImposto, DadosCalculo.ValorTaxas, _
                DadosCalculo.TestadaPrincipal)
    Bdados.GravaDados "TAB_LOG_CALCULO_IPTU", Valores, Campos, "TLC_TIM_IC ='" & DadosCalculo.Inscricao & _
                            "' and TLC_ANO =" & DadosCalculo.Ano
End Sub


Private Sub GravaValorVenalImovel()
    Dim Unidade As Integer
    Dim ValorTotalConstrucao As Double
    'GRAVA RESUMO DE VALORES DO LOTE (VALOR EDIFICACAO, DO TERRENO E FRACAO IDEAL
    
    ValorTotalConstrucao = Lote.ValorVenalEdificacoes
    Lote.ValorVenalEdificacoes = ValorTotalConstrucao
    Bdados.AtualizaDados "TAB_IMOVEL", Bdados.PreparaValor(ValorTotalConstrucao, _
        Lote.ValorVenalTerreno, ValorTotalConstrucao + Lote.ValorVenalTerreno), _
        "TIM_VALOR_EDIFIC,TIM_VALOR_TERRENO,TIM_VALOR", "TIM_IC ='" & Lote.Inscricao & "'"
End Sub


