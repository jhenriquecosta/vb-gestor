VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VsIptuBarraMansa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Iptu As New VSIptu
Private ProfundidadeMinima As Double
Private ProfundidadeMaxima As Double
Public AnoLancamento As Double
Public Type typDadosImovel
    Inscricao As String
    Ano As Integer
    AreaTerreno As Double
    NumeroFrentes As Integer
    ValorM2Terreno As Double
    ValorM2Construcao As Double
    FatorG_Profundidade As Double
    FatorK As Double
    FatorC_Topografia_Topologia As Double
    FatorF_Frentes As Double
    Aliquota As Double
    ValorTerreno As Double
    ValorEdificacaoRegular As Double
    ValorEdificacaoIrregular As Double
    AreaConstruidaRegular As Double
    AreaConstruidaIrregular As Double
    FatorDC_Conservacao As Double
    FatorAltura As Double
    FracaoIdeal As Double
    FatorTipoSubTipo As Double
    TipoSubTipo As Integer
    ValorImposto As Double
    TestadaPrincipal As Double
    Tipologia As Integer
    CodigoCobranca As Integer
    FatorCobranca As Double
    Caracterizacao As Integer
    Categoria As Integer
    SomatorioMelhoramentos As Double
    Padrao As Integer
    RevestimentoExterno As Integer
    RevestimentoInterno As Integer
    AcabamentoExterno As Integer
    AcabamentoInterno As Integer
    Piso As Integer
    InstalacaoSanitaria As Integer
    InstalacaEletrica As Integer
    Forro As Integer
    Estrutura As Integer
    Cobertura As Integer
    Esquadria As Integer
    Elevador As Integer
    Desconto As Double
End Type
Private DadosCalculo As typDadosImovel
Private Lote As typImovel
Public Sub CalculaValorTerreno()
    If DadosCalculo.FatorG_Profundidade = 0 Then DadosCalculo.FatorG_Profundidade = 1
    If DadosCalculo.FatorK = 0 Then DadosCalculo.FatorK = 1
    If DadosCalculo.FatorF_Frentes = 0 Then DadosCalculo.FatorF_Frentes = 1
    If DadosCalculo.FatorC_Topografia_Topologia = 0 Then DadosCalculo.FatorC_Topografia_Topologia = 1
    If DadosCalculo.FracaoIdeal = 0 Then DadosCalculo.FracaoIdeal = 1
    Lote.ValorVenalTerreno = Format(DadosCalculo.FatorG_Profundidade * DadosCalculo.FatorK * _
                             DadosCalculo.FatorF_Frentes * DadosCalculo.FatorC_Topografia_Topologia * _
                             DadosCalculo.FracaoIdeal, Const_Monetario)
    DadosCalculo.ValorTerreno = Lote.ValorVenalTerreno
End Sub

Private Sub CalculaFatorF_Frentes()
    Select Case DadosCalculo.NumeroFrentes
        Case Is < 2
            DadosCalculo.FatorF_Frentes = 1
        Case 2
            DadosCalculo.FatorF_Frentes = 1.2
        Case 3
            DadosCalculo.FatorF_Frentes = 1.5
        Case Else
            DadosCalculo.FatorF_Frentes = 2
    End Select
End Sub

Private Sub CalculaFatorDC_Conservacao()
    Dim Sql As String
    Dim rs As VSRecordset
    
    
    Sql = "SELECT TCO_VALOR  FROM TAB_COMPONENTE_AVANCADO WHERE TCO_GRUPO = 28 " & _
    " AND TCO_COD_COMPONENTE = " & DadosCalculo.FatorDC_Conservacao
    If Bdados.AbreTabela(Sql, rs) Then
        DadosCalculo.FatorDC_Conservacao = rs!tco_valor
    End If
End Sub

Private Sub CalculaFatorDesconto()
    Dim Sql As String
    Dim rs As VSRecordset
    
    
    Sql = "SELECT TCO_VALOR  FROM TAB_COMPONENTE_AVANCADO WHERE TCO_GRUPO = 30 " & _
    " AND TCO_COD_COMPONENTE = " & DadosCalculo.CodigoCobranca
    If Bdados.AbreTabela(Sql, rs) Then
        DadosCalculo.FatorCobranca = rs!tco_valor
    End If
End Sub

Private Sub CalculaFatorAltura()
    Dim Sql As String
    Dim rs As VSRecordset
    
    
    Sql = "SELECT TCO_VALOR  FROM TAB_COMPONENTE_AVANCADO WHERE TCO_GRUPO = 35 " & _
    " AND TCO_COD_COMPONENTE = " & DadosCalculo.FatorAltura
    If Bdados.AbreTabela(Sql, rs) Then
        DadosCalculo.FatorAltura = rs!tco_valor
    End If
End Sub


Private Sub CalculaTipoSubTipo()
    Dim Sql As String
    Dim rs As VSRecordset
    
    
    Sql = "SELECT TCO_VALOR  FROM TAB_COMPONENTE_AVANCADO WHERE TCO_GRUPO = 29 " & _
    " AND TCO_COD_COMPONENTE = " & DadosCalculo.TipoSubTipo
    If Bdados.AbreTabela(Sql, rs) Then
        DadosCalculo.FatorTipoSubTipo = rs!tco_valor
    End If
End Sub

Private Sub CalculaFatorC_Topografia_Pedologia()
    Dim Sql As String
    Dim rs As VSRecordset
    
    
    Sql = "SELECT TCO_VALOR  FROM TAB_COMPONENTE_AVANCADO WHERE TCO_GRUPO = 12 " & _
    " AND TCO_COD_COMPONENTE = " & DadosCalculo.FatorC_Topografia_Topologia
    If Bdados.AbreTabela(Sql, rs) Then
        DadosCalculo.FatorC_Topografia_Topologia = rs!tco_valor
    End If
End Sub


Public Sub CalculaValorEdificacao()

    If DadosCalculo.FatorTipoSubTipo = 0 Then DadosCalculo.FatorTipoSubTipo = 1
    If DadosCalculo.FatorDC_Conservacao = 0 Then DadosCalculo.FatorDC_Conservacao = 1
    If DadosCalculo.FatorAltura = 0 Then DadosCalculo.FatorAltura = 1
    DadosCalculo.ValorEdificacaoRegular = DadosCalculo.FatorTipoSubTipo * DadosCalculo.AreaConstruidaRegular * _
                             DadosCalculo.ValorM2Construcao * DadosCalculo.FatorDC_Conservacao * _
                             IIf(DadosCalculo.FatorAltura = 0, 1, DadosCalculo.FatorAltura)
    DadosCalculo.ValorEdificacaoIrregular = DadosCalculo.AreaConstruidaIrregular * DadosCalculo.ValorM2Construcao
    Lote.ValorVenalEdificacoes = Format(DadosCalculo.ValorEdificacaoRegular + DadosCalculo.ValorEdificacaoIrregular, Const_Monetario)
    
End Sub

Private Sub CalculaFatorG()
    DadosCalculo.FatorG_Profundidade = Format(((DadosCalculo.AreaTerreno * DadosCalculo.TestadaPrincipal) / 30) ^ (0.5), Const_Monetario)
End Sub

Private Sub BuscaValorMetroTerreno(CodLogradouro As String)
    Dim Sql As String
    Dim rs As VSRecordset
    
    Sql = "SELECT ttc_valor,ttc_peso FROM TAB_TRECHO WHERE ttc_tlg_cod_logradouro ='" & CodLogradouro & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        DadosCalculo.ValorM2Terreno = Nvl("" & rs!TTC_VALOR, 0)
        DadosCalculo.FatorK = Nvl("" & rs!ttc_peso, 0)
    End If
End Sub
Private Sub CalculaFatorK()
    DadosCalculo.FatorK = (DadosCalculo.ValorM2Terreno * 30)
End Sub

Public Function CalculaValorIptu(Optional GeraSomenteValorVenal As Boolean = False) As Double
    Dim Unidade As Integer
    Dim ValorTotalConstrucao As Double
    Dim LoteComercial As Boolean
    Dim AliquotaEspecifica As Double
    Dim Aliquota As Double
    Dim i As Byte
    Dim j As Byte
    On Error GoTo Trata
    Lote.ValorImpostoEdificacao = 0
    ValorTotalConstrucao = 0
    Lote.ValorImpostoTerreno = 0
'    Aliquota = BuscaAliquota(AnoLancamento, IIf(DadosCalculo.AreaConstruidaIrregular + DadosCalculo.AreaConstruidaIrregular = 0, 2, 1), 0, DadosCalculo.AreaTerreno)
    If (DadosCalculo.AreaConstruidaRegular + DadosCalculo.AreaConstruidaIrregular) = 0 Then
        Aliquota = TrocaPic(Iptu.BuscaParamentroIptu("ALIQUOTAS", talAliquotaTerreno, Me.AnoLancamento), ".", ",")
    Else
        Aliquota = TrocaPic(Iptu.BuscaParamentroIptu("ALIQUOTAS", talAliquotaEdificacao, Me.AnoLancamento), ".", ",")
    End If
    DadosCalculo.Aliquota = Aliquota
    
    If Not Lote.Apartamento Then Lote.ValorImpostoTerreno = Format(Lote.ValorVenalTerreno * (Aliquota / 100), Const_Monetario)
    Lote.ValorImpostoEdificacao = Format((Lote.ValorVenalEdificacoes * (Aliquota / 100)), Const_Monetario)
    DadosCalculo.ValorImposto = Format(Lote.ValorImpostoTerreno + Lote.ValorImpostoEdificacao, Const_Monetario)
    
    DadosCalculo.ValorImposto = DadosCalculo.ValorImposto * DadosCalculo.FatorCobranca
    If AnoLancamento >= 2002 Then
        If DadosCalculo.ValorImposto <= 60 Then
            DadosCalculo.ValorImposto = DadosCalculo.ValorImposto * 0.85
            DadosCalculo.Desconto = 0.15
        ElseIf DadosCalculo.ValorImposto > 61 And DadosCalculo.ValorImposto <= 100 Then
            DadosCalculo.ValorImposto = DadosCalculo.ValorImposto * 0.9
            DadosCalculo.Desconto = 0.1
        ElseIf DadosCalculo.ValorImposto > 101 And DadosCalculo.ValorImposto <= 200 Then
            DadosCalculo.ValorImposto = DadosCalculo.ValorImposto * 0.95
            DadosCalculo.Desconto = 0.05
        End If
    End If
    
    DadosCalculo.ValorImposto = Format(DadosCalculo.ValorImposto, Const_Monetario)
    'AREA DE GRAVACAO DA PARCELA UNICA
'    GravaValorVenalImovel
    If GeraSomenteValorVenal Then
        DadosCalculo.ValorImposto = 0
        Exit Function
    End If
'    GravaLog
    DoEvents
    
    If GeraCotaUnica Then 'Verifica Base de Calculo
        LoteComercial = False
        'ANALISO SE O LOTE É USADO COMERCIALMENTE(QUALQUER UNIDADE), SE SIM ESTE NAO POSSUI ISENCAO
        For i = 1 To UBound(Lote.Destino)
            If (Lote.Destino(i) <> tcoGleba) And (Lote.Destino(i) <> tcoResidencial) Then
                LoteComercial = True
                Exit For
            End If
        Next
        If (Lote.Cobranca < tcbIsentoTotal) Or Lote.ValorImpostoEdificacao = 0 Then
            'Se Valor Venal menor que base de calculo minima = isento
            'Restricoes: Tem que ser lote edificado e residencial
            If (DadosCalculo.ValorImposto < ValorMinimoBase Or _
                    ((DadosCalculo.ValorImposto) < ValorMinimoImposto)) And _
                    Lote.TipoLote = tloResidencial And Lote.ValorImpostoEdificacao > 0 And Not LoteComercial Then
                Iptu.GravaIsento CodImposto, 1, DadosCalculo.ValorImposto, CDbl(AnoLancamento)
                Lote.ValorImpostoEdificacao = 0
                Lote.ValorImpostoTerreno = 0
            End If
'            If Trim(VencimentoParcela(0)) <> "" Then GravaDadosIptu CDbl(AnoLancamento), VencimentoParcela(0), 0
        Else
            'É isento total ou imune
            Iptu.GravaIsento CodImposto, Lote.Cobranca, DadosCalculo.ValorImposto, CDbl(AnoLancamento)
        End If
    End If
    If CalculaValorIptu = 0 Then CalculaValorIptu = DadosCalculo.ValorImposto + Lote.ValorTaxas
    ValorImposto = CalculaValorIptu
'    If (GeraParcelas And (Lote.Cobranca < tcbIsentoTotal)) And (DadosCalculo.ValorImposto <> 0) Then
'        If Lote.Cobranca = tcbNormal Or Lote.Cobranca = tcbIsentoImposto Then
'            For i = 1 To 12
'                If Trim(VencimentoParcela(i)) <> "" Then
'                    GravaDadosIptu CDbl(AnoLancamento), VencimentoParcela(i), CInt(i)
'                Else
'                    Exit For
'                End If
'            Next
'        End If
'    End If
    Call GravaLog
    RegistrosProcessados = RegistrosProcessados + 1
    Lote.Isento = False
    Exit Function
Trata:
    Avisa Err.Description
    Exit Function
    Resume
End Function

Public Function GravaLog() As Boolean
    Dim Campos                       As String
    Dim Valores                      As String
    Dim Condicao                     As String
    
    Campos = " TPB_INSCRICAO,TPB_ANO,TPB_AREA_TERRENO,TPB_NUMERO_FRENTES,TPB_VALOR_M2_TERRENO,"
    Campos = Campos & " TPB_VALOR_M2_CONTRUCAO,TPB_FATOR_G_PROFUNDIDADE,TPB_FATOR_K,TPB_FATOR_C_TOPOGR_TOPOLOGIA,"
    Campos = Campos & " TPB_FATOR_F_FRENTES,TPB_ALIQUOTA,TPB_VALOR_TERRENO,TPB_VALOR_EDIFICACAO_REGULAR,"
    Campos = Campos & " TPB_VALOR_EDIFICACAO_IRREGULAR,TPB_AREA_CONTRUCAO_REGULAR,TPB_AREA_CONTRUCAO_IRREGULAR,"
    Campos = Campos & " TPB_FATOR_CD_CONSERVACAO,TPB_FATOR_ALTURA,TPB_FRACAO_IDEAL,TPB_FATOR_TIPO_SUBTIPO,"
    Campos = Campos & " TPB_VALOR_IMPOSTO,TPB_TESTADA_PRINCIPAL,TPB_TIPOLOGIA,TPB_CARACTERIZACAO,"
    Campos = Campos & " TPB_CATEGORIA,TPB_SOMATORIO_MELHORAMENTO,TPB_PADRAO,TPB_REVESTIMENTO_EXTERNO,"
    Campos = Campos & " TPB_REVESTIMENTO_INTERNO,TPB_ACABAMENTO_EXTERNO,TPB_ACABAMENTO_INTERNO,"
    Campos = Campos & " TPB_PISO,TPB_INSTALACAO_SANATARIA,TPB_INSTALACAO_ELETRICA,TPB_FORRO,"
    Campos = Campos & " TPB_ESTRUTURA,TPB_COBERTURA,TPB_ESQUADRIA,TPB_ELEVADOR,TPB_TIPO_COBRANCA,TPB_FATOR_COBRANCA,TPB_DESCONTO"
    Valores = Bdados.Converte(DadosCalculo.FracaoIdeal, TCDuplo)
    Valores = Bdados.PreparaValor(Bdados.Converte(DadosCalculo.Inscricao, tctexto), Bdados.Converte(DadosCalculo.Ano, tctexto), Bdados.Converte(DadosCalculo.AreaTerreno, TCMonetario), _
    Bdados.Converte(DadosCalculo.NumeroFrentes, TCMonetario), Bdados.Converte(DadosCalculo.ValorM2Terreno, TCMonetario), Bdados.Converte(DadosCalculo.ValorM2Construcao, TCMonetario), Bdados.Converte(DadosCalculo.FatorG_Profundidade, TCMonetario), Bdados.Converte(DadosCalculo.FatorK, TCMonetario), Bdados.Converte(DadosCalculo.FatorC_Topografia_Topologia, TCMonetario), _
    Bdados.Converte(DadosCalculo.FatorF_Frentes, TCMonetario), Bdados.Converte(DadosCalculo.Aliquota, TCMonetario), Bdados.Converte(DadosCalculo.ValorTerreno, TCMonetario), Bdados.Converte(DadosCalculo.ValorEdificacaoRegular, TCMonetario), Bdados.Converte(DadosCalculo.ValorEdificacaoIrregular, TCMonetario), _
    Bdados.Converte(DadosCalculo.AreaConstruidaRegular, TCMonetario), Bdados.Converte(DadosCalculo.AreaConstruidaIrregular, TCMonetario), Bdados.Converte(DadosCalculo.FatorDC_Conservacao, TCMonetario), Bdados.Converte(DadosCalculo.FatorAltura, TCMonetario), Bdados.Converte(DadosCalculo.FracaoIdeal, TCDuplo), Bdados.Converte(DadosCalculo.FatorTipoSubTipo, TCMonetario), _
    Bdados.Converte(DadosCalculo.ValorImposto, TCMonetario), Bdados.Converte(DadosCalculo.TestadaPrincipal, TCMonetario), Bdados.Converte(DadosCalculo.TipoSubTipo, TCMonetario), Bdados.Converte(DadosCalculo.Caracterizacao, TCMonetario), _
    Bdados.Converte(DadosCalculo.Categoria, TCMonetario), Bdados.Converte(Format(DadosCalculo.SomatorioMelhoramentos, Const_Monetario), TCMonetario), Bdados.Converte(DadosCalculo.Padrao, TCMonetario), Bdados.Converte(DadosCalculo.RevestimentoExterno, TCMonetario), _
    Bdados.Converte(DadosCalculo.RevestimentoInterno, TCMonetario), Bdados.Converte(DadosCalculo.AcabamentoExterno, TCMonetario), Bdados.Converte(DadosCalculo.AcabamentoInterno, TCMonetario), _
    Bdados.Converte(DadosCalculo.Piso, TCMonetario), Bdados.Converte(DadosCalculo.InstalacaoSanitaria, TCMonetario), Bdados.Converte(DadosCalculo.InstalacaEletrica, TCMonetario), Bdados.Converte(DadosCalculo.Forro, TCMonetario), _
    Bdados.Converte(DadosCalculo.Estrutura, TCMonetario), Bdados.Converte(DadosCalculo.Cobertura, TCMonetario), Bdados.Converte(DadosCalculo.Esquadria, TCMonetario), _
    Bdados.Converte(DadosCalculo.Elevador, TCMonetario), DadosCalculo.CodigoCobranca, Bdados.Converte(DadosCalculo.FatorCobranca, TCMonetario), Bdados.Converte(DadosCalculo.Desconto, TCMonetario))

    Condicao = "TPB_INSCRICAO = '" & DadosCalculo.Inscricao & "' and  TPB_ANO  = '" & DadosCalculo.Ano & "'"
    GravaLog = Bdados.GravaDados("TAB_PGV_BARRA", Valores, Campos, Condicao)
End Function
Private Function BuscaAliquota(Ano As Double, TipoImovel As TipoImovel, Optional Destinacao As Integer = 0, Optional Area As Double, Optional Setor As Integer) As Double
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Iptu As New VSIptu
    
    
End Function


Private Sub CalculaCategoria()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Condicao As String
    '17  Revestimento EXTERNO
    '18  Revestimento INTERNO
    '19  ACABAMENTO EXTERNO
    '20  ACABAMENTO INTERNO
    '21  PISO
    '22  INSTALACAO ELETRICA
    '23  INSTALACAO SANITARIA
    '24  FORRO
    '25  Estrutura
    '26  Cobertura
    '27  ESQUADRIAS
    '36  ELEVADOR
    Condicao = " ((TCO_GRUPO =17  AND TCO_COD_COMPONENTE = " & DadosCalculo.RevestimentoExterno & _
    " ) OR (TCO_GRUPO = 18  AND TCO_COD_COMPONENTE = " & DadosCalculo.RevestimentoInterno & _
    ") OR (TCO_GRUPO = 19  AND TCO_COD_COMPONENTE = " & DadosCalculo.AcabamentoExterno & _
    ") OR (TCO_GRUPO = 20 AND TCO_COD_COMPONENTE = " & DadosCalculo.AcabamentoInterno & _
    ") OR (TCO_GRUPO = 21  AND TCO_COD_COMPONENTE = " & DadosCalculo.Piso & _
    ") OR (TCO_GRUPO = 22  AND TCO_COD_COMPONENTE = " & DadosCalculo.InstalacaEletrica & _
    ") OR (TCO_GRUPO = 23 AND TCO_COD_COMPONENTE = " & DadosCalculo.InstalacaoSanitaria & _
    ") OR (TCO_GRUPO = 24  AND TCO_COD_COMPONENTE = " & DadosCalculo.Forro & _
    ") OR (TCO_GRUPO = 25  AND TCO_COD_COMPONENTE = " & DadosCalculo.Estrutura & _
    ") OR (TCO_GRUPO = 26  AND TCO_COD_COMPONENTE = " & DadosCalculo.Cobertura & _
    ") OR (TCO_GRUPO = 27  AND TCO_COD_COMPONENTE = " & DadosCalculo.Esquadria & _
    ") OR (TCO_GRUPO = 36  AND TCO_COD_COMPONENTE = " & DadosCalculo.Elevador & _
    ")) AND TCO_TIPOLOGIA =" & DadosCalculo.Caracterizacao
    
    Sql = "SELECT SUM(TCO_VALOR)  FROM TAB_COMPONENTE_AVANCADO_PADRAO WHERE " & Condicao
    If Bdados.AbreTabela(Sql, rs) Then
        DadosCalculo.SomatorioMelhoramentos = Nvl("" & rs(0), 0)
    End If
    Select Case DadosCalculo.Caracterizacao
        Case 2, 3
            If DadosCalculo.SomatorioMelhoramentos <= 250 Then
                DadosCalculo.Padrao = 2
            ElseIf DadosCalculo.SomatorioMelhoramentos > 250 And DadosCalculo.SomatorioMelhoramentos <= 340 Then
                DadosCalculo.Padrao = 3
            ElseIf DadosCalculo.SomatorioMelhoramentos > 340 And DadosCalculo.SomatorioMelhoramentos <= 430 Then
                DadosCalculo.Padrao = 4
            ElseIf DadosCalculo.SomatorioMelhoramentos > 430 Then
                DadosCalculo.Padrao = 5
            End If
        Case Else
            If DadosCalculo.SomatorioMelhoramentos <= 150 Then
                DadosCalculo.Padrao = 1 'PRECARIO
            ElseIf DadosCalculo.SomatorioMelhoramentos > 150 And DadosCalculo.SomatorioMelhoramentos <= 250 Then
                DadosCalculo.Padrao = 2 'POPULAR
            ElseIf DadosCalculo.SomatorioMelhoramentos > 250 And DadosCalculo.SomatorioMelhoramentos <= 340 Then
                DadosCalculo.Padrao = 3 'MEDIO
            ElseIf DadosCalculo.SomatorioMelhoramentos > 340 And DadosCalculo.SomatorioMelhoramentos <= 430 Then
                DadosCalculo.Padrao = 4 'FINO
            ElseIf DadosCalculo.SomatorioMelhoramentos > 430 Then
                DadosCalculo.Padrao = 5 'LUXO
            End If
    End Select
    DadosCalculo.Categoria = DadosCalculo.Padrao
    Sql = "SELECT * FROM cptrib.tab_cub_padrao WHERE TCP_TIPO_SUB_TIPO =" & DadosCalculo.Caracterizacao
    If Bdados.AbreTabela(Sql, rs) Then
        DadosCalculo.ValorM2Construcao = rs(DadosCalculo.Padrao)
    End If
    
End Sub

Public Function CarregaDetalheLote(Imovel As String) As Boolean
    Dim rs As VSRecordset
    Dim Sql As String
    Dim Unidade As Integer
    Dim Aux As Integer
    Dim RsAux As VSRecordset
    Dim RsEng As VSRecordset
    Dim InscImobiliaria As String
    Dim Englobado As Boolean
    Dim CodLogradouro As String
    Static Cobs As Integer
'    ApagaParcelasDuplicadas Imovel
    'ZERA OS VETORES
    On Error GoTo Trata
    Imovel = Trim(Imovel)
    Lote.Inscricao = Imovel
    Aux = 0
    CarregaDetalheLote = True
    Englobado = True
    
    Sql = "Select tdi_tgc_cod_grupo,tdi_valor_item,tdi_tco_cod_componente,tdi_tim_ic_unidade from tab_detalhe_imovel where "
    Sql = Sql & " tdi_tim_ic = '" & Imovel & "' and tdi_tgc_cod_grupo > 11 "
    Sql = Sql & "  order by  tdi_tim_ic_unidade,tdi_tgc_cod_grupo "
    If Bdados.AbreTabela(Sql, rs) Then
        Lote.Inscricao = Imovel
        DadosCalculo.Inscricao = Imovel
        Sql = "Select tim_tipo_imovel,tim_ted_cod_edificio,tim_ic_auxiliar,tim_tlg_cod_logradouro from tab_imovel where tim_ic ='" & Imovel & "' "
        If Bdados.AbreTabela(Sql, RsAux) Then
            Lote.TipoLote = RsAux!tim_tipo_imovel
            InscImobiliaria = RsAux!tim_ic_auxiliar
            Lote.Apartamento = IIf(CInt(Nvl("" & RsAux!tim_ted_cod_edificio, 0)), True, False)
            CodLogradouro = RsAux!tim_tlg_cod_logradouro
        End If
        
        rs.MoveFirst
        Aux = rs!tdi_tim_ic_unidade
        Unidade = 0
        Do While Not rs.EOF
            Select Case rs!tdi_tgc_cod_grupo
                Case 12 'Topografia/Pedologia(FATOR C)
                    DadosCalculo.FatorC_Topografia_Topologia = rs!tdi_tco_cod_componente
                Case 17 'REVESTIMENTO EXTERNO
                    DadosCalculo.RevestimentoExterno = rs!tdi_valor_item
                Case 18 'REVESTIMENTO INTERNO
                    DadosCalculo.RevestimentoInterno = rs!tdi_valor_item
                Case 19 'ACABAMENTO EXTERNO
                    DadosCalculo.AcabamentoExterno = rs!tdi_valor_item
                Case 20 'ACABAMENTO INTERNO
                    DadosCalculo.AcabamentoInterno = rs!tdi_valor_item
                Case 21 'PISO
                    DadosCalculo.Piso = rs!tdi_valor_item
                Case 22 'INSTALACAO ELETRICA
                    DadosCalculo.InstalacaEletrica = rs!tdi_valor_item
                Case 23 'INSTALACAO SANITARIA
                    DadosCalculo.InstalacaoSanitaria = rs!tdi_valor_item
                Case 24 'FORRO
                    DadosCalculo.Forro = rs!tdi_valor_item
                Case 25 'ESTRUTURA
                    DadosCalculo.Estrutura = rs!tdi_valor_item
                Case 26 'COBERTURA
                    DadosCalculo.Cobertura = rs!tdi_valor_item
                Case 27 'ESQUADRIAS
                    DadosCalculo.Esquadria = rs!tdi_valor_item
                Case 36 'ELEVADOR
                    DadosCalculo.Elevador = rs!tdi_valor_item
                Case 28 'Conservacao
                    DadosCalculo.FatorDC_Conservacao = rs!tdi_valor_item
                Case 29 'Tipo/SubTipo
                    DadosCalculo.TipoSubTipo = rs!tdi_valor_item
                Case 30 'Codigo Cobranca
                    DadosCalculo.CodigoCobranca = rs!tdi_valor_item
                Case 35 'Altura
                    DadosCalculo.FatorAltura = rs!tdi_valor_item
                Case 37 'Caracterizacao
                    DadosCalculo.Caracterizacao = rs!tdi_valor_item
                Case 100 'Testada
                    DadosCalculo.TestadaPrincipal = rs!tdi_valor_item
                Case 107 'No. Frentes
                    DadosCalculo.NumeroFrentes = rs!tdi_valor_item
                Case 108 'Area Lote
                    DadosCalculo.AreaTerreno = rs!tdi_valor_item
                Case 115 'Profundidade
                    DadosCalculo.FatorG_Profundidade = Format(rs!tdi_valor_item, Const_Monetario)
                Case 112 'Area Construida Regular
                    DadosCalculo.AreaConstruidaRegular = Format(rs!tdi_valor_item, Const_Monetario)
                Case 122 'Area Construida Irregular
                    DadosCalculo.AreaConstruidaIrregular = Format(rs!tdi_valor_item, Const_Monetario)
                Case 123 'Fracao
                    DadosCalculo.FracaoIdeal = rs!tdi_valor_item
            End Select
            rs.MoveNext
        Loop
    Else
        CarregaDetalheLote = False
        Exit Function
    End If
    BuscaValorMetroTerreno CodLogradouro
    CalculaTipoSubTipo
    CalculaCategoria
    CalculaFatorDC_Conservacao
    CalculaFatorF_Frentes
    CalculaFatorG
    CalculaFatorC_Topografia_Pedologia
    CalculaFatorDesconto
    CalculaFatorAltura
    
    If DadosCalculo.FatorK = 0 Then CalculaFatorK
    DadosCalculo.Ano = AnoLancamento
    CalculaValorTerreno
    CalculaValorEdificacao
    Exit Function
Trata:
    Avisa Err.Description
    Exit Function
    Resume '
End Function




