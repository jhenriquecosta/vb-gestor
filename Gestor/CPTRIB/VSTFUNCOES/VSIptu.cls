VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSIptu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit
Public Type typImovel
    'Dados
    Inscricao As String
    Contribuinte As String
    Endereco As String
    
    
    'Dados Numericos
    ValorVenalTerreno As Double
'    ValorVenalTerrenoFracao() As Double
    ValorVenalEdificacao()  As Double
    ValorVenalEdificacoes As Double
    valorMercadoEdificacao As Double
    valorMercadoTerreno As Double
    AreaConstruidaTotal As Double
    AreaConstruidaTotalCondominio As Double
    AreaConstruidaUnidade() As Double
    AreaLote As Double
    TestadaLote As Double
    Profundidade As Double
    TipoLote As TipoLote
    Apartamento As Boolean
    Destino() As TipoDestino
    AnoConstrucao  As Integer
    VidaUtil   As Integer
    Ocupacao As TipoOcupacaoLote

    'Imposto
    ValorImpostoTerreno As Double
    ValorImpostoEdificacao As Double
    
    'Taxas de Servico Urbano
    ValorTaxaConservacao As Double
    ValorTaxaLixo As Double
    ValorTaxaLimpeza As Double
    ValorTaxas As Double
    
    'Fatores do Imovel
    FatorConservacao As Double
    FatorDepreciacao As Double
    FatorEstrutura As Double
    FatorGleba As Double
    FatorPedologia As Double
    FatorPonderacao As Double
    FatorPadrao As Double
    FatorProfundidade As Double
    FatorSituacao  As Double
    FatorSaneamento  As Double
    FatorTestada As Double
    FatorTopografia  As Double
    FracaoIdeal As Double
    FracaoIdealLoteComum() As Double
    Cobranca As TipoCobranca
    Isento As Boolean
End Type

'ENUMS
Public Enum TipoAreaGleba
    tagDe15000 = 15000
    tagAte16000 = 16000
    tagAte18000 = 18000
    tagAte20000 = 20000
    tagAte40000 = 40000
    tagMais40000 = 40001
End Enum

Public Enum TipoCobranca
    tcbNormal = 1
    tcbIsentoImposto = 2
    tcbIsentoTotal = 3
    tcbImune = 4
End Enum

Public Enum TipoFatorTabelado
    tftSituacao = 6
    tftTopografia = 7
    tftPedologia = 8
    tftEstrutura = 10
    tftPadrao = 12
    tftConservacao = 13
    tftSaneamento = 14
End Enum

Public Enum TipoLote
    tloResidencial = 1
    tloTerreno = 2
    tloGleba = 3
End Enum

Public Enum TipoDestino
    tcoResidencial = 1
    tcoComercial = 2
    tcoIndustria = 3
    tcoPrestacao = 4
    tcoGleba = 6
    tcoOutros = 9
End Enum

Public Enum TipoOcupacaoLote
    tocNaoConstr = 1
    tocRuina = 2
    tocConstrParada = 3
    tocConstrAnda = 4
    tocConstruido = 5
    tocAgropec = 6
    tocServico = 7
End Enum

Public Enum TipoAliquota
    talAliquotaTerreno = 1
    talAliquotaComercial = 2
    talAliquotaEdificacao = 3
    talAliquotaIndustria = 4
    talAliquotaPrestacao = 5
    talAliquotaOutros = 6
End Enum

'Padroes
Private ProfundidadeMinima As Double
Private ProfundidadeMaxima As Double
Private TestadaReferencia As Double
Private ValorBasicoUnitario As Double
Private CustoBasicoUnitario As Double

'Aliquotas
Private AliquotaTerreno As Double
Private AliquotaEdificacao As Double
Private AliquotaComercial As Double

Private AliquotaIndustria As Double
Private AliquotaPrestacao As Double
Private AliquotaOutros As Double

'Variaveis Auxiliares
Private AreaEdificadaTotalCondominio As Double
Private AreaTerrenoTotalCondominio As Double
Private Lote As typImovel
Public AnoLancamento As Double
Private InfraTrecho(1 To 15) As Double

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
Public CodImpostoLixo As String
Public CodImpostoLimpeza As String
Public CodImpostoConservacao As String
Private CodImpostoITU As String
Private ValorMetroTSU As Double
Private ValorMinimoBase As Double
Private ValorMinimoImposto As Double
'Variaveis de Controle de Geracao
Public GeraParcelas As Boolean
Public GeraCotaUnica As Boolean

'Valores Gerados
Public ValorImposto As Double
Public ValorTLP As Double
Public ValorTCVL As Double
Public ValorTCL As Double

Private Sub BuscaLocalizacaoImovel(ByVal Imovel As String, ByRef Bairro As String, ByRef Logradouro As String, Trecho As String, ByRef Zona As Integer)
    'Procedimento que retorna o cod. do bairro, cod. do logradouro e nome
    'do trecho em que o lote se encontra na tabela TAB_TRECHO
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Dim RsEndereco As VSRecordset
    Sql = "Select tim_tci_im,tim_numero,tim_tlg_cod_logradouro,TIM_TBA_COD_BAIRRO,tim_complemento from tab_imovel" & _
        " where tim_ic = '" & Imovel & "' "
    If Bdados.AbreTabela(Sql, Rs) Then
        Lote.Contribuinte = Rs!tim_tci_im
        Bairro = Trim("" & Rs!tim_tba_cod_bairro)
        Logradouro = Trim("" & Rs!tim_tlg_cod_logradouro)
        Sql = "select ttl_nome,tlg_nome,tba_nome from vis_bvt where tlg_cod_logradouro ='" & Rs!tim_tlg_cod_logradouro & "'"
        If Bdados.AbreTabela(Sql, RsEndereco) Then
            Lote.Endereco = "" & RsEndereco!ttl_nome & " " & RsEndereco!tlg_nome & " " & Rs!tim_numero & " " & Rs!tim_complemento & " " & RsEndereco!tba_nome
        End If
        Sql = "SELECT TTC_COD_TRECHO,TTC_ZH FROM TAB_TRECHO WHERE TTC_TLG_COD_LOGRADOURO = '" & Rs!tim_tlg_cod_logradouro & "' AND TTC_SETOR = " & CInt(Mid(Imovel, 3, 2)) & " and TTC_QUADRA = '" & Mid(Imovel, 6, 3) & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            Trecho = Trim("" & Rs!TTC_COD_TRECHO)
            Zona = Nvl("" & Rs!TTC_ZH, 0)
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub BuscaValoresDeMercado(ByVal Imovel As String)
    'Procedimento que retorna o cod. do bairro, cod. do logradouro e nome
    'do trecho em que o lote se encontra na tabela TAB_TRECHO
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsEndereco As VSRecordset
    Sql = "Select TIM_VALOR_TERRENO_MERCADO,TIM_VALOR_EDIFICACAO_MERCADO from tab_imovel" & _
        " where tim_ic = '" & Imovel & "' "
    If Bdados.AbreTabela(Sql, Rs) Then
        Lote.valorMercadoEdificacao = Nvl("" & Rs!TIM_VALOR_EDIFICACAO_MERCADO, 0)
        Lote.valorMercadoTerreno = Nvl("" & Rs!TIM_VALOR_TERRENO_MERCADO, 0)
    End If
    Bdados.FechaTabela Rs
End Sub
Private Function CalculaCUB(ComponenteTipologia As Integer, ComponenteEstrutura As Integer, ComponentePadrao As Integer) As Double
    'Atraves dos valores de tipologia, estrutura e padrao do imovel, encontra o valor
    'do custo unitario basico na tabela TAB_CUB
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "SELECT TCU_VALOR_UNITARIO FROM TAB_CUB WHERE " & _
    "TCU_TCO_COD_COMPONENTE_TIPOLOGIA = " & ComponenteTipologia & _
    " AND TCU_TCO_COD_COMPONENTE_ESTRUTURA = " & ComponenteEstrutura & _
    " AND TCU_TCO_COD_COMPONENTE_PADRAO = " & ComponentePadrao
    CalculaCUB = 0
    If Bdados.AbreTabela(Sql, Rs) Then
        CalculaCUB = Nvl("" & Rs(0), 0)
'    Else
'        Bdados.GravaDados "tab_log", Bdados.PreparaValor("CUB", ComponenteTipologia, _
'         ComponenteEstrutura, ComponentePadrao, Bdados.Converte(Lote.Inscricao, TCTexto)), "TLO_TABELA,TLO_ComponenteTipologia," _
'         & " TLO_ComponenteEstrutura,TLO_ComponentePadrao,TLO_IC", "TLO_TABELA ='CUB' AND " & _
'         " TLO_ComponenteTipologia = " & ComponenteTipologia & " AND TLO_ComponenteEstrutura =" & _
'         ComponenteEstrutura & " AND TLO_ComponentePadrao =" & ComponentePadrao & " AND TLO_IC = '" & Lote.Inscricao & "'"
    End If
    Bdados.FechaTabela Rs
End Function

Private Sub CalculaFatorProfundidade() 'TABELA I
    'Os valores de profundidade do imovel estão estimados em faixas
    'na tabela TAB_PARAMETRO_IPTU
    If Lote.Profundidade >= ProfundidadeMinima And Lote.Profundidade <= ProfundidadeMaxima Then
        Lote.FatorProfundidade = BuscaParamentroIptu("FATOR PROFUNDIDADE", 1)
    ElseIf Lote.Profundidade > (ProfundidadeMaxima * 2) Or Lote.Profundidade < (ProfundidadeMinima / 2) Then
        Lote.FatorProfundidade = BuscaParamentroIptu("FATOR PROFUNDIDADE", 2)
    ElseIf Lote.Profundidade > (ProfundidadeMinima / 2) And Lote.Profundidade < ProfundidadeMinima Then
        Lote.FatorProfundidade = (Lote.Profundidade / ProfundidadeMinima) ^ 0.5
    ElseIf Lote.Profundidade > (ProfundidadeMaxima) And Lote.Profundidade < (2 * ProfundidadeMaxima) Then
        Lote.FatorProfundidade = (ProfundidadeMaxima / Lote.Profundidade) ^ 0.5
    End If
End Sub

Private Sub CalculaFatorTestada() 'TABELA II
    'Os valores de testada do imovel estão estimados em faixas
    'na tabela TAB_PARAMETRO_IPTU
    If Lote.TestadaLote = TestadaReferencia Then
        Lote.FatorTestada = BuscaParamentroIptu("FATOR TESTADA", 1)
    ElseIf Lote.TestadaLote <= (TestadaReferencia / 2) Then
        Lote.FatorTestada = BuscaParamentroIptu("FATOR TESTADA", 2)
    ElseIf Lote.TestadaLote >= (2 * TestadaReferencia) Then
        Lote.FatorTestada = BuscaParamentroIptu("FATOR TESTADA", 3)
    ElseIf Lote.TestadaLote > (TestadaReferencia / 2) And Lote.TestadaLote < (2 * TestadaReferencia) Then
        Lote.FatorTestada = (Lote.TestadaLote / TestadaReferencia) ^ 0.25
    End If
End Sub

Private Sub CalculaFatorPonderacao(Logradouro As String, Zona As Integer, Trecho As String)   'TABELA III
    'Os valores de ponderacao do imovel são calculados usando o somatorio dos indices
    'positivos da zona deste lote na tabela TAB_SITUACAO_PARADIGMA. Depois, fazendo a
    'a valorização do lote, comparando encontrando os componentes que tem neste e que não
    'tem na situação paradigma da zona(bairro) - situacao predominante na zona para
    'cada componente
    
    Dim SomaIndiceZona As Double 'SP
    Dim SomaIgualdadeZonaLote As Double 'RP
    Dim SomaSuperioridadeLote As Double 'AP
    Dim Sql As String
    Dim RsSp As VSRecordset
    Dim Rs As VSRecordset
    Dim I As Byte
    Dim IndicesMais As String 'Componentes positivos na Zona Homogenea/Bairro
    Dim IndicesMenos As String 'Componentes negativos na Zona Homogenea/Bairro
    'CALCULA SP
    '- Somatorio dos indices positivos na zona (SP)
    Sql = "Select  SUM(TSP_INDICE) as SP " & _
    " from TAB_SITUACAO_PARADIGMA where TSP_ZH = " & Zona
    If Bdados.AbreTabela(Sql, RsSp) Then
          SomaIndiceZona = Nvl("" & RsSp!SP, 0)
    Else
        Exit Sub
    End If
    IndicesMais = ""
    IndicesMenos = ""
    SomaSuperioridadeLote = 0
    SomaIgualdadeZonaLote = 0
    'SEPARA INDICES POSITIVOS E NEGATIVOS (O QUE TEM E O QUE NAO TEM NA ZONA HOMOGENEA/BAIRRO)
    Sql = "Select  TSP_INDICE , TSP_TGC_COD_GRUPO" & _
        " from TAB_SITUACAO_PARADIGMA where TSP_ZH = " & Zona
    If Bdados.AbreTabela(Sql, RsSp) Then
        RsSp.MoveFirst
        Do
            If RsSp!TSP_INDICE = 0 Then
                IndicesMenos = IndicesMenos & "," & RsSp!TSP_TGC_COD_GRUPO
            Else
                IndicesMais = IndicesMais & "," & RsSp!TSP_TGC_COD_GRUPO
            End If
            RsSp.MoveNext
        Loop While Not RsSp.EOF
    End If
    IndicesMais = Mid(IndicesMais, 2)
    IndicesMenos = Mid(IndicesMenos, 2)
    
    'BUSCA INDICES DO TRECHO DO LOTE DENTRO DA ZONA HOMOGENEA/BAIRRO
    Sql = "Select tdl_tgl_cod_grupo from TAB_DETALHE_LOGRADOURO,TAB_COMPONENTE_LOGRADOURO " & _
        " where tdl_tlg_cod_logradouro = '" & Logradouro & "' AND tdl_num_trecho = '" & _
        Trecho & "' AND " & "  tdl_tgl_cod_grupo = tcl_grupo AND tdl_tcl_cod_componente = tcl_cod_componente AND " & _
        " TCL_VALOR > 0 AND tcl_grupo IN ("
    If Trim(IndicesMais) <> "" Then
        If Bdados.AbreTabela(Sql & IndicesMais & ")", Rs) Then
            Rs.MoveFirst
            Do
                RsSp.MoveFirst
                Do
                    If Nvl("" & Rs!tdl_tgl_cod_grupo, 0) = Nvl("" & RsSp!TSP_TGC_COD_GRUPO, 0) Then
                        If Nvl("" & RsSp!TSP_INDICE, 0) <> 0 Then
                            'CALCULA RP
                            '- Somatorio da Infra que tem no lote e na zona
                            SomaIgualdadeZonaLote = SomaIgualdadeZonaLote + CDbl(Nvl("" & RsSp!TSP_INDICE, 0))
                        Else
                            'CALCULA AP
                            '- Somatorio da Infra que tem no lote e nao tem na zona
                            SomaSuperioridadeLote = SomaSuperioridadeLote + InfraTrecho("" & RsSp!TSP_TGC_COD_GRUPO)
                        End If
                        Exit Do
                    End If
                    RsSp.MoveNext
                Loop While Not RsSp.EOF
                Rs.MoveNext
            Loop While Not Rs.EOF
        End If
    End If
    Lote.FatorPonderacao = ((1 + SomaIgualdadeZonaLote) / (1 + SomaIndiceZona)) + SomaSuperioridadeLote
End Sub

Private Sub CalculaFatorDepreciacao()   'TABELA IV
    'Calcula o percentual de depreciacao do imovel usando a idade e o tipo do material
    Dim Idade As Double
    Idade = AnoLancamento - Lote.AnoConstrucao
    If Idade > 1000 Then
        Lote.FatorDepreciacao = 0
    Else
        Lote.VidaUtil = IIf(Lote.VidaUtil < Idade, Idade, Lote.VidaUtil)
        Lote.FatorDepreciacao = 0.3 + 0.7 * ((Lote.VidaUtil - Idade) / Lote.VidaUtil)
    End If
End Sub

Private Sub CalculaFatorGleba() 'TABELA IX
    'Os valores de indice da gleba estão estimados em faixas
    'na tabela TAB_PARAMETRO_IPTU
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "SELECT TGL_VALOR FROM TAB_GLEBA_LIMITE WHERE " & _
      Bdados.Converte(Lote.AreaLote, TCDuplo) & ">= TGL_LIMITE_INFERIOR AND " & Bdados.Converte(Lote.AreaLote, TCDuplo) & " <= TGL_LIMITE_SUPERIOR "
    If Bdados.AbreTabela(Sql, Rs) Then
        Lote.FatorGleba = Rs!TGL_VALOR
    End If
    Bdados.FechaTabela Rs
End Sub

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
        End Select
    End If
    Bdados.FechaTabela Rs
End Sub


Private Sub CalculaFracaoIdeal(Unidade As Integer)
    Dim I As Integer
    Dim Fracao As Double
    Dim AreaCondominio As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    'Usa a area total do condominio
    
    If CInt(Left(Lote.Inscricao, 3)) >= 200 Then
        Sql = "select tdi_valor_item from tab_detalhe_imovel where tdi_tgc_cod_grupo = 108 and " & _
            " tdi_tim_ic = '" & Left(Lote.Inscricao, 12) & "200'"
        If Bdados.AbreTabela(Sql, Rs) Then
            Lote.AreaLote = Format(Nvl("" & Rs(0), 0), Const_Monetario)
        Else
            Lote.FracaoIdeal = 0
            Exit Sub
        End If
    End If
    Sql = "select sum(tdi_valor_item) from tab_detalhe_imovel where tdi_tgc_cod_grupo = 112 and " & _
        " tdi_tim_ic like '" & Left(Lote.Inscricao, 12) & "%'"
    If Bdados.AbreTabela(Sql, Rs) Then
        AreaCondominio = Format(Nvl("" & Rs(0), 1), Const_Monetario)
    End If
    'ACHEI A FRACAO IDEAL = Valor Venal de todo o lote dividido pelo total da area construida
    ReDim Preserve Lote.AreaConstruidaUnidade(1 To Unidade) As Double
    Lote.FracaoIdeal = Lote.AreaConstruidaUnidade(Unidade) / CDbl(IIf(AreaCondominio = 0, 1, AreaCondominio))
    ReDim Preserve Lote.FracaoIdealLoteComum(1 To Unidade) As Double
    Lote.FracaoIdealLoteComum(Unidade) = Lote.FracaoIdeal
End Sub

Private Sub CalculaValorVenalTerreno()
    'ANTES - Lote.ValorVenalTerreno = Lote.AreaLote * Lote.FatorPedologia * Lote.FatorTestada * _
                'ValorBasicoUnitario * Lote.FatorProfundidade
                
     Lote.ValorVenalTerreno = Lote.AreaLote * Lote.FatorTestada * _
                Lote.FatorProfundidade
    If Lote.TipoLote = tloGleba And Lote.FatorGleba > 0 Then        'mudado em balsas - ederson
            ' se for gleba ele não calcula pelo fator situação e sim pelo fator gleba
        Lote.ValorVenalTerreno = Format(Lote.ValorVenalTerreno * Lote.FatorTopografia * _
                Lote.FatorGleba * Lote.FatorPonderacao * _
                IIf(CInt(Right(Lote.Inscricao, 3)) >= 200, Lote.FracaoIdeal, 1), Const_Monetario)
    Else
        Lote.ValorVenalTerreno = Format(Lote.ValorVenalTerreno * Lote.FatorTopografia * _
                IIf(Lote.FatorSituacao = 0, 1, Lote.FatorSituacao) * Lote.FatorPonderacao, Const_Monetario) '* IIf(CInt(Right(Lote.Inscricao, 3)) >= 200, Lote.FracaoIdeal, 1)
    End If
End Sub

Private Sub CalculaValorVenalEdificacao(Unidade As Integer)
    On Error GoTo Trata
    ReDim Preserve Lote.ValorVenalEdificacao(1 To Unidade) As Double
    Lote.ValorVenalEdificacao(Unidade) = Lote.AreaConstruidaUnidade(Unidade) * _
        Lote.FatorDepreciacao * Lote.FatorConservacao * CustoBasicoUnitario '* _
        IIf(CInt(Right(Lote.Inscricao, 3)) >= 200, Lote.FracaoIdealEdif, 1)
    Exit Sub
Trata:
    Lote.ValorVenalEdificacao(Unidade) = 0
End Sub

Private Function CalculaVBU(Logradouro As String, Trecho As String) As Double
    'Atraves dos nome do logradouro e do trecho do imovel, encontra o valor
    'basico unitario na tabela TAB_TRECHO
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select TTC_VALOR from TAB_TRECHO where TTC_TLG_COD_LOGRADOURO ='" & Logradouro & _
        "' and TTC_COD_TRECHO ='" & Trecho & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        CalculaVBU = Nvl("" & Rs!TTC_VALOR, 0)
    End If
End Function

Private Sub GravaDadosIptu(Periodo As Double, Vencimento As String, Parcela As Integer)
    Dim Campos As String
    Dim Valores As String
    Dim Conta As New ContaCorrente
    Dim ValorIptu As Double
    Dim CodTaxa As String
    Dim Novo As Boolean
    Dim CodPagamentoTaxa  As String
    Dim CodPagamento As String
    Dim CamposTaxa  As String
    Dim ValoresTaxa As String
    Dim Unidade As Integer
    Dim ValorTotalConstrucao As Double
    Dim CodigoImpotoPrincipal As String
    Dim CodigoImpostoAntigo As String
    Dim Tabela As String
    Dim CodBarra As New CodigoDeBarra
    On Error GoTo Trata
    If CInt(Right(Lote.Inscricao, 3)) = 200 Then Exit Sub  'SE FOR BT DO CONDOMINIO (200), NAO FAZ NADA
    Tabela = IIf(Parcela = 0, "TAB_GERACAO_TRIBUTO", "TAB_GERACAO_TRIBUTO_PARCELA")
    If (Lote.Cobranca < tcbIsentoTotal) Then
        'CodigoImpotoPrincipal = IIf(Lote.ValorVenalEdificacoes = 0, CodImpostoITU, CodImposto)
        CodigoImpotoPrincipal = CodImposto
        ValorIptu = Lote.ValorImpostoEdificacao + Lote.ValorImpostoTerreno
        'Grava na Tab_Geracao_tributo
        CodPagamento = Conta.NumPagamento(Lote.Contribuinte, CLng(Periodo), CodigoImpotoPrincipal, Trim(Lote.Inscricao), Parcela, Novo, 0)
        'Conta.GeraPagamento Lote.Contribuinte, Lote.Inscricao, CodigoImpotoPrincipal, CLng(Periodo), Vencimento, _
        ValorIptu * (100 - DescontoParcela(Parcela)) / 100, 0, 0, CodPagamento, 0, Parcela, Lote.ValorTaxas * (100 - DescontoParcela(Parcela)) / 100, , , IIf(Parcela = 0, 0, 1)
        
        Conta.GeraPagamento Lote.Contribuinte, Lote.Inscricao, CodigoImpotoPrincipal, CLng(Periodo), Vencimento, _
        ValorIptu, 0, 0, CodPagamento, 0, Parcela, Lote.ValorTaxas, , , IIf(Parcela = 0, 0, 1)
        
        Dim LinhaDigitavel As String
        Dim Cobranca As New VSCobranca
        LinhaDigitavel = CodBarra.CriaLinhaDigitavel(Lote.Inscricao, Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)), (ValorIptu + Lote.ValorTaxas) * ((100 - DescontoParcela(Parcela)) / 100), CStr(Periodo), Vencimento, Nvl(CStr(Parcela), 0))
        LinhaDigitavel = CodBarra.LinhaBarraGerada
        
        Dim Valos As String
        Dim Camp As String
        Dim con As String
        Valos = Bdados.PreparaValor("'" & LinhaDigitavel & "'")
        Camp = "tgt_linha_barcode"
        con = "tgt_cod_pagamento = " & CodPagamento
        Bdados.AtualizaDados "TAB_GERACAO_TRIBUTO", Valos, Camp, con
        
        If Parcela = 0 Then
'            CodigoImpostoAntigo = BuscaPagamentoAntigoITU_IPTU(Lote.Contribuinte, CLng(Periodo), IIf(Lote.ValorVenalEdificacoes = 0, CodImposto, CodImpostoITU), Trim(Lote.Inscricao))
'            If CodigoImpostoAntigo > 0 Then
'                Bdados.DeletaDados "tab_geracao_tributo", "tgt_cod_pagamento = " & CodigoImpostoAntigo
'                Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "TCC_CODIGO_CONTA = " & CodigoImpostoAntigo
'            End If
        Else
            Bdados.DeletaDados Tabela, "tgt_cod_pagamento <> " & CodPagamento & " and tgt_cod_pagamento_vinculado <> " & CodPagamento & " and tgt_tim_ic = '" & Lote.Inscricao & "' and tgt_parcela = " & Parcela & " and tgt_periodo = " & Periodo
        End If
    End If
    DoEvents
    'Exit Sub
    'Grava na Tabela de Taxas
    CamposTaxa = "TTD_TGT_COD_PAGAMENTO,TTD_TGT_COD_PAGAMENTO_TAXA"
    'Area generica
    Campos = "tgt_im,tgt_tim_ic,tgt_periodo,tgt_data_vencimento,tgt_parcela,tgt_data_geracao,tgt_cod_pagamento_vinculado,tgt_ativo"
    Valores = Bdados.PreparaValor(Lote.Contribuinte, Bdados.Converte(Lote.Inscricao, tctexto), Periodo, _
        Bdados.Converte(Vencimento, TCDataHora), Parcela, Bdados.Converte(Date, TCDataHora), CodPagamento, IIf(Parcela = 0, 0, 1))
    
    'TAXA DE COLETA DE LIXO
    If Lote.ValorTaxaLixo > 0 Then
        CodPagamentoTaxa = Conta.NumPagamento(Lote.Contribuinte, CLng(Periodo), CodImpostoLixo, Trim(Lote.Inscricao), Parcela, Novo, 0)
        ValoresTaxa = Bdados.PreparaValor(CodPagamento, CodPagamentoTaxa)
        Bdados.InsereDados "TAB_TAXA_DAM", ValoresTaxa, CamposTaxa
        Conta.GeraPagamento Lote.Contribuinte, Lote.Inscricao, CodImpostoLixo, CLng(AnoLancamento), Vencimento, Lote.ValorTaxaLixo, 0, 0, CodPagamentoTaxa, 0, Parcela, 0, , , IIf(Parcela = 0, 0, 1), , CodPagamento
    Else
        Bdados.DeletaDados Tabela, "tgt_tip_cod_imposto='" & CodImpostoLixo & "' and tgt_periodo=" & CLng(Periodo) & " and tgt_im='" & Lote.Contribuinte & "'  and tgt_tim_ic='" & Trim(Lote.Inscricao) & "' and tgt_parcela=" & Parcela
    End If
    
    'TAXA DE CONSERVACAO
    If Lote.ValorTaxaConservacao > 0 Then
        CodPagamentoTaxa = Conta.NumPagamento(Lote.Contribuinte, CLng(Periodo), CodImpostoConservacao, Trim(Lote.Inscricao), Parcela, Novo, 0)
        ValoresTaxa = Bdados.PreparaValor(CodPagamento, CodPagamentoTaxa)
        Bdados.InsereDados "TAB_TAXA_DAM", ValoresTaxa, CamposTaxa
        Conta.GeraPagamento Lote.Contribuinte, Lote.Inscricao, CodImpostoConservacao, CLng(AnoLancamento), Vencimento, Lote.ValorTaxaConservacao, 0, 0, CodPagamentoTaxa, 0, Parcela, 0, , , IIf(Parcela = 0, 0, 1), , CodPagamento
    Else
        Bdados.DeletaDados Tabela, "tgt_tip_cod_imposto='" & CodImpostoConservacao & "' and tgt_periodo=" & CLng(Periodo) & " and tgt_im='" & Lote.Contribuinte & "'  and tgt_tim_ic='" & Trim(Lote.Inscricao) & "' and tgt_parcela=" & Parcela
    End If
    
    'TAXA DE LIMPEZA PUBLICA
    If Lote.ValorTaxaLimpeza > 0 Then
        CodPagamentoTaxa = Conta.NumPagamento(Lote.Contribuinte, CLng(Periodo), CodImpostoLimpeza, Trim(Lote.Inscricao), Parcela, Novo, 0)
        ValoresTaxa = Bdados.PreparaValor(CodPagamento, CodPagamentoTaxa)
        Bdados.InsereDados "TAB_TAXA_DAM", ValoresTaxa, CamposTaxa
        Conta.GeraPagamento Lote.Contribuinte, Lote.Inscricao, CodImpostoLimpeza, CLng(AnoLancamento), Vencimento, Lote.ValorTaxaLimpeza, 0, 0, CodPagamentoTaxa, 0, Parcela, 0, , , IIf(Parcela = 0, 0, 1), , CodPagamento
    Else
        Bdados.DeletaDados Tabela, "tgt_tip_cod_imposto='" & CodImpostoLimpeza & "' and tgt_periodo=" & CLng(Periodo) & " and tgt_im='" & Lote.Contribuinte & "'  and tgt_tim_ic='" & Trim(Lote.Inscricao) & "' and tgt_parcela=" & Parcela
    End If
    DoEvents
    'GRAVA VALORES CALCULADOS PELA PGV
    Campos = "PGV_TIM_IC,PGV_ANO_BASE,PGV_VALOR_TERRENO,PGV_VALOR_EDIFICACAO,PGV_VALOR_IPTU," & _
            "PGV_VALOR_FPROF,PGV_VALOR_FTES,PGV_VALOR_FPOND,PGV_VALOR_FPED,PGV_VALOR_FTOP,PGV_VALOR_FSIT," & _
            "PGV_VALOR_FDEP,PGV_VALOR_FCON,PGV_VALOR_FIDE,PGV_VALOR_FIDT,PGV_VALOR_FGLE,PGV_VBU,PGV_CUB," & _
            "PGV_DATA_PROCESSAMENTO,PGV_ISENTO,PGV_BPS"
    Valores = Bdados.PreparaValor(Bdados.Converte(Lote.Inscricao, tctexto), AnoLancamento, Lote.ValorVenalTerreno, Lote.ValorVenalEdificacoes, _
            Format((Lote.ValorImpostoEdificacao + Lote.ValorImpostoTerreno) * (100 - DescontoParcela(0)) / 100, Const_Monetario), Lote.FatorProfundidade, Lote.FatorTestada, _
            Lote.FatorPonderacao, Lote.FatorPedologia, Lote.FatorTopografia, Lote.FatorSituacao, _
            Lote.FatorDepreciacao, Lote.FatorConservacao, Lote.FracaoIdeal, Lote.FracaoIdeal, Lote.FatorGleba, _
            ValorBasicoUnitario, CustoBasicoUnitario, Date, IIf(Lote.Isento, 1, 0), UBound(Lote.AreaConstruidaUnidade))
    If Parcela = 0 Then Bdados.GravaDados "TAB_PGV_VALORES", Valores, Campos, "PGV_TIM_IC ='" & Lote.Inscricao & "' AND PGV_ANO_BASE = " & AnoLancamento
    Lote.Isento = False
    Exit Sub
Trata:
    Avisa Err.Description
    Exit Sub
End Sub

Public Sub GravaValorVenalImovel()
    Dim Unidade As Integer
    Dim ValorTotalConstrucao As Double
    'GRAVA RESUMO DE VALORES DO LOTE (VALOR EDIFICACAO, DO TERRENO E FRACAO IDEAL
    For Unidade = 1 To UBound(Lote.ValorVenalEdificacao)
        ValorTotalConstrucao = ValorTotalConstrucao + Lote.ValorVenalEdificacao(Unidade)
    Next
    Lote.ValorVenalEdificacoes = ValorTotalConstrucao
    Bdados.AtualizaDados "TAB_IMOVEL", Bdados.PreparaValor(ValorTotalConstrucao, _
        Lote.ValorVenalTerreno, ValorTotalConstrucao + Lote.ValorVenalTerreno), _
        "TIM_VALOR_EDIFIC,TIM_VALOR_TERRENO,TIM_VALOR", "TIM_IC ='" & Lote.Inscricao & "'"
    If CInt(Right(Lote.Inscricao, 3)) >= 200 Then
        Bdados.AtualizaDados "TAB_DETALHE_IMOVEL", Bdados.PreparaValor(Lote.FracaoIdeal), _
        "TDI_VALOR_ITEM", "TDI_TIM_IC ='" & Lote.Inscricao & "' AND TDI_TGC_COD_GRUPO = 114"
    Else
        Dim I As Byte
        For I = 1 To UBound(Lote.FracaoIdealLoteComum)
            Bdados.AtualizaDados "TAB_DETALHE_IMOVEL", Bdados.PreparaValor(Lote.FracaoIdealLoteComum(I)), _
            "TDI_VALOR_ITEM", "TDI_TIM_IC ='" & Lote.Inscricao & "' AND TDI_TGC_COD_GRUPO = 114 and tdi_tim_ic_unidade = " & I
        Next
    End If
End Sub
Private Function BuscaPagamentoAntigoITU_IPTU(Contribuinte As String, Periodo As Long, Imposto As String, Imovel As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select tgt_cod_pagamento  from Tab_Geracao_Tributo " & _
        " where tgt_tip_cod_imposto='" & Imposto & "' and tgt_periodo=" & Periodo & _
        " and tgt_im='" & Contribuinte & "' and tgt_tim_ic='" & Imovel & "' and tgt_parcela=0"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaPagamentoAntigoITU_IPTU = Rs!TGT_COD_PAGAMENTO
    End If
    Bdados.FechaTabela Rs
End Function
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

Public Function CarregaDetalheLote(Imovel As String) As Boolean
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Unidade As Integer
    Dim Aux As Integer
    Dim Tipologia As Integer
    Dim Estrutura As Integer
    Dim Padrao As Integer
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
    'Verifica se há algum pagamento deste periodo. Se existir, o usuario deve estornar este(s) pagamentos primeiro.
    Sql = " SELECT TDR_TGT_COD_PAGAMENTO FROM TAB_DARM_RECEBIDO WHERE TDR_TIM_IC ='" & Lote.Inscricao & _
        "' AND TDR_PERIODO =" & AnoLancamento & " AND TDR_TIP_COD_IMPOSTO IN ('" & CodImposto & "','" & CodImpostoITU & "')"
    If Bdados.AbreTabela(Sql, Rs) Then
'        Util.Avisa "Foram encontrados pagamentos no período. Estorne-os antes!"
        CarregaDetalheLote = False
        Exit Function
    Else
        'SQZ (BLS, 15/04/2003): Isto impede o reaproveitamento do nº DAM, como previsto na NumPagto (!)
'        Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "TGT_TIP_COD_IMPOSTO IN ('" & CodImposto & "','" & CodImpostoConservacao & "','" & CodImpostoITU & "','" & CodImpostoLimpeza & "','" & CodImpostoLixo & "') AND TGT_TIM_IC ='" & Lote.Inscricao & "' AND TGT_PERIODO =" & AnoLancamento   'AND TGT_PERIODO = " & CDbl(AnoLancamento) & " AND TGT_PARCELA not in (" & ParcelasAtivas & ")"
'        Bdados.DeletaDados "TAB_GERACAO_TRIBUTO_PARCELA", "TGT_TIP_COD_IMPOSTO IN ('" & CodImposto & "','" & CodImpostoConservacao & "','" & CodImpostoITU & "','" & CodImpostoLimpeza & "','" & CodImpostoLixo & "') AND TGT_PERIODO = " & CDbl(AnoLancamento) & " AND TGT_TIM_IC ='" & Lote.Inscricao & "' AND TGT_PERIODO =" & AnoLancamento '" AND TGT_PARCELA not in (" & ParcelasAtivas & ")"
'        Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "TCC_TIP_COD_IMPOSTO IN ('" & CodImposto & "','" & CodImpostoConservacao & "','" & CodImpostoITU & "','" & CodImpostoLimpeza & "','" & CodImpostoLixo & "') AND TCC_PERIODO = " & CDbl(AnoLancamento) & " AND TCC_TIM_IC ='" & Lote.Inscricao & "' AND TCC_PERIODO =" & AnoLancamento '" AND TCC_PARCELA not in (" & ParcelasAtivas & ")"
    End If
    'APAGA AS ISENCOES DO LOTE
    Bdados.DeletaDados "TAB_ISENTO", "TIS_TIP_COD_IMPOSTO IN ('" & CodImposto & "','" & CodImpostoConservacao & "','" & CodImpostoITU & "','" & CodImpostoLimpeza & "','" & CodImpostoLixo & "') AND TIS_PERIODO = " & CDbl(AnoLancamento) & " AND TIS_TIM_IC ='" & Lote.Inscricao & "'"
    If CInt(Right(Imovel, 3)) = 200 Or CInt(Right(Imovel, 3)) = 600 Then Exit Function 'SE FOR BT DO CONDOMINIO (200), NAO FAZ NADA
    
    Sql = "Select tdi_tgc_cod_grupo,tdi_valor_item,tdi_tco_cod_componente,tdi_tim_ic_unidade from tab_detalhe_imovel where "
    Sql = Sql & " tdi_tim_ic = '" & Imovel & "' and  ((tdi_tgc_cod_grupo >= 1 and tdi_tgc_cod_grupo <= 13) " & _
    "  or  (tdi_tgc_cod_grupo >= 100 and tdi_tgc_cod_grupo <= 113)) "
    Sql = Sql & "  order by  tdi_tim_ic_unidade,tdi_tgc_cod_grupo "
    If Bdados.AbreTabela(Sql, Rs) Then
        Lote.Inscricao = Imovel
        
        
        Lote.TipoLote = CInt(BuscaCodigo("Select tim_tipo_imovel from tab_imovel where tim_ic ='" & Imovel & "'"))
        
        If Right(Imovel, 3) > 200 Then
            Lote.TipoLote = tloResidencial
        End If
        
        Rs.MoveFirst
        Aux = Rs!tdi_tim_ic_unidade
        Unidade = 0
        BuscaLocalizacaoImovel Imovel, Bairro, Logr, Trecho, Zona
        BuscaValoresPadroes Zona
        CalculaFatorPonderacao Logr, Zona, Trecho
        BuscaValoresDeMercado Lote.Inscricao
        Do While Not Rs.EOF
            If Aux <> Rs!tdi_tim_ic_unidade And (Rs!tdi_tim_ic_unidade <> 0 And Rs!tdi_tim_ic_unidade <> 200) Then
                If (Aux <> 0 And Aux <> 200) Then
                    'Calcula Fatores
                    CalculaFatorDepreciacao
                    CustoBasicoUnitario = CalculaCUB(Tipologia, Estrutura, Padrao) 'Custo Unitario Basico
                    'Calcula Valores Venais
                    CalculaValorVenalEdificacao Unidade
                    CalculaFracaoIdeal Unidade
                End If
                Unidade = Unidade + 1
                Aux = Rs!tdi_tim_ic_unidade
            End If
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
                Case 10 'Estrutura
                    Estrutura = Rs!tdi_tco_cod_componente
                    CalculaFatorTabelado tftEstrutura, Rs!tdi_tco_cod_componente, Unidade
                Case 11 'Tipo de Construcao(Destino)
                    BuscaDestinoLote Rs!tdi_tco_cod_componente, IIf(Unidade = 0, 1, Unidade)
                Case 12 'Padrao
                    Padrao = Rs!tdi_tco_cod_componente
                Case 13 'Conservacao
                    CalculaFatorTabelado tftConservacao, Rs!tdi_tco_cod_componente, Unidade
                Case 100 'Testada
                    Lote.TestadaLote = Rs!tdi_valor_item
                    CalculaFatorTestada
                Case 108 'Area Lote
                    Lote.AreaLote = Rs!tdi_valor_item
                Case 111 'Ano da Construcao
                    Lote.AnoConstrucao = Rs!tdi_valor_item
                Case 112 'Area Construida Unidade
                    Unidade = IIf(Unidade = 0, 1, Unidade)
                    ReDim Preserve Lote.AreaConstruidaUnidade(1 To Unidade) As Double
                    Lote.AreaConstruidaUnidade(Unidade) = Format(Rs!tdi_valor_item, Const_Monetario)
                Case 113 'Area Construida Total
                    Lote.AreaConstruidaTotal = Format(Rs!tdi_valor_item, Const_Monetario)
                Case 114 'Area Construida Total
                    Lote.Profundidade = Format(Rs!tdi_valor_item, Const_Monetario)
            End Select
            Rs.MoveNext
        Loop
    End If
    'Se for lote de condominio, busca a testada do condominio e calcula a fracao ideal dele
    If CInt(Right(Lote.Inscricao, 3)) > 200 Then
        Sql = "Select tdi_tgc_cod_grupo,tdi_valor_item,tdi_tco_cod_componente,tdi_tim_ic_unidade from tab_detalhe_imovel where tdi_tim_ic = '" & Left(Imovel, 12) & "200' and (tdi_tgc_cod_grupo >= 1 and tdi_tgc_cod_grupo <= 8 or tdi_tgc_cod_grupo = 100 or tdi_tgc_cod_grupo = 108)"
        If Bdados.AbreTabela(Sql, Rs) Then
            Rs.MoveFirst
            Do
                Select Case Rs!tdi_tgc_cod_grupo
                    Case 1 'Ocupacao
                        Lote.Ocupacao = Rs!tdi_tco_cod_componente
                    Case 6 'Situacao
                        CalculaFatorTabelado tftSituacao, Rs!tdi_tco_cod_componente, Unidade
                    Case 7 'Tografia
                        CalculaFatorTabelado tftTopografia, Rs!tdi_tco_cod_componente, Unidade
                    Case 8 'Pedologia
                        CalculaFatorTabelado tftPedologia, Rs!tdi_tco_cod_componente, Unidade
                    Case 100
                        Lote.TestadaLote = Rs!tdi_valor_item
                        CalculaFatorTestada
                    Case 108 'Area Lote
                        Lote.AreaLote = Rs!tdi_valor_item
                End Select
                Rs.MoveNext
            Loop While Not Rs.EOF
            Bdados.FechaTabela Rs
        End If
        'Se for CONDOMINIO, busca novamente Cod. Cobranca
        Sql = "Select tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic = '" & Imovel & "' and tdi_tgc_cod_grupo = 3"
        If Bdados.AbreTabela(Sql, Rs) Then
            Lote.Cobranca = Rs!tdi_valor_item
        End If
    End If
    'Calcula Fatores para a ultima unidade construida
    If (Unidade <> 0 And Unidade <> 200) Then
        CalculaFatorDepreciacao
        If Lote.TipoLote = tloGleba Then CalculaFatorGleba
        CustoBasicoUnitario = CalculaCUB(Tipologia, Estrutura, Padrao)          'Custo Unitario Basico
        'Calcula Valores Venais
        CalculaValorVenalEdificacao Unidade
        CalculaFracaoIdeal Unidade
    End If
    'Fim Fatores
    If Lote.TestadaLote > 0 And Lote.Profundidade = 0 Then
        Lote.Profundidade = Format(Lote.AreaLote / Lote.TestadaLote, Const_Monetario)
    End If
    ValorBasicoUnitario = CalculaVBU(Logr, Trecho) 'Valor Basico Unitario
    CalculaFatorProfundidade
    
    If Lote.TipoLote = tloGleba Then CalculaFatorGleba
    
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

Private Sub CarregaValoresInfra()
    Dim Sql As String
    Dim Rs As VSRecordset
    'Carrega os valores dos indices de infra - estrutura possiveis no trecho
    Sql = "Select distinct tcl_grupo,tcl_valor from tab_componente_logradouro where tcl_valor > 0 "
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do
            InfraTrecho(Rs!tcl_grupo) = Rs!tcl_valor
            Rs.MoveNext
        Loop While Not Rs.EOF
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub BuscaValoresPadroes(Zona As Integer)
    'Busca valores de profundidade minima e maxima, testada referencia
    'na tabela TAB_LOTE_PADRAO
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select TLO_TESTADA,TLO_PROF_MINIMA,TLO_PROF_MAXIMA,TLO_VALOR_MAX,TLO_VALOR_MIN from TAB_LOTE_PADRAO"
    Sql = Sql & " where TLO_ZH =" & Zona
    If Bdados.AbreTabela(Sql, Rs) Then
        ProfundidadeMinima = Rs!TLO_PROF_MINIMA
        ProfundidadeMaxima = Rs!TLO_PROF_MAXIMA
        TestadaReferencia = Rs!TLO_TESTADA
    End If
    Bdados.FechaTabela Rs
End Sub

Public Function CalculaValorIptu(Optional GeraSomenteValorVenal As Boolean = False) As Double
    Dim Unidade As Integer
    Dim ValorTotalConstrucao As Double
    Dim LoteComercial As Boolean
    Dim AliquotaEspecifica As Double
    Dim I As Byte
    Dim j As Byte
    On Error GoTo Trata
    Lote.ValorImpostoEdificacao = 0
    ValorTotalConstrucao = 0
    Lote.ValorImpostoTerreno = 0
    If Right(Lote.Inscricao, 3) = 200 Then Exit Function
    
    'AREA DE VERIFICACAO DA EXISTENCIA DE VALORES DE LOTE COM BASE NO MERCADO IMOBILIARIO
    'ANDRE - 06/09/2002
    '-------------------------------------------------------------------------------------
    If Lote.valorMercadoTerreno > 0 Then Lote.ValorVenalTerreno = Lote.valorMercadoTerreno
    If Lote.valorMercadoEdificacao > 0 Then
        For Unidade = 1 To UBound(Lote.ValorVenalEdificacao)
            Lote.ValorVenalEdificacao(Unidade) = Lote.valorMercadoEdificacao * Lote.FracaoIdealLoteComum(Unidade)
        Next
    End If
    '-------------------------------------------------------------------------------------
    'Imovel sem unidade(Terreno)
    If Lote.TipoLote = tloTerreno Or Lote.Ocupacao < 5 Then
        Lote.ValorImpostoTerreno = Lote.ValorVenalTerreno * (AliquotaTerreno / 100)
    Else
        'Calculo do imposto por unidade construida
        For Unidade = 1 To UBound(Lote.ValorVenalEdificacao)
            On Error Resume Next
            AliquotaEspecifica = 0
            Select Case Lote.Destino(Unidade)
                Case tcoResidencial
                    If Lote.Cobranca <> tcbIsentoImposto Then AliquotaEspecifica = AliquotaEdificacao
                Case tcoComercial
                    If Lote.Cobranca <> tcbIsentoImposto Then AliquotaEspecifica = AliquotaComercial
                Case tcoIndustria
                    If Lote.Cobranca <> tcbIsentoImposto Then AliquotaEspecifica = AliquotaIndustria
                Case tcoPrestacao
                    If Lote.Cobranca <> tcbIsentoImposto Then AliquotaEspecifica = AliquotaPrestacao
                Case tcoOutros
                    If Lote.Cobranca <> tcbIsentoImposto Then AliquotaEspecifica = AliquotaOutros
            End Select
            'Calcula o valor do imposto da unidade de acordo com a aliquota especifica
            Lote.ValorImpostoTerreno = Lote.ValorImpostoTerreno + Lote.ValorVenalTerreno * (AliquotaEspecifica / 100) * Lote.FracaoIdealLoteComum(Unidade)
            Lote.ValorImpostoEdificacao = Lote.ValorImpostoEdificacao + (Lote.ValorVenalEdificacao(Unidade) * (AliquotaEspecifica / 100))
        Next
    End If
    If Lote.Cobranca = tcbNormal Or Lote.Cobranca = tcbIsentoImposto Then 'Pode calcular taxas
        CalculaTaxasServico
    End If
    'AREA DE GRAVACAO DA PARCELA UNICA
    GravaValorVenalImovel
    If GeraSomenteValorVenal Then
        Lote.ValorImpostoEdificacao = 0
        Lote.ValorImpostoTerreno = 0
        CalculaValorIptu = 0
        
        Exit Function
    End If
    DoEvents
    
    'GravaDadosIptu CDbl(AnoLancamento), VencimentoCotaUnica, 0
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
                GravaIsento CodImposto, 1, CDbl(Lote.ValorImpostoTerreno) + CDbl(Lote.ValorImpostoEdificacao), CDbl(AnoLancamento)
                Lote.ValorImpostoEdificacao = 0
                Lote.ValorImpostoTerreno = 0
                Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "TGT_TIP_COD_IMPOSTO IN ('" & CodImposto & "','" & CodImpostoITU & "','" & CodImpostoConservacao & "','" & CodImpostoLimpeza & "','" & CodImpostoLixo & "') AND TGT_PERIODO = " & CDbl(AnoLancamento) & " AND TGT_TIM_IC = '" & Lote.Inscricao & "'"
                Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "TCC_TIP_COD_IMPOSTO IN ('" & CodImposto & "','" & CodImpostoITU & "','" & CodImpostoConservacao & "','" & CodImpostoLimpeza & "','" & CodImpostoLixo & "') AND TCC_TIM_IC = '" & Lote.Inscricao & "'"
            End If
            If Trim(VencimentoParcela(0)) <> "" Then GravaDadosIptu CDbl(AnoLancamento), VencimentoParcela(0), 0
        Else
            'É isento total ou imune
            GravaIsento CodImposto, Lote.Cobranca, CDbl(Lote.ValorImpostoTerreno) + CDbl(Lote.ValorImpostoEdificacao), CDbl(AnoLancamento)
            If Lote.ValorTaxaLixo <> 0 Then GravaIsento CodImpostoLixo, Lote.Cobranca, Lote.ValorTaxaLixo, CDbl(AnoLancamento)
            If Lote.ValorTaxaConservacao <> 0 Then GravaIsento CodImpostoConservacao, Lote.Cobranca, Lote.ValorTaxaConservacao, CDbl(AnoLancamento)
            If Lote.ValorTaxaLimpeza <> 0 Then GravaIsento CodImpostoLimpeza, Lote.Cobranca, Lote.ValorTaxaLimpeza, CDbl(AnoLancamento)
            Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "TGT_TIP_COD_IMPOSTO IN ('" & CodImposto & "','" & CodImpostoITU & "','" & CodImpostoConservacao & "','" & CodImpostoLimpeza & "','" & CodImpostoLixo & "') AND TGT_PERIODO = " & CDbl(AnoLancamento) & " AND TGT_TIM_IC = '" & Lote.Inscricao & "'"
            Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "TCC_TIP_COD_IMPOSTO IN ('" & CodImposto & "','" & CodImpostoITU & "','" & CodImpostoConservacao & "','" & CodImpostoLimpeza & "','" & CodImpostoLixo & "') AND TCC_PERIODO = " & CDbl(AnoLancamento) & " AND TCC_TIM_IC = '" & Lote.Inscricao & "'"
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

Public Sub GravaIsento(ByVal CodigoTributo As String, ByVal TipoIsento As Integer, ByVal ValorTributo As Double, PeriodoIsencao As Double)
    'Isento Limite Base = 1
    'Isento Imposto = 2
    'Isento Limete Tributo = 3
    'Isento Total = 4
    'Imune = 5

    Dim Valores As String
    Dim Campos As String
    
    Campos = "TIS_TCI_IM,TIS_PERIODO,TIS_VALOR_TRIBUTO,TIS_TIP_COD_IMPOSTO,TIS_TIM_IC,TIS_TIPO_ISENSAO,TIS_DATA,TIS_TUS_COD_USUARIO"
    Valores = Bdados.PreparaValor(Lote.Contribuinte, PeriodoIsencao, Format(ValorTributo, Const_Monetario), CodigoTributo, Bdados.Converte(Lote.Inscricao, tctexto), TipoIsento, Bdados.Converte(Date, TCDataHora), AplicacoesVTFuncoes.Usuario)
    Bdados.GravaDados "Tab_Isento", Valores, Campos, "TIS_TCI_IM ='" & Lote.Contribuinte & "' and TIS_TIP_COD_IMPOSTO ='" & CodigoTributo & "' and TIS_TIM_IC ='" & Lote.Inscricao & "' AND TIS_TIPO_ISENSAO=" & TipoIsento & " AND TIS_PERIODO=" & PeriodoIsencao
    Lote.Isento = True
End Sub

Private Sub CalculaTaxasServico()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim ValorTaxa As Double
    Dim PavimentacaoOK As Boolean
    'ZERA TAXAS
    Lote.ValorTaxaLimpeza = 0
    Lote.ValorTaxaConservacao = 0
    Lote.ValorTaxaLixo = 0
    Lote.ValorTaxas = 0
    PavimentacaoOK = False
    
    '*** SOMENTE PARA CONDOMINIO (FRACAO IDEAL) ***
    Lote.TestadaLote = Format(Lote.TestadaLote * IIf(CInt(Right(Lote.Inscricao, 3)) >= 200, Lote.FracaoIdeal, 1), Const_Monetario)
    Lote.AreaLote = Format(Lote.AreaLote * IIf(CInt(Right(Lote.Inscricao, 3)) >= 200, Lote.FracaoIdeal, 1), Const_Monetario)
    '*** FIM DE CONDOMINIO ***
    
    'ESTOU INDO VERIFICAR QUAIS SERVICOS ESTÃO DISPONIVEIS NESTE TRECHO DO LOGRADOURO
    Sql = "Select tdl_tcl_cod_componente,tdl_tgl_cod_grupo from tab_detalhe_logradouro where tdl_tlg_cod_logradouro='" & _
        Logr & "' and tdl_num_trecho = '" & Trecho & "' AND tdl_tgl_cod_grupo IN (8,11,14,15) ORDER BY tdl_tgl_cod_grupo ASC"
    If Bdados.AbreTabela(Sql, Rs) Then
        'O VALOR DE CADA TAXA DE SERVICO É DE 1 REAL POR M2 DE TESTADA SERVIDA PELO SERVICO
        Rs.MoveFirst
        Do
            Select Case Rs!tdl_tgl_cod_grupo
                Case 8 'PAVIMENTACAO
                    If Nvl("" & Rs!tdl_tcl_cod_componente, 0) < 4 Then
                        PavimentacaoOK = True
                    End If
                    
                Case 11 'COLETA DE LIXO
                    If Nvl("" & Rs!tdl_tcl_cod_componente, 0) <> 7 Then
                        If Lote.TipoLote <> tloTerreno And Lote.TipoLote <> tloGleba Then    'CASO FOR UM TERRENO, NÃO COBRAR COLETA DE LIXO
                            Lote.ValorTaxaLixo = Lote.TestadaLote * ValorMetroTSU
                        End If
                    End If
                
                Case 14 'LIMPEZA PUBLICA
                    If Nvl("" & Rs!tdl_tcl_cod_componente, 0) <> 2 Then
                        Lote.ValorTaxaLimpeza = Lote.TestadaLote * ValorMetroTSU
                    End If
                Case 15 'CONSERVACAO DE VIAS(PAVIMENTACAO)
                    If Nvl("" & Rs!tdl_tcl_cod_componente, 0) <> 3 Then    ' ederson -' 6 representa o valor 'não tem' do banco
                        If PavimentacaoOK Then Lote.ValorTaxaConservacao = Lote.TestadaLote * ValorMetroTSU
                    End If
            End Select
            Rs.MoveNext
        Loop While Not Rs.EOF
        Lote.ValorTaxas = Lote.ValorTaxaLimpeza + Lote.ValorTaxaConservacao + Lote.ValorTaxaLixo
    End If
    Bdados.FechaTabela Rs
    ValorTLP = Lote.ValorTaxaLimpeza
    ValorTCVL = Lote.ValorTaxaConservacao
    ValorTCL = Lote.ValorTaxaLixo
End Sub

Public Function BuscaParamentroIptu(Tabela As String, Registro As TipoAliquota, Optional Ano As String) As String
    'Funcao auxiliar que retorna valores parametrizados na tabela TAB_PARAMETRO_IPTU
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "select tpp_valor " & _
            " from tab_parametro_iptu " & _
            " where tpp_cod_tabela = (" & _
                    " SELECT tpp_cod_tabela " & _
                    " FROM tab_parametro_iptu  " & _
                    " WHERE tpp_valor ='" & Tabela & "'"
    Sql = Sql & ") and tpp_registro =" & Registro
    If Trim$(Ano) <> "" Then Sql = Sql & " AND TPP_ANO_ALIQUOTA=" & Ano
    
    BuscaParamentroIptu = ""
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaParamentroIptu = "" & Rs(0)
    Else
        BuscaParamentroIptu = 0
    End If
    
    Rs.Fechar
End Function

Public Sub ImprimeBoletoIptu(Imovel As String)
    'Imprime os Boletos de IPTU caluculados
    
    Dim Condicao As String
    Dim Report As VSRelatorio
    On Error GoTo Trata
    Condicao = "{Vis_Imovel.tim_ic}='" & Imovel & "' and {Tab_Geracao_Tributo.tgt_periodo} = " & AnoLancamento & " AND TGT_TIP_COD_IMPOSTO = '" & CodImposto & "'"
    Set Report = New VSRelatorio
    With Report
        If Not .DefinirArquivo(Bdados, App.Path & "\TdamBoleto.rpt") Then Exit Sub
        .Selecao = Condicao
        .Imprimir
    End With
    Set Rpt = Nothing
    Screen.MousePointer = 0
    Exit Sub
Trata:
    Avisa Err.Description
    Exit Sub
End Sub

Public Sub InicializarValores(Optional Imovel As String)
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
    CarregaValoresInfra
    'Busca Aliquotas de terreno, edificacao, comercio, industria, prestacao de servicos e outros
    AliquotaTerreno = TrocaPic(BuscaParamentroIptu("ALIQUOTAS", talAliquotaTerreno, Me.AnoLancamento), ".", ",")
    AliquotaComercial = TrocaPic(BuscaParamentroIptu("ALIQUOTAS", talAliquotaComercial, Me.AnoLancamento), ".", ",")
    AliquotaEdificacao = TrocaPic(BuscaParamentroIptu("ALIQUOTAS", talAliquotaEdificacao, Me.AnoLancamento), ".", ",")
    AliquotaIndustria = TrocaPic(BuscaParamentroIptu("ALIQUOTAS", talAliquotaIndustria, Me.AnoLancamento), ".", ",")
    AliquotaPrestacao = TrocaPic(BuscaParamentroIptu("ALIQUOTAS", talAliquotaPrestacao, Me.AnoLancamento), ".", ",")
    AliquotaOutros = TrocaPic(BuscaParamentroIptu("ALIQUOTAS", talAliquotaOutros, Me.AnoLancamento), ".", ",")
    
    VencimentoCotaUnica = TrocaPic(BuscaParamentroIptu("VENCIMENTO COTAS", 1), ".", ",")
    ValorMetroTSU = BuscaParamentroIptu("VALOR M2 TSU", 1)
    ValorMinimoBase = Imposto.BuscaValorMinimoBase(CodImposto, Year(Date))   'IMPOSTO MINIMO
    ValorMinimoImposto = Imposto.BuscaValorMinimoImposto(CodImposto, Year(Date))   'IMPOSTO MINIMO
    CodImpostoLixo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_TXCOLETA))
    CodImpostoConservacao = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_TXCONSERV))
    CodImpostoLimpeza = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_TXLIMPA))
     CodImpostoITU = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU))
    CodImpostoITU = CodImposto
    Parcelas = 0
    ValorImposto = 0
    ValorTCL = 0
    ValorTCVL = 0
    ValorTLP = 0
    If Not GeraParcelas Then GeraCotaUnica = True
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


Public Function Busca_IC_Auxiliar(IC As String) As String
    If Bdados.AbreTabela("SELECT tim_ic_auxiliar FROM TAB_IMOVEL WHERE TIM_IC = '" & IC & "'") Then
        Busca_IC_Auxiliar = Format(Bdados.Tabela(0), "@@.@@.@@@.@@@@.@@@")
    End If
    
End Function


Public Function Busca_Endereco_Imovel(IC As String) As String
    If Bdados.AbreTabela("select ltrim(rtrim(ttl_nome))   +' '+ ltrim(rtrim(tlg_nome)) +', '+ ltrim(rtrim(tim_numero)) + ', '+ rtrim(ltrim(tim_complemento)) + '- '+ ltrim(rtrim(tba_nome))  as Endereco from vis_endereco_imovel WHERE TIM_IC = '" & IC & "'") Then
        Busca_Endereco_Imovel = "" & Bdados.Tabela("Endereco")
    End If
End Function
'
Public Function Busca_Loteamento(IC As String) As String
    If Bdados.AbreTabela("select ltrim(rtrim(tlo_descricao)) , ' Quadra ' + ltrim(rtrim(tim_quadra)) ,' Lote '+ ltrim(rtrim(tim_lote)) from VIS_LOTE_IMOVEL WHERE TIM_IC =  '" & IC & "'") Then
        Busca_Loteamento = "" & Bdados.Tabela(0) & "" & IIf(Trim(Bdados.Tabela(1)) = "Quadra", "", Bdados.Tabela(1)) & "" & IIf(Trim(Bdados.Tabela(2)) = "Lote", "", Bdados.Tabela(2))
    End If
End Function

Public Function Busca_Edificio(IC As String) As String
    If Bdados.AbreTabela("select 'Ed. ' + ltrim(rtrim(ted_descricao)) , ' Bl ' + ltrim(rtrim(tim_bloco)) ,' Apto '+ ltrim(rtrim(TIM_APTO)) ,' Sl/Lj '+ ltrim(rtrim(TIM_SALA_LOJA)) from TAB_IMOVEL,TAB_EDIFICIO WHERE TIM_TED_COD_EDIFICIO=TED_COD_EDIFICIO AND TIM_IC =  '" & IC & "'") Then
        Busca_Edificio = "" & Bdados.Tabela(0) & "" & Bdados.Tabela(1) & "" & Bdados.Tabela(2) & "" & Bdados.Tabela(3)
    End If
End Function


Public Function Busca_Doc(Im As String) As String
    If Bdados.AbreTabela("select tci_cgc_cpf from tab_contribuinte where tci_im =   '" & Trim(Im) & "'") Then
        If Not IsNull(Bdados.Tabela(0)) Then
            Busca_Doc = "" & Bdados.Tabela(0)
        End If
    Else
        'Pego o contribuinte na tabela e imovel
        If Bdados.AbreTabela("select tci_cgc_cpf from tab_contribuinte where tci_im = (select tim_tci_im from tab_imovel where tim_ic = '" & Trim(Im) & "')") Then
            Busca_Doc = "" & Bdados.Tabela(0)
        End If
    End If
End Function

