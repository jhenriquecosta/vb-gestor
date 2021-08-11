Attribute VB_Name = "ModObrig"

Public Sub ImprimeSelecionado(Lista As Object, Razao As String, Endereco As String, _
                Optional ComNovaData As Boolean = False, Optional NovaData As String, _
                Optional DetinoDAM As TipoDestinoImpressao = tdiTela, Optional StringTaxas As String, _
                Optional TotalTaxas As Double, Optional Proprietario As String, Optional EnderecoProprietario As Object)
    Dim Sql As String
    Dim CpfCgc As String
    Dim Declaracao As New cDeclaracao
    Dim Desconto As Double
    Dim Exer As String
    Dim BaseCalculo As String
    Dim Obs As String
    Dim Imposto As New VSImposto
    Dim Lote As New BCI
    Dim Cobranca As New VSCobranca
    Dim NovoJuro As String
    Dim NovaMulta As String
    Dim Conta As New ContaCorrente
    Dim Obrig As New Obrigacao
    Dim NotaInicial As String
    Dim NotaFinal As String
    Dim strValorTaxa As String
    Dim strNomeTaxa As String
    Dim Rs_Carro As VSRecordset
    Dim Rs_Dividas As VSRecordset
    Dim IPTU As New VSIptu
    Dim RsItbi As VSRecordset
    Dim PicBarra As Object
    Dim Taxa As Double
    Dim Correcao As Double
    Dim DescMulta As Double
    Dim DescJuros As Double
    Dim DescCorrecao As Double
    
   
    With Lista
        If (Not ComNovaData) And (NovaData = .SelectedItem.SubItems(5)) Then
            NovaData = .SelectedItem.SubItems(5)
            NovoJuro = .SelectedItem.SubItems(7)
            NovaMulta = .SelectedItem.SubItems(8)
        Else
            Correcao = Conta.CalculaValoresCorrecaoAvulso(.SelectedItem.SubItems(11), .SelectedItem.SubItems(2), .SelectedItem.SubItems(5), NovaData, .SelectedItem.SubItems(6), , DescCorrecao)
            NovoJuro = Conta.CalculaValoresJurosAvulsos(.SelectedItem.SubItems(11), .SelectedItem.SubItems(2), EtcCreditoTributario, NovaData, .SelectedItem.SubItems(5), .SelectedItem.SubItems(6) + Correcao, , DescJuros)
            NovaMulta = Conta.CalculaValoresMultaAvulsos(.SelectedItem.SubItems(11), .SelectedItem.SubItems(2), EtcCreditoTributario, NovaData, .SelectedItem.SubItems(5), .SelectedItem.SubItems(6) + Correcao, , DescMulta)
        End If
        If .SelectedItem.SubItems(3) = Imposto.NomeTributo(ttr_ITBI) Then
            Sql = "SELECT * FROM TAB_TRANSFERENCIA_IMOVEL WHERE TTI_TOC_COD_OBRIGACAO =" & .SelectedItem
            If Bdados.AbreTabela(Sql, RsItbi) Then
                Cobranca.ImprimeDamITBI .SelectedItem, "" & RsItbi!TTI_TCI_IM_ADQUIRENTE, "" & RsItbi!TTI_TCI_CPF_ADQUIRENTE, _
                "" & RsItbi!TTI_CEDENTE_NOME, "" & RsItbi!TTI_CEDENTE_ENDERECO, .SelectedItem.SubItems(5), "" & RsItbi!TTI_TCI_NOME_ADQUIRENTE, _
                    "" & RsItbi!TTI_TCI_ENDERECO_ADQUIRENTE, "" & RsItbi!TTI_TIM_IC, "" & RsItbi!TTI_TCI_ENDERECO_IMOVEL, .SelectedItem.SubItems(2), .SelectedItem.SubItems(6), _
                    CDbl(Nvl(NovaMulta, 0)), CDbl(Nvl(NovoJuro, 0)), Nvl("" & RsItbi!TTI_VALOR_LOTE_INFORMADO, 0), Nvl("" & RsItbi!TTI_ALIQUOTA_PROPRIO, 0), _
                     Nvl("" & RsItbi!TTI_ALIQUOTA_FINANCIADO, 0), CDbl(Taxa), "" & RsItbi!TTI_OBSERVACAO, _
                     "" & RsItbi!TTI_ESPECIE, PicBarra, Nvl("" & RsItbi!tti_Valor_Financiado, 0), _
                     Nvl("" & RsItbi!tti_Valor_avista, 0), _
                     CDbl(Nvl(RsItbi!tti_Valor_Financiado, 0)) + CDbl(Nvl("" & RsItbi!tti_Valor_avista, 0)), _
                     "" & RsItbi!tti_processo, , .SelectedItem.SubItems(11)
                Exit Sub
            End If
        End If
        Obs = ""
        
          
        Exer = IIf(Len(.SelectedItem.SubItems(4)) = 2, .SelectedItem.SubItems(2), Right(.SelectedItem.SubItems(2), 2) & Left(.SelectedItem.SubItems(2), 4))
        strValorTaxa = Nvl(.SelectedItem.SubItems(10), 0)
        
        If .SelectedItem.SubItems(3) = Imposto.NomeTributo(ttr_IPTU) Or .SelectedItem.SubItems(3) = Imposto.NomeTributo(ttr_ITU) Then
            If CDate(.SelectedItem.SubItems(5)) >= Date Then
                If Bdados.AbreTabela("Select TPP_DESCONTO from TAB_PARAMETRO_PARCELA_IPTU WHERE tpp_parcela=" & .SelectedItem.SubItems(4) & " and TPP_ANO=" & Exer) Then
                    Desconto = .SelectedItem.SubItems(6) - (.SelectedItem.SubItems(6) * Bdados.Tabela(0) / 100)
                End If
            End If
        End If
        CpfCgc = IPTU.Busca_Doc(.SelectedItem.SubItems(1))
        If .SelectedItem.SubItems(3) = Imposto.NomeTributo(ttr_ISSQN) Then
            If Declaracao.Buscar(.SelectedItem.SubItems(1), .SelectedItem.SubItems(4)) Then
                BaseCalculo = Declaracao.Itens(4).Valor
                NotaInicial = Declaracao.Itens(1).Valor
                NotaFinal = Declaracao.Itens(2).Valor
            Else
                BaseCalculo = (.SelectedItem.SubItems(6) * 100) / ((Imposto.BuscaAliquota(Imposto.BuscaCodImposto(.SelectedItem.SubItems(3)), Left(.SelectedItem.SubItems(4), 4))) * 100)
                Obs = ""
            End If
        ElseIf .SelectedItem.SubItems(3) = Imposto.NomeTributo(ttr_IPTU) Then
            BaseCalculo = Format(Lote.ValorVenal(.SelectedItem.SubItems(1)), Const_Monetario)
        Else
            BaseCalculo = .SelectedItem.SubItems(6)
        End If
        Set Declaracao = Nothing
'        End If
      If Trim(Obs) = "" Then Obs = Trim(Util.Entrada("Observacao", "Impressão de Parcela."))
        Obrig.BuscaDetalheObrigacao .SelectedItem
        
        
        Dim DescImposto As String
        Dim EnderecoContribuinte As String
        Dim EnderecoImovel As String
        
        'CALCULO SELIC 12-05-2011 - BCP GLEYSON
        'PRECISO PEGAR O METODO DA CONTA CORRENTE (AtualizAtualizacaoMonetariaSelic) para atualizar de acordo com o vencimento (NovaData)
        Dim contaSelic As ContaCorrente
        Set contaSelic = New ContaCorrente
        contaSelic.AtualizAtualizacaoMonetariaSelic Obrig.obCodigoObrigacao, Obrig.obValorObrigacao, CDate(Obrig.obDataVencimento), CStr(NovoJuro), CStr(NovaMulta), CStr(Obrig.obValorDesconto)
        Correcao = contaSelic.PegaValorAtualizaMonetariaSelic(Obrig.obCodigoObrigacao)
        'Bdados.FechaTabela (rsContaContribuinte)
        'FECHA CALCULO SELIC
        If (Not ComNovaData) And (NovaData = .SelectedItem.SubItems(5)) Then
        Else
            NovaMulta = Conta.CalculaValoresMultaAvulsos(.SelectedItem.SubItems(11), .SelectedItem.SubItems(2), EtcCreditoTributario, NovaData, .SelectedItem.SubItems(5), .SelectedItem.SubItems(6) + Correcao + NovoJuro, , DescMulta)
        End If
        
        Taxa = Taxa + TotalTaxas
        If .SelectedItem.SubItems(3) = Imposto.NomeTributo(ttr_IPTU) Then
            If Trim(EnderecoProprietario) = "" Then
                EnderecoImovel = Endereco
                BuscaContribuinte Proprietario, Nothing, EnderecoProprietario, , etiContribuinte
                EnderecoContribuinte = EnderecoProprietario
            Else
                EnderecoContribuinte = EnderecoProprietario
            End If
        Else
            EnderecoContribuinte = Endereco
        End If
        
        'BCP
        Dim rdt As VSRecordset
        If Bdados.AbreTabela("SELECT desconto,juros,multa,atualizacao,observacao,logradouro,nome_logradouro,numero_endereco,bairro,num_documento,cod_imposto,sigla,nome,cpf_cnpj,cod_cliente FROM vis_Conta_Contribuinte where num_documento=" & Obrig.obCodigoObrigacao, rdt) Then
              If Desconto = 0 Then
                If Not IsNull(rdt(0)) Then
                    Desconto = CCur(rdt(0))
                End If
              End If
              If NovoJuro = 0 Then
                If Not IsNull(rdt(1)) Then
                    NovoJuro = CCur(rdt(1))
                End If
              End If
              If NovaMulta = 0 Then
                If Not IsNull(rdt(2)) Then
                    NovaMulta = CCur(rdt(2))
                End If
              End If
              If Correcao = 0 Then
                If Not IsNull(rdt(3)) Then
                    Correcao = CCur(rdt(3))
                End If
              End If
              If Len(Obs) = 0 Then
                If Not IsNull(rdt(4)) Then
                    Obs = rdt(4)
                End If
              End If
              
              Dim e As String
              e = IIf(IsNull(rdt("logradouro")), "", rdt("logradouro"))
             e = e & " " & IIf(IsNull(rdt("nome_logradouro")), "", rdt("nome_logradouro"))
             e = e & " " & IIf(IsNull(rdt("numero_endereco")), "", rdt("numero_endereco"))
             e = e & ", " & IIf(IsNull(rdt("bairro")), "", rdt("bairro"))
              
        End If
        
        'BCP
        'If Util.Confirma("Este Boleto irá na remessa diária") = True Then
            Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_REMESSA = NULL, TOC_OBSERVACAO='" & Obs & "' WHERE TOC_COD_OBRIGACAO=" & Obrig.obCodigoObrigacao)
        'Else
         '   Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_REMESSA = 1, TOC_OBSERVACAO='" & Obs & "' WHERE TOC_COD_OBRIGACAO=" & Obrig.obCodigoObrigacao)
        'End If
        'FIM BCP
        'FIM BCP
        DescImposto = .SelectedItem.SubItems(12) & IIf(Nvl(.SelectedItem.SubItems(13), 0) = 5, " - DAT", "")
        Cobranca.ImprimeDam RPT, rdt("num_documento"), rdt("cod_cliente"), rdt("nome"), IIf(IsNull(rdt("cpf_cnpj")), "", rdt("cpf_cnpj")), e, .SelectedItem.SubItems(1), EnderecoImovel, rdt("cod_imposto"), rdt("sigla"), _
             DescImposto, .SelectedItem.SubItems(4), .SelectedItem.SubItems(12), 4, IIf(NovaData = "", .SelectedItem.SubItems(5), NovaData), BaseCalculo, .SelectedItem.SubItems(6), NovaMulta, NovoJuro, .SelectedItem.SubItems(10), Obrig.obValorDesconto, "", Obs, _
             PicBarra, NotaInicial, NotaFinal, .SelectedItem.SubItems(14), , , , , , , , DetinoDAM, etdNormal, StringTaxas, .SelectedItem.SubItems(5), Correcao, .SelectedItem.SubItems(15), DescMulta + DescJuros + DescCorrecao
    End With
    
End Sub
Public Sub ImprimeSelecionado_Cotas_Parceladas(Lista As Object, Razao As String, Endereco As String, Optional ComNovaData As Boolean = False, Optional NovaData As String, Optional DetinoDAM As TipoDestinoImpressao = tdiTela, Optional StringTaxas As String, Optional TotalTaxas As Double, Optional Obrigacao As String)
    Dim Sql As String
    Dim CpfCgc As String
    Dim Declaracao As New cDeclaracao
    Dim Desconto As Double
    Dim Exer As String
    Dim BaseCalculo As String
    Dim Obs As String
    Dim Imposto As New VSImposto
    Dim Lote As New BCI
    Dim Cobranca As New VSCobranca
    Dim NovoJuro As String
    Dim NovaMulta As String
    Dim Conta As New ContaCorrente
    Dim Obrig As New Obrigacao
    Dim NotaInicial As String
    Dim NotaFinal As String
    Dim strValorTaxa As String
    Dim strNomeTaxa As String
    Dim Rs_Carro As VSRecordset
    Dim Rs_Dividas As VSRecordset
    Dim Correcao As Double
    
    Obs = ""
    Obs = Trim(Util.Entrada("Observacao", "Impressão de Parcela."))
    With Lista
  
        Exer = IIf(Len(.SelectedItem.SubItems(3)) = 4, .SelectedItem.SubItems(3), Right(.SelectedItem.SubItems(3), 2) & Left(.SelectedItem.SubItems(3), 4))
        strValorTaxa = 0
        If .SelectedItem.SubItems(2) = Imposto.NomeTributo(ttr_IPTU) Or .SelectedItem.SubItems(2) = Imposto.NomeTributo(ttr_ITU) Then
            If CDate(.SelectedItem.SubItems(4)) >= Date Then
                Desconto = .SelectedItem.SubItems(6) * (CInt(Nvl(Bdados.BuscaCodigo("Select TPP_DESCONTO from TAB_PARAMETRO_PARCELA_IPTU WHERE tpp_parcela=0 and TPP_ANO=" & Exer), 0)) / 100)
            End If
        End If
        'CpfCgc = .SelectedItem.SubItems(14)
        If .SelectedItem.SubItems(2) = Imposto.NomeTributo(ttr_ISSQN) Then
            If Declaracao.Buscar(.SelectedItem.SubItems(1), .SelectedItem.SubItems(3)) Then
                BaseCalculo = Declaracao.Itens(4).Valor
                NotaInicial = Declaracao.Itens(1).Valor
                NotaFinal = Declaracao.Itens(2).Valor
'                Obs = "Nº da Nota Inicial: " & declaracao.Itens(1).VALOR & " - Nº da Nota Final: " & declaracao.Itens(2).VALOR
            Else
                BaseCalculo = (.SelectedItem.SubItems(6) * 100) / ((Imposto.BuscaAliquota(Imposto.BuscaCodImposto(.SelectedItem.SubItems(2)), Left(.SelectedItem.SubItems(3), 4))) * 100)
                Obs = ""
            End If
        ElseIf .SelectedItem.SubItems(2) = Imposto.NomeTributo(ttr_IPTU) Then
            BaseCalculo = Format(Lote.ValorVenal(.SelectedItem.SubItems(1)), Const_Monetario)
        Else
            BaseCalculo = .SelectedItem.SubItems(6)
        End If
        Set Declaracao = Nothing
        If (Not ComNovaData) Or (NovaData = .SelectedItem.SubItems(4)) Then
            NovaData = .SelectedItem.SubItems(4)
            NovoJuro = .SelectedItem.SubItems(7)
            NovaMulta = .SelectedItem.SubItems(8)
        Else
            Correcao = Conta.CalculaValoresCorrecaoAvulso(.SelectedItem.SubItems(11), .SelectedItem.SubItems(2), .SelectedItem.SubItems(5), NovaData, .SelectedItem.SubItems(6))
            NovoJuro = Conta.CalculaValoresJurosAvulsos(.SelectedItem.SubItems(9), .SelectedItem.SubItems(3), EtcCreditoTributario, NovaData, .SelectedItem.SubItems(4), .SelectedItem.SubItems(6) + Correcao)
            NovaMulta = Conta.CalculaValoresMultaAvulsos(.SelectedItem.SubItems(9), .SelectedItem.SubItems(3), EtcCreditoTributario, NovaData, .SelectedItem.SubItems(4), .SelectedItem.SubItems(6) + Correcao)
        End If
        Dim PicBarra As Object
        Dim DescImposto As String
        Dim Taxa As Double
        
        'If Aplicacoes.Municipio = "PETROLINA" Then
        '    Taxa = strValorTaxa + TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
        'Else
        '    Taxa = strValorTaxa
        'End If
        Taxa = Taxa + TotalTaxas
        DescImposto = .SelectedItem.SubItems(5)  '& IIf(Nvl(.SelectedItem.SubItems(13), 0) = 5, " - DAT", "")
'        If Trim(.SelectedItem.SubItems(15)) <> "" Then Obs = "Documento de Origem: " & .SelectedItem.SubItems(15)
        Cobranca.ImprimeDam RPT, Obrigacao, .SelectedItem.SubItems(1), Razao, CpfCgc, Endereco, .SelectedItem.SubItems(1), Endereco, .SelectedItem.SubItems(9), .SelectedItem.SubItems(2), _
             DescImposto, Exer, .SelectedItem.SubItems(5), 4, NovaData, BaseCalculo, .SelectedItem.SubItems(6) - Desconto, NovaMulta, NovoJuro, CDbl(Taxa), 0, "", Obs, _
             PicBarra, NotaInicial, NotaFinal, , , , , , , , , tdiTela, etdNormal, StringTaxas, , Correcao
    End With
End Sub





Public Sub ImprimeSelecionado_Cotas_Lancadas(Lista As Object, Razao As String, Endereco As String, Optional ComNovaData As Boolean = False, Optional NovaData As String, Optional DetinoDAM As TipoDestinoImpressao = tdiTela, Optional StringTaxas As String, Optional TotalTaxas As Double, Optional Obrigacao As String)
    Dim Sql As String
    Dim CpfCgc As String
    Dim Declaracao As New cDeclaracao
    Dim Desconto As Double
    Dim Exer As String
    Dim BaseCalculo As String
    Dim Obs As String
    Dim Imposto As New VSImposto
    Dim Lote As New BCI
    Dim Cobranca As New VSCobranca
    Dim NovoJuro As String
    Dim NovaMulta As String
    Dim Conta As New ContaCorrente
    Dim Obrig As New Obrigacao
    Dim NotaInicial As String
    Dim NotaFinal As String
    Dim strValorTaxa As String
    Dim strNomeTaxa As String
    Dim Rs_Carro As VSRecordset
    Dim Rs_Dividas As VSRecordset
    
    
    Obs = ""
    Obs = Trim(Util.Entrada("Observacao", "Impressão de Parcela."))
    With Lista
  
        Exer = IIf(Len(.SelectedItem.SubItems(3)) = 4, .SelectedItem.SubItems(3), Right(.SelectedItem.SubItems(3), 2) & Left(.SelectedItem.SubItems(3), 4))
        strValorTaxa = 0
        If .SelectedItem.SubItems(2) = Imposto.NomeTributo(ttr_IPTU) Or .SelectedItem.SubItems(2) = Imposto.NomeTributo(ttr_ITU) Then
            If CDate(.SelectedItem.SubItems(4)) >= Date Then
                Desconto = .SelectedItem.SubItems(6) * (CInt(Nvl(Bdados.BuscaCodigo("Select TPP_DESCONTO from TAB_PARAMETRO_PARCELA_IPTU WHERE tpp_parcela=0 and TPP_ANO=" & Exer), 0)) / 100)
            End If
        End If
        'CpfCgc = .SelectedItem.SubItems(14)
        If .SelectedItem.SubItems(2) = Imposto.NomeTributo(ttr_ISSQN) Then
            If Declaracao.Buscar(.SelectedItem.SubItems(1), .SelectedItem.SubItems(3)) Then
                BaseCalculo = Declaracao.Itens(4).Valor
                NotaInicial = Declaracao.Itens(1).Valor
                NotaFinal = Declaracao.Itens(2).Valor
'                Obs = "Nº da Nota Inicial: " & declaracao.Itens(1).VALOR & " - Nº da Nota Final: " & declaracao.Itens(2).VALOR
            Else
                BaseCalculo = (.SelectedItem.SubItems(6) * 100) / ((Imposto.BuscaAliquota(Imposto.BuscaCodImposto(.SelectedItem.SubItems(2)), Left(.SelectedItem.SubItems(3), 4))) * 100)
                Obs = ""
            End If
        ElseIf .SelectedItem.SubItems(2) = Imposto.NomeTributo(ttr_IPTU) Then
            BaseCalculo = Format(Lote.ValorVenal(.SelectedItem.SubItems(1)), Const_Monetario)
        Else
            BaseCalculo = .SelectedItem.SubItems(6)
        End If
        Set Declaracao = Nothing
'        If .SelectedItem.SubItems(15) <> "" Then
'            'Raimundo Processo que ira processar
'            'os juros e multas para o alvara de acordo com
'            'o inicio da atividade do carro...
'            sql = "SELECT * "
'            sql = sql & " From Tab_Parametro_Imposto"
'            sql = sql & " where tpi_tip_cod_imposto =" & Bdados.Converte(.SelectedItem.SubItems(11), tctexto)
'            sql = sql & " and tpi_ano_imposto= " & Bdados.Converte(.SelectedItem.SubItems(4), tctexto)
'            If Bdados.AbreTabela(sql, Rs_Dividas) Then
'                'Pego a data de inicio de atividade do carro
'                sql = "SELECT * "
'                sql = sql & " From TAB_TRANSPORTADOR_VEICULO, TAB_ATIVIDADE_ECONOMICA"
'                sql = sql & " WHERE TTV_TCI_IM = " & Bdados.Converte(.SelectedItem.SubItems(1), tctexto)
'                sql = sql & " AND TTV_TAE_CAE = TAE_CAE"
'                sql = sql & " AND  TTV_PLACA = " & Bdados.Converte(.SelectedItem.SubItems(15), tctexto)
'                If Bdados.AbreTabela(sql, Rs_Carro) Then
'                    Dim Data_Imposto          As String
'                    Dim Prazo                 As Integer
'                    Dim Dias                  As Integer
'                    Dim Multa                 As Double
'                    Dim Juros                 As Double
'                    Dim Valor_Min             As Integer
'                    Dim Valor_Max             As Integer
'
'                    Dim Inicio_Atividade      As Date
'                    Dim Dias_Corridos         As Integer
'                    Dim Meses_Corridos        As Integer
'                    Dim Valor_Juros_Apagar    As Double
'                    Dim Valor_Multa_Apagar    As Double
'                    Dim Valor_Atividade       As Double
'
'                   'Dados da Definição de tributos..
'                    Data_Imposto = Rs_Dividas.Fields("tpi_dt_inicio_imposto")
'                    Prazo = Rs_Dividas.Fields("tpi_dias_declara")
'                    Dias = Rs_Dividas.Fields("tpi_dias_pagar")
'                    Valor_Min = Rs_Dividas.Fields("tpi_valor_min_multa")
'                    Valor_Max = Rs_Dividas.Fields("tpi_valor_max_multa")
'                    Juros = Rs_Dividas.Fields("tpi_valor_juros")
'                   'Dados da tabela de Transporte e da tabela de atividades economicas...
'                    Inicio_Atividade = Rs_Carro.Fields("TTV_INICIO_ATIVIDADE")
'                    Valor_Atividade = Rs_Carro.Fields("TAE_VALOR")
'                    'Se o inicio das atividades estive entre a data do imposto e o prazo de dias entra no if
'                    If CDate(Inicio_Atividade) >= CDate(Data_Imposto) And CDate(Inicio_Atividade) <= DateAdd("d", Dias, Data_Imposto) Then
'                       'Pego a difernca em dias...
'                        Dias_Corridos = DateDiff("d", CDate(Data_Imposto), CDate(Inicio_Atividade))
'
'                        Valor_Multa_Apagar = (Valor_Min * Valor_Atividade) / 100
'                        If Dias_Corridos > 31 Then
'                           Dim Divi As Integer
'                           Divi = Dias_Corridos \ 31
'                           Valor_Juros_Apagar = Divi * Valor_Atividade / 100
'                           NovoJuro = Valor_Juros_Apagar
'                           If Dias_Corridos - (Dias_Corridos \ 31) > 0 Then
'                               NovaMulta = ((31 - (Dias_Corridos \ 31)) * Valor_Atividade) / 100
'                           End If
'                        Else
'                            If Dias_Corridos > 0 Then
'                                NovaMulta = Valor_Multa_Apagar * Dias_Corridos
'                            End If
'                        End If
'                       'Pego a diferenca em meses...
'                        Meses_Corridos = DateDiff("m", CDate(Data_Imposto), CDate(Inicio_Atividade))
'                        Valor_Juros_Apagar = (Juros * Valor_Atividade) / 100
'                        If Meses_Corridos > 0 Then
'                            NovoJuro = Valor_Juros_Apagar * Meses_Corridos
'                        End If
'                    End If
'                End If
'            End If
'
'        Else
            If (Not ComNovaData) Or (NovaData = .SelectedItem.SubItems(4)) Then
                NovaData = .SelectedItem.SubItems(4)
                NovoJuro = .SelectedItem.SubItems(7)
                NovaMulta = 0 '.SelectedItem.SubItems(8)
            Else
                NovoJuro = Conta.CalculaValoresJurosAvulsos(.SelectedItem.SubItems(9), .SelectedItem.SubItems(3), EtcCreditoTributario, NovaData, .SelectedItem.SubItems(4), .SelectedItem.SubItems(6))
                NovaMulta = Conta.CalculaValoresMultaAvulsos(.SelectedItem.SubItems(9), .SelectedItem.SubItems(3), EtcCreditoTributario, NovaData, .SelectedItem.SubItems(4), .SelectedItem.SubItems(6))
            End If
'        End If
        Dim PicBarra As Object
        Dim DescImposto As String
        Dim Taxa As Double
       
        'If Aplicacoes.Municipio = "PETROLINA" Then
        '    Taxa = strValorTaxa + TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
        'Else
        '    Taxa = strValorTaxa
        'End If
        Taxa = Taxa + TotalTaxas
        DescImposto = .SelectedItem.SubItems(10)   '& IIf(Nvl(.SelectedItem.SubItems(13), 0) = 5, " - DAT", "")
'        If Trim(.SelectedItem.SubItems(15)) <> "" Then Obs = "Documento de Origem: " & .SelectedItem.SubItems(15)
        Cobranca.ImprimeDam RPT, Obrigacao, .SelectedItem.SubItems(1), Razao, CpfCgc, Endereco, .SelectedItem.SubItems(1), Endereco, .SelectedItem.SubItems(9), .SelectedItem.SubItems(2), _
             DescImposto, Exer, .SelectedItem.SubItems(5), 4, NovaData, BaseCalculo, .SelectedItem.SubItems(6) - Desconto, NovaMulta, NovoJuro, CDbl(Taxa), 0, "", Obs, _
             PicBarra, NotaInicial, NotaFinal, , , , , , , , , tdiTela, etdNormal, StringTaxas, .SelectedItem.SubItems(4)
    End With
End Sub

