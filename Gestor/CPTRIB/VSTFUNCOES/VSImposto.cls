VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSImposto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public DataVenc As String
Public ValorTaxas As Double
Public TaxaFixaImpressaoDAM As Double

Dim Conta As ContaCorrente
Dim ClsImposto As New VSImposto
Public Enum TipoLogradouro
    Rua = 0
    Bairro = 1
    TipoLogr = 2
End Enum
Public Enum tipoInsc
    InscContrib = 0
    InscImovel = 1
    InscGrupo = 3
    InscCpfCnpj = 4
    
End Enum

Public Enum TipoArea
    Construida = 105
    Terreno = 108
End Enum

Enum TipoBoletim
    tbo_Predial = 0
    tbo_Territorial
    tbo_Condominio
End Enum

Enum TipoGeracaoImposto
    tgi_SemParcelas = 0
    tgi_SoParcelas
    tbo_Normal
End Enum

Enum CidadeIptu
    cip_Colinas = 0
    cip_Balsas = 1
    cip_BCorda = 2
End Enum

Enum TipoImovel
    tim_TerrenoSemEdif = 0
    tim_TerrenoComEdif = 5
    tim_Edificacao = 10
    tim_Gleba = 11
End Enum

Enum TipoIsencaoImovel
    tim_SemIsencao = 1
    tii_IsentoImposto
    tii_IsentoTaxas
    tim_IsentoTotal
End Enum

Enum TipoZona
    tzo_ZonaUrbana = 0
    tzo_ZonaNaoUrbana = 6
End Enum

Enum TipoTributo
    ttr_IPTU = 1
    ttr_ISSQN = 2
    ttr_ISSQNRET = 6
    ttr_ISSQNCOMP = 9
    ttr_ISSQNSUBST = 19
    ttr_ISSQNFIXO = 35
    ttr_ITBI = 3
    ttr_ALVARA = 4
    ttr_TAXAS = 5
    ttr_IRPJ = 7
    ttr_IRPF = 8
    ttr_2VIA = 10
    ttr_AFORO = 11
    ttr_VISTORIA = 12
    ttr_DESMEMBRAMENTO = 13
    ttr_CONSTRUCAO = 14
    ttr_HABIT = 15
    ttr_RECONSTRUCAO = 16
    ttr_REMEMBRA = 17
    ttr_DATIVA = 18
    ttr_ITU = 20
    ttr_TXLIMPA = 21
    ttr_TXCONSERV = 22
    ttr_TXCOLETA = 23
    ttr_ATRANS = 24
    ttr_IPTUParcela = 31
    ttr_TSU = 32
    ttr_TEMISSAODAM = 33
    ttr_TEMISSAOALVARA = 34
    ttr_PUBLICIDADE = 36
    ttr_ISSQNEST = 37
    ttr_TXVIG = 40
    ttr_TFA = 50
    ttr_TFOP = 51
    ttr_TFAF = 52
    ttr_TFV = 53
    ttr_tfs = 54
    ttr_ISSFIXO = 55
    ttr_AUTO_INFRACAO = 56
End Enum

Private Type ComponentesImovel
    Valor As Double
    ComponenteFator As Integer
End Type

Private Componentes(1000) As ComponentesImovel
Function BuscaCoeficienteImovelIPTU(Area As TipoZona, Zona As Integer) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    
    BuscaCoeficienteImovelIPTU = 0
    Sql = "Select TGE_NOME from TAB_GERAL WHERE TGE_TIPO = 708 AND TGE_CODIGO = " & IIf(Area + Zona <> 0, Area + Zona, 3)
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaCoeficienteImovelIPTU = Nvl("" & Rs(0), 1)
    End If
    Bdados.FechaTabela Rs
End Function

Function BuscaCoeficienteEdificacao(IC As String, Unidade As Integer) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Padrao As Integer
    
    Sql = "select tdi_valor_item from TAB_DETALHE_IMOVEL where tdi_tim_ic='" & Trim(IC) & _
        "' and tdi_tgc_cod_grupo=12  and tdi_tim_ic_unidade=" & Unidade
    If Bdados.AbreTabela(Sql, Rs) Then
        Padrao = Nvl("" & Rs(0), 1)
    End If
    BuscaCoeficienteEdificacao = 0
    If Padrao = 0 Then Exit Function
    Sql = "Select TGE_NOME from TAB_GERAL WHERE TGE_TIPO =707 AND TGE_CODIGO = " & Padrao
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaCoeficienteEdificacao = Nvl("" & Rs(0), 1)
    End If
    Bdados.FechaTabela Rs
End Function

Public Function BuscaDataGeracaoDam(NumPago As String) As String
    Dim Rs As VSRecordset
    Dim Sql As String
    Sql = "Select tgt_data_geracao from tab_geracao_tributo where tgt_cod_pagamento='" & NumPago & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaDataGeracaoDam = IIf(IsNull(Rs!tgt_data_geracao), Date, Rs!tgt_data_geracao)
    Else
        BuscaDataGeracaoDam = Date
    End If
    Bdados.FechaTabela Rs
End Function

Public Function BuscaValorMinimoImposto(Imposto As String, Ano As String) As Double
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Condicao As String
    Sql = "Select tpi_valor_min_imposto from Tab_Parametro_Imposto " & _
        " where tpi_tip_cod_imposto='" & Imposto & "'"
    Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Ano, 4)) & "'"
    If Bdados.AbreTabela(Sql & Condicao, Rs) Then
        BuscaValorMinimoImposto = Nvl("" & Rs(0), 0)
    End If
    Bdados.FechaTabela Rs
End Function

Public Function BuscaValorMinimoBase(Imposto As String, Ano As String) As Double
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Condicao As String
    Sql = "Select tpi_valor_min_base_calc from Tab_Parametro_Imposto " & _
        " where tpi_tip_cod_imposto='" & Imposto & "'"
    Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Ano, 4)) & "'"
    If Bdados.AbreTabela(Sql & Condicao, Rs) Then
        BuscaValorMinimoBase = Rs(0)
    End If
    Bdados.FechaTabela Rs
End Function
Private Function BuscaValorMinimoImpostoBairro(CodBairro As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    On Error Resume Next
    Sql = "Select tba_valor_minimo_imposto from tab_bairro where tba_cod_bairro =" & CodBairro
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaValorMinimoImpostoBairro = CDbl(Nvl("" & Rs!tba_valor_minimo_imposto, 0))
    End If
End Function

Public Function CalculaAlvara(Im As String, Periodo As Integer, Taxa As Double, NomeImposto As String, ByVal DtVenc As String, CodPagamento As String, CodTaxa As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim Imposto As String
    Dim DataInicio As String
    Dim ValorAlvara As Double
    Dim RsData As VSRecordset
    Dim Novo As Boolean
    Dim Estima As Boolean
    Dim CAE As String
    Dim FatorAlvara As Double
    Dim AtividadesSecundarias As Integer
    Dim PercAtivSecund As Double
    Sql = "select tci_inicio_atividade,tci_tae_cae,tci_fator_alvara,tci_tae_cae_secund,tci_tae_cae_terc from tab_contribuinte where tci_im ='" & Im & "'"
    If Bdados.AbreTabela(Sql, RsData) Then
        If IsNull(RsData(0)) Then Exit Function
        'PEGA PARAMETROS
        DataInicio = RsData!TCI_INICIO_ATIVIDADE
        CAE = RsData!tci_tae_cae
        FatorAlvara = RsData!tci_fator_alvara
        If Format(Year(DataInicio), "0000") = Left(CStr(Periodo), 4) Then
            DataInicio = 13 - Month(DataInicio)
        End If
        PercAtivSecund = CDbl(Nvl(BuscaNaGeral("ALVARA SECUNDARIO", 1), 0))
        AtividadesSecundarias = PercAtivSecund * CDbl(IIf(Nvl("" & RsData!tci_tae_cae_secund, 0) = 0, 0, 1) + IIf(Nvl("" & RsData!tci_tae_cae_terc, 0) = 0, 0, 1))
        Imposto = BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_nome_imposto='" & NomeImposto & "'")
        DtVenc = BuscaDataVencimento(Imposto, CLng(Periodo))
        CodPagamento = Conta.NumPagamento(Im, CLng(Periodo), Imposto, "", 0, Novo, 0)
        'VERIFICA SE ATIVIDADE É ESTIMATIVA
        Sql = "SELECT   TAT_LIMITE_INFERIOR,TAT_LIMITE_SUPERIOR,TAT_LIMITE_VALOR,TAT_LIMITE_FATOR FROM TAB_ATIVIDADE_ESTIMADA, TAB_ATIVIDADE_ECONOMICA WHERE TAE_CAE=TAT_TAE_CAE AND TAE_ESTIMATIVO=2 AND TAT_TAE_CAE = " & CAE & " order by TAT_LIMITE_INFERIOR asc"
        If Bdados.AbreTabela(Sql, Rs) Then
            Rs.MoveFirst
            Do
                If Nvl("" & Rs!TAT_LIMITE_SUPERIOR, 0) = 0 Then 'ULTIMA FAIXA
                    If FatorAlvara >= Rs!TAT_LIMITE_INFERIOR Then
                        ValorAlvara = IIf(Rs!TAT_LIMITE_FATOR <> 0, FatorAlvara, 1) * Rs!TAT_LIMITE_VALOR
                    End If
                ElseIf FatorAlvara >= Rs!TAT_LIMITE_INFERIOR And FatorAlvara <= Rs!TAT_LIMITE_SUPERIOR Then
                    ValorAlvara = IIf(CDbl(Nvl(Trim(Rs!TAT_LIMITE_FATOR), 0)) <> 0, FatorAlvara, 1) * Rs!TAT_LIMITE_VALOR
                    Exit Do
                End If
                Rs.MoveNext
            Loop While Not Rs.EOF
            Rs.Fechar
        Else
            Sql = "Select tci_fator_alvara,tae_valor,tip_cod_imposto,tae_fator from Tab_Contribuinte, tab_atividade_economica,Tab_Imposto where "
            Sql = Sql & IIf(Len(Im) = 11, "tci_im", "tci_cgc_cpf") & " = '" & Im & "' and tci_tae_cae = tae_cae and tip_cod_imposto ='" & Imposto & "'"
            
            If Bdados.AbreTabela(Sql, Rs) Then
                ValorAlvara = IIf(Rs!tae_fator = 0, Rs!TAE_VALOR, Rs!tci_fator_alvara * Rs!TAE_VALOR)
            End If
        End If
        'CALCULA ALVARA PROPORCIONAL
        If IsNumeric(DataInicio) Then
            ValorAlvara = ValorAlvara * (CInt(DataInicio) / 12)
            CalculaAlvara = ValorAlvara
            If DateDiff("d", RsData(0), DtVenc) < 0 Then
                DtVenc = UltimoDiaDoMes(IIf(Day(Date) <= 20, Date, Date + 20))
            End If
        Else
            CalculaAlvara = ValorAlvara
        End If
        'GRAVA VALOR ALVARA
        CalculaAlvara = ValorAlvara + ValorAlvara * (AtividadesSecundarias / 100)
        If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
            Dim Descontos As New CalculoObrigacao
            Dim Desconto As Double
            'PEGO O DESCONTO DE ACORDO COM O BAIRRO E O FATOR
            Desconto = Descontos.Pega_Desconto_Alvara(Im)
            If Desconto <> 0 Then
                CalculaAlvara = CalculaAlvara - (Desconto * ValorAlvara / 100)
            End If
            
        End If
'        Conta.GeraPagamento Im, "", Imposto, CLng(Periodo), DtVenc, CalculaAlvara, 0, 0, CodPagamento, 0, 0, Taxa, CodTaxa
    End If
    Bdados.FechaTabela Rs
End Function
Private Function CalculaValorTerreno(ByVal Imovel As String, Area As TipoArea) As Double
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Logradouro As Double
    Sql = "Select tim_tlg_cod_logradouro from Tab_Imovel where tim_ic='" & Imovel & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        Logradouro = Rs(0)
        Sql = "select    tvl_valor * tdi_valor_item from Tab_Valor_Terreno,Tab_Detalhe_Imovel " & _
            " where tvl_tlg_cod_logradouro='" & Logradouro & "' and tdi_tim_ic='" & Imovel & "'"
        Sql = Sql & " and tdi_tco_cod_componente=" & Area
        If Bdados.AbreTabela(Sql, Rs) Then
            CalculaValorTerreno = Rs(0)
        End If
    End If
    Bdados.FechaTabela Rs
End Function

Private Function CalculaTaxasIptu(ByVal Imovel As String, ByVal Contribuinte As String, ByVal Imposto As String, ByVal Periodo As Long, ByRef ValorIptu As Double, Optional Cidade As CidadeIptu, Optional Zona As Integer, Optional TipoImovel As Integer) As Double
    Dim Sql As String
    Dim RsTax As VSRecordset
    Dim Rs As VSRecordset
    Dim Aliquota As Double
    Dim Referencia As Integer
    Dim Relacao As String
    Dim CodIPTU As String
    Dim Taxa As String
    Dim ValorTaxa As Double
    Select Case Cidade
        Case cip_Balsas
            Sql = "Select tdi_valor_item from TAB_DETALHE_IMOVEL where tdi_tim_ic ='" & IIf(CInt(Right(Trim(Imovel), 3)) >= 200, Left(Imovel, 12) & "200", Imovel) & "' and tdi_tgc_cod_grupo = 100 "
            If Bdados.AbreTabela(Sql, Rs) Then
                ValorTaxa = Rs(0) + ValorTaxa
                CalculaTaxasIptu = ValorTaxa * IIf(Zona = 1, IIf(TipoImovel = 1, 3, 2), IIf(Zona = 5, 0, IIf(TipoImovel = 1, 1, 0)))
            End If
            Bdados.FechaTabela Rs
        Case Else
            Sql = "SELECT tax.tti_tip_cod_imposto,mon.tmo_valor*tax.tti_aliquota, " & _
                " tax.tti_tco_cod_componente ,tax.tti_tco_cod_componente_relacao " & _
                " FROM Tab_Taxa_Iptu as tax,Tab_Monetaria as mon" & _
                " GROUP BY tax.tti_tip_cod_imposto , mon.tmo_valor*tax.tti_aliquota, " & _
                " tax.tti_tco_cod_componente,tax.tti_tco_cod_componente_relacao," & _
                " tax.tti_tmo_cod_unidade , Mon.tmo_unidade, tmo_data " & _
                " Having tax.tti_tmo_cod_unidade = mon.tmo_unidade and  tmo_data = " & _
                " (Select MAX(tmo_data) FROM Tab_Monetaria " & _
                " where tmo_unidade = Mon.tmo_unidade)"
            If Bdados.AbreTabela(Sql, RsTax) Then
                RsTax.MoveFirst
                Do While Not RsTax.EOF
                    Taxa = RsTax(0)
                    Referencia = RsTax!tti_tco_cod_componente
                    Relacao = RsTax!tti_tco_cod_componente_relacao
                    Aliquota = RsTax(1)
                    Sql = "SELECT im.tdi_valor_item" & _
                        " FROM Tab_Detalhe_Imovel as Im " & _
                        " WHERE Im.tdi_tim_ic='" & Imovel & _
                        "' and IM.tdi_tco_cod_componente=" & Referencia
                    If Bdados.AbreTabela(Sql, Rs) Then
                        ValorTaxa = Rs(0) * Aliquota
                        ValorIptu = ValorTaxa + ValorIptu
'                        AtualizaTaxasIptu Imovel, Taxa, Periodo, Rs(0)
                    End If
                    RsTax.MoveNext
                Loop
            End If
            Bdados.FechaTabela Rs
            Bdados.FechaTabela RsTax
    End Select
End Function

Private Sub AtualizaTaxasIptu(Imovel As String, Imposto As String, Periodo As Long, Valor As Double)
    Dim Valores As String
    Dim Campos As String
    
    Campos = "tti_tim_ic,tti_tip_cod_imposto,tti_periodo,tti_valor"
    Valores = Bdados.PreparaValor(Bdados.Converte(Imovel, tctexto), Bdados.Converte(Imposto, tctexto), Bdados.Converte(Periodo, TCDuplo), Bdados.Converte(Valor, TCDuplo))
    Bdados.GravaDados "tab_taxa_imovel", Valores, Campos, "tti_tim_ic ='" & Imovel & "' and tti_tip_cod_imposto='" & Imposto & "' and tti_periodo =" & Periodo
End Sub
Private Function BuscaValorRef(Referencia As Integer) As Double
    Dim Rs As VSRecordset
    Dim Query As String
    Query = "SELECT tmo_valor from tab_monetaria as tmo where tmo_unidade=" & _
    "Referencia and tmo_data = (Select MAX(tmo_data) FROM Tab_Monetaria" & _
    " WHERE tmo_unidade = tmo.tmo_unidade)"
    If Bdados.AbreTabela(Query, Rs) Then
        BuscaValorRef = Rs(0)
    End If
    Bdados.FechaTabela Rs
End Function

Private Function BuscaValorM2(ByVal CodLogradouro As String) As Double
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Logradouro As Double
    Sql = "select    tvl_valor from Tab_Valor_Terreno " & _
        " where tvl_tlg_cod_logradouro='" & CodLogradouro & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaValorM2 = Rs(0)
    End If
    Bdados.FechaTabela Rs
End Function

Public Function BuscaCodImposto(SiglaImposto As String, Optional CorrelativoImposto As String) As String
    Dim Query As String
    Dim Rs As VSRecordset
    Dim Flag As Byte
    On Error Resume Next
    Query = "SELECT tip_cod_imposto From Tab_Imposto "
    
    If Trim(SiglaImposto) <> "" Then
        Query = Query & " Where tip_sigla_imposto ='" & SiglaImposto & "' AND TIP_TIC_COD_CATEGORIA <> 8"
        Flag = 1
    End If
    If Trim(CorrelativoImposto) <> "" Then Query = Query & IIf(Flag = 1, " AND ", " WHERE ") & " TIP_COD_CORRELATIVO ='" & CorrelativoImposto & "'"
    If Bdados.AbreTabela(Query, Rs) Then
         BuscaCodImposto = Rs(0)
    End If
    Bdados.FechaTabela Rs
End Function
Public Function Forma_Calculo(Imposto As String, Periodo As String)
    Dim Sql                           As String
    Dim Rs                            As VSRecordset
    
    Sql = "Select * from tab_parametro_imposto "
    Sql = Sql & " WHERE "
    Sql = Sql & " TPI_TIP_COD_IMPOSTO = '" & Imposto & "'"
    Sql = Sql & " AND TPI_ANO_IMPOSTO =  '" & ClsImposto.BuscaAnoImposto(Imposto, Right(Periodo, 4)) & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Not IsNull(Rs.Fields("tpi_forma_calculo")) Then
            Forma_Calculo = Rs.Fields("tpi_forma_calculo")
        Else
            Avisa "Definição de tributo inválida." & vbCrLf & "Campo : Forma de Cálculo"
        End If
    Else
        Avisa "Definição de tributo inválida." & vbCrLf & "Campo : Forma de Cálculo"
    End If

End Function
Public Function BuscaDataVencimento(CodigoImposto As String, ByVal Ano As String) As String
    Dim Sql As String
    Dim Condicao As String
    Dim Rs As VSRecordset
    Dim DiaSemana As Integer
    Dim Mes As Integer
    
    Dim Vencimento As String
    On Error GoTo Trata
    '*******PERIDO VENCIMENTO*******
    '1-ULTIMO DIA DO MES
    '2-DENTRO DO MESMO MES
    '3-NO MES SEGUINTE
    '4-DIA ESPECIFICO
    '5-NO MESMO DIA
    '6-DIAS APOS LANCAMENTO
    BuscaDataVencimento = ""
    If CodigoImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) Then
        Sql = "SELECT TPP_VENCIMENTO FROM Tab_Parametro_Parcela_Iptu WHERE TPP_PARCELA=0 AND TPP_ANO=" & Ano
        If Bdados.AbreTabela(Sql, Rs) Then
            BuscaDataVencimento = "" & Rs!TPP_VENCIMENTO
        End If
        Bdados.FechaTabela Rs
    End If
    
    If BuscaDataVencimento = "" Then
        Sql = "Select TPI_PERIODO_VENCIMENTO,TPI_DIA_BASE,TPI_MES_BASE from Tab_Parametro_Imposto " & _
            " where tpi_tip_cod_imposto='" & CodigoImposto & "'"
        Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(CodigoImposto, Right(Ano, 4)) & "'"
        If Not Bdados.AbreTabela(Sql & Condicao, Rs) Then
            Exit Function
        End If
        If Trim(Ano) = "" Then Ano = Year(Date)
        Select Case CInt(Nvl("" & Rs!TPI_PERIODO_VENCIMENTO, 0))
            Case 1 'ULTIMO DIA DO MES REFERENCIA
                Vencimento = UltimoDiaDoMes(Date)
                
            Case 2 'DENTRO DO MESMO MES REFERENCIA
                Vencimento = Format(Nvl("" & Rs!TPI_DIA_BASE, 1), "00") & "/" & Format(Month(Date), "00") & "/" & Format(Right(Ano, 4), "0000")
                 If Left(Vencimento, 2) = "00" Then
                    Vencimento = "01" & Right(Vencimento, Len(Vencimento) - 2)
                End If
                If DateDiff("m", Vencimento, Date) = 0 And DateDiff("d", Date, Vencimento) < 0 Then
                    Vencimento = UltimoDiaDoMes(Date)
                End If
                
            Case 3 'NO MES REFERENCIA SEGUINTE
                If Len(Ano) = 6 Then
                    Mes = Left(Ano, 2)
                Else
                    Mes = Month(Date)
                End If
                If Mes < 12 Then
                    Vencimento = Format(Nvl("" & Rs!TPI_DIA_BASE, 1), "00") & "/" & Format(Mes + 1, "00") & "/" & Format(Right(Ano, 4), "0000")
                Else
                    Vencimento = Format(Nvl("" & Rs!TPI_DIA_BASE, 1), "00") & "/" & "01" & "/" & Format(CInt(Right(Ano, 4)) + 1, "0000")
                End If
            Case 4 'DIA ESPECIFICO
                Vencimento = Format(Nvl("" & Rs!TPI_DIA_BASE, 1), "00") & "/" & Format(Nvl("" & Rs!TPI_MES_BASE, 1), "00") & "/" & Format(Right(Ano, 4), "0000")
            Case 5 'NO MESMO DIA DO LANÇAMENTO
                Vencimento = Date
            Case 6 'DIAS APOS LANCAMENTO
                Vencimento = DateAdd("d", Nvl("" & Rs!TPI_DIA_BASE, 1), Date)
        End Select
        
        If Not IsDate(Vencimento) Then
            Informa "Definicão inválida do vencimento em MANUTENCÃO DE TRIBUTOS.  Vencimento encontrado = " & Vencimento & "."
            Vencimento = ""
            Exit Function
        End If
        If IsDate(Trim(Vencimento)) Then
            DiaSemana = Weekday(Vencimento)
            If DiaSemana = 1 Then
                Vencimento = DateAdd("d", 1, Vencimento)
            ElseIf DiaSemana = 7 Then
                Vencimento = DateAdd("d", 2, Vencimento)
            End If
        End If
        
        If Year(Vencimento) > 2030 Or Year(Vencimento) < 1910 Then
            Vencimento = Left(Vencimento, 6) & "2000"
        End If
        BuscaDataVencimento = Vencimento
        Bdados.FechaTabela Rs
        Exit Function
        If Rs!TPI_PERIODIC_DECLARA = 2 Then
            If IsDate("01/" & Format(Left(Ano, 2), "00") & "/" & Format(Right(Ano, 4), "0000")) Then
                BuscaDataVencimento = DateAdd("d", Rs!tpi_dias_pagar - 1, "01/" & Format(Left(Ano, 2), "00") & "/" & Format(Right(Ano, 4), "0000"))
            Else
                BuscaDataVencimento = ""
            End If
        ElseIf Rs!TPI_PERIODIC_DECLARA = 1 Then
            If Trim(Ano) <> "" Then
                BuscaDataVencimento = DateAdd("d", _
                Rs(0), _
                "01/" & _
                IIf(Len(Ano) = 4, "01", Left(Ano, 2)) & _
                "/" & _
                IIf(Len(Ano) = 4, Ano, Right(Ano, 4)))
            Else
                BuscaDataVencimento = 0
            End If
        ElseIf Rs(1) = 5 Then
            BuscaDataVencimento = Date + Rs!tpi_dias_pagar
        End If
        Bdados.FechaTabela Rs
    End If
    Exit Function
Trata:
    Avisa Err.Description
    Exit Function
    Resume
End Function

'Public Function CriaIptu_Novo(Contribuintes As VSRecordset, PeriodoInicial As Integer, PeriodoFinal As Integer) As Double
'    Dim QueryIptu As String
'    Dim RsImovel As VSRecordset
'    Dim RsIptu As VSRecordset
'    Dim ValorCalculado As Double
'    Dim Valores As String
'    Dim Campos As String
'    Dim DataVencimento As String
'    Dim i As Long
'    Dim CodImposto As String
'    Dim RefComp As Integer
'    Dim Total As Double
'    Dim Taxas As Double
'    Dim Reducao As Double
'    Dim Reg As Double
'    Dim CodConta  As Double
'    Dim Aliquota As Double
'    Dim Componente As Integer
'    Dim AreaConstruida As Double
'    Dim Rs As VSRecordset
'    Dim Sql As String
'    Dim ValorImovel As Double
'    Dim TotalPavimentos As Integer
'    Dim ValorPorPavimento As Double
'    Dim RsPav As VSRecordset
'    Dim MaxCotas As Byte
'    Dim RsDesconto As VSRecordset
'    Dim TotalUnico As Double
'    Dim ImpostoMinimo As Double
'    Dim Desconto As Double
'    Dim Novo As Boolean
'    Dim ValorTerreno As Double
'    Dim CodPagamento As Double
'    Dim CodParticular As Integer
'    Dim AliquotaEdif As Double
'    Dim AliquotaTerr As Double
'    Dim ValorPadraoTsu As Double
'    Dim ValorConstrucao As Double
'    Dim Testada As Double
'    Dim EstadoConservacao As Double
'    Dim AreaTerreno As Double
'    Dim Condicao As String
'    Dim CodTaxa  As String
'    Screen.MousePointer = 11
'    CodImposto = BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
'    Dim Gleba As Boolean
'    Reg = 0
'    ValorPadraoTsu = BuscaNaGeral("VALOR TSU", 1)
'    Gleba = False
'    Componente = 0
'    AreaConstruida = 0
'    CodParticular = IIf(Temp.PegaParametro(BDados,BDados,"MUNICIPIO") = 1, 80, 81)
'    AliquotaEdif = CDbl(nvl(BuscaNaGeral("ALIQUOTA IPTU SHELENA", 2), 0))
'    AliquotaTerr = CDbl(nvl(BuscaNaGeral("ALIQUOTA IPTU SHELENA", 1), 0))
'
'    'Pega Maximo de cotas em parcelamento
'    Sql = "SELECT * from Tab_Parametro_Parcelamento"
'    If Bdados.AbreTabela(Sql, Rs) Then
'        MaxCotas = Rs!tpp_max_cotas
'    End If
'    Bdados.FechaTabela Rs
'    'Pega desconto em cota unica
'    Sql = "Select TGE_NOME from tab_geral where TGE_TIPO = 755 and TGE_CODIGO > 0"
'    If Bdados.AbreTabela(Sql, RsDesconto) Then
'        TotalUnico = TotalUnico * ((100 - RsDesconto(0)) / 100)
'    End If
'    Bdados.FechaTabela Rs
'    For i = PeriodoInicial To PeriodoFinal
'        DataVencimento = BuscaDataVencimento(BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)), CInt(i))
'        Reducao = BuscaReducao(CodImposto, Right(i, 2))
'        Aliquota = BuscaAliquota(CodImposto, Right(i, 2))
'        'Pega valor mínimo de imposto
'        Sql = "SELECT tpi_valor_min_imposto FROM TAB_PARAMETRO_IMPOSTO WHERE tpi_tip_cod_imposto='" & CodImposto & "'" ' AND tpi_ano_imposto='" & IIf(Len(CStr(i)) = 4, Right(i, 2), Mid(i, 4, 2)) & "'"
'        Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(CodImposto, Left(i, 4)) & "'"
'        If Bdados.AbreTabela(Sql & Condicao, Rs) Then
'            ImpostoMinimo = Rs(0)
'        End If
'        Bdados.FechaTabela Rs
'        If IsDate(DataVencimento) Then
'            Contribuintes.MoveFirst
'            Do While Not Contribuintes.EOF
'                QueryIptu = "SELECT tdi_tco_cod_componente,tdi_valor_item FROM Tab_Detalhe_Imovel WHERE tdi_tim_ic ='" & Contribuintes!tim_ic & "'"
'                Gleba = False
'                If Bdados.AbreTabela(QueryIptu, RsImovel) Then
'                    RsImovel.MoveFirst
'                    Taxas = 0
'                    ValorConstrucao = 0
'                    Do While Not RsImovel.EOF
'                        Componente = RsImovel!tdi_tco_cod_componente
'                        If Componente = CodParticular Then Exit Do
'                        QueryIptu = "Select Tab_Componente.tco_Valor*Tab_Detalhe_Imovel.tdi_valor_item " & _
'                        " FROM Tab_Componente, Tab_Detalhe_Imovel " & _
'                        " WHERE Tab_Componente.tco_cod_componente=" & RsImovel(0) & " and Tab_Componente.tco_tmu_cod_municipio =" & AplicacoesVTFuncoes.Codigo_Municipio & _
'                        " and Tab_Detalhe_Imovel.tdi_tco_cod_componente IN (SELECT DISTINCT(tco_cod_componente_fator_calculo) " & _
'                        " FROM Tab_Componente WHERE tco_cod_componente = " & RsImovel(0) & " and tco_tmu_cod_municipio =" & AplicacoesVTFuncoes.Codigo_Municipio & _
'                        ") and Tab_Detalhe_Imovel.tdi_tim_ic = '" & Contribuintes!tim_ic & "'"
'                        If Bdados.AbreTabela(QueryIptu, RsIptu) Then
'                            ValorConstrucao = ValorConstrucao + RsIptu(0)
'                            If Componente = 21 Or Componente = 22 Or Componente = 24 Or Componente = 2 Then
'                                Taxas = Taxas + 1
'                            End If
'                            If Componente = 111 Or Componente = 112 Or Componente = 113 Then
'                                Testada = Testada + RsIptu(0)
'                            End If
'                            If Componente = 108 Then 'Area construida
'                                Sql = " SELECT tdi_valor_item FROM Tab_Detalhe_Imovel WHERE TDI_TIM_IC = '" & Contribuintes!tim_ic & "' AND tdi_tco_cod_componente = 108"
'                                If Bdados.AbreTabela(Sql, Rs) Then
'                                    AreaConstruida = Rs!tdi_valor_item
'                                End If
'                                Rs.Fechar
'                            ElseIf Componente = 115 Then 'Pavimentos
'                                TotalPavimentos = RsImovel(1)
'                                Sql = "Select * from tab_componente where tco_cod_componente = 115"
'                                If Bdados.AbreTabela(Sql, RsPav) Then
'                                    ValorPorPavimento = RsPav!tco_valor
'                                End If
'                            ElseIf Componente = 103 Then 'Area Terreno
'                                AreaTerreno = RsImovel(1)
'                            ElseIf Componente = 59 Then
'                                Gleba = True
'                            ElseIf Componente >= 64 And Componente <= 66 Then
'                                EstadoConservacao = RsIptu(0)
'                            End If
'                        End If
'                        Bdados.FechaTabela RsIptu
'                        RsImovel.MoveNext
'                    Loop
'                    If Componente <> CodParticular Then
'                        ValorTerreno = CalculaValorTerreno(Contribuintes!tim_ic, Terreno)
'                        Total = ValorConstrucao + ValorTerreno
'                        Taxas = Testada * ValorPadraoTsu * Taxas
'                        If Gleba = True Then
'                            Sql = "select  tvl_valor from Tab_Valor_Terreno where tvl_tlg_cod_logradouro='" & BuscaCodigo("SELECT TIM_TLG_COD_LOGRADOURO FROM TAB_IMOVEL WHERE TIM_IC ='" & Contribuintes!tim_ic & "'") & "'"
'                            If Bdados.AbreTabela(Sql, Rs) Then
'                                Total = Total * Rs(0)
'                            End If
'                        End If
'                        Bdados.FechaTabela Rs
'
'                        ValorImovel = Total
'                        If TotalPavimentos >= 1 Then
'                            ValorPorPavimento = (ValorPorPavimento * (TotalPavimentos - 1)) * ValorImovel
'                        End If
'                        Total = Total + ValorPorPavimento
'                        ValorImovel = Total * IIf(EstadoConservacao <= 0, 1, EstadoConservacao)
'                        If AplicacoesVTFuncoes.Codigo_Municipio = 1 Then
'                            Total = ((ValorImovel * Reducao) * Aliquota) ' + Taxas
'                        Else
'                            Total = ((ValorImovel * Reducao) * IIf(AreaConstruida <= 0, AliquotaTerr, AliquotaEdif)) '+ Taxas
'                        End If
'                        TotalUnico = Total * ((100 - Desconto) / 100)
''INJUSTICA - CPI
'                        Sql = "Select TGE_NOME from TAB_GERAL WHERE TGE_TIPO =720 AND TGE_CODIGO > 0 "
'                        If Bdados.AbreTabela(Sql) Then
'                            If AreaTerreno > Bdados.Tabela(0) Then
'                                Total = ((ValorTerreno * EstadoConservacao * Reducao) * IIf(AreaConstruida <= 0, AliquotaTerr, AliquotaEdif)) '+ Taxas
'                                TotalUnico = Total
'                                ValorImovel = ValorTerreno
'                            End If
'                        End If
'                        If Total < ImpostoMinimo Then Total = ImpostoMinimo
''FIM DA INJUSTICA
'                        CodPagamento = Conta.GeraCodPagamento(CodImposto)
'                        CodTaxa = "11220100"
'                        DataVencimento = "20/06/2002"
'                        Conta.GeraPagamento Contribuintes!tim_tci_im, Contribuintes!tim_ic, CodImposto, i, DataVencimento, Total, 0, 0, CodPagamento, 0, 0, Taxas, CodTaxa
'
'                        Valores = Bdados.PreparaValor(Bdados.Converte(ValorConstrucao, TCDuplo), Bdados.Converte(ValorTerreno, TCDuplo), Bdados.Converte(ValorConstrucao + ValorConstrucao, TCDuplo))
'                        Call Bdados.AtualizaDados("Tab_IMOVEL", Valores, "TIM_VALOR_EDIFIC,TIM_VALOR_TERRENO,tim_valor", "tim_ic='" & Contribuintes!tim_ic & "'")
'
'                        Reg = Reg + 1
'                        CriaIptu_Novo = Reg
'                        DoEvents
'                    End If
'                End If
'                Total = 0
'                Taxas = 0
'                ValorPorPavimento = 0
'                TotalPavimentos = 0
'                ValorImovel = 0
'                ValorConstrucao = 0
'                ValorTerreno = 0
'                AreaConstruida = 0
'                Componente = 0
'                Contribuintes.MoveNext
'                DoEvents
'            Loop
'        End If
'        DoEvents
'        DoEvents
'        Conta.DefineIsentos CodImposto, i
'    Next
'    Bdados.FechaTabela RsImovel
'    Bdados.FechaTabela RsIptu
'    Screen.MousePointer = 0
'End Function

Public Function CriaIptu(Contribuintes As VSRecordset, PeriodoInicial As Integer, PeriodoFinal As Integer, Optional Contador As Object) As Double
    Dim QueryIptu As String
    Dim RsImovel As VSRecordset
    Dim RsIptu As VSRecordset
    Dim ValorCalculado As Double
    Dim Valores As String
    Dim Campos As String
    Dim DataVencimento As String
    Dim I As Long
    Dim CodImposto As String
    Dim RefComp As Integer
    Dim Total As Double
    Dim Taxas As Double
    Dim Reducao As Double
    Dim Reg As Double
    Dim CodConta  As Double
    Dim Aliquota As Double
    Dim Componente As Integer
    Dim AreaConstruida As Double
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim ValorImovel As Double
    Dim TotalPavimentos As Integer
    Dim ValorPorPavimento As Double
    Dim RsPav As VSRecordset
    Dim MaxCotas As Byte
    Dim RsDesconto As VSRecordset
    Dim TotalUnico As Double
    Dim ImpostoMinimo As Double
    Dim Desconto As Double
    Dim Novo As Boolean
    Dim ValorTerreno As Double
    Dim CodPagamento As String
    Dim CodParticular As Integer
    Dim AliquotaEdif As Double
    Dim AliquotaTerr As Double
    Dim ValorPadraoTsu As Double
    Dim ValorConstrucao As Double
    Dim Testda As Double
    Dim EstadoConservacao As Double
    Dim AreaTerreno As Double
    Dim Condicao As String
    Dim CodTaxa  As String
    Dim Adicional As Integer
    Dim RsAux1 As VSRecordset
    Dim RsAux2 As VSRecordset
    Dim ValorVenal As Double
    Dim Indice100 As Integer
    Dim AreaPadaraoGleba As Double
    Dim ValorMinimoImpostoBairro As Double
    Dim PontoComercial As Boolean
    Dim Obrig As New Obrigacao
    Dim ValorFator As Double
    Dim ValorTemp As Double
    Dim ValorM2 As Double
    Dim CodLogr As String
    Dim Iptu As New VSIptu
    Dim LimiteIsencaoAreaTotal As Double
    
    On Error GoTo Trata
    Screen.MousePointer = 11
    CodImposto = BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
    Dim Gleba As Boolean
    Reg = 0
    ValorPadraoTsu = Nvl(Temp.PegaParametro(Bdados, "VALOR TSU 2005"), 0)
    ValorMinimoImpostoBairro = BuscaValorMinimoImpostoBairro(Contribuintes!tim_tba_cod_bairro)
    Gleba = False
    Componente = 0
    AreaConstruida = 0
    ValorConstrucao = 0
    ValorTerreno = 0
    AliquotaEdif = CDbl(Nvl(Iptu.BuscaParamentroIptu("ALIQUOTAS", 2), 1))
    AliquotaTerr = CDbl(Nvl(Iptu.BuscaParamentroIptu("ALIQUOTAS", 1), 2))
    Desconto = CDbl(Nvl(Iptu.BuscaParamentroIptu("DESCONTO COTA ÚNICA", 1), 0))
    LimiteIsencaoAreaTotal = CDbl(Nvl(Iptu.BuscaParamentroIptu("LIMITE ISENCAO AREA TOTAL", 1), 0))
    DataVencimento = Iptu.BuscaParamentroIptu("VENCIMENTO COTAS", 1)

'    Sql = "Select TGE_NOME from TAB_GERAL WHERE TGE_TIPO =(select tge_tipo " & _
'        "from tab_geral where tge_nome = 'AREA GLEBA') AND TGE_CODIGO > 0 "
'    If Bdados.AbreTabela(Sql, Rs) Then AreaPadaraoGleba = Rs!TGE_NOME
'    Bdados.FechaTabela Rs

    For I = PeriodoInicial To PeriodoFinal
        
        Reducao = BuscaReducao(CodImposto, CStr(I))

        If IsDate(DataVencimento) Then
            Contribuintes.MoveFirst
            Do While Not Contribuintes.EOF
                Contador = "Calculando IPTU da IC:  " & Contribuintes!tim_ic
                DoEvents
                Sql = " SELECT tdi_tco_cod_componente,tdi_valor_item FROM Tab_Detalhe_Imovel WHERE " & _
                    " TDI_TIM_IC = '" & Contribuintes!tim_ic & "' AND tdi_tco_cod_componente in (100,108,112) order by tdi_tco_cod_componente"
                
                If Bdados.AbreTabela(Sql, Rs) Then
                    Rs.MoveFirst
                    Do
                        If Rs!tdi_tco_cod_componente = 100 Then Testda = Rs!tdi_valor_item
                        If Rs!tdi_tco_cod_componente = 108 Then AreaTerreno = Rs!tdi_valor_item
                        If Rs!tdi_tco_cod_componente = 112 Then AreaConstruida = Rs!tdi_valor_item
                        Rs.MoveNext
                    Loop While Not Rs.EOF
                End If
                If AreaConstruida >= LimiteIsencaoAreaTotal Then
                    QueryIptu = "SELECT tdi_tgc_cod_grupo,tdi_tco_cod_componente,tdi_valor_item, tco_valor,tco_cod_componente_fator_calc,tgc_categoria FROM Tab_Detalhe_Imovel inner join "
                    QueryIptu = QueryIptu & " tab_componente_avancado on tdi_tco_cod_componente = tco_cod_componente and tdi_tgc_cod_grupo = tco_grupo "
                    QueryIptu = QueryIptu & " inner join tab_grupo_componente_avancado on tco_grupo = tgc_cod_grupo "
                    QueryIptu = QueryIptu & " where tdi_tim_ic ='" & Contribuintes!tim_ic & "' order by 1 asc"
                    Gleba = False
                    If Bdados.AbreTabela(QueryIptu, RsImovel) Then
                        RsImovel.MoveFirst
                        Taxas = 0
                        ValorConstrucao = 0
                        ValorTerreno = 0
                        ValorImovel = 0
                        Do While Not RsImovel.EOF
                            Componente = RsImovel!tdi_tco_cod_componente
                            Select Case RsImovel!tco_cod_componente_fator_calc
                                Case 100
                                    ValorFator = Testda
                                Case 108
                                    ValorFator = AreaTerreno
                                Case 112
                                    ValorFator = AreaConstruida
                                Case 1000
                                    ValorFator = 1
                            End Select
                            If CDbl(RsImovel!tdi_tco_cod_componente) = CDbl(RsImovel!tdi_valor_item) Then
                                ValorTemp = ValorFator * CDbl(RsImovel!tco_valor)
                            Else
                                ValorTemp = ValorFator * CDbl(RsImovel!tdi_valor_item)
                            End If
                            If RsImovel!tgc_categoria = 1 Then 'Dados do Terreno
                               ValorTerreno = Format(ValorTerreno + ValorTemp, Const_Monetario)
                            ElseIf RsImovel!tgc_categoria = 2 Then 'Dados da Construcao
                                ValorConstrucao = Format(ValorConstrucao + ValorTemp, Const_Monetario)
                            End If
                            RsImovel.MoveNext
                        Loop
                    End If
                    ValorM2 = BuscaValorM2(Contribuintes!tim_tlg_cod_logradouro)
                    ValorTerreno = ValorTerreno * ValorM2
                    ValorImovel = Format(ValorTerreno + ValorConstrucao, Const_Monetario)
                    If AreaConstruida = 0 Then
                        Aliquota = AliquotaTerr / 100
                    Else
                        Aliquota = AliquotaEdif / 100
                    End If
                    TotalUnico = ValorImovel * Aliquota
                    CodPagamento = Conta.NumPagamento(Trim(Contribuintes!tim_ic), I, CodImposto, Trim(Contribuintes!tim_ic), 0, False, 0)
                    Obrig.CriaObrigacao CodImposto, I, I, Trim(Contribuintes!tim_ic), Format(TotalUnico, Const_Monetario), , , DataVencimento, , CodPagamento
                    Valores = Bdados.PreparaValor(Bdados.Converte(ValorConstrucao, TCDuplo), Bdados.Converte(ValorTerreno, TCDuplo), Bdados.Converte(CDbl(ValorImovel), TCDuplo))
                    Call Bdados.AtualizaDados("Tab_IMOVEL", Valores, "TIM_VALOR_EDIFIC,TIM_VALOR_TERRENO,tim_valor", "tim_ic='" & Contribuintes!tim_ic & "'")
                    Reg = Reg + 1
                    If Not Contador Is Nothing Then Contador = Reg
                    CriaIptu = Total
                End If
                DoEvents
                Contribuintes.MoveNext
            Loop
        End If
        Total = 0
        TotalUnico = 0
        Taxas = 0
        ValorPorPavimento = 0
        TotalPavimentos = 0
        ValorImovel = 0
        ValorConstrucao = 0
        ValorTerreno = 0
        EstadoConservacao = 0
        AreaConstruida = 0
        ValorVenal = 0
        Componente = 0
        Contribuintes.MoveNext
        DoEvents
    Next
    Bdados.FechaTabela RsImovel
    Bdados.FechaTabela RsIptu
    Screen.MousePointer = 0
    Exit Function
Trata:
    If Err.Number = 3021 Then
        Exit Function
    Else
        Avisa Err.Description
        Exit Function
        Resume
    End If
End Function

Public Function CriaIptu_Old(Contribuintes As VSRecordset, PeriodoInicial As Integer, PeriodoFinal As Integer, Optional Contador As Object) As Double
    Dim QueryIptu As String
    Dim RsImovel As VSRecordset
    Dim RsIptu As VSRecordset
    Dim ValorCalculado As Double
    Dim Valores As String
    Dim Campos As String
    Dim DataVencimento As String
    Dim I As Long
    Dim CodImposto As String
    Dim RefComp As Integer
    Dim Total As Double
    Dim Taxas As Double
    Dim Reducao As Double
    Dim Reg As Double
    Dim CodConta  As Double
    Dim Aliquota As Double
    Dim Componente As Integer
    Dim AreaConstruida As Double
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim ValorImovel As Double
    Dim TotalPavimentos As Integer
    Dim ValorPorPavimento As Double
    Dim RsPav As VSRecordset
    Dim MaxCotas As Byte
    Dim RsDesconto As VSRecordset
    Dim TotalUnico As Double
    Dim ImpostoMinimo As Double
    Dim Desconto As Double
    Dim Novo As Boolean
    Dim ValorTerreno As Double
    Dim CodPagamento As String
    Dim CodParticular As Integer
    Dim AliquotaEdif As Double
    Dim AliquotaTerr As Double
    Dim ValorPadraoTsu As Double
    Dim ValorConstrucao As Double
    Dim Testda As Double
    Dim EstadoConservacao As Double
    Dim AreaTerreno As Double
    Dim Condicao As String
    Dim CodTaxa  As String
    Dim Adicional As Integer
    Dim RsAux1 As VSRecordset
    Dim RsAux2 As VSRecordset
    Dim ValorVenal As Double
    Dim Indice100 As Integer
    Dim AreaPadaraoGleba As Double
    Dim ValorMinimoImpostoBairro As Double
    Dim PontoComercial As Boolean
    On Error GoTo Trata
    Screen.MousePointer = 11
    CodImposto = BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
    Dim Gleba As Boolean
    Reg = 0
    ValorPadraoTsu = Nvl(Temp.PegaParametro(Bdados, "VALOR TSU 2005"), 0)
    ValorMinimoImpostoBairro = BuscaValorMinimoImpostoBairro(Contribuintes!tim_tba_cod_bairro)
    Gleba = False
    Componente = 0
    AreaConstruida = 0
    ValorConstrucao = 0
    ValorTerreno = 0
    CodParticular = IIf(AplicacoesVTFuncoes.Municipio = "COLINAS", 80, 81)
    AliquotaEdif = CDbl(Nvl(BuscaNaGeral("ALIQUOTA IPTU", 1), 0))
    AliquotaTerr = CDbl(Nvl(BuscaNaGeral("ALIQUOTA IPTU", 2), 0))

    'Pega Maximo de cotas em parcelamento
    Sql = "SELECT * from Tab_Parametro_Parcelamento"
    If Bdados.AbreTabela(Sql, Rs) Then
        MaxCotas = Rs!tpp_max_cotas
    End If

    Sql = "Select TGE_NOME from TAB_GERAL WHERE TGE_TIPO =(select tge_tipo " & _
        "from tab_geral where tge_nome = 'AREA GLEBA') AND TGE_CODIGO > 0 "
    If Bdados.AbreTabela(Sql, Rs) Then AreaPadaraoGleba = Rs!TGE_NOME
    Bdados.FechaTabela Rs
    'Pega desconto em cota unica
    Sql = "Select TGE_NOME from tab_geral where TGE_TIPO = 755 and TGE_CODIGO > 0"
    If Bdados.AbreTabela(Sql, RsDesconto) Then
        If IsNumeric(RsDesconto!TGE_NOME) Then TotalUnico = TotalUnico * ((100 - RsDesconto!TGE_NOME) / 100)
    End If
    Bdados.FechaTabela Rs

    CarregaComponetes

    For I = PeriodoInicial To PeriodoFinal
        DataVencimento = Imposto.BuscaDataVencimento(CodImposto, Imposto.BuscaAnoImposto(CodImposto, CStr(I)))
        
'        DataVencimento =BuscaDataVencimento(BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)), CInt(I))

        Reducao = BuscaReducao(CodImposto, CStr(I))
        Aliquota = BuscaAliquota(CodImposto, CStr(I))
        'Pega valor mínimo de imposto
        Sql = "SELECT tpi_valor_min_imposto FROM TAB_PARAMETRO_IMPOSTO WHERE " & _
                "tpi_tip_cod_imposto='" & CodImposto & "'" ' AND tpi_ano_imposto='" & _
                IIf(Len(CStr(i)) = 4, Right(i, 2), Mid(i, 4, 2)) & "'"
        Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(CodImposto, Left(I, 4)) & "'"
        If Bdados.AbreTabela(Sql & Condicao, Rs) Then
            ImpostoMinimo = Rs(0)
        End If
        Bdados.FechaTabela Rs
        If IsDate(DataVencimento) Then
            Contribuintes.MoveFirst
            Do While Not Contribuintes.EOF
                Sql = " SELECT tdi_valor_item FROM Tab_Detalhe_Imovel WHERE " & _
                    " TDI_TIM_IC = '" & Contribuintes!tim_ic & "' AND tdi_tco_cod_componente =113"
                If Bdados.AbreTabela(Sql, Rs) Then
                    AreaConstruida = Rs(0)
                End If
                If AreaConstruida = 0 Then
                    ValorConstrucao = 0
                    PontoComercial = False
                    Sql = " SELECT tdi_tco_cod_componente ,tdi_valor_item FROM Tab_Detalhe_Imovel WHERE " & _
                    " TDI_TIM_IC = '" & Contribuintes!tim_ic & "' AND (tdi_tco_cod_componente =108 OR tdi_tco_cod_componente  = 100) ORDER BY tdi_tco_cod_componente "
                    If Bdados.AbreTabela(Sql, Rs) Then
                        Rs.MoveFirst
                        If Rs!tdi_tco_cod_componente = 100 Then
                            Testda = Rs(1)
                            Rs.MoveNext
                            If Rs!tdi_tco_cod_componente = 108 Then AreaTerreno = Rs(1)
                        ElseIf Rs!tdi_tco_cod_componente = 108 Then
                            AreaTerreno = Rs(1)
                        End If
                    End If
                Else
                    QueryIptu = "SELECT tdi_tgc_cod_grupo,tdi_tco_cod_componente,tdi_valor_item FROM Tab_Detalhe_Imovel WHERE " & _
                            "tdi_tim_ic ='" & Contribuintes!tim_ic & "' order by Tdi_tco_cod_componente asc"
                    Gleba = False
                    If Bdados.AbreTabela(QueryIptu, RsImovel) Then
                        RsImovel.MoveFirst
                        Taxas = 0
                        ValorConstrucao = 0
                        Do While Not RsImovel.EOF
                            Componente = RsImovel!tdi_tco_cod_componente
                            If RsImovel!tdi_tgc_cod_grupo > UBound(Componentes) Then Componente = 0
                            If Componentes(Componente).ComponenteFator = 1000 Then
                                'VALO FIXO
                                If Componente < 100 Then ValorConstrucao = ValorConstrucao + Componentes(Componente).Valor
                                If Componente = 100 Then '101 Or Componente = 103 Or Componente = 105 Then
                                        Testda = RsImovel(1)
                                End If
                                If Componente = 110 Then
                                    'Pavimentos
                                    TotalPavimentos = RsImovel(1)
                                ElseIf Componente = 108 Then
                                    'Area Terreno
                                    AreaTerreno = RsImovel(1)
                                End If
                            Else 'PRECISA SER CALCULADO
                                QueryIptu = " SELECT tdi_valor_item * " & Bdados.Converte(Componentes(Componente).Valor, TCDuplo) & "  FROM " & _
                                        " Tab_Detalhe_Imovel " & _
                                " WHERE tab_detalhe_imovel.tdi_tgc_cod_grupo = " & _
                                Componentes(Componente).ComponenteFator & _
                                " and Tab_Detalhe_Imovel.tdi_tim_ic = '" & Contribuintes!tim_ic & "'"
                                If Bdados.AbreTabela(QueryIptu, RsIptu) Then
                                    If Componente < 100 Then ValorConstrucao = ValorConstrucao + RsIptu(0)
                                    If AplicacoesVTFuncoes.Codigo_Municipio <> 1 And (Componente >= 200 And Componente <= 300) Then
                                        Adicional = 100
                                    End If
                                    If Componente = 100 + Adicional Then
                                        Testda = Testda + RsIptu(0)
                                    End If
                                    If Componente = 110 + Adicional Then
                                        'Pavimentos
                                        TotalPavimentos = RsImovel(1)
                                    ElseIf Componente = 108 + Adicional Then
                                        'Area Terreno
                                        AreaTerreno = RsImovel(1)
                                    ElseIf Componente = 59 Then
                                        Gleba = True
                                    ElseIf Componente >= 64 And Componente <= 66 Then
                                        EstadoConservacao = RsIptu(0)
                                    ElseIf Componente = 37 Then
                                        PontoComercial = True
                                    End If
                                End If
                            End If
                            Adicional = 0
                            Bdados.FechaTabela RsIptu
                            RsImovel.MoveNext
                        Loop
                    End If
                End If
                If Componente <> CodParticular Then
                '                        Call CalculaTaxasIptu(Contribuintes!TIM_ic, Contribuintes!tim_tci_im, CodImposto,I, Taxas)
                    ValorTerreno = CalculaValorTerreno(Contribuintes!tim_ic, Terreno + Adicional)
                   If Contribuintes!tim_tba_cod_bairro <> Temp.PegaParametro(Bdados, "BAIRRO ISENTO TAXA") Then If I >= 2005 Then Taxas = Testda * ValorPadraoTsu
                    If Gleba = True Then
                        Sql = "select  tvl_valor from Tab_Valor_Terreno where tvl_tlg_cod_logradouro='" & _
                        BuscaCodigo("SELECT TIM_TLG_COD_LOGRADOURO FROM TAB_IMOVEL WHERE TIM_IC ='" & Contribuintes!tim_ic & "'") & "'"
                        If Bdados.AbreTabela(Sql, Rs) Then
                            Total = Total * Rs(0)
                        End If
                    End If
                    Bdados.FechaTabela Rs
                    ValorVenal = CDbl(ValorConstrucao + ValorTerreno) * Reducao
                    If TotalPavimentos >= 1 Then
                        ValorPorPavimento = (ValorPorPavimento * (TotalPavimentos - 1)) * ValorVenal
                    End If
                    Aliquota = CDbl(IIf(AreaConstruida = 0, AliquotaTerr, AliquotaEdif)) / 100
                    Total = Format((ValorVenal * Aliquota) + Taxas, Const_Monetario)
                    TotalUnico = Format(Total * ((100 - Desconto) / 100), Const_Monetario)
                    If AreaTerreno >= AreaPadaraoGleba And Not PontoComercial Then
                        Taxas = Testda * ValorPadraoTsu / 3
                        ValorTerreno = AreaTerreno * CDbl(Nvl(Temp.PegaParametro(Bdados, "M2 GLEBA"), 0.1))
                        Total = (IIf(ValorConstrucao < ValorTerreno, ValorConstrucao, ValorTerreno) * Aliquota)
                        If Total < ValorMinimoImpostoBairro Then Total = ValorMinimoImpostoBairro
                        Total = Total '+ Taxas
                        TotalUnico = Total
                        ValorVenal = ValorTerreno
                    Else
                        ValorTaxas = Taxas
                        If Total < ValorMinimoImpostoBairro Then Total = ValorMinimoImpostoBairro
                        Total = Total + Taxas
                    End If
'                    CodPagamento = Conta.NumPagamento(Trim(Contribuintes!tim_tci_im), i, CodImposto, Trim(Contribuintes!tim_ic), 0, False, 0)
'                    Conta.GeraPagamento Trim(Contribuintes!tim_tci_im), Trim(Contribuintes!tim_ic), CodImposto, i, DataVencimento, Format(TotalUnico, Const_Monetario), 0, 0, CodPagamento, 0, 0, Taxas, CodTaxa
                    Valores = Bdados.PreparaValor(Bdados.Converte(ValorConstrucao * Reducao, TCDuplo), Bdados.Converte(ValorTerreno * Reducao, TCDuplo), Bdados.Converte(CDbl(ValorTerreno + ValorConstrucao) * Reducao, TCDuplo))
                    Call Bdados.AtualizaDados("Tab_IMOVEL", Valores, "TIM_VALOR_EDIFIC,TIM_VALOR_TERRENO,tim_valor", "tim_ic='" & Contribuintes!tim_ic & "'")
                    Reg = Reg + 1
                    If Not Contador Is Nothing Then Contador = Reg
                    CriaIptu_Old = Total
                    DoEvents
                End If
                Contribuintes.MoveNext
            Loop
        End If
        Total = 0
        TotalUnico = 0
        Taxas = 0
        ValorPorPavimento = 0
        TotalPavimentos = 0
        ValorImovel = 0
        ValorConstrucao = 0
        ValorTerreno = 0
        EstadoConservacao = 0
        AreaConstruida = 0
        ValorVenal = 0
        Componente = 0
        Contribuintes.MoveNext
        DoEvents
    Next
    Bdados.FechaTabela RsImovel
    Bdados.FechaTabela RsIptu
    Screen.MousePointer = 0
    Exit Function
Trata:
    If Err.Number = 3021 Then
        Exit Function
    Else
        Avisa Err.Description
        Exit Function
        Resume
    End If
End Function

Public Function BuscaReducao(Imposto As String, Ano As String)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Condicao As String
    Sql = "SELECT tpi_reducao FROM TAB_PARAMETRO_IMPOSTO where tpi_tip_cod_imposto='" & Imposto & "'"
    Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Ano) & "'"
    If Bdados.AbreTabela(Sql & Condicao, Rs) Then
        BuscaReducao = (100 - Rs(0)) / 100
    End If
    If BuscaReducao = 0 Then BuscaReducao = 1
    Bdados.FechaTabela Rs
End Function

Public Function BuscaAliquota(Imposto As String, Ano As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Condicao As String
    Dim Valor As String
    
    Sql = "SELECT tpi_aliquota FROM TAB_PARAMETRO_IMPOSTO where tpi_tip_cod_imposto='" & Imposto & "'"
    Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Ano) & "'"
    If Not Bdados.AbreTabela(Sql & Condicao, Rs) Then
        BuscaAliquota = 0
        Exit Function
    Else
        Valor = Rs(0) / 100
        BuscaAliquota = Valor
    End If
End Function

Private Sub AtualizaValorVenal(Imovel As String, Total As Double, Optional Unidade As Integer)
    Dim Valores As String
    Valores = Bdados.PreparaValor(Bdados.Converte(Total, TCDuplo))
    Call Bdados.AtualizaDados("TAB_IMOVEL", Valores, "tim_valor", "tim_ic='" & Imovel & "'")
End Sub

Private Sub AtualizaValorVenalBCI(Imovel As String, Total As Double)
    Dim Valores As String
    Valores = Bdados.PreparaValor(Bdados.Converte(Total, TCDuplo))
    Call Bdados.AtualizaDados("Tab_IMOVEL", Valores, "tim_valor", "tim_ic='" & Imovel & "'")
End Sub

Public Function BuscaNomeCAE(Atividade As String) As String

    Dim RsAtiv As VSRecordset
    Dim Sql As String
    
    Sql = "Select tae_nome from tab_atividade_economica where tae_cae = " & Val(Atividade)
    If Bdados.AbreTabela(Sql, RsAtiv) Then
        BuscaNomeCAE = RsAtiv(0)
    End If
    Bdados.FechaTabela RsAtiv
End Function

Public Function BuscaCAE(Atividade As String) As String

    Dim RsAtiv As VSRecordset
    Dim Sql As String
    
    Sql = "Select tae_cae from tab_atividade_economica where tae_nome = '" & Atividade & "'"
    If Bdados.AbreTabela(Sql, RsAtiv) Then
        BuscaCAE = RsAtiv(0)
    End If
    Bdados.FechaTabela RsAtiv
End Function

Public Function BuscaCodMunicipio(NomeMunicipio As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "Select tmu_cod_municipio FROM Tab_Municipio WHERE tmu_nome = '" & NomeMunicipio & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaCodMunicipio = Rs(0)
    End If
    Bdados.FechaTabela Rs
End Function



Sub GeraIptu_valorterreno(Cidade As CidadeIptu, Contribuintes As VSRecordset, PeriodoInicial As Integer, PeriodoFinal As Integer, GeraParcelas As Boolean)
    Dim ValorVenalEdificacao As Double
    Dim ValorVenalTerreno As Double
    Dim ValorVenalImovel As Double
    Dim AreaConstruida As Double
    Dim AreaTerreno As Double
    Dim CoeficienteFracaoIdeal As Double
    Dim CoeficienteTerra As Double
    Dim CoeficienteEdific As Double
    Dim ValorTotalEdificacoes As Double
    Dim AliquotaTerreno As Double
    Dim AliquotaPredio As Double
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim TipoImovel As Byte
    Dim ValorImposto As Double
    Dim CodImposto As String
    Dim ValorTaxas As Double
    Dim I As Long
    Dim j As Byte
    Dim Fim As Boolean
    Dim Conta As New ContaCorrente
    Dim DataVenc As String
    Dim Reg  As Double
    Dim InscImovel As String
    Dim ImovelAtual As String
    Dim Zona As Byte
    Dim MinImposto As Double
    Dim Velhas As Double
    Const DataVencimento = "30/10/2001"
    Select Case Cidade
            Case cip_Balsas
                    Contribuintes.MoveFirst
                    CodImposto = BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
                    MinImposto = BuscaValorMinimoImposto(CodImposto, Year(Date))
                    InscImovel = ""
                    Fim = False
                    Velhas = 0
                    Do While Not Contribuintes.EOF 'PARA CADA IMOVEL SELECIONADO
                        TipoImovel = Contribuintes!tim_tipo_imovel
                        If InscImovel <> Contribuintes!tdi_tim_ic & Format(Contribuintes!tdi_tim_ic_unidade, "000") Then
                                    Zona = Contribuintes!TIM_ZONA
                                    AliquotaTerreno = 0
                                    AliquotaPredio = 0
                                    CoeficienteEdific = 0
                                    CoeficienteTerra = 0
                                    ValorTaxas = 0
                                    AreaTerreno = 0
                                    AreaConstruida = 0
                                    ValorVenalTerreno = 0
                                    ValorVenalEdificacao = 0
                                    ValorTotalEdificacoes = 0
                                    'CoeficienteFracaoIdeal = 0
                                   If TipoImovel = 2 Then 'SE APENAS TERRENO
                                           AliquotaTerreno = BuscaAliquotaImovelIPTU(tim_TerrenoSemEdif, Nvl("" & Zona, 5))
                                           CoeficienteTerra = BuscaCoeficienteImovelIPTU(tzo_ZonaUrbana, Nvl("" & Zona, 5))
                                   Else 'SE TERRENO COM PREDIO
                                           AliquotaTerreno = BuscaAliquotaImovelIPTU(tim_TerrenoComEdif, Nvl("" & Zona, 5))
                                           AliquotaPredio = BuscaAliquotaImovelIPTU(tim_Edificacao, Nvl("" & Zona, 5))
                                           CoeficienteEdific = BuscaCoeficienteEdificacao(Contribuintes!tdi_tim_ic, Contribuintes!tdi_tim_ic_unidade)
                                           CoeficienteTerra = BuscaCoeficienteImovelIPTU(tzo_ZonaUrbana, Nvl("" & Zona, 5))
                                   End If
                                   'CALCULANDO IPTU
                                   AreaTerreno = BuscaValorArea(Contribuintes!tdi_tim_ic, Contribuintes!tdi_tim_ic_unidade, 108)
                                   AreaConstruida = BuscaValorArea(Contribuintes!tdi_tim_ic, Contribuintes!tdi_tim_ic_unidade, 112)
                                   ValorVenalTerreno = AreaTerreno * CoeficienteTerra
                                   ValorVenalEdificacao = AreaConstruida * CoeficienteEdific
                                   
                                   If Contribuintes!tdi_tim_ic_unidade < 200 Then 'SE NAO FOR CONDOMINIO
                                        ImovelAtual = Contribuintes!tdi_tim_ic
                                        Do
                                            If ImovelAtual <> Nvl("" & Contribuintes!tdi_tim_ic, 0) Then 'PARA CADA EDIFICACAO DO IMOVEL
                                                Contribuintes.MovePrevious
                                                Exit Do
                                            End If
                                             AreaConstruida = BuscaValorArea(Contribuintes!tdi_tim_ic, Contribuintes!tdi_tim_ic_unidade, 112)
                                             CoeficienteEdific = BuscaCoeficienteEdificacao(Contribuintes!tdi_tim_ic, Contribuintes!tdi_tim_ic_unidade)
                                             ValorVenalEdificacao = AreaConstruida * CoeficienteEdific
                                             ValorTotalEdificacoes = ValorTotalEdificacoes + ValorVenalEdificacao 'ACUMULA AS EDIFICACOES
                                             Contribuintes.MoveNext
                                        Loop While Not Contribuintes.EOF
                                        ValorVenalEdificacao = ValorTotalEdificacoes
                                    Else 'SE FOR CONDOMINIO
                                        If Contribuintes!tdi_tim_ic_unidade = 200 Or Contribuintes!tdi_tim_ic_unidade = 600 Then 'BT
                                            CoeficienteFracaoIdeal = BuscaFracaoIdealImposto(Contribuintes!tdi_tim_ic, ValorVenalTerreno)
                                        Else 'PARA CADA LOTE DO CONDOMINIO
                                            AreaConstruida = BuscaValorArea(Contribuintes!tdi_tim_ic, Contribuintes!tdi_tim_ic_unidade, 112)
                                            ValorVenalTerreno = AreaConstruida * CoeficienteFracaoIdeal
                                        End If
                                    End If
                                    If Contribuintes.EOF Then 'FUI ATE O FIM DE EDIFICACOES NORMAIS
                                        Contribuintes.MovePrevious
                                        Fim = True
                                    End If
                                    'NAO GERA PARA O CONDOMINIO EM SI. SO PARA IMOVEIS COMUNS E LOTES DO CONDOMINIO
                                    If Contribuintes!tdi_tim_ic_unidade <> 200 And Contribuintes!tdi_tim_ic_unidade <> 600 Then
                                        ValorVenalImovel = ValorVenalTerreno + ValorVenalEdificacao
                                        Call Bdados.AtualizaDados("TAB_IMOVEL", "TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC", Bdados.PreparaValor(ValorVenalTerreno, ValorVenalEdificacao), "tim_ic ='" & Contribuintes!tim_ic & "' AND tim_unidade =" & Contribuintes!TIM_UNIDADE)
                                    End If
                            'End If
                        ElseIf Contribuintes!tdi_tim_ic_unidade < 200 Then
                            InscImovel = Contribuintes!tdi_tim_ic & Format(Contribuintes!tdi_tim_ic_unidade, "000")
                            Contribuintes.MoveNext
                        End If
                        If Contribuintes!tdi_tim_ic_unidade >= 200 Or Fim Then
                            InscImovel = Contribuintes!tdi_tim_ic & Format(Contribuintes!tdi_tim_ic_unidade, "000")
                        End If
                        Contribuintes.MoveNext
                        DoEvents
                    Loop
                    Reg = Reg - 1
                    Fim = False
                    DoEvents
            Case cip_Colinas
                    
    End Select
End Sub

Function BuscaFracaoIdealImposto(IC As String, ValorTerreno As Double) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    
    'BUSCA TOTAL AREAS CONSTRUIDAS
    Sql = "select sum(tdi_valor_item)  from TAB_DETALHE_IMOVEL where "
    If CInt(Right(Trim(IC), 3)) >= 200 Then
        Sql = Sql & " tdi_tim_ic > '" & Trim(IC) & "' AND tdi_tim_ic <= '" & Left(IC, 12) & "300" & _
            "' and tdi_tgc_cod_grupo=112 "
    Else
        Sql = Sql & " tdi_tim_ic = '" & IC & "' and tdi_tgc_cod_grupo=112 "
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaFracaoIdealImposto = IIf(ValorTerreno = 0, 1, ValorTerreno) / Nvl("" & Rs(0), IIf(ValorTerreno = 0, 1, ValorTerreno))
    End If
    Bdados.FechaTabela Rs
End Function

Function BuscaFracaoIdealArea(IC As String, AreaLote As Double) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    On Error Resume Next
    'BUSCA TOTAL AREAS CONSTRUIDAS
    Sql = "select SUM(tdi_valor_item)  from TAB_DETALHE_IMOVEL where "
    If CInt(Right(IC, 3)) < 200 Then
        Sql = Sql & " tdi_tim_ic='" & IC & "' and tdi_tgc_cod_grupo=112 "
    Else
        Sql = Sql & " tdi_tim_ic > '" & IC & "' AND tdi_tim_ic <= '" & Left(IC, 12) & "300'"
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
            BuscaFracaoIdealArea = AreaLote / IIf("" & Rs(0) = 0, AreaLote, Rs(0))
    End If
    Bdados.FechaTabela Rs
End Function

Public Function BuscaValorArea(Imovel As String, Unidade As Integer, Tipo As Integer) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
        
    Sql = "Select tdi_valor_item from TAB_DETALHE_IMOVEL where tdi_tim_ic ='" & Trim(Imovel) & _
    "' and tdi_tim_ic_unidade=" & Unidade & " and tdi_tco_cod_componente =" & Tipo
    
    ' ÉDERSON, NÃO TENHO CERTEZA QUANTO A UTILIZAÇÃO DA UNIDADE AQUI
    If Unidade > 200 And Tipo = 108 Then
        Unidade = 200
        Imovel = Left(Imovel, Len(Trim(Imovel)) - 3) & Unidade
        Sql = "Select tdi_valor_item from TAB_DETALHE_IMOVEL where tdi_tim_ic ='" & Trim(Imovel) & _
        "' and tdi_tim_ic_unidade=" & Unidade & " and tdi_tco_cod_componente =" & Tipo
    End If
    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaValorArea = Nvl("" & Rs(0), 0)
    End If
    Bdados.FechaTabela Rs
End Function
Function BuscaAliquotaImovelIPTU(Imovel As TipoImovel, Zona As Integer) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Imov As Integer
    Dim Zoneamento As Integer
    
    
    If Imovel = tim_Gleba Or Imovel = tim_Edificacao Then
        Zona = 1
    End If
    Imov = Imovel
    Zoneamento = Zona
    
    BuscaAliquotaImovelIPTU = 0
    Sql = "Select TGE_NOME from TAB_GERAL WHERE TGE_TIPO =709 AND TGE_CODIGO = " & IIf(CInt(Imov + Zoneamento) <> 0, CInt(Imov + Zoneamento), 1)
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaAliquotaImovelIPTU = Nvl("" & Rs(0), 1) / 100
    End If
    Bdados.FechaTabela Rs
End Function

Private Function GeraIptuTemporario(Area As String) As String
    If Area <= 30 Then
        GeraIptuTemporario = "0"
    ElseIf Area <= 50 Then
        GeraIptuTemporario = "10"
    ElseIf Area <= 80 Then
        GeraIptuTemporario = "20"
    ElseIf Area <= 100 Then
        GeraIptuTemporario = "30"
    ElseIf Area <= 150 Then
        GeraIptuTemporario = "40"
    Else
        GeraIptuTemporario = "50"
    End If
End Function

Public Sub GravaComponente(InscCadastral As String, ValorVenal As String, AreaConstruida As String, AreaNaoConstruida As String, Formulario As Object)
    Dim Controle As Control
    Dim Rs As VSRecordset
    Dim CodItem As String
    Dim ValorItem As String
    Dim Valores As String
    Dim Campos As String
    Dim ValorGerado As Single
    Dim Sql As String
    Dim Componente As Double
    Bdados.DeletaDados "TAB_DETALHE_IMOVEL", "TDI_TIM_IC ='" & InscCadastral & "'"
    For Each Controle In Formulario.Controls
        If IsNumeric(Controle.Tag) Then
            If CInt(Controle.Tag) < 100 Then
                CodItem = BuscaCodItem(Controle.Text, Controle.Tag)
                ValorItem = Controle.ListIndex
            Else
'                If AplicacoesVTFuncoes.Codigo_Municipio <> 1 And (Componente <> 1000) Then
'                    CodItem = CDbl(Controle.Tag) + 100
'                Else
'                    CodItem = Controle.Tag
'                End If
                CodItem = Controle.Tag
                If CodItem = 1000 Then
                    ValorItem = 1
                Else
                    ValorItem = Controle.Text
                End If
            End If
            If Trim(CodItem) <> "" Then
                Valores = Bdados.PreparaValor(InscCadastral, 0, Controle.Tag, CodItem, ValorItem)
                Campos = "tdi_tim_ic,tdi_tim_ic_unidade,tdi_tgc_cod_grupo,tdi_tco_cod_componente,tdi_valor_item"
                Call Bdados.InsereDados("Tab_Detalhe_Imovel", Valores, Campos)
            End If
        End If
    Next
    
End Sub

Public Function BuscaCodItem(NomeItem As String, Grupo As String) As String
    Dim Sql As String
    Dim RsItem As VSRecordset
    Sql = "SELECT tco_cod_componente FROM tab_componente_avancado WHERE tco_descricao_componente = '" & NomeItem & "' and tco_tmu_cod_municipio=" & AplicacoesVTFuncoes.Codigo_Municipio
    Sql = Sql & " and tco_grupo = " & Grupo
    If Bdados.AbreTabela(Sql, RsItem) Then
        BuscaCodItem = RsItem(0)
    End If
    Bdados.FechaTabela RsItem
End Function

Public Function BuscaCodItemAvancado(NomeItem As String, Grupo As String) As String
    Dim Sql As String
    Dim RsItem As VSRecordset
    Sql = "SELECT tco_cod_componente FROM tab_componente_avancado WHERE tco_descricao_componente = '" & NomeItem & "'"
    Sql = Sql & " and tco_grupo = " & Grupo
    If Bdados.AbreTabela(Sql, RsItem) Then
        BuscaCodItemAvancado = RsItem(0)
    End If
    Bdados.FechaTabela RsItem
End Function

Public Function GeraValorItem(Item As String, ValorVenal As String, AreaContruida As String, AreaNaoContruida As String) As String
    Dim RsAux As VSRecordset
    Dim RsPrin As VSRecordset
    Dim Sql As String
    Dim Valor As String
    Sql = "Select tco_valor, tco_unid_moneta, tco_base_calculo FROM tab_componente_avancado " & _
        " WHERE tco_cod_componente = " & Item & " and tco_tmu_cod_municipio =" & AplicacoesVTFuncoes.Codigo_Municipio
    If Bdados.AbreTabela(Sql, RsPrin) Then
        Select Case RsPrin(1) 'Moeda -> 0 - Real, 1 - Ufir, 2 - Ufm, 3 - Perc
            Case 0
                Valor = IIf(Item = 108, GeraIptuTemporario(CStr(CDbl(AreaContruida) + CDbl(AreaNaoContruida))), CDbl(RsPrin(0)))
            Case 1, 2
                Sql = "Select tmo_valor FROM Tab_Monetaria " & _
                    " WHERE tmo_unidade = 1 and tmo_data = (Select max(tmo_data) " & _
                    " FROM Tab_Monetaria)"
                If Bdados.AbreTabela(Sql, RsAux) Then
                    Valor = CDbl(RsPrin(0)) * CDbl(RsAux(0))
                    If RsPrin(1) = "2" Then
                        Sql = "Select tmo_valor FROM Tab_Monetaria " & _
                        " WHERE tmo_unidade = 2 and tmo_data = (Select max(tmo_data) " & _
                        " FROM Tab_Monetaria)"
                        If Bdados.AbreTabela(Sql, RsAux) Then
                            Valor = CDbl(Valor) * CDbl(RsAux(0))
                        Else
                            Call Util.Avisa("Não existe valor de UFM cadastrado.")
                        End If
                        Bdados.FechaTabela RsAux
                    End If
                Else
                    Call Util.Avisa("Não existe valor de UFIR cadastrado.")
                End If
                Bdados.FechaTabela RsAux
            Case 3
                Valor = CDbl(ValorVenal) - (CDbl(ValorVenal) * CDbl(RsPrin(2))) / 100
                Valor = CDbl(Valor) - (CDbl(Valor) * CDbl(RsPrin(0))) / 100
        End Select
    Else
        Call Util.Avisa("Valor do componente " & Item & " não cadastrado.")
    End If
    GeraValorItem = Valor
    Bdados.FechaTabela RsPrin
    Bdados.FechaTabela RsAux
End Function

Public Function GeraNumNota(Ano As Integer, Operacao As Integer) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Nota  As String
    Sql = "Update Tab_Num_Correlativo set tnc_sequencia =tnc_sequencia+1 where tnc_tipo_operacao = " & Operacao
        '& " And tnc_Ano = " & Ano
    Bdados.Executa Sql
    Sql = "SELECT tnc_sequencia from tab_num_correlativo where " _
        & " tnc_tipo_operacao = " & Operacao
        '& " And tnc_Ano = " & Ano
    If Bdados.AbreTabela(Sql, Rs) Then
        Nota = Operacao & Format(IIf(IsNull(Rs(0)), 1, Rs(0)), "00000")
        GeraNumNota = Nota
    Else
        Call Util.Avisa("Problemas ao gerar Numero da Nota.")
        GeraNumNota = "00000000-00"
    End If
    Bdados.FechaTabela Rs
End Function
Public Function GeraNumCorrelativo(Valor As Integer, Operacao As Integer) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim novoNumero  As Long
    Sql = "Update Tab_Num_Correlativo set tnc_sequencia =tnc_sequencia+" & Valor & " where tnc_tipo_operacao = " & Operacao
        
    Bdados.Executa Sql
    Sql = "SELECT tnc_sequencia from tab_num_correlativo where " _
        & " tnc_tipo_operacao = " & Operacao
        '& " And tnc_Ano = " & Ano
    If Bdados.AbreTabela(Sql, Rs) Then
        novoNumero = IIf(IsNull(Rs(0)), 1, Rs(0))
        GeraNumCorrelativo = novoNumero
    Else
        
    End If
    Bdados.FechaTabela Rs
End Function

Public Function GeraInscMunicipal(Ano As Integer, Operacao As Integer, Imposto As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Inscricao  As String
    Dim Posicao As Double
'    Sql = "Select tcd_numero from tab_correlativo_disponivel ORDER BY tcd_numero ASC"
'    If Bdados.AbreTabela(Sql, Rs) Then
'        Rs.MoveFirst
'        Posicao = Rs(0)
'        Bdados.DeletaDados "tab_correlativo_disponivel", "tcd_numero =" & Posicao
'    Else
        Sql = "Update Tab_Num_Correlativo set tnc_sequencia = tnc_sequencia+1 where tnc_tipo_operacao = " & Operacao
        Bdados.Executa Sql
        Sql = "SELECT tnc_sequencia from tab_num_correlativo where " _
            & " tnc_tipo_operacao = " & Operacao
        If Not Bdados.AbreTabela(Sql, Rs) Then
            Call Util.Avisa("Problemas ao gerar Inscrição Municipal.")
            GeraInscMunicipal = "00000000-00"
            Exit Function
        Else
            Posicao = IIf(IsNull(Rs(0)), 1, Rs(0))
        End If
'    End If
    Inscricao = Operacao & Format(Posicao, "000000")
    Inscricao = Inscricao & CStr(GeraDV(Inscricao, 9))
    Inscricao = Inscricao & GeraDV(Inscricao, 10)
    GeraInscMunicipal = Format(Inscricao, "00000000-00")
    Bdados.FechaTabela Rs
End Function
Public Function GeraInscCadastral(Ano As Integer, Operacao As Integer, Imposto As String) As String
    
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Imovel  As String
    Sql = "Update Tab_Num_Correlativo set tnc_sequencia =tnc_sequencia+1 where tnc_tipo_operacao = " & Operacao
        '& " And tnc_Ano = " & Ano
    Bdados.Executa Sql
    Sql = "SELECT tnc_sequencia from tab_num_correlativo where " _
        & " tnc_tipo_operacao = " & Operacao
        '& " And tnc_Ano = " & Ano
    If Bdados.AbreTabela(Sql, Rs) Then
        Imovel = Operacao & Format(IIf(IsNull(Rs(0)), 1, Rs(0)), "000000")
        Imovel = Imovel & GeraDV(Imovel, 9)
        GeraInscCadastral = Format(Imovel, "00000000-0")
    Else
        Call Util.Avisa("Problemas ao gerar Inscrição Cadastral.")
        GeraInscCadastral = "00000000-0"
    End If
    Bdados.FechaTabela Rs
End Function

Function GeraDV(Inscricao As String, Base As Integer) As Integer
    Dim I As Integer
    Dim DvInsc As Integer
    Dim Inscricao_Temp As String
    
    DvInsc = 0
    Inscricao_Temp = Format(Inscricao, "000000000000")
    For I = 1 To Len(Inscricao_Temp)
        DvInsc = DvInsc + Val(Mid(Inscricao_Temp, I, 1)) * (((Len(Inscricao_Temp) - I) Mod (Base - 1)) + 2)
    Next
    DvInsc = DvInsc Mod 10
    If Inscricao_Temp > 1 Then
        DvInsc = 10 - DvInsc
    Else
        DvInsc = 0
    End If
    If DvInsc >= 10 Then DvInsc = DvInsc Mod 10
    GeraDV = DvInsc
End Function

Public Function ImovelIsento(InscricaoImovel As String, Unidade) As TipoIsencaoImovel
    Dim Sql As String
    Dim Rs As VSRecordset
    ' VERIFICA ISENÇÃO DO IMÓVEL
    Sql = "Select tdi_tco_cod_componente from tab_detalhe_imovel where tdi_tim_ic ='" & _
    InscricaoImovel & "' and tdi_tgc_cod_grupo = 3"
    If Unidade >= 200 Then Sql = Sql & " and tdi_tim_ic_unidade=" & Unidade
    If Bdados.AbreTabela(Sql, Rs) Then
        ImovelIsento = IIf(Rs!tdi_tco_cod_componente < 5, Rs!tdi_tco_cod_componente, 1)
    Else
        ImovelIsento = tim_SemIsencao
    End If
    Bdados.FechaTabela Rs
End Function

Public Sub ImprimeFC(Inscricao As String, Rpt As Object)
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TFC_Empresa.rpt") Then Exit Sub
        .titulo = "FAC - Ficha Cadastral do Contribuinte"
        If Trim(Inscricao) <> "" Then .Selecao = "{Tab_Contribuinte.tci_im} ='" & Inscricao & "'"
        Dim r As VSRecordset
        If Bdados.AbreTabela("SELECT VCI_ENDERECO FROM VIS_CONTRIBUINTE WHERE VCI_IM='" & Inscricao & "'", r) Then
            .Formulas "VT_ENDERECO", r(0)
        End If
        .Visualizar
    End With
    Set Rpt = Nothing
End Sub


Function VerificaImovelAntigo(CodigoLogr As Long, NumeroImovel As String, Complemento As String) As Boolean
    Dim Query As String
    Dim Rs As VSRecordset
    Query = "SELECT tim_tlg_cod_logradouro FROM Tab_Imovel WHERE "
    Query = Query & "tim_tlg_cod_logradouro = '" & CodigoLogr
    Query = Query & "' and tim_numero ='" & NumeroImovel & "'"
    Query = Query & " and tim_complemento ='" & Complemento & "'"
    If Bdados.AbreTabela(Query, Rs) Then
        VerificaImovelAntigo = True
    Else
        VerificaImovelAntigo = False
    End If
    Bdados.FechaTabela Rs
End Function

Function BuscaCodLogr(Logr As String) As Integer
    Dim RsLogr As VSRecordset
    Dim Sql As String
    Sql = "Select ttl_cod_tip_Logr from tab_tipo_logr Where ttl_nome = '" & Logr & "'"
    If Bdados.AbreTabela(Sql, RsLogr) Then
        BuscaCodLogr = RsLogr(0)
    End If
    Bdados.FechaTabela RsLogr
End Function

Public Sub ImprimeCartaoIM()
    
End Sub

Public Function VerificaEmpresaAntiga(CgcCpf As String, Razao As String) As Integer
    Dim RsEmp As VSRecordset
    Dim Sql As String
    
    If Trim(CgcCpf) <> "" Then
            Sql = "SELECT tci_cgc_cpf,tci_nome from Tab_Contribuinte where  tci_cgc_cpf = '" & CgcCpf & "'"
            Sql = Sql & " and tci_tsc_cod_sit_cad = 1"
            If Bdados.AbreTabela(Sql, RsEmp) Then
                VerificaEmpresaAntiga = 1
                Bdados.FechaTabela RsEmp
                Exit Function
            End If
    End If
    
    If Trim(Razao) <> "" Then
            Sql = "SELECT tci_cgc_cpf,tci_nome from Tab_Contribuinte where  tci_nome = '" & Trim(Razao) & "'"
            Sql = Sql & " and tci_tsc_cod_sit_cad = 1"
            If Bdados.AbreTabela(Sql, RsEmp) Then
                VerificaEmpresaAntiga = 2
            End If
    End If
    Bdados.FechaTabela RsEmp
End Function

Public Function FormataInscricao(Inscricao As String, TipoInscricacao As tipoInsc) As String
    If TipoInscricacao = InscContrib Then
        If Len(Inscricao) = 10 Then
            Inscricao = Format(Inscricao, "00000000-00")
        ElseIf Len(Inscricao) = 11 Then
            'BCP AQUI ESTA MUDANDO O VALOR DA INSCRICAO
            'Inscricao = Format(Inscricao, "000000000-00")
        ElseIf Len(Inscricao) = 12 Then
            Inscricao = Format(Inscricao, "0000000000-00")
        End If
    ElseIf TipoInscricacao = InscImovel Then
        If Len(Inscricao) = 9 Then Inscricao = Format(Inscricao, "00000000-0")
    End If
    FormataInscricao = Inscricao
End Function

Public Sub BuscaLogradouro(Logr As TipoLogradouro, ComboReferencia As Object, ComboCarregar As Object)
    Dim Query As String
    Dim Rs As VSRecordset
    Static Codigo As Byte
    
    DoEvents
    If Logr = Rua Then
        Query = "SELECT distinct(tlg_nome), tlg_ttl_cod_TIP_logr FROM tab_logradouro " & _
        "WHERE tlg_ttl_cod_TIP_logr = (SELECT ttl_cod_tip_Logr FROM " & _
        "Tab_tipo_Logr WHERE ttl_nome = '" & ComboReferencia & "') and tlg_tmu_cod_municipio=" & AplicacoesVTFuncoes.Codigo_Municipio
    ElseIf Logr = Bairro Then
        Query = "SELECT tba_nome FROM tab_bairro " & _
        "WHERE tba_cod_bairro in (select tlg_tba_cod_bairro FROM " & _
        "Tab_Logradouro WHERE tlg_ttl_cod_TIP_logr = " & Codigo & _
        " and tlg_nome  = '" & ComboReferencia & "') and tba_tmu_cod_municipio=" & AplicacoesVTFuncoes.Codigo_Municipio
    ElseIf Logr = TipoLogr Then
        Query = "SELECT ttl_nome FROM TAB_TIPO_LOGR " & _
        "WHERE TTL_COD_TIP_LOGR in (SELECT distinct(tlg_ttl_cod_tip_logr) FROM " & _
        "TAB_LOGRADOURO WHERE tlg_tba_cod_bairro = " & ComboReferencia.ListIndex + 1 & " and tlg_tmu_cod_municipio=" & AplicacoesVTFuncoes.Codigo_Municipio & ")"
    End If
    
    ComboCarregar.Clear
    If Bdados.AbreTabela(Query, Rs) Then
        Rs.MoveFirst
        If Logr = Rua Then Codigo = Rs(1)
        Do While Not Rs.EOF
            ComboCarregar.AddItem Rs(0)
            Rs.MoveNext
        Loop
    End If
    DoEvents
    Bdados.FechaTabela Rs
End Sub

Public Function PegaCodLogr(Bairro As String, TipoLogr As String, Logr As String) As Double
    Dim Query As String
    Dim RsLogr As VSRecordset
    Query = "SELECT tlg_cod_logradouro FROM tab_logradouro WHERE "
    Query = Query & " tlg_tba_cod_bairro = (select tba_cod_bairro FROM Tab_Bairro WHERE tba_nome ='" & Bairro & "' and TBA_TMU_COD_MUNICIPIO =" & AplicacoesVTFuncoes.Codigo_Municipio & ")"
    Query = Query & "  and tlg_ttl_cod_tip_logr = (SELECT ttl_cod_tip_Logr FROM Tab_Tipo_Logr WHERE ttl_nome ='" & TipoLogr & "')"
    Query = Query & " and tlg_nome = '" & Logr & "' and tlg_tmu_cod_municipio=" & AplicacoesVTFuncoes.Codigo_Municipio
    If Bdados.AbreTabela(Query, RsLogr) Then
        PegaCodLogr = RsLogr(0)
    End If
    Bdados.FechaTabela RsLogr
End Function
Public Function ContribuinteHabilitado(Inscricao As String, Situacao As String) As Boolean
    Dim RsSit As VSRecordset
    Dim Sql As String
    ContribuinteHabilitado = True
    Sql = "Select tsc_nome,tci_tsc_cod_sit_cad from Tab_Contribuinte, Tab_sit_cadastral" & _
        " where tsc_cod_sit_cad = tci_tsc_cod_sit_cad and tci_im = '" & Inscricao & "'"
    If Bdados.AbreTabela(Sql, RsSit) Then
        ContribuinteHabilitado = IIf(RsSit!tci_tsc_cod_sit_cad = 1, True, False)
        Situacao = "" & RsSit!tsc_nome
    Else
        Sql = "Select tsc_nome,tci_tsc_cod_sit_cad from Tab_Contribuinte, Tab_sit_cadastral" & _
        " where tsc_cod_sit_cad = tci_tsc_cod_sit_cad and tci_im = '" & Inscricao & "'"
        If Bdados.AbreTabela(Sql, RsSit) Then
            ContribuinteHabilitado = IIf(RsSit!tci_tsc_cod_sit_cad = 1, True, False)
            Situacao = RsSit!tsc_nome
        End If
    End If
    Bdados.FechaTabela RsSit
    
End Function

Public Sub ImprimeBoletoIptu(Rs As Object, Rpt)

End Sub

Private Sub Class_Initialize()
    Set Conta = New ContaCorrente
End Sub

Public Function NomeTributo(Tributo As TipoTributo) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    On Error Resume Next
    Sql = "SELECT TGE_NOME FROM TAB_GERAL WHERE  TGE_TIPO = (SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME ='NOME TRIBUTO'  ) and TGE_CODIGO =" & Tributo
    If Bdados.AbreTabela(Sql, Rs) Then
        NomeTributo = Rs(0)
    End If
    Bdados.FechaTabela Rs
End Function
Public Function PegaNomeUsuario(Usuario As String) As String
    Dim Sql As String
    
    Sql = "Select tus_nome from tab_usuario where tus_cod_usuario = '" & Usuario & "'"
    If Bdados.AbreTabela(Sql) Then
        PegaNomeUsuario = "" & Bdados.Tabela(0)
    Else
        PegaNomeUsuario = ""
    End If
End Function
Public Function BuscaAnoImposto(Imposto As String, Ano As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim AnoIdeal As String
    Dim TemSemAno As Boolean
    On Error GoTo Trata
    TemSemAno = False
    Ano = Edita.TiraTudo(Ano)
   ' Sql = "Select tpi_ano_imposto from Tab_Parametro_Imposto where tpi_tip_cod_imposto='" & Trim(Imposto) & "' order by tpi_ano_imposto asc"
    
    
    Dim tabRs As New Recordset
    Dim Bd As New Connection
    Dim strconn As String
    strconn = Bdados.Conexao.DBConnection.ConnectionString
    Bd.Open strconn
    
    Set tabRs = Bd.Execute("SP_SEL_PARAMETROIMPOSTO_ANO '" & Trim(Imposto) & "'")
  
    If Not tabRs.EOF Then
        tabRs.MoveFirst
        Do While Not tabRs.EOF
            AnoIdeal = "" & tabRs!tpi_ano_imposto
            If Len(Trim(AnoIdeal)) = 0 Then
                TemSemAno = True
            Else
                If CInt(Nvl(Ano, 0)) <= CInt(AnoIdeal) Then
                    TemSemAno = False
                    Exit Do
                End If
            End If
            tabRs.MoveNext
        Loop
    End If
    tabRs.Close
    Bd.Close
    
    BuscaAnoImposto = IIf(TemSemAno, "", AnoIdeal)
    Exit Function
Trata:
    Avisa Err.Description
    Exit Function
    Resume
End Function

Public Function PegaSubConta(Conta As String) As String
    Dim Sql                        As String
    Dim Rs                         As VSRecordset
    
    Sql = "SELECT TIP_SUB_CONTA "
    Sql = Sql & " FROM TAB_IMPOSTO "
    Sql = Sql & " WHERE TIP_COD_IMPOSTO = '" & Conta & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Not IsNull(Rs.Fields("TIP_SUB_CONTA")) Then
            PegaSubConta = Rs.Fields("TIP_SUB_CONTA")
        Else
            PegaSubConta = ""
        End If
    Else
        PegaSubConta = ""
    End If
End Function

Public Function BuscaNumeroIncidencia(Contribuinte As String, Periodo As Long, CodImposto As String) As Integer
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select count(*) +1 from tab_geracao_tributo where tgt_im ='" & Contribuinte & _
        "' and tgt_periodo = " & Periodo & " and tgt_tip_cod_imposto = '" & CodImposto & "' and tgt_incidencia > 0"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaNumeroIncidencia = Rs(0)
    End If
End Function

'BALSAS - 01/05/2002
Public Function GeraIptu(Cidade As CidadeIptu, Contribuintes As VSRecordset, PeriodoInicial As Integer, PeriodoFinal As Integer, TipoGeracao As TipoGeracaoImposto) As Double
    Dim ValorVenalEdificacao As Double
    Dim ValorVenalTerreno As Double
    Dim ValorVenalImovel As Double
    Dim AreaConstruida As Double
    Dim AreaTerreno As Double
    Dim CoeficienteFracaoIdeal As Double
    Dim CoeficienteTerra As Double
    Dim CoeficienteEdific As Double
    Dim ValorTotalEdificacoes As Double
    Dim AliquotaTerreno As Double
    Dim AliquotaPredio As Double
    Dim Rs As VSRecordset
    Dim TipoIsencao As TipoIsencaoImovel
    Dim Sql As String
    Dim TipoImovel As Byte
    Dim ValorImposto As Double
    Dim CodImposto As String
    Dim I As Long
    Dim j As Byte
    Dim Fim As Boolean
    Dim CodParcelamento As Double
    Dim Conta As New ContaCorrente
    Dim Reg  As Double
    Dim InscImovel As String
    Dim ImovelAtual As String
    Dim Zona As Byte
    Dim MinImposto As Double
    Dim CodPagamento As String
    Dim CodigoTaxa As String
    'Const DataVencimento = "30/12/2002"
    Select Case Cidade
    
            Case cip_BCorda
                If Contribuintes.EOF Then Exit Function
                'Do While Not Contribuintes.EOF 'PARA CADA IMOVEL SELECIONADO
                    Call CriaIptu(Contribuintes, PeriodoInicial, PeriodoFinal)
                 '   Contribuintes.MoveNext
                'Loop
            Case cip_Balsas
                If Contribuintes.EOF Then Exit Function
                Sql = "select distinct tdi_tim_ic,tdi_tim_ic_unidade from tab_detalhe_imovel " & _
                    " where tdi_tim_ic = '" & Contribuintes!tim_ic & "'"
                CodImposto = BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
                MinImposto = BuscaValorMinimoImposto(CodImposto, Year(Date)) 'IMPOSTO MINIMO
                CodigoTaxa = BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto = 'TSU'")
                InscImovel = ""
                Fim = False
                Contribuintes.MoveFirst
                Do While Not Contribuintes.EOF 'PARA CADA IMOVEL SELECIONADO
                    TipoImovel = Contribuintes!tim_tipo_imovel
                    If PeriodoInicial < 2001 Then
                        TipoImovel = 2
                        TipoIsencao = tii_IsentoTaxas
                    End If
                    Sql = "select distinct tdi_tim_ic,tdi_tim_ic_unidade from tab_detalhe_imovel " & _
                                        " where tdi_tim_ic = '" & Trim(Contribuintes!tim_ic) & "'"
                    If Bdados.AbreTabela(Sql, Rs, Registros) Then
                        Rs.MoveFirst 'PRIMEIRO IMOVEL DA LISTA
                        Zona = Nvl("" & Contribuintes!TIM_ZONA, 1)
                        'ZERANDO VARIAVEIS ACUMULATIVAS
                        AliquotaTerreno = 0
                        AliquotaPredio = 0
                        CoeficienteEdific = 0
                        CoeficienteTerra = 0
                        ValorTaxas = 0
                        AreaTerreno = 0
                        AreaConstruida = 0
                        ValorVenalTerreno = 0
                        ValorVenalEdificacao = 0
                        ValorTotalEdificacoes = 0
                        'VERIFICANDO TIPO DA ISENCAO
                        TipoIsencao = ImovelIsento(Rs!tdi_tim_ic, Rs!tdi_tim_ic_unidade)
                        If PeriodoInicial < 2001 Then
                            TipoIsencao = tii_IsentoTaxas
                        End If
                        'PODE CALCULAR IPTU
                        If TipoIsencao <> tii_IsentoImposto And TipoIsencao <> tim_IsentoTotal Then
                           If TipoImovel = 2 Then 'SE FOR APENAS TERRENO
                                   AliquotaTerreno = BuscaAliquotaImovelIPTU(tim_TerrenoSemEdif, Nvl("" & Zona, 5))
                                   CoeficienteTerra = BuscaCoeficienteImovelIPTU(tzo_ZonaUrbana, Nvl("" & Zona, 5))
                           Else 'SE FOR TERRENO COM PREDIO
                                   AliquotaTerreno = BuscaAliquotaImovelIPTU(tim_TerrenoComEdif, Nvl("" & Zona, 5))
                                   AliquotaPredio = BuscaAliquotaImovelIPTU(tim_Edificacao, Nvl("" & Zona, 5))
                                   CoeficienteEdific = BuscaCoeficienteEdificacao(Contribuintes!tim_ic, Contribuintes!TIM_UNIDADE)
                                   CoeficienteTerra = BuscaCoeficienteImovelIPTU(tzo_ZonaUrbana, Nvl("" & Zona, 5))
                           End If
                           'CALCULANDO IPTU
                           AreaTerreno = BuscaValorArea(Rs!tdi_tim_ic, Rs!tdi_tim_ic_unidade, 108)
                           AreaConstruida = BuscaValorArea(Rs!tdi_tim_ic, Rs!tdi_tim_ic_unidade, 112)
                           ValorVenalTerreno = AreaTerreno * CoeficienteTerra
                           ValorVenalEdificacao = AreaConstruida * CoeficienteEdific
                           
                            If Rs!tdi_tim_ic_unidade < 200 Then 'SE NAO FOR CONDOMINIO
                                ImovelAtual = "" & Rs!tdi_tim_ic
                                Do
                                    'REGISTRO DE OUTRO IMOVEL
                                    If ImovelAtual <> Nvl("" & Rs!tdi_tim_ic, 0) Then 'PARA CADA EDIFICACAO DO IMOVEL
                                        Rs.MovePrevious
                                        Exit Do
                                    End If
                                     AreaConstruida = BuscaValorArea(Rs!tdi_tim_ic, Rs!tdi_tim_ic_unidade, 112)
                                     CoeficienteEdific = BuscaCoeficienteEdificacao(Rs!tdi_tim_ic, Rs!tdi_tim_ic_unidade)
                                     ValorVenalEdificacao = AreaConstruida * CoeficienteEdific
                                     ValorTotalEdificacoes = ValorTotalEdificacoes + ValorVenalEdificacao 'ACUMULA AS EDIFICACOES
                                     Rs.MoveNext
                                Loop While Not Rs.EOF
                                ValorVenalEdificacao = ValorTotalEdificacoes
                            Else 'SE FOR CONDOMINIO
                                If Rs!tdi_tim_ic_unidade = 200 Or Rs!tdi_tim_ic_unidade = 600 Then 'BT
                                    CoeficienteFracaoIdeal = BuscaFracaoIdealImposto(Rs!tdi_tim_ic, ValorVenalTerreno)
                                Else 'PARA CADA LOTE DO CONDOMINIO
                                    If CoeficienteFracaoIdeal = 0 Then
                                        CoeficienteFracaoIdeal = BuscaFracaoIdealImposto(Rs!tdi_tim_ic, ValorVenalTerreno)
                                    End If
                                    AreaConstruida = BuscaValorArea(Rs!tdi_tim_ic, Rs!tdi_tim_ic_unidade, 112)
                                    ValorVenalTerreno = AreaTerreno * CoeficienteFracaoIdeal
                                    ValorVenalEdificacao = AreaConstruida * CoeficienteFracaoIdeal
                                End If
                            End If
                            If Rs.EOF Then 'ACABOU A RELACAO DE IMOVEIS
                                Rs.MovePrevious
                                Fim = True
                            End If
                            'NAO GERA PARA O CONDOMINIO EM SI. SO PARA OS LOTES DO CONDOMINIO
                            'GAMBIARRA SILMAR
                            If Rs!tdi_tim_ic_unidade <> 200 Then '(RETIREI - ANDRE[03/09/2003]
                                ValorVenalImovel = ValorVenalTerreno + ValorVenalEdificacao
                                'ValorVenalImovel = ValorVenalTerreno + ValorVenalEdificacao
                                Call Bdados.AtualizaDados("Tab_Imovel", Bdados.PreparaValor(ValorVenalTerreno, ValorVenalEdificacao), "TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC", "tim_ic ='" & Rs!tdi_tim_ic & "'") ' AND tim_unidade =" & Contribuintes!tdi_tim_ic_unidade)
                                ValorImposto = (ValorVenalTerreno * AliquotaTerreno) + (ValorVenalEdificacao * AliquotaPredio)
                                
                                If TipoIsencao <> tii_IsentoTaxas Then 'PODE CALCULAR T.S.U
                                    If Zona < 5 Then
                                        ValorTaxas = CalculaTaxasIptu(Contribuintes!tim_ic, Contribuintes!tim_tci_im, CodImposto, PeriodoFinal, ValorImposto, cip_Balsas, CInt(Zona), CInt(TipoImovel))
                                    Else 'NA ZONA 5 NAO HÁ T.S.U
                                        ValorTaxas = 0
                                    End If
                                Else
                                    ValorTaxas = 0
                                End If
                                
                                If Rs!tdi_tim_ic_unidade > 200 Then ' NO CASO DE CONDOMINIO
                                    ValorTaxas = ValorTaxas * BuscaFracaoIdealArea(Rs!tdi_tim_ic, AreaConstruida)
                                End If
                             End If  '(RETIREI - ANDRE[03/09/2003]
                        Else 'SO CALCULA T.S.U
                            ValorImposto = 0
                            If Rs!tdi_tim_ic_unidade < 200 Then  'SE NAO FOR CONDOMINIO
                                ImovelAtual = Rs!tdi_tim_ic
                                Do 'PASSA TODOS OS BC'S
                                    'REGISTRO DE OUTRO IMOVEL
                                    If ImovelAtual <> Nvl("" & Rs!tdi_tim_ic, 0) Then 'PARA CADA EDIFICACAO DO IMOVEL
                                        Rs.MovePrevious
                                        Exit Do
                                    End If
                                     Rs.MoveNext
                                Loop While Not Rs.EOF
                            End If
                            If Rs.EOF Then Rs.MovePrevious
                            If TipoIsencao = tii_IsentoImposto Then 'SO ISENTO DE IMPOSTO
                                If Zona < 5 Then
                                    ValorTaxas = CalculaTaxasIptu(Rs!tdi_tim_ic, "" & Contribuintes!tim_tci_im, CodImposto, PeriodoFinal, ValorImposto, cip_Balsas, CInt(Zona), CInt(TipoImovel))
                                Else 'NA ZONA 5 NAO HÁ T.S.U
                                    ValorTaxas = 0
                                End If
                            End If
                        End If
                        'Grava IPTU GERADO
'                        ValorImposto = ValorImposto + ValorTaxas
                        If (ValorImposto) >= MinImposto Or TipoImovel <> 1 Then
                            Call AtualizaValorVenal(Contribuintes!tim_ic, Format(ValorVenalImovel, "#,##0.00"), Contribuintes!TIM_UNIDADE)
                            For I = PeriodoInicial To PeriodoFinal
'*********************************************************************************************************
                                'GAMBIARRA SILMAR
'                                If PeriodoInicial < 2001 Then
'                                    DataVenc = BuscaDataVencimento(CodImposto, CStr(PeriodoInicial))
'                                Else
'                                    DataVenc = DataVencimento
'                                End If
'                                If PeriodoInicial < Year(Date) Then
'                                    DataVenc = UltimoDiaDoMes(Date)
'                                End If
                                
                                DataVenc = Imposto.BuscaDataVencimento(CodImposto, I)
                                
'*********************************************************************************************************
                                CodPagamento = Conta.NumPagamento(Contribuintes!tim_tci_im, I, CodImposto, Contribuintes!tim_ic, 0, True, 0)
                                
                                Conta.GeraPagamento Contribuintes!tim_tci_im, Contribuintes!tim_ic, CodImposto, I, DataVenc, Format(ValorImposto, "#,##0.00"), _
                                    Format(Conta.CalculaValoresMultaAvulsos(CodImposto, I, EtcCreditoTributario, Date, DataVenc, ValorImposto), "#,##0.00"), _
                                    Format(Conta.CalculaValoresJurosAvulsos(CodImposto, I, EtcCreditoTributario, Date, DataVenc, ValorImposto), "#,##0.00"), _
                                     CodPagamento _
                                    , 0, 0, Format(ValorTaxas, "#,##0.00"), CodigoTaxa, , , , , etgIPTU
                            
'                                Call Conta.MovimentaContaContribuinte(CodPagamento, Contribuintes!tim_tci_im, Contribuintes!tim_ic, CodImposto, i, EtcCreditoTributario, CInt(j), Date)
                                If TipoGeracao <> tgi_SemParcelas Then
                                    'Call Conta.CriaParcelamento(Contribuintes!tim_tci_im, Contribuintes!TIM_IC, CStr(i), CStr(i), CodImposto, 3, 0, DataVenc, EtcCreditoTributario, CodParcelamento, ValorImposto + ValorTaxas, ValorImposto + ValorTaxas)
                                    
                                End If
                            Next
                        Else
                            If TipoImovel > 1 Then
                                Conta.GeraPagamento Contribuintes!tim_tci_im, Contribuintes!tim_ic, CodImposto, I, DataVenc, Format(0, "#,##0.00"), _
                                        Format(0, "#,##0.00"), _
                                        Format(0, "#,##0.00"), _
                                        Conta.GeraCodPagamento(EtsIptu), 0, CInt(j), Format(0, "#,##0.00")
                            End If
                        End If
                        Reg = Reg + 1
                    End If
                    If Contribuintes!TIM_UNIDADE >= 200 Or Fim Then
                        InscImovel = Contribuintes!tim_ic
                    End If
                    Contribuintes.MoveNext
                    DoEvents
                Loop
                Reg = Reg - 1
                Fim = False
                DoEvents
            Case cip_Colinas
                    
    End Select
    GeraIptu = ValorImposto
End Function

Public Function BuscaNomeCompletoImposto(CodImposto As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select tip_nome_imposto from tab_imposto where tip_cod_imposto = '" & CodImposto & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaNomeCompletoImposto = Rs!tip_nome_imposto
    End If
    Bdados.FechaTabela Rs
End Function

Public Function DataVencimentoNova(Optional DataPadrao As String) As String

    TCIS103.Tag = DataPadrao
    
    TCIS103.Show 1
    TCIS103.Tag = DataPadrao
    If NovaDataVencimento <> "" Then
        DataVencimentoNova = NovaDataVencimento
    End If
'    Dim NovaData  As String
'    Dim Aux As Byte
'    Aux = 0
'    Do
'        If Aux > 0 Then Avisa "Informe uma data válida."
'        NovaData = InputBox("Informe a data de vencimento da parcela", "Data de Vencimento", DataPadrao)
'        NovaData = Format(NovaData, "dd/mm/yyyy")
'        If NovaData = "" Then
'            Util.Avisa "Operacão Cancelada pelo usuário."
'            NovaData = ""
'            Exit Function
'        End If
'        If Len(NovaData) < 6 Then
'            Util.Avisa "Data inválida."
'            NovaData = DataVencimentoNova(NovaData)
'            Exit Function
'        End If
'        NovaData = Mid(NovaData, 1, 2) + "/" + Mid(NovaData, 3, 2) + "/" + Right(NovaData, 4)
'
'        If Not IsDate(NovaData) Then
'            Util.Avisa "Data inválida."
'            NovaData = DataVencimentoNova(NovaData)
'            Exit Function
'        End If
'
'        If Len(Format(NovaData, "dd/mm/yyyy")) < 10 Then
'            Util.Avisa "Data inválida."
'            NovaData = DataVencimentoNova(NovaData)
'            Exit Function
'        End If
'        If CDate(Format(NovaData, "dd/mm/yyyy")) < Date Then
'            Util.Avisa "Data de vencimento não pode ser menor que data atual."
'            NovaData = DataVencimentoNova(NovaData)
'            Exit Function
'        End If
'        Aux = 1
''        NovaData = Format(CStr(NovaData), "dd/mm/yyyy")
'    Loop While Not IsDate(NovaData)
'    DataVencimentoNova = NovaData
End Function

Private Sub CarregaComponetes()
    Dim Sql As String
    Dim Rs As VSRecordset
    On Error GoTo Trata
    
    Sql = "Select tco_cod_componente,tco_valor,tco_cod_componente_fator_calc " & _
        " from tab_componente_avancado where tab_componente_avancado.tco_tmu_cod_municipio = " & _
        AplicacoesVTFuncoes.Codigo_Municipio
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            Componentes(Rs!tco_cod_componente).Valor = Rs!tco_valor
            Componentes(Rs!tco_cod_componente).ComponenteFator = Rs!tco_cod_componente_fator_calc
            Rs.MoveNext
        Loop
    End If
    Exit Sub
Trata:
    Avisa Err.Description
    Exit Sub
    Resume
End Sub


Public Function BuscaValorFixoTaxa(Imposto As String, Ano As String) As Double
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Condicao As String
    Sql = "Select tpi_valor_taxa_fixa from Tab_Parametro_Imposto " & _
        " where tpi_tip_cod_imposto='" & Imposto & "'"
    Condicao = " and tpi_ano_imposto= '" & ClsImposto.BuscaAnoImposto(Imposto, Left(Ano, 4)) & "'"
    If Bdados.AbreTabela(Sql & Condicao, Rs) Then
        BuscaValorFixoTaxa = Rs(0)
    End If
    Bdados.FechaTabela Rs
End Function

