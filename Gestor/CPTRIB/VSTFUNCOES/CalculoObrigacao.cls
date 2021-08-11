VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalculoObrigacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TipoLanc
    etlHomologacao = 1
    etlDeclaracao = 2
End Enum

Public TipoLancamento As TipoLanc
Public GravaObrigacao As Boolean
Public GravaObrigacaoAcessoria As Boolean
Public TaxaUnificada As Boolean
Public DadosIptu As VSIptu
Public Function AlvaraVeiculo(Im As String, Periodo As String, CodImposto As String, ByVal DtVenc As String, Optional PlacaVeiculo As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    'Eu Raimundo mundei o seguinte sql
    'O cliente pediu que fosse impresso um alvara para cada veiculo
    'mas o valor do alvara seria o da atividade informada no cadastro não a principal
    'Antes
    'Sql = "Select TAE_VALOR "
    'Sql = Sql & " from TAB_CONTRIBUINTE,TAB_ATIVIDADE_ECONOMICA "
    'Sql = Sql & " where TCI_IM ='" & IM & "'"
    'Sql = Sql & " AND TCI_TAE_CAE_TRANSPORTE = TAE_CAE"
    'Depois
    Sql = " SELECT TAE_VALOR"
    Sql = Sql & " From TAB_TRANSPORTADOR_VEICULO, TAB_ATIVIDADE_ECONOMICA"
    Sql = Sql & " WHERE TTV_TCI_IM = " & Bdados.Converte(Im, tctexto)
    Sql = Sql & " AND TTV_TAE_CAE = TAE_CAE"
    Sql = Sql & " AND  TTV_PLACA = " & Bdados.Converte(PlacaVeiculo, tctexto)
    If Bdados.AbreTabela(Sql, Rs) Then
        AlvaraVeiculo = Rs!TAE_VALOR
    Else
        Sql = "Select TAE_VALOR "
        Sql = Sql & " from TAB_ATIVIDADE_ECONOMICA "
        Sql = Sql & " where TAE_CAE = " & Temp.PegaParametro(Bdados, "CAE TRANSPORTADOR")
        If Bdados.AbreTabela(Sql, Rs) Then
            AlvaraVeiculo = Rs!TAE_VALOR
        End If
    End If
End Function

Public Function Alvara(Im As String, Periodo As String, CodImposto As String, ByRef DtVenc As String, Optional PlacaVeiculo As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim InicioAtividade As String
    Dim ValorAlvara As Double
    Dim RsData As VSRecordset
    Dim CAE As String
    Dim FatorAlvara As Double
    Dim AtividadesSecundarias As Integer
    Dim PercAtivSecund As Double
    Dim Desconto As Double
    
    GravaObrigacao = True
    TipoLancamento = etlHomologacao
    If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Or _
        Temp.PegaParametro(Bdados, "ALVARA IPM") = "SIM" Then
        Alvara = AlvaraPetrolina(Im, Periodo, CodImposto, DtVenc)
        Exit Function
    End If
    'Se possui placa gera para atividade
    If Trim(PlacaVeiculo) <> "" And CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) Then
        Alvara = AlvaraVeiculo(Im, Periodo, CodImposto, DtVenc, PlacaVeiculo)
        'Raimundo processo para gerar alvara para cada atividade...
        Dim RA_RS As VSRecordset
        Dim Dados  As String
        Sql = "select * "
        Sql = Sql & " From TAB_TRANSPORTADOR_VEICULO "
        Sql = Sql & " where TTV_TCI_IM = " & Bdados.Converte(Im, tctexto)
        Sql = Sql & " And TTV_PLACA = " & Bdados.Converte(PlacaVeiculo, tctexto)
        If Bdados.AbreTabela(Sql, RA_RS) Then
            
            Dados = RA_RS.Fields("TTV_INICIO_ATIVIDADE")
            Dados = (13 - Month(RA_RS.Fields("TTV_INICIO_ATIVIDADE")))
            Dim x As Double
            Dim Base As Double
            Base = 12
            x = (Dados * Alvara) / Base
            Alvara = x
        End If
        Exit Function
    End If
    
    If AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Then
        Sql = "SELECT TCI_FATOR_ALVARA,TAE_VALOR FROM TAB_CONTRIBUINTE,TAB_ATIVIDADE_ECONOMICA WHERE TCI_IM = '" & Im & "' AND tci_tae_cae = tae_cae"
        If Bdados.AbreTabela(Sql, Rs) Then
            ValorAlvara = CDbl(Nvl("" & Rs!TAE_VALOR, 0)) * CDbl(Nvl("" & Rs!tci_fator_alvara, 0)) * CDbl(TrocaPic(Nvl(Temp.PegaParametro(Bdados, "VALOR UFM"), 1), ".", ","))
            Dim ValorMinimo As Double
            Dim ValorMaximo As Double
            
            ValorMinimo = TrocaPic(Nvl(Temp.PegaParametro(Bdados, "ALVARA VALOR MINIMO"), 0), ".", ",")
            ValorMaximo = TrocaPic(Nvl(Temp.PegaParametro(Bdados, "ALVARA VALOR MAXIMO"), 0), ".", ",")
            
            If ValorAlvara < ValorMinimo Then
                ValorAlvara = ValorMinimo
            ElseIf ValorAlvara > ValorMaximo Then
                ValorAlvara = ValorMaximo
            End If
'            ValorAlvara = ValorAlvara + 1.21 + 3.63
            Alvara = ValorAlvara
            GravaObrigacao = True
        End If
        Exit Function
    End If
    Sql = "select tci_inicio_atividade,tci_tae_cae,tci_fator_alvara,tci_tae_cae_secund," & _
        "tci_tae_cae_terc from tab_contribuinte where tci_im ='" & Im & "'"
    If Bdados.AbreTabela(Sql, RsData) Then
        If IsNull(RsData(0)) Then Exit Function
        'PEGA PARAMETROS
        InicioAtividade = IIf(IsNull(RsData!TCI_INICIO_ATIVIDADE), "01/01/1900", RsData!TCI_INICIO_ATIVIDADE)
        If Periodo < Year(InicioAtividade) And UCase(AplicacoesVTFuncoes.Municipio) <> "BARRA MANSA" Then
            GravaObrigacao = False
            Exit Function
        End If
        CAE = RsData!tci_tae_cae
        FatorAlvara = CDbl(Nvl(Edita.TrocaPic(RsData!tci_fator_alvara, ".", ","), 0))
        If Format(Year(InicioAtividade), "0000") = Left(Periodo, 4) Then
            InicioAtividade = 13 - Month(InicioAtividade)
        End If
        PercAtivSecund = CDbl(Nvl(BuscaNaGeral("ALVARA SECUNDARIO", 1), 0))
        AtividadesSecundarias = PercAtivSecund * CDbl(IIf(Nvl("" & RsData!tci_tae_cae_secund, 0) = 0, 0, 1) + IIf(Nvl("" & RsData!tci_tae_cae_terc, 0) = 0, 0, 1))
        DtVenc = Imposto.BuscaDataVencimento(CodImposto, CLng(Periodo))
        'VERIFICA SE ATIVIDADE É ESTIMATIVA
        Sql = "SELECT   TAT_LIMITE_INFERIOR,TAT_LIMITE_SUPERIOR,TAT_LIMITE_VALOR,TAT_LIMITE_FATOR "
        Sql = Sql & " FROM TAB_ATIVIDADE_ESTIMADA WHERE TAT_TAE_CAE = " & CAE & " and TAT_ANO =" & Periodo & " order by TAT_LIMITE_INFERIOR asc"
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
            
            Sql = "Select TAV_FATOR,TCI_FATOR_alvara,"
            Sql = Sql & " TAV_VALOR_ALVARA,tip_cod_imposto "
            Sql = Sql & " from Tab_Contribuinte, TAB_ATIVIDADE_VALOR_TRIBUTOS,Tab_Imposto "
            Sql = Sql & " where "
            
            'Depois
            If UCase(AplicacoesVTFuncoes.Municipio) = "BALSAS" Then
                If Len(Im) = 11 Or Len(Im) = 12 Then
                    Sql = Sql & " tci_im = '" & Im & "'"
                Else
                    Sql = Sql & " tci_cgc_cpf = '" & Im & "'"
                End If
            Else
                If Trim(Im) <> "" Then Sql = Sql & " tci_im = '" & Im & "'"
            End If
            Sql = Sql & " and tci_tae_cae = TAV_TAE_CAE and tip_cod_imposto='" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) & "'"
            Sql = Sql & " AND TAV_ANO =" & Imposto.BuscaAnoImposto(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)), Left(Periodo, 4))
            
            If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
                'DEVE SER MUDADO, ANTES DE RODAR A 1ª VEZ EM BARRA MANSA -
                'ALTERACAO DA TABELA DE ATIVIDADE POR ANDRE EM 08-01-2006(COLINAS/MA)
                Dim RSMax As VSRecordset
                Dim SqlMax As String
                Dim MaxAtividadeSecundaria   As String
                
                SqlMax = "SELECT MAX(TAE_VALOR) AS TOTAL"
                SqlMax = SqlMax & " FROM VIS_ATIVIDADE_SECUNDARIA "
                SqlMax = SqlMax & "  WHERE TAS_TCI_IM = '" & Im & "'"
                If Bdados.AbreTabela(SqlMax, RSMax) Then
                     MaxAtividadeSecundaria = RSMax.Fields("TOTAL")
                End If
            End If
            If Bdados.AbreTabela(Sql, Rs) Then
                If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
                    If CDbl(Rs!TAV_VALOR_ALVARA) > CDbl(MaxAtividadeSecundaria) Then
                        ValorAlvara = IIf(Rs!TAV_FATOR < 2, Rs!TAV_VALOR_ALVARA, Rs!tci_fator_alvara * Rs!TAV_VALOR_ALVARA)
                    Else
                        ValorAlvara = IIf(Rs!TAV_FATOR < 2, MaxAtividadeSecundaria, Rs!tci_fator_alvara * MaxAtividadeSecundaria)
                    End If
            
                Else
                    ValorAlvara = IIf(Rs!TAV_FATOR < 2, Rs!TAV_VALOR_ALVARA, CDbl(Edita.TrocaPic(Rs!tci_fator_alvara, ".", ",")) * Rs!TAV_VALOR_ALVARA)
                End If
            End If
        End If
        If UCase(AplicacoesVTFuncoes.Municipio) = "PETROLINA" Then
            If Year(RsData!TCI_INICIO_ATIVIDADE) = Periodo Then
                If Val(Month(RsData!TCI_INICIO_ATIVIDADE)) >= 6 Then
                    Alvara = ValorAlvara / 2
                Else
                    Alvara = ValorAlvara
                End If
            End If
        Else
            'CALCULA ALVARA PROPORCIONAL
            If Imposto.Forma_Calculo(CodImposto, Left(Periodo, 4)) = 1 Then
                 If IsNumeric(InicioAtividade) Then
                     ValorAlvara = ValorAlvara * (CInt(InicioAtividade) / 12)
                     Alvara = ValorAlvara
                     If DateDiff("d", RsData(0), DtVenc) < 0 Then
                         DtVenc = UltimoDiaDoMes(IIf(Day(Date) <= 20, Date, Date + 20))
                     End If
                 Else
                     Alvara = ValorAlvara
                 End If
            Else
                Alvara = ValorAlvara
            End If
         End If
        'GRAVA VALOR ALVARA
        If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
            Alvara = ValorAlvara + ValorAlvara * (AtividadesSecundarias / 100)
            'PEGO O DESCONTO DE ACORDO COM O BAIRRO E O FATOR
            Desconto = Pega_Desconto_Alvara(Im)
            If Desconto <> 0 Then
                ValorAlvara = ValorAlvara - (Desconto * ValorAlvara / 100)
                Alvara = ValorAlvara
            End If
            Desconto = Pega_Desconto_Micro_Empresa(Im)
            If Desconto <> 0 Then
                Alvara = Alvara - (Desconto * Alvara / 100)
            End If
        End If
    End If
    
    GravaObrigacaoAcessoria = False
    
End Function
Public Function Pega_Desconto_Micro_Empresa(Contribuinte As String) As Integer
    Dim Sql                                              As String
    Dim Rs                                               As VSRecordset
    
    Sql = "select tci_porte_empresa from tab_contribuinte where tci_im  = '" & Contribuinte & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If CStr(Trim(Rs.Fields("tci_porte_empresa"))) = "1" Then
            Pega_Desconto_Micro_Empresa = Temp.PegaParametro(Bdados, "DESCONTO MICRO")
        End If
    End If
End Function

Public Function CalculaPropocionalidade(InicoAtividade As String, Valor As Double) As String
    Dim Dias As Integer
    Dias = 13 - Month(InicoAtividade)
    If Year(InicoAtividade) = Year(Date) Then
        CalculaPropocionalidade = Valor * (Dias / 12)
    Else
        CalculaPropocionalidade = Valor
    End If
End Function


Public Function Pega_Desconto_Formatura(Im As String, Periodo As String) As Double
    Dim Sql                                As String
    Dim Rs                                 As VSRecordset
    Dim DATA_SOMADA                        As Date
    'PEGO A DATA DE FORMATURA DO CONTRIBUINTE
    Sql = "SELECT * FROM TAB_VARIAVEL_ECONOMICO WHERE TVE_TCI_IM = '" & Im & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Rs.Fields("TVE_DATA_FORMATURA") <> "" Then
            If DateDiff("M", CDate(Rs.Fields("TVE_DATA_FORMATURA")), CDate(Day(Date) & "/" & Month(Date) & "/" & Periodo)) <= 36 Then
                Pega_Desconto_Formatura = Temp.PegaParametro(Bdados, "DESCONTO_ALVARA_FORMATURA")
            End If
        End If
    End If
End Function
Public Function Pega_Desconto_Funcionario_SUS(Im As String) As Double
    Dim Sql                                      As String
    Dim Rs                                       As VSRecordset
    
    Sql = "SELECT * FROM TAB_VARIAVEL_ECONOMICO WHERE TVE_TCI_IM = '" & Im & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Rs.Fields("TVE_FUNCIONARIO_SUS") <> "" Then
            If CStr(Rs.Fields("TVE_FUNCIONARIO_SUS")) = "1" Then
                Pega_Desconto_Funcionario_SUS = Temp.PegaParametro(Bdados, "DESCONTO_ALVARA_FUNCIONARIO_SUS")
            Else
                Pega_Desconto_Funcionario_SUS = 0
            End If
        End If
    End If
End Function
Public Function Pega_Desconto_Alvara(Contribuinte As String) As Double
    Dim Sql                             As String
    Dim Rs                              As VSRecordset
    Dim RsDesconto                      As VSRecordset
    
    'PEGO OS SEGUINTES DADOS NA TAB_CONTRIBUINTE...
    Sql = "SELECT TCI_COD_BAIRRO,TCI_FATOR_ALVARA "
    Sql = Sql & " FROM TAB_CONTRIBUINTE WHERE TCI_IM = '" & Contribuinte & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        'CHECO AS DEFINIÇÕES PARA DESCONTOS...
        'SE O BAIRRO NÃO ESTIVER NESSA TABELA ENTÃO ELE ESTÁ DEFINIDO COMO GERAL
        
        Sql = "SELECT * FROM TAB_DESCONTO_ALVARA WHERE TDA_BAIRRO = '" & Rs.Fields("TCI_COD_BAIRRO") & "'"
        If Bdados.AbreTabela(Sql, RsDesconto) Then
            'TESTO SE O FATOR DO ALVARA DA TAB_CONTRIBUINTE ESTÁ DENTRO DE UM LIMITE INFORMADO...
            Do Until RsDesconto.EOF
                If Rs.Fields("TCI_FATOR_ALVARA") >= RsDesconto.Fields("TDA_VALOR_INICIAL") And Rs.Fields("TCI_FATOR_ALVARA") <= RsDesconto.Fields("TDA_VALOR_FINAL") Then
                    Pega_Desconto_Alvara = RsDesconto.Fields("TDA_PERCENTUAL")
                End If
                RsDesconto.MoveNext
            Loop
        Else
            'PEGO OS LIMITES NA TAB_DESCONTO_ALVARA_GERAL
            Sql = "SELECT * FROM TAB_DESCONTO_ALVARA_GERAL"
            If Bdados.AbreTabela(Sql, RsDesconto) Then
            'TESTO SE O FATOR DO ALVARA DA TAB_CONTRIBUINTE ESTÁ DENTRO DE UM LIMITE INFORMADO...
                Do Until RsDesconto.EOF
                    If Rs.Fields("TCI_FATOR_ALVARA") >= RsDesconto.Fields("TDG_VALOR_INICIAL") And Rs.Fields("TCI_FATOR_ALVARA") <= RsDesconto.Fields("TDG_VALOR_FINAL") Then
                        Pega_Desconto_Alvara = RsDesconto.Fields("TDG_PERCENTUAL")
                    End If
                    RsDesconto.MoveNext
                Loop
            End If
        End If
    End If
End Function
Public Function Calcula_IssFixo_Barra_Mansa(Contribuinte As String) As String
    
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsFator As VSRecordset
    
'    'PEGO A ATIVIDADE DO CONTRIBUINTE...
'    Sql = "SELECT tci_tae_cae FROM TAB_CONTRIBUINTE WHERE TCI_IM  = '" & Contribuinte & "'"
'    If Bdados.AbreTabela(Sql, Rs) Then
'        'Pego o valor da taxa...
'         Sql = "Select TAE_ALIQUOTA_SPL "
'         Sql = Sql & " from TAB_ATIVIDADE_ECONOMICA "
'         Sql = Sql & " WHERE TAE_CAE =  '" & Rs.Fields("tci_tae_cae") & "'"
'         If Bdados.AbreTabela(Sql, Rs) Then
'             Calcula_IssFixo_Barra_Mansa = Rs.Fields("TAE_ALIQUOTA_SPL") * Temp.PegaParametro(Bdados, "UFM")
'         Else
'             Calcula_IssFixo_Barra_Mansa = 0
'         End If
'    Else
'        Calcula_IssFixo_Barra_Mansa = 0
'    End If
    'PEGO O TIPO DE RECOLHIMENTO SE FOR  = 3 FIXO ANUAL ENTÃO O CONTRIBUINTE É UM PROFISSIONAL LIBERAL...
    
    Sql = "SELECT * FROM TAB_CONTRIBUINTE WHERE TCI_IM = '" & Contribuinte & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If CStr(Rs.Fields("TCI_TIPO_RECOLHIMENTO_ISS")) = "4" Then 'RECOLHIMENTO FIXO ANUAL
            'CHECO O GRAU DE ESTUDO DO CONTRBUINTE...
            If CStr(Pega_NivelEscolar(Rs.Fields("TCI_TAE_CAE"))) = "1" Then    '3º GRAU
                'CHECO SE O CONTRIBUINTE É ESTABELECIDO...
                If CStr(Rs.Fields("TCI_ESTAB")) = "1" Then
                   'PEGO O VALOR DO FATOR...
                    Sql = "SELECT * FROM VIS_VALORES_ALVARA WHERE TGE_CODIGO = 1"
                    If Bdados.AbreTabela(Sql, RsFator) Then
                        Calcula_IssFixo_Barra_Mansa = RsFator.Fields("TGE_NOME") * Temp.PegaParametro(Bdados, "UFM")
                    End If
                Else
                    Sql = "SELECT * FROM VIS_VALORES_ALVARA WHERE TGE_CODIGO = 2"
                    If Bdados.AbreTabela(Sql, RsFator) Then
                        Calcula_IssFixo_Barra_Mansa = RsFator.Fields("TGE_NOME") * Temp.PegaParametro(Bdados, "UFM")
                    End If
                End If
            ElseIf CStr(Pega_NivelEscolar(Rs.Fields("TCI_TAE_CAE"))) = "2" Then '2º GRAU
                'CHECO SE O CONTRIBUINTE É ESTABELECIDO...
                If CStr(Rs.Fields("TCI_ESTAB")) = "1" Then
                    'PEGO O VALOR DO FATOR...
                    Sql = "SELECT * FROM VIS_VALORES_ALVARA WHERE TGE_CODIGO = 3"
                    If Bdados.AbreTabela(Sql, RsFator) Then
                        Calcula_IssFixo_Barra_Mansa = RsFator.Fields("TGE_NOME") * Temp.PegaParametro(Bdados, "UFM")
                    End If
                Else
                    Sql = "SELECT * FROM VIS_VALORES_ALVARA WHERE TGE_CODIGO = 4"
                    If Bdados.AbreTabela(Sql, RsFator) Then
                        Calcula_IssFixo_Barra_Mansa = RsFator.Fields("TGE_NOME") * Temp.PegaParametro(Bdados, "UFM")
                    End If
                End If
            End If
        End If
    End If
End Function
Public Function Pega_NivelEscolar(Atividade As String) As String
    Dim Sql As String
    
    Sql = "Select * from Tab_Atividade_Economica where TAE_CAE  = '" & Atividade & "'"
    If Bdados.AbreTabela(Sql) Then
        Pega_NivelEscolar = Bdados.Tabela("TAE_NIVEL_ESCOLAR")
    Else
        Pega_NivelEscolar = ""
    End If
End Function
Public Function Calcula_Taxas(Contribuinte As String, Taxa As String, Periodo As String) As String
    Dim Sql                                As String
    Dim Rs                                 As VSRecordset
    Dim RsTaxa                             As VSRecordset
    Dim Mutiplicador                       As Double
    Dim RsMutiplicador                     As VSRecordset
    Dim ValorReal                          As Double
    Dim ValorBase                          As Double
    
    'PEGO A ATIVIDADE DO CONTRIBUINTE...
    Sql = "SELECT tci_tae_cae,TCI_INICIO_ATIVIDADE FROM TAB_CONTRIBUINTE WHERE TCI_IM  = '" & Contribuinte & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
    'Pego o mutiplicador
    Sql = "Select * from tab_variavel_economico where tve_tci_im = '" & Contribuinte & "'"
    If Bdados.AbreTabela(Sql, RsMutiplicador) Then
        If Taxa = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_TFA)) Then
            Mutiplicador = RsMutiplicador.Fields("tve_anuncios")
        ElseIf Taxa <> Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_tfs)) Then
            Mutiplicador = RsMutiplicador.Fields("tve_quantidade_item")
        End If
    End If
    'Pego o valor da taxa...
     Sql = "Select tta_valor_real "
     Sql = Sql & " from tab_taxa_atividade,tab_contribuinte "
     Sql = Sql & " where  tci_tae_cae  =  tta_tae_codigo  "
     Sql = Sql & " and tta_tae_codigo =  '" & Rs.Fields("tci_tae_cae") & "'"
     Sql = Sql & " and tci_im = '" & Contribuinte & "'"
     Sql = Sql & " and tta_codigo_taxa  = '" & Taxa & "'"
     If Bdados.AbreTabela(Sql, RsTaxa) Then
        If Mutiplicador <> 0 Then
            ValorBase = RsTaxa.Fields("tta_valor_real") * Mutiplicador
        Else
            ValorBase = RsTaxa.Fields("tta_valor_real")
        End If
        ValorReal = ValorBase
        If Imposto.Forma_Calculo(Taxa, Left(Periodo, 4)) = 1 Then
            ValorReal = CStr(CalculaPropocionalidade(Rs.Fields("TCI_INICIO_ATIVIDADE"), ValorBase))
        End If
        Calcula_Taxas = ValorReal
     Else
         Calcula_Taxas = 0
     End If
    Else
        Calcula_Taxas = 0
    End If
End Function
Public Function PublicidadeAnuncio(Im As String, CodImposto As String, Optional DocOrigem As String, Optional ByRef DtVenc As String)
    Dim Sql As String
    
    Sql = "Select TAN_VALOR AS TOTAL,TAN_DATA_INSTALACAO from tab_anuncio where TAN_TCI_IM = '" & Im & "' and TAN_MOVIMENTO = '" & CodImposto & "'"
    If Trim(DocOrigem) <> "" Then Sql = Sql & " and  TAN_DOC_ORIGEM ='" & DocOrigem & "'"
    If Bdados.AbreTabela(Sql) Then
        PublicidadeAnuncio = Bdados.Tabela("TOTAL")
        If DateDiff("d", Bdados.Tabela("TAN_DATA_INSTALACAO"), DtVenc) < 0 Then
            DtVenc = UltimoDiaDoMes(IIf(Day(Date) <= 20, Date, Date + 20))
        End If
    Else
        PublicidadeAnuncio = 0
    End If
End Function
Public Function AlvaraPetrolina(Im As String, Periodo As String, CodImposto As String, ByRef DtVenc As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim InicioAtividade As String
    Dim ValorAlvara As Double
    Dim RsData As VSRecordset
    Dim CAE As String
    Dim FatorAlvara As Double
    Dim AtividadesSecundarias As Integer
    Dim PercAtivSecund As Double
    Dim strSiglaAlvara As String
    Dim dblValorExcedente As Double
    'DEVE SER MUDADO, ANTES DE RODAR A 1ª VEZ EM PETROLINA -
    'ALTERACAO DA TABELA DE ATIVIDADE POR ANDRE EM 08-01-2006(COLINAS/MA)
    
    strSiglaAlvara = Imposto.NomeTributo(ttr_ALVARA)
    
    GravaObrigacao = True
    Sql = "select tci_inicio_atividade,tci_tae_cae,tci_fator_alvara,tci_tae_cae_secund," & _
        "tci_tae_cae_terc from tab_contribuinte where tci_im ='" & Im & "'"
    If Bdados.AbreTabela(Sql, RsData) Then
        If IsNull(RsData(0)) Then Exit Function
        'PEGA PARAMETROS
        InicioAtividade = RsData!TCI_INICIO_ATIVIDADE
        If Periodo < Year(InicioAtividade) Then
            GravaObrigacao = False
            Exit Function
        End If
        CAE = RsData!tci_tae_cae
        FatorAlvara = Nvl("" & RsData!tci_fator_alvara, 1)
        If Format(Year(InicioAtividade), "0000") = Left(Periodo, 4) Then
            InicioAtividade = 13 - Month(InicioAtividade)
        End If
        PercAtivSecund = CDbl(Nvl(BuscaNaGeral("ALVARA SECUNDARIO", 1), 0))
        AtividadesSecundarias = PercAtivSecund * CDbl(IIf(Nvl("" & RsData!tci_tae_cae_secund, 0) = 0, 0, 1) + IIf(Nvl("" & RsData!tci_tae_cae_terc, 0) = 0, 0, 1))
        
        DtVenc = Imposto.BuscaDataVencimento(CodImposto, CLng(Periodo))
        
        If AplicacoesVTFuncoes.Municipio = "ARARIPINA" Or AplicacoesVTFuncoes.Municipio = "SANTA MARIA DA BOA VISTA" Or AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE" Or Temp.PegaParametro(Bdados, "ALVARA IPM") = "SIM" Then
            'VERIFICA SE ATIVIDADE É ESTIMATIVA
            'ESTIMATIVA GERAL PARA TODAS AS ATIVIDADES
            If Temp.PegaParametro(Bdados, "ALVARA ESTIMATIVA GERAL") = "SIM" Then
                Sql = "SELECT   TAT_LIMITE_INFERIOR,TAT_LIMITE_SUPERIOR,TAT_LIMITE_VALOR,TAT_LIMITE_FATOR, TAT_VALOR_EXCEDENTE FROM TAB_ATIVIDADE_ESTIMADA WHERE TAT_TAE_CAE = 0 AND TAT_ANO =" & Periodo & " order by TAT_LIMITE_INFERIOR asc"
            Else
                Sql = "SELECT   TAT_LIMITE_INFERIOR,TAT_LIMITE_SUPERIOR,TAT_LIMITE_VALOR,TAT_LIMITE_FATOR, TAT_VALOR_EXCEDENTE FROM TAB_ATIVIDADE_ESTIMADA WHERE TAT_TAE_CAE = " & CAE & " AND TAT_ANO =" & Periodo & " order by TAT_LIMITE_INFERIOR asc"
            End If
            If Bdados.AbreTabela(Sql, Rs) Then
                Rs.MoveFirst
                Do
                    If Nvl("" & Rs!TAT_LIMITE_SUPERIOR, 0) = 0 Then 'ULTIMA FAIXA
                        If FatorAlvara >= Rs!TAT_LIMITE_INFERIOR Then
                            If Rs!TAT_LIMITE_FATOR <> 0 Then
                                ValorAlvara = FatorAlvara * Rs!TAT_LIMITE_VALOR
                            Else
                                ValorAlvara = 1 * Rs!TAT_LIMITE_VALOR
                            End If
                            'ValorAlvara = IIf(Rs!TAT_LIMITE_FATOR <> 0, FatorAlvara, 1) * Rs!TAT_LIMITE_VALOR
                            
                            'CALCULA O VALOR EXCEDENTE
                            If Nvl("" & Rs!TAT_VALOR_EXCEDENTE, 0) > 0 Then
                                If Rs!TAT_LIMITE_INFERIOR > 0 Then
                                    If FatorAlvara Mod Rs!TAT_LIMITE_INFERIOR = 0 Then
                                        dblValorExcedente = ((FatorAlvara / Rs!TAT_LIMITE_INFERIOR) - 1) * Rs!TAT_VALOR_EXCEDENTE
                                    Else
                                        dblValorExcedente = CInt((FatorAlvara / Rs!TAT_LIMITE_INFERIOR) - 0.5) * Rs!TAT_VALOR_EXCEDENTE
                                    End If
                                    ValorAlvara = ValorAlvara + dblValorExcedente
                                End If
                            End If
                            
                        End If
                    ElseIf FatorAlvara >= Rs!TAT_LIMITE_INFERIOR And FatorAlvara <= Rs!TAT_LIMITE_SUPERIOR Then
                        If CDbl(Nvl(Trim(Rs!TAT_LIMITE_FATOR), 0)) <> 0 Then
                            ValorAlvara = FatorAlvara * Rs!TAT_LIMITE_VALOR
                        Else
                            ValorAlvara = 1 * Rs!TAT_LIMITE_VALOR
                        End If
                        'ValorAlvara = IIf(CDbl(Nvl(Trim(Rs!TAT_LIMITE_FATOR), 0)) <> 0, FatorAlvara, 1) * Rs!TAT_LIMITE_VALOR
                        
                        'CALCULA O VALOR EXCEDENTE
                        If Nvl("" & Rs!TAT_VALOR_EXCEDENTE, 0) > 0 Then
                            If Rs!TAT_LIMITE_INFERIOR > 0 Then
                                If FatorAlvara Mod Rs!TAT_LIMITE_INFERIOR = 0 Then
                                    dblValorExcedente = ((FatorAlvara / Rs!TAT_LIMITE_INFERIOR) - 1) * Rs!TAT_VALOR_EXCEDENTE
                                Else
                                    dblValorExcedente = CInt((FatorAlvara / Rs!TAT_LIMITE_INFERIOR) - 0.5) * Rs!TAT_VALOR_EXCEDENTE
                                End If
                                ValorAlvara = ValorAlvara + dblValorExcedente
                            End If
                        End If
                        Exit Do
                    End If
                    Rs.MoveNext
                Loop While Not Rs.EOF
                Rs.Fechar
            Else
                Sql = "Select tci_fator_alvara,TAV_VALOR_ALVARA,"
                Sql = Sql & " tip_cod_imposto,TAT_FATOR "
                Sql = Sql & " from Tab_Contribuinte, TAB_ATIVIDADE_VALOR_TRIBUTOS,Tab_Imposto "
                Sql = Sql & " where tci_im  = '" & Im & "' and tci_tae_cae = TAV_TAE_CAE and tip_sigla_imposto='" & strSiglaAlvara & _
                        "' AND TAV_ANO =" & Periodo
                If Bdados.AbreTabela(Sql, Rs) Then
                    If Rs!tci_fator_alvara > 0 Then
                        ValorAlvara = IIf(Rs!TAV_VALOR_ALVARA = 0, Rs!TAV_VALOR_ALVARA, Rs!TAT_FATOR * Rs!TAV_VALOR_ALVARA)
                    Else
                        ValorAlvara = Rs!TAV_VALOR_ALVARA
                    End If
                End If
            End If
        End If
        'CALCULA ALVARA PROPORCIONAL
        If UCase(AplicacoesVTFuncoes.Municipio) = "PETROLINA" Then
            If Year(RsData!TCI_INICIO_ATIVIDADE) = Periodo Then
                If Val(Month(RsData!TCI_INICIO_ATIVIDADE)) >= 6 Then
                    AlvaraPetrolina = ValorAlvara / 2
                Else
                    AlvaraPetrolina = ValorAlvara
                End If
            End If
        Else
            If IsNumeric(InicioAtividade) Then
                ValorAlvara = ValorAlvara * (CInt(InicioAtividade) / 12)
                AlvaraPetrolina = ValorAlvara
                If DateDiff("d", RsData(0), DtVenc) < 0 Then
                    DtVenc = UltimoDiaDoMes(IIf(Day(Date) <= 20, Date, Date + 20))
                End If
            Else
                AlvaraPetrolina = ValorAlvara
            End If
        End If
        'GRAVA VALOR ALVARA
        AlvaraPetrolina = ValorAlvara + ValorAlvara * (AtividadesSecundarias / 100)
    End If
    TipoLancamento = etlHomologacao
    GravaObrigacaoAcessoria = False
    
End Function
Public Function GeraIssEstimativo(Im As String, Periodo As String) As Double
    Dim Sql                          As String
    Dim Rs                           As VSRecordset
    
    Sql = "Select *  from tab_contribuinte_estimado "
    Sql = Sql & " where tce_tci_im = '" & Im & "'"
    Sql = Sql & " and  tce_exercicio = '" & Left(Periodo, 4) & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Rs.Fields("TCE_STATUS") <> 2 Then
            GeraIssEstimativo = Rs.Fields("TCE_VALOR_MENSAL")
        Else
            Avisa "Contribuinte com a situação cancelada do regime estimativo."
        End If
    End If
End Function

Public Function Issqn(Im As String, Periodo As String, CodImposto As String, ByVal DtVenc As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Decl As New cDeclaracao
    
    Sql = "select tci_inicio_atividade,tci_tae_cae,"
    Sql = Sql & "  tae_valor,tci_tipo_recolhimento_iss ,tci_tga_cod_grupo "
    Sql = Sql & "  from tab_contribuinte,"
    Sql = Sql & " tab_atividade_economica "
    Sql = Sql & " where (tci_im ='" & Im & "' OR TCI_CGC_CPF ='" & Im & "') "
    Sql = Sql & " and tci_tae_cae= tae_cae"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Len(CStr(Periodo)) = 4 Then
            GravaObrigacao = IIf(Periodo < CDbl(Format(Rs!TCI_INICIO_ATIVIDADE, "YYYY")), False, True)
        Else
            GravaObrigacao = IIf(Periodo < CDbl(Format(Rs!TCI_INICIO_ATIVIDADE, "YYYYMM")), False, True)
        End If
        If GravaObrigacao Then
'            If Rs!tci_tipo_recolhimento_iss + 1 = 1 Or Rs!tci_tipo_recolhimento_iss + 1 = 3 Then 'RECOLHIMENTO MENSAL
                If Decl.Buscar(Im, Right(Periodo, 2) & Left(Periodo, 4), , CodImposto) Then
                    Issqn = Decl.ImpostoDevido
                    TipoLancamento = etlDeclaracao
                Else
                    'If Rs!tci_tga_cod_grupo <> 4 Then
                    '    Util.Avisa "O grupo de atividade deste contribuinte não é Prestação de Serviço. A obrigação não pode ser gerada."
                    '    GravaObrigacao = False
                    'Else
                        TipoLancamento = etlDeclaracao
                    'End If
                End If
'            ElseIf Rs!tci_tipo_recolhimento_iss + 1 = 2 Then 'RECOLHIMENTO ANUAL
'                Issqn = Rs!tae_valor
'                TipoLancamento = etlHomologacao
'            End If
        Else
            Avisa "Periodo de referencia incompativel com inicio de atividade em " & Rs!TCI_INICIO_ATIVIDADE & "."
                'LUCAS - 25/04/2004
        End If
    Else
        If Decl.Buscar(Im, Right(Periodo, 2) & Left(Periodo, 4), , CodImposto) Then
            Issqn = Decl.ImpostoDevido
            TipoLancamento = etlDeclaracao
        Else
            TipoLancamento = etlDeclaracao
        End If
    End If
    GravaObrigacaoAcessoria = False
End Function



Public Function Iptu(Imovel As String, Periodo As String, CodImposto As String, ByRef DtVenc As String, Optional GeraSomenteValorVenal As Boolean = False, Optional TaxasInclusas As Double) As Double
'SQz - Balsas 07062003
    Dim CalculoIptu As New VSIptu
    Dim rsCob As VSRecordset, Sql As String
    On Error GoTo Trata
'    CalculoIptu.AnoLancamento = Periodo
'    CalculoIptu.InicializarValores Imovel
'    CalculoIptu.GeraCotaUnica = True
'    CalculoIptu.GeraParcelas = False
'    GravaObrigacaoAcessoria = False
'    Set DadosIptu = Nothing
'    GravaObrigacao = False
'    If CalculoIptu.CarregaDetalheLote(Trim(Imovel)) Then
'        Iptu = CalculoIptu.CalculaValorIptu
'        Set DadosIptu = CalculoIptu
'        GravaObrigacao = True
'        GravaObrigacaoAcessoria = True
'    End If
'    TipoLancamento = etlHomologacao

'*****************************
    
    If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
        ' SE O ANO FOR MENOR QUE O DO PGV, GERA COM O MODELO PADRÃO ANTIGO DO IPTU (padrão bci)
        If CDbl(Periodo) < CDbl(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
            
            Sql = "Select * from tab_imovel where tim_ic ='" & Imovel & "'"
            Bdados.AbreTabela Sql, rsCob
            Imposto.ValorTaxas = 0
            Iptu = Imposto.GeraIptu(cip_Balsas, rsCob, CInt(Periodo), CInt(Periodo), tgi_SemParcelas)
            Iptu = Iptu + Imposto.ValorTaxas
            DtVenc = Imposto.DataVenc
            GravaObrigacaoAcessoria = True: TaxaUnificada = True
            
            Exit Function
        End If
        If CDbl(Periodo) >= CDbl(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
            CalculoIptu.AnoLancamento = Periodo
            CalculoIptu.InicializarValores Imovel
        End If
        
        ' SE O ANO FOR MENOR QUE O DO PGV, GERA COM O MODELO PADRÃO ANTIGO DO IPTU
        If CDbl(Periodo) < CDbl(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
            Iptu = Imposto.GeraIptu(cip_Balsas, rsCob, CInt(Periodo), CInt(Periodo), tgi_SemParcelas)
        Else
            ' SE O ANO FOR COMPATÍVEL COM O PGV GERA ENTÃO O IPTU PGV
            If CalculoIptu.CarregaDetalheLote(Trim(Imovel)) Then
                Iptu = CalculoIptu.CalculaValorIptu(GeraSomenteValorVenal)
                GravaObrigacaoAcessoria = True
                
                Set DadosIptu = CalculoIptu
            End If
        End If
    
    Else
        Sql = "SELECT * FROM  Vis_Imovel ,Tab_Contribuinte where tim_tci_im=tci_im AND TIM_SITUACAO_LOTE <> 1"
        Sql = Sql & " AND TBA_TMU_COD_MUNICIPIO = " & AplicacoesVTFuncoes.Codigo_Municipio & " AND tlg_tmu_cod_municipio = " & AplicacoesVTFuncoes.Codigo_Municipio
        Sql = Sql & " and tim_ic ='" & Imovel & "'"
        
        Sql = Sql & " order by tim_ic ASC ,tim_unidade ASC"
        Bdados.AbreTabela Sql, rsCob
        Iptu = Imposto.CriaIptu(rsCob, CInt(Periodo), CInt(Periodo))
        TaxasInclusas = Imposto.ValorTaxas
        DtVenc = Imposto.BuscaDataVencimento(CodImposto, CStr(Periodo))
        Exit Function
    End If
    DtVenc = CalculoIptu.VencimentoCotaUnica
    Exit Function
Trata:
    Avisa Err.Description
    Exit Function
    Resume
End Function

Public Function IptuSinfaz(Imovel As String, Periodo As String, CodImposto As String, ByRef DtVenc As String, Optional GeraSomenteValorVenal As Boolean = False) As Double
'SQz - Balsas 07062003
    Dim CalculoIptu As New VsIptuPetrolina
    Dim rsCob As VSRecordset, Sql As String

    
    CalculoIptu.AnoLancamento = Periodo
    CalculoIptu.InicializaValores Imovel
    
    If CalculoIptu.CarregaDetalheLote(Trim(Imovel)) Then
        IptuSinfaz = CalculoIptu.CalculaValorIptu(GeraSomenteValorVenal)
    End If
    DtVenc = CalculoIptu.VencimentoCotaUnica
End Function

Public Function IptuBarra(Imovel As String, Periodo As String, CodImposto As String, ByRef DtVenc As String, Optional GeraSomenteValorVenal As Boolean = False) As Double
    Dim CalculoIptu As New VsIptuBarraMansa
    Dim rsCob As VSRecordset, Sql As String
    
    CalculoIptu.AnoLancamento = Periodo
    If CalculoIptu.CarregaDetalheLote(Trim(Imovel)) Then
        IptuBarra = CalculoIptu.CalculaValorIptu(GeraSomenteValorVenal)
    End If
    'DtVenc = CalculoIptu.VencimentoCotaUnica
End Function
Public Function TaxaFixa(Contribuinte As String, CodImposto As String, Ano As String, Optional multiplicador As Double) As Double
    Dim Imposto As New VSImposto
    Dim Sql As String
    Dim Rs As VSRecordset
    If CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_TEMISSAOALVARA)) Then
        Sql = "SELECT TCI_ALVARA_LIBERADO FROM TAB_CONTRIBUINTE WHERE TCI_IM = '" & Contribuinte & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            If Nvl("" & Rs!TCI_ALVARA_LIBERADO, 0) = 0 Then
                Bdados.AtualizaDados "TAB_CONTRIBUINTE", Bdados.PreparaValor(1), "TCI_ALVARA_LIBERADO", "TCI_IM = '" & Contribuinte & "'"
            Else
                Exit Function
            End If
        End If
    End If
    If CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_PUBLICIDADE)) And multiplicador = 0 Then
        Sql = "SELECT tci_fator_alvara FROM TAB_CONTRIBUINTE WHERE TCI_IM = '" & Contribuinte & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            multiplicador = Nvl("" & Rs!tci_fator_alvara, 0)
        End If
    End If
    
    TaxaFixa = Format(Imposto.BuscaValorFixoTaxa(CodImposto, Ano) * IIf(multiplicador > 0, multiplicador, 1), Const_Monetario)

End Function


Public Function IssqnFixo(Im As String, Periodo As String, CodImposto As String, ByVal DtVenc As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Decl As New cDeclaracao
    

    Sql = "select tci_inicio_atividade,"
    Sql = Sql & " tci_tae_cae,"
    Sql = Sql & " TAE_ALIQUOTA_SPL,"
    Sql = Sql & " tci_tipo_recolhimento_iss ,"
    Sql = Sql & " tci_tga_cod_grupo "
    Sql = Sql & " from tab_contribuinte,tab_atividade_economica "
    Sql = Sql & " where (tci_im ='" & Im & "' OR TCI_CGC_CPF ='" & Im & "') and tci_tae_cae= tae_cae"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Len(CStr(Periodo)) = 4 Then
            GravaObrigacao = IIf(Periodo < CDbl(Format(Rs!TCI_INICIO_ATIVIDADE, "YYYY")), False, True)
        Else
            GravaObrigacao = IIf(Periodo < CDbl(Format(Rs!TCI_INICIO_ATIVIDADE, "YYYYMM")), False, True)
        End If
        If GravaObrigacao Then
            If UCase(AplicacoesVTFuncoes.Municipio) = "PETROLINA" Then
                If Year(Rs!TCI_INICIO_ATIVIDADE) = Periodo Then
                    If Val(Month(Rs!TCI_INICIO_ATIVIDADE)) > 6 Then
                        IssqnFixo = Rs!TAE_ALIQUOTA_SPL * TrocaPic(Temp.PegaParametro(Bdados, "UFM"), ".", ",") / 2
                    Else
                        IssqnFixo = Rs!TAE_ALIQUOTA_SPL * TrocaPic(Temp.PegaParametro(Bdados, "UFM"), ".", ",")
                    End If
                Else
                    IssqnFixo = Rs!TAE_ALIQUOTA_SPL * TrocaPic(Temp.PegaParametro(Bdados, "UFM"), ".", ",")
                End If
            Else
                IssqnFixo = Rs!TAE_ALIQUOTA_SPL
            End If
            TipoLancamento = etlHomologacao
        Else
            Avisa "Periodo de referencia incompativel com inicio de atividade em " & Rs!TCI_INICIO_ATIVIDADE & "."
        End If
    End If
    GravaObrigacaoAcessoria = False
        
End Function

