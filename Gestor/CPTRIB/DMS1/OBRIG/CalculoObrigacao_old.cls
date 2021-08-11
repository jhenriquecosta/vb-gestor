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
Public Function Alvara(IM As String, Periodo As String, CodImposto As String, ByVal DtVenc As String) As Double
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
    
    GravaObrigacao = True
    Sql = "select tci_inicio_atividade,tci_tae_cae,tci_fator_alvara,tci_tae_cae_secund," & _
        "tci_tae_cae_terc from tab_contribuinte where tci_im ='" & IM & "'"
    If Bdados.AbreTabela(Sql, RsData) Then
        If IsNull(RsData(0)) Then Exit Function
        'PEGA PARAMETROS
        InicioAtividade = RsData!tci_inicio_atividade
        If Periodo < Year(InicioAtividade) Then
            GravaObrigacao = False
            Exit Function
        End If
        CAE = RsData!tci_tae_cae
        FatorAlvara = RsData!tci_fator_alvara
        If Format(Year(InicioAtividade), "0000") = Left(Periodo, 4) Then
            InicioAtividade = 13 - Month(InicioAtividade)
        End If
        PercAtivSecund = CDbl(Nvl(BuscaNaGeral("ALVARA SECUNDARIO", 1), 0))
        AtividadesSecundarias = PercAtivSecund * CDbl(IIf(Nvl("" & RsData!tci_tae_cae_secund, 0) = 0, 0, 1) + IIf(Nvl("" & RsData!tci_tae_cae_terc, 0) = 0, 0, 1))
        DtVenc = Imposto.BuscaDataVencimento(CodImposto, CLng(Periodo))
        'VERIFICA SE ATIVIDADE É ESTIMATIVA
        Sql = "SELECT   TAT_LIMITE_INFERIOR,TAT_LIMITE_SUPERIOR,TAT_LIMITE_VALOR,TAT_LIMITE_FATOR FROM TAB_ATIVIDADE_ESTIMADA WHERE TAT_TAE_CAE = " & CAE & " order by TAT_LIMITE_INFERIOR asc"
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
            Sql = "Select tci_fator_alvara,tae_valor,tip_cod_imposto,tae_fator from Tab_Contribuinte, tab_atividade_economica,Tab_Imposto where " & IIf(Len(IM) = 11, "tci_im", "tci_cgc_cpf") & " = '" & IM & "' and tci_tae_cae = tae_cae and tip_sigla_imposto='ALVARA'"
            If Bdados.AbreTabela(Sql, Rs) Then
                ValorAlvara = IIf(Rs!tae_fator = 0, Rs!tae_valor, Rs!tci_fator_alvara * Rs!tae_valor)
            End If
        End If
        'CALCULA ALVARA PROPORCIONAL
        If IsNumeric(InicioAtividade) Then
            ValorAlvara = ValorAlvara * (CInt(InicioAtividade) / 12)
            Alvara = ValorAlvara
            If DateDiff("d", RsData(0), DtVenc) < 0 Then
                DtVenc = UltimoDiaDoMes(IIf(Day(Date) <= 20, Date, Date + 20))
            End If
        Else
            Alvara = ValorAlvara
        End If
        'GRAVA VALOR ALVARA
        Alvara = ValorAlvara + ValorAlvara * (AtividadesSecundarias / 100)
    End If
    TipoLancamento = etlHomologacao
    GravaObrigacaoAcessoria = False
    
End Function



Public Function Issqn(IM As String, Periodo As String, CodImposto As String, ByVal DtVenc As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Decl As New cDeclaracao
    
    Sql = "select tci_inicio_atividade,tci_tae_cae,tae_valor,tci_tipo_recolhimento_iss +1 as tci_tipo_recolhimento_iss " & _
            "from tab_contribuinte,tab_atividade_economica where tci_im ='" & IM & "' and tci_tae_cae= tae_cae"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Len(CStr(Periodo)) = 4 Then
            GravaObrigacao = IIf(Periodo < CDbl(Format(Rs!tci_inicio_atividade, "YYYY")), False, True)
        Else
            GravaObrigacao = IIf(Periodo < CDbl(Format(Rs!tci_inicio_atividade, "YYYYMM")), False, True)
        End If
        If GravaObrigacao Then
            If Rs!tci_tipo_recolhimento_iss = 1 Or Rs!tci_tipo_recolhimento_iss = 3 Then 'RECOLHIMENTO MENSAL
                If Decl.Buscar(IM, Periodo) Then
                    Issqn = Decl.ImpostoDevido
                    TipoLancamento = etlDeclaracao
                End If
            ElseIf Rs!tci_tipo_recolhimento_iss = 2 Then 'RECOLHIMENTO ANUAL
                Issqn = Rs!tae_valor
                TipoLancamento = etlHomologacao
            End If
        End If
    End If
    GravaObrigacaoAcessoria = False
End Function



Public Function Iptu(Imovel As String, Periodo As String, CodImposto As String, ByRef DtVenc As String) As Double
'SQz - Balsas 07062003
    Dim CalculoIptu As New VSIptu
    Dim rsCob As VSRecordset, Sql As String

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
        If CInt(Periodo) < CInt(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
            Sql = "SELECT * FROM  Vis_Imovel ,Tab_Contribuinte where tim_tci_im=tci_im AND TIM_SITUACAO_LOTE <> 1"
            Sql = Sql & " AND TBA_TMU_COD_MUNICIPIO = " & AplicacoesVTFuncoes.Codigo_Municipio & " AND tlg_tmu_cod_municipio = " & AplicacoesVTFuncoes.Codigo_Municipio
            If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
                If CInt(Right(Imovel, 3)) <> 200 Then
                    Sql = Sql & " and tim_ic ='" & Imovel & "'"
                Else
                    Sql = Sql & " and tim_ic > '" & Imovel & "' and  tim_ic  <= '" & Left(Imovel, 12) & "300'"
                End If
            Else
                Sql = Sql & " and tim_ic ='" & Imovel & "'"
            End If
            Sql = Sql & " order by tim_ic ASC ,tim_unidade ASC"
            Bdados.AbreTabela Sql, rsCob
            
            Imposto.ValorTaxas = 0
            Iptu = Imposto.GeraIptu(cip_Balsas, rsCob, CInt(Periodo), CInt(Periodo), tgi_SemParcelas)
            Iptu = Iptu + Imposto.ValorTaxas
            DtVenc = Imposto.DataVenc
            GravaObrigacaoAcessoria = True: TaxaUnificada = True
            
            Exit Function
        End If
'        rsCob.MoveFirst
        If CInt(Periodo) >= CInt(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
            CalculoIptu.AnoLancamento = Periodo
            CalculoIptu.InicializarValores Imovel
        End If
        ' SE O ANO FOR MENOR QUE O DO PGV, GERA COM O MODELO PADRÃO ANTIGO DO IPTU
        If CInt(Periodo) < CInt(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
            Iptu = Imposto.GeraIptu(cip_Balsas, rsCob, CInt(Periodo), CInt(Periodo), tgi_SemParcelas)
        Else
            ' SE O ANO FOR COMPATÍVEL COM O PGV GERA ENTÃO O IPTU PGV
            If CalculoIptu.CarregaDetalheLote(Trim(Imovel)) Then
                Iptu = CalculoIptu.CalculaValorIptu()
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
        DtVenc = Imposto.BuscaDataVencimento(CodImposto, CStr(Periodo))
        Exit Function
    End If
    DtVenc = CalculoIptu.VencimentoCotaUnica
End Function
