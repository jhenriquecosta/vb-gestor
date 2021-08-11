VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Calculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function Alvara(Im As String, Periodo As Integer, Taxa As Double, CodImposto As String, ByVal DtVenc As String, CodPagamento As Double, CodTaxa As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim DataInicio As String
    Dim ValorAlvara As Double
    Dim RsData As VSRecordset
    Dim CAE As String
    Dim FatorAlvara As Double
    Dim AtividadesSecundarias As Integer
    Dim PercAtivSecund As Double
    
    Sql = "select tci_inicio_atividade,tci_tae_cae,tci_fator_alvara,tci_tae_cae_secund," & _
        "tci_tae_cae_terc from tab_contribuinte where tci_im ='" & Im & "'"
    If Bdados.AbreTabela(Sql, RsData) Then
        If IsNull(RsData(0)) Then Exit Function
        'PEGA PARAMETROS
        DataInicio = RsData!tci_inicio_atividade
        CAE = RsData!tci_tae_cae
        FatorAlvara = RsData!tci_fator_alvara
        If Format(Year(DataInicio), "0000") = Left(CStr(Periodo), 4) Then
            DataInicio = 13 - Month(DataInicio)
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
            Sql = "Select tci_fator_alvara,tae_valor,tip_cod_imposto,tae_fator from Tab_Contribuinte, tab_atividade_economica,Tab_Imposto where " & IIf(Len(Im) = 11, "tci_im", "tci_cgc_cpf") & " = '" & Im & "' and tci_tae_cae = tae_cae and tip_sigla_imposto='ALVARA'"
            If Bdados.AbreTabela(Sql, Rs) Then
                ValorAlvara = IIf(Rs!tae_fator = 0, Rs!tae_valor, Rs!tci_fator_alvara * Rs!tae_valor)
            End If
        End If
        'CALCULA ALVARA PROPORCIONAL
        If IsNumeric(DataInicio) Then
            ValorAlvara = ValorAlvara * (CInt(DataInicio) / 12)
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
End Function
