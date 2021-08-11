VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iCertidao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_CodNegativa As Double
Private m_Im As String
Private m_DataNegativa As String
Private m_Ic As String
Private m_Finalidade As String
Private m_Observacao As String
Private m_Validade As String
Private m_PeriodoInicial As String
Private m_PeriodoFinal As String
Private m_CodUsuario As String
Private m_VencimentoFinal As String
Private m_VencimentoInicial As String
Private m_Tipo As Integer
Private m_Imposto As String
Public Enum TipoCertidao

    tcCND = 60
    tcCPD = 90
    tcCPND = 30
    tcCNE = 66
    
End Enum

Public Property Get Tipo() As Integer
    Tipo = m_Tipo
End Property

Public Property Let Tipo(ByVal Value As Integer)
    m_Tipo = Value
End Property

Public Property Get VencimentoInicial() As String
    VencimentoInicial = m_VencimentoInicial
End Property

Public Property Let VencimentoInicial(ByVal Value As String)
    m_VencimentoInicial = Value
End Property

Public Property Get VencimentoFinal() As String
    VencimentoFinal = m_VencimentoFinal
End Property

Public Property Let VencimentoFinal(ByVal Value As String)
    m_VencimentoFinal = Value
End Property

Public Property Get CodUsuario() As String
    CodUsuario = m_CodUsuario
End Property

Public Property Let CodUsuario(ByVal Value As String)
    m_CodUsuario = Value
End Property

Public Property Get PeriodoFinal() As String
    PeriodoFinal = m_PeriodoFinal
End Property

Public Property Let PeriodoFinal(ByVal Value As String)
    m_PeriodoFinal = Value
End Property

Public Property Get Imposto() As String
    Imposto = m_Imposto
End Property

Public Property Let Imposto(ByVal Value As String)
    m_Imposto = Value
End Property

Public Property Get PeriodoInicial() As String
    PeriodoInicial = m_PeriodoInicial
End Property

Public Property Let PeriodoInicial(ByVal Value As String)
    m_PeriodoInicial = Value
End Property

Public Property Get Validade() As String
    Validade = m_Validade
End Property

Public Property Let Validade(ByVal Value As String)
    m_Validade = Value
End Property

Public Property Get Finalidade() As String
    Finalidade = m_Finalidade
End Property

Public Property Let Finalidade(ByVal Value As String)
    m_Finalidade = Value
End Property




Public Property Get Observacao() As String
    Observacao = m_Observacao
End Property

Public Property Let Observacao(ByVal Value As String)
    m_Observacao = Value
End Property



Public Property Get Ic() As String
    Ic = m_Ic
End Property

Public Property Let Ic(ByVal Value As String)
    m_Ic = Value
End Property

Public Property Get DataNegativa() As String
    DataNegativa = m_DataNegativa
End Property

Public Property Let DataNegativa(ByVal Value As String)
    m_DataNegativa = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Property Get CodNegativa() As Double
    CodNegativa = m_CodNegativa
End Property

Public Property Let CodNegativa(ByVal Value As Double)
    m_CodNegativa = Value
End Property

Public Function Buscar(CodNegativa As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "select * from tab_certidao_negativa where TCN_COD_NEGATIVA = '" & CodNegativa & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        Buscar = True
        m_CodNegativa = "" & rs!TCN_COD_NEGATIVA
        m_Im = "" & rs!TCN_TCI_IM
        m_DataNegativa = "" & rs!TCN_DATA_NEGATIVA
        m_Ic = "" & rs!TCN_TIM_IC
        m_Finalidade = "" & rs!TCN_FINALIDADE
        m_Validade = "" & rs!TCN_VALIDADE
        m_PeriodoInicial = "" & rs!TCN_PERIODO_INICIAL
        m_PeriodoFinal = "" & rs!TCN_PERIODO_FINAL
        m_CodUsuario = "" & rs!TCN_TUS_COD_USUARIO
        m_Tipo = "" & rs!TCN_TIPO
        m_Imposto = "" & rs!TCN_TIP_COD_IMPOSTO
    End If
End Function

Public Function Gravar() As Boolean
    Dim Valores As String
    Dim campos As String
    Dim condicao As String
    campos = "TCN_COD_NEGATIVA, TCN_TCI_IM, TCN_DATA_NEGATIVA, TCN_FINALIDADE, TCN_VALIDADE, TCN_TUS_COD_USUARIO, TCN_TIPO, TCN_TIM_IC,TCN_TIP_COD_IMPOSTO,TCN_OBSERVACAO"
    Valores = Bdados.PreparaValor(m_CodNegativa, Bdados.Converte(m_Im, tctexto), m_DataNegativa, m_Finalidade, m_Validade, m_CodUsuario, m_Tipo, m_Ic, m_Imposto, m_Observacao)
    
    If Trim(m_PeriodoInicial) <> "" And Trim(m_PeriodoFinal) <> "" Then
        campos = campos & "TCN_PERIODO_INICIAL, TCN_PERIODO_FINAL"
        Valores = Valores & Bdados.PreparaValor(m_PeriodoInicial, m_PeriodoFinal)
    End If
    condicao = " TCN_COD_NEGATIVA = " & m_CodNegativa
    If Bdados.GravaDados("tab_certidao_negativa", Valores, campos, condicao) Then Gravar = True
End Function

Public Function Excluir(CodNegativa) As Boolean
    Dim condicao As String
    condicao = "TCN_COD_NEGATIVA = '" & m_CodNegativa & "'"
    If Bdados.DeletaDados("tab_certidao_negativa", condicao) Then Excluir = True
End Function

Public Function BuscarContribuinte(Im As String, Optional RetNome As String, Optional RetCgc As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    
    Sql = "Select tci_nome,tci_logradouro,tci_nome_logradouro," & _
    " tci_numero, tci_complemento, tci_bairro, tci_cidade, tci_uf,tci_cgc_cpf " & _
    " FROM tab_Contribuinte where tci_im='" & Im & "' and tci_tsc_cod_sit_cad = 1"
    If Bdados.AbreTabela(Sql, rs) Then
        RetNome = "" & rs!tci_nome
        RetCgc = "" & rs!TCI_CGC_CPF
        BuscarContribuinte = True
    End If
    Bdados.FechaTabela rs
End Function

Public Sub BuscarCertidoes(Grid As Object, Tipo As TipoCertidao, Optional Im As String, Optional Ic As String, Optional CodCertidao As String)
    Dim Sql As String
    Sql = "SELECT TCN_COD_NEGATIVA AS Numero, " & _
            " TCN_FINALIDADE as Finalidade, " & _
            " TCN_PERIODO_INICIAL as Per_Inicial, " & _
            " TCN_PERIODO_FINAL as Per_Final, " & _
            " TCN_VALIDADE as Validade, " & _
            " TCN_DATA_NEGATIVA as Emissao,TCN_OBSERVACAO AS Observação " & _
        " FROM TAB_CERTIDAO_NEGATIVA" & _
        " where TCN_TIPO = '" & Tipo & "'"
    If Trim$(Im) <> "" Then
        Sql = Sql & " and TCN_TCI_IM ='" & Im & "'"
    End If
    If Trim$(Ic) <> "" Then
        Sql = Sql & " and TCN_TIM_IC ='" & Ic & "'"
    End If
    If Trim(CodCertidao) <> "" Then
        Sql = Sql & " and TCN_COD_NEGATIVA =" & CodCertidao
    End If
    Sql = Sql & " ORDER BY TCN_VALIDADE DESC"
    Grid.Preencher Bdados, Sql
End Sub

Public Function BuscaTexto(NomeTexto As String) As String
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = '" & NomeTexto & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        BuscaTexto = "" & rs!TPT_TEXTO
    End If
End Function

Public Function GravarTexto(NomeTexto As String, Texto As String) As Boolean
    Dim Valores As String
    Dim campos As String
    Valores = Bdados.PreparaValor(NomeTexto, Texto)
    campos = "tpt_parametro,TPT_TEXTO"
    Bdados.GravaDados "TAB_PARAMETRO_TEXTO", Valores, campos, "TPT_PARAMETRO = '" & NomeTexto & "'"
End Function

Public Function DebitosAbertos(Grid As Object, Optional Im As String, Optional Ic As String, Optional CodTributo As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String) As Boolean
    Dim Sql As String
    Dim CodArr As String
    Dim CodLanc As String
    Dim PeriodoData As String
    Dim Obrig As New Obrigacao
    
    If Trim(Im) <> "" Then
        CodArr = " tdr_im ='" & Im & "'"
        CodLanc = " tgt_im ='" & Im & "'"
    End If
    If Trim(Ic) <> "" Then
        CodArr = " tdr_tim_ic ='" & Ic & "'"
        CodLanc = " tgt_tim_ic ='" & Ic & "'"
    End If
    Sql = "SELECT tgt_cod_pagamento AS Documento," & _
          " tip_sigla_imposto AS Imposto ," & _
          " tgt_periodo as Periodo," & _
          " tgt_Valor_tributo as Valor," & _
          " TGT_DATA_VENCIMENTO AS Vencimento" & _
          " FROM TAB_GERACAO_TRIBUTO, TAB_IMPOSTO " & _
          " where tgt_tip_cod_imposto = tip_cod_imposto" & _
          " and tgt_COD_PAGAMENTO NOT IN " & _
                        " (SELECT TDR_tgt_cod_pagamento FROM TAB_DARM_RECEBIDO WHERE " & CodArr & " )" & _
                        " AND " & CodLanc & " AND TGT_TIP_COD_IMPOSTO <> '" & Const_Extrato & "'" & _
                        " AND tgt_tip_cod_imposto <>'" & Const_Notificacao & "'"
    If Trim(CodTributo) <> "" Then
        Sql = Sql & " and tgt_tip_cod_imposto = '" & CodTributo & "'"
    End If
        
    PeriodoInicial = Edita.TiraPic(PeriodoInicial, "/")
    PeriodoFinal = IIf(PeriodoFinal <> "", PeriodoFinal, Format(Date, "mm/yyyy"))
    PeriodoFinal = Edita.TiraPic(PeriodoFinal, "/")
    
    If Len(PeriodoInicial) > 0 Then
        PeriodoData = "01/" & IIf(Len(PeriodoInicial) = 6, Left(PeriodoInicial, 2) & "/", "01/")
        PeriodoData = PeriodoData & Right(PeriodoInicial, 4)
        Sql = Sql & " and tgt_data_vencimento >= " & Bdados.Converte(PeriodoData, TCDataHora)
    End If

    If Len(PeriodoFinal) > 0 Then
        PeriodoData = "01/" & IIf(Len(PeriodoFinal) = 6, Left(PeriodoFinal, 2) & "/", "12/") & Right(PeriodoFinal, 4)
        PeriodoData = Util.UltimoDiaDoMes(CDate(PeriodoData))
        Sql = Sql & " and tgt_data_vencimento >= " & Bdados.Converte(PeriodoData, TCDataHora)
    End If
    
    Sql = Sql & " ORDER BY tip_sigla_imposto, tgt_periodo"
    If Grid.Preencher(Bdados, Sql) Then DebitosAbertos = True
    
End Function

Public Function DebitosAbertosVencidos(Grid As Object, Optional Im As String, Optional Ic As String, Optional CodTributo As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String) As Boolean
    Dim Sql As String
    Dim CodArr As String
    Dim CodLanc As String
    Dim PeriodoData As String
    
    If Trim(Im) <> "" Then
        CodArr = " tdr_im ='" & Im & "'"
        CodLanc = " tgt_im ='" & Im & "'"
    End If
    If Trim(Ic) <> "" Then
        CodArr = " tdr_tim_ic ='" & Ic & "'"
        CodLanc = " tgt_tim_ic ='" & Ic & "'"
    End If
    Sql = "SELECT tgt_cod_pagamento AS Documento," & _
          " tip_sigla_imposto AS Imposto ," & _
          " tgt_periodo as Periodo," & _
          " tgt_Valor_tributo as Valor," & _
          " TGT_DATA_VENCIMENTO AS Vencimento" & _
          " FROM TAB_GERACAO_TRIBUTO, TAB_IMPOSTO " & _
          " where tgt_tip_cod_imposto = tip_cod_imposto" & _
          " and tgt_COD_PAGAMENTO NOT IN" & _
                        " (SELECT TDR_tgt_cod_pagamento FROM TAB_DARM_RECEBIDO WHERE " & CodArr & " )" & _
                        " AND " & CodLanc & " AND TGT_TIP_COD_IMPOSTO <> '" & Const_Extrato & "'" & _
                        " AND tgt_tip_cod_imposto <>'" & Const_Notificacao & "'"
                        
    If Trim(CodTributo) <> "" Then
        Sql = Sql & " and tgt_tip_cod_imposto = '" & CodTributo & "'"
    End If
    
    PeriodoInicial = Edita.TiraPic(PeriodoInicial, "/")
    PeriodoFinal = IIf(PeriodoFinal <> "", PeriodoFinal, Format(Date, "mm/yyyy"))
    PeriodoFinal = Edita.TiraPic(PeriodoFinal, "/")

    If Len(PeriodoInicial) > 0 Then
        PeriodoData = "01/" & IIf(Len(PeriodoInicial) = 6, Left(PeriodoInicial, 2) & "/", "01/")
        PeriodoData = PeriodoData & Right(PeriodoInicial, 4)
        Sql = Sql & " and tgt_data_vencimento >= " & Bdados.Converte(PeriodoData, TCDataHora)
    End If

    If Len(PeriodoFinal) > 0 Then
        PeriodoData = "01/" & IIf(Len(PeriodoFinal) = 6, Left(PeriodoFinal, 2) & "/", "12/") & Right(PeriodoFinal, 4)
        PeriodoData = Util.UltimoDiaDoMes(CDate(PeriodoData))
        Sql = Sql & " and tgt_data_vencimento <= " & Bdados.Converte(PeriodoData, TCDataHora)
    End If
   
    Sql = Sql & " ORDER BY tip_sigla_imposto, tgt_periodo"
    If Grid.Preencher(Bdados, Sql) Then DebitosAbertosVencidos = True
End Function

Public Sub PreencherCboImposto(Combo As Object)
    Dim Sql As String
    
    Sql = "Select  tip_cod_imposto , tip_sigla_imposto " & Bdados.Concatena & " ' - ' " & Bdados.Concatena & " tip_nome_imposto  From TAB_IMPOSTO "
    
    Sql = Sql & " order by tip_sigla_imposto asc"
    Combo.Preencher Bdados, Sql, 1
End Sub
Public Sub PreencherGridImposto(Grid As Object)
    Dim Sql As String
    
    Sql = "Select  tip_sigla_imposto " & Bdados.Concatena & " ' - ' " & Bdados.Concatena & " tip_nome_imposto,tip_cod_imposto  From TAB_IMPOSTO "
    
    Sql = Sql & " order by tip_sigla_imposto asc"
    
    Grid.Preencher Bdados, Sql
End Sub

