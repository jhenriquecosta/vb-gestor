VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNotaAvulsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_NumNota As Double
Private m_StatusNota As Integer
Private m_DataEmissao As String
Private m_IdentidadeRemetente As String
Private m_IdentidadeDestinatario As String
Private m_ValorNota As Double
Private m_ValorImposto As Double
Private m_CodUsuario As String
Private m_Aliquota As Double
Private m_Periodo As String
Private m_ObsCancelamento As String
Private m_CodPagamento As String
Private m_CodImposto As String
Private m_DataRecepcao As String
Private m_Material As Double
Private m_IRRF As Double
Private m_INSS_Valor As Double
Private m_INSS_Indice As Double
Private m_IRRF_INDICE As Double
Public Property Let ObsCancelamento(ByVal Value As String)
    m_ObsCancelamento = Value
End Property

Public Property Get ObsCancelamento() As String
    ObsCancelamento = m_ObsCancelamento
End Property

Public Property Get INSS_Valor() As Double
    INSS_Valor = m_INSS_Valor
End Property

Public Property Let INSS_Valor(ByVal Value As Double)
    m_INSS_Valor = Value
End Property

Public Property Get INSS_Indice() As Double

    INSS_Indice = m_INSS_Indice
End Property

Public Property Let INSS_Indice(ByVal Value As Double)

    m_INSS_Indice = Value

End Property

Public Property Get IRRF() As Double
    IRRF = m_IRRF
End Property

Public Property Let IRRF(ByVal Value As Double)
    m_IRRF = Value
End Property

Public Property Get IRRF_INDICE() As Double
    IRRF_INDICE = m_IRRF_INDICE
End Property

Public Property Let IRRF_INDICE(ByVal Value As Double)
    m_IRRF_INDICE = Value
End Property

Public Property Get Material() As Double
    Material = m_Material
End Property

Public Property Let Material(ByVal Value As Double)
    m_Material = Value
End Property

Public Property Get DataRecepcao() As String
    DataRecepcao = m_DataRecepcao
End Property

Public Property Let DataRecepcao(ByVal Value As String)
    m_DataRecepcao = Value
End Property

Public Property Get CodImposto() As String
    CodImposto = m_CodImposto
End Property

Public Property Let CodImposto(ByVal Value As String)
    m_CodImposto = Value
End Property

Public Property Get CodPagamento() As String
    CodPagamento = m_CodPagamento
End Property

Public Property Let CodPagamento(ByVal Value As String)
    m_CodPagamento = Value
End Property

Public Property Get Periodo() As String
    Periodo = m_Periodo
End Property

Public Property Let Periodo(ByVal Value As String)
    m_Periodo = Value
End Property

Public Property Get Aliquota() As Double
    Aliquota = m_Aliquota
End Property

Public Property Let Aliquota(ByVal Value As Double)
    m_Aliquota = Value
End Property

Public Property Get CodUsuario() As String
    CodUsuario = m_CodUsuario
End Property

Public Property Let CodUsuario(ByVal Value As String)
    m_CodUsuario = Value
End Property

Public Property Get ValorImposto() As Double
    ValorImposto = m_ValorImposto
End Property

Public Property Let ValorImposto(ByVal Value As Double)
    m_ValorImposto = Value
End Property

Public Property Get ValorNota() As Double
    ValorNota = m_ValorNota
End Property

Public Property Let ValorNota(ByVal Value As Double)
    m_ValorNota = Value
End Property

Public Property Get IdentidadeDestinatario() As String
    IdentidadeDestinatario = m_IdentidadeDestinatario
End Property

Public Property Let IdentidadeDestinatario(ByVal Value As String)
    m_IdentidadeDestinatario = Value
End Property

Public Property Get IdentidadeRemetente() As String
    IdentidadeRemetente = m_IdentidadeRemetente
End Property

Public Property Let IdentidadeRemetente(ByVal Value As String)
    m_IdentidadeRemetente = Value
End Property

Public Property Get DataEmissao() As String
    DataEmissao = m_DataEmissao
End Property

Public Property Let DataEmissao(ByVal Value As String)
    m_DataEmissao = Value
End Property

Public Property Get statusNota() As Integer
    statusNota = m_StatusNota
End Property

Public Property Let statusNota(ByVal Value As Integer)
    m_StatusNota = Value
End Property

Public Property Get NumNota() As Double
    NumNota = m_NumNota
End Property

Public Property Let NumNota(ByVal Value As Double)
    m_NumNota = Value
End Property

Public Function Buscar(NumNota As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "select * from tab_nota_avulsa where tna_numero_nota = '" & NumNota & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        m_NumNota = "" & rs!tna_numero_nota
        m_StatusNota = "" & rs!tna_status_nota
        m_DataEmissao = "" & rs!tna_data_emissao
        m_IdentidadeRemetente = "" & rs!tna_tca_identidade_remetente
        m_IdentidadeDestinatario = "" & rs!TNA_TCA_IDENTIDADE_DEST
        m_ValorNota = "" & rs!tna_valor_nota
        m_ValorImposto = "" & rs!tna_valor_imposto
        m_CodUsuario = "" & rs!tna_tus_cod_usuario
        m_Aliquota = "" & rs!tna_aliquota
        m_Periodo = "" & rs!tna_periodo
        m_CodPagamento = "" & rs!tna_cod_pagamento
        m_CodImposto = "" & rs!tna_TIP_COD_IMPOSTO
        m_DataRecepcao = "" & rs!tna_DATA_RECEPCAO
        m_Material = "" & Nvl("" & rs!tna_MATERIAL, 0)
        m_IRRF = "" & Nvl("" & rs!tna_IRRF, 0)
        m_INSS_Indice = "" & Nvl("" & rs!tna_INSS_Indice, 0)
        m_INSS_Valor = "" & Nvl("" & rs!tna_INSS_Valor, 0)
        m_IRRF_INDICE = "" & Nvl("" & rs!tna_IRRF_INDICE, 0)
        
        Buscar = True
    End If
End Function

Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    campos = " tna_numero_nota, tna_status_nota," _
        & " tna_data_emissao, tna_tca_identidade_remetente," _
        & " TNA_TCA_IDENTIDADE_DEST, tna_valor_nota," _
        & " tna_valor_imposto, tna_tus_cod_usuario," _
        & " tna_aliquota, tna_periodo," _
        & " tna_cod_pagamento, TNA_TIP_COD_IMPOSTO," _
        & " TNA_DATA_RECEPCAO, TNA_MATERIAL," _
        & " TNA_IRRF,TNA_INSS_INDICE,TNA_INSS_VALOR,TNA_IRRF_INDICE"
    valores = Bdados.PreparaValor(m_NumNota, m_StatusNota, _
        m_DataEmissao, m_IdentidadeRemetente, _
        m_IdentidadeDestinatario, m_ValorNota, _
        m_ValorImposto, m_CodUsuario, _
        m_Aliquota, m_Periodo, _
        m_CodPagamento, m_CodImposto, _
        m_DataRecepcao, m_Material, _
        m_IRRF, m_INSS_Indice, m_INSS_Valor, m_IRRF_INDICE)
    If Bdados.InsereDados("Tab_nota_avulsa", valores, campos) Then
        Salvar = True
    End If
End Function

Public Function PreencheCboAtividade(Combo As Object)
    Dim Sql As String
    Sql = " SELECT Tae_nome, tae_cae   from tab_atividade_economica"
    Combo.Preencher Bdados, Sql
End Function


Public Function BuscaAliquota() As String
    BuscaAliquota = CDbl(Imposto.BuscaAliquota(BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto ='" & Imposto.NomeTributo(ttr_ISSQN) & "'"), Right(Year(Date), 4))) * 100
End Function

Public Function PreencherGrid(Grid As Object, Optional NumNota As String, Optional NomeContribuinte As String) As Boolean
    Dim Sql As String, sqlaux As String
    Dim rs As VSRecordset
    Sql = "Select  tna_numero_nota as Num_Nota,tna_tca_identidade_remetente AS IM_CNPJ_CPF,tna_cod_pagamento AS Cod_Pagamento,tna_periodo as Período,tna_data_emissao as Dt_Emissão, tna_valor_nota as Vl_Nota," & _
        " TNA_TCA_IDENTIDADE_DEST as Destinatário from tab_nota_avulsa "
    If Trim(NumNota) <> "" Then
        sqlaux = sqlaux & " and tna_numero_nota = '" & NumNota & "'"
    End If
    If Trim(NomeContribuinte) <> "" Then
        sqlaux = sqlaux & " and tna_tca_identidade_remetente in (select tci_im from tab_contribuinte " & _
                    " where tci_nome like '" & NomeContribuinte & "%' or tci_nome like '%" & NomeContribuinte & "%')" & _
                    " or tna_tca_identidade_remetente in (select tca_identidade from tab_contribuinte_avulso where " & _
                    " tca_nome like '" & NomeContribuinte & "%' or tca_nome like '% " & NomeContribuinte & "%') AND tna_status_nota in (0,3)"
    End If
    If sqlaux <> "" Then
        Sql = Sql & " where " & Right(sqlaux, Len(sqlaux) - 4)
    End If
    Sql = Sql & " order by tna_numero_nota"
    If Grid.Preencher(Bdados, Sql) Then PreencherGrid = True
End Function
Public Function RetornaSQL(Optional DataInicial As String, Optional DataFinal As String, Optional ordem As String) As String
    Dim Sql As String, sqlaux As String
    Dim rs As VSRecordset
    Sql = "SELECT   Tab_Nota_Avulsa.tna_numero_nota, Tab_Nota_Avulsa.tna_status_nota, Tab_Nota_Avulsa.tna_data_emissao, " _
         & " Tab_Nota_Avulsa.tna_tca_identidade_remetente, Tab_Nota_Avulsa.tna_tca_identidade_destinatario, Tab_Nota_Avulsa.tna_valor_nota, " _
         & " Tab_Nota_Avulsa.tna_valor_imposto, Tab_Nota_Avulsa.tna_periodo,Tab_Contribuinte.tci_nome , Tab_Contribuinte.tci_fantasia" _
         & " FROM Tab_Nota_Avulsa INNER JOIN Tab_Contribuinte ON Tab_Nota_Avulsa.tna_tca_identidade_remetente = Tab_Contribuinte.tci_im" _
         & " WHERE Tab_Nota_Avulsa.tna_data_emissao >=" & Bdados.Converte(DataInicial, TCDataHora) & " and Tab_Nota_Avulsa.tna_data_emissao <=" & Bdados.Converte(DataFinal, TCDataHora) _
         & " ORDER BY " & ordem
    RetornaSQL = Sql
End Function
Public Function PreencherGridComServico(Grid As Object, Optional NumNota As String, Optional NomeContribuinte As String) As Boolean
    Dim Sql As String, sqlaux As String
    Dim rs As VSRecordset
    Sql = "SELECT Tab_Nota_Avulsa.tna_numero_nota as Numero, Tab_Nota_Avulsa.tna_data_emissao as Emissao,  Tab_Nota_Avulsa.tna_periodo as Periodo, " & _
    "Tab_Nota_Avulsa.tna_valor_nota as Valor , Tab_Item_Nota_Avulsa.tin_descricao_servico as Servico, Tab_Item_Nota_Avulsa.tin_codigo as Item " & _
    "FROM Tab_Item_Nota_Avulsa INNER JOIN Tab_Nota_Avulsa ON Tab_Item_Nota_Avulsa.tin_tna_numero_nota = Tab_Nota_Avulsa.tna_numero_nota"
    If Trim(NumNota) <> "" Then
        sqlaux = sqlaux & " and Tab_Nota_Avulsa.tna_numero_nota = '" & NumNota & "'"
    End If
    If Trim(NomeContribuinte) <> "" Then
        sqlaux = sqlaux & " and Tab_Nota_Avulsa.tna_tca_identidade_remetente in (select tci_im from tab_contribuinte " & _
                    " where tci_nome like '" & NomeContribuinte & "%' or tci_nome like '%" & NomeContribuinte & "%')" & _
                    " or tna_tca_identidade_remetente in (select tca_identidade from tab_contribuinte_avulso where " & _
                    " tca_nome like '" & NomeContribuinte & "%' or tca_nome like '% " & NomeContribuinte & "%') AND tna_status_nota in (0,3)"
    End If
    If sqlaux <> "" Then
        Sql = Sql & " where " & Right(sqlaux, Len(sqlaux) - 4)
    End If
    Sql = Sql & " order by Tab_Nota_Avulsa.tna_numero_nota"
    If Grid.Preencher(Bdados, Sql) Then PreencherGridComServico = True
End Function


Public Function Excluir(NumNota As String, Motivo As String) As Boolean
    'BCP - GLEYSON
    If Bdados.Executa("UPDATE TAB_NOTA_AVULSA SET TNA_STATUS_NOTA=3, TNA_OBS_CANCELAMENTO='" & Motivo & "' WHERE TNA_NUMERO_NOTA=" & NumNota) = True Then
        'Bdados.DeletaDados "TAB_OBRIGACAO_CONTRIBUINTE", "TOC_NUM_DOC_ORIGEM = " & Bdados.Converte(NumNota, tctexto)
        'Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "TCC_CODIGO_CONTA = " & Bdados.Converte(NumNota, tctexto)
    End If
    
    'If Bdados.DeletaDados("TAB_NOTA_AVULSA", "tna_numero_nota =" & NumNota) = True Then
     '   Bdados.DeletaDados "TAB_OBRIGACAO_CONTRIBUINTE", "TOC_NUM_DOC_ORIGEM = " & Bdados.Converte(NumNota, tctexto)
    'End If
End Function




