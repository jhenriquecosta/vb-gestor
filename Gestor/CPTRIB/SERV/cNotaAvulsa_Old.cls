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
Private m_CodPagamento As String
Private m_CodImposto As String
Private m_DataRecepcao As String
Private m_Material As Double
Private m_IRRF As Double
Private m_INSS_Indice As Double
Private m_INSS_Valor As Double

Public Property Get IRRF() As Double
    IRRF = m_IRRF
End Property

Public Property Let IRRF(ByVal Value As Double)
    m_IRRF = Value
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

Public Property Get StatusNota() As Integer
    StatusNota = m_StatusNota
End Property

Public Property Let StatusNota(ByVal Value As Integer)
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
        & " TNA_IRRF"
    valores = Bdados.PreparaValor(m_NumNota, m_StatusNota, _
        m_DataEmissao, m_IdentidadeRemetente, _
        m_IdentidadeDestinatario, m_ValorNota, _
        m_ValorImposto, m_CodUsuario, _
        m_Aliquota, m_Periodo, _
        m_CodPagamento, m_CodImposto, _
        m_DataRecepcao, m_Material, _
        m_IRRF)
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
    Sql = "Select  tna_numero_nota as Num_Nota,tna_tca_identidade_remetente AS IM_CNPJ_CPF,tna_cod_pagamento AS Cod_Pagamento,tna_periodo as Per�odo,tna_data_emissao as Dt_Emiss�o, tna_valor_nota as Vl_Nota," & _
        " TNA_TCA_IDENTIDADE_DEST as Destinat�rio from tab_nota_avulsa "
    If Trim(NumNota) <> "" Then
        sqlaux = sqlaux & " and tna_numero_nota = '" & NumNota & "'"
    End If
    If Trim(NomeContribuinte) <> "" Then
        sqlaux = sqlaux & " and tna_tca_identidade_remetente in (select tci_im from tab_contribuinte " & _
                    " where tci_nome like '" & NomeContribuinte & "%' or tci_nome like '%" & NomeContribuinte & "%')" & _
                    " or tna_tca_identidade_remetente in (select tca_identidade from tab_contribuinte_avulso where " & _
                    " tca_nome like '" & NomeContribuinte & "%' or tca_nome like '% " & NomeContribuinte & "%') AND tna_status_nota = 0"
    End If
    If sqlaux <> "" Then
        Sql = Sql & " where " & Right(sqlaux, Len(sqlaux) - 4)
    End If
    Sql = Sql & " order by tna_numero_nota"
    If Grid.Preencher(Bdados, Sql) Then PreencherGrid = True
End Function

Public Function Excluir(NumNota As String) As Boolean
    If Bdados.DeletaDados("TAB_NOTA_AVULSA", "tna_numero_nota =" & NumNota) = True Then
        Bdados.DeletaDados "TAB_OBRIGACAO_CONTRIBUINTE", "TOC_NUM_DOC_ORIGEM = " & Bdados.Converte(NumNota, tctexto)
    End If
End Function


