VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eAtividadeEstimada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_CodAtividade As Double
Private m_LimiteInferior As Double
Private m_LimiteSuperior As Double
Private m_LimiteValor As Double
Private m_LimiteFator As Double
Private m_ValorExcedente As Double
Private m_Valor_UFM       As Double
Private m_Ano As Integer

Public Property Get Ano() As Integer
    Ano = m_Ano
End Property

Public Property Let Ano(ByVal Value As Integer)
    m_Ano = Value
End Property


Public Property Get ValorUFM() As Double
    ValorUFM = m_Valor_UFM
End Property

Public Property Let ValorUFM(ByVal Value As Double)
    m_Valor_UFM = Value
End Property


Public Property Get ValorExcedente() As Double
    ValorExcedente = m_ValorExcedente
End Property

Public Property Let ValorExcedente(ByVal Value As Double)
    m_ValorExcedente = Value
End Property

Public Property Get LimiteFator() As Double
    LimiteFator = m_LimiteFator
End Property

Public Property Let LimiteFator(ByVal Value As Double)
    m_LimiteFator = Value
End Property

Public Property Get LimiteValor() As Double
    LimiteValor = m_LimiteValor
End Property

Public Property Let LimiteValor(ByVal Value As Double)
    m_LimiteValor = Value
End Property

Public Property Get LimiteSuperior() As Double
    LimiteSuperior = m_LimiteSuperior
End Property

Public Property Let LimiteSuperior(ByVal Value As Double)
    m_LimiteSuperior = Value
End Property

Public Property Get LimiteInferior() As Double
    LimiteInferior = m_LimiteInferior
End Property

Public Property Let LimiteInferior(ByVal Value As Double)
    m_LimiteInferior = Value
End Property

Public Property Get CodAtividade() As Double
    CodAtividade = m_CodAtividade
End Property

Public Property Let CodAtividade(ByVal Value As Double)
    m_CodAtividade = Value
End Property

Public Function Buscar(CodAtividade As String, LimiteInferior As String, Ano As Integer) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "select * from tab_atividade_estimada where TAT_TAE_CAE = '" & CodAtividade & "' and TAT_LIMITE_INFERIOR = '" & LimiteInferior & "' and TAT_ANO =" & Ano
    If Bdados.AbreTabela(Sql, Rs) Then
        m_CodAtividade = "" & Rs!TAT_TAE_CAE
        m_LimiteInferior = "" & Rs!TAT_LIMITE_INFERIOR
        m_LimiteSuperior = "" & Rs!TAT_LIMITE_SUPERIOR
        m_LimiteValor = "" & Rs!TAT_LIMITE_VALOR
        m_LimiteFator = "" & Rs!TAT_LIMITE_FATOR
        m_ValorExcedente = "" & Rs!TAT_VALOR_EXCEDENTE
        m_Ano = Ano
        Buscar = True
    End If
End Function

Public Function Gravar() As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Campos = "TAT_TAE_CAE, TAT_LIMITE_INFERIOR, TAT_LIMITE_SUPERIOR, TAT_LIMITE_VALOR, TAT_LIMITE_FATOR, TAT_VALOR_EXCEDENTE,TAT_VALOR_UFM,TAT_ANO"
    Valores = Bdados.PreparaValor(m_CodAtividade, m_LimiteInferior, m_LimiteSuperior, m_LimiteValor, m_LimiteFator, m_ValorExcedente, m_Valor_UFM, m_Ano)
    Condicao = "TAT_TAE_CAE=" & m_CodAtividade & " and TAT_LIMITE_INFERIOR = " & m_LimiteInferior & " and TAT_ANO =" & m_Ano
    If Bdados.GravaDados("Tab_Atividade_estimada", Valores, Campos, Condicao) Then Gravar = True
End Function

Public Function Excluir(CodAtividade As String, LimiteInferior As String) As Boolean
    Dim Condicao As String
    Condicao = "TAT_TAE_CAE=" & CodAtividade & " and TAT_LIMITE_INFERIOR = " & LimiteInferior & " and TAT_ANO =" & m_Ano
    If Bdados.DeletaDados("Tab_Atividade_estimada", Condicao) Then
        Excluir = True
    End If
End Function

Public Function PreencherGrd(Grid As Object, CodAtividade As String, Optional Ano As Integer) As Boolean
    Dim Sql As String
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "SELECT TAT_TAE_CAE as CAE, TAT_LIMITE_INFERIOR as Inferior, TAT_LIMITE_SUPERIOR as Superior,TAT_VALOR_UFM as UFM, " & _
                " TAT_LIMITE_VALOR as Valor, TGE_NOME AS  [Mult. pela Base],TAT_VALOR_EXCEDENTE as [Valor Exedente] " & _
                " FROM Tab_Atividade_estimada, VIS_SIM_NAO " & _
                " WHERE TAT_TAE_CAE = " & CodAtividade & " AND " & _
                " TGE_CODIGO = (TAT_LIMITE_FATOR + 1) "
        If Ano > 0 Then
            Sql = Sql & " AND TAT_ANO =" & Ano
        Else
            Sql = Sql & " AND TAT_ANO = (SELECT MAX(TAT_ANO) FROM TAB_ATIVIDADE_ESTIMADA WHERE TAT_TAE_CAE = " & CodAtividade & ")"
        End If
        Sql = Sql & " ORDER BY TAT_LIMITE_INFERIOR "
    Else
        Sql = "SELECT TAT_TAE_CAE as CAE, TAT_LIMITE_INFERIOR as Inferior, TAT_LIMITE_SUPERIOR as Superior,TAT_VALOR_UFM as UFM, " & _
                " TAT_LIMITE_VALOR as Valor, TGE_NOME AS  Mult_pela_Base,TAT_VALOR_EXCEDENTE as Valor_Exedente " & _
                " FROM Tab_Atividade_estimada, VIS_SIM_NAO " & _
                " WHERE TAT_TAE_CAE = " & CodAtividade
        If Ano > 0 Then Sql = Sql & " AND TAT_ANO =" & Ano
        Sql = Sql & " ORDER BY TAT_LIMITE_INFERIOR "
    End If
    If Grid.Preencher(Bdados, Sql, 1400) Then PreencherGrd = True
End Function

Public Function PreencherAtividadesEstimativas(Grid As Object) As Boolean
    Dim Sql As String
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "Select tae_cae as Código, tae_nome as Atividade, " & _
                " tga_nome as Grupo," & Bdados.Converte("tae_valor", TCDuplo) & " as [Valor(R$)], " & _
                " tae_desc_fator as Fator " & _
                " from Tab_Atividade_Economica, Tab_Grupo_Atividade " & _
                " where tae_tga_cod_grupo = tga_cod_grupo and TAE_ESTIMATIVO = 2"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "Select tae_cae as Código, tae_nome as Atividade, " & _
                " tga_nome as Grupo,TO_CHAR(tae_valor,'999999.99') as Valor, " & _
                " tae_desc_fator as Fator " & _
                " from Tab_Atividade_Economica, Tab_Grupo_Atividade " & _
                " where tae_tga_cod_grupo = tga_cod_grupo and TAE_ESTIMATIVO = 2"
    End If
    If Grid.Preencher(Bdados, Sql) Then PreencherAtividadesEstimativas = True
End Function
