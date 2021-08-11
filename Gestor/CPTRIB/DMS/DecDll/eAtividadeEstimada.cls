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

Public Function Buscar(CodAtividade As String, LimiteInferior As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "select * from tab_atividade_estimada where TAT_TAE_CAE = '" & CodAtividade & "' and TAT_LIMITE_INFERIOR = '" & LimiteInferior & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        m_CodAtividade = "" & Rs!TAT_TAE_CAE
        m_LimiteInferior = "" & Rs!TAT_LIMITE_INFERIOR
        m_LimiteSuperior = "" & Rs!TAT_LIMITE_SUPERIOR
        m_LimiteValor = "" & Rs!TAT_LIMITE_VALOR
        m_LimiteFator = "" & Rs!TAT_LIMITE_FATOR
        Buscar = True
    End If
End Function

Public Function Gravar() As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Campos = "TAT_TAE_CAE, TAT_LIMITE_INFERIOR, TAT_LIMITE_SUPERIOR, TAT_LIMITE_VALOR, TAT_LIMITE_FATOR"
    Valores = Bdados.PreparaValor(m_CodAtividade, m_LimiteInferior, m_LimiteSuperior, m_LimiteValor, m_LimiteFator)
    Condicao = "TAT_TAE_CAE=" & m_CodAtividade & " and TAT_LIMITE_INFERIOR = " & m_LimiteInferior
    If Bdados.GravaDados("Tab_Atividade_estimada", Valores, Campos, Condicao) Then Gravar = True
End Function

Public Function Excluir(CodAtividade As String, LimiteInferior As String) As Boolean
    Dim Condicao As String
    Condicao = "TAT_TAE_CAE=" & CodAtividade & " and TAT_LIMITE_INFERIOR = " & LimiteInferior
    If Bdados.DeletaDados("Tab_Atividade_estimada", Condicao) Then
        Excluir = True
    End If
End Function

Public Function PreencherGrd(Grid As Object, CodAtividade As String) As Boolean
    Dim Sql As String
    Sql = "SELECT TAT_TAE_CAE as CAE, TAT_LIMITE_INFERIOR as Inferior, TAT_LIMITE_SUPERIOR as Superior, " & _
            " TAT_LIMITE_VALOR as Valor, TGE_NOME AS  [Mult. pela Base] " & _
            " FROM Tab_Atividade_estimada, VIS_SIM_NAO " & _
            " WHERE TAT_TAE_CAE = " & CodAtividade & " AND " & _
            " TGE_CODIGO = (TAT_LIMITE_FATOR + 1)"
    If Grid.Preencher(Bdados, Sql, 1400) Then PreencherGrd = True
End Function

Public Function PreencherAtividadesEstimativas(Grid As Object) As Boolean
    Dim Sql As String
    Sql = "Select tae_cae as Código, tae_nome as Atividade, " & _
            " tga_nome as Grupo," & Bdados.Converte("tae_valor", TCDuplo) & " as [Valor(R$)], " & _
            " tae_desc_fator as Fator " & _
            " from Tab_Atividade_Economica, Tab_Grupo_Atividade " & _
            " where tae_tga_cod_grupo = tga_cod_grupo and TAE_ESTIMATIVO = 2"
    If Grid.Preencher(Bdados, Sql) Then PreencherAtividadesEstimativas = True
End Function
