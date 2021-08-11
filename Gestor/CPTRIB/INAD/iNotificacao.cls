VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iNotificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_CodNotificacao As Double
Private m_Im As String
Private m_DataEmissao As String
Private m_Vencimento As String
Private m_ValorNotificacao As Double
Private m_CodUsuario As String
Private m_TipoNotificacao As Integer

Public Property Get TipoNotificacao() As Integer
    TipoNotificacao = m_TipoNotificacao
End Property

Public Property Let TipoNotificacao(ByVal Value As Integer)
    m_TipoNotificacao = Value
End Property

Public Property Get CodUsuario() As String
    CodUsuario = m_CodUsuario
End Property

Public Property Let CodUsuario(ByVal Value As String)
    m_CodUsuario = Value
End Property

Public Property Get ValorNotificacao() As Double
    ValorNotificacao = m_ValorNotificacao
End Property

Public Property Let ValorNotificacao(ByVal Value As Double)
    m_ValorNotificacao = Value
End Property

Public Property Get Vencimento() As String
    Vencimento = m_Vencimento
End Property

Public Property Let Vencimento(ByVal Value As String)
    m_Vencimento = Value
End Property

Public Property Get DataEmissao() As String
    DataEmissao = m_DataEmissao
End Property

Public Property Let DataEmissao(ByVal Value As String)
    m_DataEmissao = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Property Get CodNotificacao() As Double
    CodNotificacao = m_CodNotificacao
End Property

Public Property Let CodNotificacao(ByVal Value As Double)
    m_CodNotificacao = Value
End Property

Public Function Buscar(CodNotificacao As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "select * from tab_notificacao where tnt_cod_notificacao = '" & CodNotificacao & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        Buscar = True
        m_CodNotificacao = "" & Rs!TNT_COD_NOTIFICACAO
        m_Im = "" & Rs!TNT_TCI_IM
        m_DataEmissao = "" & Rs!TNT_DT_EMISSAO
        m_Vencimento = "" & Rs!TNT_VENCIMENTO
        m_ValorNotificacao = "" & Rs!TNT_VALOR_NOTIFICACAO
        m_CodUsuario = "" & Rs!TNT_TUS_COD_USUARIO
        m_TipoNotificacao = "" & Rs!TNT_TIPO_NOTIFICACAO
    End If
End Function

Public Function Gravar() As Boolean
    Dim campos As String, valores As String, condicao As String
    campos = "TNT_COD_NOTIFICACAO, TNT_TCI_IM, TNT_DT_EMISSAO, TNT_VENCIMENTO, TNT_VALOR_NOTIFICACAO, TNT_TUS_COD_USUARIO, TNT_TIPO_NOTIFICACAO"
    valores = Bdados.PreparaValor(m_CodNotificacao, m_Im, m_DataEmissao, m_Vencimento, m_ValorNotificacao, m_CodUsuario, m_TipoNotificacao)
    condicao = "TNT_COD_NOTIFICACAO =  '" & m_CodNotificacao & "'"
    If Bdados.GravaDados("tab_notificacao", valores, campos, condicao) Then Gravar = True
End Function

Public Function ExibirNotificacoes(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    Dim condicao As String
    
    condicao = ""
    If Trim$(Im) <> "" Then
        condicao = " and TNT_TCI_IM='" & Im & "'"
    End If
    condicao = condicao & " and TNT_TIPO_NOTIFICACAO = 1"
    Sql = "SELECT TNT_COD_NOTIFICACAO AS Número, " & _
             " TNT_DT_EMISSAO as Emissão, " & _
             " TNT_VENCIMENTO as Vencimento, " & _
             FuncaoReal("TNT_VALOR_NOTIFICACAO") & " as Valor" & _
             " FROM TAB_NOTIFICACAO "
    If condicao <> "" Then
        Sql = Sql & " WHERE " & Right(condicao, Len(condicao) - 4)
    End If
    Sql = Sql & " ORDER BY TNT_VENCIMENTO"
    If Grid.Preencher(Bdados, Sql) Then
        ExibirNotificacoes = True
        Grid.Mensagem = "Total : " & Format(Grid.Colunas(4).Soma, "currency")
    Else
        Grid.Mensagem = "Registros não encontrados"
    End If
End Function
