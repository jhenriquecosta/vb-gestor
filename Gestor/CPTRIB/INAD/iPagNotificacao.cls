VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iPagNotificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_CodNotificacao As Double
Private m_CodPagamento As Double
Private m_Valor As Double
Private m_CodImposto As String

Public Property Get CodImposto() As String
    CodImposto = m_CodImposto
End Property

Public Property Let CodImposto(ByVal Value As String)
    m_CodImposto = Value
End Property

Public Property Get Valor() As Double
    Valor = m_Valor
End Property

Public Property Let Valor(ByVal Value As Double)
    m_Valor = Value
End Property

Public Property Get CodPagamento() As Double
    CodPagamento = m_CodPagamento
End Property

Public Property Let CodPagamento(ByVal Value As Double)
    m_CodPagamento = Value
End Property

Public Property Get CodNotificacao() As Double
    CodNotificacao = m_CodNotificacao
End Property

Public Property Let CodNotificacao(ByVal Value As Double)
    m_CodNotificacao = Value
End Property

Public Function Buscar(CodNotificacao As Double, CodPagamento As Double) As Boolean
    Dim Sql As String, Rs As VSRecordset
    Sql = "Select * from TAB_PAGAMENTO_NOTIFICACAO where TPN_TNO_COD_NOTIFICACAO = '" & CodNotificacao & "' and TPN_TGT_COD_PAGAMENTO = '" & CodPagamento & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        Buscar = True
        m_CodNotificacao = "" & Rs!tpn_TNO_COD_NOTIFICACAO
        m_CodPagamento = "" & Rs!tpn_TGT_COD_PAGAMENTO
        m_Valor = "" & Rs!tpn_SUB_VALOR
        m_CodImposto = "" & Rs!tpn_TIP_COD_IMPOSTO
    End If
End Function

Public Function Gravar() As Boolean
    Dim campos As String, valores As String, condicao As String
    campos = "tpn_TNO_COD_NOTIFICACAO, tpn_TGT_COD_PAGAMENTO, tpn_SUB_VALOR, tpn_TIP_COD_IMPOSTO"
    valores = Bdados.PreparaValor(m_CodNotificacao, m_CodPagamento, m_Valor, m_CodImposto)
    condicao = "TPN_TNO_COD_NOTIFICACAO = '" & m_CodNotificacao & "' and TPN_TGT_COD_PAGAMENTO = '" & m_CodPagamento & "'"
    If Bdados.GravaDados("TAB_PAGAMENTO_NOTIFICACAO", valores, campos, condicao) Then
        Gravar = True
    End If
End Function
