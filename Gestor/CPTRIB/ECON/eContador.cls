VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eContador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Im As String
Private m_CRC As String
Private m_Contador As String
Private m_CPF As String
Private m_CGCEscritorio As String

Public Property Get CGCEscritorio() As String
    CGCEscritorio = m_CGCEscritorio
End Property

Public Property Let CGCEscritorio(ByVal Value As String)
    m_CGCEscritorio = Value
End Property

Public Property Get Cpf() As String
    Cpf = m_CPF
End Property

Public Property Let Cpf(ByVal Value As String)
    m_CPF = Value
End Property

Public Property Get Contador() As String
    Contador = m_Contador
End Property

Public Property Let Contador(ByVal Value As String)
    m_Contador = Value
End Property

Public Property Get Crc() As String
    Crc = m_CRC
End Property

Public Property Let Crc(ByVal Value As String)
    m_CRC = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Function Buscar(Im As String, Optional Crc As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "Select * from tab_contador_empresa where TCE_TCI_IM = '" & Im & "'"
    If Trim(Crc) <> "" Then
        Sql = Sql & " and TCE_CRC = '" & Crc & "'"
    End If
    If Bdados.AbreTabela(Sql, rs) Then
        Buscar = True
        m_Im = "" & rs!TCE_TCI_IM
        m_CRC = "" & rs!TCE_CRC
        m_Contador = "" & rs!TCE_CONTADOR
        m_CPF = "" & rs!TCE_CPF
        m_CGCEscritorio = "" & rs!tce_cgc_escritorio
    End If
End Function

Public Function Excluir(Im As String, Optional Crc As String) As Boolean
    Dim condicao As String
    condicao = "TCE_TCI_IM = '" & Im & "'"
    If Trim(Crc) <> "" Then
        condicao = condicao & " and TCE_CRC = '" & Crc & "'"
    End If
    If Bdados.DeletaDados("tab_contador_empresa", condicao) Then Excluir = True
End Function

Public Sub PreencherCboContador(Combo As Object)
    Dim Sql As String
    Sql = "SELECT tci_nome, tci_Im FROM TAB_contribuinte WHERE tci_tae_cae in (select tae_cae from tab_atividade_economica where tae_nome like '%CONTAB%')"
    Combo.Preencher Bdados, Sql
End Sub

Public Function Salvar() As Boolean
    Dim campos As String, valores As String, condicao As String
    campos = "TCE_TCI_IM, TCE_CRC, TCE_CONTADOR, TCE_CPF, TCE_CGC_ESCRITORIO"
    valores = Bdados.PreparaValor(Bdados.Converte(m_Im, tctexto), m_CRC, m_Contador, m_CPF, m_CGCEscritorio)
    condicao = "tce_tci_im = '" & m_Im & "' AND tce_crc = '" & m_CRC & "'"
    If Bdados.GravaDados("tab_contador_empresa", valores, campos, condicao) Then Salvar = True
End Function

Public Function GravarHistorico(CodMudanca As String, Im As String) As Boolean
    Dim Sql As String
    Sql = "Select TAB_CONTADOR_EMPRESA.* ," & CodMudanca & " from TAB_CONTADOR_EMPRESA where TCE_TCI_IM = '" & Im & "'"
    If Bdados.Executa("Insert into TAB_CONTADOR_EMPRESA_HISTORICO  " & Sql) Then GravarHistorico = True
End Function

