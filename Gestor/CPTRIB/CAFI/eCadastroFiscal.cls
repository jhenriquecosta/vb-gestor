VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eCadastroFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Im As String
Private m_Icad As String
Private m_Tipo_Cadastro As Integer
Private m_Status As Integer
Private m_Data_Cadastro As String
Private m_Data_Cancelamento As String
Public Enum ecTipoCadastro
    ecAnuncio = 1
    ecAparelhoTransporte = 2
    ecMobiliario = 3
    ecAmbulanteEventual = 4
    ecHorarioEspecial = 5
    ecImobiliario = 6
    ecMaquinaMotor = 7
    ecObraParticular = 8
    ecOcupacaoPermanencia = 9
    ecSanitario = 10
    ecTransportePassageiro = 11
    ecUtilizacaoPassagem = 12
End Enum
Public Enum ecStstus
        ecAtivo = 1
        ecCancelado = 2
End Enum


    
Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let icad(ByVal Value As String)
    m_Icad = Value
End Property

Public Property Get icad() As String
    icad = m_Icad
End Property

Public Property Let Tipo_Cadastro(ByVal Value As String)
    m_Tipo_Cadastro = Value
End Property

Public Property Get Tipo_Cadastro() As String
    Tipo_Cadastro = m_Tipo_Cadastro
End Property

Public Property Let Status(ByVal Value As String)
    m_Status = Value
End Property

Public Property Get Status() As String
    Status = m_Status
End Property

Public Property Let Data_Cadastro(ByVal Value As String)
    m_Data_Cadastro = Value
End Property

Public Property Get Data_Cadastro() As String
    Data_Cadastro = m_Data_Cadastro
End Property

Public Property Let Data_Cancelamento(ByVal Value As String)
    m_Data_Cancelamento = Value
End Property

Public Property Get Data_Cancelamento() As String
    Data_Cancelamento = m_Data_Cancelamento
End Property
    

Public Function PreencherGrd(Grid As Object, icad As String, Im As String) As Boolean
    Dim Sql As String
    Dim condicao As String
    condicao = ""
    
    Sql = "SELECT * FROM TAB_CADASTRO_FISCAL"
    If icad = "" Then
        If condicao = "" Then condicao = condicao & " TCF_COD_CADASTRO = & ICad"
        Else: condicao = condicao & "AND  TCF_COD_CADASTRO = & ICad"
    End If
    If Im = "" Then
        If condicao = "" Then condicao = condicao & " TCF_COD_CADASTRO = '" & Im & "'"
        Else: condicao = condicao & "AND  TCF_COD_CADASTRO = '" & Im & "'"
    End If
    If condicao <> "" Then Sql = Sql & Where & condicao
    If Grid.Preencher(Bdados, Sql) Then PreencherGrd = True
End Function

Public Function Buscar(icad As String, Im As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Dim condicao As String
    condicao = ""
    
    Sql = "SELECT * FROM TAB_CADASTRO_FISCAL"
    If icad = "" Then
        If condicao = "" Then condicao = condicao & " TCF_COD_CADASTRO = & ICad"
        Else: condicao = condicao & "AND  TCF_COD_CADASTRO = & ICad"
    End If
    If Im = "" Then
        If condicao = "" Then condicao = condicao & " TCF_COD_CADASTRO = '" & Im & "'"
        Else: condicao = condicao & "AND  TCF_COD_CADASTRO = '" & Im & "'"
    End If
    If condicao <> "" Then Sql = Sql & Where & condicao
    If Bdados.AbreTabela(Sql, rs) Then
          m_Icad = rs!TCF_COD_CADASTRO
          m_Im = rs!TCF_TCI_IM
          m_Tipo_Cadastro = rs!TCF_TIPO_CADASTRO
          m_Status = rs!TCF_STATUS
          m_Data_Cadastro = rs!TCF_DATA_CADASTRO
          m_Data_Cancelamento = rs!TCF_DATA_CANCELAMENTO
        Buscar = True
    End If
End Function

Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
      
     campos = "TCF_COD_CADASTRO,TCF_TCI_IM,TCF_TIPO_CADASTRO,TCF_STATUS ,TCF_DATA_CADASTRO"
    valores = Bdados.PreparaValor(m_Icad, Bdados.Converte(m_Im, tctexto), m_Tipo_Cadastro, m_Status, m_Data_Cadastro)
    condicao = "TCF_COD_CADASTRO = " & icad & " AND  TCF_TCI_IM = '" & Im & "'"
    If Bdados.GravaDados("TAB_CADASTRO_FISCAL", valores, campos, condicao) Then Salvar = True
End Function

Public Function Baixa() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
      
     campos = "TCF_STATUS ,TCF_DATA_CANCELAMENTO"
    valores = Bdados.PreparaValor(m_Status, m_Data_Cancelamento)
    condicao = "TCF_COD_CADASTRO = " & icad & " AND  TCF_TCI_IM = '" & Im & "'"
    If Bdados.AtualizaDados("TAB_CADASTRO_FISCAL", valores, campos, condicao) Then Baixa = True
End Function
