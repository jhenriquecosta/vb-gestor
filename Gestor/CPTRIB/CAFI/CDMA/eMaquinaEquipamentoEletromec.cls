VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eMaquinaEquipamentoEletromec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Icod As String
Private m_Serie As String
Private m_Tipo As String
Private m_Status As String
Private m_Localizacao As String
Private m_Im As String
Public cadastro As New eCadastroFiscal
Dim Codigo As New ContaCorrente
    
Public Property Let Icod(ByVal Value As String)
    m_Icod = Value
End Property

Public Property Get Icod() As String
    Icod = m_Icod
End Property
    
Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property


Public Property Let Localizacao(ByVal Value As String)
    m_Localizacao = Value
End Property

Public Property Get Localizacao() As String
    Localizacao = m_Localizacao
End Property

Public Property Let Horario1(ByVal Value As String)
    m_Horario1 = Value
End Property

Public Property Get Horario1() As String
    Horario1 = m_Horario1
End Property

Public Property Let Status(ByVal Value As String)
    m_Status = Value
End Property

Public Property Get Status() As String
    Status = m_Status
End Property

Public Property Let Tipo(ByVal Value As String)
    m_Tipo = Value
End Property

Public Property Get Tipo() As String
    Tipo = m_Tipo
End Property

Public Property Let Serie(ByVal Value As String)
    m_Serie = Value
End Property

Public Property Get Serie() As String
    Serie = m_Serie
End Property
    
    
Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    Dim codCadastro As String
    Set Codigo = New ContaCorrente
    Set cadastro = New eCadastroFiscal
       
    If (m_Icod = "") Then
        codCadastro = Codigo.GeraCodPagamento(46)
        With cadastro
            .icad = codCadastro
            .Im = m_Im
            .Tipo_Cadastro = ecMaquinaMotor
            .Status = ecAtivo
            .Data_Cadastro = Date
        End With
        If Not cadastro.Salvar Then Exit Function
    Else
        codCadastro = m_Icod
    End If
     
        campos = "TME_COD_CADASTRO,TME_SERIE,TME_TIPO,TME_STATUS,TME_LOCALIZACAO"
        valores = Bdados.PreparaValor(codCadastro, m_Serie, m_Tipo, m_Status, m_Localizacao)
        condicao = "TME_COD_CADASTRO =" & codCadastro
        Bdados.AbreTrans
        If Bdados.GravaDados("TAB_CAD_MAQUINA_EQUIPAMENTO", valores, campos, condicao) Then
        Salvar = True

            Bdados.GravaTrans
        Else
            Bdados.CancelaTrans
        End If
    
End Function

Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    Sql = " SELECT TME_COD_CADASTRO AS Código, "
    Sql = Sql & " TME_SERIE AS Série, "
    Sql = Sql & " Status_Equipamento ,"
    Sql = Sql & " Tipo AS Tipo, "
    Sql = Sql & " TME_LOCALIZACAO As Localização, "
    Sql = Sql & " TME_TIPO,"
    Sql = Sql & " TME_STATUS,"
    Sql = Sql & " TCF_TCI_IM"
    Sql = Sql & " From VIS_CAD_MAQUINA"
    Sql = Sql & " WHERE TCF_STATUS= " & ecAtivo & ""
    If Im <> "" Then Sql = Sql & "and  TCF_TCI_IM = '" & Im & "'"
    
    If Not Grid.Preencher(Bdados, Sql, 1300, 1500, 2000, 2000, 4000, 0, 0, 0) Then
        Util.Avisa "Contribuinte não possui Cadastro de Máquina e Equipamento Eletromecânico."
    End If
End Function


Public Function PreencherGrdConsulta(Grid As Object, Im As String, cod As String, SerieN As String, StatusCad As String) As Boolean
    Dim Sql As String
    Dim condicao As String

    If (cod <> "") Then condicao = condicao & " AND TME_COD_CADASTRO = " & cod
    If (SerieN <> "") Then condicao = condicao & " AND TME_SERIE = '" & SerieN & "'"
    If (StatusCad <> "") Then condicao = condicao & " AND TCF_STATUS = " & StatusCad
    
    Sql = " SELECT  TME_COD_CADASTRO AS Código, "
    Sql = Sql & " tci_nome AS Contribuinte, "
    Sql = Sql & " Status_Equipamento AS Status, "
    Sql = Sql & " TCF_DATA_CADASTRO AS Data_Cadastro, "
    Sql = Sql & " TCF_DATA_CANCELAMENTO AS Data_Baixa, "
    Sql = Sql & " TME_SERIE AS Série, "
    Sql = Sql & " Tipo AS Tipo, "
    Sql = Sql & " Status_Equipamento AS Status_Equipamento,"
    Sql = Sql & " TME_LOCALIZACAO As Localização"
    Sql = Sql & " From VIS_CAD_MAQUINA"
    Sql = Sql & " WHERE 1=1"
    If Im <> "" Then Sql = Sql & " and TCF_TCI_IM = '" & Im & "'"
    Sql = Sql & condicao
    If Not Grid.Preencher(Bdados, Sql, 1100, 5000, 1100, 1200, 1150, 1500, 2000, 2000, 4000, 0, 0) Then
     Util.Avisa "Contribuinte não possui Cadastro de Máquina e Equipamento Eletromecânico."
    End If
End Function

