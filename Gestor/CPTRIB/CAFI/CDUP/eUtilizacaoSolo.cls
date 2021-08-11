VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eUtilizacaoSolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Icod As String
Private m_Destinacao As String
Private m_Equipamento As String
Private m_AreaOcupada As String

Private m_Im As String
Public cadastro As New eCadastroFiscal
Dim Codigo As New ContaCorrente
    
Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property
    
Public Property Let Icod(ByVal Value As String)
    m_Icod = Value
End Property

Public Property Get Icod() As String
    Icod = m_Icod
End Property

Public Property Let Destinacao(ByVal Value As String)
    m_Destinacao = Value
End Property

Public Property Get Destinacao() As String
    Destinacao = m_Destinacao
End Property

Public Property Let Equipamento(ByVal Value As String)
    m_Equipamento = Value
End Property

Public Property Get Equipamento() As String
    Equipamento = m_Equipamento
End Property

Public Property Let AreaOcupada(ByVal Value As String)
    m_AreaOcupada = Value
End Property

Public Property Get AreaOcupada() As String
    AreaOcupada = m_AreaOcupada
End Property


Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    Dim codCadastro As String
    Set Codigo = New ContaCorrente
    Set cadastro = New eCadastroFiscal
       
    If (m_Icod = "") Then
        codCadastro = Codigo.GeraCodPagamento(51)
        With cadastro
            .icad = codCadastro
            .Im = m_Im
            .Tipo_Cadastro = ecUtilizacaoPassagem
            .Status = ecAtivo
            .Data_Cadastro = Date
        End With
        If Not cadastro.Salvar Then Exit Function
    Else
        codCadastro = m_Icod
    End If
     
        campos = "TUS_COD_CADASTRO , TUS_DESTINACAO, TUS_EQUIPAMENTO, TUS_AREA_OCUPADA"
        valores = Bdados.PreparaValor(codCadastro, m_Destinacao, m_Equipamento, m_AreaOcupada)
        condicao = "TUS_COD_CADASTRO = " & codCadastro
        Bdados.AbreTrans
        If Bdados.GravaDados("TAB_CAD_UTILIZACAO_SUBSOLO", valores, campos, condicao) Then
        Salvar = True

            Bdados.GravaTrans
        Else
            Bdados.CancelaTrans
        End If
    
End Function

Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    Sql = "SELECT  TUS_COD_CADASTRO AS Código, "
    Sql = Sql & " TUS_DESTINACAO AS Destinação, "
    Sql = Sql & " TUS_EQUIPAMENTO AS Equipamento, "
    Sql = Sql & " TUS_AREA_OCUPADA As Área_Ocupada,"
    Sql = Sql & " TCF_TCI_IM"
    Sql = Sql & " From VIS_CAD_UTILIZACAO_SUBSOLO"
    Sql = Sql & " WHERE TCF_STATUS= " & ecAtivo & " "
    If Im <> "" Then Sql = Sql & " AND   TCF_TCI_IM = '" & Im & "'"
    
    If Not Grid.Preencher(Bdados, Sql, 1200, 3880, 3815, 1500, 0) Then
        Util.Avisa "Contribuinte não possui cadastro de Utilização e de Passagem em Subsolo e Solo d Logradouros Públicos."
    End If
End Function


Public Function PreencherGrdConsulta(Grid As Object, Im As String, cod As String, StatusCad As String) As Boolean
    Dim Sql As String
    Dim condicao As String

    If (cod <> "") Then condicao = condicao & " AND TUS_COD_CADASTRO = " & cod
    If (StatusCad <> "") Then condicao = condicao & " AND TCF_STATUS = " & StatusCad

    Sql = "SELECT TUS_COD_CADASTRO AS Código, "
    Sql = Sql & " tci_nome AS Nome, "
    Sql = Sql & " TCF_DATA_CADASTRO as Data_Cadastro, "
    Sql = Sql & " STATUS as Status, "
    Sql = Sql & " TCF_DATA_CANCELAMENTO as Data_Cancelamento,"
    Sql = Sql & " TUS_DESTINACAO AS Destinação, "
    Sql = Sql & " TUS_EQUIPAMENTO AS Equipamento, "
    Sql = Sql & " TUS_AREA_OCUPADA AS Área_Ocupada"
    Sql = Sql & " From VIS_CAD_UTILIZACAO_SUBSOLO"
    Sql = Sql & " WHERE 1=1"
    If Im <> "" Then Sql = Sql & " and TCF_TCI_IM  = '" & Im & "'"
    Sql = Sql & condicao
    
    If Not Grid.Preencher(Bdados, Sql, 1100, 5000, 1100, 1200, 1150, 3880, 3815, 1500) Then
     Util.Avisa "Contribuinte não possui cadastro de Utilização e de Passagem em Subsolo e Solo d Logradouros Públicos."
    End If
End Function

