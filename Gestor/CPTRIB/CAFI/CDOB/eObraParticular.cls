VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eObraParticular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Icod As String
Private m_TipoIntervencao As Integer
Private m_DataInicio As String
Private m_Previsao As String
Private m_AreaAtingida As String
Private m_Motivo As String
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

Public Property Let TipoIntervencao(ByVal Value As String)
    m_TipoIntervencao = Value
End Property

Public Property Get TipoIntervencao() As String
    TipoIntervencao = m_TipoIntervencao
End Property

Public Property Let DataInicio(ByVal Value As String)
    m_DataInicio = Value
End Property

Public Property Get DataInicio() As String
    DataInicio = m_DataInicio
End Property

Public Property Let PREVISAO(ByVal Value As String)
    m_Previsao = Value
End Property

Public Property Get PREVISAO() As String
    PREVISAO = m_Previsao
End Property

Public Property Let AreaAtingida(ByVal Value As String)
    m_AreaAtingida = Value
End Property

Public Property Get AreaAtingida() As String
    AreaAtingida = m_AreaAtingida
End Property

Public Property Let Motivo(ByVal Value As String)
    m_Motivo = Value
End Property

Public Property Get Motivo() As String
    Motivo = m_Motivo
End Property
    
Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    Dim codCadastro As String
    Set Codigo = New ContaCorrente
    Set cadastro = New eCadastroFiscal
       
    If (m_Icod = "") Then
        codCadastro = Codigo.GeraCodPagamento(47)
        With cadastro
            .icad = codCadastro
            .Im = m_Im
            .Tipo_Cadastro = ecObraParticular
            .Status = ecAtivo
            .Data_Cadastro = Date
        End With
        If Not cadastro.Salvar Then Exit Function
    Else
        codCadastro = m_Icod
    End If
     
        campos = "TOP_COD_CADASTRO , TOP_TIPO_INTERVENCAO, TOP_DATA_INICIO_OBRA, TOP_PREVISAO, TOP_AREA_ATINGIDA, TOP_MOTIVO"
        valores = Bdados.PreparaValor(codCadastro, m_TipoIntervencao, m_DataInicio, m_Previsao, m_AreaAtingida, m_Motivo)
        condicao = "TOP_COD_CADASTRO = " & codCadastro
        Bdados.AbreTrans
        If Bdados.GravaDados("TAB_CAD_OBRA_PARTICULAR", valores, campos, condicao) Then
            Salvar = True
            Bdados.GravaTrans
        Else
            Bdados.CancelaTrans
        End If
    
End Function
    
Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    Sql = " SELECT TOP_COD_CADASTRO as Código, "
    Sql = Sql & " TGE_NOME as Intervenção,"
    Sql = Sql & " TOP_DATA_INICIO_OBRA as Início_Obra,"
    Sql = Sql & " TOP_PREVISAO as Previsão, "
    Sql = Sql & " TOP_AREA_ATINGIDA as Área_Atingida, "
    Sql = Sql & " TOP_MOTIVO as Motivo, "
    Sql = Sql & " TOP_TIPO_INTERVENCAO,"
    Sql = Sql & " TCF_TCI_IM"
    Sql = Sql & " From VIS_CAD_OBRA"
    Sql = Sql & " WHERE TCF_STATUS= " & ecAtivo & " "
    If Im <> "" Then Sql = Sql & " AND   TCF_TCI_IM = '" & Im & "'"
    
    If Not Grid.Preencher(Bdados, Sql, 1300, 2000, 1200, 1200, 1500, 4000, 0, 0) Then
        Util.Avisa "Contribuinte não possui Cadastro de Obra Particular."
    End If
End Function


Public Function PreencherGrdConsulta(Grid As Object, Im As String, cod As String, TipoIntervencao As String, StatusCad As String) As Boolean
    Dim Sql As String
    Dim condicao As String
    
    If (cod <> "") Then condicao = condicao & " AND TOP_COD_CADASTRO  = " & cod
    If (TipoIntervencao <> "") Then condicao = condicao & " AND TOP_TIPO_INTERVENCAO =  " & TipoIntervencao
    If (StatusCad <> "") Then condicao = condicao & " AND TCF_STATUS = " & StatusCad
    
    Sql = " SELECT TOP_COD_CADASTRO as Código, "
    Sql = Sql & " tci_nome as Contribuinte, "
    Sql = Sql & " TCF_DATA_CADASTRO as Data_Cadastro, "
    Sql = Sql & " STATUS as Status, "
    Sql = Sql & " TCF_DATA_CANCELAMENTO  as Data_Baixa, "
    Sql = Sql & " TGE_NOME as Intervenção, "
    Sql = Sql & " TOP_DATA_INICIO_OBRA as Data_Início, "
    Sql = Sql & " TOP_PREVISAO as Previsão, "
    Sql = Sql & " TOP_AREA_ATINGIDA as Área_Atingida, "
    Sql = Sql & " TOP_MOTIVO as Motivo"
    Sql = Sql & " FROM VIS_CAD_OBRA where  1=1"
    If Im <> "" Then Sql = Sql & " and  TCF_TCI_IM = '" & Im & "'"
    Sql = Sql & condicao
    If Not Grid.Preencher(Bdados, Sql, 1100, 5000, 1100, 1200, 1150, 3000, 1200, 1200, 2000, 4000) Then
        Util.Avisa "Contribuinte não possui Cadastro de Obra Particular."
    End If
End Function
    

