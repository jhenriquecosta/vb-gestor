VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eAparelhoTransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Icad As String
Private m_Veiculo As String
Private m_Marca As String
Private m_Modelo As String
Private m_AnoFabricacao As String
Private m_Placa As String
Private m_Municipio As String
Private m_Uf As String
Private m_Licenca As String
Private m_Atividade As String
Private m_Chassi As String
Private m_Inicio_Atividade_Veiculo As String
Private m_Complemento As String
Private m_Im As String
Public cadastro As New eCadastroFiscal
Dim Codigo As New ContaCorrente

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Property Get Complemento() As String
    Complemento = m_Complemento
End Property

Public Property Let Complemento(ByVal Value As String)
    m_Complemento = Value
End Property

Public Property Get IniAtividadeCarro() As String
    IniAtividadeCarro = m_Inicio_Atividade_Veiculo
End Property

Public Property Let IniAtividadeCarro(ByVal Value As String)
    m_Inicio_Atividade_Veiculo = Value
End Property


Public Property Get Chassi() As String
    Chassi = m_Chassi
End Property

Public Property Let Chassi(ByVal Value As String)
    m_Chassi = Value
End Property

Public Property Get Licenca() As String
    Licenca = m_Licenca
End Property

Public Property Let Licenca(ByVal Value As String)
    m_Licenca = Value
End Property

Public Property Get Atividade() As String
    Atividade = m_Atividade
End Property

Public Property Let Atividade(ByVal Value As String)
    m_Atividade = Value
End Property

Public Property Get UF() As String
    UF = m_Uf
End Property

Public Property Let UF(ByVal Value As String)
    m_Uf = Value
End Property

Public Property Get Municipio() As String
    Municipio = m_Municipio
End Property

Public Property Let Municipio(ByVal Value As String)
    m_Municipio = Value
End Property

Public Property Get Placa() As String
    Placa = m_Placa
End Property

Public Property Let Placa(ByVal Value As String)
    m_Placa = Value
End Property

Public Property Get AnoFabricacao() As String
    AnoFabricacao = m_AnoFabricacao
End Property

Public Property Let AnoFabricacao(ByVal Value As String)
    m_AnoFabricacao = Value
End Property

Public Property Get Modelo() As String
    Modelo = m_Modelo
End Property

Public Property Let Modelo(ByVal Value As String)
    m_Modelo = Value
End Property

Public Property Get Marca() As String
    Marca = m_Marca
End Property

Public Property Let Marca(ByVal Value As String)
    m_Marca = Value
End Property

Public Property Get Veiculo() As String
    Veiculo = m_Veiculo
End Property

Public Property Let Veiculo(ByVal Value As String)
    m_Veiculo = Value
End Property

Public Property Get icad() As String
    icad = m_Icad
End Property

Public Property Let icad(ByVal Value As String)
    m_Icad = Value
End Property

Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    Dim codCadastro As String
    Set Codigo = New ContaCorrente
       
    If (m_Icad = "") Then
        codCadastro = Codigo.GeraCodPagamento(50)
        With cadastro
            .icad = codCadastro
            .Im = m_Im
            .Tipo_Cadastro = ecAparelhoTransporte
            .STATUS = ecAtivo
            .Data_Cadastro = Date
        End With
        If Not cadastro.Salvar Then Exit Function
    Else
        codCadastro = m_Icad
    End If
     
        campos = " TAP_COD_CADASTRO,TAP_VEICULO,TAP_MARCA,TAP_MODELO,TAP_ANO_FABRICACAO,TAP_PLACA,TAP_MUNICIPIO,TAP_UF,TAP_LICENCA,TAP_CHASSI,TAP_ATIVIDADE,TAP_DATA_INICIO_ATIVIDADE,TAP_COMPLEMENTO"
        valores = Bdados.PreparaValor(codCadastro, m_Veiculo, m_Marca, m_Modelo, m_AnoFabricacao, m_Placa, m_Municipio, m_Uf, m_Licenca, m_Chassi, m_Atividade, m_Inicio_Atividade_Veiculo, m_Complemento)
        condicao = "TAP_COD_CADASTRO = " & codCadastro & ""
        Bdados.AbreTrans
        If Bdados.GravaDados("TAB_CAD_APARELHO_TRANSPORTE", valores, campos, condicao) Then
            Salvar = True
            Bdados.GravaTrans
        Else
            Bdados.CancelaTrans
        End If
    
End Function

Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    
    Sql = " SELECT TAP_COD_CADASTRO AS Código, "
    Sql = Sql & " Atividade AS Atividade,"
    Sql = Sql & " TAP_VEICULO AS Veículo,"
    Sql = Sql & " TAP_MARCA AS Marca,"
    Sql = Sql & " TAP_MODELO AS Modelo,"
    Sql = Sql & " TAP_ANO_FABRICACAO AS Ano,"
    Sql = Sql & " TAP_PLACA AS Placa,"
    Sql = Sql & " TAP_CHASSI AS Chassi,"
    Sql = Sql & " TAP_LICENCA AS Licença,"
    Sql = Sql & " TAP_MUNICIPIO AS Município, UF,"
    Sql = Sql & "  TAP_DATA_INICIO_ATIVIDADE AS Data_Início_Atividade,"
    Sql = Sql & " TAP_COMPLEMENTO AS Complemento,"
    Sql = Sql & " TAP_UF,"
    Sql = Sql & " TAP_ATIVIDADE,"
    Sql = Sql & " TCF_TCI_IM"
    Sql = Sql & " From VIS_CAD_APARELHO_TRANSPORTE"
    Sql = Sql & " WHERE TCF_STATUS= " & ecAtivo & " "
    
    If Im <> "" Then Sql = Sql & " AND   TCF_TCI_IM = '" & Im & "'"
    
    If Not Grid.Preencher(Bdados, Sql, 1300, 5000, 2500, 2500, 2500, 1000, 2000, 1000, 2000, 3000, 1500, 2000, 4000, 0, 0, 0) Then
        If Im <> "" Then
            Util.Avisa "Contribuinte não possui Cadastro de Ambulante, Evenual e Feirante."
        Else
            Util.Avisa "Buscar sem resultados."
        End If
    End If
End Function

Public Function PreencherGrdConsulta(Grid As Object, Im As String, cod As String, PlacaCad As String, ChassiCad As String, StatusCad As String) As Boolean
    Dim Sql As String
    Dim condicao As String
    
    If (cod <> "") Then condicao = condicao & " AND TAP_COD_CADASTRO = " & cod
    If (PlacaCad <> "") Then condicao = condicao & " AND TAP_PLACA =  '" & PlacaCad & "'"
    If (ChassiCad <> "") Then condicao = condicao & " AND TAP_CHASSI =  '" & ChassiCad & "'"
    If (StatusCad <> "") Then condicao = condicao & " AND TCF_STATUS = " & StatusCad
    
    Sql = " SELECT TAP_COD_CADASTRO AS Código, "
    Sql = Sql & " tci_nome AS Contribuinte,"
    Sql = Sql & " TCF_DATA_CADASTRO AS Data_Cadastro,"
    Sql = Sql & " STATUS AS Status,"
    Sql = Sql & " TCF_DATA_CANCELAMENTO AS Data_Cancelamento,"
    Sql = Sql & " TAP_VEICULO AS Veículo,"
    Sql = Sql & " TAP_MARCA AS Marca,"
    Sql = Sql & " TAP_MODELO AS Modelo,"
    Sql = Sql & " TAP_ANO_FABRICACAO AS Ano,"
    Sql = Sql & " TAP_PLACA AS Placa,"
    Sql = Sql & " TAP_CHASSI AS Chassi,"
    Sql = Sql & " TAP_LICENCA AS Licença,"
    Sql = Sql & " UF AS UF,"
    Sql = Sql & " TAP_MUNICIPIO AS Município,"
    Sql = Sql & " TAP_ATIVIDADE AS Atividade,"
    Sql = Sql & " TAP_DATA_INICIO_ATIVIDADE AS Data_Início_Atividade,"
    Sql = Sql & " TAP_COMPLEMENTO As Complemento"
    Sql = Sql & " From VIS_CAD_APARELHO_TRANSPORTE where 1=1"
    If Im <> "" Then Sql = Sql & " and  TCF_TCI_IM = '" & Im & "'"
    
    If condicao <> "" Then Sql = Sql & condicao
    

    If Not Grid.Preencher(Bdados, Sql, 1100, 5000, 1100, 1200, 1150, 2000, 2000, 2000, 800, 2000, 2000, 2000, 1800, 1800, 3000, 1500, 4000, 0) Then
        Util.Avisa "Contribuinte não possui Cadastro de Aparelho e Transporte."
    End If
    
End Function

