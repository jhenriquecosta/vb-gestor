VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eTransportador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Im As String
Private m_Veiculo As String
Private m_Marca As String
Private m_CodModelo As String
Private m_AnoFabricacao As Integer
Private m_Placa As String
Private m_Municipio As String
Private m_Uf As String
Private m_Licensa As String
Private m_Atividade As String
Private m_Chassi As String
Private m_Inicio_Atividade_Veiculo As Date

Public Property Get IniAtividadeCarro() As Date
    IniAtividadeCarro = m_Inicio_Atividade_Veiculo
End Property

Public Property Let IniAtividadeCarro(ByVal Value As Date)
    m_Inicio_Atividade_Veiculo = Value
End Property


Public Property Get Chassi() As String
    Chassi = m_Chassi
End Property

Public Property Let Chassi(ByVal Value As String)
    m_Chassi = Value
End Property

Public Property Get Licensa() As String
    Licensa = m_Licensa
End Property

Public Property Let Licensa(ByVal Value As String)
    m_Licensa = Value
End Property

Public Property Get atividade() As String
    atividade = m_Atividade
End Property

Public Property Let atividade(ByVal Value As String)
    m_Atividade = Value
End Property

Public Property Get Uf() As String
    Uf = m_Uf
End Property

Public Property Let Uf(ByVal Value As String)
    m_Uf = Value
End Property

Public Property Get municipio() As String
    municipio = m_Municipio
End Property

Public Property Let municipio(ByVal Value As String)
    m_Municipio = Value
End Property

Public Property Get Placa() As String
    Placa = m_Placa
End Property

Public Property Let Placa(ByVal Value As String)
    m_Placa = Value
End Property

Public Property Get AnoFabricacao() As Integer
    AnoFabricacao = m_AnoFabricacao
End Property

Public Property Let AnoFabricacao(ByVal Value As Integer)
    m_AnoFabricacao = Value
End Property

Public Property Get CodModelo() As String
    CodModelo = m_CodModelo
End Property

Public Property Let CodModelo(ByVal Value As String)
    m_CodModelo = Value
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

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property


Public Function Buscar(Im As String, Placa As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "select * from tab_transportador_veiculo where TTV_TCI_IM = '" & Im & "' TTV_PLACA = '" & Placa & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        m_Im = "" & Rs!TTV_TCI_IM
        m_Veiculo = "" & Rs!TTV_VEICULO
        m_Marca = "" & Rs!TTV_MARCA
        m_CodModelo = "" & Rs!TTV_COD_MODELO
        m_AnoFabricacao = "" & Rs!TTV_ANO_FABRICACAO
        m_Placa = "" & Rs!TTV_PLACA
        m_Municipio = "" & Rs!TTV_MUNICIPIO
        m_Uf = "" & Rs!TTV_UF
        m_Licensa = "" & Rs!TTV_LICENCA
        m_Chassi = "" & Rs!TTV_chassi
        Buscar = True
    End If
End Function

Public Function Salvar() As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Campos = "TTV_TCI_IM, TTV_VEICULO, TTV_MARCA, TTV_COD_MODELO, TTV_ANO_FABRICACAO, TTV_PLACA, TTV_MUNICIPIO, TTV_UF, TTV_LICENCA, TTV_CHASSI,TTV_TAE_CAE,TTV_INICIO_ATIVIDADE"
    Valores = Bdados.PreparaValor(m_Im, m_Veiculo, m_Marca, m_CodModelo, m_AnoFabricacao, m_Placa, m_Municipio, m_Uf, m_Licensa, Nvl(m_Chassi, 0), m_Atividade, m_Inicio_Atividade_Veiculo)
    Condicao = "TTV_TCI_IM = '" & Im & "' and TTV_PLACA = '" & m_Placa & "'"
    If Bdados.GravaDados("tab_transportador_veiculo", Valores, Campos, Condicao) Then Salvar = True
End Function

Public Function Excluir(Im As String, Optional Placa As String) As Boolean
    Dim Condicao As String
    Condicao = "TTV_TCI_IM = '" & Im & "'"
    If Trim(Placa) <> "" Then
        Condicao = Condicao & " and TTV_PLACA = '" & Placa & "'"
    End If
    If Bdados.DeletaDados("tab_transportador_veiculo", Condicao) Then Excluir = True
End Function


Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = " select TTV_VEICULO as Veículo, "
        Sql = Sql & " TTV_MARCA AS Marca, "
        Sql = Sql & " TTV_COD_MODELO as Modelo,"
        Sql = Sql & " TTV_ANO_FABRICACAO as [Ano Fabricação], "
        Sql = Sql & " TTV_PLACA as Placa, "
        Sql = Sql & " TTV_CHASSI As Chassi,"
        Sql = Sql & " TTV_MUNICIPIO as Cidade, "
        Sql = Sql & " TTV_UF as Uf, TTV_LICENCA as Licença, "
        Sql = Sql & " ltrim(str(TAE_CAE)) + ' - ' +  tae_nome as Atividade,TTV_INICIO_ATIVIDADE AS [Ini Atividade]"
        Sql = Sql & " From Tab_atividade_Economica,"
        Sql = Sql & " TAB_TRANSPORTADOR_VEICULO"
        Sql = Sql & " where  TAE_CAE = TTV_TAE_CAE"
        Sql = Sql & " and  TTV_TCI_IM = '" & Im & "'"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = " select TTV_VEICULO as Veículo, "
        Sql = Sql & " TTV_MARCA AS Marca, "
        Sql = Sql & " TTV_COD_MODELO as Modelo,"
        Sql = Sql & " TTV_ANO_FABRICACAO as Ano_Fabricação, "
        Sql = Sql & " TTV_PLACA as Placa, "
        Sql = Sql & " TTV_CHASSI As Chassi,"
        Sql = Sql & " TTV_MUNICIPIO as Cidade, "
        Sql = Sql & " TTV_UF as Uf, TTV_LICENCA as Licença, "
        Sql = Sql & " LTRIM(TO_CHAR(TAE_CAE)) " & Bdados.Concatena & "' - '" & Bdados.Concatena & "  tae_nome as Atividade,TTV_INICIO_ATIVIDADE AS Ini_Atividade"
        Sql = Sql & " From Tab_atividade_Economica,"
        Sql = Sql & " TAB_TRANSPORTADOR_VEICULO"
        Sql = Sql & " where  TAE_CAE = TTV_TAE_CAE"
        Sql = Sql & " and  TTV_TCI_IM = '" & Im & "'"
    End If
    If Grid.Preencher(Bdados, Sql) Then PreencherGrd = True
End Function

Public Function VerificaChassi(Chassi As String, RetIm As String) As Boolean
    Dim Sql As String, Rs As VSRecordset
    Sql = "SELECT TTV_TCI_IM FROM TAB_TRANSPORTADOR_VEICULO WHERE ttv_chassi ='" & Chassi & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        RetIm = "" & Rs!TTV_TCI_IM
        VerificaChassi = True
    End If
End Function

Public Function GravarHistorico(CodMudanca As String, Im As String) As Boolean
    Dim Sql As String
    Sql = "Select TAB_TRANSPORTADOR_VEICULO.* ," & CodMudanca & " from TAB_TRANSPORTADOR_VEICULO where TTV_TCI_IM = '" & Im & "'"
    If Bdados.Executa("Insert into TAB_TRANSPORTADOR_VEICULO_HISTORICO " & Sql) Then GravarHistorico = True
End Function
