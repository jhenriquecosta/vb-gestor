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
Private m_Licensa As Integer
Private m_Chassi As String

Public Property Get Chassi() As String
    Chassi = m_Chassi
End Property

Public Property Let Chassi(ByVal Value As String)
    m_Chassi = Value
End Property

Public Property Get Licensa() As Integer
    Licensa = m_Licensa
End Property

Public Property Let Licensa(ByVal Value As Integer)
    m_Licensa = Value
End Property

Public Property Get Uf() As String
    Uf = m_Uf
End Property

Public Property Let Uf(ByVal Value As String)
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
        m_Chassi = "" & Rs!TTV_CHASSI
        Buscar = True
    End If
End Function

Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    campos = "TTV_TCI_IM, TTV_VEICULO, TTV_MARCA, TTV_COD_MODELO, TTV_ANO_FABRICACAO, TTV_PLACA, TTV_MUNICIPIO, TTV_UF, TTV_LICENCA, TTV_CHASSI"
    valores = Bdados.PreparaValor(m_Im, m_Veiculo, m_Marca, m_CodModelo, m_AnoFabricacao, m_Placa, m_Municipio, m_Uf, m_Licensa, m_Chassi)
    condicao = "TTV_TCI_IM = '" & Im & "' and TTV_chassi = '" & m_Chassi & "'"
    If Bdados.GravaDados("tab_transportador_veiculo", valores, campos, condicao) Then Salvar = True
End Function

Public Function Excluir(Im As String, Optional Placa As String) As Boolean
    Dim condicao As String
    condicao = "TTV_TCI_IM = '" & Im & "'"
    If Trim(Placa) <> "" Then
        condicao = condicao & " and TTV_PLACA = '" & Placa & "'"
    End If
    If Bdados.DeletaDados("tab_transportador_veiculo", condicao) Then Excluir = True
End Function


Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    Sql = "select TTV_VEICULO as Veículo," & _
            " TTV_MARCA AS Marca," & _
            " TTV_COD_MODELO as Modelo," & _
            " TTV_ANO_FABRICACAO as [Ano Fabricação]," & _
            " TTV_PLACA as Placa," & _
            " TTV_CHASSI As Chassi," & _
            " TTV_MUNICIPIO as Cidade," & _
            " TTV_UF as Uf," & _
            " TTV_LICENCA as Licença" & _
            " From TAB_TRANSPORTADOR_VEICULO" & _
            " where TTV_TCI_IM = '" & Im & "'"
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
