VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAforManu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Livro As Double
Private m_FolhaAtual As Double
Private m_FolhaTotal As Double
Private m_Status As Integer
Private m_CodUsuario As String
Private m_DataAbertura As String

Public Property Get DataAbertura() As String
    DataAbertura = m_DataAbertura
End Property

Public Property Let DataAbertura(ByVal Value As String)
    m_DataAbertura = Value
End Property

Public Property Get CodUsuario() As String
    CodUsuario = m_CodUsuario
End Property

Public Property Let CodUsuario(ByVal Value As String)
    m_CodUsuario = Value
End Property

Public Property Get Status() As Integer
    Status = m_Status
End Property

Public Property Let Status(ByVal Value As Integer)
    m_Status = Value
End Property

Public Property Get FolhaTotal() As Double
    FolhaTotal = m_FolhaTotal
End Property

Public Property Let FolhaTotal(ByVal Value As Double)
    m_FolhaTotal = Value
End Property

Public Property Get FolhaAtual() As Double
    FolhaAtual = m_FolhaAtual
End Property

Public Property Let FolhaAtual(ByVal Value As Double)
    m_FolhaAtual = Value
End Property

Public Property Get Livro() As Double
    Livro = m_Livro
End Property

Public Property Let Livro(ByVal Value As Double)
    m_Livro = Value
End Property

Public Function Buscar(Livro As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "SELECT TAM_LIVRO," & _
            " TAM_FOLHA_ATUAL," & _
            " TAM_FOLHA_TOTAL," & _
            " TAM_STATUS," & _
            " TAM_TUS_COD_USUARIO," & _
            " TAM_DATA_ABERTURA" & _
            " From TAB_AFORAMENTO_MANUTENCAO" & _
            " WHERE TAM_LIVRO = '" & Livro & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        Buscar = True
        m_Livro = "" & rs!TAM_LIVRO
        m_FolhaAtual = "" & rs!TAM_FOLHA_ATUAL
        m_FolhaTotal = "" & rs!TAM_FOLHA_TOTAL
        m_Status = "" & rs!TAM_STATUS
        m_CodUsuario = "" & rs!TAM_TUS_COD_USUARIO
        m_DataAbertura = "" & rs!TAM_DATA_ABERTURA
    End If
End Function

Public Function Gravar(Livro As String) As Boolean
    Dim Valores As String, Campos As String, condicao As String
    Campos = "TAM_LIVRO, TAM_FOLHA_ATUAL, TAM_FOLHA_TOTAL, TAM_STATUS, TAM_TUS_COD_USUARIO, TAM_DATA_ABERTURA"
    Valores = Bdados.PreparaValor(Livro, m_FolhaAtual, m_FolhaTotal, m_Status, m_CodUsuario, m_DataAbertura)
    condicao = "tam_livro = '" & Livro & "'"
    If Bdados.GravaDados("TAB_AFORAMENTO_MANUTENCAO", Valores, Campos, condicao) Then
        Gravar = True
    End If
End Function

Public Function PreencherGrid(grd As Object) As Boolean
    Dim Sql As String
    Sql = "SELECT TAM_LIVRO AS Livro," & _
            " TAM_FOLHA_ATUAL AS Folhas, " & _
            " TAM_FOLHA_TOTAL As Total, " & _
            " TAM_STATUS AS Situacao" & _
        " FROM TAB_AFORAMENTO_MANUTENCAO" & _
        " ORDER BY TAM_LIVRO"
    grd.Preencher Bdados, Sql, (grd.Width * 25 / 100), (grd.Width * 25 / 100), (grd.Width * 25 / 100), (grd.Width * 25 / 100)
    grd.Mensagem = "Situacao : [1] Aberto, [0] Fechado"
End Function


Public Function GravarCorrelativo(Ordem As Integer) As Boolean
    Dim Valores As String, Campos As String
    Campos = "tnc_ano, tnc_tipo_operacao, tnc_sequencia, tnc_descricao"
    Valores = Bdados.PreparaValor(0, 33, Ordem - 1, "ORDEM AFORAMENTO")
    If Bdados.GravaDados("TAB_NUM_CORRELATIVO", Valores, Campos, "tnc_tipo_operacao=33") Then
        GravarCorrelativo = True
    End If
End Function
