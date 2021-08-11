VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cContribuinteAvulso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Identidade As String
Private m_Nome As String
Private m_Endereco As String
Private m_Bairro As String
Private m_Cidade As String
Private m_Uf As String
Private m_Cep As String
Private m_Telefone As String
Private m_CodUsuario As String
Private m_AtividadeEconomica As Long

Public Property Get AtividadeEconomica() As Long
    AtividadeEconomica = m_AtividadeEconomica
End Property

Public Property Let AtividadeEconomica(ByVal Value As Long)
    m_AtividadeEconomica = Value
End Property

Public Property Get CodUsuario() As String
    CodUsuario = m_CodUsuario
End Property

Public Property Let CodUsuario(ByVal Value As String)
    m_CodUsuario = Value
End Property

Public Property Get Telefone() As String
    Telefone = m_Telefone
End Property

Public Property Let Telefone(ByVal Value As String)
    m_Telefone = Value
End Property

Public Property Get Cep() As String
    Cep = m_Cep
End Property

Public Property Let Cep(ByVal Value As String)
    m_Cep = Value
End Property

Public Property Get Uf() As String
    Uf = m_Uf
End Property

Public Property Let Uf(ByVal Value As String)
    m_Uf = Value
End Property

Public Property Get Cidade() As String
    Cidade = m_Cidade
End Property

Public Property Let Cidade(ByVal Value As String)
    m_Cidade = Value
End Property

Public Property Get Bairro() As String
    Bairro = m_Bairro
End Property

Public Property Let Bairro(ByVal Value As String)
    m_Bairro = Value
End Property

Public Property Get Endereco() As String
    Endereco = m_Endereco
End Property

Public Property Let Endereco(ByVal Value As String)
    m_Endereco = Value
End Property

Public Property Get Nome() As String
    Nome = m_Nome
End Property

Public Property Let Nome(ByVal Value As String)
    m_Nome = Value
End Property

Public Property Get Identidade() As String
    Identidade = m_Identidade
End Property

Public Property Let Identidade(ByVal Value As String)
    m_Identidade = Value
End Property

Public Function Buscar(Identidade As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "select * from tab_contribuinte_avulso where tca_identidade = '" & Identidade & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        m_Identidade = "" & rs!tca_identidade
        m_Nome = "" & rs!tca_nome
        m_Endereco = "" & rs!tca_endereco
        m_Bairro = "" & rs!tca_bairro
        m_Cidade = "" & rs!tca_cidade
        m_Uf = "" & rs!tca_uf
        m_Cep = "" & rs!tca_cep
        m_Telefone = "" & rs!tca_telefone
        m_CodUsuario = "" & rs!tca_tus_cod_usuario
        m_AtividadeEconomica = "" & rs!tca_tae_cae
        Buscar = True
    End If
End Function

Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    valores = Bdados.PreparaValor(m_Identidade, m_Nome, m_Endereco, m_Bairro, m_Cidade, m_Uf, m_Cep, m_Telefone, m_CodUsuario, m_AtividadeEconomica)
    campos = "tca_identidade, tca_nome, tca_endereco, tca_bairro, tca_cidade, tca_uf, tca_cep, tca_telefone, tca_tus_cod_usuario, tca_tae_cae"
    If Bdados.InsereDados("Tab_contribuinte_avulso", valores, campos) Then
        Salvar = True
    End If
End Function
