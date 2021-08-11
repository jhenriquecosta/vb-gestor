VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eRepresentante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Im As String
Private m_CPF As String
Private m_Nome As String
Private m_TipoLogr As String
Private m_Logr As String
Private m_Numero As String
Private m_Complemento As String
Private m_Bairro As String
Private m_Telefone As String
Private m_Cidade As String
Private m_Uf As String
Private m_ImRepresentante As String

Public Property Get ImRepresentante() As String
    ImRepresentante = m_ImRepresentante
End Property

Public Property Let ImRepresentante(ByVal Value As String)
    m_ImRepresentante = Value
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

Public Property Get Telefone() As String
    Telefone = m_Telefone
End Property

Public Property Let Telefone(ByVal Value As String)
    m_Telefone = Value
End Property

Public Property Get Bairro() As String
    Bairro = m_Bairro
End Property

Public Property Let Bairro(ByVal Value As String)
    m_Bairro = Value
End Property

Public Property Get Complemento() As String
    Complemento = m_Complemento
End Property

Public Property Let Complemento(ByVal Value As String)
    m_Complemento = Value
End Property

Public Property Get Numero() As String
    Numero = m_Numero
End Property

Public Property Let Numero(ByVal Value As String)
    m_Numero = Value
End Property

Public Property Get Logr() As String
    Logr = m_Logr
End Property

Public Property Let Logr(ByVal Value As String)
    m_Logr = Value
End Property

Public Property Get TipoLogr() As String
    TipoLogr = m_TipoLogr
End Property

Public Property Let TipoLogr(ByVal Value As String)
    m_TipoLogr = Value
End Property

Public Property Get Nome() As String
    Nome = m_Nome
End Property

Public Property Let Nome(ByVal Value As String)
    m_Nome = Value
End Property

Public Property Get Cpf() As String
    Cpf = m_CPF
End Property

Public Property Let Cpf(ByVal Value As String)
    m_CPF = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Function Buscar(Im As String, Optional Cpf As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    If Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "select * from tab_representante_legal_emp where trl_tci_im = '" & Im & "'"
    Else
        'Sql = "select * from tab_representante_legal_empresa where trl_tci_im = '" & Im & "'"
        Sql = "select * from tab_representante_legal_emp where trl_tci_im = '" & Im & "'"
    End If
    If Trim(Cpf) <> "" Then
        Sql = Sql & " and trl_cpf = '" & Cpf & "'"
    End If
    If Bdados.AbreTabela(Sql, rs) Then
        Buscar = True
        m_Im = "" & rs!trl_tci_im
        m_CPF = "" & rs!trl_cpf
        m_Nome = "" & rs!trl_nome
        m_TipoLogr = "" & rs!trl_tip_logr
        m_Logr = "" & rs!trl_logr
        m_Numero = "" & rs!trl_num
        m_Complemento = "" & rs!trl_complemento
        m_Bairro = "" & rs!trl_bairro
        m_Telefone = "" & rs!trl_telefone
        m_Cidade = "" & rs!trl_cidade
        m_Uf = "" & rs!trl_uf
        m_ImRepresentante = "" & rs!trl_im_representante
    End If
End Function

Public Function Salvar() As Boolean
    Dim Campos As String
    Dim Valores As String
    Dim Condicao As String
    Campos = "trl_tci_im, trl_cpf, trl_nome, trl_tip_logr, trl_logr, trl_num, trl_complemento, trl_bairro, trl_telefone, trl_cidade, trl_uf, trl_im_representante"
    Valores = Bdados.PreparaValor(m_Im, Bdados.Converte(m_CPF, tctexto), m_Nome, m_TipoLogr, m_Logr, m_Numero, m_Complemento, m_Bairro, m_Telefone, m_Cidade, m_Uf, m_ImRepresentante)
    Condicao = " trl_tci_im = '" & m_Im & "' and trl_cpf = '" & m_CPF & "'"
    If Bdados.GravaDados("tab_representante_legal_emp", Valores, Campos, Condicao) Then Salvar = True
End Function


Public Function Excluir(Im As String, Optional Cpf As String) As Boolean
    Dim Condicao As String
    Condicao = " trl_tci_im = '" & Im & "'"
    If Trim(Cpf) <> "" Then
        Condicao = Condicao & "  and trl_cpf = '" & Cpf & "'"
    End If
    If Bdados.Conexao.FormatoBanco = oracle Then
        If Bdados.DeletaDados("tab_representante_legal_emp", Condicao) Then Excluir = True
    ElseIf Bdados.Conexao.FormatoBanco = SQLServer Then
        If Bdados.DeletaDados("tab_representante_legal_empresa", Condicao) Then Excluir = True
    End If
End Function
