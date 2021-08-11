VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Im As String
Private m_CPF As String
Private m_Nome As String
Private m_Cargo As String
Private m_TipoLogr As String
Private m_Logr As String
Private m_Numero As String
Private m_Complemento As String
Private m_Bairro As String
Private m_Telefone As String
Private m_Cidade As String
Private m_Uf As String

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

Public Property Get Cargo() As String
    Cargo = m_Cargo
End Property

Public Property Let Cargo(ByVal Value As String)
    m_Cargo = Value
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

Public Function Buscar(Optional Im As String, Optional Cpf As String) As Boolean
    Dim Sql As String, sqlaux As String
    Dim rs As VSRecordset
    Sql = "select * from tab_socio_empresa"
    If Trim(Im) <> "" Then
        sqlaux = sqlaux & " and tso_tci_im = '" & Im & "'"
    End If
    If Trim(Cpf) <> "" Then
        sqlaux = sqlaux & " and tso_cpf = '" & Cpf & "'"
    End If
    If Trim(sqlaux) <> "" Then Sql = Sql & " where " & Right(sqlaux, Len(sqlaux) - 4)
    If Bdados.AbreTabela(Sql, rs) Then
        Buscar = True
        m_Im = "" & rs!tso_tci_im
        m_CPF = "" & rs!tso_cpf
        m_Nome = "" & rs!tso_nome
        m_Cargo = "" & rs!tso_cargo
        m_TipoLogr = "" & rs!tso_tip_logr
        m_Logr = "" & rs!tso_logr
        m_Numero = "" & rs!tso_num
        m_Complemento = "" & rs!tso_complemento
        m_Bairro = "" & rs!tso_bairro
        m_Telefone = "" & rs!tso_telefone
        m_Cidade = "" & rs!tso_cidade
        m_Uf = "" & rs!tso_uf
    End If
End Function

Public Function Salvar() As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim condicao As String
    Campos = "tso_tci_im, tso_cpf, tso_nome, tso_cargo, tso_tip_logr, tso_logr, tso_num, tso_complemento, tso_bairro, tso_telefone, tso_cidade, tso_uf"
    Valores = Bdados.PreparaValor(m_Im, m_CPF, m_Nome, m_Cargo, m_TipoLogr, m_Logr, m_Numero, m_Complemento, m_Bairro, m_Telefone, Cidade, m_Uf)
    condicao = "tso_tci_im = '" & m_Im & "' and tso_cpf = '" & m_CPF & "'"
    If Bdados.GravaDados("tab_socio_empresa", Valores, Campos, condicao) Then Salvar = True
End Function

Public Function Excluir(Im As String, Optional Cpf As String) As Boolean
    Dim condicao As String
    condicao = "tso_tci_im = '" & m_Im & "'"
    If Trim(Cpf) <> "" Then
         condicao = condicao & " and tso_cpf = '" & m_CPF & "'"
    End If
    If Bdados.DeletaDados("tab_socio_empresa", condicao) Then Excluir = True
End Function

Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "select tso_cpf as CPF," & _
                " tso_nome as Nome," & _
                " tso_cargo as Cargo," & _
                " tso_tip_logr as [Logr.]," & _
                " tso_logr as Endereço," & _
                " tso_num as Número," & _
                " tso_complemento as [Compl.]," & _
                " tso_bairro as Bairro," & _
                " tso_telefone as Telefone," & _
                " tso_cidade as Cidade," & _
                " tso_uf As UF" & _
                " From Tab_Socio_Empresa" & _
                " where tso_tci_im = '" & Im & "'"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "select tso_cpf as CPF," & _
                " tso_nome as Nome," & _
                " tso_cargo as Cargo," & _
                " tso_tip_logr as Tipo_Logr," & _
                " tso_logr as Endereço," & _
                " tso_num as Número," & _
                " tso_complemento as Complemento," & _
                " tso_bairro as Bairro," & _
                " tso_telefone as Telefone," & _
                " tso_cidade as Cidade," & _
                " tso_uf As UF" & _
                " From Tab_Socio_Empresa" & _
                " where tso_tci_im = '" & Im & "'"
    End If
    If Bdados.AbreTabela(Sql) Then
        'Grid.Preencher Bdados, ""
        If Grid.Preencher(Bdados, Sql) Then PreencherGrd = True
    End If
End Function

Public Function GravarHistorico(CodMudanca As String, Im As String) As Boolean
    Dim Sql As String
    Sql = "Select tab_socio_empresa.* ," & CodMudanca & " from tab_socio_empresa where tso_tci_im = '" & Im & "'"
    If Bdados.Executa("Insert into tab_socio_empresa_historico  " & Sql) Then GravarHistorico = True
End Function
