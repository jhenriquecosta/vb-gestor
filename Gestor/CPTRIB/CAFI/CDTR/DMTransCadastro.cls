VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMTransCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private m_Codigo As String
Private m_POSTO As String
Private m_CNOME As String
Private m_CDATANASCIMENTO As String
Private m_CCPF As String
Private m_CIDENTIDADE As String
Private m_CSSP As String
Private m_CENDERECO As String
Private m_CNUMERO As String
Private m_CBAIRRO As String
Private m_CCIDADE As String
Private m_CESTADO As String
Private m_CEMAIL As String
Private m_CTELEFONE As String
Private m_CNATURALIDADE As String
Private m_CESTADOCIVIL As String
Private m_CPAI As String
Private m_CMAE As String

Private m_PPROFISSAO As String
Private m_PGRAU As String
Private m_PATIVIDADE As String
Private m_PLOCAL As String
Private m_PCIDADE As String
Private m_PENDERECO As String
Private m_PNUMERO As String
Private m_PBAIRRO As String
Private m_PPUBLICOINICIO As String
Private m_PPUBLICONOME As String
Private m_PEMPRESAINICIO As String
Private m_PEMPRESANOME As String

Private m_CVCNH As String
Private m_CVCATEGORIA As String
Private m_CVDATAHABILITACAO As String
Private m_CVREGISTRO As String
Private m_CVCURSO As String
Private m_CVORGAOEMISSORCURSO As String
Private m_CVORGAOEMISSOR As String
Private m_CVVEICULO As String
Private m_CVMARCA As String
Private m_CVMODELO As String
Private m_CVANOFAB As String
Private m_CVPLACA As String
Private m_CVCHASSI As String
Private m_CVPOTENCIA As String
Private m_CVREGVEICULO As String
Private m_CVCIDADEREGISTRO As String
Private m_CVLICENCIAMENTO As String
Private m_CVUFLICENCIAMENTO As String
Private m_CVPROPRIETARIO As String
Private m_CVENDERECO As String
Private m_CVNUMERO As String
Private m_CVBAIRRO As String
Private m_CVCIDADEPROPRIETARIO As String
Private m_CVPROPRIEDADE As String
Public Property Get Codigo() As String
        Codigo = m_Codigo
End Property
Public Property Get POSTO() As String
        POSTO = m_POSTO
End Property
Public Property Get CNOME() As String
        CNOME = m_CNOME
End Property
Public Property Get CVPROPRIEDADE() As String
        CVPROPRIEDADE = m_CVPROPRIEDADE
End Property

Public Property Get CDATANASCIMENTO() As String
    CDATANASCIMENTO = m_CDATANASCIMENTO
End Property
Public Property Get CCPF() As String
        CCPF = m_CCPF
End Property
Public Property Get CIDENTIDADE() As String
        CIDENTIDADE = m_CIDENTIDADE
End Property
Public Property Get CSSP() As String
        CSSP = m_CSSP
End Property
Public Property Get CENDERECO() As String
        CENDERECO = m_CENDERECO
End Property
Public Property Get CNUMERO() As String
        CNUMERO = m_CNUMERO
End Property
Public Property Get CBAIRRO() As String
        CBAIRRO = m_CBAIRRO
End Property
Public Property Get CCIDADE() As String
        CCIDADE = m_CCIDADE
End Property

Public Property Get CESTADO() As String
        CESTADO = m_CESTADO
End Property
Public Property Get CEMAIL() As String
        CEMAIL = m_CEMAIL
End Property
Public Property Get CTELEFONE() As String
        CTELEFONE = m_CTELEFONE
End Property
Public Property Get CNATURALIDADE() As String
        CNATURALIDADE = m_CNATURALIDADE
End Property
Public Property Get CESTADOCIVIL() As String
        CESTADOCIVIL = m_CESTADOCIVIL
End Property
Public Property Get CPAI() As String
        CPAI = m_CPAI
End Property
Public Property Get CMAE() As String
        CMAE = m_CMAE
End Property
Public Property Get PPROFISSAO() As String
        PPROFISSAO = m_PPROFISSAO
End Property
Public Property Get PGRAU() As String
        PGRAU = m_PGRAU
End Property
Public Property Get PATIVIDADE() As String
        PATIVIDADE = m_PATIVIDADE
End Property
Public Property Get PLOCAL() As String
        PLOCAL = m_PLOCAL
End Property
Public Property Get PENDERECO() As String
        PENDERECO = m_PENDERECO
End Property
Public Property Get PNUMERO() As String
        PNUMERO = m_PNUMERO
End Property
Public Property Get PBAIRRO() As String
        PBAIRRO = m_PBAIRRO
End Property
Public Property Get PPUBLICOINICIO() As String
        PPUBLICOINICIO = m_PPUBLICOINICIO
End Property
Public Property Get PPUBLICONOME() As String
        PPUBLICONOME = m_PPUBLICONOME
End Property
Public Property Get PEMPRESAINICIO() As String
        PEMPRESAINICIO = m_PEMPRESAINICIO
End Property
Public Property Get PEMPRESANOME() As String
        PEMPRESANOME = m_PEMPRESANOME
End Property
Public Property Get PCIDADE() As String
        PCIDADE = m_PCIDADE
End Property

Public Property Get CVCNH() As String
        CVCNH = m_CVCNH
End Property
Public Property Get CVCATEGORIA() As String
        CVCATEGORIA = m_CVCATEGORIA
End Property
Public Property Get CVDATAHABILITACAO() As String
        CVDATAHABILITACAO = m_CVDATAHABILITACAO
End Property
Public Property Get CVREGISTRO() As String
        CVREGISTRO = m_CVREGISTRO
End Property
Public Property Get CVCURSO() As String
        CVCURSO = m_CVCURSO
End Property
Public Property Get CVORGAOEMISSORCURSO() As String
        CVORGAOEMISSORCURSO = m_CVORGAOEMISSORCURSO
End Property
Public Property Get CVORGAOEMISSOR() As String
        CVORGAOEMISSOR = m_CVORGAOEMISSOR
End Property
Public Property Get CVVEICULO() As String
        CVVEICULO = m_CVVEICULO
End Property
Public Property Get CVMARCA() As String
        CVMARCA = m_CVMARCA
End Property
Public Property Get CVMODELO() As String
        CVMODELO = m_CVMODELO
End Property
Public Property Get CVANOFAB() As String
        CVANOFAB = m_CVANOFAB
End Property
Public Property Get CVPLACA() As String
        CVPLACA = m_CVPLACA
End Property
Public Property Get CVCHASSI() As String
        CVCHASSI = m_CVCHASSI
End Property
Public Property Get CVPOTENCIA() As String
        CVPOTENCIA = m_CVPOTENCIA
End Property
Public Property Get CVREGVEICULO() As String
        CVREGVEICULO = m_CVREGVEICULO
End Property
Public Property Get CVCIDADEREGISTRO() As String
        CVCIDADEREGISTRO = m_CVCIDADEREGISTRO
End Property
Public Property Get CVLICENCIAMENTO() As String
        CVLICENCIAMENTO = m_CVLICENCIAMENTO
End Property
Public Property Get CVUFLICENCIAMENTO() As String
        CVUFLICENCIAMENTO = m_CVUFLICENCIAMENTO
End Property
Public Property Get CVPROPRIETARIO() As String
        CVPROPRIETARIO = m_CVPROPRIETARIO
End Property
Public Property Get CVENDERECO() As String
        CVENDERECO = m_CVENDERECO
End Property
Public Property Get CVNUMERO() As String
        CVNUMERO = m_CVNUMERO
End Property
Public Property Get CVBAIRRO() As String
        CVBAIRRO = m_CVBAIRRO
End Property
Public Property Get CVCIDADEPROPRIETARIO() As String
        CVCIDADEPROPRIETARIO = m_CVCIDADEPROPRIETARIO
End Property

Public Property Let Codigo(ByVal Value As String)
        m_Codigo = Value
End Property
Public Property Let CNOME(ByVal Value As String)
        m_CNOME = Value
End Property
Public Property Let CDATANASCIMENTO(ByVal Value As String)
        m_CDATANASCIMENTO = Value
End Property
Public Property Let CCPF(ByVal Value As String)
        m_CCPF = Value
End Property
Public Property Let CIDENTIDADE(ByVal Value As String)
        m_CIDENTIDADE = Value
End Property
Public Property Let CSSP(ByVal Value As String)
        m_CSSP = Value
End Property
Public Property Let CENDERECO(ByVal Value As String)
        m_CENDERECO = Value
End Property
Public Property Let CNUMERO(ByVal Value As String)
        m_CNUMERO = Value
End Property
Public Property Let CBAIRRO(ByVal Value As String)
        m_CBAIRRO = Value
End Property
Public Property Let CCIDADE(ByVal Value As String)
        m_CCIDADE = Value
End Property
Public Property Let CESTADO(ByVal Value As String)
        m_CESTADO = Value
End Property
Public Property Let CEMAIL(ByVal Value As String)
        m_CEMAIL = Value
End Property
Public Property Let CTELEFONE(ByVal Value As String)
        m_CTELEFONE = Value
End Property
Public Property Let CNATURALIDADE(ByVal Value As String)
        m_CNATURALIDADE = Value
End Property
Public Property Let CESTADOCIVIL(ByVal Value As String)
        m_CESTADOCIVIL = Value
End Property
Public Property Let CPAI(ByVal Value As String)
        m_CPAI = Value
End Property
Public Property Let CMAE(ByVal Value As String)
        m_CMAE = Value
End Property
Public Property Let PPROFISSAO(ByVal Value As String)
        m_PPROFISSAO = Value
End Property
Public Property Let PGRAU(ByVal Value As String)
        m_PGRAU = Value
End Property
Public Property Let PATIVIDADE(ByVal Value As String)
        m_PATIVIDADE = Value
End Property
Public Property Let PLOCAL(ByVal Value As String)
        m_PLOCAL = Value
End Property
Public Property Let PENDERECO(ByVal Value As String)
        m_PENDERECO = Value
End Property
Public Property Let PNUMERO(ByVal Value As String)
        m_PNUMERO = Value
End Property
Public Property Let PCIDADE(ByVal Value As String)
        m_PCIDADE = Value
End Property

Public Property Let PBAIRRO(ByVal Value As String)
        m_PBAIRRO = Value
End Property
Public Property Let PPUBLICOINICIO(ByVal Value As String)
        m_PPUBLICOINICIO = Value
End Property
Public Property Let PPUBLICONOME(ByVal Value As String)
        m_PPUBLICONOME = Value
End Property
Public Property Let PEMPRESAINICIO(ByVal Value As String)
        m_PEMPRESAINICIO = Value
End Property
Public Property Let PEMPRESANOME(ByVal Value As String)
        m_PEMPRESANOME = Value
End Property
Public Property Let CVCNH(ByVal Value As String)
        m_CVCNH = Value
End Property
Public Property Let CVCATEGORIA(ByVal Value As String)
        m_CVCATEGORIA = Value
End Property
Public Property Let CVDATAHABILITACAO(ByVal Value As String)
        m_CVDATAHABILITACAO = Value
End Property
Public Property Let CVREGISTRO(ByVal Value As String)
        m_CVREGISTRO = Value
End Property
Public Property Let CVCURSO(ByVal Value As String)
        m_CVCURSO = Value
End Property
Public Property Let CVORGAOEMISSORCURSO(ByVal Value As String)
        m_CVORGAOEMISSORCURSO = Value
End Property
Public Property Let CVORGAOEMISSOR(ByVal Value As String)
        m_CVORGAOEMISSOR = Value
End Property
Public Property Let CVVEICULO(ByVal Value As String)
        m_CVVEICULO = Value
End Property
Public Property Let CVMARCA(ByVal Value As String)
        m_CVMARCA = Value
End Property
Public Property Let CVMODELO(ByVal Value As String)
        m_CVMODELO = Value
End Property
Public Property Let CVANOFAB(ByVal Value As String)
        m_CVANOFAB = Value
End Property
Public Property Let CVPLACA(ByVal Value As String)
        m_CVPLACA = Value
End Property
Public Property Let CVCHASSI(ByVal Value As String)
        m_CVCHASSI = Value
End Property
Public Property Let CVPOTENCIA(ByVal Value As String)
        m_CVPOTENCIA = Value
End Property
Public Property Let CVREGVEICULO(ByVal Value As String)
        m_CVREGVEICULO = Value
End Property
Public Property Let CVCIDADEREGISTRO(ByVal Value As String)
        m_CVCIDADEREGISTRO = Value
End Property
Public Property Let CVLICENCIAMENTO(ByVal Value As String)
        m_CVLICENCIAMENTO = Value
End Property
Public Property Let CVUFLICENCIAMENTO(ByVal Value As String)
        m_CVUFLICENCIAMENTO = Value
End Property
Public Property Let CVPROPRIETARIO(ByVal Value As String)
        m_CVPROPRIETARIO = Value
End Property
Public Property Let CVENDERECO(ByVal Value As String)
        m_CVENDERECO = Value
End Property
Public Property Let CVNUMERO(ByVal Value As String)
        m_CVNUMERO = Value
End Property
Public Property Let CVBAIRRO(ByVal Value As String)
        m_CVBAIRRO = Value
End Property
Public Property Let CVCIDADEPROPRIETARIO(ByVal Value As String)
        m_CVCIDADEPROPRIETARIO = Value
End Property
Public Property Let POSTO(ByVal Value As String)
        m_POSTO = Value
End Property
Public Property Let CVPROPRIEDADE(ByVal Value As String)
        m_CVPROPRIEDADE = Value
End Property

Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    campos = "CODIGO ,CNOME ,CDATANASCIMENTO ,CCPF ,CIDENTIDADE ,CSSP ,CENDERECO ,CNUMERO ,CBAIRRO ,CCIDADE ,CESTADO ,CEMAIL ,CTELEFONE ," _
    & " CNATURALIDADE ,CESTADOCIVIL ,CPAI ,CMAE ,PPROFISSAO ,PGRAUESCOLAR ,PATIVIDADE ,PLOCAL ,PENDERECO ,PNUMERO ,PBAIRRO ,PPUBLICOINICIO ," _
    & " PPUBLICONOME ,PEMPRESAINICIO ,PEMPRESANOME ,CVCNH ,CVCATEGORIA ,CVDATAHABILITACAO ,CVREGISTRO ,CVCURSO ,CVORGAOEMISSORCURSO ," _
    & " CVORGAOEMISSOR ,CVVEICULO ,CVMARCA ,CVMODELO ,CVANOFAB ,CVPLACA ,CVCHASSI ,CVPOTENCIA ,CVREGVEICULO ,CVCIDADEREGISTRO ," _
    & " CVLICENCIAMENTO ,CVUFLICENCIAMENTO ,CVPROPRIETARIO ,CVENDERECO ,CVNUMERO ,CVBAIRRO ,CVCIDADEPROPRIETARIO,POSTO,CVPROPRIEDADE,PCIDADE "
    
    
    valores = Bdados.PreparaValor(m_Codigo, m_CNOME, m_CDATANASCIMENTO, m_CCPF, m_CIDENTIDADE, m_CSSP, m_CENDERECO, m_CNUMERO, m_CBAIRRO, m_CCIDADE, m_CESTADO, m_CEMAIL, _
     m_CTELEFONE, m_CNATURALIDADE, m_CESTADOCIVIL, m_CPAI, m_CMAE, m_PPROFISSAO, m_PGRAU, m_PATIVIDADE, m_PLOCAL, m_PENDERECO, m_PNUMERO, _
     m_PBAIRRO, m_PPUBLICOINICIO, m_PPUBLICONOME, m_PEMPRESAINICIO, m_PEMPRESANOME, m_CVCNH, m_CVCATEGORIA, m_CVDATAHABILITACAO, m_CVREGISTRO, _
     m_CVCURSO, m_CVORGAOEMISSORCURSO, m_CVORGAOEMISSOR, m_CVVEICULO, m_CVMARCA, m_CVMODELO, m_CVANOFAB, m_CVPLACA, m_CVCHASSI, m_CVPOTENCIA, _
     m_CVREGVEICULO, m_CVCIDADEREGISTRO, m_CVLICENCIAMENTO, m_CVUFLICENCIAMENTO, m_CVPROPRIETARIO, m_CVENDERECO, m_CVNUMERO, m_CVBAIRRO, m_CVCIDADEPROPRIETARIO, m_POSTO, m_CVPROPRIEDADE, m_PCIDADE)
    
    condicao = "CODIGO = '" & m_Codigo & "'"
    If Bdados.GravaDados("TAB_BCP_DMTRANS_CADASTRO", valores, campos, condicao) Then
        Salvar = True
    Else
        Salvar = False
    End If

End Function
Public Function PreencherGrid(Grid As Object, Optional Nome As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "Select codigo as Codigo,ccpf as Cpf, cidentidade as Identidade,cvplaca as Placa,cnome as Nome from tab_bcp_dmtrans_cadastro "
    If Trim(Nome) <> "" Then
        Sql = Sql & "where cnome like '%" & Nome & "%'"
    End If
    Sql = Sql & " order by cnome"
    If Grid.Preencher(Bdados, Sql) Then PreencherGrid = True
End Function
Public Function Buscar(Codigo As String) As Boolean
    
End Function
