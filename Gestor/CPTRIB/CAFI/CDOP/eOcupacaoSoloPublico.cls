VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eOcupacaoSoloPublico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Icod As String
Private m_ObjetoOcupacao As String
Private m_TipoOcupacao As String
Private m_DataOcupacao As String
Private m_Periodo As String
Private m_AreaOcupada As String
Private m_AtividadeExercida As String
Private m_Logradouro As String
Private m_NomeLogradouro As String
Private m_Numero As String
Private m_Complemento As String
Private m_Bairro As String
Private m_Cep As String
Private m_Cidade As String
Private m_Uf As String
Private m_Im As String
Public cadastro As New eCadastroFiscal
Dim Codigo As New ContaCorrente

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Property Let Icod(ByVal Value As String)
    m_Icod = Value
End Property

Public Property Get Icod() As String
    Icod = m_Icod
End Property

Public Property Let ObjetoOcupacao(ByVal Value As String)
    m_ObjetoOcupacao = Value
End Property

Public Property Get ObjetoOcupacao() As String
    ObjetoOcupacao = m_ObjetoOcupacao
End Property

Public Property Let TipoOcupacao(ByVal Value As String)
    m_TipoOcupacao = Value
End Property

Public Property Get TipoOcupacao() As String
    TipoOcupacao = m_TipoOcupacao
End Property

Public Property Let DataOcupacao(ByVal Value As String)
    m_DataOcupacao = Value
End Property

Public Property Get DataOcupacao() As String
    DataOcupacao = m_DataOcupacao
End Property

Public Property Let Periodo(ByVal Value As String)
    m_Periodo = Value
End Property

Public Property Get Periodo() As String
    Periodo = m_Periodo
End Property

Public Property Let AreaOcupada(ByVal Value As String)
    m_AreaOcupada = Value
End Property

Public Property Get AreaOcupada() As String
    AreaOcupada = m_AreaOcupada
End Property

Public Property Let AtividadeExercida(ByVal Value As String)
    m_AtividadeExercida = Value
End Property

Public Property Get AtividadeExercida() As String
    AtividadeExercida = m_AtividadeExercida
End Property


Public Property Get Logradouro() As String
    Logradouro = m_Logradouro
End Property

Public Property Let Logradouro(ByVal Value As String)
    m_Logradouro = Value
End Property

Public Property Get NomeLogradouro() As String
    NomeLogradouro = m_NomeLogradouro
End Property

Public Property Let NomeLogradouro(ByVal Value As String)
    m_NomeLogradouro = Value
End Property

Public Property Get Numero() As String
    Numero = m_Numero
End Property

Public Property Let Numero(ByVal Value As String)
    m_Numero = Value
End Property

Public Property Get Complemento() As String
    Complemento = m_Complemento
End Property

Public Property Let Complemento(ByVal Value As String)
    m_Complemento = Value
End Property

Public Property Get Bairro() As String
    Bairro = m_Bairro
End Property

Public Property Let Bairro(ByVal Value As String)
    m_Bairro = Value
End Property

Public Property Get Cep() As String
    Cep = m_Cep
End Property

Public Property Let Cep(ByVal Value As String)
    m_Cep = Value
End Property

Public Property Get Cidade() As String
    Cidade = m_Cidade
End Property

Public Property Let Cidade(ByVal Value As String)
    m_Cidade = Value
End Property

Public Property Get UF() As String
    UF = m_Uf
End Property

Public Property Let UF(ByVal Value As String)
    m_Uf = Value
End Property
    
Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    Dim codCadastro As String
    Set Codigo = New ContaCorrente
       
    If (m_Icod = "") Then
        codCadastro = Codigo.GeraCodPagamento(48)
        With cadastro
            .icad = codCadastro
            .Im = m_Im
            .Tipo_Cadastro = ecOcupacaoPermanencia
            .Status = ecAtivo
            .Data_Cadastro = Date
        End With
        If Not cadastro.Salvar Then Exit Function
    Else
        codCadastro = m_Icod
    End If
     
        campos = "TOS_COD_CADASTRO,TOS_OBJETO_OCUPACAO,TOS_TIPO_OCUPACAO,TOS_PERIODO,TOS_AREA_OCUPADA,TOS_ATIVIDADE_EXERCIDA,TOS_LOGRADOURO,TOS_NOME_LOGRADOURO,TOS_NUMERO,TOS_COMPLEMENTO,TOS_BAIRRO,TOS_CEP,TOS_CIDADE,TOS_UF,TOS_DATA_OCUPACAO"
        valores = Bdados.PreparaValor(codCadastro, m_ObjetoOcupacao, m_TipoOcupacao, m_Periodo, m_AreaOcupada, m_AtividadeExercida, m_Logradouro, m_NomeLogradouro, m_Numero, m_Complemento, m_Bairro, m_Cep, m_Cidade, m_Uf, m_DataOcupacao)
        condicao = "TOS_COD_CADASTRO = " & codCadastro
        Bdados.AbreTrans
        If Bdados.GravaDados("TAB_CAD_OCUPACAO_SOLO_PUBLICO", valores, campos, condicao) Then
            Salvar = True
            Bdados.GravaTrans
        Else
            Bdados.CancelaTrans
        End If
    
End Function

Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    
    Sql = " SELECT TOS_COD_CADASTRO AS Código, "
    Sql = Sql & " TOS_OBJETO_OCUPACAO AS Objeto_Ocupação, "
    Sql = Sql & " TOS_DATA_OCUPACAO AS Data_Ocupação, "
    Sql = Sql & " Tipo AS Tipo, "
    Sql = Sql & " TOS_PERIODO AS Período, "
    Sql = Sql & " TOS_AREA_OCUPADA AS Área_Ocupada, "
    Sql = Sql & " TAE_NOME AS Atividade, "
    Sql = Sql & " TOS_LOGRADOURO AS Logradouro,"
    Sql = Sql & " TOS_NOME_LOGRADOURO AS Nome_Logradouro, "
    Sql = Sql & " TOS_NUMERO AS Número, "
    Sql = Sql & " TOS_COMPLEMENTO AS Complemento, "
    Sql = Sql & " TOS_BAIRRO AS Bairro,"
    Sql = Sql & " TOS_CEP As Cep, "
    Sql = Sql & " TOS_CIDADE As Cidade, "
    Sql = Sql & " UF As UF, "
    Sql = Sql & " TOS_TIPO_OCUPACAO, "
    Sql = Sql & " TOS_ATIVIDADE_EXERCIDA, "
    Sql = Sql & " TOS_UF,"
    Sql = Sql & " TCF_TCI_IM"
    Sql = Sql & " From VIS_CAD_OCUPACAO"
    Sql = Sql & " WHERE TCF_STATUS= " & ecAtivo & " "
    If Im <> "" Then Sql = Sql & " AND   TCF_TCI_IM = '" & Im & "'"
    
    If Not Grid.Preencher(Bdados, Sql, 1100, 2000, 1700, 1500, 1000, 1500, 4500, 1700, 3000, 1000, 3000, 1500, 1000, 2000, 1500, 0, 0, 0, 0) Then
     Util.Avisa "Contribuinte não possui Cadastro de Ocupação e de Permanência em Solo de Logradouros Públicos."
    End If
End Function

Public Function PreencherGrdConsulta(Grid As Object, Im As String, cod As String, TipoOcup As String, StatusCad As String) As Boolean
    Dim Sql As String
    Dim condicao As String
    
    If (cod <> "") Then condicao = condicao & " AND TOS_COD_CADASTRO = " & cod
    If (TipoOcup <> "") Then condicao = condicao & " AND TIPO =  '" & TipoOcup & "'"
    If (StatusCad <> "") Then condicao = condicao & " AND TCF_STATUS = " & StatusCad
    
    
    
    Sql = " SELECT TOS_COD_CADASTRO AS Código, "
    Sql = Sql & " tci_nome AS Contribuinte, "
    Sql = Sql & " TCF_DATA_CADASTRO AS Data_Cadastro, "
    Sql = Sql & " STATUS AS Status, "
    Sql = Sql & " TCF_DATA_CANCELAMENTO AS Data_Baixa, "
    Sql = Sql & " TOS_OBJETO_OCUPACAO AS Objeto_Ocupação, "
    Sql = Sql & " TOS_DATA_OCUPACAO AS Data_Ocupação, "
    Sql = Sql & " Tipo AS Tipo,"
    Sql = Sql & " TOS_PERIODO AS Período, "
    Sql = Sql & " TOS_AREA_OCUPADA AS Área_Ocupada, "
    Sql = Sql & " TAE_NOME AS Atividade, "
    Sql = Sql & " TOS_LOGRADOURO AS Logradouro,"
    Sql = Sql & " TOS_NOME_LOGRADOURO AS Nome_Logradouro, "
    Sql = Sql & " TOS_NUMERO AS Número, "
    Sql = Sql & " TOS_COMPLEMENTO AS Complemento, "
    Sql = Sql & " TOS_BAIRRO AS Bairro,"
    Sql = Sql & " TOS_CEP As Cep, "
    Sql = Sql & " TOS_CIDADE As Cidade, UF As UF"
    Sql = Sql & " From VIS_CAD_OCUPACAO"
    Sql = Sql & " WHERE 1=1"
    If Im <> "" Then Sql = Sql & " AND TCF_TCI_IM = '" & Im & "'"
    If condicao <> "" Then Sql = Sql & condicao
   
    If Not Grid.Preencher(Bdados, Sql, 1100, 5000, 1100, 1200, 1150, 2000, 2000, 2000, 800, 2000, 2000, 2000, 1800, 1800, 3000, 1500, 1000, 2000, 2000) Then
        Util.Avisa "Contribuinte não possui Cadastro de Ocupação e de Permanência em Solo de Logradouros Públicos."
    End If
    
End Function

