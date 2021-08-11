VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eSanitario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Icod As String
Private m_Logradouro As String
Private m_NomeLogradouro As String
Private m_Numero As String
Private m_Complemento As String
Private m_Bairro As String
Private m_Cep As String
Private m_Cidade As String
Private m_Uf As String
Private m_Motivo As String
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

Public Property Let Motivo(ByVal Value As String)
    m_Motivo = Value
End Property

Public Property Get Motivo() As String
    Motivo = m_Motivo
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
        codCadastro = Codigo.GeraCodPagamento(49)
        With cadastro
            .icad = codCadastro
            .Im = m_Im
            .Tipo_Cadastro = ecSanitario
            .Status = ecAtivo
            .Data_Cadastro = Date
        End With
        If Not cadastro.Salvar Then Exit Function
    Else
        codCadastro = m_Icod
    End If
     
        campos = "TSA_COD_CADASTRO,TSA_LOGRADOURO,TSA_NOME_LOGRADOURO,TSA_NUMERO,TSA_COMPLEMENTO,TSA_BAIRRO,TSA_CEP,  TSA_CIDADE, TSA_UF,TSA_MOTIVO"
        valores = Bdados.PreparaValor(codCadastro, m_Logradouro, m_NomeLogradouro, m_Numero, m_Complemento, m_Bairro, m_Cep, m_Cidade, m_Uf, m_Motivo)
        condicao = " TSA_COD_CADASTRO = " & codCadastro
        Bdados.AbreTrans
        If Bdados.GravaDados("TAB_CAD_SANITARIO", valores, campos, condicao) Then
            Salvar = True
            Bdados.GravaTrans
        Else
            Bdados.CancelaTrans
        End If
    
End Function

Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    
    Sql = " SELECT TSA_COD_CADASTRO AS Código, "
    Sql = Sql & " TSA_MOTIVO AS Motivo, "
    Sql = Sql & " TSA_LOGRADOURO AS Logradouro, "
    Sql = Sql & " TSA_NOME_LOGRADOURO AS Nome_Logradouro, "
    Sql = Sql & " TSA_NUMERO AS Numero, "
    Sql = Sql & " TSA_COMPLEMENTO AS Complemento, "
    Sql = Sql & " TSA_BAIRRO AS Bairro, "
    Sql = Sql & " TSA_CEP AS CEP, "
    Sql = Sql & " TSA_CIDADE AS Cidade, "
    Sql = Sql & " UF, TSA_UF,"
    Sql = Sql & " TCF_TCI_IM"
    Sql = Sql & " FROM VIS_CAD_SANITARIO"
    Sql = Sql & " WHERE TCF_STATUS= " & ecAtivo & " "
    If Im <> "" Then Sql = Sql & " AND   TCF_TCI_IM = '" & Im & "'"
    If Not Grid.Preencher(Bdados, Sql, 1100, 3000, 1700, 3000, 1000, 3000, 1500, 1000, 2000, 1500, 0, 0) Then
     Util.Avisa "Contribuinte não possui Cadastro Sanitário"
    End If
End Function

Public Function PreencherGrdConsulta(Grid As Object, Im As String, cod As String, StatusCad As String) As Boolean
    Dim Sql As String
    Dim condicao As String
    
    If (cod <> "") Then condicao = condicao & " AND TSA_COD_CADASTRO = " & cod
    If (StatusCad <> "") Then condicao = condicao & " AND TCF_STATUS = " & StatusCad
               
    Sql = "  SELECT TSA_COD_CADASTRO AS Código, "
    Sql = Sql & " tci_nome AS Contribuinte, "
    Sql = Sql & " TCF_DATA_CADASTRO AS Data_Cadastro, "
    Sql = Sql & " STATUS AS Status, "
    Sql = Sql & " TCF_DATA_CANCELAMENTO AS Data_Baixa, "
    Sql = Sql & " TSA_MOTIVO AS Motivo, "
    Sql = Sql & " TSA_LOGRADOURO AS Logradouro,"
    Sql = Sql & " TSA_NOME_LOGRADOURO AS Nome_Logradouro, "
    Sql = Sql & " TSA_NUMERO AS Numero, "
    Sql = Sql & " TSA_COMPLEMENTO AS Complemento, "
    Sql = Sql & " TSA_BAIRRO AS Bairro,"
    Sql = Sql & " TSA_CEP As Cep, "
    Sql = Sql & " TSA_CIDADE As Cidade, "
    Sql = Sql & " UF, TSA_UF"
    Sql = Sql & " From VIS_CAD_SANITARIO"
    Sql = Sql & " WHERE 1=1"
    If Im <> "" Then Sql = Sql & " AND TCF_TCI_IM = '" & Im & "'"

    If condicao <> "" Then Sql = Sql & condicao
   
    If Not Grid.Preencher(Bdados, Sql, 1100, 5000, 1100, 1200, 1150, 3000, 1700, 3000, 1000, 3000, 1500, 1000, 2000, 1500) Then
        Util.Avisa "Contribuinte não possui Cadastro Sanitário"
    End If
    
End Function


