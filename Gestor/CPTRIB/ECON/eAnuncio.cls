VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eAnuncio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public m_ICAD As String
Public m_Im As String
Public m_Movimento As Variant
Public m_Ilumina As Integer
Public m_Zona As Integer
Public m_Dimensao As String
Public m_Area As String
    Public m_DataInstalacao As String
Public m_ICLocal As String
Public m_Endereco As String
Public m_UFM As Double
Public m_Valor_Apagar As Double
Public m_Item As Integer
Public m_Doc_Origem As String
Public m_SubPublicidade As String
Public m_Obs As String
Public m_Multiplicador As String
Public m_Logradouro As String
Public m_Bairro As String
Public m_Tipo_Logra As String

Public Property Let Logradouro(ByVal Value As String)
    m_Logradouro = Value
End Property

Public Property Get Logradouro() As String
    Logradouro = m_Logradouro
End Property


Public Property Let Bairro(ByVal Value As String)
    m_Bairro = Value
End Property

Public Property Get Bairro() As String
    Bairro = m_Bairro
End Property


Public Property Let TipoLogra(ByVal Value As String)
    m_Tipo_Logra = Value
End Property

Public Property Get TipoLogra() As String
    TipoLogra = m_Tipo_Logra
End Property


Public Property Let Multiplicador(ByVal Value As String)
    m_Multiplicador = Value
End Property

Public Property Get Multiplicador() As String
    Multiplicador = m_Multiplicador
End Property


Public Property Let Obs(ByVal Value As String)
    m_Obs = Value
End Property

Public Property Get Obs() As String
    Obs = m_Obs
End Property



Public Property Let SubPublicidade(ByVal Value As String)
    m_SubPublicidade = Value
End Property

Public Property Get SubPublicidade() As String
    SubPublicidade = m_SubPublicidade
End Property


Public Property Let Doc_Origem(ByVal Value As String)
    m_Doc_Origem = Value
End Property

Public Property Get Doc_Origem() As String
    Doc_Origem = m_Doc_Origem
End Property


Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = " SELECT TAN_ICAD AS Item, RTRIM(LTRIM(TAN_MOVIMENTO) )+' - '+ tip_nome_imposto AS Movimento,  TAN_DIMENSAO AS Dimenções,"
        Sql = Sql & " TAN_AREA As Area,  TAN_DATA_INSTALACAO As Data_Instalacao,  TAN_VALOR_UFM As [Valor em UFM],"
        Sql = Sql & " TAN_VALOR As [Valor a pagar], TAN_SUB_MOVIMENTO  As [Sub Publicidade],tan_DOC_ORIGEM AS [Doc Origem]"
        Sql = Sql & " From TAB_ANUNCIO, tab_imposto"
        Sql = Sql & " Where TAN_MOVIMENTO = tip_cod_imposto"
        Sql = Sql & " and TAN_TCI_IM = '" & Im & "' order by TAN_ICAD"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = " SELECT TAN_ICAD AS Item, RTRIM(LTRIM(TAN_MOVIMENTO) ) " & Bdados.Concatena & " ' - ' " & Bdados.Concatena & " tip_nome_imposto AS Movimento,  TAN_DIMENSAO AS Dimenções,"
        Sql = Sql & " TAN_AREA As Area,  TAN_DATA_INSTALACAO As Data_Instalacao,  TAN_VALOR_UFM As Valor_UFM,"
        Sql = Sql & " TAN_VALOR As Valor, TAN_SUB_MOVIMENTO  As Sub_Ublicidade,tan_DOC_ORIGEM AS Doc_Origem"
        Sql = Sql & " From TAB_ANUNCIO, tab_imposto"
        Sql = Sql & " Where TAN_MOVIMENTO = tip_cod_imposto"
        Sql = Sql & " and TAN_TCI_IM = '" & Im & "' order by TAN_ICAD"
    End If
    If Grid.Preencher(Bdados, Sql) Then PreencherGrd = True
End Function

Public Function Buscar(icad As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "select * from TAB_ANUNCIO where TAN_ICAD = '" & icad & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        m_ICAD = "" & rs.Fields("TAN_ICAD")
        m_Movimento = "" & rs.Fields("TAN_MOVIMENTO")
        m_Dimensao = "" & rs.Fields("TAN_DIMENSAO")
        m_Area = "" & rs.Fields("TAN_AREA")
        m_DataInstalacao = "" & rs.Fields("TAN_DATA_INSTALACAO")
        m_ICLocal = "" & rs.Fields("TAN_IC_LOCAL")
        m_Tipo_Logra = "" & rs.Fields("TAN_TLl_COD_TIP_LOGR")
        m_UFM = "" & rs.Fields("TAN_VALOR_UFM")
        m_Valor_Apagar = "" & rs.Fields("TAN_VALOR")
        m_Multiplicador = "" & rs.Fields("TAN_MUTIPLICADOR")
        m_Logradouro = "" & rs.Fields("TAN_tlg_cod_logradouro")
        m_Bairro = "" & rs.Fields("TAN_TBA_COD_BAIRRO")
        m_Obs = "" & rs.Fields("TAN_OBS")
        m_SubPublicidade = "" & rs.Fields("TAN_SUB_MOVIMENTO")
        Buscar = True
    End If
End Function

Public Function Salvar() As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim condicao As String
    
    Campos = "TAN_TCI_IM,"
    Campos = Campos & "TAN_ICAD,"
    Campos = Campos & "TAN_MOVIMENTO,"
    Campos = Campos & "TAN_DIMENSAO,"
    Campos = Campos & "TAN_AREA,"
    Campos = Campos & "TAN_DATA_INSTALACAO,"
    Campos = Campos & "TAN_DATA_CADASTRO,"
    Campos = Campos & "TAN_IC_LOCAL,"
    Campos = Campos & "TAN_TLl_COD_TIP_LOGR,"
    Campos = Campos & "TAN_VALOR_UFM,"
    Campos = Campos & "TAN_VALOR,"
    Campos = Campos & "TAN_MUTIPLICADOR,"
    Campos = Campos & "TAN_tlg_cod_logradouro,"
    Campos = Campos & "TAN_TBA_COD_BAIRRO,"
    Campos = Campos & "TAN_OBS,"
    Campos = Campos & "TAN_SUB_MOVIMENTO,tan_doc_origem"
    
    Valores = Bdados.PreparaValor(m_Im, m_ICAD, Trim(m_Movimento), m_Dimensao, Bdados.Converte(m_Area, TCDuplo), m_DataInstalacao, Format(Date, "dd/MM/yyyy"), m_ICLocal, m_Tipo_Logra, Bdados.Converte(m_UFM, TCDuplo), Bdados.Converte(m_Valor_Apagar, TCDuplo), Bdados.Converte(m_Multiplicador, TCDuplo), m_Logradouro, m_Bairro, m_Obs, m_SubPublicidade, m_Doc_Origem)
    condicao = "TAN_ICAD = " & m_ICAD & " AND  TAN_TCI_IM = " & m_Im
    If Bdados.GravaDados("TAB_ANUNCIO", Valores, Campos, condicao) Then Salvar = True
End Function

Public Function Excluir(icad As String) As Boolean
    Dim condicao As String
    condicao = "TAN_icad = '" & icad & "'"
    
    If Bdados.DeletaDados("TAB_ANUNCIO", condicao) Then Excluir = True
End Function

Public Sub MontarGrid(Grid As Object)
    Grid.ColumnHeaders.Clear
    Grid.ColumnHeaders.Add , , "Item"
    Grid.ColumnHeaders.Add , , "Movimento"
    Grid.ColumnHeaders.Add , , "Descrição"
    Grid.ColumnHeaders.Add , , "Área Total"
    Grid.ColumnHeaders.Add , , "Data Instalação"
    Grid.ColumnHeaders.Add , , "Valor UFM"
    Grid.ColumnHeaders.Add , , "Valor a pagar"
    Grid.ColumnHeaders.Add , , "Sub Movimento"
    Grid.ColumnHeaders.Add , , "Doc.Origem"
End Sub

Public Property Let Item(ByVal Value As Integer)
    m_Item = Value
End Property

Public Property Get Item() As Integer
    Item = m_Item
End Property

Public Property Let icad(ByVal Value As String)
    m_ICAD = Value
End Property

Public Property Get icad() As String
    icad = m_ICAD
End Property

Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let Valor_UFM(ByVal Value As Double)
    m_UFM = Value
End Property

Public Property Get Valor_UFM() As Double
    Valor_UFM = m_UFM
End Property


Public Property Let Valor_Apagar(ByVal Value As Double)
    m_Valor_Apagar = Value
End Property

Public Property Get Valor_Apagar() As Double
    Valor_Apagar = m_Valor_Apagar
End Property


Public Property Let Movimento(ByVal Value As Variant)
    m_Movimento = Value
End Property

Public Property Get Movimento() As Variant
    Movimento = m_Movimento
End Property

Public Property Let Ilumina(ByVal Value As Integer)
    m_Ilumina = Value
End Property

Public Property Get Ilumina() As Integer
    Ilumina = m_Ilumina
End Property

Public Property Let Zona(ByVal Value As Integer)
    m_Zona = Value
End Property

Public Property Get Zona() As Integer
    Zona = m_Zona
End Property

Public Property Let Dimensao(ByVal Value As String)
    m_Dimensao = Value
End Property

Public Property Get Dimensao() As String
    Dimensao = m_Dimensao
End Property

Public Property Let Area(ByVal Value As String)
    m_Area = Value
End Property

Public Property Get Area() As String
    Area = m_Area
End Property

Public Property Let DataInstalacao(ByVal Value As String)
    m_DataInstalacao = Value
End Property

Public Property Get DataInstalacao() As String
    DataInstalacao = m_DataInstalacao
End Property

Public Property Let ICLocal(ByVal Value As String)
    m_ICLocal = Value
End Property

Public Property Get ICLocal() As String
    ICLocal = m_ICLocal
End Property

Public Property Let Endereco(ByVal Value As String)
    m_Endereco = Value
End Property

Public Property Get Endereco() As String
    Endereco = m_Endereco
End Property

Public Function GeraIcad(Im As String)
    Dim Sql As String
    Dim rs As VSRecordset
    
    
    
End Function

