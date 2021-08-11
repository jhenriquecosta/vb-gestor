VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCPParametroIPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Periodo As Integer
Private m_Bairro As Integer
'CODO
Private m_Grupo As Integer
Private m_Componente As Integer
Private m_ValorPadraoConstrutivo As Double
Private m_AliquotaPadraoConstrutivo As Double
'FIM CODO
Private m_ValorTerreno As Double
Private m_ValorImovelResidencial As Double
Private m_ValorImovelComercial As Double
Private m_AliquotaImovelResidencial As Double
Private m_AliquotaImovelComercial As Double
Private m_AliquotaTerrenoMurado As Double
Private m_AliquotaTerrenoBaldio As Double
Private m_logradouro As Long

Public Property Let ValorPadraoConstrutivo(ByVal Value As Double)
    m_ValorPadraoConstrutivo = Value
End Property
Public Property Get ValorPadraoConstrutivo() As Double
    ValorPadraoConstrutivo = m_ValorPadraoConstrutivo
End Property

Public Property Let AliquotaPadraoConstrutivo(ByVal Value As Double)
    m_AliquotaPadraoConstrutivo = Value
End Property
Public Property Get AliquotaPadraoConstrutivo() As Double
    AliquotaPadraoConstrutivo = m_AliquotaPadraoConstrutivo
End Property
Public Property Get Componente() As Integer
    Componente = m_Componente
End Property

Public Property Let Componente(ByVal Value As Integer)
    m_Componente = Value
End Property
Public Property Get Grupo() As Integer
    Grupo = m_Grupo
End Property

Public Property Let Grupo(ByVal Value As Integer)
    m_Grupo = Value
End Property
Public Property Get Periodo() As Integer
    Periodo = m_Periodo
End Property

Public Property Let Periodo(ByVal Value As Integer)
    m_Periodo = Value
End Property

Public Property Get Bairro() As Integer
    Bairro = m_Bairro
End Property

Public Property Let Bairro(ByVal Value As Integer)
    m_Bairro = Value
End Property

Public Property Get ValorTerreno() As Double
    ValorTerreno = m_ValorTerreno
End Property

Public Property Let ValorTerreno(ByVal Value As Double)
    m_ValorTerreno = Value
End Property
Public Property Get ValorImovelResidencial() As Double
    ValorImovelResidencial = m_ValorImovelResidencial
End Property

Public Property Let ValorImovelResidencial(ByVal Value As Double)
    m_ValorImovelResidencial = Value
End Property
Public Property Get ValorImovelComercial() As Double
    ValorImovelComercial = m_ValorImovelComercial
End Property

Public Property Let ValorImovelComercial(ByVal Value As Double)
    m_ValorImovelComercial = Value
End Property

Public Property Get AliquotaImovelResidencial() As Double
    AliquotaImovelResidencial = m_AliquotaImovelResidencial
End Property

Public Property Let AliquotaImovelResidencial(ByVal Value As Double)
    m_AliquotaImovelResidencial = Value
End Property
Public Property Get AliquotaImovelComercial() As Double
    AliquotaImovelComercial = m_AliquotaImovelComercial
End Property

Public Property Let AliquotaImovelComercial(ByVal Value As Double)
    m_AliquotaImovelComercial = Value
End Property
Public Property Get AliquotaTerrenoMurado() As Double
    AliquotaTerrenoMurado = m_AliquotaTerrenoMurado
End Property

Public Property Let AliquotaTerrenoMurado(ByVal Value As Double)
    m_AliquotaTerrenoMurado = Value
End Property
Public Property Get AliquotaTerrenoBaldio() As Double
    AliquotaTerrenoBaldio = m_AliquotaTerrenoBaldio
End Property

Public Property Let AliquotaTerrenoBaldio(ByVal Value As Double)
    m_AliquotaTerrenoBaldio = Value
End Property
Public Function Buscar(Periodo As Integer, Bairro As Integer) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "SELECT * FROM TAB_BCP_PARAMETRO_IPTU WHERE TPI_PERIODO=" & Periodo & " AND TPI_TBA_COD_BAIRRO=" & Bairro
    If Bdados.AbreTabela(Sql, rs) Then
        m_Periodo = CInt(rs!TPI_PERIODO)
        m_Bairro = CInt(rs!TPI_TBA_COD_BAIRRO)
        m_ValorTerreno = CDbl(rs!TPI_VALOR_TERRENO)
        m_ValorImovelResidencial = CDbl(rs!TPI_VALOR_IMOVEL_RESIDENCIAL)
        m_ValorImovelComercial = CDbl(rs!TPI_VALOR_IMOVEL_COMERCIAL)
        m_AliquotaImovelResidencial = CDbl(rs!TPI_ALIQUOTA_IMOVEL_RESIDENCIAL)
        m_AliquotaImovelComercial = CDbl(rs!TPI_ALIQUOTA_IMOVEL_COMERCIAL)
        m_AliquotaTerrenoMurado = CDbl(rs!TPI_ALIQUOTA_TERRENO_MURADO)
        m_AliquotaTerrenoBaldio = CDbl(rs!TPI_ALIQUOTA_TERRENO_BALDIO)
        Buscar = True
    End If
End Function
'CODO
Public Function BuscarPadraoConstrutivo(Periodo As Integer, setor As Integer, Grupo As Integer, Componente As Integer) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "SELECT * FROM TAB_BCP_PARAMETRO_IPTU WHERE TPI_PERIODO=" & Periodo & " AND TPI_SETOR=" & setor & " AND TPI_GRUPO=" & Grupo & " AND TPI_COMPONENTE=" & Componente
    Dim bd As New Connection
    Dim rst As New Recordset
    bd.ConnectionString = Bdados.Conexao.ConnectionString
    bd.Open
    Set rst = bd.Execute(Sql)
    
    If Not rst.EOF Then
        m_Periodo = CInt(rst!TPI_PERIODO)
        'm_Bairro = CInt(rs!TPI_TBA_COD_BAIRRO)
        'm_ValorTerreno = CDbl(rs!TPI_VALOR_TERRENO)
        'm_ValorImovelResidencial = CDbl(rs!TPI_VALOR_IMOVEL_RESIDENCIAL)
        'm_ValorImovelComercial = CDbl(rs!TPI_VALOR_IMOVEL_COMERCIAL)
        'm_AliquotaImovelResidencial = CDbl(rs!TPI_ALIQUOTA_IMOVEL_RESIDENCIAL)
        'm_AliquotaImovelComercial = CDbl(rs!TPI_ALIQUOTA_IMOVEL_COMERCIAL)
       'm_AliquotaTerrenoMurado = CDbl(rs!TPI_ALIQUOTA_TERRENO_MURADO)
        'm_AliquotaTerrenoBaldio = CDbl(rs!TPI_ALIQUOTA_TERRENO_BALDIO)
         
        'CODO
        m_Grupo = CInt(rst!TPI_GRUPO)
        m_Componente = CInt(rst!TPI_COMPONENTE)
        m_AliquotaPadraoConstrutivo = CDbl(rst!TPI_ALIQUOTA_PADRAO_CONSTRUTIVO)
        m_ValorPadraoConstrutivo = CDbl(rst!TPI_VALOR_PADRAO_CONSTRUTIVO)
        'FIM CODO
        BuscarPadraoConstrutivo = True
    End If
    rst.Close
    bd.Close
    
End Function
'FIM CODO
Public Function BuscarPorSetor(Periodo As Integer, Bairro As Integer, setor As Integer, quadra As Integer, pac As Boolean, logradouro As Long) As Boolean
    m_logradouro = logradouro
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "SELECT * FROM TAB_BCP_PARAMETRO_IPTU WHERE TPI_PERIODO=" & Periodo & " AND TPI_SETOR=" & setor & " AND TPI_QUADRA=0"
    
    Dim bd As New Connection
    bd.ConnectionString = Bdados.Conexao.DBConnection.ConnectionString
    bd.Open
    Dim rstab As New Recordset
    Set rstab = bd.Execute(Sql)
    
    
    If Not rstab.EOF Then
        m_Periodo = CInt(rstab!TPI_PERIODO)
        m_Bairro = CInt(rstab!TPI_TBA_COD_BAIRRO)
        m_ValorTerreno = CDbl(rstab!TPI_VALOR_TERRENO)
        m_ValorImovelResidencial = CDbl(rstab!TPI_VALOR_IMOVEL_RESIDENCIAL)
        m_ValorImovelComercial = CDbl(rstab!TPI_VALOR_IMOVEL_COMERCIAL)
        m_AliquotaImovelResidencial = CDbl(rstab!TPI_ALIQUOTA_IMOVEL_RESIDENCIAL)
        m_AliquotaImovelComercial = CDbl(rstab!TPI_ALIQUOTA_IMOVEL_COMERCIAL)
        m_AliquotaTerrenoMurado = CDbl(rstab!TPI_ALIQUOTA_TERRENO_MURADO)
        m_AliquotaTerrenoBaldio = CDbl(rstab!TPI_ALIQUOTA_TERRENO_BALDIO)
        BuscarPorSetor = True
    End If
    rstab.Close
    bd.Close
    'METODO PARA CODO COM OBJETIVO DE ATENDER A PAGINA 194 DO CTM
    '3003   'SANTOS DUMONT
    '21003  'AUGUSTO TEIXEIRA
    '56007  'AFONSO PENA
    '170007 '28 JULHO
    '199007 'PC FERREIRA BAYMA
    '158007 'JOAO PESSOA
    '160003 'VITORINO FREIRE
    '41007  'ANTONIO LOGES
    '174007 '13 DE MAIO
    '32002  'CRISTOVAO COLOMBO
    '310007 'COLATINO BORBOREM
    Select Case setor
        Case 1
        Case 2
        Case 3
            Select Case logradouro
                Case 56007
                    If quadra = 1 Or quadra = 2 Or quadra = 5 Or quadra = 12 Or quadra = 13 Or quadra = 14 Or quadra = 15 Or quadra = 20 Or quadra = 21 Or quadra = 22 Or quadra = 30 Then
                        m_ValorTerreno = CDbl(200)
                    End If
                Case 170007
                    If quadra = 9 Or quadra = 11 Or quadra = 12 Or quadra = 15 Or quadra = 16 Or quadra = 20 Or quadra = 22 Or quadra = 23 Then
                        m_ValorTerreno = CDbl(150)
                    End If
                Case 199007
                    If quadra = 4 Or quadra = 6 Then
                        m_ValorTerreno = CDbl(200)
                    End If
                Case 158007
                    If quadra = 14 Or quadra = 21 Then
                        m_ValorTerreno = CDbl(100)
                    End If
            End Select
        Case 4
            If logradouro = 31007 Then
                If quadra = 3 Then
                    m_ValorTerreno = CDbl(150)
                End If
            End If
        Case 5
            Select Case logradouro
                Case 158007
                    If quadra = 6 Or quadra = 7 Or quadra = 8 Or quadra = 9 Or quadra = 10 Or quadra = 11 Then
                        m_ValorTerreno = CDbl(100)
                    End If
                Case 160003
                    If quadra = 22 Or quadra = 23 Or quadra = 24 Or quadra = 30 Then
                        m_ValorTerreno = CDbl(100)
                    End If
                Case 41007
                    If quadra = 5 Or quadra = 20 Or quadra = 24 Or quadra = 25 Or quadra = 26 Or quadra = 27 Or quadra = 28 Then
                        m_ValorTerreno = CDbl(100)
                    End If
                Case 174007
                    If quadra = 2 Or quadra = 3 Or quadra = 6 Or quadra = 7 Or quadra = 8 Then
                        m_ValorTerreno = CDbl(100)
                    End If
            End Select
        Case 6
            Case 21003
                    If quadra = 1 Or quadra = 2 Or quadra = 7 Or quadra = 8 Or quadra = 9 Or quadra = 10 Or quadra = 12 Or quadra = 13 Or quadra = 14 Or quadra = 15 Or quadra = 16 Then
                        m_ValorTerreno = CDbl(100)
                    End If
        Case 7
        Case 8
        Case 9
            If logradouro = 160003 Then
                If quadra = 3 Or quadra = 4 Or quadra = 5 Or quadra = 22 Or quadra = 23 Or quadra = 24 Or quadra = 30 Then
                    m_ValorTerreno = CDbl(100)
                End If
            End If
            
        Case 10
        Case 11
            If logradouro = 3003 Then
                If quadra = 2 Or quadra = 3 Or quadra = 4 Or quadra = 5 Or quadra = 6 Or quadra = 7 Then
                    m_ValorTerreno = CDbl(100)
                End If
            End If
        Case 12
            If logradouro = 3003 Then
                If quadra = 1 Or quadra = 2 Or quadra = 12 Or quadra = 13 Or quadra = 26 Or quadra = 27 Or quadra = 28 Then
                    m_ValorTerreno = CDbl(100)
                End If
            End If
        Case 13
        Case 14
            If logradouro = 32002 Then
                If quadra = 1 Or quadra = 8 Or quadra = 9 Or quadra = 13 Or quadra = 14 Or quadra = 15 Or quadra = 16 Or quadra = 19 Or quadra = 27 Or quadra = 28 Or quadra = 80 Or quadra = 86 Then
                    m_ValorTerreno = CDbl(100)
                End If
            End If
        Case 15
            If logradouro = 3003 Then
                If quadra = 1 Or quadra = 4 Or quadra = 10 Or quadra = 21 Or quadra = 68 Then
                    m_ValorTerreno = CDbl(100)
                End If
            End If
    End Select
    'If setor = 14 Then
     '   If Bairro = 12 Then
      '      m_ValorTerreno = 25
       ' ElseIf Bairro = 10 Then
        '    m_ValorTerreno = 15
        ' Else
          '  m_ValorTerreno = 1
        'End If
    'ElseIf setor = 10 And pac = True Then
     '   m_ValorTerreno = 36
    'End If
End Function




