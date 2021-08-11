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
Private m_ValorTerreno As Double
Private m_ValorImovelResidencial As Double
Private m_ValorImovelComercial As Double
Private m_AliquotaImovelResidencial As Double
Private m_AliquotaImovelComercial As Double
Private m_AliquotaTerrenoMurado As Double
Private m_AliquotaTerrenoBaldio As Double
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
Public Function BuscarPorSetor(Periodo As Integer, Bairro As Integer, setor As Integer, quadra As Integer, pac As Boolean) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "SELECT * FROM TAB_BCP_PARAMETRO_IPTU WHERE TPI_PERIODO=" & Periodo & " AND TPI_SETOR=" & setor & " AND TPI_QUADRA=" & quadra
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
        BuscarPorSetor = True
    Else
        Sql = "SELECT * FROM TAB_BCP_PARAMETRO_IPTU WHERE TPI_PERIODO=" & Periodo & " AND TPI_SETOR=" & setor & " AND TPI_QUADRA=0"
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
            BuscarPorSetor = True
        End If
    End If
    If setor = 14 Then
        If Bairro = 12 Then
            m_ValorTerreno = 25
        ElseIf Bairro = 10 Then
            m_ValorTerreno = 15
        Else
            m_ValorTerreno = 1
        End If
    ElseIf setor = 10 And pac = True Then
        m_ValorTerreno = 36
    End If
End Function




