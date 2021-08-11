VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'TCL_DISTRITO
'TCL_SETOR
'TCL_QUADRA
'TCL_QTD_BT
'TCL_QTD_BP
'TCL_QTD_BC
'TCL_DATA_ABERTURA
'TCL_STATUS
'TCL_STATUS_DATA
'TCL_TUS_COD_USUARIO
Private m_Distrito As String
Private m_Setor As String
Private m_Quadra As String
Private m_QtdBT As Double
Private m_DigBT As Double
Private m_QtdBP As Double
Private m_DigBP As Double
Private m_QtdBC As Double
Private m_DigBC As Double
Private m_DataAbertura As Date
Private m_Status As StatusGradeLote
Private m_StatusData As Date
Private m_Usuario As String

Public Property Get Usuario() As String
    Usuario = m_Usuario
End Property

Public Property Let Usuario(ByVal Value As String)
    m_Usuario = Value
End Property

Public Property Get StatusData() As Date
    StatusData = m_StatusData
End Property

Public Property Let StatusData(ByVal Value As Date)
    m_StatusData = Value
End Property

Public Property Get Status() As StatusGradeLote
    Status = m_Status
End Property

Public Property Let Status(ByVal Value As StatusGradeLote)
    m_Status = Value
End Property

Public Property Get DataAbertura() As Date
    DataAbertura = m_DataAbertura
End Property

Public Property Let DataAbertura(ByVal Value As Date)
    m_DataAbertura = Value
End Property

Public Property Get QtdBC() As Double
    QtdBC = m_QtdBC
End Property

Public Property Let QtdBC(ByVal Value As Double)
    m_QtdBC = Value
End Property

Public Property Get QtdBP() As Double
    QtdBP = m_QtdBP
End Property

Public Property Let QtdBP(ByVal Value As Double)
    m_QtdBP = Value
End Property

Public Property Get QtdBT() As Double
    QtdBT = m_QtdBT
End Property

Public Property Let QtdBT(ByVal Value As Double)
    m_QtdBT = Value
End Property

Public Property Get Quadra() As String
    Quadra = m_Quadra
End Property

Public Property Let Quadra(ByVal Value As String)
    m_Quadra = Value
End Property

Public Property Get Setor() As String
    Setor = m_Setor
End Property

Public Property Let Setor(ByVal Value As String)
    m_Setor = Value
End Property

Public Property Get Distrito() As String
    Distrito = m_Distrito
End Property

Public Property Let Distrito(ByVal Value As String)
    m_Distrito = Value
End Property

Public Function Gravar() As Boolean
    Dim sCampos As String
    Dim sValores As String
    Dim sCondicao As String
    On Error GoTo Trata
    sCampos = "TCL_DISTRITO,TCL_SETOR,TCL_QUADRA,TCL_QTD_BT,TCL_QTD_BP,TCL_QTD_BC,TCL_DATA_ABERTURA,TCL_TUS_COD_USUARIO,TCL_STATUS,TCL_STATUS_DATA"
    sValores = Bdados.PreparaValor(Bdados.Converte(m_Distrito, vsclass.tctexto), Bdados.Converte(m_Setor, vsclass.tctexto), Bdados.Converte(m_Quadra, vsclass.tctexto), m_QtdBT, m_QtdBP, m_QtdBC, m_DataAbertura, Aplicacoes.Usuario, m_Status, m_StatusData)
    sCondicao = "TCL_DISTRITO = '" & m_Distrito & "' AND TCL_SETOR='" & m_Setor & "' AND TCL_QUADRA = '" & m_Quadra & "'"
    Gravar = Bdados.GravaDados("TAB_CONTROLE_LOTE", sValores, sCampos, sCondicao)
    Exit Function
Trata:
    Util.Erro Err.Description
End Function

Public Sub FecharQuadras()
    Dim Sql As String
    Dim Rs As VSRecordset
    Screen.MousePointer = 11
    Sql = "SELECT TCL_DISTRITO,TCL_SETOR,TCL_QUADRA FROM TAB_CONTROLE_LOTE" 'WHERE TCL_STATUS <> 2 "
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            FechaLote Trim("" & Rs!TCL_DISTRITO), Trim("" & Rs!TCL_SETOR), Trim("" & Rs!TCL_QUADRA)
            DoEvents
            Rs.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
End Sub


Public Function Buscar(Distrito As String, Setor As String, Quadra As String, Optional Status As StatusGradeLote) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    On Error GoTo Trata
    Sql = "SELECT * FROM VIS_LOTE_DIGITACAO WHERE TCL_DISTRITO = '" & Distrito & "' AND TCL_SETOR='" & Setor & "' AND TCL_QUADRA = '" & Quadra & "'"
    Sql = Sql & IIf(Status <> 0, " AND TCL_STATUS =" & Status, "")
    If Bdados.AbreTabela(Sql, Rs) Then
        m_Distrito = "" & Rs!TCL_DISTRITO
        m_Setor = "" & Rs!TCL_SETOR
        m_Quadra = "" & Rs!TCL_QUADRA
        m_QtdBT = "" & Rs!TCL_QTD_BT
        m_DigBT = Nvl("" & Rs!VBT_QTD, 0)
        m_QtdBP = "" & Rs!TCL_QTD_BP
        m_DigBP = Nvl("" & Rs!VBP_QTD, 0)
        m_QtdBC = "" & Rs!TCL_QTD_BC
        m_DigBC = Nvl("" & Rs!VBC_QTD, 0)
        m_DataAbertura = "" & Rs!TCL_DATA_ABERTURA
        m_Usuario = "" & Rs!VBT_USUARIO
        m_Status = "" & Rs!TCL_STATUS
        m_StatusData = "" & Rs!TCL_STATUS_DATA
        Buscar = True
    End If
    Exit Function
Trata:
    Util.Erro Err.Description
    'Resume
End Function

Public Function Excluir(ByRef Erro As String) As Boolean
    On Error GoTo Trata
    If m_DigBT = 0 Then
        Excluir = Bdados.DeletaDados("TAB_CONTROLE_LOTE", "TCL_DISTRITO = '" & m_Distrito & "' AND TCL_SETOR='" & m_Setor & "' AND TCL_QUADRA = '" & m_Quadra & "'")
    Else
        Erro = "Lote não vazio."
    End If
    Exit Function
Trata:
    Util.Erro Err.Description
End Function

Public Function PreencherGrid(grid As Object, Optional Situacao As StatusGradeLote) As Boolean
    Dim Sql As String
    Sql = "SELECT " & _
                    "TCL_DISTRITO AS Distrito, " & _
                    "TCL_SETOR AS Setor, " & _
                    "TCL_QUADRA AS Quadra, " & _
                    "TCL_QTD_BT AS BT, " & _
                    "TCL_QTD_BP AS BP, " & _
                    "TCL_QTD_BC AS BC," & _
                    "SUM(VBT_QTD) " & _
                "FROM VIS_LOTE_DIGITACAO "
    If Situacao > 0 Then
        Sql = Sql & " WHERE TCL_STATUS =" & Situacao
    End If
    Sql = Sql & " GROUP BY " & _
                    "TCL_DISTRITO, " & _
                    "TCL_SETOR, " & _
                    "TCL_QUADRA, " & _
                    "TCL_QTD_BT, " & _
                    "TCL_QTD_BP, " & _
                    "TCL_QTD_BC"
    Sql = Sql & " ORDER BY 1,2,3"
    grid.Preencher Bdados, Sql, (grid.Width * 16 / 100), (grid.Width * 16 / 100), (grid.Width * 16 / 100), (grid.Width * 16 / 100), (grid.Width * 16 / 100), (grid.Width * 16 / 100), 0
    PreencherGrid = grid.ListItems.Count > 0
End Function

Public Function ExibirBatimento(grid As Object, Optional Usuario As String, Optional Situacao As String) As Boolean
    Dim Sql As String, Where As String
    
    Sql = "SELECT " & _
                    "TCL_DISTRITO AS Distrito, " & _
                    "TCL_SETOR AS Setor, " & _
                    "TCL_QUADRA AS Quadra, " & _
                    "TCL_QTD_BT AS [BT Inf], " & _
                    "SUM(VBT_QTD) as [BT Dig], " & _
                    "TCL_QTD_BP AS [BP Inf], " & _
                    "SUM(VBP_QTD) as [BP Dig], " & _
                    "TCL_QTD_BC AS [BC Inf]," & _
                    "SUM(VBC_QTD) as [BC Dig], " & _
                    "TGE_NOME as Situacao " & _
                    " " & _
                "FROM VIS_LOTE_DIGITACAO "
    Where = ""
    If Usuario <> "" Then
        Where = " WHERE VBT_USUARIO ='" & Usuario & "'"
    End If
    If Situacao <> "" Then
        Where = Where & IIf(Where = "", " WHERE ", " AND ") & " TGE_NOME='" & Situacao & "'"
    End If
    Sql = Sql & Where
    Sql = Sql & " GROUP BY " & _
                    "TCL_DISTRITO, " & _
                    "TCL_SETOR, " & _
                    "TCL_QUADRA, " & _
                    "TCL_QTD_BT, " & _
                    "TCL_QTD_BP, " & _
                    "TCL_QTD_BC," & _
                    "TGE_NOME"
    Sql = Sql & " ORDER BY 1,2,3"
    grid.Preencher Bdados, Sql
    If grid.ListItems.Count > 0 Then
        grid.Mensagem = "BT : " & grid.Colunas(5).Soma & "       BP : " & grid.Colunas(7).Soma & "       BC : " & grid.Colunas(9).Soma & "       Total : " & grid.Colunas(5).Soma + grid.Colunas(7).Soma + grid.Colunas(9).Soma
    Else
        grid.Mensagem = "Nenhum registro encontrado."
    End If
    ExibirBatimento = grid.ListItems.Count > 0
End Function

Public Function FechaLote(Distrito As String, Setor As String, Quadra As String, Optional AtualizaSomenteFechados As Boolean = False) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim BTs As Double
    Dim BPs As Double
    Dim BCs As Double
    
    FechaLote = False
    Sql = "select tdi_tgc_cod_grupo , count(*) as Total from tab_detalhe_imovel " & _
        " where tdi_tgc_cod_grupo in (10,108) and tdi_tim_ic like '" & Distrito & Setor & Quadra & "%' group by tdi_tgc_cod_grupo"
    If Bdados.AbreTabela(Sql, Rs) Then
        Do
            If Rs!tdi_tgc_cod_grupo = 10 Then BPs = Rs!Total
            If Rs!tdi_tgc_cod_grupo = 108 Then BTs = Rs!Total
            Rs.MoveNext
        Loop While Not Rs.EOF
    End If
    Bdados.FechaTabela Rs
    Sql = "select count(*) as Total from tab_imovel " & _
        " where tim_unidade > 0 and tim_unidade not in (200,600) and tim_ic like '" & Distrito & Setor & Quadra & "%'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BCs = Rs!Total
    End If
    Bdados.FechaTabela Rs
    Sql = "Select * from TAB_CONTROLE_LOTE where TCL_DISTRITO = '" & Distrito & "' AND TCL_SETOR='" & Setor & "' AND TCL_QUADRA = '" & Quadra & "' and TCL_QTD_BT =" & BTs & " and TCL_QTD_BP =" & BPs & " and TCL_QTD_BC = " & BCs
    If Bdados.AbreTabela(Sql, Rs) Then
        FechaLote = True
        Bdados.AtualizaDados "TAB_CONTROLE_LOTE", Bdados.PreparaValor(2), "TCL_STATUS", "TCL_DISTRITO = '" & Distrito & _
            "' AND TCL_SETOR='" & Setor & "' AND TCL_QUADRA = '" & Quadra & _
            "' and TCL_QTD_BT =" & BTs & " and TCL_QTD_BP =" & BPs & " and TCL_QTD_BC = " & BCs
    Else
        FechaLote = False
        If AtualizaSomenteFechados Then Exit Function
        Bdados.AtualizaDados "TAB_CONTROLE_LOTE", Bdados.PreparaValor(1), "TCL_STATUS", "TCL_DISTRITO = '" & Distrito & _
            "' AND TCL_SETOR='" & Setor & "' AND TCL_QUADRA = '" & Quadra & _
            "' "
    End If
End Function


