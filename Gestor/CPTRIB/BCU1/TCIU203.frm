VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{F859D2AC-288F-4C12-BA09-A9F93FADC107}#1.0#0"; "orioncontrols.ocx"
Begin VB.Form TCIU203a 
   BackColor       =   &H00FBEDE8&
   Caption         =   "Gerenciamento de IMOVEIS"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16500
   ForeColor       =   &H00000000&
   Icon            =   "TCIU203.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8985
   ScaleMode       =   0  'User
   ScaleWidth      =   18660.31
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CboRelatorio 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5040
      TabIndex        =   14
      Top             =   8520
      Width           =   3495
   End
   Begin VB.Frame FraFiltro 
      BackColor       =   &H00FBEDE8&
      Caption         =   "Filtro"
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   16455
      Begin ORIONControls.txtORION txtInsc 
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "Inscricao"
         Text            =   ""
         AlinhamentoRotulo=   1
         EnterEqvTab     =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin ORIONControls.cmdORION cmdFiltro 
         Height          =   375
         Left            =   15360
         TabIndex        =   5
         ToolTipText     =   "Credito referentes a venda de veiculos"
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "Filtrar"
         Acao            =   5
         CorFundo        =   16776960
         CorFoco         =   65535
      End
      Begin ORIONControls.cboORION cboEndereco 
         Height          =   510
         Left            =   4680
         TabIndex        =   2
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   900
         Caption         =   "Endereco"
         Text            =   ""
         AutoFocaliza    =   0   'False
         TipoLetras      =   0
         Alinhamento     =   1
         EnterEqvTab     =   0   'False
         Editavel        =   -1  'True
         Sorted          =   0   'False
      End
      Begin ORIONControls.txtORION txtPessoa 
         Height          =   495
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   873
         Caption         =   "Pessoa"
         Text            =   ""
         AlinhamentoRotulo=   1
         EnterEqvTab     =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin ORIONControls.cboORION cboTipo 
         Height          =   510
         Left            =   13200
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   900
         Caption         =   "Tipo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         TipoLetras      =   0
         Alinhamento     =   1
         EnterEqvTab     =   0   'False
         Editavel        =   -1  'True
         Sorted          =   0   'False
      End
      Begin ORIONControls.cboORION cboBairro 
         Height          =   510
         Left            =   9120
         TabIndex        =   3
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   900
         Caption         =   "Bairro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         TipoLetras      =   0
         Alinhamento     =   1
         EnterEqvTab     =   0   'False
         Editavel        =   -1  'True
         Sorted          =   0   'False
      End
   End
   Begin MSComctlLib.ListView Lst1 
      Height          =   6735
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   11880
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ORIONControls.cmdORION cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   15120
      TabIndex        =   7
      ToolTipText     =   "Credito referentes a venda de veiculos"
      Top             =   8520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Sair"
      Acao            =   7
      CorFundo        =   16776960
      CorFoco         =   65535
   End
   Begin ORIONControls.cmdORION cmdIncluir 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Credito referentes a venda de veiculos"
      Top             =   8520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Incluir"
      Acao            =   1
      CorFundo        =   16776960
      CorFoco         =   65535
   End
   Begin ORIONControls.cmdORION cmdAnterior 
      Height          =   375
      Left            =   12240
      TabIndex        =   9
      ToolTipText     =   "Gravar os dados informados na tabela"
      Top             =   8520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Anterior"
      Acao            =   8
      CorFundo        =   16776960
      CorFoco         =   65535
      Icone           =   "TCIU203.frx":08CA
   End
   Begin ORIONControls.cmdORION cmdProxima 
      Height          =   375
      Left            =   13680
      TabIndex        =   10
      ToolTipText     =   "Gravar os dados informados na tabela"
      Top             =   8520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Proxima"
      Acao            =   8
      CorFundo        =   16776960
      CorFoco         =   65535
      Icone           =   "TCIU203.frx":0BE4
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   1138
      Formulario      =   "Relaçao de IMOVEIS"
      Descricao       =   ""
      Icone           =   "TCIU203.frx":0EFE
   End
   Begin ORIONControls.cmdORION cmdImprimir 
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      ToolTipText     =   "Credito referentes a venda de veiculos"
      Top             =   8520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Imprimir"
      Acao            =   4
      CorFundo        =   16776960
      CorFoco         =   65535
   End
End
Attribute VB_Name = "TCIU203a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_Placa As String
Public m_Lancto As Double
Private SqlStr As String
Private Rst As Recordset
Private SqlRel As String
Private lColoriu As Boolean
Dim vetorValor() As Variant
Dim lFiltrou As Boolean

Sub CabLancto()
      
      Lst1.ListItems.Clear
      Lst1.ColumnHeaders.Clear
      Lst1.ColumnHeaders.Add , , "Inscricao", 2000
      Lst1.ColumnHeaders.Add , , "Contribuinte", 3500
      Lst1.ColumnHeaders.Add , , "Logr", 1000
      Lst1.ColumnHeaders.Add , , "Endereco", 3500
      Lst1.ColumnHeaders.Add , , "Numero", 1500
      Lst1.ColumnHeaders.Add , , "Bairro", 3500
       Lst1.ColumnHeaders.Add , , "Tipo", 1500
      Lst1.View = lvwReport
   
End Sub

Sub LoadLancto(Sql As String)


On Error GoTo erros
    Dim mVlr As Currency
    Dim mVlP As Currency
    Dim mVlA As Currency
    
    Dim Rs As Recordset
   '    Dim bd As New Connection
   ' Dim strconn As String
    'strconn = Bdados.Conexao.DBConnection.ConnectionString
    ''bd.Open strconn
    'bd = abreConexao()
    'Set Rst = New ADODB.Recordset
   
    'Rst.CursorLocation = adUseClient
    'Rst.PageSize = 27
    'If Sql = "" Then Sql = "SELECT top 50 * FROM VIEW_IMOVEL"
    
    
    'Rst.Open Sql, bd, adOpenStatic, adLockReadOnly, adAsyncFetchNonBlocking
    'Set Rst.ActiveConnection = Nothing
    
    'Call ExibePag(1)
    
    'fechaConexao (bd)
    
    Exit Sub
erros:
  '  MOSTRAERRO Me.Name

End Sub

Sub ExibePag(pag As Double)
   
 '  If Rst.EOF Then
 '     Util.Informa "Não foram encontrados mais registros no criterio especificado!"
 '     Exit Sub
 '  End If
   If pag = 0 Then pag = 1
   Rst.AbsolutePage = pag
   TamanhoPagina = Rst.PageSize
   Contador = 1
    
    
    
   Lst1.ListItems.Clear
   mVlr = 0
   
   While (Not Rst.EOF) And (Contador <= TamanhoPagina)
   
        Set Lst = Lst1.ListItems.Add(, , carregaCampo(Rst!Inscricao))
        Lst.SubItems(1) = carregaCampo(Rst!Pessoa)
        Lst.SubItems(2) = carregaCampo(Rst("Logradouro"))
        Lst.SubItems(3) = carregaCampo(Rst!Endereco)
        Lst.SubItems(4) = carregaCampo(Rst!Numero)
        Lst.SubItems(5) = carregaCampo(Rst!Bairro)
        Lst.SubItems(6) = carregaCampo(Rst!Tipo)
        
        Contador = Contador + 1
        Rst.MoveNext
    Wend

End Sub

Private Sub cmdAnterior_Click()
    If Rst.AbsolutePage > 2 And Rst.AbsolutePage <> -2 Then
       Call ExibePag(Rst.AbsolutePage - 2)
    Else
        If Rst.AbsolutePage = -3 Then
           Call ExibePag(Rst.PageCount - 1)
        Else
           Avisa " Não há mais pagina"
           Call ExibePag(1)
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub




Private Sub cmdFiltro_Click()
Dim Sql As String
Dim bCanceladas As Boolean

Criterio2 = ""
Dim sqlInsc
Dim sqlPessoa
Dim sqlEndereco
Dim sqlBairro
Dim sqlTipo

If txtInsc <> "" Then sqlInsc = "Inscricao Like '%" & txtInsc & "%'"
If txtPessoa <> "" Then sqlPessoa = "Pessoa Like '%" & txtPessoa & "%'"
If cboEndereco <> "" Then sqlEndereco = "Endereco Like '%" & cboEndereco & "%'"
If cboBairro <> "" Then sqlBairro = "Bairro  ='" & cboBairro & "'"
If cboTipo <> "" Then sqlTipo = "Tipo = '" & cboTipo & "'"

Dim col As New Collection
col.Add sqlInsc
col.Add sqlPessoa
col.Add sqlEndereco
col.Add sqlBairro
col.Add sqlTipo

Sql = "SELECT * FROM VIEW_IMOVEL"
Sql = montaSqlWhere(Sql, col)

LoadLancto Sql

End Sub
 
Private Sub cmdIncluir_Click()
    
    formCadImovel.Show vbModal
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdProxima_Click()
    If Rst.AbsolutePage <> -3 Then
      Call ExibePag(Rst.AbsolutePage)
      cmdAnterior.Enabled = True
    Else
      Avisa " Não Há Mais Pagina! "
      Call ExibePag(Rst.PageCount)
    End If
End Sub

Private Sub Form_Activate()
  '  LoadLancto ""
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
  
   ' If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub Form_Load()
    CabLancto
    
    'carregando combos
    carregaCombo cboEndereco, "SELECT TLG_NOME FROM TAB_LOGRADOURO ORDER BY TLG_NOME"
    carregaCombo cboBairro, "SELECT TBA_NOME FROM TAB_BAIRRO ORDER BY TBA_NOME"
    carregaCombo cboTipo, "EXEC sp_buscar_tabela_geral 'TIPO LOTE'"

End Sub

Private Sub Lst1_DblClick()

    formCadImovel.txtInscImob = Lst1.SelectedItem
    formCadImovel.Show vbModal

End Sub
