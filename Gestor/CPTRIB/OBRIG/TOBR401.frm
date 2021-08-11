VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form TOBR4011 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   16455
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView Lst2 
      Height          =   2295
      Left            =   12360
      TabIndex        =   40
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4048
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
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
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   16410
      _ExtentX        =   28945
      _ExtentY        =   1138
      Icone           =   "TOBR401.frx":0000
   End
   Begin VB.Frame FraMenagem 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1800
      TabIndex        =   28
      Top             =   8040
      Width           =   4110
      Begin VB.Label LblMensagem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTRIBUINTE NOTIFICADO EM 01/01/2004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   29
         Top             =   240
         Width           =   4005
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   22
      Top             =   8310
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmd100Maiores 
         Height          =   375
         Left            =   9720
         TabIndex        =   39
         Top             =   120
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   661
         Caption         =   "&Maiores"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdAnaliseMensal 
         Height          =   375
         Left            =   11040
         TabIndex        =   38
         Top             =   120
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   661
         Caption         =   "&Alvará"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdRelatorio 
         Height          =   375
         Left            =   12360
         TabIndex        =   14
         Top             =   120
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   661
         Caption         =   "&Relatorio"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   15120
         TabIndex        =   16
         Top             =   120
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   13680
         TabIndex        =   15
         Top             =   120
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Index           =   3
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Width           =   12300
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   300
         Left            =   4470
         TabIndex        =   20
         Top             =   930
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   529
         Caption         =   "Nome/Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboRestricao 
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Tag             =   "Tributo"
         Top             =   1605
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   556
         Caption         =   "Restrição"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   7200
         TabIndex        =   7
         Tag             =   "Tributo"
         ToolTipText     =   "746 - STATUS OBRIGACAO"
         Top             =   1590
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   11280
         TabIndex        =   13
         Top             =   1965
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         Caption         =   "&Filtrar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   12240
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3390
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   540
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtExercicioInicial 
         Height          =   300
         Left            =   270
         TabIndex        =   8
         Tag             =   "Periodo Inicial"
         Top             =   1980
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         Caption         =   "Periodo Inicial"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtExercicioFinal 
         Height          =   300
         Left            =   2730
         TabIndex        =   9
         Tag             =   "Periodo Final"
         Top             =   1980
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         Caption         =   "Periodo Final"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoFinal 
         Height          =   300
         Left            =   8340
         TabIndex        =   11
         Top             =   1980
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         Caption         =   ""
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   5190
         TabIndex        =   10
         Tag             =   "Periodo Inicial"
         Top             =   1980
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   529
         Caption         =   "Periodo(dd/mm/aaaa)"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   3840
         TabIndex        =   2
         Top             =   540
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   7350
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   540
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtDAM 
         Height          =   300
         Left            =   8760
         TabIndex        =   3
         Top             =   600
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Número DAM"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   870
         TabIndex        =   0
         Tag             =   "Tributo"
         Top             =   135
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   705
         TabIndex        =   1
         Top             =   540
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   529
         Caption         =   "Inscricão"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   300
         Left            =   9720
         TabIndex        =   12
         Tag             =   "Periodo Final"
         Top             =   1980
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         Caption         =   "Valor"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         Requerido       =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtCpf 
         Height          =   300
         Left            =   600
         TabIndex        =   4
         Top             =   930
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   529
         Caption         =   " CPF/CNPJ"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaCpf 
         Height          =   315
         Left            =   3390
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   900
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.cboVISUAL cboAtividade 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Tag             =   "Tributo"
         Top             =   1245
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   556
         Caption         =   "Atividade"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   21
         Top             =   1590
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   17
      Top             =   90
      Width           =   375
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   795
   End
   Begin VTOcx.grdVISUAL GrdTaxas 
      Height          =   1620
      Left            =   4800
      TabIndex        =   26
      Top             =   8640
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   2858
      Caption         =   "Taxas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      CheckBox        =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtEnderecoContrib 
      Height          =   300
      Left            =   0
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "Periodo Final"
      Top             =   0
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   529
      Caption         =   "Periodo Final"
      Text            =   ""
      Restricao       =   2
      Requerido       =   0   'False
      MinLen          =   4
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtVISUAL1 
      Height          =   300
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   529
      Caption         =   "Inscricão"
      Text            =   ""
      Restricao       =   2
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin MSComctlLib.ListView Lst1 
      Height          =   4695
      Left            =   0
      TabIndex        =   33
      Top             =   3120
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
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
   Begin VTOcx.cmdVISUAL cmdAnterior 
      Height          =   375
      Left            =   13680
      TabIndex        =   34
      Top             =   7920
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "Anterior"
      Acao            =   8
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
      Icone           =   "TOBR401.frx":031A
   End
   Begin VTOcx.cmdVISUAL cmdProximo 
      Height          =   375
      Left            =   15120
      TabIndex        =   35
      Top             =   7920
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "Proximo"
      Acao            =   8
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
      Icone           =   "TOBR401.frx":063C
   End
   Begin VTOcx.cmdVISUAL cmdImprimir 
      Height          =   375
      Left            =   12240
      TabIndex        =   36
      Top             =   7920
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      Caption         =   "&DAM"
      Acao            =   4
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VB.Label LblOBS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6120
      TabIndex        =   32
      Top             =   8280
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "Geral"
      Visible         =   0   'False
      Begin VB.Menu mnuReimprime 
         Caption         =   "Reimprime"
         Index           =   0
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Consultar Parcelamento Espontâneo"
         Index           =   2
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Consultar Parcelamento de Ofício"
         Index           =   3
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Consultar Processo de DAT"
         Index           =   5
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Consultar Dados do Pagamento"
         Index           =   6
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Consultar Notas Fiscais"
         Index           =   7
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Cancelar"
         Index           =   9
      End
   End
End
Attribute VB_Name = "TOBR4011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Obrig As New Obrigacao
Dim Conta As New ContaCorrente
Dim Cobranca As New VSCobranca
Dim NovoJuro As String
Dim NovaMulta As String
Dim NovaData As String
Dim InscProprietario As String
Public String_Taxas    As String
Public Total_Taxas     As Double
Dim rst As Recordset
Dim sqlRelatorio As String
Dim sqlCubo As String
Dim sqlMaiores As String


Sub CabLancto()
      
      Lst1.ListItems.Clear
      Lst1.ColumnHeaders.Clear
      Lst1.ColumnHeaders.Add , , "Documento", 1000
      Lst1.ColumnHeaders.Add , , "Inscrição", 1650
      Lst1.ColumnHeaders.Add , , "Cpf/Cnpj", 1650
      Lst1.ColumnHeaders.Add , , "Contribuinte", 2500
      Lst1.ColumnHeaders.Add , , "Endereco", 2500
      Lst1.ColumnHeaders.Add , , "Tributo", 1000
      Lst1.ColumnHeaders.Add , , "Status", 1500
      Lst1.ColumnHeaders.Add , , "Gerado Em", 1000
      Lst1.ColumnHeaders.Add , , "Vencto", 1000
      Lst1.ColumnHeaders.Add , , "Pagto", 1000
      
      Lst1.ColumnHeaders.Add , , "Valor", 1000
      Lst1.ColumnHeaders.Add , , "Taxa", 1000
      Lst1.ColumnHeaders.Add , , "Obs", 2000
      Lst1.ColumnHeaders.Add , , "Atividade", 2000
      
      
      
      Lst1.ColumnHeaders.Item(11).Alignment = lvwColumnRight
      Lst1.ColumnHeaders.Item(12).Alignment = lvwColumnRight
      
      Lst1.View = lvwReport
   
End Sub

Sub loadSoma(Sql As String)
      Lst2.ListItems.Clear
      Lst2.ColumnHeaders.Clear
      Lst2.ColumnHeaders.Add , , "Situacao", 2300
      Lst2.ColumnHeaders.Add , , "Valor", 1350
      Lst2.ColumnHeaders.Item(2).Alignment = lvwColumnRight
      Lst2.View = lvwReport
      
      Dim Bd As Connection
      Dim Rs As New Recordset
      Dim mTot As Currency
      Dim Lst
       abreConexao Bd
      Set Rs = New ADODB.Recordset
   
      Rs.CursorLocation = adUseClient
      Rs.PageSize = 27
      Rs.Open Sql, Bd, adOpenStatic, adLockReadOnly, adAsyncFetchNonBlocking
      Set Rs.ActiveConnection = Nothing
      
      While (Not Rs.EOF)
   
        Set Lst = Lst2.ListItems.Add(, , carregaCampo(Rs!situacao))
        Lst.SubItems(1) = Format(carregaCampo(Rs!Valor), "##,##0.00")
        mTot = mTot + carregaCampo(Rs!Valor)
        Rs.MoveNext
        
      Wend
      Dim cor
      Set Lst = Lst2.ListItems.Add(, , "TOTAL")
      Lst.SubItems(1) = Format(mTot, "##,##0.00")
        cor = &HFF0000
        Lst.ForeColor = cor
        Lst.Bold = True
        Lst.ListSubItems(1).Bold = True
        
      
      fechaConexao Bd
      
      
End Sub

Sub LoadLancto(Sql As String)


On Error GoTo erros
    Dim mVlr As Currency
    Dim mVlP As Currency
    Dim mVlA As Currency
    Dim Bd As New Connection
    Dim Rs As Recordset
   
   '    Dim bd As New Connection
   ' Dim strconn As String
    'strconn = Bdados.Conexao.DBConnection.ConnectionString
    ''bd.Open strconn
     abreConexao Bd
    Set rst = New ADODB.Recordset
   
    rst.CursorLocation = adUseClient
    rst.PageSize = 27
    If Sql = "" Then Sql = "SELECT top 50 * FROM VIEW_IMOVEL"
    
    
    rst.Open Sql, Bd, adOpenStatic, adLockReadOnly, adAsyncFetchNonBlocking
    Set rst.ActiveConnection = Nothing
    
    Call ExibePag(1)
    
    fechaConexao Bd
    
    Exit Sub
erros:
  '  MOSTRAERRO Me.Name

End Sub

Sub ExibePag(pag As Double)
   
 '  If Rst.EOF Then
 '     Util.Informa "Não foram encontrados mais registros no criterio especificado!"
 '     Exit Sub
 '  End If
   Dim Contador As Integer
   Dim mVlr As Integer
   Dim tamanhoPagina As Integer
   Dim Lst
   If pag = 0 Then pag = 1
   rst.AbsolutePage = pag
   tamanhoPagina = rst.PageSize
   Contador = 1
    
    
    
   Lst1.ListItems.Clear
   mVlr = 0
   
   While (Not rst.EOF) And (Contador <= tamanhoPagina)
   
        Set Lst = Lst1.ListItems.Add(, , carregaCampo(rst!Documento))
        Lst.SubItems(1) = carregaCampo(rst!Inscricao)
        
        Lst.SubItems(2) = carregaCampo(rst!CPFCNPJ)
        
        Lst.SubItems(3) = carregaCampo(rst!Contribuinte)
        Lst.SubItems(4) = carregaCampo(rst!Endereco)
        
        Lst.SubItems(5) = carregaCampo(rst("nome_imposto"))
        Lst.SubItems(6) = carregaCampo(rst!situacao)
        Lst.SubItems(7) = Format(carregaCampo(rst!Geracao), "dd/mm/YY")
        Lst.SubItems(8) = Format(carregaCampo(rst!Vencto), "dd/mm/YY")
        Lst.SubItems(9) = Format(carregaCampo(rst!Pagamento), "dd/mm/YY")
        
        Lst.SubItems(10) = Format(carregaCampo(rst!Valor), "##,##0.00")
        Lst.SubItems(11) = Format(carregaCampo(rst!Taxa), "##,##0.00")
        Lst.SubItems(12) = carregaCampo(rst!Observacao)
        Lst.SubItems(13) = carregaCampo(rst!Atividade)
        
        
     '   Lst.ListSubItems(9).ForeColor = Cor
      
         ' 12/trans
         ' 8 cancelado
         Dim cor, x
         If rst!Status = 3 Then ' pago
            cor = &HFF&
         ElseIf rst!Status = 2 Then 'aberto
            cor = &HFF0000
         ElseIf rst!Status = 12 Then 'aberto
            cor = &HFF0000
         End If
         
            Lst.ForeColor = cor
            For x = 1 To Lst.ListSubItems.Count
                Lst.ListSubItems(x).ForeColor = cor
            Next
        ' End If
        'cor = &HFF0000
      '   cor = &HFF0000
        cor = &H0&
        Lst.ListSubItems(10).ForeColor = cor
        Lst.ListSubItems(10).Bold = True
        
        Contador = Contador + 1
        rst.MoveNext
    Wend

End Sub

Private Function CriticaCampos() As Boolean
    CriticaCampos = True
    If Not Edita.CriticaCampos(Me) Then
        CriticaCampos = False
        Exit Function
    End If
    If Len(txtPeriodoInicial) <> Len(txtPeriodoFinal) Then
        Avisa "Período inconsistente."
        txtPeriodoInicial.SetFocus
        CriticaCampos = False
        Exit Function
    End If
    If Len(txtPeriodoInicial) > 4 Then
        If Right(Trim(txtPeriodoInicial), 4) <> Right(Trim(txtPeriodoFinal), 4) Then
            Avisa "Período deve ser dentro do mesmo ano."
            txtPeriodoInicial.SetFocus
            CriticaCampos = False
        End If
    End If
End Function



Private Sub cmd100Maiores_Click()
Dim Sql As String
Sql = "SELECT Identity(INT,1,1) As Posicao, CC.Contribuinte as Nome, CAST(SUM(cc.VALOR) AS Money) As SubTotal Into TAB_CURVA_ABC"
Sql = Sql & "  FROM View_Obrigacao As CC"
Sql = Sql & " WHERE " & sqlMaiores
Sql = Sql & " GROUP BY cc.Contribuinte ORDER BY SubTotal DESC"

Dim Bd As Connection
Dim regs As Long

abreConexao Bd
Bd.Execute "DROP TABLE TAB_CURVA_ABC"
Bd.Execute Sql, regs
fechaConexao Bd
form100MaioresContribuintes.Show


End Sub

Private Sub cmdAnaliseMensal_Click()
    formContribuinteAlvara.Show

End Sub

Private Sub cmdAnterior_Click()
    
    If rst.AbsolutePage > 2 And rst.AbsolutePage <> -2 Then
       Call ExibePag(rst.AbsolutePage - 2)
    Else
        If rst.AbsolutePage = -3 Then
           Call ExibePag(rst.PageCount - 1)
        Else
           Avisa " Não há mais pagina"
           Call ExibePag(1)
        End If
    End If
    
End Sub


Private Sub cmdBuscar_Click()
    Dim Obrig As Obrigacao
    Dim Inscri As String
    Set Obrig = New Obrigacao
    
    Inscri = txtIm
    
    Dim Sql As String
    Dim bCanceladas As Boolean
    
   ' Criterio2 = ""
    Dim sqlCpf
    Dim sqlImposto
    Dim sqlInscr
    Dim sqlImovel
    Dim sqlRestricao
    Dim sqlStatus
    Dim sqlExercicio
    Dim sqlPeriodo
    Dim sqlDocto
    Dim sqlValor
    Dim txtInicial
    Dim txtFinal
    Dim sqlAtiv
    
    

    CabLancto
    
    txtInicial = txtPeriodoInicial
    txtFinal = txtPeriodoFinal
    
    
    
    If txtInicial <> "" And txtFinal <> "" Then
        If CDate(txtInicial) > CDate(txtFinal) Then
           Util.Mensagem "Digitacao Invalida" & vbCr & " Porque a ordem das datas está incorreta !"
           Exit Sub
        End If
        sqlPeriodo = " Vencto Between '" & DtoSQL(txtInicial) & "' And '" & DtoSQL(txtFinal) & "'"
    End If
     
    If cboAtividade <> "" Then sqlAtiv = "Atividade='" & cboAtividade.Text & "'"
    If txtCpf <> "" Then sqlCpf = "CpfCnpj='" & txtCpf & "'"
    If txtValor <> "" Then sqlValor = " Valor = " & NumeroSQL(txtValor)
    If txtDAM <> "" Then sqlDocto = "Documento Like %'" & txtDAM & "%'"
    If cboImposto <> "" Then sqlImposto = "Imposto='" & cboImposto.coluna(0).Valor & "'"
    If cboStatus <> "" Then sqlStatus = "Status=" & cboStatus.coluna(1).Valor
    If CInt(cboRestricao.coluna(1).Valor) = 1 Then sqlRestricao = "STATUS NOT IN (" & etsCreditoPago & "," & etsCreditoIsento & ")"
    If CInt(cboRestricao.coluna(1).Valor) = 2 Then sqlRestricao = "STATUS = " & etsCreditoPago
    
     If Trim(txtIm) <> "" Then
            If VBA.InStr(1, txtIm, "-") > 0 Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then  'INSCRICAO MUNICIPAL
                 If Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM" Then
                    Dim SqlAux As String
                    Dim Rs As VSRecordset
                    Dim Inscricoes
                    SqlAux = "select tim_ic from tab_imovel where  tim_tci_im = '" & txtIm & "'"
                    If Bdados.AbreTabela(SqlAux, Rs) Then
                        Rs.MoveFirst
                        Inscricoes = ""
                        Do While Not Rs.EOF
                            Inscricoes = Inscricoes & "'" & Trim(Rs!TIM_IC) & "',"
                            Rs.MoveNext
                        Loop
                        Inscricoes = Left(Inscricoes, Len(Inscricoes) - 1)
                        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Or True Then
                            sqlInscr = "((INSCRICAO = '" & Trim(txtIm) & "' AND TIPO_INSCRICAO = 2) or (inscricao in (  " & Trim(Inscricoes) & ") AND TIPO_INSCRICAO = 1))"
                        Else
                            sqlInscr = "((INSCRICAO = '" & Trim(txtIm) & "' or inscricao in (  " & Inscricoes & "))"
                        End If
                    Else
                        sqlInscr = "INSCRICAO = '" & Trim(txtIm) & "' AND TIPO_INSCRICAO = 2"
                    End If
                Else
                    sqlInscr = "INSCRICAO = '" & Trim(txtIm) & "' AND TIPO_INSCRICAO = 2"
                End If
           End If
      ElseIf Trim(txtImovel) <> "" Then   ' INSCRICAO CADASTRAL(imovel)
                sqlInscr = "INSCRICAO LIKE '" & Trim(txtImovel) & "%' AND TIPO_INSCRICAO = 1"
        End If

    Dim col As New Collection
    Dim SqlSoma As String
    
    col.Add sqlPeriodo
    col.Add sqlValor
    col.Add sqlImposto
    col.Add sqlStatus
    col.Add sqlRestricao
    col.Add sqlInscr
    col.Add sqlDocto
    col.Add sqlCpf
    col.Add sqlAtiv
  
    
  '  col.Add sqlImovel
  
    
    Sql = "SELECT * FROM VIEW_OBRIGACAO"
    SqlSoma = "SELECT SITUACAO,SUM(VALOR) AS VALOR FROM VIEW_OBRIGACAO"
    
    Sql = montaSqlWhere(Sql, col)
    SqlSoma = montaSqlWhere(SqlSoma, col)
    
    'Sql = montaSQL(Sql)
    sqlMaiores = montaClausulaSqlWhere(col)
    sqlCubo = Sql
    sqlRelatorio = Sql & " ORDER BY CONTRIBUINTE,SITUACAO,NOME_IMPOSTO,GERACAO"
    SqlSoma = SqlSoma & " GROUP BY SITUACAO"
    
    LoadLancto Sql
    loadSoma SqlSoma
End Sub

Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    cboImposto.SetFocus
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub


Private Sub cmdImprimir_Click()
If Lst1.SelectedItem Is Nothing Then Exit Sub
   TOBR401.MostraObrigacao (Lst1.SelectedItem)
  ' formObrigConsulta.Show
   
   Unload TOBR401
  
End Sub

Private Sub cmdPesquisaCpf_Click()

  AplicacoesVTFuncoes.BuscaInscricao InscCpfCnpj, txtCpf
  
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdProxima_Click()
   
End Sub

Private Sub cmdProximo_Click()

 If rst.AbsolutePage <> -3 Then
      Call ExibePag(rst.AbsolutePage)
      cmdAnterior.Enabled = True
    Else
      Avisa " Não Há Mais Pagina! "
      Call ExibePag(rst.PageCount)
    End If

End Sub

Private Sub cmdRelatorio_Click()
    
    Dim CondRelatorio As String
    Dim FORMULA As String
    Dim Inscricao As String
    Dim relCons As New relObrigConsulta
    configRelatorio relCons, sqlRelatorio
    relCons.Show
  
  
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Activate()
    If Left(Trim(Me.Tag), 1) = "C" Then
        txtIm = Mid(Me.Tag, 2)
        txtIm_LostFocus
    ElseIf Left(Trim(Me.Tag), 1) = "I" Then
        txtImovel = Mid(Me.Tag, 2)
        txtImovel_LostFocus
    ElseIf Len(Trim(Me.Tag)) > 0 Then
        txtIm = Me.Tag
        txtIm_LostFocus
    End If
    FraMenagem.Visible = False
        LblMensagem.Visible = False
        txtIm.SetFocus
End Sub
Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Obrig.PreencheComboTributo cboImposto, False
    cboStatus.PreencherGeral Bdados, "STATUS OBRIGACAO"
    cboRestricao.PreencherGeral Bdados, "RESTRICAO DAM"
    carregaCombo cboAtividade, "select tae_nome from tab_atividade_economica order by tae_nome"
    
End Sub

Private Sub ExibeContribuinte(Ic As String)
    If Len(Ic) < 15 Then
        txtImovel = ""
        txtIm = Ic
        txtIm = BuscaContribuinte(Ic, txtRazao, txtEndereco)
    Else
        txtIm = ""
        txtImovel = Ic
        txtImovel = BuscaContribuinte(Ic, txtRazao, txtEndereco, InscProprietario, etiImovel)
    End If
End Sub




Private Sub txtCpf_LostFocus()
'
'If Len(txtCpf.Text) <= 11 Then
'   txtCpf.Formato = formCPF
'Else
'   txtCpf.Formato = formCGC
'End If

End Sub

Private Sub txtIm_LostFocus()
    Dim Ic As String
    If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        If Len(txtIm) = 10 Or Len(txtIm) = 11 Then
            Ic = Imposto.FormataInscricao(txtIm, InscContrib)
        Else
            Ic = txtIm
        End If
    Else
        Ic = txtIm
    End If
    If Trim(txtIm) <> "" Then
        txtIm = BuscaContribuinte(Ic, txtRazao, txtEndereco)
        If Trim(txtIm) = "" Then
            Avisa "Inscricão não encontrada"
            txtRazao = ""
            txtEndereco = ""
            
            txtIm.SetFocus
        End If
    End If
    
End Sub
Private Sub Pega_taxas()
    Dim i As Integer
    Dim Pos As Integer
    String_Taxas = ""
    Total_Taxas = 0
    For i = 1 To GrdTaxas.ListItems.Count
        If GrdTaxas.ListItems(i).Checked Then
            Pos = InStr(GrdTaxas.ListItems(i).SubItems(1), "-") - 1
            If String_Taxas = "" Then
                String_Taxas = String_Taxas & " [ " & Left(GrdTaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(GrdTaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            Else
                String_Taxas = String_Taxas & ", [ " & Left(GrdTaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(GrdTaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            End If
            Total_Taxas = Total_Taxas + CCur(GrdTaxas.ListItems(i).SubItems(2))
        End If
    Next
End Sub

Private Sub txtImovel_LostFocus()
    Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        If Trim(txtImovel) = 14 Then txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, InscProprietario, etiImovel)
       ' If Trim(txtImovel) = "" Then
       '     Avisa "Inscricão não encontrada"
       '     txtRazao = ""
       '     txtEndereco = ""
       '     txtImovel.SetFocus
       ' End If
    End If
End Sub
