VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form frmImprimir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   1138
      Icone           =   "ImprimirNovo.frx":0000
   End
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   1605
      Left            =   60
      TabIndex        =   10
      Top             =   1110
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   2831
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "ImprimirNovo.frx":0A42
      Images          =   "ImprimirNovo.frx":0AC6
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   1185
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   2090
         _Version        =   131082
         TabGuid         =   "ImprimirNovo.frx":160E
         Begin VTOcx.cboVISUAL cboImpressora 
            Height          =   315
            Left            =   330
            TabIndex        =   13
            Top             =   150
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   556
            Caption         =   "Impressora"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtPagInicial 
            Height          =   315
            Left            =   180
            TabIndex        =   1
            Top             =   555
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            Caption         =   "Página Inicial"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtPagFinal 
            Height          =   315
            Left            =   2295
            TabIndex        =   2
            Top             =   540
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            Caption         =   "Página Final"
            Text            =   ""
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   1185
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   2090
         _Version        =   131082
         TabGuid         =   "ImprimirNovo.frx":1636
         Begin VTOcx.cboVISUAL cboPapel 
            Height          =   315
            Left            =   510
            TabIndex        =   3
            Top             =   30
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   556
            Caption         =   "Papel"
         End
         Begin VTOcx.cboVISUAL cboOrientacao 
            Height          =   315
            Left            =   90
            TabIndex        =   4
            Top             =   390
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   556
            Caption         =   "Tipo impressão"
         End
         Begin VTOcx.txtVISUAL txtTamFonte 
            Height          =   315
            Left            =   1710
            TabIndex        =   5
            Top             =   750
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            Caption         =   "Tamanho Fonte"
            Text            =   ""
         End
      End
   End
   Begin VTOcx.txtVISUAL txtTitulo 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   720
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      Caption         =   "Titulo"
      Text            =   ""
   End
   Begin Cabecalho.rodVISUAL rodRodape 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   2835
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   873
      CorFundo        =   12632256
      CorFrente       =   4210752
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   405
         Left            =   2505
         TabIndex        =   6
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "Imprimir"
         Acao            =   4
         CorFrente       =   4210752
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   3690
         TabIndex        =   7
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   714
         Caption         =   "Sair"
         Acao            =   7
         CorFrente       =   4210752
      End
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
    
    Dim Impressora As Printer
    For Each Impressora In Printers
        If cboImpressora = Impressora.DeviceName Then
            Set Printer = Impressora
            Exit For
        End If
    Next

    OpcoesImpressao = txtTitulo & "||" & cboPapel.ListIndex & "||" & txtPagInicial & "||" & txtPagFinal & "||" & txtTamFonte & "||" & cboOrientacao.ListIndex
    Unload Me
    DoEvents
End Sub

Private Sub cmdSair_Click()
    OpcoesImpressao = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim Util As New VSUtil
    Dim AchouImpressora As Boolean
    
    txtTitulo = Util.ParseString(Tag, "||", 1)
    cboPapel.ListIndex = Util.ParseString(Tag, "||", 2)
    txtPagInicial = Util.ParseString(Tag, "||", 3)
    txtPagFinal = Util.ParseString(Tag, "||", 4)
    txtTamFonte = Util.ParseString(Tag, "||", 5)
    cboOrientacao.ListIndex = Util.ParseString(Tag, "||", 6) - 1
    
    Dim Impressora As Printer
    
    For Each Impressora In Printers
       cboImpressora.AddItem Impressora.DeviceName
       AchouImpressora = True
    Next
    If AchouImpressora Then
       cboImpressora = Printer.DeviceName
    Else
        Avisa "Impressora não encontrada."
    End If
    
End Sub

Private Sub Form_Load()
    cboPapel.AddItem "A4"
    cboPapel.AddItem "Matricial"
    cboOrientacao.AddItem "Vertical"
    cboOrientacao.AddItem "Horizontal"
End Sub
