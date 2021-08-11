VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form GMEN104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GMEN104"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   4875
      TabIndex        =   15
      Top             =   5355
      Width           =   1185
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   315
      Left            =   6120
      TabIndex        =   14
      Top             =   5340
      Width           =   1185
   End
   Begin ActiveTabs.SSActiveTabs TadDados 
      Height          =   5130
      Left            =   45
      TabIndex        =   0
      Top             =   105
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   9049
      _Version        =   131082
      TabCount        =   1
      Tabs            =   "GMEN104.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4740
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   8361
         _Version        =   131082
         TabGuid         =   "GMEN104.frx":0048
         Begin VB.CommandButton CmdArquivo 
            Caption         =   "..."
            Height          =   315
            Left            =   6750
            TabIndex        =   13
            Top             =   3690
            Width           =   285
         End
         Begin VB.TextBox txtArquivo 
            Height          =   285
            Left            =   1080
            TabIndex        =   12
            Top             =   3705
            Width           =   5610
         End
         Begin VB.TextBox txtNomeExibicao 
            Height          =   285
            Left            =   1080
            TabIndex        =   10
            Top             =   3060
            Width           =   5610
         End
         Begin VB.CheckBox ChkAlertaReceberArquivo 
            Caption         =   "Exibir alerta ao recerber arquivos"
            Height          =   195
            Left            =   1425
            TabIndex        =   7
            Top             =   2160
            Width           =   5160
         End
         Begin VB.CheckBox ChkAlertaReceberMensagem 
            Caption         =   "Exibir alerta ao receber mensagem"
            Height          =   195
            Left            =   1425
            TabIndex        =   6
            Top             =   1860
            Width           =   4755
         End
         Begin VB.CheckBox ChkAlertaContatosOline 
            Caption         =   "Exibir alertas quando os contatos ficarem oline"
            Height          =   195
            Left            =   1425
            TabIndex        =   5
            Top             =   1530
            Width           =   5265
         End
         Begin VB.CheckBox ChkExectarGestor 
            Caption         =   "Executar o Gestor Mensseger automaticamente ao fazer logon no Gestor"
            Height          =   195
            Left            =   1425
            TabIndex        =   3
            Top             =   840
            Width           =   5865
         End
         Begin VB.Label Label5 
            Caption         =   "Colocar arquivos recebidos de outros nesta pasta"
            Height          =   165
            Left            =   1050
            TabIndex        =   11
            Top             =   3465
            Width           =   3195
         End
         Begin VB.Label Label4 
            Caption         =   "Nome de Exibição"
            Height          =   165
            Left            =   1065
            TabIndex        =   9
            Top             =   2850
            Width           =   1380
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000005&
            X1              =   2730
            X2              =   7035
            Y1              =   2655
            Y2              =   2655
         End
         Begin VB.Line Line5 
            X1              =   2730
            X2              =   7035
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label3 
            Caption         =   "Mensagem"
            Height          =   225
            Left            =   1065
            TabIndex        =   8
            Top             =   2520
            Width           =   930
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000005&
            X1              =   2715
            X2              =   7020
            Y1              =   1335
            Y2              =   1335
         End
         Begin VB.Line Line3 
            X1              =   2715
            X2              =   7020
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label Label2 
            Caption         =   "Aleras"
            Height          =   225
            Left            =   1215
            TabIndex        =   4
            Top             =   1260
            Width           =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   2730
            X2              =   7035
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Line Line1 
            X1              =   2730
            X2              =   7035
            Y1              =   495
            Y2              =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Entrar No Gestor Mensseger"
            Height          =   225
            Left            =   480
            TabIndex        =   2
            Top             =   390
            Width           =   2085
         End
      End
   End
End
Attribute VB_Name = "GMEN104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSair_Click()
    Unload Me
End Sub
