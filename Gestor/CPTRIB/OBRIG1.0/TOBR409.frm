VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TOBR409 
   Caption         =   "TOBR409"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   1138
      Formulario      =   "Liberação de Contribuinte Notificado."
      Icone           =   "TOBR409.frx":0000
   End
   Begin VTOcx.txtVISUAL txtSenha 
      Height          =   405
      Left            =   3810
      TabIndex        =   0
      Top             =   2730
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   714
      Caption         =   ""
      Text            =   ""
      CaracterSenha   =   "*"
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdOK 
      Height          =   330
      Left            =   3810
      TabIndex        =   1
      Top             =   3270
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   582
      Caption         =   "Liberar"
      Acao            =   3
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   3645
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   688
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   330
      Left            =   4965
      TabIndex        =   2
      Top             =   3270
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   582
      Caption         =   "Cancelar"
      Acao            =   2
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1005
      X2              =   6330
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   1005
      X2              =   6330
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Label LblAviso 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1005
      TabIndex        =   5
      Top             =   1005
      Width           =   5250
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha de Liberacao"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1890
      TabIndex        =   4
      Top             =   2820
      Width           =   1785
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   345
      Picture         =   "TOBR409.frx":031A
      Top             =   1185
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   30
      Picture         =   "TOBR409.frx":0624
      Top             =   825
      Width           =   480
   End
End
Attribute VB_Name = "TOBR409"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    SenhaLiberacao = txtSenha
    Unload Me
End Sub

Private Sub cmdSair_Click()
    SenhaLiberacao = ""
    Unload Me
End Sub

Private Sub Form_Load()
    LblAviso = Temp.PegaParametro(Bdados, "AVISO NOTIFICADO")
End Sub
