VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Begin VB.Form TMPU202 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   60
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   820
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Alignment       =   6
      RoundedCorners  =   0   'False
      Begin Threed.SSPanel lblForm 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   60
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   661
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   128
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Inclusão"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lblModulo 
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   503
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   128
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ECAL101"
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         AutoSize        =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lblHora 
         Height          =   285
         Left            =   60
         TabIndex        =   5
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   128
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "00:00:00"
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         AutoSize        =   3
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1470
      TabIndex        =   6
      Top             =   150
      Width           =   375
   End
   Begin Threed.SSCommand cmdSalvar 
      Height          =   435
      Left            =   2220
      TabIndex        =   0
      Top             =   630
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TMPU202.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU202.frx":001C
      Caption         =   "&Salvar"
      ButtonStyle     =   3
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdSair 
      Height          =   435
      Left            =   3660
      TabIndex        =   1
      Top             =   630
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TMPU202.frx":0038
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU202.frx":0054
      Caption         =   "Sai&r"
      ButtonStyle     =   3
      PictureAlignment=   6
   End
End
Attribute VB_Name = "TMPU202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Tipologia As Integer
    Dim Estrutura As Integer
    Dim Padrao As Integer
    Dim Valores As String
    Dim Campos As String
    Dim ConsultaPreliminar As String
    Dim ConsultaPredominate As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Campos = "TCU_COD_ITEM,TCU_TCO_COD_COMPONENTE_TIPOLOGIA,TCU_TCO_COD_COMPONENTE_ESTRUTURA,TCU_TCO_COD_COMPONENTE_PADRAO,TCU_VALOR_UNITARIO"
        Informa "Transação completada."
        Edita.LimpaCampos Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
End Sub
