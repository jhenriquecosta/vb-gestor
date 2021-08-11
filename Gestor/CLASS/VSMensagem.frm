VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VSMensagem 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFD3B3&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mensagem"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6180
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand cmdBotao 
      Cancel          =   -1  'True
      Height          =   435
      Index           =   6
      Left            =   2640
      TabIndex        =   2
      Top             =   1830
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   11683841
      BackColor       =   16762261
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&OK"
      ButtonStyle     =   3
   End
   Begin VB.TextBox txtEntrada 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B24801&
      Height          =   330
      Left            =   885
      MaxLength       =   150
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1035
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox lblMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFD3B3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   915
      Left            =   930
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "VSMensagem.frx":0000
      Top             =   810
      Width           =   4995
   End
   Begin VB.Image imgInforma 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   195
      Picture         =   "VSMensagem.frx":001B
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEntrada 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   195
      Picture         =   "VSMensagem.frx":08E5
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgExclama 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   195
      Picture         =   "VSMensagem.frx":11AF
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMsg 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   195
      Picture         =   "VSMensagem.frx":1A79
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAvisa 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   195
      Picture         =   "VSMensagem.frx":2343
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgErro 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   195
      Picture         =   "VSMensagem.frx":3185
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Título..............."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B24801&
      Height          =   270
      Left            =   915
      TabIndex        =   1
      Top             =   360
      Width           =   5220
   End
End
Attribute VB_Name = "VSMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBotao_Click(Index As Integer)
    On Error GoTo Trata
    T.OpcaoBotao = Index
    T.Resposta = Trim$(txtEntrada)
    Unload Me
Trata:
End Sub

Private Sub Form_Activate()
    cmdBotao(6).Visible = True
    Beep
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub
