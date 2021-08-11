VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form GMEN103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GMEN103"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand CmdEnviar 
      Height          =   600
      Left            =   7200
      TabIndex        =   2
      Top             =   5010
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1058
      _Version        =   196610
      PictureFrames   =   1
      Picture         =   "GMEN103.frx":0000
   End
   Begin VB.TextBox txtMensagem 
      Appearance      =   0  'Flat
      Height          =   885
      Left            =   225
      TabIndex        =   1
      Top             =   4860
      Width           =   8040
   End
   Begin VB.TextBox txtTexto 
      BorderStyle     =   0  'None
      Height          =   3825
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   540
      Width           =   8010
   End
   Begin VB.Label LblConvidado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Raimundo"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   870
      TabIndex        =   4
      Top             =   270
      Width           =   7200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Para:"
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   450
      TabIndex        =   3
      Top             =   255
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   6225
      Left            =   15
      Picture         =   "GMEN103.frx":0508
      Top             =   0
      Width           =   8445
   End
End
Attribute VB_Name = "GMEN103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ImgImagem_Click()

End Sub
