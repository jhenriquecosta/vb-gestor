VERSION 5.00
Begin VB.Form sobre_frm 
   BorderStyle     =   0  'None
   Caption         =   "DECLARAÇÕES ELETRÔNICAS"
   ClientHeight    =   3180
   ClientLeft      =   2100
   ClientTop       =   1770
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   90
      Picture         =   "SOBRE.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   885
      TabIndex        =   4
      Top             =   30
      Width           =   885
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   870
      ScaleHeight     =   825
      ScaleWidth      =   3915
      TabIndex        =   6
      Top             =   1380
      Width           =   3975
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Esta é uma versão de Teste, não utilize para apresentar sua Declaração."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   330
         TabIndex        =   7
         Top             =   360
         Width           =   3465
      End
   End
   Begin VB.PictureBox SSPanel2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   0
      Left            =   1020
      ScaleHeight     =   585
      ScaleWidth      =   3645
      TabIndex        =   8
      Top             =   810
      Width           =   3705
   End
   Begin VB.PictureBox SSCommand1 
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4230
      ScaleHeight     =   375
      ScaleWidth      =   1365
      TabIndex        =   5
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   3195
      Left            =   0
      Top             =   0
      Width           =   5745
   End
   Begin VB.Label label_versao 
      Caption         =   "Versão 1.00"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   2610
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"SOBRE.frx":204A
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   210
      TabIndex        =   2
      Top             =   2370
      Width           =   3285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Secretaria de Estado da Fazenda "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   330
      Width           =   4140
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "GOVERNO DO ESTADO DO MARANHÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   60
      Width           =   4935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   195
      X2              =   5185
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   5250
      Y1              =   2340
      Y2              =   2340
   End
End
Attribute VB_Name = "sobre_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   CentrarFormu Me
   label_versao.Caption = "Versão " & Mid(Versao, 2, 1) & "." & Mid(Versao, 3, 1) '
End Sub


Private Sub SSCommand1_Click()
Dim gArquivo As String
Dim reg As String
Dim nFile As Integer
On Error GoTo ErroFile:
nFile = FreeFile(0)
gArquivo = App.Path + "\Dief.ini"
Open gArquivo For Input Shared As nFile
Input #nFile, reg
Close #nFile
If Mid(reg, 9, 5) <> "False" Then
    Unload Me
    Dec010.Show 1
Else
    Unload Me
End If

ErroFile:
If Err = 53 Then
     Unload Me
     Dec010.Show 1
End If

End Sub


