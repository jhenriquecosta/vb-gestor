VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{D8A7CA9C-BFF7-11D5-9D50-00D0590D0C80}#1.0#0"; "cTreeOpt.ocx"
Begin VB.Form THOM405 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "THOM405"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin cTreeOpt.XTreeOpt TrebDados 
      Height          =   4950
      Left            =   45
      TabIndex        =   2
      Top             =   690
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8731
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      Indentation     =   256.251983642578
      Style           =   5
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL CmdSair 
         Height          =   345
         Left            =   8115
         TabIndex        =   3
         Top             =   75
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   609
         Caption         =   "Sair"
         Acao            =   7
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1138
      Icone           =   "THOM405.frx":0000
   End
End
Attribute VB_Name = "THOM405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub
