VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form PATS101 
   BackColor       =   &H00DDF1FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PATS101.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1950
      TabIndex        =   3
      Top             =   1560
      Width           =   2610
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   1349
      CorFundo        =   14545407
      CorFrente       =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Local do arquivo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   675
      TabIndex        =   4
      Top             =   1605
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Insira o disco fornecido contendo  a atualização do sistema e clique em Atualizar."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   0
      Left            =   300
      TabIndex        =   2
      Top             =   900
      Width           =   4815
      WordWrap        =   -1  'True
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4170
      TabIndex        =   1
      ToolTipText     =   "Deseja sair?"
      Top             =   2115
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PATS101.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PATS101.frx":08E6
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdAtualizar 
      Height          =   435
      Left            =   3030
      TabIndex        =   0
      ToolTipText     =   "Atualizar sistema"
      Top             =   2115
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PATS101.frx":0902
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PATS101.frx":091E
      Caption         =   "&Atualizar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
End
Attribute VB_Name = "PATS101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAtualizar_Click()
    On Error GoTo Trata
    
    Screen.MousePointer = 11
    If Not Instala.AchouAtualizacao(Bdados, Mid(Drive1.Drive, 1, 2)) Then
        Util.Avisa "Insira um disquete correto."
    Else
        Util.Informa "Atualização Completada."
    End If
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Trata
    
    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

