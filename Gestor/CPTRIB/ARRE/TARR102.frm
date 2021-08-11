VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TARR102 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Gestão de Receitas"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPerFinal 
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   3330
      MaxLength       =   4
      TabIndex        =   7
      Tag             =   "Periodo Final"
      Top             =   1110
      Width           =   1395
   End
   Begin VB.TextBox txtPerInicial 
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Periodo Inicial"
      Top             =   1125
      Width           =   1395
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   3000
      Left            =   60
      TabIndex        =   1
      Top             =   1950
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   5292
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5730
      TabIndex        =   10
      Top             =   1170
      Width           =   855
   End
   Begin VB.Label lblConta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7410
      TabIndex        =   9
      Top             =   1110
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fim"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2880
      TabIndex        =   8
      Top             =   1185
      Width           =   300
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Período de  Arrecadação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   810
      Width           =   7635
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   390
      TabIndex        =   5
      Top             =   1200
      Width           =   465
   End
   Begin Threed.SSCommand cmdSair 
      Height          =   435
      Left            =   6675
      TabIndex        =   3
      ToolTipText     =   "Deseja sair?"
      Top             =   5025
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdGerar 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   5580
      TabIndex        =   2
      Top             =   5025
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Gerar"
      ButtonStyle     =   4
      PictureAlignment=   6
      ShapeSize       =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Imóveis Importados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   1665
      Width           =   7665
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "TARR102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NomeArq As String = "VTExpIPTU.arr"
Private Sub cmdGerar_Click()
    Dim Cobranca As New VSCobranca
    Dim RsZona As VSRecordset
    Dim RsImovel As VSRecordset
    Dim sql As String
    
    'Call Cobranca.GeraArquivoIPTU(txtPerInicial, App.Path & "\" & NomeArq)
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid Grid, ColumnHeader
End Sub


Private Sub txtBanco_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    On Error GoTo trata

    Screen.MousePointer = 0
    Exit Sub
trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

