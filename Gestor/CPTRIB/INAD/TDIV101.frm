VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TDIV101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TDIV101.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   2640
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1085
      Modulo          =   "Divida Ativa"
      Begin VTOcx.cmdVISUAL cmdDivida 
         Height          =   375
         Left            =   5820
         TabIndex        =   5
         Top             =   150
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         Caption         =   "&Gerar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   150
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8550
         TabIndex        =   7
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   1905
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   660
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   3360
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
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtEndereco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1350
         Width           =   8145
      End
      Begin VB.TextBox txtNome 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3480
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   960
         Width           =   6105
      End
      Begin VB.TextBox txtIm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   2
         Top             =   960
         Width           =   1665
      End
      Begin VB.TextBox txtExercicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1470
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "Exercício"
         Top             =   210
         Width           =   1185
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   1
         Left            =   30
         TabIndex        =   11
         Top             =   210
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
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
         Caption         =   "Exercício"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   960
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
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
         Caption         =   "Contribuinte"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Tag             =   "Tributo"
         Top             =   570
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1138
      Icone           =   "TDIV101.frx":2123
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   9
      Top             =   660
      Width           =   375
   End
End
Attribute VB_Name = "TDIV101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto

Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    txtExercicio.SetFocus
End Sub
Private Sub cmdDivida_Click()
    Dim DAT As New cDividaAtiva
    Dim Sql As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
        
    If txtNome.Text = "" Then
        Util.Avisa "Informe uma inscrição válida"
        txtIm.SetFocus
        Exit Sub
    End If
    
    Sql = "SELECT * FROM VIS_OBRIGACAO_ATRASO WHERE IM = " & txtIm.Text
    If Not Bdados.AbreTabela(Sql) Then
        Util.Avisa "Este contribuinte não está inadiplente"
        txtIm.SetFocus
        Bdados.FechaTabela
        Exit Sub
    End If
    Bdados.FechaTabela
    
    If Not Util.Confirma("Deseja inscrever o contribuinte em dívida ativa") Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    If DAT.GeraDivida(txtExercicio, txtIm, CStr(cboImposto.Coluna(0).Valor)) Then
        Informa "Dívida Ativa Constituída."
    Else
        Util.Avisa "Não existem obrigações para o contribuinte : " & vbCrLf & txtNome
    End If
'    cmdImprime_Click
    Screen.MousePointer = 0
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdsair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Obrig.PreencheComboTributo cboImposto, False
End Sub


Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtInicio_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValidade_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_LostFocus()
'    CarregaEnderecoImovel txtIc, txtEndereco
End Sub

Private Sub txtIm_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    If Trim(txtIm) <> "" Then
        txtIm = BuscaContribuinte(txtIm, txtNome, txtEndereco)
        If Trim(txtIm) = "" Then
            Avisa "Contribuinte não encontrado."
            txtIm.SetFocus
        End If
    End If
End Sub

