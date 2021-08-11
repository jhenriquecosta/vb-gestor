VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCOB404 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCOB404.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   16
      Top             =   6255
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6450
         TabIndex        =   7
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdBusca 
         Height          =   375
         Left            =   5250
         TabIndex        =   6
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7650
         TabIndex        =   8
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421104
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1545
      Left            =   60
      TabIndex        =   11
      Top             =   765
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   2725
      Altura          =   1905
      Caption         =   " Opções de Busca"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3900
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1140
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   4755
         TabIndex        =   15
         Top             =   1185
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
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
         Caption         =   "Exercício"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   1185
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
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
         Caption         =   "Inscrição Cadastral"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   675
         TabIndex        =   13
         Top             =   810
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
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
         Caption         =   "Tipo Isenção"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   1125
         TabIndex        =   12
         Top             =   435
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
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
         Caption         =   "Tributo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VB.ComboBox cboImposto 
         DataField       =   "ttl_nome"
         DataSource      =   "dtTipLogr"
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
         Height          =   330
         Left            =   1815
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "GRUPOATIVIDADE"
         Top             =   375
         Width           =   6750
      End
      Begin VB.ComboBox cboIsencao 
         DataField       =   "ttl_nome"
         DataSource      =   "dtTipLogr"
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
         Height          =   330
         ItemData        =   "TCOB404.frx":2123
         Left            =   1815
         List            =   "TCOB404.frx":2136
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "GRUPOATIVIDADE"
         Top             =   757
         Width           =   2865
      End
      Begin VB.TextBox txtIc 
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
         Left            =   1815
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1140
         Width           =   2055
      End
      Begin VB.TextBox txtExercicio1 
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
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1140
         Width           =   1485
      End
      Begin VB.TextBox txtExercicio2 
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
         Left            =   7065
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1140
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   930
      Width           =   375
   End
   Begin VTOcx.grdVISUAL lstAtv 
      Height          =   3795
      Left            =   75
      TabIndex        =   5
      Top             =   2400
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   6694
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   1138
      Icone           =   "TCOB404.frx":21AC
   End
End
Attribute VB_Name = "TCOB404"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBusca_Click()
    Dim Sql As String
    Dim RsPref As VSRecordset
    Dim RsCTM As VSRecordset
    Dim Anterior As String
    Sql = "Select TIS_TCI_IM as Contribuinte, TIS_TIM_IC as Imovel,TIS_PERIODO as Periodo,TIS_VALOR_TRIBUTO as Valor," & _
        "TIS_TIPO_ISENSAO as Tipo, Tip_Sigla_Imposto as Tributo from Tab_Isento, Tab_Imposto where TIS_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
    
    If Trim(cboImposto) <> "" Then Sql = Sql & " and tip_nome_imposto ='" & cboImposto & "'"
    If Trim(cboIsencao) <> "" Then Sql = Sql & " and TIS_TIPO_ISENSAO = " & Left(cboIsencao, 1)
    If Trim(txtIc) <> "" Then Sql = Sql & " and TIS_TIM_IC LIKE '" & Trim(txtIc) & "%'"
    If Trim(txtExercicio1) <> "" And Trim(txtExercicio2) <> "" Then
        Sql = Sql & " and TIS_periodo >= " & IIf(Len(txtExercicio1) = 4, txtExercicio1, Right(txtExercicio1, 4) & Left(txtExercicio1, 2)) & " and TIS_periodo <= " & IIf(Len(txtExercicio2) = 4, txtExercicio2, Right(txtExercicio2, 4) & Left(txtExercicio2, 2))
    End If
    If lstAtv.Preencher(Bdados, Sql, 1300, 2000, 900, 1000, 650, 1000) Then
        lstAtv.Mensagem = "Total de Isenção: R$" & Format(lstAtv.Colunas(4).Soma, Const_Monetario)
    End If
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    lstAtv.Preencher Bdados, ""
    cboImposto.SetFocus
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtIc
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub


Private Sub Form_Activate()
    Dim Sql As String
        
    AtualizaCabecalho lstAtv
    '1 - Isento por Limite da Base
    '2 - Isento de Imposto
    '3 - Isento por Limite Tributo
    '4 - Isento Total
    '5 - Imune

End Sub

Private Sub Form_Load()
    On Error Resume Next
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    Call Edita.AtualizaCombo(Bdados, cboImposto, "Select   tip_nome_imposto From TAB_IMPOSTO")
    cboImposto.AddItem " "
End Sub

Private Sub Timer_Timer()
    On Error Resume Next
    
    
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub Timer1_Timer()

End Sub

Private Sub txtDescAtiv_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtMult_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub


Private Sub txtExercicio1_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtExercicio2_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub
