VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCTB101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCTB101.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   1815
      Index           =   2
      Left            =   30
      TabIndex        =   9
      Top             =   690
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   3201
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
      Begin VB.TextBox txtConvenio 
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
         Left            =   5280
         MaxLength       =   14
         TabIndex        =   4
         Tag             =   "Data Pagamento"
         Top             =   960
         Width           =   1635
      End
      Begin VB.TextBox txtDtDesativa 
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
         Left            =   5310
         MaxLength       =   14
         TabIndex        =   6
         Top             =   1350
         Width           =   1635
      End
      Begin VB.TextBox txtDtAtivacao 
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
         Left            =   1980
         MaxLength       =   14
         TabIndex        =   5
         Top             =   1350
         Width           =   1650
      End
      Begin VB.ComboBox cboSituacao 
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
         ItemData        =   "TCTB101.frx":2123
         Left            =   1980
         List            =   "TCTB101.frx":212D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Tipo"
         Top             =   960
         Width           =   1665
      End
      Begin VB.ComboBox cboAgente 
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
         Left            =   1980
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Tipo"
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtCodSucursal 
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
         Left            =   1980
         MaxLength       =   14
         TabIndex        =   1
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox txtNumConta 
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
         Left            =   5280
         MaxLength       =   14
         TabIndex        =   2
         Tag             =   "Data Pagamento"
         Top             =   600
         Width           =   1635
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   4
         Left            =   480
         TabIndex        =   10
         Top             =   660
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   318
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
         Caption         =   "Código Sucursal"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   0
         Left            =   3630
         TabIndex        =   11
         Top             =   630
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   318
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
         Caption         =   "Num. Conta"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   270
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   423
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
         Caption         =   "Agente Arrecadador"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   3
         Left            =   150
         TabIndex        =   14
         Top             =   990
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   423
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
         Caption         =   "Situação Conta"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Top             =   1410
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   318
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
         Caption         =   "Data Ativação"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   5
         Left            =   3660
         TabIndex        =   16
         Top             =   1380
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   318
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
         Caption         =   "Data Desativação"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   6
         Left            =   3630
         TabIndex        =   19
         Top             =   990
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   318
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
         Caption         =   "No. Convênio"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   870
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   1138
      Icone           =   "TCTB101.frx":2141
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   5970
      TabIndex        =   8
      Top             =   2550
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   375
      Left            =   4740
      TabIndex        =   7
      Top             =   2550
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCTB101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim NumAgente  As Double
Private Sub cboAgente_Click()
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "Select tar_cod_agente from tab_agente_ARRECADADOR where tar_nome_agente ='" & cboAgente & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        NumAgente = rs(0)
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Condicao As String
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Campos = "tcb_tar_cod_agente,tcb_cod_sucursal,tcb_num_conta,tcb_data_ativacao,tcb_tus_cod_usuario,tcb_convenio,TCB_STATUS"
    Valores = Bdados.PreparaValor(NumAgente, txtCodSucursal, txtNumConta, Bdados.Converte(txtDtAtivacao, TCDataHora), Aplicacoes.Usuario, Nvl(txtConvenio, 0), IIf(cboSituacao.ListIndex = -1, 0, cboSituacao.ListIndex))
    If Trim(txtDtDesativa) <> "" Then
        Valores = Valores = Bdados.PreparaValor(Bdados.Converte(txtDtDesativa, TCDataHora))
        Campos = Campos & ",tcb_data_desativacao"
    End If
    Bdados.GravaDados "TAB_CONTA_BANCARIA", Valores, Campos, _
    "tcb_tar_cod_agente=" & NumAgente & " and tcb_cod_sucursal='" & txtCodSucursal & "' and tcb_num_conta='" & txtNumConta & "'"
    Util.Informa "Transação Realizada com Sucesso."
    Edita.LimpaCampos Me
    cboAgente.SetFocus
End Sub

Private Sub Form_Load()
    Dim rs As VSRecordset
    cboAgente.Clear
    cabVisual.Exibir Bdados, Me.Name, App.Path
    AtualizaCombo Bdados, cboAgente, "Select tar_nome_agente from tab_agente_arrecadador where tar_ativo =0"
    End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDtAtivacao_LostFocus()
    txtDtAtivacao = Edita.FormataTexto(txtDtAtivacao, Data)
End Sub

Private Sub txtDtDesativa_LostFocus()
    txtDtDesativa = Edita.FormataTexto(txtDtDesativa, Data)
End Sub

Private Sub txtNumConta_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    
    Sql = "SELECT tcb_tar_cod_agente,tcb_cod_sucursal,tcb_num_conta,tcb_data_ativacao,tcb_tus_cod_usuario,tcb_convenio,TCB_STATUS,tcb_data_desativacao " & _
        " FROM TAB_CONTA_BANCARIA WHERE tcb_tar_cod_agente =" & NumAgente & " AND tcb_cod_sucursal = '" & _
        txtCodSucursal & "' AND tcb_num_conta ='" & txtNumConta & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        cboSituacao.ListIndex = Nvl("" & rs!TCB_STATUS, 0)
        txtConvenio = "" & rs!tcb_convenio
        txtDtAtivacao = "" & rs!tcb_data_ativacao
        txtDtDesativa = "" & rs!tcb_data_desativacao
    End If
End Sub
