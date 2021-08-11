VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU103"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodBairro 
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
      Left            =   0
      MaxLength       =   2
      TabIndex        =   22
      Tag             =   "Distrito"
      Top             =   -450
      Width           =   315
   End
   Begin VB.TextBox txtCodLogr 
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
      Left            =   0
      MaxLength       =   11
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   -570
      Width           =   1035
   End
   Begin Threed.SSFrame fra 
      Height          =   1545
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   660
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   2725
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
      Caption         =   "Referência Cadastral / Localização do Imóvel"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtCodReduzido 
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
         Left            =   1500
         TabIndex        =   23
         Top             =   330
         Width           =   1965
      End
      Begin VB.TextBox txtValorTerr 
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
         Left            =   5940
         TabIndex        =   8
         Top             =   1110
         Width           =   2055
      End
      Begin VB.TextBox txtValorEdif 
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
         Left            =   2550
         TabIndex        =   7
         Top             =   1110
         Width           =   1425
      End
      Begin VB.TextBox txtIcAnterior 
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
         Left            =   6420
         TabIndex        =   0
         Top             =   345
         Width           =   1545
      End
      Begin VB.TextBox txtTipoLogrBt 
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
         Left            =   1500
         MaxLength       =   11
         TabIndex        =   2
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox txtLogrBt 
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
         Left            =   2610
         TabIndex        =   3
         Tag             =   "Nome Contribuinte"
         Top             =   720
         Width           =   2565
      End
      Begin VB.TextBox txtBairroBt 
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
         Left            =   5940
         TabIndex        =   4
         Tag             =   "Nome Contribuinte"
         Top             =   720
         Width           =   2025
      End
      Begin VB.ComboBox cboTipoImovel 
         DataField       =   "ttl_nome"
         DataSource      =   "dtTipLogr"
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
         Height          =   330
         ItemData        =   "TCIU103.frx":0000
         Left            =   9600
         List            =   "TCIU103.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Tipo Imovel"
         Top             =   330
         Width           =   1455
      End
      Begin VB.TextBox txtNumero 
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
         Left            =   8310
         MaxLength       =   10
         TabIndex        =   5
         Top             =   720
         Width           =   555
      End
      Begin VB.TextBox txtComplemento 
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
         Left            =   9600
         TabIndex        =   6
         Top             =   720
         Width           =   1425
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   8910
         TabIndex        =   13
         Top             =   765
         Width           =   645
         _ExtentX        =   1138
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
         Caption         =   "Compl.:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   1
         Left            =   8040
         TabIndex        =   14
         Top             =   780
         Width           =   390
         _ExtentX        =   688
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
         Caption         =   "N.º:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   5310
         TabIndex        =   15
         Top             =   795
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   476
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
         Caption         =   "Bairro:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   7
         Left            =   9150
         TabIndex        =   16
         Top             =   360
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   476
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
         Caption         =   "Tipo:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   72
         Left            =   5160
         TabIndex        =   17
         Top             =   405
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "Insc. Anterior:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   1155
         Width           =   1860
         _ExtentX        =   3281
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
         Caption         =   "Valor Venal Edificação:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   4170
         TabIndex        =   19
         Top             =   1155
         Width           =   1725
         _ExtentX        =   3043
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
         Caption         =   "Valor Venal Terreno:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   82
         Left            =   420
         TabIndex        =   20
         Top             =   765
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   370
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
         Caption         =   "Logradouro:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   24
         Top             =   375
         Width           =   1425
         _ExtentX        =   2514
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
         Caption         =   "Cad. Imobiliário:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   9960
      TabIndex        =   11
      Top             =   2280
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   375
      Left            =   8805
      TabIndex        =   9
      Top             =   2280
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdLimpar 
      Height          =   375
      Left            =   7650
      TabIndex        =   10
      Top             =   2280
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Limpar  "
      Acao            =   6
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   1138
      Icone           =   "TCIU103.frx":0024
   End
End
Attribute VB_Name = "TCIU103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLimpar_Click()
    LimpaCampos Me
    InscricaoAntiga = ""
    InscricaoCadastral = ""
    txtCodReduzido.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Lote As New BCI
    Dim Ic As String
    
    Ic = Trim(txtCodReduzido)
    If Lote.AtualizaValoresMercado(Ic, txtValorTerr, txtValorEdif) Then
        cmdLimpar_Click
        Informa "Dados gravados com sucesso."
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub txtCodLogr_LostFocus()
    Dim Query As String
    Dim rs As VSRecordset
    If Trim(txtCodLogr) = "" Then Exit Sub
    Query = "SELECT TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome, " & _
        " TAB_BAIRRO.TBA_NOME FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
        " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
         " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & txtCodLogr & "'"
    If Bdados.AbreTabela(Query, rs) Then
        txtTipoLogrBt = rs(0)
        txtLogrBt = rs(1)
    Else
        Avisa "Código de logradouro inválido."
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtCodReduzido_Validate(Cancel As Boolean)
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Tem As String
    
    If txtCodReduzido = "" Then Exit Sub
    
    Sql = "Select * from tab_imovel where tim_ic ='" & Trim(txtCodReduzido) & "'"
    
    If Bdados.AbreTabela(Sql, rs) Then
        txtIcAnterior = "" & IIf(rs!tim_ic_anterior = 0, "", rs!tim_ic_anterior)
        cboTipoImovel.ListIndex = rs!tim_tipo_imovel - 1
        txtCodLogr = "" & rs!tim_tlg_cod_logradouro
        DoEvents
        txtCodLogr_LostFocus
        txtNumero = "" & rs!tim_numero
        txtComplemento = "" & rs!tim_complemento
        txtLoteamento = "" & rs!tim_loteamento
        txtQuadra = "" & rs!tim_QUADRA
        txtLote = "" & rs!tim_lote
        txtCodReduzido = "" & rs!TIM_IC
        txtIcAnterior = "" & rs!tim_ic_anterior
        txtCodBairro = "" & rs!tim_TBA_COD_BAIRRO
        txtCodBairro_LostFocus
        txtValorEdif = Format(IIf(IsNull(rs!tim_VALOR_EDIFICACAO_MERCADO) Or CDbl(Nvl("" & rs!tim_VALOR_EDIFICACAO_MERCADO, 0)) = 0, "" & rs!tim_VALOR_EDIFIC, "" & rs!tim_VALOR_EDIFICACAO_MERCADO), Const_Monetario)
        txtValorTerr = Format(IIf(IsNull(rs!tim_VALOR_TERRENO_MERCADO) Or CDbl(Nvl("" & rs!tim_VALOR_TERRENO_MERCADO, 0)) = 0, "" & rs!tim_VALOR_TERRENO, "" & rs!tim_VALOR_TERRENO_MERCADO), Const_Monetario)
        
        Screen.MousePointer = 0
    Else
        Avisa "Lote não encontrado."
    End If
End Sub

Private Sub txtCodBairro_LostFocus()
    Dim rs As VSRecordset
    Dim Sql As String
    If Trim(txtCodBairro) <> "" Then
        Sql = " select TBA_NOME from TAB_BAIRRO where tba_cod_bairro=" & txtCodBairro
        If Bdados.AbreTabela(Sql, rs) Then
            txtBairroBt = rs(0)
        Else
            Avisa "Bairro inexistente."
        End If
        Bdados.FechaTabela rs
    End If
    
End Sub

Private Sub txtValorEdif_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValorEdif_LostFocus()
    txtValorEdif = Format(txtValorEdif, Const_Monetario)
End Sub

Private Sub txtValorTerr_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValorTerr_LostFocus()
    txtValorTerr = Format(txtValorTerr, Const_Monetario)
End Sub
