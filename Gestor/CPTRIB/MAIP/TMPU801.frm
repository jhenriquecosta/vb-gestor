VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TMPU801 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fra 
      Height          =   1545
      Index           =   0
      Left            =   45
      TabIndex        =   13
      Top             =   690
      Width           =   9225
      _ExtentX        =   16272
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
      Alignment       =   2
      ShadowStyle     =   1
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
         Left            =   4110
         MaxLength       =   5
         TabIndex        =   5
         Tag             =   "Lote"
         Top             =   1065
         Width           =   885
      End
      Begin VB.TextBox txtZona 
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
         Height          =   330
         Left            =   3405
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1050
         Width           =   465
      End
      Begin VB.TextBox txtLogrInicial 
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
         Left            =   6075
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Lote"
         Top             =   1065
         Width           =   885
      End
      Begin VB.TextBox txtLogrFinal 
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
         Left            =   6990
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "Unidade"
         Top             =   1065
         Width           =   885
      End
      Begin VB.TextBox txtDistrito 
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
         Left            =   1875
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "Distrito"
         Top             =   1065
         Width           =   375
      End
      Begin VB.TextBox txtSetor 
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
         Left            =   2295
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "Setor"
         Top             =   1065
         Width           =   375
      End
      Begin VB.TextBox txtQuadra 
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
         Left            =   2700
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "Quadra"
         Top             =   1065
         Width           =   675
      End
      Begin VB.ComboBox cboLado 
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
         ItemData        =   "TMPU801.frx":0000
         Left            =   150
         List            =   "TMPU801.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Lado"
         Top             =   1050
         Width           =   1485
      End
      Begin VB.TextBox txtCodLogr 
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
         Left            =   150
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "Valor"
         Top             =   390
         Width           =   1515
      End
      Begin VB.TextBox txtNumTrecho 
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
         Left            =   7905
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Valor"
         Top             =   1065
         Width           =   1035
      End
      Begin VB.ComboBox cboTipoLogr 
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
         ItemData        =   "TMPU801.frx":0004
         Left            =   1875
         List            =   "TMPU801.frx":0011
         Locked          =   -1  'True
         TabIndex        =   11
         Tag             =   "Tipo Logradouro"
         Text            =   "cboTipoLogr"
         Top             =   390
         Width           =   1665
      End
      Begin VB.ComboBox cboLogr 
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
         ItemData        =   "TMPU801.frx":0032
         Left            =   3585
         List            =   "TMPU801.frx":003F
         Locked          =   -1  'True
         TabIndex        =   12
         Tag             =   "Nome Logradouro"
         Text            =   "cboLogr"
         Top             =   390
         Width           =   2415
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   1
         Left            =   7905
         TabIndex        =   14
         Top             =   840
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         _Version        =   196610
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
         Caption         =   "Nº do Trecho"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   150
         TabIndex        =   15
         Top             =   810
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   476
         _Version        =   196610
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
         Caption         =   "Lado"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   2
         Left            =   150
         TabIndex        =   16
         Top             =   150
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   476
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   "Código Logradouro"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   5
         Left            =   1875
         TabIndex        =   17
         Top             =   810
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   476
         _Version        =   196610
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
         Caption         =   "Dist."
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   7
         Left            =   2295
         TabIndex        =   18
         Top             =   840
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   370
         _Version        =   196610
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
         Caption         =   "Set."
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   8
         Left            =   2700
         TabIndex        =   19
         Top             =   840
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   370
         _Version        =   196610
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
         Caption         =   "Qd."
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   9
         Left            =   6075
         TabIndex        =   21
         Top             =   840
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   370
         _Version        =   196610
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
         Caption         =   "Logr Inicial"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   10
         Left            =   6990
         TabIndex        =   22
         Top             =   840
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   370
         _Version        =   196610
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
         Caption         =   "Logr Final"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Command1"
         Default         =   -1  'True
         Height          =   255
         Left            =   3090
         TabIndex        =   23
         Top             =   2550
         Width           =   375
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   28
         Left            =   3405
         TabIndex        =   24
         Top             =   840
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   370
         _Version        =   196610
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
         Caption         =   "Zona"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   0
         Left            =   4110
         TabIndex        =   25
         Top             =   840
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   370
         _Version        =   196610
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
         Caption         =   "Cod Bairro"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin VTOcx.grdVISUAL lstBvt 
      Height          =   2340
      Left            =   60
      TabIndex        =   26
      Top             =   2295
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   4128
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   8130
      TabIndex        =   10
      Top             =   4680
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
      Left            =   6945
      TabIndex        =   9
      Top             =   4680
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   1138
      Icone           =   "TMPU801.frx":0060
   End
End
Attribute VB_Name = "TMPU801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cadastro As VSImposto
Dim Click As Boolean

Private Sub AtualizaLista(Logr As String)
    Dim Sql As String
    Dim rs As VSRecordset
    If Bdados.Conexao.FormatoBanco = SQLServer Then
    Sql = "SELECT  Tab_Geral.TGE_NOME as Lado,  " & _
        " Tab_Valor_Terreno.tvl_distrito as [Distrito],Tab_Valor_Terreno.tvl_setor as Setor," & _
            " Tab_Valor_Terreno.tvl_quadra AS Quadra,Tab_Valor_Terreno.tvl_logr_inicial as [Logr Inicio], " & _
            "Tab_Valor_Terreno.tvl_logr_final as [Logr Fim], Tab_Valor_Terreno.tvl_num_trecho as [Num Trecho],Tab_Valor_Terreno.tvl_zona as [Zona], tvl_tba_cod_bairro as [Bairro] " & _
            " FROM TAB_BAIRRO,TAB_TIPO_LOGR," & _
        " TAB_LOGRADOURO,Tab_Geral,Tab_Valor_Terreno WHERE " & _
        " Tab_Valor_Terreno.tvl_tlg_cod_logradouro = TAB_LOGRADOURO.tlg_cod_logradouro AND " & _
        " TAB_TIPO_LOGR.TTL_COD_TIP_LOGR = TAB_LOGRADOURO.tlg_ttl_cod_tip_logr AND " & _
        " TAB_BAIRRO.TBA_COD_BAIRRO = TAB_LOGRADOURO.tlg_tba_cod_bairro " & _
        " AND Tab_Valor_Terreno.tvl_lado =Tab_Geral.TGE_CODIGO " & _
         " AND Tab_Geral.TGE_CODIGO > 0 and Tab_Geral.TGE_TIPO =8 and TAB_LOGRADOURO.tlg_cod_logradouro='" & Logr & _
        "' and tlg_tmu_cod_municipio = " & Aplicacoes.Codigo_Municipio & " and TBA_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio & _
        " ORDER BY TAB_BAIRRO.TBA_NOME, TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome "
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "SELECT  Tab_Geral.TGE_NOME as Lado,  " & _
        " Tab_Valor_Terreno.tvl_distrito as Distrito,Tab_Valor_Terreno.tvl_setor as Setor," & _
            " Tab_Valor_Terreno.tvl_quadra AS Quadra,Tab_Valor_Terreno.tvl_logr_inicial as Logr_Inicio, " & _
            "Tab_Valor_Terreno.tvl_logr_final as Logr_Fim, Tab_Valor_Terreno.tvl_num_trecho as Num_Trecho,Tab_Valor_Terreno.tvl_zona as Zona, tvl_tba_cod_bairro as Bairro " & _
            " FROM TAB_BAIRRO,TAB_TIPO_LOGR," & _
        " TAB_LOGRADOURO,Tab_Geral,Tab_Valor_Terreno WHERE " & _
        " Tab_Valor_Terreno.tvl_tlg_cod_logradouro = TAB_LOGRADOURO.tlg_cod_logradouro AND " & _
        " TAB_TIPO_LOGR.TTL_COD_TIP_LOGR = TAB_LOGRADOURO.tlg_ttl_cod_tip_logr AND " & _
        " TAB_BAIRRO.TBA_COD_BAIRRO = TAB_LOGRADOURO.tlg_tba_cod_bairro " & _
        " AND Tab_Valor_Terreno.tvl_lado =Tab_Geral.TGE_CODIGO " & _
         " AND Tab_Geral.TGE_CODIGO > 0 and Tab_Geral.TGE_TIPO =8 and TAB_LOGRADOURO.tlg_cod_logradouro='" & Logr & _
        "'  ORDER BY TAB_BAIRRO.TBA_NOME, TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome "
    End If
    lstBvt.ListItems.Clear
    lstBvt.Preencher Bdados, Sql, 1400
End Sub

Private Sub cboLado_Click()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Contador As Double
    
    Sql = "Select COUNT(tvl_tlg_cod_logradouro) + 1 from tab_valor_terreno where tvl_tlg_cod_logradouro='" & txtCodLogr & "' and tvl_lado = " & cboLado.ListIndex + 1
    If Bdados.AbreTabela(Sql, rs) Then
        Contador = rs(0)
        txtNumTrecho = "" & Contador & Left(cboLado, 1)
    Else
        txtNumTrecho = "1" & Left(cboLado, 1)
    End If
    Bdados.FechaTabela rs
    Sql = "Select tvl_logr_final from tab_valor_terreno where tvl_tlg_cod_logradouro='" & txtCodLogr & "' and tvl_lado = " & cboLado.ListIndex + 1 & " AND tvl_num_trecho ='" & Contador & Left(cboLado, 1) & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        If IsNull(rs(0)) Then
            txtLogrInicial.Enabled = True
            txtLogrInicial = ""
        Else
            txtLogrInicial = rs(0)
            txtLogrInicial.Enabled = False
        End If
    Else
        txtLogrInicial.Enabled = True
        txtLogrInicial = ""
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub cboTipoLogr_Click()
    Call cadastro.BuscaLogradouro(Rua, cboTipoLogr, cboLogr)
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim CodLogradouro As Long
    Dim Valores As String
    Dim Campos As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    CodLogradouro = txtCodLogr
    Campos = "tvl_tlg_cod_logradouro,tvl_lado,tvl_distrito,tvl_setor,tvl_quadra,tvl_logr_inicial,tvl_logr_final,tvl_num_trecho,tVL_tba_cod_bairro,tVL_zona"
    Valores = Bdados.PreparaValor(CodLogradouro, cboLado.ListIndex + 1, txtDistrito, txtSetor, txtQuadra, txtLogrInicial, txtLogrFinal, txtNumTrecho, txtCodBairro, Nvl(txtZona, 1))
    Call Bdados.GravaDados("Tab_Valor_Terreno", Valores, Campos, "tvl_tlg_cod_logradouro='" & CodLogradouro & "' and tvl_lado=" & cboLado.ListIndex + 1 & " and tvl_num_trecho='" & txtNumTrecho & "'")
    AtualizaLista txtCodLogr
    Informa "Transação completada."
    cboLado_Click
    txtLogrFinal = ""
    txtLogrInicial.SetFocus
End Sub

Private Sub Form_Activate()
    Call Edita.AtualizaCombo(Bdados, cboLogr, "Select tlg_nome From Tab_Logradouro where tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio)
    Call Edita.AtualizaCombo(Bdados, cboTipoLogr, "Select ttl_nome From Tab_Tipo_Logr")
    Call Edita.AtualizaCombo(Bdados, cboLado, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_CODIGO < 3 and TGE_TIPO =8 ORDER BY TGE_CODIGO ASC")
    Set cadastro = New VSImposto
    AtualizaCabecalho lstBvt
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    
End Sub

Private Sub lstBvt_Click()
    On Error Resume Next
    Click = True
    cboLado = lstBvt.SelectedItem
    txtDistrito = lstBvt.SelectedItem.SubItems(1)
    txtSetor = lstBvt.SelectedItem.SubItems(2)
    txtQuadra = lstBvt.SelectedItem.SubItems(3)
    txtLogrInicial = lstBvt.SelectedItem.SubItems(4)
    txtLogrFinal = lstBvt.SelectedItem.SubItems(5)
    txtNumTrecho = lstBvt.SelectedItem.SubItems(6)
    txtZona = lstBvt.SelectedItem.SubItems(7)
    txtCodBairro = lstBvt.SelectedItem.SubItems(8)
End Sub

Private Sub lstBvt_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstBvt, ColumnHeader
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub lstBvt_DblClick()
    If Confirma("Deseja excluir o trecho " & txtNumTrecho & " da " & cboTipoLogr & "  " & cboLogr & ".") Then
        Bdados.DeletaDados "Tab_Valor_Terreno", "tvl_tlg_cod_logradouro='" & txtCodLogr & "' and tvl_lado=" & cboLado.ListIndex + 1 & " and tvl_num_trecho='" & txtNumTrecho & "'"
    End If
    Avisa "Trecho excluído."
    AtualizaLista txtCodLogr
End Sub

Private Sub txtCodLogr_KeyPress(KeyAscii As Integer)
    Click = False
End Sub

Private Sub txtCodLogr_LostFocus()
    Dim rs As VSRecordset
    Dim Sql As String
    
    If Trim(txtCodLogr) <> "" And Not Click Then
        Sql = "Select TTL_NOME,tlg_nome from tab_logradouro, tab_tipo_logr where tlg_cod_logradouro='" & txtCodLogr & "' and tlg_ttl_cod_tip_logr = TTL_COD_TIP_LOGR and tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
        If Bdados.AbreTabela(Sql, rs) Then
            cboTipoLogr = rs(0)
            cboLogr = rs(1)
            AtualizaLista txtCodLogr
        Else
            Avisa "Código de logradouro inexistente."
        End If
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtDistrito_Change()
    If Len(txtDistrito) = txtDistrito.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDistrito_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtQuadra_Change()
     If Len(txtQuadra) = txtQuadra.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtQuadra_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtSetor_Change()
    If Len(txtSetor) = txtSetor.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSetor_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub
