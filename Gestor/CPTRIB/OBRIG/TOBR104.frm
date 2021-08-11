VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TOBR104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TOBR104-Geração de IPTU"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleMode       =   0  'User
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdValores 
      Height          =   2055
      Left            =   60
      TabIndex        =   37
      Top             =   3750
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   3625
      CorBorda        =   32768
      Caption         =   "Valores"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   0
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   21
      Top             =   6675
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   12
         Top             =   120
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   661
         Caption         =   "Gerar &Cota Única"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   13
         Top             =   120
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         Caption         =   "Gerar &Parcelas"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   4495
         TabIndex        =   15
         Top             =   120
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   6360
         TabIndex        =   16
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2985
      Left            =   75
      TabIndex        =   20
      Top             =   720
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   5265
      Altura          =   1905
      Caption         =   " Detalhes"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtVencimento 
         Height          =   300
         Left            =   1800
         TabIndex        =   39
         Top             =   2235
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   529
         Caption         =   "Vencto"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin Threed.SSPanel lblCont 
         Height          =   285
         Left            =   4320
         TabIndex        =   35
         Top             =   2250
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   196610
         ForeColor       =   8388608
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   9
         Left            =   105
         TabIndex        =   34
         Top             =   2280
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
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
         Caption         =   "Período"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtAnoIni 
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
         Left            =   3915
         MaxLength       =   4
         TabIndex        =   33
         Tag             =   "TOC_PERIODO"
         Top             =   2235
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox txtAnoFi 
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
         Left            =   825
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "TOC_PERIODO"
         Top             =   2235
         Width           =   885
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   11
         Left            =   3060
         TabIndex        =   31
         Top             =   360
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   397
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
         Caption         =   "Cod Logr"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   6
         Left            =   5385
         TabIndex        =   30
         Top             =   960
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
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
         Caption         =   "Tipo Imóvel"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   29
         Top             =   1635
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
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
         Caption         =   "Valor Venal"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   8
         Left            =   1740
         TabIndex        =   28
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
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
         Caption         =   "Insc. Municipal"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   49
         Left            =   105
         TabIndex        =   27
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
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
         Caption         =   "Insc. Cadastral"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   7
         Left            =   3060
         TabIndex        =   26
         Top             =   960
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   397
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
         Caption         =   "Setor"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   5
         Left            =   4020
         TabIndex        =   25
         Top             =   960
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   397
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
         Caption         =   "Quadra"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   24
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
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
         Caption         =   "Bairro"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   4020
         TabIndex        =   23
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
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
         Caption         =   "Logradouro"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtic 
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
         Left            =   105
         TabIndex        =   0
         Tag             =   "TIM_IC"
         Top             =   600
         Width           =   1560
      End
      Begin VB.ComboBox cboBairro 
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
         Left            =   105
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "TBA_NOME"
         Top             =   1207
         Width           =   2910
      End
      Begin VB.ComboBox cboLogr 
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
         Left            =   5385
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "TLG_NOME"
         Top             =   600
         Width           =   1920
      End
      Begin VB.ComboBox cboTipoLogr 
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
         Left            =   4020
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "TTL_NOME"
         Top             =   600
         Width           =   1305
      End
      Begin VB.TextBox txtQuadra 
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
         Left            =   4020
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1215
         Width           =   915
      End
      Begin VB.TextBox txtSecao 
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
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1215
         Width           =   915
      End
      Begin VB.TextBox txtIM 
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
         Left            =   1725
         MaxLength       =   11
         TabIndex        =   1
         Tag             =   "TIM_TCI_IM"
         Top             =   600
         Width           =   1275
      End
      Begin VB.TextBox txtValor01 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   105
         TabIndex        =   9
         Top             =   1860
         Width           =   1110
      End
      Begin VB.TextBox txtValor02 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1785
         TabIndex        =   10
         Top             =   1860
         Width           =   1185
      End
      Begin VB.ComboBox cboTipo 
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
         Left            =   5385
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1207
         Width           =   1920
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
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   2
         Tag             =   "TIM_TLG_COD_LOGRADOURO"
         Top             =   600
         Width           =   915
      End
      Begin Threed.SSPanel lblIsentos 
         Height          =   330
         Left            =   15
         TabIndex        =   36
         Top             =   2640
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   12045006
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         X1              =   0
         X2              =   7380
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "até"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   1350
         TabIndex        =   32
         Top             =   1905
         Width           =   285
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   1138
      Formulario      =   "IPTU"
      Icone           =   "TOBR104.frx":0000
   End
   Begin VB.PictureBox prgIptu 
      Height          =   195
      Left            =   4545
      ScaleHeight     =   135
      ScaleWidth      =   2670
      TabIndex        =   18
      Top             =   2715
      Width           =   2730
   End
   Begin VB.Timer tmr 
      Interval        =   10
      Left            =   2550
      Top             =   1050
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   180
      Width           =   375
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   795
      Left            =   75
      TabIndex        =   22
      Top             =   5820
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   1402
      Altura          =   1905
      Caption         =   " Relatórios"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboRelatorio 
         Height          =   315
         Left            =   90
         TabIndex        =   38
         Top             =   360
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         Caption         =   "Relatório"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VB.ComboBox cboCampoZero 
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
         Left            =   4680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   375
         Width           =   2625
      End
   End
End
Attribute VB_Name = "TOBR104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As VSImposto
Dim Cobranca As New VSCobranca
Dim CalculoIptu As New VSIptu
Dim sqlValores As String
Dim observacaoIPTU As String

'CONSTANTES TAB_GRUPO_COMPONENTE_AVANCADO
Public Enum GrupoComponente
    OCUPACAO_LOTE_26 = 26
    LIMITE_MURO_BALDIO_33 = 33
    SITUACAO_LOCALIZACAO_43 = 43
    TOPOGRAFIA_44 = 44
    PEDOLOGIA_45 = 45
    TIPO_IMOVEL_77 = 77
    
    AREA_TERRENO_108 = 108
    AREA_IMOVEL_113 = 113
    
    
    RESIDENCIAL_HORIZONTAL_39 = 39
    RESIDENCIAL_VERTICAL_77 = 77
    COMERCIAL_78 = 78
    OUTROS_4_81 = 81
End Enum
Private Function IPTUParcelado(CodObrigacao As String) As Boolean
    Dim Sql As String
       
    Sql = "SELECT TCO_COD_OBRIGACAO_PARCELA FROM TAB_COTAS_OBRIGACAO WHERE TCO_TOC_COD_OBRIGACAO =" & CodObrigacao
    IPTUParcelado = Bdados.AbreTabela(Sql)
End Function

Private Function AbreSelecao(Record As VSRecordset) As Boolean

    ' BUSCA OS IMÓVEIS PARA O QUAL SERÁ GERADO O IMPOSTO
    Dim Query As String
    Dim Tabela As String
    Dim Operador As String
    Dim Sql As String
    'PARA GERAR  O IMPOSTO
    Tabela = "SELECT * FROM  Vis_Imovel ,Tab_Contribuinte where tim_tci_im=tci_im  and Tab_Contribuinte.tci_nome not like '%PREFEITURA%' "
    'Tabela = Tabela & " AND TBA_TMU_COD_MUNICIPIO = " & AplicacoesVTFuncoes.Codigo_Municipio & " AND tlg_tmu_cod_municipio = " & AplicacoesVTFuncoes.Codigo_Municipio
    sqlValores = "SELECT TOC_INSCRICAO as IC,TOC_COD_OBRIGACAO, TOC_DATA_VENCIMENTO, tip_sigla_imposto as Imposto, TOC_VALOR_OBRIGACAO  as Valor " & _
        " FROM TAB_OBRIGACAO_CONTRIBUINTE, VIS_IMOVEL, TAB_IMPOSTO " & _
        " WHERE TOC_INSCRICAO = tim_ic AND TOC_tip_cod_imposto=tip_cod_imposto  AND toc_tip_cod_imposto IN ('" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "', '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "') and toc_periodo=" & txtAnoFi
    If Trim(txtic) <> "" Then
        If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
            If CInt(Right(txtic, 3)) <> 200 Then
                Sql = Sql & " and tim_ic ='" & txtic & "'"
            Else
                Sql = Sql & " and tim_ic > '" & txtic & "' and  tim_ic  <= '" & Left(txtic, 12) & "300'"
            End If
        Else
            Sql = Sql & " and tim_ic ='" & txtic & "'"
        End If
    End If
    If Trim(txtIM) <> "" Then
        Sql = Sql & " and tim_tci_im = '" & txtIM & "'"
    End If
    If Trim(cboTipoLogr) <> "" Then
        Sql = Sql & " and TTL_NOME = '" & cboTipoLogr & "'"
    End If
    If Trim(cboLogr) <> "" Then
        Sql = Sql & " and tlg_nome = '" & cboLogr & "'"
    End If
    If Trim(cboBairro) <> "" Then
        Sql = Sql & " and TBA_NOME = '" & cboBairro & "'"
    End If
    If Trim(txtValor01) <> "" And Trim(txtValor02) <> "" Then
        Sql = Sql & " and tim_valor >=" & CDbl(txtValor01) & " and tim_valor <=" & CDbl(txtValor02)
    End If
    If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
        If Trim(txtSecao) <> "" Then
            Sql = Sql & " and tim_ic like '__" & Format(txtSecao, "00") & "%'"
        End If
        If Trim(txtQuadra) <> "" Then
            Sql = Sql & " and tim_ic like '____" & Format(txtQuadra, "000") & "%'"
        End If
    End If
    If Trim(cboTipo) <> "" Then
        Sql = Sql & " and tim_tipo_imovel = " & cboTipo.ListIndex + 1
    End If
    'GAMBIARRA SILMAR
'    Sql = Sql & " AND TIM_IC NOT IN (SELECT DISTINCT tgt_tim_ic FROM Tab_Geracao_Tributo WHERE tgt_data_vencimento = CONVERT(DATETIME, '30/12/2002', 103) AND tgt_tip_cod_imposto = '11120200')"
    Tabela = Tabela & Sql & _
    "  order by tim_ic ASC ,tim_unidade ASC"
    sqlValores = sqlValores & Sql
    Screen.MousePointer = 11
    Tabela = Replace(Tabela, "TTL_NOME", "TIPOLOGRADOURO")
    Tabela = Replace(Tabela, "tlg_nome", "LOGRADOURO")
    AbreSelecao = Bdados.AbreTabela(Tabela, Record)  '(tabela)
    
End Function

Private Sub GeraExtratos()
    Dim CodPagamento As String
    Dim Sql As String
    Dim RsObrig As VSRecordset
    Dim rsImovel As VSRecordset
    Dim CodImposto  As String
    Dim Conta As New ContaCorrente
    Dim DataVenc As String
    Dim valorTotal As Double
    Dim Cont As Double
    Dim Total As Double
    Dim Campos As String
    Dim Valores As String
    Screen.MousePointer = 11
    If AbreSelecao(rsImovel) Then
        rsImovel.MoveFirst
        Cont = 0
        CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
        DataVenc = Imposto.BuscaDataVencimento(CodImposto, Year(Date))
        Do While Not rsImovel.EOF
            If cboRelatorio.Text = "IMPRIMIR BOLETO EM EXTRATO" Then
                Conta.ExecutaAtualizacao Trim(rsImovel!TIM_IC), etiImovel, False, , , DataVenc, , , , , CodImposto
                Sql = "Select TCC_INSCRICAO, TCC_CODIGO_CONTA,TCC_SALDO_ATUAL,TCC_PERIODO FROM " & _
                    " TAB_CONTA_CONTRIBUINTE WHERE TCC_INSCRICAO ='" & Trim(rsImovel!TIM_IC) & _
                    "' AND TCC_TIP_COD_IMPOSTO ='" & CodImposto & "' AND TCC_STATUS_CONTA IN (" & Const_NaoPagos & ")"
                If Bdados.AbreTabela(Sql, RsObrig) Then
                    CodPagamento = Conta.GeraCodPagamento(EtsExtratoPagamento)
                    Campos = "TPE_INSCRICAO, TPE_COD_PAGAMENTO_EXTRATO, TPE_TGT_COD_PAGAMENTO,TPE_TIP_COD_IMPOSTO,TPE_SUB_VALOR,TPE_TIPO_DOCUMENTO,TPE_SUB_PERIODO"
                    Do While Not RsObrig.EOF
                        Valores = Bdados.PreparaValor(Bdados.Converte(RsObrig!TCC_INSCRICAO, tctexto), CodPagamento, _
                                   RsObrig!TCC_CODIGO_CONTA, CodImposto, Bdados.Converte(RsObrig!TCC_SALDO_ATUAL, TCMonetario), _
                                    4, RsObrig!TCC_PERIODO)
                        Bdados.InsereDados "TAB_PAGAMENTO_EXTRATO", Valores, Campos
                        valorTotal = CDbl(Nvl(RsObrig!TCC_SALDO_ATUAL, 0)) + valorTotal
                        RsObrig.MoveNext
                    Loop
                    Conta.GeraPagamento Trim(rsImovel!TIM_IC), Trim(rsImovel!tim_tci_im), Const_Extrato, _
                     Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), _
                    DataVenc, CDbl(valorTotal), 0, 0, CodPagamento, 0, 0, 0, , EtcCreditoTributario
                    ImprimeExtrato CodPagamento, DataVenc, Trim(rsImovel!TIM_IC), valorTotal, CodImposto
                End If
            Else
                Sql = "Select TOC_COD_OBRIGACAO,TOC_INSCRICAO, TOC_TIP_COD_IMPOSTO,TOC_VALOR_OBRIGACAO,TOC_PERIODO,TOC_DATA_VENCIMENTO FROM " & _
                    " TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_INSCRICAO ='" & Trim(rsImovel!TIM_IC) & _
                    "' AND TOC_TIP_COD_IMPOSTO ='" & CodImposto & "' AND TOC_PERIODO =" & txtAnoFi
                If Bdados.AbreTabela(Sql, RsObrig) Then
                    With RPT
                        If Not IPTUParcelado(RsObrig!TOC_COD_OBRIGACAO) Then
                            If Not RPT.DefinirArquivo(Bdados, App.Path + "\TDAMExtratoBarra_TITULO_IPTU_2008_UNICA.Rpt") Then
                                Avisa "Arquivo do extrato não foi encontrado."
                                Screen.MousePointer = 0
                                Exit Sub
                            End If
                        Else
                            If Not RPT.DefinirArquivo(Bdados, App.Path + "\TDAMExtratoBarra_TITULO_IPTU_2008.rpt") Then
                                Avisa "Arquivo do extrato não foi encontrado."
                                Screen.MousePointer = 0
                                Exit Sub
                            End If
                        End If
                        Cobranca.ImprimeDamBarra RPT, Trim(RsObrig!TOC_INSCRICAO), RsObrig!TOC_TIP_COD_IMPOSTO, _
                            CStr(CDbl(Nvl(Trim(RsObrig!TOC_VALOR_OBRIGACAO), 0)) * 0.7), RsObrig!TOC_PERIODO, _
                            Nothing, RsObrig!TOC_DATA_VENCIMENTO, 0, RsObrig!TOC_COD_OBRIGACAO
                        .Selecao = "{TAB_OBRIGACAO_CONTRIBUINTE.TOC_COD_OBRIGACAO} =" & RsObrig!TOC_COD_OBRIGACAO
                        .Arvore = False
                        .Imprimir
                    End With
                    Screen.MousePointer = 0
                    Set RPT = Nothing
                End If
            End If
            rsImovel.MoveNext
            Cont = Cont + 1
            valorTotal = 0
            lblCont = Cont & " registros lidos"
            DoEvents
        Loop
        Avisa "Impressão concluída para " & Cont - 1 & " registros!"
    Else
        Informa "Nenhum registro encontrado."
    End If
    Screen.MousePointer = 0
End Sub

Private Sub ImprimeBoleto()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsComp As VSRecordset
    Dim Operador As String
    Dim Aux As Byte
    Dim AreaTotal As String
    Dim AreaConstruida As String
    Dim Desconto As Double
    Dim RsDesconto As VSRecordset
    Dim RsAux As VSRecordset
    Dim ValorMetro As Double
    Dim NomeLogr As String
    Dim Logr As String
    Dim CodigoLogr As String
    Dim Bairro As String
    Dim Cobranca As New VSCobranca
    Dim Conta As New ContaCorrente
    Dim CodImposto As String
    Dim Controle As Control
    Dim NomeImposto As String
    Dim Sigla As String
    Dim EnderecoCont As String
    Dim Obrig As New Obrigacao
    
    Screen.MousePointer = 11
    Aux = 0
    CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
    Sql = "select tip_sigla_imposto, tip_nome_imposto from tab_imposto where tip_cod_imposto ='" & CodImposto & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        Sigla = Rs!TIP_sigla_IMPOSTO
        NomeImposto = Rs!tip_nome_imposto
    End If
    Bdados.FechaTabela Rs
    Sql = "SELECT TIM_IC,TIM_tlg_cod_logradouro,tci_logradouro ,tci_nome_logradouro,tci_NUMERO,tci_BAIRRO,tci_cidade, " & _
        " TOC_COD_OBRIGACAO,TOC_INSCRICAO,tci_nome,TIM_IC,TTL_NOME,tlg_nome,tim_numero,TBA_NOME, TOC_PERIODO,tci_uf," & _
        " TOC_PARCELA,TOC_DATA_VENCIMENTO,tim_valor,TOC_VALOR_OBRIGACAO,TOC_VALOR_MULTA,TOC_valor_juros,TOC_TOTAL_TAXA_INCLUSA,tci_cgc_cpf" & _
        " FROM VIS_IMOVEL,TAB_OBRIGACAO_CONTRIBUINTE where TOC_INSCRICAO = tim_ic and TOC_tip_cod_imposto ='" & CodImposto & "'"
    For Each Controle In Controls
        If Controle.Tag <> "" Then
            If Trim(Controle.Text) <> "" Then
                Sql = Sql & " and " & Controle.Tag & " = '" & Trim(Controle.Text) & "'"
            End If
        End If
    Next
    If AplicacoesVTFuncoes.municipio <> "COLINAS" Then
        If AplicacoesVTFuncoes.municipio = "SANTA MARIA DA BOA VISTA" Then
            If Trim(txtSecao) <> "" Then Sql = Sql & " AND substring(tim_ic,2,4) = '" & txtSecao & "'"
            If Trim(txtQuadra) <> "" Then Sql = Sql & " AND substring(tim_ic,6,3) = '" & txtQuadra & "'"
        Else
            If Trim(txtSecao) <> "" Then Sql = Sql & " AND substring(tim_ic,3,2) = '" & txtSecao & "'"
            If Trim(txtQuadra) <> "" Then Sql = Sql & " AND substring(tim_ic,5,3) = '" & txtQuadra & "'"
        End If
    Else
        If Trim(txtSecao) <> "" Then Sql = Sql & " AND SUBSTRING(tim_ic_anterior,3,2) = '" & txtSecao & "'"
        If Trim(txtQuadra) <> "" Then Sql = Sql & " AND substring(tim_ic_anterior,5,3) = '" & txtQuadra & "'"
    End If
    Dim S As String
    S = RPT.Arquivo
            
    Sql = Sql & " ORDER BY tim_IC ASC"
    If Bdados.AbreTabela(Sql, Rs) Then
'        MontaGrid Bdados,lstIptu, Sql, 1400
        
        DoEvents
        Rs.MoveFirst
        Sql = "Select TGE_NOME from tab_geral where TGE_TIPO = 755 and TGE_CODIGO > 0"
        If Bdados.AbreTabela(Sql, RsDesconto) Then
            If IsNumeric(RsDesconto(0)) Then Desconto = RsDesconto(0)
        End If
        Do While Not Rs.EOF
            Sql = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & Rs!TIM_IC & _
                "' and (tdi_tco_cod_componente=113 or tdi_tco_cod_componente=108)"
            If Bdados.AbreTabela(Sql, RsComp) Then
                RsComp.MoveFirst
                Do While Not RsComp.EOF
                    If RsComp(0) = 108 Then
                        AreaTotal = RsComp(1)
                    ElseIf RsComp(0) = 113 Then
                        AreaConstruida = RsComp(1)
                    End If
                    RsComp.MoveNext
                Loop
            End If
            Bdados.FechaTabela RsComp
            
            Sql = "select tvl_valor  as ValorMetro from TAB_VALOR_TERRENO where tvl_tlg_cod_logradouro='" & Rs!tim_tlg_cod_logradouro & "'"
            If Bdados.AbreTabela(Sql, RsAux) Then
                ValorMetro = RsAux!ValorMetro
            End If
            EnderecoCont = Rs!tci_logradouro & " " & Rs!tci_nome_logradouro & " " & Rs!tci_NUMERO & " " & Rs!tci_BAIRRO & " " & Rs!tci_cidade & " " & Rs!tci_UF
            Bdados.FechaTabela RsAux
            Sql = "select ttl_nome as Logr,TLG_NOME AS Nome  from tab_logradouro,tab_tipo_logr where tlg_cod_logradouro='" & Rs!tim_tlg_cod_logradouro & "' and tlg_ttl_cod_tip_logr = ttl_cod_tip_logr and tlg_tmu_cod_municipio =" & AplicacoesVTFuncoes.Codigo_Municipio
            If Bdados.AbreTabela(Sql, RsAux) Then
                Logr = RsAux!Logr
                NomeLogr = RsAux!Nome
            End If
            Bdados.FechaTabela RsAux
            Obrig.BuscaDetalheObrigacao "" & Rs!TOC_COD_OBRIGACAO
            Dim robs As VSRecordset
            If Bdados.AbreTabela("SELECT TOC_OBS FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO=" & Rs!TOC_COD_OBRIGACAO, robs) Then
                If IsNull(robs(0)) Then
                    observacaoIPTU = ""
                Else
                    observacaoIPTU = robs(0)
                End If
            End If
            
            '' If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TDAMBarra_TITULO_IPTU.Rpt") Then
                ' Avisa "Arquivo não foi encontrado."
                'Screen.MousePointer = 0
                'Exit Sub
            'End If
            Cobranca.ImprimeDamIPTU RPT, "" & Rs!TOC_COD_OBRIGACAO, "" & Rs!TOC_INSCRICAO, "" & Rs!tci_nome, "" & Rs!TCI_CGC_CPF, EnderecoCont, _
             "" & Rs!TIM_IC, "" & Rs!TTL_NOME & " " & " " & Rs!tlg_nome & " " & Rs!tim_numero & " " & Rs!TBA_NOME, CodImposto, _
            Sigla, NomeImposto, "" & Rs!TOC_PERIODO, "" & Rs!TOC_PARCELA, 1, "" & Rs!TOC_DATA_VENCIMENTO, Nvl("" & Rs!tim_valor, 0), "" & Rs!TOC_VALOR_OBRIGACAO, "" & Rs!TOC_VALOR_MULTA, "" & Rs!TOC_valor_juros, "" & Rs!TOC_TOTAL_TAXA_INCLUSA, Obrig.obValorDesconto, "", observacaoIPTU, , , , , ValorMetro, , CDbl(Nvl(AreaTotal, 0)), CDbl(Nvl(AreaConstruida, 0)), , , , tdiImpressora, , , txtVencimento
'            Util.Pausa 1000
            AreaTotal = 0
            AreaConstruida = 0
            Rs.MoveNext
            DoEvents
        Loop
    Else
        Avisa "Nenhum Registro encontrado."
    End If
    Screen.MousePointer = 0
    Bdados.FechaTabela Rs
    Bdados.FechaTabela RsAux
    Bdados.FechaTabela RsComp
End Sub

Private Sub ImprimeExtrato(CodExtrato As String, DataVenc As String, Inscricao As String, ValorExtrato As Double, CodImposto As String)
    Dim Barra As Boolean
    Dim Cobranca As New VSCobranca
    Dim CgcPref As String
    Dim M As Boolean
    Dim i As Integer
    On Error GoTo trata
        

    If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
        If Not RPT.DefinirArquivo(Bdados, App.Path + "\TDAMExtratoBarra_TITULO_IPTU.rpt") Then
            Avisa "Arquivo do extrato não foi encontrado."
            Screen.MousePointer = 0
            Exit Sub
        End If
    Else
        If Not RPT.DefinirArquivo(Bdados, App.Path + "\TDAMExtratoBarra.rpt") Then
            Avisa "Arquivo do extrato não foi encontrado."
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If

    With RPT
        
        .Formulas "DATAVENCIMENTO", DataVenc
        .Formulas "PARCELA", "UNICA"
        .Formulas "VENCIMENTONORMAL", DataVenc
        If UCase(AplicacoesVTFuncoes.municipio) = "PETROLINA" Then
            .Formulas "TXDAM", TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
        End If
        .Formulas "PERIDO", Format(Year(Date), "0000")
        .Formulas "EMISSAO", Format(Date, "DD/MM/YYYY")
        .Formulas "VT_CARTEIRA", Temp.PegaParametro(Bdados, "CARTEIRA")
        .Formulas "VT_TAXAEXPEDIENTE", Format(CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0)), "###,###,###,#,##0.00")
        .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "Cliente")) & " - CNPJ Nº " & Temp.PegaParametro(Bdados, "CGC CLIENTE")
         .Formulas "VT_CodigoTributo", CodImposto & " - IPTU"
        .Selecao = "{TAB_PAGAMENTO_EXTRATO.TPE_COD_PAGAMENTO_EXTRATO} =" & CodExtrato
'        If Barra Then
            If AplicacoesVTFuncoes.municipio = "BARRA MANSA" Then
                .Formulas "VT_NOME_DAM", Temp.PegaParametro(Bdados, "NOME DAM")
            End If
        Cobranca.ImprimeDamBarra RPT, Inscricao, Const_Extrato, CStr(Format(CDbl(Nvl(Trim(ValorExtrato), 0)) + CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0)), "###,###,###,#,##0.00")), Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), Nothing, DataVenc, 0, CodExtrato
        .Arvore = False
        .Imprimir
    End With
    Screen.MousePointer = 0
    Set RPT = Nothing
    Exit Sub
trata:
    Avisa Err.Number & " - " & Err.Description
    Exit Sub
    Resume
End Sub

Sub ImprimeRelatorios()
    Dim RPT As VSRelatorio
    Dim RsCob As VSRecordset
    Dim FiltroRptIptu As String
    Dim Rs As VSRecordset
    
    Set RPT = New VSRelatorio
    FiltroRptIptu = ""
    If cboBairro <> "" Then
        FiltroRptIptu = " and {VIS_LISTAGEM_IPTU.tci_bairro} = '" & cboBairro & "'"
    End If
    If Trim(txtIM) <> "" Then
        FiltroRptIptu = FiltroRptIptu & " and {VIS_IMOVEL.TIM_tci_im} = '" & txtIM & "'"
    End If
    If Trim(txtSecao) <> "" Then
        FiltroRptIptu = FiltroRptIptu & " and MID({VIS_LISTAGEM_IPTU.tim_ic},3,2) = '" & Format(txtSecao, "00") & "'"
    End If
    If Trim(txtQuadra) <> "" Then
        FiltroRptIptu = FiltroRptIptu & " and MID({VIS_LISTAGEM_IPTU.tim_ic},5,4) = '" & Format(txtQuadra, "0000") & "'"
    End If
    If cboTipo <> "" Then
        FiltroRptIptu = FiltroRptIptu & " and {VIS_LISTAGEM_IPTU.tim_tipo_imovel} = " & cboTipo.ListIndex + 1
    End If
    If Trim(txtCodLogr) <> "" Then
        FiltroRptIptu = FiltroRptIptu & " and {VIS_LISTAGEM_IPTU.tim_tlg_cod_logradouro} = '" & txtCodLogr & "'"
    End If
    
    If Trim(txtValor01) <> "" And Trim(txtValor02) <> "" Then
        FiltroRptIptu = FiltroRptIptu & " and {VIS_LISTAGEM_IPTU.tim_valor} in " & Edita.TrocaPic(Edita.TrocaPic(txtValor01, ".", " "), ",", ".") & " to " & Edita.TrocaPic(Edita.TrocaPic(txtValor02, ".", " "), ",", ".")
    End If
    
    
    If cboRelatorio.Text = "IMPRIMIR BOLETO EM EXTRATO" Or cboRelatorio.Text = "IMPRIMIR BOLETO IPTU 2008" Then
        If Confirma("Deseja gerar o extrato para o(s) imovel(is) selecionados?") Then
            Screen.MousePointer = 11
            GeraExtratos
            Screen.MousePointer = 0
        End If
    ElseIf cboRelatorio.ListIndex = 3 Then
        If txtAnoIni = "" And txtAnoFi = "" Then Util.Avisa ("Informe o Período"): Exit Sub
        Screen.MousePointer = 11
        If Trim(txtic) <> "" Then
            If CInt(Right(txtic, 3)) <> 200 Then
                FiltroRptIptu = FiltroRptIptu & " and {VIS_LISTAGEM_IPTU.tim_ic}  = '" & txtic & "'"
            Else
                FiltroRptIptu = FiltroRptIptu & " and ({VIS_LISTAGEM_IPTU.tim_ic} = '" & txtic & "'"
            End If
        End If
        FiltroRptIptu = FiltroRptIptu & " and {VIS_IPTU_LANCADO.Período} = " & IIf(txtAnoIni <> "", txtAnoIni, txtAnoFi)
        If cboRelatorio.ListIndex = 2 Then 'Listagem
            If Not RPT.DefinirArquivo(Bdados, App.Path & "\TIPTUListagem.rpt") Then Exit Sub
        Else
            If Not RPT.DefinirArquivo(Bdados, App.Path & "\TIPTUListagem2.rpt") Then Exit Sub
        End If
        RPT.Selecao = Mid(FiltroRptIptu, 5)
        RPT.Formulas "VTPeriodo", IIf(txtAnoIni <> "", txtAnoIni, txtAnoFi)
        RPT.Arvore = False
        RPT.Visualizar
    ElseIf (cboRelatorio.ListIndex = 0) Or (cboRelatorio.ListIndex = 1) Then  'Boleto de IPTU, BOLETOS DE PARCELAS DO IPTU
        If AplicacoesVTFuncoes.Codigo_Municipio <> 2 Then
            ' If Temp.PegaParametro(Bdados, "TIPO IPTU") = 1 Then
           '     If AbreSelecao(RsCob) Then
           '         RsCob.MoveFirst
           '         CalculoIptu.AnoLancamento = Nvl(txtAnoIni, Year(Now))
           '         Do
            '            CalculoIptu.ImprimeBoletoIptu Trim(RsCob!tim_ic)
            '            RsCob.MoveNext
            '        Loop While Not RsCob.EOF
            '    End If
            '    Informa "Impressão de Boletos  Finalizada."
            'Else
                If Util.Confirma("Deseja gerar boleto para todos as ordens de pagamento selecionadas?") Then
                    ImprimeBoleto
                End If
           ' End If
        Else
            If IsNumeric(txtAnoIni) And Len(txtAnoIni) = 4 Then
                Screen.MousePointer = 11
                FiltroRptIptu = ""
                If Trim(txtic) <> "" Then
                    'If CInt(Right(txtIc, 3)) <> 200 Then
                        FiltroRptIptu = " and {VIS_IMOVEL.tim_ic} LIKE '" & txtic & "*'"
                    'Else
                    '    FiltroRptIptu = " and ({VIS_IMOVEL.tim_ic} = '" & txtIc & "'"
                    'End If
                End If
                If Trim(txtIM) <> "" Then
                    FiltroRptIptu = FiltroRptIptu & " and {VIS_IMOVEL.TIM_tci_im} = '" & txtIM & "'"
                End If
                If Trim(txtSecao) <> "" Then
                    FiltroRptIptu = FiltroRptIptu & " and MID({VIS_IMOVEL.TIM_IC},3,2) = '" & Format(txtSecao, "00") & "'"
                End If
                If Trim(txtQuadra) <> "" Then
                    FiltroRptIptu = FiltroRptIptu & " and MID({VIS_IMOVEL.TIM_IC},5,4) = '" & Format(txtQuadra, "0000") & "'"
                End If
                If cboBairro <> "" Then
                    FiltroRptIptu = FiltroRptIptu & " and {VIS_IMOVEL.TBA_NOME} = '" & cboBairro & "'"
                End If
                If cboTipo <> "" Then
                    FiltroRptIptu = FiltroRptIptu & " and {VIS_IMOVEL.tim_tipo_imovel} = " & cboTipo.ListIndex + 1
                End If
                If Trim(txtCodLogr) <> "" Then
                    FiltroRptIptu = FiltroRptIptu & " and {VIS_IMOVEL.tim_tlg_cod_logradouro} = '" & txtCodLogr & "'"
                End If
                If Trim(txtValor01) <> "" And Trim(txtValor02) <> "" Then
                    FiltroRptIptu = FiltroRptIptu & " and {VIS_IMOVEL.tim_valor} >=" & CDbl(txtValor01) & " and {VIS_IMOVEL.tim_valor} <=" & CDbl(txtValor02)
                End If
                
                Dim Operador As String
                
'                    0 - Listagem IPTU Gerados
'                    1 - Carnê do IPTU Grafica
'                    2 - Comparativo
'                    3 - VBUs Nulos
'                    4 - Lista de Logradouros/VBU
'                    5 - Listagem IPTU com valor zero
'                    6 - Carnê IPTU em Folha A4

                If cboRelatorio.ListIndex = 5 Then
                    If Not RPT.DefinirArquivo(Bdados, App.Path & "\TDAMBoletoGraficaIsento.rpt") Then Exit Sub
                    Operador = " <> "
                ElseIf cboRelatorio.ListIndex = 1 Then ' BOLETO A SER IMPRESSO NO MODELO PRÉ-IMPRESSO
                    If Not RPT.DefinirArquivo(Bdados, App.Path & "\TDAMBoletoGrafica.rpt") Then Exit Sub
                    Operador = " = "
                ElseIf cboRelatorio.ListIndex = 0 Then ' BOLETO A SER IMPRESSO EM FOLHA BRANCA A4
                    
                    If AplicacoesVTFuncoes.municipio = "BALSAS" Then
                        If Not RPT.DefinirArquivo(Bdados, App.Path & "\TDAMBARRA_IPTU.rpt") Then Exit Sub
                    Else
                        Dim Sql As String
                        FiltroRptIptu = FiltroRptIptu
                    End If
                    Operador = " = "
                End If
                'SUB-TSU
                RPT.SubRelatorio = "TSU"
                RPT.Formulas "Vt_cod_iptu", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "'"
                RPT.Formulas "Vt_cod_itu", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "'"
                RPT.Selecao = "{Tab_Geracao_Tributo.tgt_tim_ic} = {?Pm-Tab_Geracao_Tributo.tgt_tim_ic}" & _
                " AND {Tab_Geracao_Tributo.TGT_PARCELA} = 0 and {Tab_Geracao_Tributo.tgt_periodo} = " & txtAnoIni & _
                " and ({Tab_Imposto.tip_cod_imposto} <> {@VT_COD_IPTU} and {Tab_Imposto.tip_cod_imposto} <> {@VT_COD_ITU})"
                
                If cboRelatorio.ListIndex = 1 Or cboRelatorio.ListIndex = 0 Then
                    'SUB-COTA UNICA
                    RPT.SubRelatorio = ""
                    RPT.SubRelatorio = "COTA_UNICA"
                    RPT.Selecao = "{Tab_Geracao_Tributo.tgt_tim_ic} = {?Pm-Vis_Imovel.tim_ic} AND CDBL({Tab_Geracao_Tributo.tgt_periodo}) = " & txtAnoIni & _
                    " AND ({Tab_Geracao_Tributo.tgt_tip_cod_imposto} " & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & _
                    "' OR {Tab_Geracao_Tributo.tgt_tip_cod_imposto} " & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "')"
                End If
                
                'SUB-COTAS POSTERIORES
                RPT.SubRelatorio = ""
                RPT.SubRelatorio = "SUB_COTAS"
                
                RPT.Selecao = ""
                Dim Barra As String
                
                Barra = "{Tab_Geracao_Tributo_Parcela.tgt_tim_ic} = {?Pm-Vis_Imovel.tim_ic} AND " & _
                " CDBL({Tab_Geracao_Tributo_Parcela.tgt_periodo}) = " & txtAnoIni & _
                " AND ({Tab_Geracao_Tributo_Parcela.tgt_tip_cod_imposto} " & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & _
                "' OR {Tab_Geracao_Tributo_Parcela.tgt_tip_cod_imposto} " & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "')"
                RPT.Selecao = Barra
                
                'SUB-DIVIDA ATIVA
                RPT.SubRelatorio = ""
                RPT.SubRelatorio = "DividaAtiva"
                RPT.Formulas "VT_COD_DATIVA", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_DATIVA)) & "'"
                RPT.SubRelatorio = ""
                                             
                'RELATORIO PRINCIPAL
                Dim ValorMin As Double
                ValorMin = Imposto.BuscaValorMinimoImposto(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)), txtAnoIni)
                RPT.Formulas "Vt_Ano", txtAnoIni
                RPT.Formulas "Vt_cod_iptu", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "'"
                RPT.Formulas "Vt_cod_itu", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "'"
                RPT.Formulas "Vt_cod_dativa", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_DATIVA)) & "'"
                'Rpt.Formulas "Men", txtMens
                RPT.Formulas "Men_DA", "Contribuinte em Dívida Ativa. Impostos Atrasados Podem Ser Parcelados. Regularize Seu IPTU"
                RPT.Formulas "Men_OK", "Imposto Gera Educação e Saúde. Obrigado!"
                RPT.Selecao = " CDBL({Tab_Geracao_Tributo.tgt_periodo}) = CDBL({@VT_ANO}) " & _
                " AND ({Tab_Geracao_Tributo.tgt_tip_cod_imposto}" & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & _
                "' OR {Tab_Geracao_Tributo.tgt_tip_cod_imposto}" & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "') " & FiltroRptIptu & _
                " AND (cdbl({Tab_Geracao_Tributo.tgt_valor_tributo}) + cdbl({Tab_Geracao_Tributo.tgt_taxa_expediente})) " & IIf(ValorMin = 0, " > ", " >= ") & ValorMin
                'Raimudo
                
                RPT.Arvore = False
                RPT.Visualizar
            Else
                Avisa "Ano inválido."
                txtAnoFi.SetFocus
            End If
        End If
    ElseIf cboRelatorio.ListIndex = 8 Then 'Comparativo
        If Not RPT.DefinirArquivo(Bdados, App.Path & "\TIPTUComparacao.rpt") Then Exit Sub
        RPT.Arvore = False
        RPT.Selecao = Mid(FiltroRptIptu, 5) 'Mid(FiltroRptIptu & IIf(cboRelatorio.ListIndex = 2, " and ISNULL({TAB_TRECHO.TTC_VALOR})", ""), 5)
        RPT.Visualizar
    ElseIf cboRelatorio.ListIndex = 9 Then  ' VBU NULOS
        RPT.Arvore = False
        If Not RPT.DefinirArquivo(Bdados, App.Path & "\TVbuZero.rpt") Then Exit Sub
        
        RPT.Selecao = "isnull({TAB_TRECHO.TTC_VALOR} )" & IIf(Trim(Mid(FiltroRptIptu, 5)) <> "", FiltroRptIptu, "")
        RPT.Visualizar
    ElseIf cboRelatorio.ListIndex = 4 Then  'VBU
        RPT.Arvore = False
        FiltroRptIptu = "({TAB_GERACAO_TRIBUTO.TGT_TIP_COD_IMPOSTO} = '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "' or {TAB_GERACAO_TRIBUTO.TGT_TIP_COD_IMPOSTO} = '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "')"
        If Not RPT.DefinirArquivo(Bdados, App.Path & "\TVbuListagem_Lista.rpt") Then Exit Sub
        
        If Trim(txtCodLogr) <> "" Then
            RPT.Selecao = "and {TAB_TRECHO.TTC_tlg_cod_logradouro} = '" & txtCodLogr & "'"
        End If
        RPT.Visualizar
    ElseIf cboRelatorio.ListIndex = 5 Then ' Listagem de valor nulo
        Screen.MousePointer = 11
        If Trim(txtic) <> "" Then
            If CInt(Right(txtic, 3)) <> 200 Then
                FiltroRptIptu = FiltroRptIptu & " and {VIS_LISTAGEM_IPTU.tim_ic}  = '" & txtic & "'"
            Else
                FiltroRptIptu = FiltroRptIptu & " and ({VIS_LISTAGEM_IPTU.tim_ic} = '" & txtic & "'"
            End If
        End If
        ' FILTRA UM CAMPO COM VALOR IGUAL A ZERO 0,00
        If cboCampoZero.ListIndex = 0 Then
            FiltroRptIptu = FiltroRptIptu & " and {VIS_LISTAGEM_IPTU.AREA_LOTE} = 0"
        ElseIf cboCampoZero.ListIndex = 1 Then
            FiltroRptIptu = FiltroRptIptu & " and {VIS_LISTAGEM_IPTU.TIM_VALOR_TERRENO} = 0"
        ElseIf cboCampoZero.ListIndex = 2 Then
            FiltroRptIptu = FiltroRptIptu & " and {VIS_LISTAGEM_IPTU.TIM_VALOR_EDIFIC} = 0"
        ElseIf cboCampoZero.ListIndex = 3 Then
            FiltroRptIptu = FiltroRptIptu & " and {VIS_LISTAGEM_IPTU.TESTADA} = 0"
        ElseIf cboCampoZero.ListIndex = 4 Then
            FiltroRptIptu = FiltroRptIptu & " and {VIS_COMPARACAO_IPTU.IPTU2002} = 0"
        ElseIf cboCampoZero.ListIndex = 5 Then
            FiltroRptIptu = FiltroRptIptu & " and {VIS_TAX_CL.tgt_valor_tributo} = 0"
        ElseIf cboCampoZero.ListIndex = 6 Then
            FiltroRptIptu = FiltroRptIptu & " and {VIS_TAX_CVL.tgt_valor_tributo} = 0"
        ElseIf cboCampoZero.ListIndex = 7 Then
            FiltroRptIptu = FiltroRptIptu & " and {VIS_TAX_LP.tgt_valor_tributo} = 0"
        Else
            Screen.MousePointer = 0
            CalculoIptu.GeraParcelas = False
            CalculoIptu.GeraCotaUnica = True
            FiltroRptIptu = ""
            Util.Informa "Selecione um campo com valor 0,00 para ser exibido."
            cboCampoZero.SetFocus
            Exit Sub
        End If
        
        If Not RPT.DefinirArquivo(Bdados, App.Path & "\TIPTUListagemValorZero.rpt") Then Exit Sub
        RPT.Formulas "FiltroAplicado", Mid(cboCampoZero, 5) & " com valor 0,00"
        RPT.Selecao = Mid(FiltroRptIptu, 5)
        RPT.Arvore = False
        RPT.Visualizar
    ElseIf cboRelatorio.ListIndex = 7 Then ' Estatisticas
        If Not RPT.DefinirArquivo(Bdados, App.Path & "\TEstatisticaIPTU.rpt") Then Exit Sub
        RPT.Arvore = False
        RPT.Visualizar
    End If
End Sub

'Private Sub cboRelatorio_Click()
'    cboCampoZero.Visible = IIf(cboRelatorio.ListIndex <> 5, False, True)
'End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Valores As String
    Dim Campos As String
    Dim DataReab As Date
    Dim RsAux As VSRecordset
    Dim InscricaoMunicipal As String
    Dim InscricaoCadastral As String
    Dim CodLogr As Long
    Dim SitCadastral As String
    Dim i As Integer
    Dim Registros As Long
    lblCont = "0"
    Dim Sql As String
    Dim Query As String
    Dim DtVenc As String
    Dim Aux As Byte
    Dim RsCob As VSRecordset
    Dim Obrig As New Obrigacao
    Dim CodIPTU As String
    Dim Cont As Double
    Dim ValorCalculado As Double
    Dim VencimentoCotaUnica As String
    Dim IPTU As New VSIptu
    
    Dim pac As Boolean
    
    txtic = Trim(txtic)
    CodIPTU = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
    txtAnoIni = txtAnoFi
    Select Case cmd(Index).Caption
        Case "Gerar &Cota Única"
            If Trim(txtAnoFi) = "" Then
                Avisa "Informe o Ano para geração do imposto."
                Exit Sub
            End If
            If Not Util.Confirma("Gerar IPTU?") Then Exit Sub
            CodIPTU = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
            If AbreSelecao(RsCob) Then
                If AplicacoesVTFuncoes.municipio = "BARRA DO CORDA" Then
                    Cont = 0
                    Dim cria As New VSImposto
                    ValorCalculado = Imposto.CriaIptu(RsCob, CInt(txtAnoIni), CInt(txtAnoIni), lblIsentos)
                    
                    If grdValores.Preencher(Bdados, sqlValores) Then
                        grdValores.Mensagem = "Soma : " & Format$(grdValores.Colunas(3).Soma, "currency") & " x Menor : " & Format$(grdValores.Colunas(3).Min, "currency") & " x Maior : " & Format$(grdValores.Colunas(3).Max, "currency") & " x Média : " & Format$(grdValores.Colunas(3).Media, "currency")
                        Avisa "Dados gerados com sucesso!"
                    End If
                    DoEvents
                    Exit Sub
                    
                End If
                If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
                    ' SE O ANO FOR MENOR QUE O DO PGV, GERA COM O MODELO PADRÃO ANTIGO DO IPTU (padrão bci)
                    If CInt(Nvl(txtAnoFi, Year(Date))) < CInt(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
                        Call Imposto.GeraIptu(cip_Balsas, RsCob, CInt(txtAnoIni), CInt(txtAnoFi), tgi_SemParcelas)
                        Screen.MousePointer = 0
                        Call Util.Informa("Geração de imposto finalizada.")
                        If grdValores.Preencher(Bdados, sqlValores) Then grdValores.Mensagem = "Soma : " & Format$(grdValores.Colunas(3).Soma, "currency") & " x Menor : " & Format$(grdValores.Colunas(3).Min, "currency") & " x Maior : " & Format$(grdValores.Colunas(3).Max, "currency") & " x Média : " & Format$(grdValores.Colunas(3).Media, "currency")
                        Exit Sub
                    End If
                    RsCob.MoveFirst
                    If CInt(Nvl(txtAnoFi, Year(Date))) >= CInt(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
                        CalculoIptu.AnoLancamento = Nvl(txtAnoIni, Year(Now))
                        CalculoIptu.InicializarValores (Trim(RsCob!TIM_IC))
                        If CalculoIptu.GeraCotaUnica Then
                            CalculoIptu.GeraParcelas = Confirma("Deseja gerar as parcelas automaticamente?")
                        End If
                    End If
                    DoEvents
                    
                    Dim SqlIPTU As String
                    Dim Rs As VSRecordset
                    Dim rsLog  As VSRecordset
                    Dim Bairro As Integer, logradouro As Long
                    Dim areaTerreno As Double, areaImovel As Double
                    Dim ocupacao As Integer, limite As Integer, situacao As Integer, topografia As Integer, pedodologia As Integer
                    Dim Periodo As Integer, setor As Integer, quadra As Integer, Tipo As Integer
                    
                    Dim aliqSituacao As Double, aliqTopgrafia As Double, aliqPedologia As Double
                    Dim vvT As Double, vvE As Double, vvI As Double, valorIPTU As Double
                    Dim valorM2Terreno As Double, valorM2Edificacao As Double
                    
                    'CODO
                    Dim Grupo As Integer, Componente As Integer
                    
                    Dim imovel As String
                    Dim nomeLogradouro As String 'PARA GRAJAU
                    
                    'TEMP
                    Dim imoveis As Integer
                    Dim valorTotal As Double
                    imoveis = 0
                    valorTotal = 0
                    Dim tabRs As New Recordset
                    Dim Bd As New Connection
                    Dim strconn As String
                    strconn = Bdados.Conexao.DBConnection.ConnectionString
                    Bd.ConnectionString = strconn
                     
                        
                    Do
                        lblIsentos = "Calculando Valor da IC:  " & RsCob!TIM_IC
                        ' SE O ANO FOR MENOR QUE O DO PGV, GERA COM O MODELO PADRÃO ANTIGO DO IPTU
                        If CInt(Nvl(txtAnoFi, Year(Date))) < CInt(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
                            Call Imposto.GeraIptu(cip_Balsas, RsCob, CInt(txtAnoIni), CInt(txtAnoFi), tgi_SemParcelas)
                        Else
                        ' SE O ANO FOR COMPATÍVEL COM O PGV GERA ENTÃO O IPTU PGV
                            'If CalculoIptu.CarregaDetalheLote(Trim(RsCob!tim_ic)) Then
                                'CalculoIptu.CalculaValorIptu
                                'CALCULO NORMAL
                                'If CalculoIptu.ValorImposto > 0 Then Obrig.CriaObrigacao CodIPTU, txtAnoIni, txtAnoIni, Trim("" & RsCob!tim_ic), Format(CalculoIptu.ValorImposto, Const_Monetario), etsCreditoOriginalAberto, , CalculoIptu.VencimentoCotaUnica
                                'CALCULO BCP 18/05/2011
                                'METODO ELABORADO PARA A NECESSIDADE DA PREFEITURA DE GRAJAU E CODO
                                imovel = Trim(RsCob!TIM_IC)
                                pac = RsCob!tim_pac
                                nomeLogradouro = ""
                                Tipo = CInt(RsCob!tim_tipo_imovel)
                                Periodo = CInt(txtAnoFi)
                                setor = CInt(Mid(Trim(RsCob!TIM_IC), 3, 2))
                                quadra = CInt(Mid(Trim(RsCob!TIM_IC), 5, 3))
                                logradouro = CLng(RsCob!tim_tlg_cod_logradouro)
                                SqlIPTU = "select tim_tba_cod_bairro,tim_tlg_cod_logradouro from tab_imovel where tim_ic='" & imovel & "'"
                                Dim rstab As ADODB.Recordset
                                Set rstab = abreConexao(Bd, SqlIPTU)
                                If Not rstab.EOF Then
                                    ' If Bdados.AbreTabela("select tlg_nome from tab_logradouro where tlg_cod_logradouro=" & rs("tim_tlg_cod_logradouro"), rsLog) Then
                                    '     nomeLogradouro = rsLog("tlg_nome")
                                    ' End If
                                    ocupacao = CInt(pegaCaracteristicaImovel(imovel, GrupoComponente.OCUPACAO_LOTE_26))
                                    limite = CInt(pegaCaracteristicaImovel(imovel, GrupoComponente.LIMITE_MURO_BALDIO_33))
                                    situacao = CInt(pegaCaracteristicaImovel(imovel, GrupoComponente.SITUACAO_LOCALIZACAO_43))
                                    topografia = CInt(pegaCaracteristicaImovel(imovel, GrupoComponente.TOPOGRAFIA_44))
                                    pedodologia = CInt(pegaCaracteristicaImovel(imovel, GrupoComponente.PEDOLOGIA_45))
                                    areaTerreno = Format(CDbl(pegaCaracteristicaImovel(imovel, GrupoComponente.AREA_TERRENO_108)), "#,#,##0.00")
                                    areaImovel = Format(CDbl(pegaCaracteristicaImovel(imovel, GrupoComponente.AREA_IMOVEL_113)), "#,#,##0.00")
                                    
                                    Bairro = IIf(IsNull(rstab("tim_tba_cod_bairro")), 1, rstab("tim_tba_cod_bairro"))
                                    'Bdados.FechaTabela rs
                                    
                                    Dim parametroIPTU As New BCPParametroIPTU
                                    
                                    'DEPOIS DE PEGAR O VALORES, PEGO AS ALIQUOTAS CORRESPONDENTES NA TAB_BCP_PARAMETRO_IPTU
                                    If parametroIPTU.BuscarPorSetor(Periodo, Bairro, setor, quadra, pac, logradouro) Then
                                        aliqSituacao = Format(CDbl(pegaValorComponente(SITUACAO_LOCALIZACAO_43, situacao)), "#,#,##0.00")
                                        aliqTopgrafia = Format(CDbl(pegaValorComponente(TOPOGRAFIA_44, topografia)), "#,#,##0.00")
                                        aliqPedologia = Format(CDbl(pegaValorComponente(PEDOLOGIA_45, pedodologia)), "#,#,##0.00")
                                        
                                        valorM2Terreno = Format(CDbl(parametroIPTU.ValorTerreno), "#,#,##0.00")
                                        vvT = Format(CDbl(areaTerreno) * CDbl(valorM2Terreno) * CDbl(aliqSituacao) * CDbl(aliqTopgrafia) * CDbl(aliqPedologia), "#,#,##0.00")
                                        
                                        If Tipo = 2 Then 'TERRITORIAL
                                            vvI = vvT
                                            If limite = 1 Then ' SEM MURO / BALDIO
                                                valorIPTU = Format(CDbl(vvT) * (CDbl(parametroIPTU.AliquotaTerrenoBaldio) / CDbl(100)), "#,#,##0.00")
                                            Else ' COM MUTO / CALCADA
                                                valorIPTU = Format(CDbl(vvT) * (CDbl(parametroIPTU.AliquotaTerrenoMurado) / CDbl(100)), "#,#,##0.00")
                                            End If
                                        Else ' PREDIAL / CONDOMINIO
                                            Grupo = pegarGrupo(imovel)
                                            Componente = pegarComponente(imovel, Grupo)
                                            
                                            Dim padraoConstrutivo As New BCPParametroIPTU
                                            If padraoConstrutivo.BuscarPadraoConstrutivo(Periodo, setor, Grupo, Componente) Then
                                                Dim valorPadrao As Double, aliqPadrao As Double
                                                valorPadrao = Format(padraoConstrutivo.ValorPadraoConstrutivo, "#,#,##0.00")
                                                aliqPadrao = Format(padraoConstrutivo.AliquotaPadraoConstrutivo, "#,#,##0.00")
                                                
                                                vvE = Format((CDbl(areaImovel) * CDbl(valorPadrao)), "#,#,##0.00")
                                                vvI = vvE + vvT
                                                
                                                valorIPTU = Format(CDbl(vvI) * (CDbl(aliqPadrao) / 100), "#,#,##0.00")
                                            End If
                                        End If
                                            
                                    End If
                                    Dim lRes As Boolean
                                    lRes = fechaConexao(Bd, rstab)
                                     '**************************
                                     If vvI >= 15000 Then
                                        imoveis = imoveis + 1
                                        valorTotal = valorTotal + valorIPTU
                                       Dim rsVerificacao As VSRecordset
                                        If Bdados.AbreTabela("SELECT * FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_PERIODO=" & Periodo & " AND  TOC_INSCRICAO='" & imovel & "' AND TOC_TIP_COD_IMPOSTO='11120201'", rsVerificacao) Then
                                        Else
                                            '''''''''Obrig.CriaObrigacao CodIPTU, txtAnoIni, txtAnoIni, Trim("" & RsCob!tim_ic), Format(iptuValorTotal, Const_Monetario), etsCreditoOriginalAberto, , CalculoIptu.VencimentoCotaUnica
                                           Obrig.CriaObrigacao CodIPTU, txtAnoIni, txtAnoIni, Trim("" & RsCob!TIM_IC), Format(valorIPTU, Const_Monetario), etsCreditoOriginalAberto, , txtVencimento
                                           Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_OBS='" & observacaoIPTU & "', TOC_REMESSA=" & Format(Now, "yyyyMM") & "WHERE TOC_COD_OBRIGACAO=" & Obrig.obCodigoObrigacao)
                                           observacaoIPTU = ""
                                        End If
                                        'FIM CALCULO BCP - PODE SER TEMPORARIO
                                        lblIsentos = "  Finalizando Unidade " & RsCob!TIM_IC
                                        lblCont = "  Registros Processados: " & CalculoIptu.RegistrosProcessados
                                        lblIsentos = "  O lote " & Trim(RsCob!TIM_IC) & " já registra pagamento para este periodo. Regularize esta situação."
                                    End If
                                End If
                        End If
                        
                        DoEvents
                        RsCob.MoveNext
                    Loop While Not RsCob.EOF
                Else
                    RsCob.MoveFirst
                    Cont = 0
                    Do
                        lblIsentos = "Calculando Valor da IC:  " & RsCob!TIM_IC
'                        lblCont = cadastro.CriaIptu(RsCob, CInt(Nvl(txtAnoIni, Year(Now))), CInt(Nvl(txtAnoIni, Year(Now))), lblCont)
                        Obrig.CriaObrigacao CodIPTU, txtAnoIni, txtAnoIni, Trim("" & RsCob!TIM_IC), , etsCreditoOriginalAberto
                        lblIsentos = "  Finalizando Unidade " & RsCob!TIM_IC
                        Cont = Cont + 1
                        lblCont = "  Registros Processados: " & Cont
                        DoEvents
                        RsCob.MoveNext
                    Loop While Not RsCob.EOF
                End If
            Else
                Call Util.Informa("Nenhum lote encontrado.")
                Screen.MousePointer = 0
                lblIsentos = ""
                Exit Sub
            End If
            prgIptu.Visible = False
            Call Util.Informa("Geração de imposto finalizada.")
            sqlValores = Replace(sqlValores, "TTL_NOME", "TIPOLOGRADOURO")
            sqlValores = Replace(sqlValores, "tlg_nome", "LOGRADOURO")
     
            If grdValores.Preencher(Bdados, sqlValores) Then
                grdValores.Mensagem = "Soma : " & Format$(grdValores.Colunas(3).Soma, "currency") & " x Menor : " & Format$(grdValores.Colunas(3).Min, "currency") & " x Maior : " & Format$(grdValores.Colunas(3).Max, "currency") & " x Média : " & Format$(grdValores.Colunas(3).Media, "currency")
            End If
            'BCP TECNOLOGIAS -
            Dim o As Integer
            'For o = 1 To grdValores.ListItems.Count
             '   Dim iptuObrig As New Obrigacao
              '  Dim ob As String
                'Dim dt  As String
               ' dt = Format(txtVencimento, "dd/MM/yyyy")
                'ob = grdValores.ListItems(o).SubItems(1)
                'dt = Right(dt, 4) & "-" & Left(dt, 2) & "-" & Mid(dt, 4, 2)
                'Bdados.Executa ("update tab_obrigacao_contribuinte set toc_data_geracao='" & dt & "',toc_data_vencimento='" & dt & "' where toc_cod_obrigacao=" & ob)
                'Bdados.Executa ("update tab_conta_contribuinte set tcc_data_movimento='" & dt & "',tcc_data_vencimento='" & dt & "' where tcc_codigo_conta=" & ob)
            'Next o
            '''''
            lblIsentos = ""
            txtic.SetFocus
            Screen.MousePointer = 0
        Case "&Imprimir"
            ImprimeRelatorios
        Case "Gerar &Parcelas"
            If Not Util.Confirma("Gerar Parcelas?") Then Exit Sub
                CalculoIptu.GeraParcelas = True
                CalculoIptu.GeraCotaUnica = False
                cmd_Click 3
        Case "Sai&r"
            
            Unload Me
    End Select
    
    Screen.MousePointer = 0
    CalculoIptu.GeraParcelas = False
    CalculoIptu.GeraCotaUnica = True
End Sub

Private Function pegaCaracteristicaImovel(Inscricao As String, Grupo As GrupoComponente) As Double
    'PEGA O COMPONENTE DO DETALHE DE IMOVEL, POSTERIORMENTE CADA COMPONENTE TEM SEU VALOR
    Dim Valor As Double
    Dim Componente As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "select tdi_tco_cod_componente as componente, tdi_valor_item as valor from TAB_DETALHE_IMOVEL where tdi_tim_ic = '" & Inscricao & "' and tdi_tgc_cod_grupo=" & Grupo
    
    If Grupo = AREA_IMOVEL_113 Or Grupo = AREA_TERRENO_108 Then
        
        Dim Bd As New Connection
        Bd.ConnectionString = Bdados.Conexao.DBConnection.ConnectionString
        Bd.Open
    Dim rstab As New Recordset
    Set rstab = Bd.Execute(Sql)
    
    
        If Not rstab.EOF Then
            Valor = IIf(IsNull(rstab("valor")), 0, rstab("valor"))
        End If
        rstab.Close
        Bd.Close
        
    Else
      
       Dim bd1 As New Connection
       Dim rst As New Recordset
       bd1.ConnectionString = Bdados.Conexao.DBConnection.ConnectionString
       bd1.Open
       Set rst = bd1.Execute(Sql)
       
        If Not rst.EOF Then
            Valor = IIf(IsNull(rst("componente")), 1, rst("componente"))
        End If
        rst.Close
        bd1.Close
        
    End If
    pegaCaracteristicaImovel = Valor
End Function
Private Function pegaValorComponente(Grupo As GrupoComponente, Componente As Integer) As Double
    'PEGA O COMPONENTE DO DETALHE DE IMOVEL, POSTERIORMENTE CADA COMPONENTE TEM SEU VALOR
    Dim Valor As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "SELECT tco_valor AS valor, tco_cod_componente AS componente, tco_grupo AS grupo, tco_descricao_componente From Tab_Componente_AVANCADO Where tco_grupo=" & Grupo & " and tco_cod_componente=" & Componente
    
    Dim Bd As New Connection
    Bd.ConnectionString = Bdados.Conexao.DBConnection.ConnectionString
    Bd.Open
    Dim rstab As New Recordset
    Set rstab = Bd.Execute(Sql)
    
    
    If Not rstab.EOF Then
        Valor = IIf(IsNull(rstab("valor")), 1, rstab("valor"))
    End If
    pegaValorComponente = Valor
End Function
Private Function pegarComponente(Inscricao As String, Grupo As Integer) As Integer
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Componente As Integer
    Sql = "SELECT tdi_tco_cod_componente  From TAB_DETALHE_IMOVEL " _
    & " WHERE tdi_tim_ic = '" & Inscricao & "' AND tdi_tgc_cod_grupo=" & Grupo
    
    Dim Bd As New Connection
    Bd.ConnectionString = Bdados.Conexao.DBConnection.ConnectionString
    Bd.Open
    Dim rstab As New Recordset
    Set rstab = Bd.Execute(Sql)
    
    
    If Not rstab.EOF Then
        Componente = IIf(IsNull(rstab("tdi_tco_cod_componente")), 3, rstab("tdi_tco_cod_componente")) ' 3=1-C
    End If
    pegarComponente = Componente
    rstab.Close
    Bd.Close
    
End Function
Private Function pegarGrupo(Inscricao As String) As Integer
    Dim Grupo As Integer
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "SELECT tdi_tgc_cod_grupo  From TAB_DETALHE_IMOVEL " _
    & " WHERE tdi_tim_ic = '" & Inscricao & "' AND tdi_tgc_cod_grupo IN (39, 77, 78, 81)"
    
    Dim Bd As New Connection
    Bd.ConnectionString = Bdados.Conexao.DBConnection.ConnectionString
    Bd.Open
    Dim rstab As New Recordset
    Set rstab = Bd.Execute(Sql)
    
    
    If Not rstab.EOF Then
        Grupo = IIf(IsNull(rstab("tdi_tgc_cod_grupo")), GrupoComponente.RESIDENCIAL_HORIZONTAL_39, rstab("tdi_tgc_cod_grupo"))
    End If
    pegarGrupo = Grupo
    rstab.Close
    Bd.Close
    
End Function

Private Function retornaVuc(Grupo As Integer, Componente As Integer) As Double
    Dim Valor As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "select tco_valor from tab_componente_avancado where tco_grupo=" & Grupo & " and tco_cod_componente=" & Componente
    If Bdados.AbreTabela(Sql, Rs) Then
        Valor = IIf(IsNull(Rs("tco_valor")), 1, Rs("tco_valor"))
    Else
        Valor = 1
    End If
    retornaVuc = Valor
End Function
Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim Controle As Control
    Dim i As Byte
    Set cadastro = New VSImposto
    Call Edita.AtualizaCombo(Bdados, cboLogr, "Select DISTINCT(tlg_nome) From Tab_Logradouro ")
    Call Edita.AtualizaCombo(Bdados, cboTipoLogr, "Select DISTINCT(ttl_nome) From Tab_Tipo_Logr")
    Call Edita.AtualizaCombo(Bdados, cboBairro, "Select DISTINCT(tba_nome) From Tab_Bairro ")
    cboLogr.AddItem ""
    cboTipoLogr.AddItem ""
    cboBairro.AddItem ""
    cboRelatorio.AddItem ""
    cboRelatorio.AddItem "IMPRIMIR BOLETO EM FOLHA A4"
    cboRelatorio.AddItem "IMPRIMIR BOLETO EM EXTRATO"
    cboRelatorio.AddItem "IMPRIMIR BOLETO IPTU 2008"
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    cboCampoZero.Visible = False
    txtVencimento = Format(Now, "dd/MM/yyyy")
    txtAnoFi = Format(Now, "yyyy")
    
End Sub

Private Sub grdValores_Click()

End Sub

Private Sub txtAnoFi_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtAnoFi_Validate(Cancel As Boolean)
    txtAnoIni = txtAnoFi
End Sub

Private Sub txtAnoIni_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_LostFocus()
'    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
'        txtic = cadastro.FormataInscricao(txtic, InscImovel)
'    End If
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    txtIM = cadastro.FormataInscricao(txtIM, InscContrib)
End Sub

Private Sub txtValor01_LostFocus()
    txtValor01 = Edita.FormataTexto(txtValor01, Monetario)
End Sub

Private Sub txtValor02_LostFocus()
    txtValor02 = Edita.FormataTexto(txtValor02, Monetario)
End Sub

Private Function abreConexao(Bd As Connection, strSql As String) As Recordset
   Dim tabRs As New Recordset
                  
  Bd.Open
  Set tabRs = Bd.Execute(strSql)
  Set abreConexao = tabRs
End Function
Private Function fechaConexao(Bd As Connection, rsObj As Recordset) As Boolean
  rsObj.Close
  Bd.Close
  fechaConexao = True
  
End Function


