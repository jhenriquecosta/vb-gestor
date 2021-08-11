VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCOB101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleMode       =   0  'User
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   40
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCOB101.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.grdVISUAL grdValores 
      Height          =   2055
      Left            =   60
      TabIndex        =   39
      Top             =   3750
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   3625
      CorBorda        =   32768
      Caption         =   "Valores"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   0
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   22
      Top             =   6675
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   3
         Left            =   2595
         TabIndex        =   12
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         Caption         =   "&Cota Única"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   3945
         TabIndex        =   13
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Parcelas"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   5115
         TabIndex        =   16
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   6285
         TabIndex        =   17
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2985
      Left            =   75
      TabIndex        =   21
      Top             =   720
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   5265
      Altura          =   1905
      Caption         =   " Detalhes"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lblCont 
         Height          =   285
         Left            =   2520
         TabIndex        =   37
         Top             =   2250
         Width           =   4695
         _ExtentX        =   8281
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
         TabIndex        =   36
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
         Left            =   1875
         MaxLength       =   4
         TabIndex        =   35
         Tag             =   "TGT_PERIODO"
         Top             =   2235
         Visible         =   0   'False
         Width           =   255
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
         Tag             =   "TGT_PERIODO"
         Top             =   2235
         Width           =   885
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   11
         Left            =   3060
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         ItemData        =   "TCOB101.frx":2123
         Left            =   105
         List            =   "TCOB101.frx":2130
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
         ItemData        =   "TCOB101.frx":2151
         Left            =   5385
         List            =   "TCOB101.frx":215E
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
         ItemData        =   "TCOB101.frx":217F
         Left            =   4020
         List            =   "TCOB101.frx":2181
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
         Tag             =   "TIM_IC"
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
         Tag             =   "TIM_IC"
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
         ItemData        =   "TCOB101.frx":2183
         Left            =   5385
         List            =   "TCOB101.frx":218D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "TTL_NOME"
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
         Tag             =   "TIM_IC"
         Top             =   600
         Width           =   915
      End
      Begin Threed.SSPanel lblIsentos 
         Height          =   330
         Left            =   15
         TabIndex        =   38
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
         TabIndex        =   34
         Top             =   1905
         Width           =   285
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   1138
      Icone           =   "TCOB101.frx":21A7
   End
   Begin VB.PictureBox prgIptu 
      Height          =   195
      Left            =   4545
      ScaleHeight     =   135
      ScaleWidth      =   2670
      TabIndex        =   19
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
      TabIndex        =   18
      Top             =   180
      Width           =   375
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   795
      Left            =   75
      TabIndex        =   23
      Top             =   5820
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   1402
      Altura          =   1905
      Caption         =   " Relatórios"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   435
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   397
         _Version        =   196610
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Relatório"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.ComboBox cboRelatorio 
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
         ItemData        =   "TCOB101.frx":24C1
         Left            =   960
         List            =   "TCOB101.frx":24E3
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   375
         Width           =   3675
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
         ItemData        =   "TCOB101.frx":25EB
         Left            =   4680
         List            =   "TCOB101.frx":2607
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   375
         Width           =   2625
      End
   End
End
Attribute VB_Name = "TCOB101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim cadastro As VSImposto
Dim Cobranca As New VSCobranca
Dim CalculoIptu As New VSIptu
Dim sqlValores As String
Private Function AbreSelecao(Record As VSRecordset) As Boolean
    ' BUSCA OS IMÓVEIS PARA O QUAL SERÁ GERADO O IMPOSTO
    Dim Query As String
    Dim Tabela As String
    Dim Operador As String
    Dim Sql As String
    'PARA GERAR  O IMPOSTO
    Tabela = "SELECT * FROM  Vis_Imovel ,Tab_Contribuinte where tim_tci_im=tci_im AND TIM_SITUACAO_LOTE <> 1"
    Tabela = Tabela & " AND TBA_TMU_COD_MUNICIPIO = " & Aplicacoes.Codigo_Municipio & " AND tlg_tmu_cod_municipio = " & Aplicacoes.Codigo_Municipio
    sqlValores = "SELECT tgt_tim_ic as IC, tip_sigla_imposto as Imposto, tgt_valor_tributo + tgt_taxa_expediente as Valor" & _
        " FROM TAB_GERACAO_TRIBUTO, VIS_IMOVEL, TAB_CONTRIBUINTE, TAB_IMPOSTO" & _
        " WHERE tgt_tim_ic=tim_ic AND tgt_im=tci_im AND tgt_tip_cod_imposto=tip_cod_imposto AND tgt_tip_cod_imposto IN ('" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "', '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "') and tgt_periodo=" & txtAnoFi
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
            Sql = Sql & " and tim_ic like '__" & txtSecao & "%'"
        End If
        If Trim(txtQuadra) <> "" Then
            Sql = Sql & " and tim_ic like '____0" & txtQuadra & "%'"
        End If
    End If
    'GAMBIARRA SILMAR
'    Sql = Sql & " AND TIM_IC NOT IN (SELECT DISTINCT tgt_tim_ic FROM Tab_Geracao_Tributo WHERE tgt_data_vencimento = CONVERT(DATETIME, '30/12/2002', 103) AND tgt_tip_cod_imposto = '11120200')"
    Tabela = Tabela & Sql & _
    "  order by tim_ic ASC ,tim_unidade ASC"
    sqlValores = sqlValores & Sql
    Screen.MousePointer = 11
    AbreSelecao = Bdados.AbreTabela(Tabela, Record)  '(tabela)
    
End Function

Private Sub ImprimeBoleto()
    Dim Sql As String
    Dim rs As VSRecordset
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
    
    Screen.MousePointer = 11
    Aux = 0
    CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
    Sql = "select tip_sigla_imposto, tip_nome_imposto from tab_imposto where tip_cod_imposto ='" & CodImposto & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        Sigla = rs!TIP_sigla_IMPOSTO
        NomeImposto = rs!tip_nome_imposto
    End If
    Bdados.FechaTabela rs
    Sql = "SELECT * FROM VIS_IMOVEL,tab_geracao_tributo where tgt_tim_ic = tim_ic and tgt_tip_cod_imposto ='" & CodImposto & "' and tgt_im = tim_tci_im "
    For Each Controle In Controls
        If Controle.Tag <> "" Then
            If Trim(Controle.Text) <> "" Then
                Sql = Sql & " and " & Controle.Tag & " = '" & Trim(Controle.Text) & "'"
            End If
        End If
    Next
    Sql = Sql & " ORDER BY tim_IC ASC"
    If Bdados.AbreTabela(Sql, rs) Then
'        MontaGrid Bdados,lstIptu, Sql, 1400
        
        DoEvents
        rs.MoveFirst
        Sql = "Select TGE_NOME from tab_geral where TGE_TIPO = 755 and TGE_CODIGO > 0"
        If Bdados.AbreTabela(Sql, RsDesconto) Then
            Desconto = RsDesconto(0)
        End If
        Do While Not rs.EOF
            Sql = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & rs!TIM_IC & _
                "' and (tdi_tco_cod_componente=110 or tdi_tco_cod_componente=108)"
            If Bdados.AbreTabela(Sql, RsComp) Then
                RsComp.MoveFirst
                Do While Not RsComp.EOF
                    If RsComp(0) = 110 Then
                        AreaTotal = RsComp(1)
                    ElseIf RsComp(0) = 108 Then
                        AreaConstruida = RsComp(1)
                    End If
                    RsComp.MoveNext
                Loop
            End If
            Bdados.FechaTabela RsComp
            
            Sql = "select tvl_valor  as ValorMetro from TAB_VALOR_TERRENO where tvl_tlg_cod_logradouro='" & rs!TIM_tlg_cod_logradouro & "'"
            If Bdados.AbreTabela(Sql, RsAux) Then
                ValorMetro = RsAux!ValorMetro
            End If
            EnderecoCont = rs!tci_logradouro & " " & rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_BAIRRO
            Bdados.FechaTabela RsAux
            Sql = "select ttl_nome as Logr,TLG_NOME AS Nome  from tab_logradouro,tab_tipo_logr where tlg_cod_logradouro='" & rs!TIM_tlg_cod_logradouro & "' and tlg_ttl_cod_tip_logr = ttl_cod_tip_logr and tlg_tmu_cod_municipio =" & Aplicacoes.Codigo_Municipio
            If Bdados.AbreTabela(Sql, RsAux) Then
                Logr = RsAux!Logr
                NomeLogr = RsAux!Nome
            End If
            Bdados.FechaTabela RsAux
            Cobranca.ImprimeDam Rpt, rs!tgt_cod_pagamento, rs!TGT_im, rs!tci_nome, "", EnderecoCont, _
             rs!TIM_IC, rs!TTL_NOME & " " & " " & rs!tlg_nome & " " & rs!tim_numero & " " & rs!TBA_NOME, CodImposto, _
            Sigla, NomeImposto, rs!TGT_PERIODO, IIf(rs!TGT_PARCELA = 0, "UNICA", rs!TGT_PARCELA), 1, rs!TGT_DATA_VENCIMENTO, rs!tim_valor, rs!TGT_VALOR_TRIBUTO, rs!TGT_VALOR_MULTA, rs!tgt_valor_juros, rs!tgt_taxa_expediente, 0, "", "", , , , , ValorMetro, , CDbl(AreaTotal), CDbl(AreaConstruida), , , , tdiImpressora
'            Util.Pausa 1000
            AreaTotal = 0
            AreaConstruida = 0
            rs.MoveNext
        Loop
    Else
        Avisa "Nenhum Registro encontrado."
    End If
    Screen.MousePointer = 0
    Bdados.FechaTabela rs
    Bdados.FechaTabela RsAux
    Bdados.FechaTabela RsComp
End Sub

Sub ImprimeRelatorios()
    Dim Rpt As VSRelatorio
    Dim RsCob As VSRecordset
    Dim FiltroRptIptu As String
    Dim rs As VSRecordset
    
    Set Rpt = New VSRelatorio
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
    
    If cboRelatorio.ListIndex = 2 Or cboRelatorio.ListIndex = 3 Then   'Listagem
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
            If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TIPTUListagem.rpt") Then Exit Sub
        Else
            If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TIPTUListagem2.rpt") Then Exit Sub
        End If
        Rpt.Selecao = Mid(FiltroRptIptu, 5)
        Rpt.Formulas "VTPeriodo", IIf(txtAnoIni <> "", txtAnoIni, txtAnoFi)
        Rpt.Arvore = False
        Rpt.Visualizar
    ElseIf (cboRelatorio.ListIndex = 0) Or (cboRelatorio.ListIndex = 1) Then  'Boleto de IPTU, BOLETOS DE PARCELAS DO IPTU
        If Aplicacoes.Codigo_Municipio <> 2 Then
            If Temp.PegaParametro(Bdados, "TIPO IPTU") = 1 Then
                If AbreSelecao(RsCob) Then
                    RsCob.MoveFirst
                    CalculoIptu.AnoLancamento = Nvl(txtAnoIni, Year(Now))
                    Do
                        CalculoIptu.ImprimeBoletoIptu Trim(RsCob!TIM_IC)
                        rs.MoveNext
                    Loop While Not rs.EOF
                End If
                Informa "Impressão de Boletos  Finalizada."
            Else
                If Util.Confirma("Deseja gerar boleto para todos as ordens de pagamento selecionadas?") Then
                    ImprimeBoleto
                End If
            End If
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
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TDAMBoletoGraficaIsento.rpt") Then Exit Sub
                    Operador = " <> "
                ElseIf cboRelatorio.ListIndex = 1 Then ' BOLETO A SER IMPRESSO NO MODELO PRÉ-IMPRESSO
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TDAMBoletoGrafica.rpt") Then Exit Sub
                    Operador = " = "
                ElseIf cboRelatorio.ListIndex = 0 Then ' BOLETO A SER IMPRESSO EM FOLHA BRANCA A4
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TDAMBoletoA4.rpt") Then Exit Sub
                    Operador = " = "
                End If
                'SUB-TSU
                Rpt.SubRelatorio = "TSU"
                Rpt.Formulas "Vt_cod_iptu", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "'"
                Rpt.Formulas "Vt_cod_itu", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "'"
                Rpt.Selecao = "{Tab_Geracao_Tributo.tgt_tim_ic} = {?Pm-Tab_Geracao_Tributo.tgt_tim_ic}" & _
                " AND {Tab_Geracao_Tributo.TGT_PARCELA} = 0 and {Tab_Geracao_Tributo.tgt_periodo} = " & txtAnoIni & _
                " and ({Tab_Imposto.tip_cod_imposto} <> {@VT_COD_IPTU} and {Tab_Imposto.tip_cod_imposto} <> {@VT_COD_ITU})"
                
                If cboRelatorio.ListIndex = 1 Or cboRelatorio.ListIndex = 0 Then
                    'SUB-COTA UNICA
                    Rpt.SubRelatorio = ""
                    Rpt.SubRelatorio = "COTA_UNICA"
                    Rpt.Selecao = "{Tab_Geracao_Tributo.tgt_tim_ic} = {?Pm-Vis_Imovel.tim_ic} AND CDBL({Tab_Geracao_Tributo.tgt_periodo}) = " & txtAnoIni & _
                    " AND ({Tab_Geracao_Tributo.tgt_tip_cod_imposto} " & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & _
                    "' OR {Tab_Geracao_Tributo.tgt_tip_cod_imposto} " & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "')"
                End If
                
                'SUB-COTAS POSTERIORES
                Rpt.SubRelatorio = ""
                Rpt.SubRelatorio = "SUB_COTAS"
                
                Rpt.Selecao = ""
                Dim Barra As String
                
                Barra = "{Tab_Geracao_Tributo_Parcela.tgt_tim_ic} = {?Pm-Vis_Imovel.tim_ic} AND " & _
                " CDBL({Tab_Geracao_Tributo_Parcela.tgt_periodo}) = " & txtAnoIni & _
                " AND ({Tab_Geracao_Tributo_Parcela.tgt_tip_cod_imposto} " & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & _
                "' OR {Tab_Geracao_Tributo_Parcela.tgt_tip_cod_imposto} " & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "')"
                Rpt.Selecao = Barra
                
                'SUB-DIVIDA ATIVA
                Rpt.SubRelatorio = ""
                Rpt.SubRelatorio = "DividaAtiva"
                Rpt.Formulas "VT_COD_DATIVA", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_DATIVA)) & "'"
                Rpt.SubRelatorio = ""
                                             
                'RELATORIO PRINCIPAL
                Dim ValorMin As Double
                ValorMin = Imposto.BuscaValorMinimoImposto(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)), txtAnoIni)
                Rpt.Formulas "Vt_Ano", txtAnoIni
                Rpt.Formulas "Vt_cod_iptu", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "'"
                Rpt.Formulas "Vt_cod_itu", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "'"
                Rpt.Formulas "Vt_cod_dativa", "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_DATIVA)) & "'"
                'Rpt.Formulas "Men", txtMens
                Rpt.Formulas "Men_DA", "Contribuinte em Dívida Ativa. Impostos Atrasados Podem Ser Parcelados. Regularize Seu IPTU"
                Rpt.Formulas "Men_OK", "Imposto Gera Educação e Saúde. Obrigado!"
                Rpt.Selecao = " CDBL({Tab_Geracao_Tributo.tgt_periodo}) = CDBL({@VT_ANO}) " & _
                " AND ({Tab_Geracao_Tributo.tgt_tip_cod_imposto}" & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & _
                "' OR {Tab_Geracao_Tributo.tgt_tip_cod_imposto}" & Operador & "'" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "') " & FiltroRptIptu & _
                " AND (cdbl({Tab_Geracao_Tributo.tgt_valor_tributo}) + cdbl({Tab_Geracao_Tributo.tgt_taxa_expediente})) " & IIf(ValorMin = 0, " > ", " >= ") & ValorMin
                'Raimudo
                
                Rpt.Arvore = False
                Rpt.Visualizar
            Else
                Avisa "Ano inválido."
                txtAnoFi.SetFocus
            End If
        End If
    ElseIf cboRelatorio.ListIndex = 8 Then 'Comparativo
        If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TIPTUComparacao.rpt") Then Exit Sub
        Rpt.Arvore = False
        Rpt.Selecao = Mid(FiltroRptIptu, 5) 'Mid(FiltroRptIptu & IIf(cboRelatorio.ListIndex = 2, " and ISNULL({TAB_TRECHO.TTC_VALOR})", ""), 5)
        Rpt.Visualizar
    ElseIf cboRelatorio.ListIndex = 9 Then  ' VBU NULOS
        Rpt.Arvore = False
        If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TVbuZero.rpt") Then Exit Sub
        
        Rpt.Selecao = "isnull({TAB_TRECHO.TTC_VALOR} )" & IIf(Trim(Mid(FiltroRptIptu, 5)) <> "", FiltroRptIptu, "")
        Rpt.Visualizar
    ElseIf cboRelatorio.ListIndex = 4 Then  'VBU
        Rpt.Arvore = False
        FiltroRptIptu = "({TAB_GERACAO_TRIBUTO.TGT_TIP_COD_IMPOSTO} = '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "' or {TAB_GERACAO_TRIBUTO.TGT_TIP_COD_IMPOSTO} = '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "')"
        If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TVbuListagem_Lista.rpt") Then Exit Sub
        
        If Trim(txtCodLogr) <> "" Then
            Rpt.Selecao = "and {TAB_TRECHO.TTC_tlg_cod_logradouro} = '" & txtCodLogr & "'"
        End If
        Rpt.Visualizar
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
        
        If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TIPTUListagemValorZero.rpt") Then Exit Sub
        Rpt.Formulas "FiltroAplicado", Mid(cboCampoZero, 5) & " com valor 0,00"
        Rpt.Selecao = Mid(FiltroRptIptu, 5)
        Rpt.Arvore = False
        Rpt.Visualizar
    ElseIf cboRelatorio.ListIndex = 7 Then ' Estatisticas
        If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TEstatisticaIPTU.rpt") Then Exit Sub
        Rpt.Arvore = False
        Rpt.Visualizar
    End If
End Sub

Private Sub cboRelatorio_Click()
    cboCampoZero.Visible = IIf(cboRelatorio.ListIndex <> 5, False, True)
End Sub

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
    
    txtic = Trim(txtic)
                    
    Select Case cmd(Index).Caption
        Case "&Cota Única"
            If Not Util.Confirma("Gerar IPTU?") Then Exit Sub
            CodIPTU = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
            If AbreSelecao(RsCob) Then
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
                    Do
                        lblIsentos = "Calculando Valor da IC:  " & RsCob!TIM_IC
                        ' SE O ANO FOR MENOR QUE O DO PGV, GERA COM O MODELO PADRÃO ANTIGO DO IPTU
                        If CInt(Nvl(txtAnoFi, Year(Date))) < CInt(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
                            Call Imposto.GeraIptu(cip_Balsas, RsCob, CInt(txtAnoIni), CInt(txtAnoFi), tgi_SemParcelas)
                        Else
                        ' SE O ANO FOR COMPATÍVEL COM O PGV GERA ENTÃO O IPTU PGV
                            If CalculoIptu.CarregaDetalheLote(Trim(RsCob!TIM_IC)) Then
                                CalculoIptu.CalculaValorIptu
                                If CalculoIptu.ValorImposto > 0 Then Obrig.CriaObrigacao CodIPTU, txtAnoIni, txtAnoIni, Trim("" & RsCob!TIM_IC), Format(CalculoIptu.ValorImposto, Const_Monetario), etsCreditoOriginalAberto, , CalculoIptu.VencimentoCotaUnica
                                lblIsentos = "  Finalizando Unidade " & RsCob!TIM_IC
                                lblCont = "  Registros Processados: " & CalculoIptu.RegistrosProcessados
                            Else
                                lblIsentos = "  O lote " & Trim(RsCob!TIM_IC) & " já registra pagamento para este periodo. Regularize esta situação."
                            End If
                        End If
                        DoEvents
                        RsCob.MoveNext
                    Loop While Not RsCob.EOF
                Else
                    lblCont = cadastro.CriaIptu(RsCob, CInt(Nvl(txtAnoIni, Year(Now))), CInt(Nvl(txtAnoIni, Year(Now))))
                End If
            Else
                Call Util.Informa("Nenhum lote encontrado.")
                Screen.MousePointer = 0
                lblIsentos = ""
                Exit Sub
            End If
            prgIptu.Visible = False
            Call Util.Informa("Geração de imposto finalizada.")
            If grdValores.Preencher(Bdados, sqlValores) Then
                grdValores.Mensagem = "Soma : " & Format$(grdValores.Colunas(3).Soma, "currency") & " x Menor : " & Format$(grdValores.Colunas(3).Min, "currency") & " x Maior : " & Format$(grdValores.Colunas(3).Max, "currency") & " x Média : " & Format$(grdValores.Colunas(3).Media, "currency")
            End If
            lblIsentos = ""
            txtic.SetFocus
            Screen.MousePointer = 0
        Case "&Imprimir"
            ImprimeRelatorios
        Case "&Parcelas"
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


Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
            
    Dim Controle As Control
    Dim i As Byte
    Set cadastro = New VSImposto
    Call Edita.AtualizaCombo(Bdados, cboLogr, "Select DISTINCT(tlg_nome) From Tab_Logradouro where tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio)
    Call Edita.AtualizaCombo(Bdados, cboTipoLogr, "Select DISTINCT(ttl_nome) From Tab_Tipo_Logr")
    Call Edita.AtualizaCombo(Bdados, cboBairro, "Select DISTINCT(tba_nome) From Tab_Bairro where TBA_TMU_COD_MUNICIPIO=" & Aplicacoes.Codigo_Municipio)
    'txtMens = Temp.PegaParametro(Bdados, "MENSAGEM IPTU")
    cboLogr.AddItem ""
    cboTipoLogr.AddItem ""
    cboBairro.AddItem ""
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    cboCampoZero.Visible = False
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
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtic = cadastro.FormataInscricao(txtic, InscImovel)
    End If
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
