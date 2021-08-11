VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TIMP401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fra 
      Height          =   645
      Index           =   1
      Left            =   30
      TabIndex        =   25
      Top             =   690
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   1138
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Enabled         =   0   'False
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
      Caption         =   "Período"
      ShadowStyle     =   1
      Begin VB.TextBox txtCodImposto 
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
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   0
         Tag             =   "Código"
         Top             =   210
         Width           =   975
      End
      Begin VB.TextBox txtAnoImposto 
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
         Left            =   8940
         MaxLength       =   4
         TabIndex        =   1
         Top             =   210
         Width           =   465
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   9
         Left            =   90
         TabIndex        =   26
         Top             =   240
         Width           =   1650
         _ExtentX        =   2910
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
         Caption         =   "Código do Tributo:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   12
         Left            =   7650
         TabIndex        =   27
         Top             =   240
         Width           =   1110
         _ExtentX        =   1958
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
         Caption         =   "Ano Imposto:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtNomeImposto 
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
         Left            =   2700
         MaxLength       =   50
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   210
         Width           =   4425
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   195
      Left            =   1890
      TabIndex        =   28
      Top             =   840
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6780
      Top             =   1380
   End
   Begin VTOcx.grdVISUAL lstImposto 
      Height          =   3105
      Left            =   60
      TabIndex        =   52
      Top             =   4890
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5477
      CorTitulo       =   16711680
      CorCaption      =   16777215
   End
   Begin Threed.SSFrame fra 
      Height          =   3435
      Index           =   0
      Left            =   30
      TabIndex        =   29
      Top             =   1380
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6059
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Enabled         =   0   'False
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
      Begin Threed.SSFrame SSFrame1 
         Height          =   765
         Left            =   3090
         TabIndex        =   21
         Top             =   2610
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1349
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Receitas diversas"
         Begin VTOcx.cboVISUAL cboReceitaAMais 
            Height          =   510
            Left            =   150
            TabIndex        =   22
            Top             =   180
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   900
            Caption         =   "Valor a Mais"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
         End
         Begin VTOcx.cboVISUAL cboReceitaAMenos 
            Height          =   510
            Left            =   1950
            TabIndex        =   23
            Top             =   180
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   900
            Caption         =   "Valor a Menos"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
         End
      End
      Begin VB.ComboBox CboJuros 
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
         ItemData        =   "TIMP401.frx":0000
         Left            =   7650
         List            =   "TIMP401.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "Tipo de Juros"
         Top             =   2160
         Width           =   1665
      End
      Begin VB.TextBox txtRedutor 
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
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   20
         Tag             =   "Redutor"
         Top             =   2970
         Width           =   885
      End
      Begin VB.TextBox txtJuros 
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
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   19
         Tag             =   "Juros"
         Top             =   2970
         Width           =   825
      End
      Begin VB.TextBox txtAliquota 
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
         Left            =   120
         MaxLength       =   8
         TabIndex        =   18
         Tag             =   "Aliquota"
         Top             =   2970
         Width           =   885
      End
      Begin VB.TextBox txtDiasPag 
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
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   4
         Tag             =   "Dias Pagar"
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox txtDiasDecl 
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
         Left            =   1890
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "Dias Declarar"
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox txtDtInicio 
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
         Left            =   60
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Data Inicio"
         Top             =   360
         Width           =   1185
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   6
         Left            =   60
         TabIndex        =   36
         Top             =   90
         Width           =   1650
         _ExtentX        =   2910
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
         Caption         =   "Data Início Imposto"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   7
         Left            =   1890
         TabIndex        =   37
         Top             =   90
         Width           =   1650
         _ExtentX        =   2910
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
         Caption         =   "Dias para Declarar"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   8
         Left            =   3600
         TabIndex        =   38
         Top             =   90
         Width           =   1650
         _ExtentX        =   2910
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
         Caption         =   "Dias para Pagar"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSFrame fra 
         Height          =   825
         Index           =   2
         Left            =   60
         TabIndex        =   30
         Top             =   930
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   1455
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
         Caption         =   "Tipo"
         ShadowStyle     =   1
         Begin VTOcx.cboVISUAL cboObrigacao 
            Height          =   510
            Left            =   6990
            TabIndex        =   54
            Top             =   180
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   900
            Caption         =   "Obrigacão"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
         End
         Begin VB.ComboBox cboTipoIC 
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
            ItemData        =   "TIMP401.frx":0025
            Left            =   5280
            List            =   "TIMP401.frx":0027
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Tag             =   "Logradouro"
            Top             =   390
            Width           =   1605
         End
         Begin VB.ComboBox cboTipoInsc 
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
            ItemData        =   "TIMP401.frx":0029
            Left            =   3570
            List            =   "TIMP401.frx":002B
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Tag             =   "Logradouro"
            Top             =   420
            Width           =   1545
         End
         Begin VB.ComboBox cboTipoContrib 
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
            ItemData        =   "TIMP401.frx":002D
            Left            =   90
            List            =   "TIMP401.frx":002F
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "Logradouro"
            Top             =   420
            Width           =   1605
         End
         Begin VB.ComboBox cboTipoTributo 
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
            ItemData        =   "TIMP401.frx":0031
            Left            =   1740
            List            =   "TIMP401.frx":0033
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Tag             =   "Logradouro"
            Top             =   420
            Width           =   1605
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   32
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
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
            Caption         =   "Contribuinte"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   1
            Left            =   1740
            TabIndex        =   33
            Top             =   180
            Width           =   600
            _ExtentX        =   1058
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
            Caption         =   "Tributo"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   10
            Left            =   3570
            TabIndex        =   39
            Top             =   180
            Width           =   1575
            _ExtentX        =   2778
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
            Caption         =   "Inscrição Municipal"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   21
            Left            =   5280
            TabIndex        =   53
            Top             =   150
            Width           =   1575
            _ExtentX        =   2778
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
            Caption         =   "Inscrição Cadastral"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   825
         Index           =   4
         Left            =   60
         TabIndex        =   40
         Top             =   1770
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   1455
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
         Caption         =   "Valor da Multa Mora (%)"
         ShadowStyle     =   1
         Begin VB.TextBox txtValMaxMulta 
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
            Left            =   1380
            MaxLength       =   8
            TabIndex        =   12
            Tag             =   "Valor Máximo"
            Top             =   420
            Width           =   1185
         End
         Begin VB.TextBox txtValMinMulta 
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
            Left            =   120
            MaxLength       =   8
            TabIndex        =   11
            Tag             =   "Valor Mínimo"
            Top             =   420
            Width           =   1065
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   11
            Left            =   120
            TabIndex        =   41
            Top             =   210
            Width           =   1110
            _ExtentX        =   1958
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
            Caption         =   "Valor Mínimo"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   13
            Left            =   1380
            TabIndex        =   42
            Top             =   210
            Width           =   1140
            _ExtentX        =   2011
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
            Caption         =   "Valor Máximo"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   825
         Index           =   5
         Left            =   2790
         TabIndex        =   43
         Top             =   1770
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1455
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
         Caption         =   "Valor Mínimo:"
         ShadowStyle     =   1
         Begin VB.TextBox txtValMinBase 
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
            Left            =   120
            MaxLength       =   8
            TabIndex        =   13
            Tag             =   "Base de Cálculo"
            Top             =   420
            Width           =   1245
         End
         Begin VB.TextBox txtValMinImposto 
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
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   14
            Tag             =   "Imposto"
            Top             =   420
            Width           =   1185
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   16
            Left            =   120
            TabIndex        =   44
            Top             =   210
            Width           =   1320
            _ExtentX        =   2328
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
            Caption         =   "Base de Cálculo"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   17
            Left            =   1530
            TabIndex        =   45
            Top             =   210
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
            Caption         =   "Imposto"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   14
         Left            =   120
         TabIndex        =   46
         Top             =   2730
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "Aliquota(%)"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   15
         Left            =   1170
         TabIndex        =   47
         Top             =   2730
         Width           =   750
         _ExtentX        =   1323
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
         Caption         =   "Juros(%)"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   18
         Left            =   2040
         TabIndex        =   48
         Top             =   2730
         Width           =   930
         _ExtentX        =   1640
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
         Caption         =   "Redutor(%)"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSFrame fraTaxa 
         Height          =   825
         Left            =   5670
         TabIndex        =   49
         Top             =   1770
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1455
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
         Caption         =   "Taxa"
         ShadowStyle     =   1
         Begin Threed.SSCheck chkTaxa 
            Height          =   195
            Left            =   90
            TabIndex        =   15
            Top             =   390
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   344
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Fixa"
         End
         Begin VB.TextBox txtValorTaxa 
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
            Left            =   900
            MaxLength       =   8
            TabIndex        =   16
            Top             =   390
            Width           =   855
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   19
            Left            =   930
            TabIndex        =   50
            Top             =   180
            Width           =   765
            _ExtentX        =   1349
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
            Caption         =   "Valor(R$)"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   20
         Left            =   7650
         TabIndex        =   51
         Top             =   1920
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "Tipo de Juros"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSFrame fra 
         Height          =   885
         Index           =   3
         Left            =   5580
         TabIndex        =   31
         Top             =   30
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   1561
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
         Caption         =   "Periodicidade"
         ShadowStyle     =   1
         Begin VB.ComboBox cboPeriodoCalc 
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
            ItemData        =   "TIMP401.frx":0035
            Left            =   1890
            List            =   "TIMP401.frx":0037
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Tag             =   "Calculo"
            Top             =   480
            Width           =   1815
         End
         Begin VB.ComboBox cboPeriodoDecl 
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
            ItemData        =   "TIMP401.frx":0039
            Left            =   90
            List            =   "TIMP401.frx":003B
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Tag             =   "Declaracao"
            Top             =   480
            Width           =   1605
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   3
            Left            =   60
            TabIndex        =   34
            Top             =   240
            Width           =   945
            _ExtentX        =   1667
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
            Caption         =   " Declaração"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   4
            Left            =   1890
            TabIndex        =   35
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
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
            Caption         =   "Cálculo"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   55
      Top             =   8070
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   1138
      Icone           =   "TIMP401.frx":003D
   End
End
Attribute VB_Name = "TIMP401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub CarregaCombos()
    Call Edita.AtualizaCombo(Bdados, cboTipoContrib, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =1 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboTipoTributo, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =2 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboPeriodoDecl, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =3 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboPeriodoCalc, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =4 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboTipoInsc, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =6 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboTipoIC, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =6 ORDER BY TGE_CODIGO ASC")
End Sub

Private Sub cboTipoTributo_Click()
    If cboTipoTributo = "TAXA" Then
        fraTaxa.Visible = True
    Else
        fraTaxa.Visible = False
    End If
End Sub

Private Sub chkTaxa_Click(Value As Integer)
    If Value Then
        fraTaxa.Visible = True
        txtAliquota = 0
        txtAliquota.Enabled = False
        txtValorTaxa.Enabled = True
    Else
        txtAliquota.Enabled = True
        txtValorTaxa.Enabled = False
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Select Case Index
        Case 0
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Valores = Bdados.PreparaValor(txtCodImposto, txtAnoImposto, Bdados.Converte(txtDtInicio, TCDataHora), txtDiasDecl, _
                    txtDiasPag, cboPeriodoDecl.ListIndex + 1, cboPeriodoCalc.ListIndex + 1, cboTipoContrib.ListIndex + 1, cboTipoTributo.ListIndex + 1, _
                    cboTipoInsc.ListIndex + 1, Bdados.Converte(txtValMinMulta, TCDuplo), Bdados.Converte(txtValMaxMulta, TCDuplo), Bdados.Converte(txtValMinBase, TCDuplo), Bdados.Converte(txtValMinImposto, TCDuplo), _
                    Bdados.Converte(txtAliquota, TCDuplo), Bdados.Converte(txtJuros, TCDuplo), Bdados.Converte(txtRedutor, TCDuplo), Bdados.Converte(IIf(Trim(txtValorTaxa) = "", 0, txtValorTaxa), TCDuplo), CboJuros.ListIndex, cboTipoIC.ListIndex + 1, cboObrigacao.Coluna(1).Valor, Nvl("" & cboReceitaAMais.Coluna(1).Valor, 0), Nvl("" & cboReceitaAMenos.Coluna(1).Valor, 0))
            Campos = "tpi_tip_cod_imposto,tpi_ano_imposto,tpi_dt_inicio_imposto,tpi_dias_declara,tpi_dias_pagar,"
            Campos = Campos & "TPI_PERIODIC_DECLARA,TPI_PERIODIC_CALCULO,tpi_tipo_contribuinte,tpi_tipo_tributo,"
            Campos = Campos & "tpi_tipo_inscricao,tpi_valor_min_multa,tpi_valor_max_multa,"
            Campos = Campos & "tpi_valor_min_base_calc,tpi_valor_min_imposto,"
            Campos = Campos & "tpi_aliquota,tpi_valor_juros,tpi_reducao,tpi_valor_taxa_fixa,TPI_JUROS_CAPTALIZADOS,tpi_tipo_ic,TPI_GERA_OBRIGACAO,TPI_RECEITA_A_MAIS,TPI_RECEITA_A_MENOS"
            If Bdados.GravaDados("Tab_Parametro_Imposto", Valores, Campos, "tpi_tip_cod_imposto='" & txtCodImposto & "'" & IIf(cboTipoTributo = "TAXA", "", " and tpi_ano_imposto='" & txtAnoImposto & "'")) Then
                Call Util.Informa("Transação Completada.")
            End If
            Edita.LimpaCampos Me
            Sql = "SELECT tpi_tip_cod_imposto as Código, tip_sigla_imposto as Sigla,tip_nome_imposto as Imposto, tpi_ano_imposto as Ano FROM Tab_Parametro_Imposto, Tab_Imposto WHERE tpi_tip_cod_imposto=tip_cod_imposto"
            lstImposto.Preencher Bdados, Sql, 1400
        Case 1
            Unload Me
    End Select
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImpostosRelacionados_Click()
    If Trim(txtCodImposto) = "" Then Exit Sub
    TIMP105.CarregaTributos txtCodImposto, txtAnoImposto
    TIMP105.Show 1
End Sub

Private Sub Form_Load()
    Dim Sql As String
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    Call CarregaCombos
    Sql = "SELECT tpi_tip_cod_imposto as Código, tip_sigla_imposto as Sigla,tip_nome_imposto as Imposto, tpi_ano_imposto as Ano FROM Tab_Parametro_Imposto, Tab_Imposto WHERE tpi_tip_cod_imposto=tip_cod_imposto"
    lstImposto.Preencher Bdados, Sql, 2500
    AtualizaCabecalho lstImposto
    cboObrigacao.PreencherGeral Bdados, "LANCAMENTO OBRIGACAO"
    cboReceitaAMais.PreencherGeral Bdados, "RECEITA A MAIS"
    cboReceitaAMenos.PreencherGeral Bdados, "RECEITA A MENOS"
End Sub

Private Sub lstImposto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstImposto, ColumnHeader
End Sub

Private Sub lstImposto_DblClick()
    txtCodImposto = lstImposto.SelectedItem
    txtCodImposto_LostFocus
    txtAnoImposto = lstImposto.SelectedItem.SubItems(3)
    txtAnoImposto_LostFocus
End Sub

Private Sub lstImposto_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    Dim Sql As String
    
    If Confirma("Deseja excluir o tributo " & lstImposto.SelectedItem.SubItems(1) & " ano " & lstImposto.SelectedItem.SubItems(3) & "?") Then
        Bdados.DeletaDados "Tab_Parametro_Imposto", " tpi_tip_cod_imposto =  '" & lstImposto.SelectedItem & "' and tpi_ano_imposto ='" & lstImposto.SelectedItem.SubItems(3) & "'"
        Avisa "Parâmetro excluído com sucesso! "
        txtCodImposto.Enabled = True
        Edita.LimpaCampos Me
        Sql = "SELECT tpi_tip_cod_imposto as Código, tip_sigla_imposto as Sigla,tip_nome_imposto as Imposto, tpi_ano_imposto as Ano FROM Tab_Parametro_Imposto, Tab_Imposto WHERE tpi_tip_cod_imposto=tip_cod_imposto"
        lstImposto.Preencher Bdados, Sql, 2500
    End If
End Sub

Private Sub txtAliquota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtAnoImposto_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtAnoImposto_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    On Error Resume Next
    If Trim(txtCodImposto) = "" Then Exit Sub
    If Trim(Len(Trim(txtAnoImposto))) = 1 Then
        Avisa "Formato do Ano = 'AA'"
        Exit Sub
    End If
    Sql = "SELECT * from Tab_Parametro_Imposto where tpi_tip_cod_imposto ='" & txtCodImposto & "'  and tpi_ano_imposto='" & txtAnoImposto & "'" '& "'" & IIf(Trim(txtAnoImposto) = "", "", " and tpi_ano_imposto='" & txtAnoImposto & "'")
    If Bdados.AbreTabela(Sql, rs) Then
        txtDtInicio = Format(rs!tpi_dt_inicio_imposto, "dd/mm/yyyy")
        txtDiasDecl = rs!tpi_dias_declara
        txtDiasPag = rs!tpi_dias_pagar
        cboPeriodoDecl.ListIndex = rs!TPI_PERIODIC_DECLARA - 1
        cboPeriodoCalc.ListIndex = rs!TPI_PERIODIC_CALCULO - 1
        cboTipoContrib.ListIndex = rs!tpi_tipo_contribuinte - 1
        cboTipoTributo.ListIndex = rs!tpi_tipo_tributo - 1
        cboTipoInsc.ListIndex = rs!tpi_tipo_inscricao - 1
        cboTipoIC.ListIndex = Nvl("" & rs!tpi_tipo_ic - 1, -1)
        txtValMinMulta = rs!tpi_valor_min_multa
        txtValMaxMulta = rs!tpi_valor_max_multa
        txtValMinBase = rs!tpi_valor_min_base_calc
        txtValMinImposto = rs!tpi_valor_min_imposto
        txtAliquota = rs!tpi_aliquota
        txtJuros = rs!tpi_valor_juros
        txtRedutor = rs!tpi_reducao
        cboObrigacao.SetarLinha Nvl("" & rs!TPI_GERA_OBRIGACAO, 0), 1
        cboReceitaAMais.SetarLinha Nvl("" & rs!TPI_RECEITA_A_MAIS, 0), 1
        cboReceitaAMenos.SetarLinha Nvl("" & rs!TPI_RECEITA_A_MENOS, 0), 1
        fraTaxa.Visible = IIf(rs!tpi_tipo_tributo = 2, True, False)
        CboJuros.ListIndex = IIf(IsNull(rs!TPI_JUROS_CAPTALIZADOS), 0, rs!TPI_JUROS_CAPTALIZADOS)
        If Not IsNull(rs!tpi_valor_taxa_fixa) Then
            If rs!tpi_valor_taxa_fixa > 0 Then
                txtValorTaxa = rs!tpi_valor_taxa_fixa
                txtValorTaxa = Edita.FormataTexto(txtValorTaxa, Monetario, True)
                chkTaxa.Value = ssCBChecked
            Else
                 fraTaxa.Visible = False
                txtValorTaxa = ""
                chkTaxa.Value = ssCBUnchecked
            End If
        Else
            txtValorTaxa = ""
            chkTaxa.Value = ssCBUnchecked
        End If
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtCodImposto_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtCodImposto) = "" Then Exit Sub
    If Me.ActiveControl.Name = "cmd" Then Exit Sub
    
    Sql = "Select tip_nome_imposto from Tab_Imposto Where tip_cod_imposto='" & txtCodImposto & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtNomeImposto = rs(0)
    Else
        txtNomeImposto = ""
        Avisa "Código de Tributo inválido."
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtDiasDecl_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDiasPag_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDtInicio_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDtInicio_LostFocus()
    txtDtInicio = Edita.FormataTexto(txtDtInicio, Data)
End Sub

Private Sub txtJuros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtRedutor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValMaxMulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValMaxMulta_LostFocus()
    txtValMaxMulta = Edita.FormataTexto(txtValMaxMulta, Monetario, True)
End Sub

Private Sub txtValMinBase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValMinBase_LostFocus()
    txtValMinBase = Edita.FormataTexto(txtValMinBase, Monetario, True)
End Sub

Private Sub txtValMinImposto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValMinImposto_LostFocus()
    txtValMinImposto = Edita.FormataTexto(txtValMinImposto, Monetario, True)
End Sub

Private Sub txtValMinMulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValMinMulta_LostFocus()
    'txtValMinMulta = Edita.FormataTexto(txtValMinMulta, Monetario,true)
End Sub

Private Sub txtValorTaxa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub


Private Sub txtValorTaxa_LostFocus()
    txtValorTaxa = Edita.FormataTexto(txtValorTaxa, Monetario, True)
End Sub
