VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCIU202 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   243
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCIU202.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin TabDlg.SSTab tabCad 
      Height          =   6015
      Left            =   30
      TabIndex        =   87
      Top             =   690
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Boletim Territorial"
      TabPicture(0)   =   "TCIU202.frx":2123
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(49)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstPesq"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtMotivo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Detalhe BT"
      TabPicture(1)   =   "TCIU202.frx":213F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra(5)"
      Tab(1).Control(1)=   "fra(3)"
      Tab(1).Control(2)=   "fra(4)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Boletim Predial"
      TabPicture(2)   =   "TCIU202.frx":215B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra(6)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Boletim de Condomínio"
      TabPicture(3)   =   "TCIU202.frx":2177
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra(2)"
      Tab(3).Control(1)=   "fra(7)"
      Tab(3).Control(2)=   "fra(8)"
      Tab(3).Control(3)=   "lstCond"
      Tab(3).ControlCount=   4
      Begin VB.TextBox txtMotivo 
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
         Left            =   2010
         TabIndex        =   35
         Tag             =   "Motivo"
         Top             =   5640
         Width           =   9150
      End
      Begin Threed.SSFrame fra 
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   88
         Top             =   480
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   2566
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
            Left            =   5370
            TabIndex        =   244
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox txtLote 
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
            Left            =   3270
            MaxLength       =   5
            TabIndex        =   13
            Top             =   1020
            Width           =   615
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
            Left            =   2130
            MaxLength       =   5
            TabIndex        =   12
            Top             =   990
            Width           =   555
         End
         Begin VB.TextBox txtLoteamento 
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
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   11
            Top             =   990
            Width           =   405
         End
         Begin VB.TextBox txtComplemento 
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
            Left            =   9600
            TabIndex        =   10
            Top             =   660
            Width           =   1425
         End
         Begin VB.TextBox txtNumero 
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
            Left            =   8310
            MaxLength       =   10
            TabIndex        =   9
            Top             =   660
            Width           =   555
         End
         Begin VB.ComboBox cboTipoImovel 
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
            ItemData        =   "TCIU202.frx":2193
            Left            =   540
            List            =   "TCIU202.frx":219D
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "Tipo Imovel"
            Top             =   600
            Width           =   1455
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
            Left            =   2910
            TabIndex        =   8
            Tag             =   "Logradouro"
            Top             =   630
            Width           =   1485
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
            Left            =   5130
            TabIndex        =   91
            Tag             =   "Nome Contribuinte"
            Top             =   990
            Width           =   2865
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
            Left            =   5550
            TabIndex        =   90
            Tag             =   "Nome Contribuinte"
            Top             =   630
            Width           =   2445
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
            Left            =   4440
            MaxLength       =   11
            TabIndex        =   89
            Top             =   630
            Width           =   1035
         End
         Begin VB.TextBox txtCodBairro 
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
            Left            =   4560
            MaxLength       =   50
            TabIndex        =   14
            Tag             =   "Bairro"
            Top             =   990
            Width           =   525
         End
         Begin VB.TextBox txtCodMens 
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
            Left            =   10710
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Cod Mensagem"
            Top             =   1020
            Width           =   315
         End
         Begin VB.TextBox txtIc 
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
            Index           =   0
            Left            =   1890
            MaxLength       =   2
            TabIndex        =   0
            Tag             =   "Distrito"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtIc 
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
            Index           =   1
            Left            =   2220
            MaxLength       =   2
            TabIndex        =   1
            Tag             =   "Setor"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtIc 
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
            Index           =   2
            Left            =   2550
            MaxLength       =   4
            TabIndex        =   2
            Tag             =   "Quadra"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtIc 
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
            Index           =   3
            Left            =   3060
            MaxLength       =   4
            TabIndex        =   3
            Tag             =   "Lote"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtIc 
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
            Index           =   4
            Left            =   3570
            MaxLength       =   3
            TabIndex        =   4
            Tag             =   "Unidade"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtIcAnterior 
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
            Left            =   8250
            TabIndex        =   5
            Top             =   240
            Width           =   1575
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
            Height          =   315
            Left            =   10470
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Zona"
            Top             =   210
            Width           =   555
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   2
            Left            =   8910
            TabIndex        =   92
            Top             =   720
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
            TabIndex        =   93
            Top             =   720
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
            Left            =   3960
            TabIndex        =   94
            Top             =   1050
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
            Height          =   180
            Index           =   4
            Left            =   90
            TabIndex        =   95
            Top             =   1020
            Width           =   1170
            _ExtentX        =   2064
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
            Caption         =   "Loteamento:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   5
            Left            =   1710
            TabIndex        =   96
            Top             =   1050
            Width           =   660
            _ExtentX        =   1164
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
            Caption         =   "Qd.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   6
            Left            =   2820
            TabIndex        =   97
            Top             =   1050
            Width           =   435
            _ExtentX        =   767
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
            Caption         =   "Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   7
            Left            =   90
            TabIndex        =   98
            Top             =   660
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
            Height          =   210
            Index           =   82
            Left            =   2010
            TabIndex        =   99
            Top             =   690
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
            Caption         =   "Cód. Logr:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   50
            Left            =   9630
            TabIndex        =   100
            Top             =   1080
            Width           =   990
            _ExtentX        =   1746
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
            Caption         =   "Cod. Mens.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   72
            Left            =   6960
            TabIndex        =   101
            Top             =   270
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
            Height          =   180
            Index           =   75
            Left            =   510
            TabIndex        =   102
            Top             =   285
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "Insc. Imobiliária:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   76
            Left            =   9990
            TabIndex        =   103
            Top             =   270
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
            Caption         =   "Zona:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   91
            Left            =   4020
            TabIndex        =   245
            Top             =   285
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
      Begin Threed.SSFrame fra 
         Height          =   1785
         Index           =   1
         Left            =   120
         TabIndex        =   104
         Top             =   2550
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   3149
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
         Caption         =   "Dados do Proprietário"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtUf 
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
            Left            =   10425
            MaxLength       =   50
            TabIndex        =   32
            Tag             =   "Bairro"
            Top             =   975
            Width           =   585
         End
         Begin VB.TextBox txtCodLogrContrib 
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
            Left            =   1020
            TabIndex        =   224
            Tag             =   "Logradouro"
            Top             =   585
            Width           =   645
         End
         Begin VB.TextBox txtNomeTipoLogrContrib 
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
            Left            =   2085
            MaxLength       =   11
            TabIndex        =   223
            Top             =   570
            Width           =   1035
         End
         Begin VB.TextBox txtCompContrib 
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
            Left            =   7560
            TabIndex        =   28
            Top             =   600
            Width           =   735
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
            Height          =   315
            Left            =   1470
            TabIndex        =   23
            Top             =   210
            Width           =   1305
         End
         Begin VB.TextBox txtNomeContrib 
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
            Left            =   3660
            TabIndex        =   24
            Tag             =   "Nome Contribuinte"
            Top             =   210
            Width           =   4665
         End
         Begin VB.TextBox txtNomeLogrContrib 
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
            Left            =   3225
            TabIndex        =   26
            Tag             =   "Nome Logradouro"
            Top             =   570
            Width           =   2370
         End
         Begin VB.TextBox txtCep 
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
            MaxLength       =   10
            TabIndex        =   30
            Top             =   960
            Width           =   1125
         End
         Begin VB.CommandButton cmdEnter 
            Caption         =   "Command1"
            Default         =   -1  'True
            Height          =   255
            Left            =   7740
            TabIndex        =   105
            Top             =   3090
            Width           =   375
         End
         Begin VB.TextBox txtBairroContrib 
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
            Left            =   8970
            TabIndex        =   29
            Tag             =   "Bairro"
            Top             =   630
            Width           =   2055
         End
         Begin VB.TextBox txtOcupante 
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
            TabIndex        =   33
            Top             =   1350
            Width           =   4965
         End
         Begin VB.TextBox txtCpfOcupante 
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
            Left            =   8970
            MaxLength       =   20
            TabIndex        =   34
            Top             =   1350
            Width           =   2055
         End
         Begin VB.TextBox txtCpfCgc 
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
            Left            =   9240
            MaxLength       =   20
            TabIndex        =   25
            Top             =   210
            Width           =   1785
         End
         Begin VB.TextBox txtNumeroContrib 
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
            Left            =   6180
            MaxLength       =   10
            TabIndex        =   27
            Top             =   600
            Width           =   525
         End
         Begin VB.TextBox txtMunic 
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
            Left            =   5880
            TabIndex        =   31
            Tag             =   "Município"
            Top             =   990
            Width           =   4335
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   8
            Left            =   150
            TabIndex        =   106
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
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
            Caption         =   "Insc. Municipal:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   14
            Left            =   2640
            TabIndex        =   107
            Top             =   990
            Width           =   375
            _ExtentX        =   661
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
            Caption         =   "CEP:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   15
            Left            =   4995
            TabIndex        =   108
            Top             =   1020
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Municipio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   13
            Left            =   5805
            TabIndex        =   109
            Top             =   615
            Width           =   270
            _ExtentX        =   476
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
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   17
            Left            =   8370
            TabIndex        =   110
            Top             =   660
            Width           =   555
            _ExtentX        =   979
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
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   16
            Left            =   6810
            TabIndex        =   111
            Top             =   630
            Width           =   1290
            _ExtentX        =   2275
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
            Caption         =   "Compl.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   29
            Left            =   8385
            TabIndex        =   112
            Top             =   240
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   11
            Left            =   2175
            TabIndex        =   113
            Top             =   1380
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "Ocupante:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   330
            Index           =   18
            Left            =   8085
            TabIndex        =   114
            Top             =   1380
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   12
            Left            =   120
            TabIndex        =   225
            Top             =   645
            Width           =   870
            _ExtentX        =   1535
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
            Caption         =   "Cód. Logr:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdNovo 
            Height          =   285
            Left            =   3120
            TabIndex        =   232
            Top             =   240
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            Caption         =   ""
            Acao            =   6
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   285
            Index           =   0
            Left            =   2790
            TabIndex        =   233
            Top             =   240
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   315
            Index           =   3
            Left            =   1680
            TabIndex        =   234
            Top             =   570
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin MSComctlLib.ListView lstPesq 
         Height          =   1155
         Left            =   90
         TabIndex        =   115
         Top             =   4440
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   2037
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   2540
         EndProperty
      End
      Begin Threed.SSFrame fra 
         Height          =   1395
         Index           =   5
         Left            =   -74880
         TabIndex        =   116
         Top             =   480
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   2461
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
         Caption         =   "Características do Imóvel:"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.ComboBox cboInstSanit17 
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
            ItemData        =   "TCIU202.frx":21B7
            Left            =   6930
            List            =   "TCIU202.frx":21B9
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   1950
            Width           =   3015
         End
         Begin VB.ComboBox cboInstElet18 
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
            ItemData        =   "TCIU202.frx":21BB
            Left            =   7410
            List            =   "TCIU202.frx":21BD
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   2310
            Width           =   2535
         End
         Begin VB.ComboBox cboArborizacao 
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
            ItemData        =   "TCIU202.frx":21BF
            Left            =   7680
            List            =   "TCIU202.frx":21C1
            Style           =   2  'Dropdown List
            TabIndex        =   121
            TabStop         =   0   'False
            Tag             =   "5"
            Top             =   630
            Width           =   3375
         End
         Begin VB.ComboBox cboLimites 
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
            ItemData        =   "TCIU202.frx":21C3
            Left            =   7680
            List            =   "TCIU202.frx":21C5
            Style           =   2  'Dropdown List
            TabIndex        =   120
            TabStop         =   0   'False
            Tag             =   "4"
            Top             =   270
            Width           =   3375
         End
         Begin VB.ComboBox cboCobranca 
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
            ItemData        =   "TCIU202.frx":21C7
            Left            =   2070
            List            =   "TCIU202.frx":21C9
            Style           =   2  'Dropdown List
            TabIndex        =   119
            TabStop         =   0   'False
            Tag             =   "3"
            Top             =   990
            Width           =   3375
         End
         Begin VB.ComboBox cboPatrimonio 
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
            ItemData        =   "TCIU202.frx":21CB
            Left            =   2070
            List            =   "TCIU202.frx":21CD
            Style           =   2  'Dropdown List
            TabIndex        =   118
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   630
            Width           =   3375
         End
         Begin VB.ComboBox cboOcupLote 
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
            ItemData        =   "TCIU202.frx":21CF
            Left            =   2070
            List            =   "TCIU202.frx":21D1
            Style           =   2  'Dropdown List
            TabIndex        =   117
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   270
            Width           =   3375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   0
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   36
            Top             =   270
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   1
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   37
            Top             =   630
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   2
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   38
            Top             =   990
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   3
            Left            =   7260
            MaxLength       =   3
            TabIndex        =   39
            Top             =   270
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   4
            Left            =   7260
            MaxLength       =   3
            TabIndex        =   40
            Top             =   630
            Width           =   375
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   30
            Left            =   6150
            TabIndex        =   124
            Top             =   600
            Width           =   1050
            _ExtentX        =   1852
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
            Caption         =   "Arborização:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   32
            Left            =   6480
            TabIndex        =   125
            Top             =   300
            Width           =   690
            _ExtentX        =   1217
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
            Caption         =   "Limites:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   33
            Left            =   330
            TabIndex        =   126
            Top             =   1020
            Width           =   1260
            _ExtentX        =   2223
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
            Caption         =   "Cod. Cobrança:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   34
            Left            =   60
            TabIndex        =   127
            Top             =   300
            Width           =   1530
            _ExtentX        =   2699
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
            Caption         =   "Ocupação do Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   35
            Left            =   630
            TabIndex        =   128
            Top             =   660
            Width           =   960
            _ExtentX        =   1693
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
            Caption         =   "Patrimônio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   31
            Left            =   5190
            TabIndex        =   129
            Top             =   2010
            Width           =   1680
            _ExtentX        =   2963
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
            Caption         =   "Instalação Sanitária:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   37
            Left            =   5700
            TabIndex        =   130
            Top             =   2370
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
            Caption         =   "Instalação Elétrica:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   975
         Index           =   3
         Left            =   -74910
         TabIndex        =   131
         Top             =   1860
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   1720
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
         Caption         =   "Características do Terreno:"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.ComboBox cboSit 
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
            ItemData        =   "TCIU202.frx":21D3
            Left            =   2070
            List            =   "TCIU202.frx":21D5
            Style           =   2  'Dropdown List
            TabIndex        =   134
            TabStop         =   0   'False
            Tag             =   "7"
            Top             =   600
            Width           =   3375
         End
         Begin VB.ComboBox cboPedol 
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
            ItemData        =   "TCIU202.frx":21D7
            Left            =   7650
            List            =   "TCIU202.frx":21D9
            Style           =   2  'Dropdown List
            TabIndex        =   133
            TabStop         =   0   'False
            Tag             =   "8"
            Top             =   240
            Width           =   3405
         End
         Begin VB.ComboBox cboTopogr 
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
            ItemData        =   "TCIU202.frx":21DB
            Left            =   2070
            List            =   "TCIU202.frx":21DD
            Style           =   2  'Dropdown List
            TabIndex        =   132
            TabStop         =   0   'False
            Tag             =   "6"
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   5
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   41
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   7
            Left            =   7260
            MaxLength       =   3
            TabIndex        =   43
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   6
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   42
            Top             =   570
            Width           =   375
         End
         Begin Threed.SSPanel lbl 
            Height          =   315
            Index           =   20
            Left            =   690
            TabIndex        =   135
            Top             =   600
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   556
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
            Caption         =   "Topografia:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   21
            Left            =   870
            TabIndex        =   136
            Top             =   210
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
            Caption         =   "Situação:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   240
            Index           =   22
            Left            =   6330
            TabIndex        =   137
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
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
            Caption         =   "Pedologia:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1005
         Index           =   4
         Left            =   -74910
         TabIndex        =   138
         Top             =   2820
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   1773
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
         Caption         =   "Dimensões do Terreno (m²)"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtTestadaCampo 
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
            Left            =   8280
            TabIndex        =   51
            Tag             =   "107"
            Top             =   510
            Width           =   735
         End
         Begin VB.TextBox txtAreaLote 
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
            Left            =   10350
            TabIndex        =   52
            Tag             =   "108"
            Top             =   210
            Width           =   735
         End
         Begin VB.TextBox txtTestada4 
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
            Left            =   5880
            TabIndex        =   49
            Tag             =   "105"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtTrechoLogr3 
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
            Left            =   5880
            TabIndex        =   48
            Tag             =   "104"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtTrechoLogr4 
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
            Left            =   8280
            TabIndex        =   50
            Tag             =   "106"
            Top             =   180
            Width           =   735
         End
         Begin VB.TextBox txtTestada3 
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
            Left            =   3780
            TabIndex        =   47
            Tag             =   "103"
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtTrechoLogr2 
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
            Left            =   3780
            TabIndex        =   46
            Tag             =   "102"
            Top             =   210
            Width           =   735
         End
         Begin VB.TextBox txtTestada2 
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
            Left            =   1650
            TabIndex        =   45
            Tag             =   "101"
            Top             =   570
            Width           =   735
         End
         Begin VB.TextBox txtTestadaPrin 
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
            Left            =   1650
            TabIndex        =   44
            Tag             =   "100"
            Top             =   240
            Width           =   735
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   23
            Left            =   4590
            TabIndex        =   139
            Top             =   240
            Width           =   1260
            _ExtentX        =   2223
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
            Caption         =   "Trecho Logr. 3:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   24
            Left            =   6930
            TabIndex        =   140
            Top             =   240
            Width           =   1260
            _ExtentX        =   2223
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
            Caption         =   "Trecho Logr. 4:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   330
            Index           =   25
            Left            =   2910
            TabIndex        =   141
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
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
            Caption         =   "Testada 3:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   26
            Left            =   2490
            TabIndex        =   142
            Top             =   270
            Width           =   1260
            _ExtentX        =   2223
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
            Caption         =   "Trecho Logr. 2:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   27
            Left            =   120
            TabIndex        =   143
            Top             =   240
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
            Caption         =   "Testada Principal:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   28
            Left            =   720
            TabIndex        =   144
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Testada 2:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   19
            Left            =   4950
            TabIndex        =   145
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Testada 4:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   45
            Left            =   6780
            TabIndex        =   146
            Top             =   600
            Width           =   1425
            _ExtentX        =   2514
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
            Caption         =   "Testada(Campo):"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   51
            Left            =   9210
            TabIndex        =   147
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
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
            Caption         =   "Área do Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   5325
         Index           =   6
         Left            =   -74880
         TabIndex        =   148
         Top             =   480
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   9393
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
         Caption         =   "Características das Edificações"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.ComboBox cboConservacao 
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
            ItemData        =   "TCIU202.frx":21DF
            Left            =   2040
            List            =   "TCIU202.frx":21E1
            Style           =   2  'Dropdown List
            TabIndex        =   160
            TabStop         =   0   'False
            Tag             =   "13"
            Top             =   2580
            Width           =   3615
         End
         Begin VB.ComboBox cboPadrao 
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
            ItemData        =   "TCIU202.frx":21E3
            Left            =   2040
            List            =   "TCIU202.frx":21E5
            Style           =   2  'Dropdown List
            TabIndex        =   159
            TabStop         =   0   'False
            Tag             =   "12"
            Top             =   2220
            Width           =   3615
         End
         Begin VB.ComboBox cboTipologia 
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
            ItemData        =   "TCIU202.frx":21E7
            Left            =   7440
            List            =   "TCIU202.frx":21E9
            Style           =   2  'Dropdown List
            TabIndex        =   158
            TabStop         =   0   'False
            Tag             =   "9"
            Top             =   1080
            Width           =   3615
         End
         Begin VB.ComboBox cboDestinacao 
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
            ItemData        =   "TCIU202.frx":21EB
            Left            =   2040
            List            =   "TCIU202.frx":21ED
            Style           =   2  'Dropdown List
            TabIndex        =   157
            TabStop         =   0   'False
            Tag             =   "11"
            Top             =   1860
            Width           =   3615
         End
         Begin VB.ComboBox cboEstrutura 
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
            ItemData        =   "TCIU202.frx":21EF
            Left            =   7440
            List            =   "TCIU202.frx":21F1
            Style           =   2  'Dropdown List
            TabIndex        =   156
            TabStop         =   0   'False
            Tag             =   "10"
            Top             =   1440
            Width           =   3615
         End
         Begin VB.TextBox txtAreaEdif 
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
            Left            =   7050
            TabIndex        =   65
            Tag             =   "112"
            Top             =   2280
            Width           =   1185
         End
         Begin VB.TextBox txtFracaoEdif 
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
            Left            =   9870
            TabIndex        =   66
            Tag             =   "114"
            Top             =   2280
            Width           =   1155
         End
         Begin VB.ComboBox cboSentido 
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
            ItemData        =   "TCIU202.frx":21F3
            Left            =   2040
            List            =   "TCIU202.frx":21F5
            Style           =   2  'Dropdown List
            TabIndex        =   155
            TabStop         =   0   'False
            Tag             =   "14"
            Top             =   720
            Width           =   3615
         End
         Begin VB.ComboBox cboUso 
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
            ItemData        =   "TCIU202.frx":21F7
            Left            =   7440
            List            =   "TCIU202.frx":21F9
            Style           =   2  'Dropdown List
            TabIndex        =   154
            TabStop         =   0   'False
            Tag             =   "16"
            Top             =   720
            Width           =   3615
         End
         Begin VB.ComboBox cboPredio 
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
            ItemData        =   "TCIU202.frx":21FB
            Left            =   2040
            List            =   "TCIU202.frx":21FD
            Style           =   2  'Dropdown List
            TabIndex        =   153
            TabStop         =   0   'False
            Tag             =   "15"
            Top             =   1080
            Width           =   3615
         End
         Begin VB.TextBox txtPavimento 
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
            Left            =   1650
            TabIndex        =   56
            Tag             =   "110"
            Top             =   1440
            Width           =   825
         End
         Begin VB.TextBox txtInscImobiliaria 
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
            Left            =   3330
            TabIndex        =   53
            Top             =   210
            Width           =   735
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   15
            Left            =   7050
            MaxLength       =   3
            TabIndex        =   57
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   14
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   55
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   13
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   54
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   8
            Left            =   7050
            MaxLength       =   3
            TabIndex        =   58
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   10
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   60
            Top             =   1860
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   12
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   62
            Top             =   2580
            Width           =   375
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   11
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   61
            Top             =   2220
            Width           =   375
         End
         Begin VB.TextBox txtAreaEdifTotal 
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
            Left            =   9870
            TabIndex        =   64
            Tag             =   "113"
            Top             =   1860
            Width           =   1155
         End
         Begin VB.TextBox txtAnoConst 
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
            Left            =   7050
            TabIndex        =   63
            Tag             =   "111"
            Top             =   1860
            Width           =   1185
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   9
            Left            =   7050
            MaxLength       =   3
            TabIndex        =   59
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtIc 
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
            Index           =   10
            Left            =   1650
            MaxLength       =   2
            TabIndex        =   152
            Top             =   210
            Width           =   315
         End
         Begin VB.TextBox txtIc 
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
            Index           =   11
            Left            =   1980
            MaxLength       =   2
            TabIndex        =   151
            Top             =   210
            Width           =   315
         End
         Begin VB.TextBox txtIc 
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
            Index           =   12
            Left            =   2310
            MaxLength       =   4
            TabIndex        =   150
            Top             =   210
            Width           =   495
         End
         Begin VB.TextBox txtIc 
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
            Index           =   13
            Left            =   2820
            MaxLength       =   4
            TabIndex        =   149
            Top             =   210
            Width           =   495
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   36
            Left            =   6045
            TabIndex        =   161
            Top             =   1500
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "Estrutura:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   38
            Left            =   6060
            TabIndex        =   162
            Top             =   1080
            Width           =   810
            _ExtentX        =   1429
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
            Caption         =   "Tipologia:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   240
            Index           =   39
            Left            =   510
            TabIndex        =   163
            Top             =   1920
            Width           =   960
            _ExtentX        =   1693
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
            Caption         =   "Destinação:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   40
            Left            =   840
            TabIndex        =   164
            Top             =   2280
            Width           =   630
            _ExtentX        =   1111
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
            Caption         =   "Padrão:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   240
            Index           =   41
            Left            =   330
            TabIndex        =   165
            Top             =   2670
            Width           =   1140
            _ExtentX        =   2011
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
            Caption         =   "Conservação:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   42
            Left            =   5820
            TabIndex        =   166
            Top             =   1905
            Width           =   1050
            _ExtentX        =   1852
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
            Caption         =   "Ano Constr.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   44
            Left            =   8505
            TabIndex        =   167
            Top             =   1905
            Width           =   1290
            _ExtentX        =   2275
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
            Caption         =   "Área Edif. Total:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   47
            Left            =   780
            TabIndex        =   168
            Top             =   720
            Width           =   690
            _ExtentX        =   1217
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
            Caption         =   "Sentido:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   48
            Left            =   6495
            TabIndex        =   169
            Top             =   750
            Width           =   375
            _ExtentX        =   661
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
            Caption         =   "Uso:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   52
            Left            =   870
            TabIndex        =   170
            Top             =   1080
            Width           =   600
            _ExtentX        =   1058
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
            Caption         =   "Predio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   53
            Left            =   435
            TabIndex        =   171
            Top             =   1470
            Width           =   1035
            _ExtentX        =   1826
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
            Caption         =   "Pavimentos:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin MSComctlLib.ListView lstEdific 
            Height          =   1815
            Left            =   90
            TabIndex        =   172
            Top             =   3420
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   3201
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   14
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Insc. Imobiliária"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "14"
               Text            =   "Sentido"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "15"
               Text            =   "Prédio"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "110"
               Text            =   "Pavimentos"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "16"
               Text            =   "Uso"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "9"
               Text            =   "TipoLogia"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "10"
               Text            =   "Estrutura"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "11"
               Text            =   "Destinação"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "12"
               Text            =   "Padrão"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Tag             =   "13"
               Text            =   "Conservação"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Object.Tag             =   "111"
               Text            =   "Área Constr."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Object.Tag             =   "112"
               Text            =   "Área Edificada"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Object.Tag             =   "113"
               Text            =   "Área Edificada Total"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Object.Tag             =   "114"
               Text            =   "Fração Ideal"
               Object.Width           =   2540
            EndProperty
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   9
            Left            =   120
            TabIndex        =   173
            Top             =   270
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Insc. Imobiliária:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   43
            Left            =   6030
            TabIndex        =   174
            Top             =   2302
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "Área Edif.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   46
            Left            =   8760
            TabIndex        =   175
            Top             =   2302
            Width           =   1035
            _ExtentX        =   1826
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
            Caption         =   "Fração Ideal:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdAdEdif 
            Height          =   375
            Left            =   120
            TabIndex        =   235
            Top             =   3000
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   661
            Caption         =   "&Adicionar Edificação"
            Acao            =   1
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1455
         Index           =   2
         Left            =   -74880
         TabIndex        =   176
         Top             =   480
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   2566
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
         Begin VB.TextBox txtBairro 
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
            Left            =   5610
            TabIndex        =   188
            Tag             =   "Nome Contribuinte"
            Top             =   1020
            Width           =   3675
         End
         Begin VB.TextBox txtNomeLogr 
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
            Left            =   3870
            TabIndex        =   187
            Tag             =   "Nome Contribuinte"
            Top             =   660
            Width           =   3255
         End
         Begin VB.TextBox txtLogr 
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
            Left            =   2790
            MaxLength       =   11
            TabIndex        =   186
            Top             =   660
            Width           =   1035
         End
         Begin VB.TextBox txtCodLogrBc 
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
            Left            =   1260
            TabIndex        =   185
            Top             =   660
            Width           =   1485
         End
         Begin VB.TextBox txtInscAnteriorBC 
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
            Left            =   6105
            TabIndex        =   68
            Top             =   270
            Width           =   1665
         End
         Begin VB.TextBox txtIc 
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
            Index           =   9
            Left            =   4335
            MaxLength       =   3
            TabIndex        =   67
            Tag             =   "Unidade"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtIc 
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
            Index           =   8
            Left            =   3315
            MaxLength       =   4
            TabIndex        =   184
            Tag             =   "Lote"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtIc 
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
            Index           =   7
            Left            =   2325
            MaxLength       =   4
            TabIndex        =   183
            Tag             =   "Quadra"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtIc 
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
            Index           =   6
            Left            =   1275
            MaxLength       =   2
            TabIndex        =   182
            Tag             =   "Setor"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtIc 
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
            Index           =   5
            Left            =   555
            MaxLength       =   2
            TabIndex        =   181
            Tag             =   "Distrito"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtLoteBc 
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
            Left            =   4110
            MaxLength       =   5
            TabIndex        =   180
            Top             =   1020
            Width           =   765
         End
         Begin VB.TextBox txtQuadraBc 
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
            MaxLength       =   5
            TabIndex        =   179
            Top             =   990
            Width           =   705
         End
         Begin VB.TextBox txtLoteamentoBc 
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
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   178
            Top             =   990
            Width           =   705
         End
         Begin VB.TextBox txtComplementoBc 
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
            Left            =   8670
            TabIndex        =   71
            Top             =   660
            Width           =   2355
         End
         Begin VB.TextBox txtNumeroBc 
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
            Left            =   7470
            MaxLength       =   10
            TabIndex        =   70
            Top             =   690
            Width           =   525
         End
         Begin VB.TextBox txtCepImBc 
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
            Left            =   9990
            MaxLength       =   10
            TabIndex        =   177
            Top             =   1020
            Width           =   1035
         End
         Begin VB.ComboBox cboTipoImovelBc 
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
            ItemData        =   "TCIU202.frx":21FF
            Left            =   9480
            List            =   "TCIU202.frx":2209
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Tag             =   "Logradouro"
            Top             =   255
            Width           =   1545
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   10
            Left            =   90
            TabIndex        =   189
            Top             =   720
            Width           =   1080
            _ExtentX        =   1905
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
            Caption         =   "Cod. Logr:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   54
            Left            =   8010
            TabIndex        =   190
            Top             =   720
            Width           =   660
            _ExtentX        =   1164
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
            Caption         =   "Compl.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   55
            Left            =   7200
            TabIndex        =   191
            Top             =   720
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
            Index           =   56
            Left            =   5010
            TabIndex        =   192
            Top             =   1050
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
            Height          =   180
            Index           =   57
            Left            =   90
            TabIndex        =   193
            Top             =   1020
            Width           =   1170
            _ExtentX        =   2064
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
            Caption         =   "Loteamento:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   58
            Left            =   2040
            TabIndex        =   194
            Top             =   1050
            Width           =   660
            _ExtentX        =   1164
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
            Caption         =   "Quadra:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   59
            Left            =   3600
            TabIndex        =   195
            Top             =   1050
            Width           =   750
            _ExtentX        =   1323
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
            Caption         =   "Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   60
            Left            =   9570
            TabIndex        =   196
            Top             =   1080
            Width           =   750
            _ExtentX        =   1323
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
            Caption         =   "CEP:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   300
            Index           =   61
            Left            =   9015
            TabIndex        =   197
            Top             =   330
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   529
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
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   198
            Top             =   270
            Width           =   375
            _ExtentX        =   661
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
            Caption         =   "Dist:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   73
            Left            =   945
            TabIndex        =   199
            Top             =   270
            Width           =   330
            _ExtentX        =   582
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
            Caption         =   "Set:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   83
            Left            =   1665
            TabIndex        =   200
            Top             =   270
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
            Caption         =   "Quadra:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   84
            Left            =   2835
            TabIndex        =   201
            Top             =   270
            Width           =   435
            _ExtentX        =   767
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
            Caption         =   "Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   85
            Left            =   3915
            TabIndex        =   202
            Top             =   270
            Width           =   420
            _ExtentX        =   741
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
            Caption         =   "Unid:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   86
            Left            =   4845
            TabIndex        =   203
            Top             =   300
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
         Begin VB.Shape Shape2 
            Height          =   405
            Left            =   105
            Top             =   210
            Width           =   4665
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1785
         Index           =   7
         Left            =   -74880
         TabIndex        =   204
         Top             =   1950
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   3149
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
         Caption         =   "Dados do Proprietário"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtCompContribBc 
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
            Left            =   7290
            TabIndex        =   78
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtIMBc 
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
            MaxLength       =   11
            TabIndex        =   72
            Top             =   210
            Width           =   1305
         End
         Begin VB.TextBox txtNomeContribBc 
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
            TabIndex        =   73
            Tag             =   "Nome Contribuinte"
            Top             =   210
            Width           =   4965
         End
         Begin VB.TextBox txtNomeLogrContribBc 
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
            TabIndex        =   76
            Tag             =   "Nome Logradouro"
            Top             =   570
            Width           =   2415
         End
         Begin VB.TextBox txtCepBc 
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
            MaxLength       =   10
            TabIndex        =   80
            Top             =   960
            Width           =   1125
         End
         Begin VB.TextBox txtBairroContribBc 
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
            Left            =   8955
            TabIndex        =   79
            Tag             =   "Bairro"
            Top             =   630
            Width           =   2055
         End
         Begin VB.ComboBox cboTipoLogrContribBc 
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
            ItemData        =   "TCIU202.frx":2223
            Left            =   1470
            List            =   "TCIU202.frx":2230
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Tag             =   "Logradouro"
            Top             =   570
            Width           =   1365
         End
         Begin VB.TextBox txtOcupanteBc 
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
            TabIndex        =   83
            Top             =   1365
            Width           =   4965
         End
         Begin VB.TextBox txtCpfOcupanteBc 
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
            Left            =   8955
            MaxLength       =   20
            TabIndex        =   84
            Top             =   1350
            Width           =   2055
         End
         Begin VB.TextBox txtCpfCgcBc 
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
            Left            =   8955
            MaxLength       =   20
            TabIndex        =   74
            Top             =   210
            Width           =   2055
         End
         Begin VB.ComboBox cboUFBc 
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
            ItemData        =   "TCIU202.frx":2251
            Left            =   10215
            List            =   "TCIU202.frx":225E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Tag             =   "UF"
            Top             =   990
            Width           =   795
         End
         Begin VB.TextBox txtNumeroContribBc 
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
            Left            =   5880
            MaxLength       =   10
            TabIndex        =   77
            Top             =   600
            Width           =   525
         End
         Begin VB.TextBox txtMunicBc 
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
            Left            =   5880
            TabIndex        =   81
            Tag             =   "Município"
            Top             =   1005
            Width           =   4275
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   62
            Left            =   150
            TabIndex        =   205
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
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
            Caption         =   "Insc. Municipal:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   210
            Index           =   63
            Left            =   390
            TabIndex        =   206
            Top             =   600
            Width           =   1035
            _ExtentX        =   1826
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
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   64
            Left            =   2640
            TabIndex        =   207
            Top             =   990
            Width           =   375
            _ExtentX        =   661
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
            Caption         =   "CEP:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   65
            Left            =   4995
            TabIndex        =   208
            Top             =   1020
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Municipio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   66
            Left            =   5580
            TabIndex        =   209
            Top             =   600
            Width           =   270
            _ExtentX        =   476
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
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   67
            Left            =   8370
            TabIndex        =   210
            Top             =   660
            Width           =   555
            _ExtentX        =   979
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
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   68
            Left            =   6510
            TabIndex        =   211
            Top             =   630
            Width           =   1290
            _ExtentX        =   2275
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
            Caption         =   "Compl.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   69
            Left            =   8085
            TabIndex        =   212
            Top             =   240
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   70
            Left            =   2175
            TabIndex        =   213
            Top             =   1380
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "Ocupante:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   330
            Index           =   71
            Left            =   8085
            TabIndex        =   214
            Top             =   1410
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   615
         Index           =   8
         Left            =   -74880
         TabIndex        =   215
         Top             =   3750
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   1085
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
         Caption         =   "Características do Imóvel:"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.ComboBox Combo13 
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
            ItemData        =   "TCIU202.frx":227F
            Left            =   6930
            List            =   "TCIU202.frx":2281
            Style           =   2  'Dropdown List
            TabIndex        =   218
            Top             =   1950
            Width           =   3015
         End
         Begin VB.ComboBox Combo12 
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
            ItemData        =   "TCIU202.frx":2283
            Left            =   7410
            List            =   "TCIU202.frx":2285
            Style           =   2  'Dropdown List
            TabIndex        =   217
            Top             =   2310
            Width           =   2535
         End
         Begin VB.ComboBox cboCobrancaBc 
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
            ItemData        =   "TCIU202.frx":2287
            Left            =   2010
            List            =   "TCIU202.frx":2289
            Style           =   2  'Dropdown List
            TabIndex        =   216
            TabStop         =   0   'False
            Tag             =   "3"
            Top             =   195
            Width           =   6525
         End
         Begin VB.TextBox txtCodComponente 
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
            Index           =   20
            Left            =   1455
            MaxLength       =   3
            TabIndex        =   85
            Top             =   210
            Width           =   495
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   74
            Left            =   105
            TabIndex        =   219
            Top             =   270
            Width           =   1260
            _ExtentX        =   2223
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
            Caption         =   "Cod. Cobrança:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   77
            Left            =   5190
            TabIndex        =   220
            Top             =   2010
            Width           =   1680
            _ExtentX        =   2963
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
            Caption         =   "Instalação Sanitária:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   78
            Left            =   5700
            TabIndex        =   221
            Top             =   2370
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
            Caption         =   "Instalação Elétrica:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin MSComctlLib.ListView lstCond 
         Height          =   1530
         Left            =   -74910
         TabIndex        =   222
         Top             =   4410
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   2699
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "14"
            Text            =   "IC Anterior"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "15"
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "110"
            Text            =   "Nº"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "16"
            Text            =   "Complemento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "9"
            Text            =   "IM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "10"
            Text            =   "Contribuinte"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "11"
            Text            =   "CPF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "12"
            Text            =   "Tipo Logr"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "13"
            Text            =   "Logradouro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "111"
            Text            =   "Nº"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "112"
            Text            =   "Complemento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Tag             =   "113"
            Text            =   "Bairro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Tag             =   "114"
            Text            =   "CEP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Municipio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "UF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Ocupante"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Cpf Ocupante"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Cod. Cobrança"
            Object.Width           =   2540
         EndProperty
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   49
         Left            =   105
         TabIndex        =   231
         Top             =   5670
         Width           =   1875
         _ExtentX        =   3307
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
         Caption         =   "Motivo da Modificação:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSFrame fra 
         Height          =   585
         Index           =   9
         Left            =   105
         TabIndex        =   236
         Top             =   1905
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   1032
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
         Caption         =   "Aforamento"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtLivroAforamento 
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
            Left            =   3375
            MaxLength       =   5
            TabIndex        =   18
            Top             =   188
            Width           =   615
         End
         Begin VB.TextBox txtFichaAforamento 
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
            Left            =   1935
            MaxLength       =   5
            TabIndex        =   17
            Top             =   188
            Width           =   705
         End
         Begin VB.TextBox txtNumAforamento 
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
            Left            =   360
            MaxLength       =   5
            TabIndex        =   16
            Top             =   188
            Width           =   840
         End
         Begin VB.TextBox txtDataAforamento 
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
            Left            =   5835
            MaxLength       =   10
            TabIndex        =   20
            Top             =   188
            Width           =   1215
         End
         Begin VB.TextBox txtFolhaAforamento 
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
            Left            =   4695
            MaxLength       =   50
            TabIndex        =   19
            Top             =   180
            Width           =   585
         End
         Begin VB.TextBox txtRegistro 
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
            Left            =   8010
            MaxLength       =   50
            TabIndex        =   21
            Top             =   180
            Width           =   585
         End
         Begin VB.TextBox txtDtRegistro 
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
            Left            =   9990
            MaxLength       =   10
            TabIndex        =   22
            Top             =   195
            Width           =   990
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   79
            Left            =   4095
            TabIndex        =   237
            Top             =   240
            Width           =   495
            _ExtentX        =   873
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
            Caption         =   "Folha:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   80
            Left            =   90
            TabIndex        =   238
            Top             =   233
            Width           =   225
            _ExtentX        =   397
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
            Caption         =   "Nº:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   81
            Left            =   1455
            TabIndex        =   239
            Top             =   240
            Width           =   480
            _ExtentX        =   847
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
            Caption         =   "Ficha:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   87
            Left            =   2865
            TabIndex        =   229
            Top             =   240
            Width           =   480
            _ExtentX        =   847
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
            Caption         =   "Livro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   88
            Left            =   5355
            TabIndex        =   230
            Top             =   240
            Width           =   405
            _ExtentX        =   714
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
            Caption         =   "Data:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   89
            Left            =   7230
            TabIndex        =   226
            Top             =   225
            Width           =   735
            _ExtentX        =   1296
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
            Caption         =   "Registro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   90
            Left            =   8820
            TabIndex        =   227
            Top             =   240
            Width           =   1155
            _ExtentX        =   2037
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
            Caption         =   "Data Registro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
   End
   Begin VB.TextBox txtFatorFixo 
      Height          =   285
      Left            =   8640
      TabIndex        =   86
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   4560
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   228
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   1138
      Icone           =   "TCIU202.frx":228B
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   2
      Left            =   10230
      TabIndex        =   240
      Top             =   6765
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   1
      Left            =   9060
      TabIndex        =   241
      Top             =   6765
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   4
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   0
      Left            =   7890
      TabIndex        =   242
      Top             =   6765
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCIU202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim Cadastro As VSImposto
Dim NovoContrib As Boolean
Dim Sql As String
Private Boletim As TipoBoletim
Dim Consultando As Boolean
Dim Lote As New BCI

Private Sub AtualizaCodComponente(Combo As ComboBox)
    On Error Resume Next
    With Combo
        If .ListIndex + 1 = 0 Then
            txtCodComponente(Val(.Tag) - 1).Text = ""
        Else
            txtCodComponente(Val(.Tag) - 1).Text = .ListIndex + 1
        End If
    End With
End Sub

Function TotalProva(Valor As String) As String
    Static Total As Double
    If Trim(Valor) = "" Then Valor = "0"
    Total = CDbl(Valor) + Total
    TotalProva = Total
End Function

Public Sub HabilitaCaixa(Status As Boolean)
    txtIM.Enabled = Not Status
    txtNomeContrib.Enabled = Status
    
    txtNomeLogrContrib.Enabled = Status
    txtNumeroContrib.Enabled = Status
    txtCompContrib.Enabled = Status
    txtBairroContrib.Enabled = Status
    txtCep.Enabled = Status
    txtMunic.Enabled = Status
    txtIM = ""
    txtNomeContrib = ""
    
    txtNomeLogrContrib = ""
    txtNumeroContrib = ""
    txtCompContrib = ""
    txtBairroContrib = ""
    txtCep = ""
    txtMunic = ""
    txtCpfCgc = ""
    txtOcupante = ""
    If Status Then txtNomeContrib.SetFocus
End Sub

Private Sub cboArborizacao_Click()
    AtualizaCodComponente cboArborizacao
End Sub

Private Sub cboArborizacao_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboArborizacao
End Sub

Private Sub cboCobranca_Click()
    AtualizaCodComponente cboCobranca
End Sub

Private Sub cboCobranca_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboCobranca
End Sub

Private Sub cboConservacao_Click()
    AtualizaCodComponente cboConservacao
End Sub

Private Sub cboConservacao_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboConservacao
End Sub

Private Sub cboDestinacao_Click()
    AtualizaCodComponente cboDestinacao
End Sub

Private Sub cboDestinacao_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboDestinacao
End Sub

Private Sub cboEstrutura_Click()
    AtualizaCodComponente cboEstrutura
End Sub

Private Sub cboEstrutura_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboEstrutura
End Sub

Private Sub cboLimites_Click()
    AtualizaCodComponente cboLimites
End Sub

Private Sub cboLimites_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboLimites
End Sub

Private Sub cboOcupLote_Click()
    AtualizaCodComponente cboOcupLote
End Sub

Private Sub cboOcupLote_KeyPress(KeyAscii As Integer)
    cboOcupLote_Click
End Sub

Private Sub cboPadrao_Click()
    AtualizaCodComponente cboPadrao
End Sub

Private Sub cboPadrao_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboPadrao
End Sub

Private Sub cboPatrimonio_Click()
    AtualizaCodComponente cboPatrimonio
End Sub

Private Sub cboPatrimonio_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboPatrimonio
End Sub

Private Sub cboPedol_Click()
    AtualizaCodComponente cboPedol
End Sub

Private Sub cboPedol_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboPedol
End Sub

Private Sub cboPredio_Click()
    AtualizaCodComponente cboPredio
End Sub

Private Sub cboPredio_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboPredio
End Sub

Private Sub cboSentido_Click()
    AtualizaCodComponente cboSentido
End Sub

Private Sub cboSentido_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboSentido
End Sub

Private Sub cboSit_Click()
    AtualizaCodComponente cboSit
End Sub

Private Sub cboSit_KeyPress(KeyAscii As Integer)
    AtualizaCodComponente cboSit
End Sub

Private Sub cboTipoImovel_Click()
    If cboTipoImovel = "PREDIAL" Then
        Boletim = tbo_Predial
    Else
        Boletim = tbo_Territorial
    End If

End Sub

Private Sub cmdAdCond_Click()
     'NOVIDADE
    Dim ItmX As Object
    Dim i As Byte
    
    On Error Resume Next
    Set ItmX = lstCond.ListItems.Add(, , txtIc(9))
    ItmX.SubItems(1) = txtInscAnteriorBC
    ItmX.SubItems(2) = cboTipoImovelBc.ListIndex + 1
    ItmX.SubItems(3) = txtNumeroBc
    ItmX.SubItems(4) = txtComplementoBc
    ItmX.SubItems(5) = IIf(Trim(txtIMBc) = "", "", txtIMBc)
    ItmX.SubItems(6) = txtNomeContribBc
    ItmX.SubItems(7) = txtCpfCgcBc
    ItmX.SubItems(8) = cboTipoLogrContribBc
    ItmX.SubItems(9) = txtNomeLogrContribBc
    ItmX.SubItems(10) = txtNumeroBc
    ItmX.SubItems(11) = txtComplementoBc
    ItmX.SubItems(12) = txtBairroContribBc
    ItmX.SubItems(13) = txtCepImBc
    ItmX.SubItems(14) = txtMunicBc
    ItmX.SubItems(15) = cboUFBc
    ItmX.SubItems(16) = txtOcupanteBc
    ItmX.SubItems(17) = txtCpfOcupanteBc
    ItmX.SubItems(18) = txtCodComponente(20)
    txtIc(9).SetFocus
End Sub

Private Sub cmdAdEdif_Click()
    Dim ItmX As Object
    Dim i As Byte
    If Trim(txtInscImobiliaria) = "" Then
        Avisa "Informe a unidade."
        txtInscImobiliaria.SetFocus
        Exit Sub
    End If
    For i = 8 To 15
        If Trim(txtCodComponente(i)) = "" Then
            Avisa "Informe todos os campos."
            txtCodComponente(i).SetFocus
            Exit Sub
        End If
    Next
    
    Set ItmX = lstEdific.ListItems.Add(, , txtInscImobiliaria)
    ItmX.SubItems(1) = txtCodComponente(13)
    ItmX.SubItems(2) = txtCodComponente(14)
    ItmX.SubItems(3) = txtPavimento
    ItmX.SubItems(4) = txtCodComponente(15)
    For i = 8 To 13
        ItmX.SubItems(i - 3) = txtCodComponente(i)
    Next
    ItmX.SubItems(10) = txtAnoConst
    ItmX.SubItems(11) = IIf(Trim(txtAreaEdif) = "", 0, txtAreaEdif)
    ItmX.SubItems(12) = IIf(Trim(txtAreaEdifTotal) = "", 0, txtAreaEdifTotal)
    ItmX.SubItems(13) = IIf(Trim(txtFracaoEdif) = "", 0, txtFracaoEdif)
    
    For i = 8 To 15
        txtCodComponente(i) = ""
    Next
    txtAnoConst = ""
    txtAreaEdif = ""
    txtAreaEdifTotal = ""
    txtFracaoEdif = ""
    txtPavimento = ""
    txtInscImobiliaria = ""
    txtInscImobiliaria.SetFocus
    
End Sub

Private Sub cmdEnter_Click()
        SendKeys "{Tab}"
End Sub

Private Sub cmdImprime_Click()
    If Me.Tag <> "" Then
        With Rpt
            If Not .DefinirArquivo(Bdados, App.Path & "\TCIU201.rpt") Then Exit Sub
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
            .Selecao = "{TAB_IMOVEL.tim_ic} = '" & Me.Tag & "'"
            .Titulo = "Ficha Cadastral"
            .Arvore = False
            .Visualizar
            DoEvents
        End With
        Set Rpt = Nothing
    End If
End Sub

Private Sub cmdNovo_Click()
    Static Status As Boolean
    Status = Not Status
    HabilitaCaixa Status
End Sub

Private Sub cmdOpcao_Click(Index As Integer)
    Dim rs As VSRecordset
    Select Case Index
        Case 0
            Sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where tci_nome like '" & txtNomeContrib & "%' or tci_nome like '% " & txtNomeContrib & "%'"
            Sql = Sql & " and tci_tsc_cod_sit_cad =1"
            If Not Bdados.AbreTabela(Sql, rs) Then
                Call Util.Avisa("Nenhum contribuinte encontrado.")
            End If
            Bdados.FechaTabela rs
            MontaGrid Bdados, lstPesq, Sql, 1400
        Case 1
            NovoContrib = True
            txtIM = ""
            Call HabilitaCaixa(True)
            txtCep = Temp.PegaParametro(Bdados, "CEP")
            txtNomeContrib.SetFocus
    End Select
End Sub

Private Sub Form_Activate()
    
    Dim i As Byte
    If Me.Tag <> "" Then
        Consultando = True
        For i = 0 To 8
            fra(i).Enabled = False
        Next
        txtIc(0) = Left(Me.Tag, 2)
        txtIc(1) = Mid(Me.Tag, 3, 2)
        txtIc(2) = Mid(Me.Tag, 5, 4)
        txtIc(3) = Mid(Me.Tag, 9, 4)
        txtIc(4) = IIf(Right(Me.Tag, 3) < 200, "000", IIf(Right(Me.Tag, 3) < 600, "200", "600"))
        Call txtic_LostFocus(4)
        tabCad.Tab = 0
        cmd(0).Enabled = False
        cmd(1).Enabled = False
    End If
    DoEvents
    Consultando = False
End Sub

Private Sub Form_Load()
    
    Dim Controle As Control
    Dim i As Byte
    Dim rs As VSRecordset
    Set Cadastro = New VSImposto
    
    Call Edita.AtualizaCombo(Bdados, cboTipoLogrContribBc, "Select ttl_nome From Tab_Tipo_Logr")
    Call AtualizaUF(cboUFBc)
    
    For Each Controle In Controls
        If IsNumeric(Controle.Tag) Then
            If Val(Controle.Tag) < 20 Then Call Edita.AtualizaCombo(Bdados, Controle, "Select tco_descricao_componente From Tab_Componente_Avancado Where tco_grupo = " & Controle.Tag & " order by tco_cod_componente asc")
        End If
    Next
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    NovoContrib = True
    Bdados.FechaTabela rs
    Boletim = tbo_Territorial
End Sub

Private Sub lstCond_DblClick()
    If lstCond.SelectedItem Is Nothing Then Exit Sub
    Dim ItmX As Object
    Dim i As Byte
    
    On Error Resume Next
    txtIc(9) = lstCond.SelectedItem
    txtInscAnteriorBC = lstCond.SelectedItem.SubItems(1)
    cboTipoImovelBc.ListIndex = lstCond.SelectedItem.SubItems(2) - 1
    txtNumeroBc = lstCond.SelectedItem.SubItems(3)
    txtComplementoBc = lstCond.SelectedItem.SubItems(4)
    txtIMBc = lstCond.SelectedItem.SubItems(5)
    txtNomeContribBc = lstCond.SelectedItem.SubItems(6)
    txtCpfCgcBc = lstCond.SelectedItem.SubItems(7)
    cboTipoLogrContribBc = lstCond.SelectedItem.SubItems(8)
    txtNomeLogrContribBc = lstCond.SelectedItem.SubItems(9)
    txtNumeroBc = lstCond.SelectedItem.SubItems(10)
    txtComplementoBc = lstCond.SelectedItem.SubItems(11)
    txtBairroContribBc = lstCond.SelectedItem.SubItems(12)
    txtCepImBc = lstCond.SelectedItem.SubItems(13)
    txtMunicBc = lstCond.SelectedItem.SubItems(14)
    cboUFBc = lstCond.SelectedItem.SubItems(15)
    txtOcupanteBc = lstCond.SelectedItem.SubItems(16)
    txtCpfOcupanteBc = lstCond.SelectedItem.SubItems(17)
    txtCodComponente(20) = lstCond.SelectedItem.SubItems(18)
    lstCond.ListItems.Remove lstCond.SelectedItem.Index
    DoEvents
End Sub


Private Sub lstEdific_Click()
    Dim i As Byte
    Dim Sql As String
    Dim rs As VSRecordset
    On Error Resume Next
    If lstEdific.SelectedItem Is Nothing Then Exit Sub
    If Trim(txtCodComponente(13)) <> "" Then
        If Not Confirma("Existe uma unidade edificada em aberto. Deseja exclui-la?") Then
            Exit Sub
        End If
    End If
    txtInscImobiliaria = Right(lstEdific.SelectedItem, 3)
    txtCodComponente(13) = lstEdific.SelectedItem.SubItems(1)
    txtCodComponente(14) = lstEdific.SelectedItem.SubItems(2)
    txtPavimento = lstEdific.SelectedItem.SubItems(3)
    txtCodComponente(15) = lstEdific.SelectedItem.SubItems(4)
    For i = 8 To 12
        txtCodComponente(i) = lstEdific.SelectedItem.SubItems(i - 3)
    Next
    txtAnoConst = lstEdific.SelectedItem.SubItems(10)
    txtAreaEdif = lstEdific.SelectedItem.SubItems(11)
    txtAreaEdifTotal = lstEdific.SelectedItem.SubItems(12)
    txtFracaoEdif = lstEdific.SelectedItem.SubItems(13)
    
    'ElseIf CInt(Nvl(txtInscImobiliaria, 0)) >= 200 Then
        'CONSULTA BC
    If lstEdific.SelectedItem >= 200 Then
        'tabCad.TabEnabled(2) = True
        txtIc(5) = txtIc(0)
        txtIc(6) = txtIc(1)
        txtIc(7) = txtIc(2)
        txtIc(8) = txtIc(3)
        txtCodLogrBc = txtCodLogr
        txtLogr = txtTipoLogrBt
        txtNomeLogr = txtLogrBt
        txtNumeroBc = txtNumero
        txtLoteamentoBc = txtLoteamento
        txtQuadraBc = txtQuadra
        txtLoteBc = txtLote
        txtBairro = txtBairroBt
        txtCepImBc = txtCep
        txtIc(9) = lstEdific.SelectedItem
        txtIc(9).Enabled = True
        
        Sql = "SELECT TIM_complemento,tim_tci_im,tim_ocupante,tim_cgc_cpf_ocupante,tim_tipo_imovel from tab_imovel where tim_ic ='" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & "' AND TIM_UNIDADE=" & lstEdific.SelectedItem
        If Bdados.AbreTabela(Sql, rs) Then
            txtComplementoBc = "" & rs(0)
            txtIMBc = "" & rs(1)
            txtIMBc_LostFocus
            txtOcupanteBc = "" & rs(2)
            txtCpfOcupanteBc = "" & rs(3)
            cboTipoImovelBc.ListIndex = rs(4) - 1
            DoEvents
            Sql = "select tdi_tco_cod_componente from tab_detalhe_imovel where tdi_tgc_cod_grupo = 3 and tdi_tim_ic_unidade = " & txtIc(9) & " and tdi_tim_ic ='" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & "'"
            If Bdados.AbreTabela(Sql, rs) Then
                txtCodComponente(20) = rs(0)
            End If
            Bdados.FechaTabela rs
        End If
    End If
    If Me.Tag = "" Then
        lstEdific.ListItems.Remove lstEdific.SelectedItem.Index
    End If
    'FIM CONSULTA BC
End Sub

Private Sub lstPesq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaGrid lstPesq, ColumnHeader
End Sub

Private Sub lstPesq_DblClick()
    On Error Resume Next
    txtIM = lstPesq.SelectedItem
    Call txtIm_LostFocus
End Sub

Private Sub tabCad_Click(PreviousTab As Integer)
    'NOVIDADE
    If tabCad.Tab = 2 Then
        If Trim(txtInscImobiliaria) = "" And Trim(txtIc(4)) <> "" Then
            txtInscImobiliaria.Enabled = True
            txtIc(10) = txtIc(0)
            txtIc(11) = txtIc(1)
            txtIc(12) = txtIc(2)
            txtIc(13) = txtIc(3)
            If txtInscImobiliaria.Enabled Then txtInscImobiliaria.SetFocus
        End If
    ElseIf tabCad.Tab = 3 Then
        If Trim(txtIc(4)) <> "" Then
            txtIc(5) = txtIc(0)
            txtIc(6) = txtIc(1)
            txtIc(7) = txtIc(2)
            txtIc(8) = txtIc(3)
            txtIc(9).Enabled = True
            
            cboTipoImovelBc.ListIndex = cboTipoImovel.ListIndex
            txtCodLogrBc = txtCodLogr
            txtLogr = txtTipoLogrBt
            txtNomeLogr = txtLogrBt
            txtNumeroBc = txtNumero
            txtLoteamentoBc = txtLoteamento
            txtQuadraBc = txtQuadra
            txtLoteBc = txtLote
            txtBairro = txtBairroBt
            txtCepImBc = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
        End If
    End If
End Sub



Private Sub txtAconstr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtAnoAq_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtAreaNao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtAreaEdif_Change()
    On Error Resume Next
    If Trim(txtAreaEdif) <> "" Then
        If Not IsNumeric(txtAreaEdif) Then Exit Sub
        txtFracaoEdif = Format(CDbl(Nvl(txtAreaEdif, 1)) / CDbl(Nvl(txtAreaEdifTotal, 1)), "#0.000,0000")
    Else
        txtFracaoEdif = ""
    End If
End Sub

Private Sub txtAreaEdifTotal_Change()
   Call txtAreaEdif_Change
End Sub

Private Sub txtAreaLote_LostFocus()
    tabCad.Tab = 2
    txtInscImobiliaria.SetFocus
End Sub

Private Sub txtBairroContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCepImBc_LostFocus()
    txtCepImBc = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
End Sub

Private Sub txtCodBairro_LostFocus()
    Dim rs As VSRecordset
    Dim Sql As String
    If Trim(txtCodBairro) <> "" Then
        Sql = " select TBA_NOME from TAB_BAIRRO where tba_cod_bairro=" & txtCodBairro & " and tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
        If Bdados.AbreTabela(Sql, rs) Then
            txtBairroBt = rs(0)
        Else
            Avisa "Bairro inexistente."
            txtCodBairro.SetFocus
        End If
    Else
        txtBairroBt = ""
        Bdados.FechaTabela rs
    End If
    
End Sub

Private Sub txtCodComponente_Change(Index As Integer)
    Dim Controle As Control
    On Error GoTo trata
     If Index = 20 Then
        cboCobrancaBc.ListIndex = Nvl(txtCodComponente(Index).Text, 0) - 1
        Exit Sub
    End If
    For Each Controle In Controls
        If Controle.Tag = Index + 1 Then
            Controle.ListIndex = Util.Nvl(txtCodComponente(Index).Text, 0) - 1
            Exit For
        End If
    Next
trata:
    If Err.Number = 380 Then
        txtCodComponente(Index).SetFocus
    End If
End Sub

Private Sub txtCodComponente_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodLogr_LostFocus()
    Dim Query As String
    Dim rs As VSRecordset
    If Trim(txtCodLogr) = "" Then Exit Sub
    Query = "SELECT TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome, " & _
        " TAB_BAIRRO.TBA_NOME FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
        " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
         " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & txtCodLogr & "' and tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio & " and tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
    If Bdados.AbreTabela(Query, rs) Then
        txtTipoLogrBt = rs(0)
        txtLogrBt = rs(1)
    Else
        Avisa "Código de logradouro inválido."
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtCodLogrBc_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    
    Sql = "Select "
End Sub

Private Sub txtCodLogrContrib_LostFocus()
    Dim Query As String
    Dim rs As VSRecordset
    If Trim(txtCodLogrContrib) <> "" Then
        If Trim(txtCodLogrContrib) = "" Then Exit Sub
        Query = "SELECT TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome, " & _
            " TAB_BAIRRO.TBA_NOME FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
            " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
             " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & txtCodLogrContrib & "' and tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
        If Bdados.AbreTabela(Query, rs) Then
            txtNomeTipoLogrContrib = rs(0)
            txtNomeLogrContrib = rs(1)
            txtBairroContrib = rs(2)
            txtMunic = Aplicacoes.Municipio
        Else
            Avisa "Código de logradouro inválido."
            txtCodLogr.SetFocus
        End If
        Bdados.FechaTabela rs
        txtCep = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
        txtUf = Temp.PegaParametro(Bdados, "ESTADO CLIENTE")
        
    End If
End Sub

Private Sub txtCodMens_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtCodReduzido_Change()

End Sub

Private Sub txtCompContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFracao_Change()

End Sub

Private Sub txtCpfCgc_LostFocus()
    If Len(Trim(txtCpfCgc)) = 11 Then
        If Not Util.ValidaCpf(Trim(txtCpfCgc)) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCpfCgc.SetFocus
            Exit Sub
        End If
        txtCpfCgc = Edita.FormataTexto(txtCpfCgc, Cpf)
    ElseIf Len(Trim(txtCpfCgc)) = 14 And Mid(Trim(txtCpfCgc), 4, 1) <> "." Then
        txtCpfCgc.MaxLength = 20
        txtCpfCgc = Trim(txtCpfCgc)
        txtCpfCgc = Edita.FormataTexto(txtCpfCgc, Cgc)
    ElseIf Trim(txtCpfCgc) <> "" And Len(Trim(txtCpfCgc)) <> 18 And Mid(Trim(txtCpfCgc), 4, 1) <> "." Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
        txtCpfCgc.SetFocus
    End If
End Sub

Private Sub txtCpfCgcBc_LostFocus()
    If Len(txtCpfCgcBc) = 11 Then
        If Not Util.ValidaCpf(txtCpfCgcBc) Then
            Call Util.Informa("Número de CPF inválido.")
             txtCpfCgcBc.SetFocus
            Exit Sub
        End If
        txtCpfCgcBc = Edita.FormataTexto(txtCpfCgcBc, Cpf)
    ElseIf Len(txtCpfCgcBc) = 14 And Mid(txtCpfCgcBc, 4, 1) <> "." Then
        txtCpfCgcBc.MaxLength = 20
        txtCpfCgcBc = Edita.FormataTexto(txtCpfCgcBc, Cgc)
    ElseIf Trim(txtCpfCgcBc) <> "" And Len(txtCpfCgcBc) <> 18 And Mid(txtCpfCgcBc, 4, 1) <> "." Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
        txtCpfCgcBc.SetFocus
    End If
End Sub

Private Sub txtCpfOcupante_LostFocus()
    If Len(txtCpfOcupante) = 11 Then
        If Not Util.ValidaCpf(txtCpfOcupante) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCpfOcupante.SetFocus
            Exit Sub
        End If
        txtCpfOcupante = Edita.FormataTexto(txtCpfOcupante, Cpf)
    End If
    tabCad.Tab = 1
    DoEvents
    txtCodComponente(0).SetFocus
End Sub

Private Sub txtCpfOcupanteBc_LostFocus()
    If Len(txtCpfOcupanteBc) = 11 Then
        If Not Util.ValidaCpf(txtCpfOcupanteBc) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCpfOcupanteBc.SetFocus
            Exit Sub
        End If
        txtCpfOcupanteBc = Edita.FormataTexto(txtCpfOcupanteBc, Cpf)
    End If
End Sub

Private Sub txtDataAforamento_LostFocus()
    txtDataAforamento = Edita.FormataTexto(txtDataAforamento, Data)
End Sub

Private Sub txtDtRegistro_LostFocus()
    txtDtRegistro = Edita.FormataTexto(txtDtRegistro, Data)
End Sub

Private Sub txtFracaoEdif_LostFocus()
    If CInt(Nvl(Trim(txtIc(4)), 0)) >= 200 Then
        tabCad.Tab = 3
        DoEvents
        txtIc(9).SetFocus
    End If
End Sub

Private Sub txtIc_Change(Index As Integer)
    If Len(txtIc(Index)) = txtIc(Index).MaxLength Then
       SendKeys "{ENTER}"
    End If
End Sub

Private Sub txtic_LostFocus(Index As Integer)
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Tem As String
    Dim Temp As String
    
    If Index = 9 Then
        
        If Trim(txtIc(9)) = "" Then Exit Sub
        Screen.MousePointer = 11
        If AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
            Sql = "Select * from tab_imovel where tim_ic_auxiliar ='" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & "'"
        Else
            Sql = "Select * from tab_imovel where (tIM_ic ='" & _
            txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & "' AND TIM_UNIDADE =" & txtIc(9) & ")" & Temp
        End If
        If Bdados.AbreTabela(Sql, rs) Then
            txtIMBc = rs!tim_tci_im
            txtIMBc_LostFocus
            cboTipoImovelBc.ListIndex = rs!tim_tipo_imovel - 1
            txtInscAnteriorBC = "" & rs!tim_ic_anterior
            txtOcupanteBc = "" & rs!tim_ocupante
            txtCpfCgcBc = "" & rs!tim_cgc_cpf_ocupante
            Sql = "Select * from tab_detalhe_imovel where TDI_TIM_IC ='" & _
            txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & "' AND tdi_tim_ic_unidade = " & CInt(Nvl(txtIc(9), 0)) & " and tdi_tgc_cod_grupo = 3" ' order by tdi_tco_cod_componente asc"
            If Bdados.AbreTabela(Sql, rs) Then
                txtCodComponente(20) = rs!TDI_VALOR_ITEM
            End If
        End If
        Bdados.FechaTabela rs
    ElseIf Index = 4 Then
        If Trim(txtIc(4)) = "" Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        'Imperatriz 16.01.03 (Queiroz)
        'Sql = "Select * from tab_imovel where (tIM_ic ='" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & IIf(CInt(txtIc(4)) < 200, "000", txtIc(4)) & "'" & IIf(CInt(txtIc(4)) >= 200, " AND TIM_UNIDADE =" & txtIc(4), "") & ") " & Temp
        Sql = "Select * from tab_imovel where tIM_ic ='" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & txtIc(4) & "'" '& IIf(CInt(txtIc(4)) >= 200, " AND TIM_UNIDADE =" & txtIc(4), "") & ") " & Temp
        If Bdados.AbreTabela(Sql, rs) Then
            txtIcAnterior = "" & IIf(rs!tim_ic_anterior = 0, "", rs!tim_ic_anterior)
            cboTipoImovel.ListIndex = rs!tim_tipo_imovel - 1
            txtCodLogr = "" & rs!tim_tlg_cod_logradouro
            txtCodLogr_LostFocus
            txtNumero = "" & rs!tim_numero
            txtComplemento = "" & rs!tim_complemento
            txtLoteamento = "" & rs!tim_loteamento
            txtQuadra = "" & rs!tim_QUADRA
            txtLote = "" & rs!tim_lote
            txtOcupante = "" & rs!tim_ocupante
            txtCpfOcupante = "" & rs!tim_cgc_cpf_ocupante
            
            txtIcAnterior = "" & rs!tim_ic_anterior
            txtCodMens = "" & rs!tim_COD_MENSAGEM
            txtNumAforamento = "" & rs!tim_AFORAMENTO_NUMERO
            txtFichaAforamento = "" & rs!tim_AFORAMENTO_FICHA
            txtLivroAforamento = "" & rs!tim_AFORAMENTO_LIVRO
            txtFolhaAforamento = "" & rs!tim_AFORAMENTO_FOLHA
            txtDataAforamento = "" & rs!tim_AFORAMENTO_DATA
            'VOU PEGAR O CONTRIBUINTE
            txtIM = "" & rs!tim_tci_im
            txtIm_LostFocus
            txtZona = "" & rs!tim_ZONA
            txtCodBairro = "" & rs!tim_TBA_COD_BAIRRO
            txtCodBairro_LostFocus
            'VOU PEGAR OS DETALHES
            'Temp = " or (TDI_TIM_IC ='" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & "' AND tdi_tim_ic_unidade =" & CInt(txtic(4)) & ")"
            Sql = "Select * from TAB_DETALHE_IMOVEL where (TDI_TIM_IC ='" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & txtIc(4) & "' AND tdi_tim_ic_unidade = " & CInt(txtIc(4)) & ") " & Temp '& " order by tdi_tco_cod_componente asc"
            If Bdados.AbreTabela(Sql, rs) Then
                rs.MoveFirst
                Do While Not rs.EOF
                    If rs!tdi_tgc_cod_grupo <= 8 Then
                        On Error Resume Next
                        txtCodComponente(rs!tdi_tgc_cod_grupo - 1) = rs!TDI_VALOR_ITEM + 1
                        On Error GoTo 0
                    Else
                        Dim Controle As Control
                        On Error Resume Next
                        For Each Controle In Controls
                            If IsNumeric(Controle.Tag) Then
                                If CInt(Controle.Tag) = rs!tdi_tgc_cod_grupo Then
                                    Controle.Text = rs!TDI_VALOR_ITEM
                                End If
                            End If
                        Next
                        On Error GoTo 0
                    End If
                    rs.MoveNext
                Loop
            End If
            'Vou pegar as construcões 'sergio
            Dim i As Byte
            Dim ItmX As Object
            If CInt(txtIc(4)) = 0 Then
                Temp = "  TDI_TIM_IC ='" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & txtIc(4) & "' and tdi_tim_ic_unidade > 0"
            Else
                Temp = " ( TDI_TIM_IC >'" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & txtIc(4) & "' AND TDI_TIM_IC <'" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & "300'" & ")"
            End If
            Sql = "Select * from tab_detalhe_imovel where " & Temp & " order by tdi_tim_ic_unidade asc, tdi_tgc_cod_grupo asc"
            lstEdific.ListItems.Clear
            If Bdados.AbreTabela(Sql, rs) Then
                rs.MoveFirst
                Set ItmX = lstEdific.ListItems.Add(, , Format(rs!tdi_tim_ic_unidade, "000"))
                Dim Conta As Byte
                Conta = 1
                Do While Not rs.EOF
                    If Format(Nvl(rs!tdi_tim_ic_unidade, 0), "000") <> ItmX Then
                        Set ItmX = lstEdific.ListItems.Add(, , Format(rs!tdi_tim_ic_unidade, "000"))
                        Conta = Conta + 1
                    End If
                    'If Rs!tdi_tgc_cod_grupo >= 14 And Rs!tdi_tgc_cod_grupo <= 15 Then
                    For i = 2 To 14
                        If CInt(lstEdific.ColumnHeaders(i).Tag) = CInt(rs!tdi_tgc_cod_grupo) Then
                            ItmX.SubItems(i - 1) = rs!TDI_VALOR_ITEM
                            Screen.MousePointer = 0
                            Exit For
                        End If
                    Next
                    rs.MoveNext
                Loop
                txtIc(10) = txtIc(0)
                txtIc(11) = txtIc(1)
                txtIc(12) = txtIc(2)
                txtIc(13) = txtIc(3)
            End If
            For i = 0 To 4
                txtIc(i).Enabled = False
            Next
            'Vou pegar os condominios
            If CInt(Trim(txtIc(4))) < 200 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            Dim Campos As String
            Campos = "TIM_UNIDADE, tim_ic_anterior,tim_tipo_imovel,tim_numero,  tim_complemento, " _
                    & "tim_tci_im , " _
                    & "  tci_nome,tci_cgc_cpf,tci_logradouro," _
                    & " tci_nome_logradouro, tci_numero," _
                    & "tci_complemento, tci_bairro,tci_cep,tci_cidade,tci_UF," _
                    & "tim_ocupante,tim_cgc_cpf_ocupante" ',tdi_tco_cod_componente "
            
            Sql = "Select " & Campos & " from tab_imovel,tab_contribuinte,tab_detalhe_imovel " & _
                " where (TIM_IC >'" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & txtIc(4) & _
                "' AND TIM_IC <'" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & "300'" & ")" & _
                " and tim_tci_im = tci_im  and tim_ic = tdi_tim_ic and tdi_tgc_cod_grupo = 3 order by tim_unidade asc"
                
            lstCond.ListItems.Clear
            If Bdados.AbreTabela(Sql, rs) Then
                rs.MoveFirst
                Conta = 1
                Do While Not rs.EOF
                    Set ItmX = lstCond.ListItems.Add(, , Format(rs!tim_unidade, "000"))
                    For i = 1 To 18
                        On Error Resume Next
                        ItmX.SubItems(i) = CStr("" & rs(CInt(i)))
                    Next
                    rs.MoveNext
                Loop
                txtIc(10) = txtIc(0)
                txtIc(11) = txtIc(1)
                txtIc(12) = txtIc(2)
                txtIc(13) = txtIc(3)
            End If
            For i = 0 To 4
                txtIc(i).Enabled = False
            Next
        Else
            Avisa "Imóvel não cadastrado."
            ' LIMPA A TELA E DEVOLVE O FOCO PARA O CAMPO IC
            cmd_Click 0
        End If
        Bdados.FechaTabela rs
    End If
    Screen.MousePointer = 0
    NovoContrib = False
End Sub

Private Sub txtIcAnterior_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCpfCgc_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    Dim rs As VSRecordset
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIM) <> "" Then
        txtIM = Cadastro.FormataInscricao(txtIM, InscContrib)
        Sql = "Select tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, " & _
        " tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF from Tab_Contribuinte where tci_im = '" & txtIM & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtNomeContrib = rs(0) 'Rs!tci_Nome
            txtNomeTipoLogrContrib = rs(1)
            txtNomeLogrContrib = rs(2) '!tci_nome_logradouro
            txtNumeroContrib = rs(3) '!tci_numero
            txtCompContrib = rs(4) '!tci_complemento
            txtBairroContrib = rs(5) '!tci_bairro
            txtCep = rs(6) '!tci_cep
        
            txtMunic = rs(7)
            txtUf = rs(8) '!tci_UF
            
            txtCpfCgc = "" & rs(9)
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtIM.Enabled = True
            txtIM.SetFocus
        End If
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtIMBc_LostFocus()
    Dim rs As VSRecordset
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIMBc) <> "" Then
        txtIMBc = Cadastro.FormataInscricao(txtIMBc, InscContrib)
        Sql = "Select tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF from Tab_Contribuinte where tci_im = '" & txtIMBc & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtNomeContribBc = rs(0)  'Rs!tci_Nome
            cboTipoLogrContribBc.ListIndex = Cadastro.BuscaCodLogr(rs(1)) - 1
            txtNomeLogrContribBc = rs(2)  '!tci_nome_logradouro
            txtNumeroContribBc = rs(3)  '!tci_numero
            txtCompContribBc = rs(4)  '!tci_complemento
            txtBairroContribBc = rs(5)  '!tci_bairro
            txtCepBc = rs(6)  '!tci_cep
            txtMunicBc = rs(7)
            cboUFBc = rs(8)  '!tci_UF
            txtCpfCgcBc = "" & rs(9)
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtIMBc.Enabled = True
            txtIMBc.SetFocus
        End If
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtLoteamento_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.Maiuscula(KeyAscii)
End Sub

Private Sub txtMunic_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMunic_LostFocus()
    If Trim(txtMunic) = "" Then txtMunic = Aplicacoes.Municipio
End Sub

Private Sub txtNomeContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomeLogrContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumeroContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtOcupante_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtQuadra_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtSecao_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtServCalc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtServIlum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtServLimp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtTtotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtUnids_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtZona_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub cmd_Click(Index As Integer)
    On Error Resume Next
    Dim Valores As String
    Dim Campos As String
    Dim DataReab As Date
    Dim RsAux As VSRecordset
    Dim rs As VSRecordset
    Dim InscricaoMunicipal  As String
    Dim InscricaoCadastral As String
    Dim CodLogr As Long
    Dim DtVenc As String
    Dim SitCadastral As String
    Static Unidades As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Cadastro As New VSImposto
    Select Case cmd(Index).Caption
        Case "&Salvar"
                If Trim(txtCodBairro) = "" Then
                    Util.Informa "Falta a definição do bairro."
                    txtCodBairro.SetFocus
                    tabCad.Tab = 0
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                'VERIFCANDO BP'S
                If cboTipoImovel = "PREDIAL" Then
                    If Not Lote.VerificaDigitacaoBP(lstEdific, txtCodBairro, tabCad) Then
                        txtCodComponente(13).SetFocus
                        Exit Sub
                    End If
                End If
                txtFatorFixo.Tag = "1000"
                CodLogr = txtCodLogr
                InscricaoCadastral = txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & txtIc(4)
                InscricaoMunicipal = txtIM
                Screen.MousePointer = 11
                If Not Lote.VerificaFechamentoAreas(lstEdific) Then Exit Sub
                Lote.GravaHistorico InscricaoCadastral, txtMotivo
                'GRAVANDO BT
                Lote.CarregaDadosContribuinte InscricaoMunicipal, txtNomeContrib, txtCpfCgc, "", txtLogr, txtNomeLogrContrib, _
                        txtNumeroContrib, txtCompContrib, "", txtBairroContrib, txtCep, txtMunic, txtUf
                If Not Lote.InsereContribuinte() Then Exit Sub
                'SQz (BLS, 15/04/2003): Ao apagar perde-se a informação do valor venal (!)
                Lote.ApagaImovel InscricaoCadastral
                Lote.CarregaDadosImovel InscricaoCadastral, txtIcAnterior, txtIc(4), "0", "", "", CStr(CodLogr), txtCodBairro, _
                     txtNumero, txtComplemento, txtLote, txtQuadra, txtLoteamento, Boletim, txtOcupante, txtCpfOcupante, _
                     txtCodMens, txtZona, txtNumAforamento, txtFichaAforamento, txtLivroAforamento, txtFolhaAforamento, txtRegistro, txtDataAforamento, txtDtRegistro, , , , txtSecao
                
                If Not Lote.InsereTerritorio() Then Exit Sub
                Lote.ApagaDetalheImovel InscricaoCadastral
                Call Lote.GravaComponentes(InscricaoCadastral, Me, 1, 8, False, txtIc(4), 0)
                Call Lote.GravaComponentes(InscricaoCadastral, Me, 100, 109, True, txtIc(4), 0)

                'GRAVANDO BP
                Lote.GravaBP lstEdific, txtCodMens, txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3), txtIc(4)
                'GRAVANDO BC'S
                If CInt(Nvl(Trim(txtIc(4)), 0)) >= 200 Then
                    cboCobrancaBc.Tag = "3"
                    cboCobranca.Tag = ""
                    If lstCond.ListItems.Count > 0 Then
                        For j = 1 To lstCond.ListItems.Count 'Para cada edificacao
                            lstCond.ListItems(j).Selected = True
                            InscricaoMunicipal = txtIMBc
                            CodLogr = txtCodLogr
                            InscricaoCadastral = txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & lstCond.SelectedItem
                            'INSERE CONTRIBUINTE
                            If lstCond.SelectedItem.ListSubItems(5) = "" Then
                                InscricaoMunicipal = Cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
                            Else
                                InscricaoMunicipal = lstCond.SelectedItem.ListSubItems(5)
                            End If
                            
                            Lote.CarregaDadosContribuinte InscricaoMunicipal, lstCond.SelectedItem.ListSubItems(6), _
                                    "", lstCond.SelectedItem.ListSubItems(20), lstCond.SelectedItem.ListSubItems(8), _
                                    lstCond.SelectedItem.ListSubItems(9), lstCond.SelectedItem.ListSubItems(10), _
                                     lstCond.SelectedItem.ListSubItems(11), lstCond.SelectedItem.ListSubItems(21), _
                                    lstCond.SelectedItem.ListSubItems(12), lstCond.SelectedItem.ListSubItems(13), _
                                    lstCond.SelectedItem.ListSubItems(14), lstCond.SelectedItem.ListSubItems(15)
                            Lote.InsereContribuinte
                            
                            'INSERE IMOVEL
                            Lote.CarregaDadosImovel InscricaoCadastral, "", lstCond.SelectedItem, lstCond.SelectedItem.ListSubItems(19), _
                                    InscricaoCadastral & txtIc(4), "", txtCodLogrBc, txtCodBairro, _
                                    lstCond.SelectedItem.ListSubItems(3), lstCond.SelectedItem.ListSubItems(4), _
                                    Trim(txtLoteBc), Trim(txtQuadraBc), Trim(txtLoteamentoBc), lstCond.SelectedItem.ListSubItems(2), _
                                    lstCond.SelectedItem.ListSubItems(16), lstCond.SelectedItem.ListSubItems(17), _
                                    Nvl(txtCodMens, 0), Nvl(txtZona, 1)
                            Lote.InsereTerritorio
                            
                            'INSERE COD. COBRANÇA
                            Call Lote.GravaComponente(InscricaoCadastral, lstCond.SelectedItem, lstCond.SelectedItem.ListSubItems(18), 3, 0)
                        Next
                    End If
                End If
                atualizarEnderecoContribuinte InscricaoCadastral, txtTipoLogrBt, txtLogrBt, txtNumero, txtComplemento, txtBairroBt
                atualizarContribuinte InscricaoCadastral, InscricaoMunicipal
                'LIMPA TELA
                Informa "Registro gravado com sucesso."
                Call cmd_Click(0)
                DoEvents
        Case "&Novo"
            Call Edita.LimpaCampos(Me)
            cboCobrancaBc.Tag = ""
            cboCobranca.Tag = "3"
            lstEdific.ListItems.Clear
            lstCond.ListItems.Clear
            tabCad.Tab = 0
            Unidades = 0
            Screen.MousePointer = 0
            For i = 0 To 4
                txtIc(i).Enabled = True
            Next
            txtIc(0).SetFocus
        Case "Sai&r"
            NovoContrib = True
            Unload Me
    End Select
End Sub

Private Function atualizarEnderecoContribuinte(Ic As String, Tipo As String, Logradouro As String, Numero As String, Complemento As String, Bairro As String) As Boolean
    Dim Sql As String
    
    Sql = "UPDATE TAB_CONTRIBUINTE " & _
            " SET tci_logradouro='" & Tipo & "', " & _
                " tci_nome_logradouro='" & Logradouro & "', " & _
                " tci_numero='" & Numero & "', " & _
                " tci_complemento='" & Complemento & "', " & _
                " tci_bairro='" & Bairro & "'" & _
            " WHERE tci_tim_ic='" & Ic & "'"
    atualizarEnderecoContribuinte = Bdados.Executa(Sql)
End Function

Private Function atualizarContribuinte(Ic As String, IM As String) As Boolean
    Dim Sql As String
    
    Sql = "UPDATE TAB_GERACAO_TRIBUTO" & _
            " SET tgt_im='" & IM & "'" & _
            " WHERE tgt_tim_ic='" & Ic & "'"
    atualizarContribuinte = Bdados.Executa(Sql)
    
    Sql = "UPDATE TAB_DARM_RECEBIDO" & _
            " SET tdr_im='" & IM & "'" & _
            " WHERE tdr_tim_ic='" & Ic & "'"
    atualizarContribuinte = atualizarContribuinte And Bdados.Executa(Sql)
End Function

