VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Begin VB.Form TCIU101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU101"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   247
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1138
      Icone           =   "TCIU101.frx":0000
   End
   Begin TabDlg.SSTab tabCad 
      Height          =   6015
      Left            =   45
      TabIndex        =   104
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
      TabCaption(0)   =   "BT (Boletim Territorial)"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstPesq"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "BT (cont.)"
      TabPicture(1)   =   "TCIU101.frx":282A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra(4)"
      Tab(1).Control(1)=   "fra(3)"
      Tab(1).Control(2)=   "fra(5)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "BP (Boletim Predial)"
      TabPicture(2)   =   "TCIU101.frx":2846
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra(6)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "BC (Boletim de Condom?nio)"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra(2)"
      Tab(3).Control(1)=   "fra(7)"
      Tab(3).Control(2)=   "fra(8)"
      Tab(3).Control(3)=   "lstCond"
      Tab(3).Control(4)=   "cmdAdCond"
      Tab(3).ControlCount=   5
      Begin Threed.SSFrame fra 
         Height          =   1875
         Index           =   0
         Left            =   120
         TabIndex        =   105
         Top             =   540
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   3307
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
         Caption         =   "Dados do Im?vel"
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
            Left            =   8010
            TabIndex        =   245
            Top             =   240
            Visible         =   0   'False
            Width           =   1965
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
            Left            =   6810
            MaxLength       =   5
            TabIndex        =   11
            Top             =   1020
            Width           =   1005
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
            Left            =   10440
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Zona"
            Top             =   240
            Width           =   555
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
            Left            =   4980
            TabIndex        =   5
            Top             =   240
            Width           =   1605
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
            Left            =   3360
            MaxLength       =   3
            TabIndex        =   4
            Tag             =   "Unidade"
            Top             =   240
            Width           =   375
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
            Left            =   2850
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
            Index           =   2
            Left            =   2340
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
            Index           =   1
            Left            =   2010
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
            Index           =   0
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   0
            Tag             =   "Distrito"
            Top             =   240
            Width           =   315
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
            Left            =   10665
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "Cod Mensagem"
            Top             =   1380
            Width           =   315
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
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   10
            Tag             =   "Bairro"
            Top             =   1012
            Width           =   525
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
            Left            =   2850
            MaxLength       =   11
            TabIndex        =   17
            Top             =   630
            Width           =   1140
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
            Left            =   4050
            TabIndex        =   18
            Tag             =   "Nome Contribuinte"
            Top             =   630
            Width           =   3765
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
            Left            =   1740
            TabIndex        =   19
            Tag             =   "Nome Contribuinte"
            Top             =   1020
            Width           =   4215
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
            Left            =   1170
            TabIndex        =   7
            Tag             =   "Logradouro"
            Top             =   630
            Width           =   1185
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
            Left            =   9570
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Tag             =   "Tipo Imovel"
            Top             =   1005
            Width           =   1455
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
            Left            =   10440
            MaxLength       =   10
            TabIndex        =   9
            Top             =   630
            Width           =   555
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
            Left            =   8580
            TabIndex        =   8
            Top             =   630
            Width           =   1425
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
            Left            =   6360
            MaxLength       =   5
            TabIndex        =   14
            Top             =   1410
            Width           =   555
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
            Left            =   7440
            MaxLength       =   5
            TabIndex        =   15
            Top             =   1410
            Width           =   615
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   2
            Left            =   7890
            TabIndex        =   106
            Top             =   675
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   1
            Left            =   10140
            TabIndex        =   107
            Top             =   675
            Width           =   390
            _ExtentX        =   688
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
            Caption         =   "N.?:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   3
            Left            =   570
            TabIndex        =   108
            Top             =   1065
            Width           =   705
            _ExtentX        =   1244
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
            Caption         =   "Bairro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   5
            Left            =   6030
            TabIndex        =   109
            Top             =   1455
            Width           =   660
            _ExtentX        =   1164
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
            Caption         =   "Qd.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   6
            Left            =   6990
            TabIndex        =   110
            Top             =   1455
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   7
            Left            =   9090
            TabIndex        =   120
            Top             =   1065
            Width           =   450
            _ExtentX        =   794
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
            Caption         =   "Tipo:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   82
            Left            =   270
            TabIndex        =   124
            Top             =   675
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
            Caption         =   "C?d. Logr:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   50
            Left            =   9675
            TabIndex        =   125
            Top             =   1425
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   72
            Left            =   3810
            TabIndex        =   126
            Top             =   285
            Width           =   1185
            _ExtentX        =   2090
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
            Caption         =   "Insc. Anterior:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   75
            Left            =   285
            TabIndex        =   127
            Top             =   285
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "Insc. Imobili?ria:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   76
            Left            =   10035
            TabIndex        =   128
            Top             =   285
            Width           =   465
            _ExtentX        =   820
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
            Caption         =   "Zona:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   345
            Index           =   2
            Left            =   2370
            TabIndex        =   238
            Top             =   600
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cboVISUAL cboLoteamento 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   1410
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   556
            Caption         =   "Loteamento"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   91
            Left            =   6210
            TabIndex        =   244
            Top             =   1065
            Width           =   525
            _ExtentX        =   926
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
            Caption         =   "Sec?o:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   4
            Left            =   6660
            TabIndex        =   246
            Top             =   285
            Visible         =   0   'False
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
            Caption         =   "Cad. Imobili?rio:"
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
         Left            =   90
         TabIndex        =   111
         Top             =   3030
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
         Caption         =   "03 - Dados do Propriet?rio"
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
            Left            =   10380
            MaxLength       =   50
            TabIndex        =   40
            Tag             =   "Bairro"
            Top             =   960
            Width           =   585
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
            Left            =   2280
            MaxLength       =   11
            TabIndex        =   33
            Top             =   570
            Width           =   1035
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
            Left            =   1005
            TabIndex        =   32
            Tag             =   "Logradouro"
            Top             =   585
            Width           =   735
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
            Left            =   5850
            TabIndex        =   39
            Tag             =   "Munic?pio"
            Top             =   970
            Width           =   4335
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
            Left            =   6195
            MaxLength       =   10
            TabIndex        =   35
            Top             =   590
            Width           =   465
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
            Left            =   9210
            MaxLength       =   20
            TabIndex        =   31
            Top             =   210
            Width           =   1785
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
            Left            =   8940
            MaxLength       =   20
            TabIndex        =   42
            Top             =   1350
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
            Left            =   3030
            TabIndex        =   41
            Top             =   1350
            Width           =   4965
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
            Left            =   8940
            TabIndex        =   37
            Tag             =   "Bairro"
            Top             =   590
            Width           =   2055
         End
         Begin VB.CommandButton cmdEnter 
            Caption         =   "Command1"
            Default         =   -1  'True
            Height          =   255
            Left            =   7740
            TabIndex        =   117
            Top             =   3090
            Width           =   375
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
            TabIndex        =   38
            Top             =   970
            Width           =   1125
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
            Left            =   3390
            TabIndex        =   34
            Tag             =   "Nome Logradouro"
            Top             =   590
            Width           =   2415
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
            Left            =   1560
            TabIndex        =   29
            Tag             =   "Nome Contribuinte"
            Top             =   210
            Width           =   4665
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
            Left            =   6570
            MaxLength       =   14
            TabIndex        =   30
            Top             =   210
            Width           =   1305
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
            Left            =   7530
            TabIndex        =   36
            Top             =   590
            Width           =   735
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   8
            Left            =   975
            TabIndex        =   112
            Top             =   255
            Width           =   525
            _ExtentX        =   926
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
            Caption         =   "Nome:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   14
            Left            =   2640
            TabIndex        =   113
            Top             =   1015
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
            Caption         =   "CEP:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   15
            Left            =   4965
            TabIndex        =   114
            Top             =   1015
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Municipio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   13
            Left            =   5910
            TabIndex        =   115
            Top             =   630
            Width           =   270
            _ExtentX        =   476
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
            Caption         =   "N.?:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   17
            Left            =   8340
            TabIndex        =   116
            Top             =   635
            Width           =   555
            _ExtentX        =   979
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
            Caption         =   "Bairro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   16
            Left            =   6750
            TabIndex        =   119
            Top             =   630
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
            Caption         =   "Compl.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   29
            Left            =   8355
            TabIndex        =   121
            Top             =   255
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   11
            Left            =   2175
            TabIndex        =   122
            Top             =   1395
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "Ocupante:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   18
            Left            =   8055
            TabIndex        =   123
            Top             =   1395
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   49
            Left            =   510
            TabIndex        =   234
            Top             =   630
            Width           =   450
            _ExtentX        =   794
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
            Caption         =   "Logr:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   345
            Index           =   1
            Left            =   540
            TabIndex        =   28
            Top             =   195
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            Caption         =   ""
            Acao            =   6
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   345
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   195
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   345
            Index           =   3
            Left            =   1770
            TabIndex        =   239
            Top             =   570
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   90
            Left            =   6270
            TabIndex        =   243
            Top             =   270
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
            Caption         =   "IM:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin MSComctlLib.ListView lstPesq 
         Height          =   1065
         Left            =   90
         TabIndex        =   129
         Top             =   4860
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   1879
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
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
         TabIndex        =   130
         Top             =   540
         Width           =   11085
         _ExtentX        =   19553
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
         Caption         =   "Caracter?sticas do Im?vel:"
         Alignment       =   2
         ShadowStyle     =   1
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
            Left            =   7200
            MaxLength       =   3
            TabIndex        =   51
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
            Index           =   3
            Left            =   7200
            MaxLength       =   3
            TabIndex        =   49
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
            Index           =   2
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   47
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
            Index           =   1
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   45
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
            Index           =   0
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   43
            Top             =   270
            Width           =   375
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
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   44
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   270
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
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   46
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   630
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
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   48
            TabStop         =   0   'False
            Tag             =   "3"
            Top             =   990
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
            Left            =   7620
            Style           =   2  'Dropdown List
            TabIndex        =   50
            TabStop         =   0   'False
            Tag             =   "4"
            Top             =   270
            Width           =   3375
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
            Left            =   7620
            Style           =   2  'Dropdown List
            TabIndex        =   52
            TabStop         =   0   'False
            Tag             =   "5"
            Top             =   630
            Width           =   3375
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
            Left            =   7410
            Style           =   2  'Dropdown List
            TabIndex        =   132
            Top             =   2310
            Width           =   2535
         End
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
            Left            =   6930
            Style           =   2  'Dropdown List
            TabIndex        =   131
            Top             =   1950
            Width           =   3015
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   30
            Left            =   6090
            TabIndex        =   133
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
            Caption         =   "Arboriza??o:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   32
            Left            =   6420
            TabIndex        =   134
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
            TabIndex        =   135
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
            Caption         =   "Cod. Cobran?a:"
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
            TabIndex        =   136
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
            Caption         =   "Ocupa??o do Lote:"
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
            TabIndex        =   137
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
            Caption         =   "Patrim?nio:"
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
            TabIndex        =   138
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
            Caption         =   "Instala??o Sanit?ria:"
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
            TabIndex        =   139
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
            Caption         =   "Instala??o El?trica:"
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
         Left            =   -74880
         TabIndex        =   140
         Top             =   1950
         Width           =   11085
         _ExtentX        =   19553
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
         Caption         =   "Caracter?sticas do Terreno:"
         Alignment       =   2
         ShadowStyle     =   1
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
            TabIndex        =   55
            Top             =   570
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
            Left            =   7200
            MaxLength       =   3
            TabIndex        =   57
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
            Index           =   5
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   53
            Top             =   240
            Width           =   375
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
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   54
            TabStop         =   0   'False
            Tag             =   "6"
            Top             =   240
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
            Left            =   7590
            Style           =   2  'Dropdown List
            TabIndex        =   58
            TabStop         =   0   'False
            Tag             =   "8"
            Top             =   240
            Width           =   3405
         End
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
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   56
            TabStop         =   0   'False
            Tag             =   "7"
            Top             =   600
            Width           =   3375
         End
         Begin Threed.SSPanel lbl 
            Height          =   315
            Index           =   20
            Left            =   690
            TabIndex        =   141
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
            TabIndex        =   142
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
            Caption         =   "Situa??o:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   240
            Index           =   22
            Left            =   6270
            TabIndex        =   143
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
         Left            =   -74880
         TabIndex        =   144
         Top             =   2940
         Width           =   11085
         _ExtentX        =   19553
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
         Caption         =   "Dimens?es do Terreno (m?)"
         Alignment       =   2
         ShadowStyle     =   1
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
            TabIndex        =   59
            Tag             =   "100"
            Top             =   240
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
            TabIndex        =   60
            Tag             =   "101"
            Top             =   570
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
            TabIndex        =   61
            Tag             =   "102"
            Top             =   210
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
            TabIndex        =   62
            Tag             =   "103"
            Top             =   540
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
            TabIndex        =   65
            Tag             =   "106"
            Top             =   180
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
            TabIndex        =   63
            Tag             =   "104"
            Top             =   240
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
            TabIndex        =   64
            Tag             =   "105"
            Top             =   600
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
            Left            =   10230
            TabIndex        =   67
            Tag             =   "108"
            Top             =   210
            Width           =   735
         End
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
            TabIndex        =   66
            Tag             =   "107"
            Top             =   510
            Width           =   735
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   23
            Left            =   4590
            TabIndex        =   145
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
            TabIndex        =   146
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
            TabIndex        =   147
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
            TabIndex        =   148
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
            TabIndex        =   149
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
            TabIndex        =   150
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
            TabIndex        =   151
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
            TabIndex        =   152
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
            Left            =   9090
            TabIndex        =   153
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
            Caption         =   "?rea do Lote:"
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
         TabIndex        =   154
         Top             =   540
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
         Caption         =   "Caracter?sticas das Edifica??es"
         Alignment       =   2
         ShadowStyle     =   1
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
            TabIndex        =   166
            Top             =   210
            Visible         =   0   'False
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
            Index           =   12
            Left            =   2310
            MaxLength       =   4
            TabIndex        =   165
            Top             =   210
            Visible         =   0   'False
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
            Index           =   11
            Left            =   1980
            MaxLength       =   2
            TabIndex        =   164
            Top             =   210
            Visible         =   0   'False
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
            Index           =   10
            Left            =   1650
            MaxLength       =   2
            TabIndex        =   163
            Top             =   210
            Visible         =   0   'False
            Width           =   315
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
            Left            =   1620
            MaxLength       =   3
            TabIndex        =   74
            Top             =   2655
            Width           =   375
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
            Left            =   6990
            TabIndex        =   78
            Tag             =   "111"
            Top             =   1500
            Width           =   1185
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
            Left            =   6990
            TabIndex        =   80
            Tag             =   "113"
            Top             =   2258
            Width           =   1155
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
            Left            =   6990
            MaxLength       =   3
            TabIndex        =   76
            Top             =   728
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
            Left            =   6990
            MaxLength       =   3
            TabIndex        =   77
            Top             =   1121
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
            Left            =   6990
            MaxLength       =   3
            TabIndex        =   75
            Top             =   210
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
            Left            =   1620
            MaxLength       =   3
            TabIndex        =   73
            Top             =   2265
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
            TabIndex        =   69
            Top             =   728
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
            TabIndex        =   70
            Top             =   1121
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
            Index           =   15
            Left            =   1620
            MaxLength       =   3
            TabIndex        =   72
            Top             =   1875
            Width           =   375
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
            TabIndex        =   68
            Top             =   210
            Width           =   735
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
            TabIndex        =   71
            Tag             =   "110"
            Top             =   1500
            Width           =   825
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
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   162
            TabStop         =   0   'False
            Tag             =   "15"
            Top             =   1113
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
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   161
            TabStop         =   0   'False
            Tag             =   "16"
            Top             =   1860
            Width           =   3615
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
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   160
            TabStop         =   0   'False
            Tag             =   "14"
            Top             =   720
            Width           =   3615
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
            Left            =   6990
            TabIndex        =   81
            Tag             =   "114"
            Top             =   2648
            Width           =   1155
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
            Left            =   6990
            TabIndex        =   79
            Tag             =   "112"
            Top             =   1868
            Width           =   1185
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
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   159
            TabStop         =   0   'False
            Tag             =   "10"
            Top             =   2640
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
            Left            =   7380
            Style           =   2  'Dropdown List
            TabIndex        =   158
            TabStop         =   0   'False
            Tag             =   "11"
            Top             =   202
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
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   157
            TabStop         =   0   'False
            Tag             =   "9"
            Top             =   2250
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
            Left            =   7380
            Style           =   2  'Dropdown List
            TabIndex        =   156
            TabStop         =   0   'False
            Tag             =   "12"
            Top             =   720
            Width           =   3615
         End
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
            Left            =   7380
            Style           =   2  'Dropdown List
            TabIndex        =   155
            TabStop         =   0   'False
            Tag             =   "13"
            Top             =   1113
            Width           =   3615
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   36
            Left            =   735
            TabIndex        =   167
            Top             =   2700
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   38
            Left            =   750
            TabIndex        =   168
            Top             =   2310
            Width           =   810
            _ExtentX        =   1429
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
            Caption         =   "Tipologia:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   39
            Left            =   5940
            TabIndex        =   169
            Top             =   255
            Width           =   960
            _ExtentX        =   1693
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
            Caption         =   "Destina??o:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   40
            Left            =   6270
            TabIndex        =   170
            Top             =   773
            Width           =   630
            _ExtentX        =   1111
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
            Caption         =   "Padr?o:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   41
            Left            =   5760
            TabIndex        =   171
            Top             =   1166
            Width           =   1140
            _ExtentX        =   2011
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
            Caption         =   "Conserva??o:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   42
            Left            =   5850
            TabIndex        =   172
            Top             =   1545
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   44
            Left            =   5610
            TabIndex        =   173
            Top             =   2303
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
            Caption         =   "?rea Edif. Total:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   47
            Left            =   870
            TabIndex        =   174
            Top             =   780
            Width           =   690
            _ExtentX        =   1217
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
            Caption         =   "Sentido:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   48
            Left            =   1185
            TabIndex        =   175
            Top             =   1920
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
            Caption         =   "Uso:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   52
            Left            =   960
            TabIndex        =   176
            Top             =   1170
            Width           =   600
            _ExtentX        =   1058
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
            Caption         =   "Predio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   53
            Left            =   525
            TabIndex        =   177
            Top             =   1560
            Width           =   1035
            _ExtentX        =   1826
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
            Caption         =   "Pavimentos:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin MSComctlLib.ListView lstEdific 
            Height          =   2115
            Left            =   120
            TabIndex        =   178
            Top             =   3090
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   3731
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
            Appearance      =   0
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
               Text            =   "Insc. Imobili?ria"
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
               Text            =   "Pr?dio"
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
               Text            =   "Destina??o"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "12"
               Text            =   "Padr?o"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Tag             =   "13"
               Text            =   "Conserva??o"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Object.Tag             =   "111"
               Text            =   "?rea Constr."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Object.Tag             =   "112"
               Text            =   "?rea Edificada"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Object.Tag             =   "113"
               Text            =   "?rea Edificada Total"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Object.Tag             =   "114"
               Text            =   "Fra??o Ideal"
               Object.Width           =   2540
            EndProperty
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   9
            Left            =   210
            TabIndex        =   179
            Top             =   255
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "Sub-Unidade:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   43
            Left            =   6060
            TabIndex        =   180
            Top             =   1913
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "?rea Edif.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   46
            Left            =   5865
            TabIndex        =   181
            Top             =   2693
            Width           =   1035
            _ExtentX        =   1826
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
            Caption         =   "Fra??o Ideal:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdAdEdif 
            Height          =   375
            Left            =   8685
            TabIndex        =   82
            Top             =   2595
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   661
            Caption         =   "&Adicionar Edifica??o"
            Acao            =   1
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1815
         Index           =   2
         Left            =   -74880
         TabIndex        =   182
         Top             =   540
         Width           =   11085
         _ExtentX        =   19553
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
         Caption         =   "Refer?ncia Cadastral / Localiza??o do Im?vel"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VTOcx.cboVISUAL cboEdificio 
            Height          =   315
            Left            =   630
            TabIndex        =   87
            Top             =   1020
            Width           =   6525
            _ExtentX        =   11509
            _ExtentY        =   556
            Caption         =   "Edificio"
            Text            =   ""
            AutoFocaliza    =   0   'False
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
            Left            =   9450
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Tag             =   "Logradouro"
            Top             =   277
            Width           =   1545
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
            Left            =   9960
            MaxLength       =   10
            TabIndex        =   193
            Top             =   1410
            Width           =   1035
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
            TabIndex        =   88
            Top             =   667
            Width           =   525
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
            Left            =   8790
            TabIndex        =   86
            Top             =   667
            Width           =   2205
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
            TabIndex        =   192
            Top             =   1410
            Width           =   705
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
            TabIndex        =   191
            Top             =   1410
            Width           =   705
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
            TabIndex        =   190
            Top             =   1410
            Width           =   765
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
            Left            =   540
            MaxLength       =   2
            TabIndex        =   189
            Tag             =   "Distrito"
            Top             =   285
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
            Index           =   6
            Left            =   1260
            MaxLength       =   2
            TabIndex        =   188
            Tag             =   "Setor"
            Top             =   285
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
            Index           =   7
            Left            =   2310
            MaxLength       =   4
            TabIndex        =   187
            Tag             =   "Quadra"
            Top             =   285
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
            Index           =   8
            Left            =   3300
            MaxLength       =   4
            TabIndex        =   186
            Tag             =   "Lote"
            Top             =   285
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
            Index           =   9
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   83
            Tag             =   "Unidade"
            Top             =   285
            Width           =   375
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
            Left            =   6120
            TabIndex        =   84
            Top             =   285
            Width           =   1665
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
            TabIndex        =   103
            Top             =   667
            Width           =   1485
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
            TabIndex        =   185
            Top             =   667
            Width           =   1035
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
            TabIndex        =   184
            Tag             =   "Nome Contribuinte"
            Top             =   667
            Width           =   3255
         End
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
            TabIndex        =   183
            Tag             =   "Nome Contribuinte"
            Top             =   1410
            Width           =   3525
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   10
            Left            =   90
            TabIndex        =   194
            Top             =   712
            Width           =   1080
            _ExtentX        =   1905
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
            Caption         =   "Cod. Logr:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   54
            Left            =   8100
            TabIndex        =   195
            Top             =   712
            Width           =   660
            _ExtentX        =   1164
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   55
            Left            =   7200
            TabIndex        =   196
            Top             =   712
            Width           =   390
            _ExtentX        =   688
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
            Caption         =   "N.?:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   56
            Left            =   5010
            TabIndex        =   197
            Top             =   1455
            Width           =   705
            _ExtentX        =   1244
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
            Caption         =   "Bairro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   57
            Left            =   90
            TabIndex        =   198
            Top             =   1455
            Width           =   1170
            _ExtentX        =   2064
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
            Caption         =   "Loteamento:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   58
            Left            =   2040
            TabIndex        =   199
            Top             =   1455
            Width           =   660
            _ExtentX        =   1164
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   59
            Left            =   3600
            TabIndex        =   200
            Top             =   1455
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
            Caption         =   "Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   60
            Left            =   9540
            TabIndex        =   201
            Top             =   1455
            Width           =   360
            _ExtentX        =   635
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
            Caption         =   "CEP:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   61
            Left            =   8970
            TabIndex        =   202
            Top             =   330
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
            Caption         =   "Tipo:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   203
            Top             =   330
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
            AutoSize        =   2
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   73
            Left            =   930
            TabIndex        =   204
            Top             =   330
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
            AutoSize        =   2
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   83
            Left            =   1650
            TabIndex        =   205
            Top             =   330
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
            AutoSize        =   2
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   84
            Left            =   2820
            TabIndex        =   206
            Top             =   330
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
            AutoSize        =   2
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   85
            Left            =   3900
            TabIndex        =   207
            Top             =   330
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
            AutoSize        =   2
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   86
            Left            =   4890
            TabIndex        =   208
            Top             =   330
            Width           =   1185
            _ExtentX        =   2090
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
            Caption         =   "Insc. Anterior:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VB.Shape Shape2 
            Height          =   405
            Left            =   90
            Top             =   240
            Width           =   4695
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1785
         Index           =   7
         Left            =   -74880
         TabIndex        =   209
         Top             =   2400
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
         Caption         =   "Dados do Propriet?rio"
         Alignment       =   2
         ShadowStyle     =   1
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
            TabIndex        =   98
            Tag             =   "Munic?pio"
            Top             =   970
            Width           =   4305
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
            TabIndex        =   94
            Top             =   590
            Width           =   525
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
            Left            =   10215
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Tag             =   "UF"
            Top             =   962
            Width           =   795
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
            Left            =   8940
            MaxLength       =   20
            TabIndex        =   91
            Top             =   210
            Width           =   2055
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
            Left            =   8940
            MaxLength       =   20
            TabIndex        =   101
            Top             =   1350
            Width           =   2055
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
            Left            =   3090
            TabIndex        =   100
            Top             =   1350
            Width           =   4875
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Tag             =   "Logradouro"
            Top             =   582
            Width           =   1305
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
            Left            =   8940
            TabIndex        =   96
            Tag             =   "Bairro"
            Top             =   590
            Width           =   2055
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
            Left            =   3090
            MaxLength       =   10
            TabIndex        =   97
            Top             =   970
            Width           =   1125
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
            Left            =   3090
            TabIndex        =   93
            Tag             =   "Nome Logradouro"
            Top             =   590
            Width           =   2415
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
            Left            =   3090
            TabIndex        =   90
            Tag             =   "Nome Contribuinte"
            Top             =   210
            Width           =   4875
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
            Left            =   1440
            MaxLength       =   11
            TabIndex        =   89
            Top             =   210
            Width           =   1305
         End
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
            Left            =   7230
            TabIndex        =   95
            Top             =   590
            Width           =   735
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   62
            Left            =   150
            TabIndex        =   210
            Top             =   255
            Width           =   1275
            _ExtentX        =   2249
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
            Caption         =   "Insc. Municipal:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   63
            Left            =   390
            TabIndex        =   211
            Top             =   635
            Width           =   1035
            _ExtentX        =   1826
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
            Caption         =   "Logradouro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   64
            Left            =   2670
            TabIndex        =   212
            Top             =   1020
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
            Caption         =   "CEP:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   65
            Left            =   4935
            TabIndex        =   213
            Top             =   1020
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Municipio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   66
            Left            =   5580
            TabIndex        =   214
            Top             =   635
            Width           =   270
            _ExtentX        =   476
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
            Caption         =   "N.?:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   67
            Left            =   8340
            TabIndex        =   215
            Top             =   630
            Width           =   555
            _ExtentX        =   979
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
            Caption         =   "Bairro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   68
            Left            =   6510
            TabIndex        =   216
            Top             =   635
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
            Caption         =   "Compl.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   69
            Left            =   8055
            TabIndex        =   217
            Top             =   255
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   70
            Left            =   2205
            TabIndex        =   218
            Top             =   1395
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "Ocupante:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   71
            Left            =   8055
            TabIndex        =   219
            Top             =   1395
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   675
         Index           =   8
         Left            =   -74880
         TabIndex        =   220
         Top             =   4185
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   1191
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
         Caption         =   "Caracter?sticas do Im?vel:"
         Alignment       =   2
         ShadowStyle     =   1
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
            Height          =   330
            Index           =   20
            Left            =   1530
            MaxLength       =   3
            TabIndex        =   102
            Top             =   225
            Width           =   420
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
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   223
            TabStop         =   0   'False
            Tag             =   "3"
            Top             =   225
            Width           =   6450
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
            Left            =   7410
            Style           =   2  'Dropdown List
            TabIndex        =   222
            Top             =   2310
            Width           =   2535
         End
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
            Left            =   6930
            Style           =   2  'Dropdown List
            TabIndex        =   221
            Top             =   1950
            Width           =   3015
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   74
            Left            =   195
            TabIndex        =   224
            Top             =   300
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
            Caption         =   "Cod. Cobran?a:"
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
            TabIndex        =   225
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
            Caption         =   "Instala??o Sanit?ria:"
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
            TabIndex        =   226
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
            Caption         =   "Instala??o El?trica:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   645
         Index           =   9
         Left            =   120
         TabIndex        =   227
         Top             =   2370
         Visible         =   0   'False
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   1138
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
            TabIndex        =   26
            Top             =   255
            Width           =   990
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
            TabIndex        =   25
            Top             =   247
            Width           =   585
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
            TabIndex        =   23
            Top             =   247
            Width           =   585
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
            TabIndex        =   24
            Top             =   255
            Width           =   1215
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
            Left            =   390
            MaxLength       =   5
            TabIndex        =   20
            Top             =   255
            Width           =   840
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
            TabIndex        =   21
            Top             =   255
            Width           =   705
         End
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
            TabIndex        =   22
            Top             =   255
            Width           =   615
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   79
            Left            =   4095
            TabIndex        =   228
            Top             =   300
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
            TabIndex        =   229
            Top             =   300
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
            Caption         =   "N?:"
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
            TabIndex        =   230
            Top             =   300
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
            TabIndex        =   231
            Top             =   300
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
            TabIndex        =   232
            Top             =   300
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
            Index           =   12
            Left            =   7230
            TabIndex        =   241
            Top             =   300
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
            Index           =   89
            Left            =   8820
            TabIndex        =   242
            Top             =   300
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
      Begin MSComctlLib.ListView lstCond 
         Height          =   1020
         Left            =   -74910
         TabIndex        =   233
         Top             =   4920
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   1799
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
         Appearance      =   0
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
            Text            =   "N?"
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
            Text            =   "N?"
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
            Text            =   "Cod. Cobran?a"
            Object.Width           =   2540
         EndProperty
      End
      Begin VTOcx.cmdVISUAL cmdAdCond 
         Height          =   375
         Left            =   -66120
         TabIndex        =   240
         Top             =   4350
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "Adicionar &Condom?nio"
         Acao            =   1
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.TextBox txtFatorFixo 
      Height          =   285
      Left            =   8640
      TabIndex        =   118
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   4560
      Width           =   375
   End
   Begin VTOcx.cmdVISUAL cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   2
      Left            =   10200
      TabIndex        =   235
      Top             =   6750
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
      Left            =   9030
      TabIndex        =   236
      Top             =   6750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   4
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   0
      Left            =   7860
      TabIndex        =   237
      Top             =   6750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCIU101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As VSImposto
Dim NovoContrib As Boolean
Dim Sql As String
Dim Lote As New BCI
Private Boletim As TipoBoletim

Private Function ImovelJaCadastrado()
    Dim rs As VSRecordset
    Dim Sql As String
    Dim aux As String
       
    If Trim(txtic(4)) = "" Then
        Screen.MousePointer = 0
        Exit Function
    End If
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        Sql = "Select * from tab_imovel where tim_ic_auxiliar ='" & txtic(0).Text & txtic(1).Text & txtic(2).Text & txtic(3).Text & txtic(4).Text & "'"
    Else
        Sql = "Select * from tab_imovel where (tIM_ic ='" & txtic(0).Text & txtic(1).Text & txtic(2).Text & txtic(3).Text & txtic(4).Text & "'" & IIf(CInt(txtic(4)) >= 200, " AND TIM_UNIDADE =" & txtic(4), "") & ") " & aux
    End If
    If Bdados.AbreTabela(Sql, rs) Then
        ImovelJaCadastrado = True
    End If
End Function

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

Function TotalProva(VALOR As String) As String
    Static Total As Double
    If Trim(VALOR) = "" Then VALOR = "0"
    Total = CDbl(VALOR) + Total
    TotalProva = Total
End Function

Public Sub HabilitaCaixa(Status As Boolean)
    txtIM.Enabled = Not Status
    txtNomeContrib.Enabled = Status
    txtNomeTipoLogrContrib.Enabled = Status
    txtNomeLogrContrib.Enabled = Status
    txtNumeroContrib.Enabled = Status
    txtCompContrib.Enabled = Status
    txtBairroContrib.Enabled = Status
    txtCEP.Enabled = Status
    txtMunic.Enabled = Status
    txtUF.Enabled = Status
    txtIM = ""
    txtNomeContrib = ""
    txtNomeTipoLogrContrib = ""
    txtNomeLogrContrib = ""
    txtNumeroContrib = ""
    txtCompContrib = ""
    txtBairroContrib = ""
    txtCEP = ""
    txtMunic = ""
    txtUF = ""
End Sub

Private Sub cboArborizacao_Click()
'    AtualizaCodComponente cboArborizacao
End Sub

Private Sub cboArborizacao_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboArborizacao
End Sub

Private Sub cboCobranca_Click()
'    AtualizaCodComponente cboCobranca
End Sub

Private Sub cboCobranca_KeyPress(KeyAscii As Integer)
'   AtualizaCodComponente cboCobranca
End Sub

Private Sub cboConservacao_Click()
'    AtualizaCodComponente cboConservacao
End Sub

Private Sub cboConservacao_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboConservacao
End Sub

Private Sub cboDestinacao_Click()
    'AtualizaCodComponente cboDestinacao
End Sub

Private Sub cboDestinacao_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboDestinacao
End Sub

Private Sub cboEstrutura_Click()
    'AtualizaCodComponente cboEstrutura
End Sub

Private Sub cboEstrutura_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboEstrutura
End Sub

Private Sub cboLimites_Click()
    'AtualizaCodComponente cboLimites
End Sub

Private Sub cboLimites_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboLimites
End Sub

Private Sub cboOcupLote_Click()
    'AtualizaCodComponente cboOcupLote
End Sub

Private Sub cboOcupLote_KeyPress(KeyAscii As Integer)
    'cboOcupLote_Click
End Sub

Private Sub cboPadrao_Click()
    'AtualizaCodComponente cboPadrao
End Sub

Private Sub cboPadrao_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboPadrao
End Sub

Private Sub cboPatrimonio_Click()
    'AtualizaCodComponente cboPatrimonio
End Sub

Private Sub cboPatrimonio_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboPatrimonio
End Sub

Private Sub cboPedol_Click()
    'AtualizaCodComponente cboPedol
End Sub

Private Sub cboPedol_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboPedol
End Sub

Private Sub cboPredio_Click()
    'AtualizaCodComponente cboPredio
End Sub

Private Sub cboPredio_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboPredio
End Sub

Private Sub cboSentido_Click()
    'AtualizaCodComponente cboSentido
End Sub

Private Sub cboSentido_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboSentido
End Sub

Private Sub cboSit_Click()
    'AtualizaCodComponente cboSit
End Sub

Private Sub cboSit_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboSit
End Sub

Private Sub cboTipoImovel_Click()
    If cboTipoImovel = "PREDIAL" Then
        tabCad.TabEnabled(0) = True
        tabCad.TabEnabled(1) = True
        tabCad.TabEnabled(2) = True
        tabCad.TabEnabled(3) = True
        Boletim = tbo_Predial
    Else
        tabCad.TabEnabled(2) = False
        tabCad.TabEnabled(3) = False
        
        Boletim = tbo_Territorial
    End If

End Sub

Private Sub cboTipologia_Click()
    'AtualizaCodComponente cboTipologia
End Sub

Private Sub cboTipologia_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboTipologia
End Sub

Private Sub cboTopogr_Click()
    'AtualizaCodComponente cboTopogr
End Sub

Private Sub cboTopogr_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboTopogr
End Sub

Private Sub cboUso_Click()
    'AtualizaCodComponente cboUso
End Sub

Private Sub cboUso_KeyPress(KeyAscii As Integer)
    'AtualizaCodComponente cboUso
    
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
    Dim InscricaoReduzida    As String
    Dim CodLogr As Long
    Dim DtVenc As String
    Dim SitCadastral As String
    Static Unidades As Integer
    Dim Conta As New ContaCorrente
    Dim i As Integer
    Dim j As Integer
    Dim cadastro As New VSImposto
    Select Case cmd(Index).Caption
        Case "&Salvar"
                If ImovelJaCadastrado Then
                    Util.Informa "Este im?vel j? est? cadastrado"
                    txtic(4).SetFocus
                    tabCad.Tab = 0
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                If Trim(txtCodBairro) = "" Then
                    Util.Informa "Falta a defini??o do bairro."
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
                InscricaoCadastral = txtic(0) & txtic(1) & txtic(2) & txtic(3) & txtic(4)
                
                InscricaoMunicipal = txtIM
                Screen.MousePointer = 11
                If Not Lote.VerificaFechamentoAreas(lstEdific) Then Exit Sub
                'GRAVANDO BT
                Lote.CarregaDadosContribuinte InscricaoMunicipal, txtNomeContrib, txtCpfCgc, txtCodLogrContrib, txtNomeTipoLogrContrib, txtNomeLogrContrib, _
                        txtNumeroContrib, txtCompContrib, "", txtBairroContrib, txtCEP, txtMunic, txtUF
                If Not Lote.InsereContribuinte(NovoContrib) Then Exit Sub
                
'                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
'                    Lote.ApagaImovel InscricaoReduzida
'                    Lote.ApagaDetalheImovel InscricaoReduzida
'                Else
'                    Lote.ApagaImovel InscricaoCadastral
'                    Lote.ApagaDetalheImovel InscricaoCadastral
'                End If
                
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                    InscricaoReduzida = Conta.GeraCodPagamento("CADASTRO IMOBILIARIO")
                    Lote.CarregaDadosImovel InscricaoCadastral, txtICAnterior, txtic(4), "0", "", "", CStr(CodLogr), txtCodBairro, _
                         txtNumero, txtComplemento, txtLote, txtQuadra, CStr(cboLoteamento.Coluna(0).VALOR), Boletim, txtOcupante, txtCPFOcupante, _
                         txtCodMens, txtZona, txtNumAforamento, txtFichaAforamento, txtLivroAforamento, txtFolhaAforamento, txtRegistro, txtDataAforamento, txtDtRegistro, , InscricaoReduzida, txtSecao, CStr(cboEdificio.Coluna(0).VALOR)
                Else
                    Lote.CarregaDadosImovel InscricaoCadastral, txtICAnterior, txtic(4), "0", "", "", CStr(CodLogr), txtCodBairro, _
                         txtNumero, txtComplemento, txtLote, txtQuadra, CStr(cboLoteamento.Coluna(0).VALOR), Boletim, txtOcupante, txtCPFOcupante, _
                         txtCodMens, txtZona, txtNumAforamento, txtFichaAforamento, txtLivroAforamento, txtFolhaAforamento, txtRegistro, txtDataAforamento, txtDtRegistro
                End If
                If Not Lote.InsereTerritorio() Then Exit Sub
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                    Call Lote.GravaComponentes(InscricaoReduzida, Me, 1, 8, False, Nvl(txtic(4), 0), 0)
                    Call Lote.GravaComponentes(InscricaoReduzida, Me, 100, 109, True, Nvl(txtic(4), 0), 0)
                Else
                    Call Lote.GravaComponentes(InscricaoCadastral, Me, 1, 8, False, Nvl(txtic(4), 0), 0)
                    Call Lote.GravaComponentes(InscricaoCadastral, Me, 100, 109, True, Nvl(txtic(4), 0), 0)
                End If
                'GRAVANDO BP
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                    Lote.GravaBP lstEdific, txtCodMens, InscricaoReduzida, 0
                Else
                    Lote.GravaBP lstEdific, txtCodMens, txtic(0) & txtic(1) & txtic(2) & txtic(3), txtic(4)
                End If
                'GRAVANDO BC'S
                If CInt(Nvl(Trim(txtic(4)), 0)) >= 200 Then
                    cboCobrancaBc.Tag = "3"
                    cboCobranca.Tag = ""
                    If lstCond.ListItems.Count > 0 Then
                        For j = 1 To lstCond.ListItems.Count 'Para cada edificacao
                            lstCond.ListItems(j).Selected = True
                            InscricaoMunicipal = txtIMBc
                            CodLogr = txtCodLogr
                            InscricaoCadastral = txtic(0) & txtic(1) & txtic(2) & txtic(3) & lstCond.SelectedItem
                            'INSERE CONTRIBUINTE
                            If lstCond.SelectedItem.ListSubItems(5) = "" Then
                                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                                   InscricaoMunicipal = Conta.GeraCodPagamento("CADASTRO ECONOMICO")
                                Else
                                    InscricaoMunicipal = cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
                                End If
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
                                    InscricaoCadastral & txtic(4), "", txtCodLogrBc, txtCodBairro, _
                                    lstCond.SelectedItem.ListSubItems(3), lstCond.SelectedItem.ListSubItems(4), _
                                    Trim(txtLoteBc), Trim(txtQuadraBc), Trim(txtLoteamentoBc), lstCond.SelectedItem.ListSubItems(2), _
                                    lstCond.SelectedItem.ListSubItems(16), lstCond.SelectedItem.ListSubItems(17), _
                                    Nvl(txtCodMens, 0), Nvl(txtZona, 1), , , , , , , , , InscricaoReduzida, txtSecao, CStr(cboEdificio.Coluna(0).VALOR)
                            Lote.InsereTerritorio
                            
                            'INSERE COD. COBRAN?A
                            Call Lote.GravaComponente(InscricaoCadastral, lstCond.SelectedItem, lstCond.SelectedItem.ListSubItems(18), 3, 0)
                        Next
                    End If
                End If
                'LIMPA TELA
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                    Informa "Registro gravado com sucesso. Cadastro gerado: " & InscricaoReduzida & "."
                Else
                    Informa "Registro gravado com sucesso."
                End If
                Call cmd_Click(0)
                DoEvents
        Case "&Novo"
            NovoContrib = True
            Call Edita.LimpaCampos(Me)
            cboCobrancaBc.Tag = ""
            cboCobranca.Tag = "3"
            lstEdific.ListItems.Clear
            lstCond.ListItems.Clear
            tabCad.Tab = 0
                  txtCEP = "" & Temp.PegaParametro(Bdados, "CEP MUNICIPIO")
            Unidades = 0
            Screen.MousePointer = 0
            txtic(0).SetFocus
        Case "Sai&r"
            NovoContrib = True
            Unload Me
    End Select
End Sub


Private Sub cmdAdCond_Click()
    'NOVIDADE
    Dim ItmX As Object
    Dim i As Byte
    
    On Error Resume Next
    Set ItmX = lstCond.ListItems.Add(, , txtic(9))
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
    txtic(9).SetFocus
End Sub

Private Sub cmdAdEdif_Click()
    On Error GoTo trata
    Dim ItmX As Object
    Dim i As Byte
    If Trim(txtInscImobiliaria) = "" Then
        Informa "Informe a unidade."
        txtInscImobiliaria.SetFocus
        Exit Sub
    End If
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
    txtInscImobiliaria = Format(CInt(txtInscImobiliaria) + 1, "000")
    txtCodComponente(13).SetFocus
trata:
End Sub

Private Sub cmdEnter_Click()
        SendKeys "{Tab}"
End Sub

Private Sub cmdNovo_Click()
    Static Status As Boolean
    Status = Not Status
    HabilitaCaixa Status
End Sub

Private Sub cmdOpcao_Click(Index As Integer)
    Dim rs As VSRecordset
    Dim Sql As String
    Select Case Index
        Case 0
            NovoContrib = False
            Sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where tci_nome like '" & txtNomeContrib & "%' or tci_nome like '% " & txtNomeContrib & "%'"
            Sql = Sql & " and tci_tsc_cod_sit_cad =1"
            If Not Bdados.AbreTabela(Sql, rs) Then
                Call Util.Avisa("Nenhum contribuinte encontrado.")
                SendKeys "{tab}"
            End If
            Bdados.FechaTabela rs
            MontaGrid Bdados, lstPesq, Sql, 1400
        Case 1
            NovoContrib = True
            txtIM = ""
            txtCodLogrContrib = ""
            Call HabilitaCaixa(True)
            txtCpfCgc = ""
            txtCPFOcupante = ""
            txtOcupante = ""
            txtNomeContrib.SetFocus
        Case 2, 3
           ' TMPU701.Tag = Me.Name
    End Select
End Sub

Private Sub Form_Activate()
    txtic(0).SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim Controle As Control
    Dim i As Byte
    Dim rs As VSRecordset
    Set cadastro = New VSImposto
    
    Call Edita.AtualizaCombo(Bdados, cboTipoLogrContribBc, "Select ttl_nome From Tab_Tipo_Logr")
    
    
    Call AtualizaUF(cboUFBc)
    
    For Each Controle In Controls
        If IsNumeric(Controle.Tag) Then
            If Val(Controle.Tag) < 20 Then Call Edita.AtualizaCombo(Bdados, Controle, "Select convert(varchar,tco_cod_componente) " & Bdados.Concatena & "'-'" & Bdados.Concatena & " tco_descricao_componente From Tab_Componente_Avancado Where tco_grupo = " & Controle.Tag & " order by tco_cod_componente asc")
        End If
    Next
    HabilitaCaixa False
    cboLoteamento.Preencher Bdados, "Select TLO_COD_LOTEAMENTO,TLO_DESCRICAO from TAB_LOTEAMENTO ORDER BY TLO_DESCRICAO", 1
    cboEdificio.Preencher Bdados, "Select TED_COD_EDIFICIO,TED_DESCRICAO from TAB_EDIFICIO ORDER BY TED_DESCRICAO", 1
    txtNomeContrib.Enabled = True
    fra(9).Visible = Nvl(Temp.PegaParametro(Bdados, "AFORAMENTO TCIU101"), 1)
    txtic(2).MaxLength = Nvl(Temp.PegaParametro(Bdados, "CAMPO QUADRA"), 4)
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    NovoContrib = True
    Bdados.FechaTabela rs
    Boletim = tbo_Territorial
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        fra(9).Visible = False
    End If
End Sub

Private Sub lstCond_DblClick()
    If lstCond.SelectedItem Is Nothing Then Exit Sub
    Dim ItmX As Object
    Dim i As Byte
    
    On Error Resume Next
    txtic(9) = lstCond.SelectedItem
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
End Sub

Private Sub lstEdific_DblClick()
    Dim i As Integer
    If lstEdific.SelectedItem Is Nothing Then Exit Sub
    If Trim(txtInscImobiliaria) <> "" Then
        If Confirma("Existe uma unidade edificada em aberto. Deseja exclui-la?") Then
            lstEdific.ListItems.Remove lstEdific.SelectedItem.Index
        Else
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
            lstEdific.ListItems.Remove lstEdific.SelectedItem.Index
        End If
    End If
End Sub

Private Sub lstPesq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaGrid lstPesq, ColumnHeader
End Sub

Private Sub lstPesq_DblClick()
    txtIM = lstPesq.SelectedItem
    Call txtIm_LostFocus
End Sub

Private Sub tabCad_Click(PreviousTab As Integer)
    'NOVIDADE
    If tabCad.Tab = 2 Then
        If Trim(txtInscImobiliaria) = "" And Trim(txtic(4)) <> "" Then
            txtInscImobiliaria.Enabled = True
            txtic(10) = txtic(0)
            txtic(11) = txtic(1)
            txtic(12) = txtic(2)
            txtic(13) = txtic(3)
            txtInscImobiliaria.SetFocus
            DoEvents
        End If
    ElseIf tabCad.Tab = 3 Then
        If Trim(txtic(4)) <> "" Then
            txtic(5) = txtic(0)
            txtic(6) = txtic(1)
            txtic(7) = txtic(2)
            txtic(8) = txtic(3)
            txtic(9).Enabled = True
            
            cboTipoImovelBc.ListIndex = cboTipoImovel.ListIndex
            txtCodLogrBc = txtCodLogr
            txtLogr = txtTipoLogrBt
            txtNomeLogr = txtLogrBt
            txtNumeroBc = txtNumero
            txtLoteamentoBc = cboLoteamento.Coluna(0).VALOR
            txtQuadraBc = txtQuadra
            txtLoteBc = txtLote
            txtBairro = txtBairroBt
            txtCepImBc = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
            DoEvents
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

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

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
    txtInscImobiliaria.Enabled = True
    txtInscImobiliaria.SetFocus
    'DoEvents
End Sub

Private Sub txtBairroContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCep_LostFocus()
    If Trim$(txtCEP) = "" Then
        txtCEP = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
    Else
        If IsNumeric(txtCEP) Then txtCepImBc = Edita.FormataTexto(txtCEP, CEP)
    End If
End Sub

Private Sub txtCepImBc_LostFocus()
    If IsNumeric(txtCepImBc) Then txtCepImBc = Edita.FormataTexto(txtCepImBc, CEP)
End Sub

Private Sub txtCodBairro_Change()
    txtBairroBt = ""
End Sub

Private Sub txtCodBairro_LostFocus()
    On Error GoTo TrataErro
    Dim rs As VSRecordset
    Dim Sql As String
    If Trim(txtCodBairro) <> "" Then
        Sql = " select TBA_NOME from TAB_BAIRRO where tba_cod_bairro=" & txtCodBairro & " and tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
        If Bdados.AbreTabela(Sql, rs) Then
            txtBairroBt = rs(0)
        Else
            Avisa "Bairro inexistente."
            txtCodBairro.SetFocus
            Exit Sub
        End If
    Else
        txtBairroBt = ""
    End If
    Bdados.FechaTabela rs
    
    Exit Sub
TrataErro:
    Util.Erro Err.Description
End Sub

Private Sub txtCodComponente_Change(Index As Integer)
    Dim Controle As Control
    On Error GoTo trata
    If Index = 20 Then
        'cboCobrancaBc.ListIndex = Nvl(txtCodComponente(Index).Text, 0) - 1
        cboCobrancaBc.ListIndex = ListIndexDe(cboCobrancaBc, txtCodComponente(Index).Text & "-" & DescricaoComponente(cboCobrancaBc.Tag, txtCodComponente(Index).Text))
        Exit Sub
    End If
    For Each Controle In Controls
        If Controle.Tag = Index + 1 Then
            'Controle.ListIndex = Util.Nvl(txtCodComponente(Index).Text, 0) - 1
            Controle.ListIndex = ListIndexDe(Controle, txtCodComponente(Index).Text & "-" & DescricaoComponente(Controle.Tag, txtCodComponente(Index).Text))
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

Private Sub txtCodLogr_Change()
    txtTipoLogrBt = ""
    txtLogrBt = ""
End Sub

Private Sub txtCodLogr_LostFocus()
    On Error GoTo TrataErro
    Dim Query As String
    Dim rs As VSRecordset
    If Trim(txtCodLogr) = "" Then Exit Sub
    Query = "SELECT TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome, " & _
        " TAB_BAIRRO.TBA_NOME FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
        " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
         " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & txtCodLogr & "' and tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
    If Bdados.AbreTabela(Query, rs) Then
        'Rs.MoveFirst
        txtTipoLogrBt = rs(0)
        txtLogrBt = rs(1)
        
    Else
        Avisa "C?digo de logradouro inv?lido."
        txtCodLogr.SetFocus
    End If
    Bdados.FechaTabela rs
    
    Exit Sub
TrataErro:
    If Err.Number = 3265 Then
        Resume Next
    Else
        Util.Erro Err.Description
    End If
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
            " TAB_BAIRRO.TBA_NOME,tlg_cep FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
            " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
             " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & txtCodLogrContrib & "' and tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
        If Bdados.AbreTabela(Query, rs) Then
            txtNomeTipoLogrContrib = rs(0)
            txtNomeLogrContrib = rs(1)
            txtBairroContrib = rs(2)
            txtMunic = Aplicacoes.Municipio
            txtCEP = rs.Fields("tlg_cep")
        Else
            Avisa "C?digo de logradouro inv?lido."
            txtCodLogr.SetFocus
        End If
        Bdados.FechaTabela rs
        txtCEP = Temp.PegaParametro(Bdados, "CEP CLIENTE") '' & "-" &  Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
        txtUF = Temp.PegaParametro(Bdados, "ESTADO CLIENTE")
        
    End If
End Sub

Private Sub txtCodMens_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
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

    If txtCpfCgc = "99999999999" Or txtCpfCgc = "999.999.999-99" Or txtCpfCgc = "00000000000" Or txtCpfCgc = "000.000.000-00" Then
        Util.Avisa "CPF inv?lido."
        txtCpfCgc.SetFocus
    End If
    If Len(txtCpfCgc) = 11 Then
        If Not Util.ValidaCpf(txtCpfCgc) Then
            Call Util.Informa("N?mero de CPF inv?lido.")
            txtCpfCgc.SetFocus
            Exit Sub
        End If
        txtCpfCgc = Edita.FormataTexto(txtCpfCgc, Cpf)
        'txtCpfOcupante = txtCpfCgc
    ElseIf Len(txtCpfCgc) = 14 And Mid(txtCpfCgc, 4, 1) <> "." Then
        txtCpfCgc.MaxLength = 20
        txtCpfCgc = Edita.FormataTexto(txtCpfCgc, Cgc)
        'txtCpfOcupante = txtCpfCgc
    ElseIf Trim(txtCpfCgc) <> "" And Len(txtCpfCgc) <> 18 And Mid(txtCpfCgc, 4, 1) <> "." Then
        Call Util.Informa("N?mero de CNPJ ou CPF inv?lido.")
        txtCpfCgc.SetFocus
    End If
End Sub

Private Sub txtCpfCgcBc_LostFocus()
    If Len(txtCpfCgcBc) = 11 Then
        If Not Util.ValidaCpf(txtCpfCgcBc) Then
            Call Util.Informa("N?mero de CPF inv?lido.")
             txtCpfCgcBc.SetFocus
            Exit Sub
        End If
        txtCpfCgcBc = Edita.FormataTexto(txtCpfCgcBc, Cpf)
    ElseIf Len(txtCpfCgcBc) = 14 And Mid(txtCpfCgcBc, 4, 1) <> "." Then
        txtCpfCgcBc.MaxLength = 20
        txtCpfCgcBc = Edita.FormataTexto(txtCpfCgcBc, Cgc)
    ElseIf Trim(txtCpfCgcBc) <> "" And Len(txtCpfCgcBc) <> 18 And Mid(txtCpfCgcBc, 4, 1) <> "." Then
        Call Util.Informa("N?mero de CNPJ ou CPF inv?lido.")
        txtCpfCgcBc.SetFocus
    End If
End Sub

Private Sub txtCpfOcupante_LostFocus()
    If Len(txtCPFOcupante) = 11 Then
        If Not Util.ValidaCpf(txtCPFOcupante) Then
            Call Util.Informa("N?mero de CPF inv?lido.")
            txtCPFOcupante.SetFocus
            Exit Sub
        End If
        txtCPFOcupante = Edita.FormataTexto(txtCPFOcupante, Cpf)
    End If
    tabCad.Tab = 1
    DoEvents
    txtCodComponente(0).SetFocus
End Sub

Private Sub txtCpfOcupanteBc_LostFocus()
    If Len(txtCpfOcupanteBc) = 11 Then
        If Not Util.ValidaCpf(txtCpfOcupanteBc) Then
            Call Util.Informa("N?mero de CPF inv?lido.")
            txtCpfOcupanteBc.SetFocus
            Exit Sub
        End If
        txtCpfOcupanteBc = Edita.FormataTexto(txtCpfOcupanteBc, Cpf)
    End If
End Sub

Private Sub txtDataAforamento_LostFocus()
    txtDataAforamento = Edita.FormataTexto(txtDataAforamento, Data)
End Sub

Private Sub txtDtRegistro_Validate(Cancel As Boolean)
    txtDtRegistro = Edita.FormataTexto(txtDtRegistro, Data)
End Sub

Private Sub txtIc_Change(Index As Integer)
    If Len(txtic(Index)) = txtic(Index).MaxLength Then
       SendKeys "{ENTER}"
    End If
End Sub

Private Sub txtIc_Validate(Index As Integer, Cancel As Boolean)
    If Index = 4 Then
        If Trim(txtic(4)) = "" Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        If ImovelJaCadastrado Then
            Util.Avisa "Este im?vel j? est? cadastrado"
            Cancel = True
        End If
    End If
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
        If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
            txtIM = cadastro.FormataInscricao(txtIM, InscContrib)
        End If
        Sql = "Select  tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, " & _
        " tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF,TCI_COD_LOGRADOURO from Tab_Contribuinte where tci_im = '" & txtIM & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtNomeContrib = "" & rs(0) 'Rs!tci_Nome
            'txtOcupante = txtNomeContrib
            txtNomeTipoLogrContrib = CStr("" & rs(1))
            txtNomeLogrContrib = "" & rs(2) '!tci_nome_logradouro
            txtNumeroContrib = "" & rs(3) '!tci_numero
            txtCompContrib = "" & rs(4) '!tci_complemento
            txtBairroContrib = "" & rs(5) '!tci_bairro
            txtCEP = "" & Temp.PegaParametro(Bdados, "CEP MUNICIPIO")
            txtMunic = rs(7)
            txtUF = rs(8) '!tci_UF
            txtCpfCgc = "" & rs(9)
            txtCodLogrContrib = "" & rs("TCI_COD_LOGRADOURO")
            'txtCpfOcupante = txtCpfCgc
        Else
            Call Util.Informa("Contribuinte n?o cadastrado.")
            txtIM.Enabled = True
            txtIM.SetFocus
        End If
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtIMBc_LostFocus()
    'NOVIDADE
    Dim rs As VSRecordset
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIMBc) <> "" Then
        txtIMBc = cadastro.FormataInscricao(txtIMBc, InscContrib)
        Sql = "Select tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF from Tab_Contribuinte where tci_im = '" & txtIMBc & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtNomeContribBc = rs(0)  'Rs!tci_Nome
            cboTipoLogrContribBc.ListIndex = cadastro.BuscaCodLogr(rs(1)) - 1
            txtNomeLogrContribBc = rs(2)  '!tci_nome_logradouro
            txtNumeroContribBc = rs(3)  '!tci_numero
            txtCompContribBc = rs(4)  '!tci_complemento
            txtBairroContribBc = rs(5)  '!tci_bairro
            txtCepBc = rs(6)  '!tci_cep
            txtMunicBc = rs(7)
            cboUFBc = rs(8)  '!tci_UF
            txtCpfCgcBc = "" & rs(9)
        Else
            Call Util.Informa("Contribuinte n?o cadastrado.")
            txtIMBc.Enabled = True
            txtIMBc.SetFocus
        End If
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtInscImobiliaria_LostFocus()
    txtInscImobiliaria = Format(txtInscImobiliaria, "000")
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    'KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtLoteamento_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
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

Private Sub txtNomeContrib_LostFocus()
    'txtOcupante = txtNomeContrib
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
    'KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtSecao_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.Maiuscula(KeyAscii)
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

