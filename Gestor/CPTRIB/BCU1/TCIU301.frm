VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU301 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU301"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fra 
      Height          =   1785
      Index           =   1
      Left            =   30
      TabIndex        =   34
      Top             =   2550
      Width           =   11115
      _ExtentX        =   19606
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
      Begin VB.TextBox txtCompContrib 
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
         Left            =   7290
         TabIndex        =   48
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtIM 
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
         Left            =   1470
         TabIndex        =   47
         Top             =   210
         Width           =   1305
      End
      Begin VB.TextBox txtNomeContrib 
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
         Left            =   3060
         TabIndex        =   46
         Top             =   210
         Width           =   4965
      End
      Begin VB.TextBox txtNomeLogrContrib 
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
         Left            =   3060
         TabIndex        =   45
         Top             =   570
         Width           =   2415
      End
      Begin VB.TextBox txtCep 
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
         Left            =   3060
         MaxLength       =   10
         TabIndex        =   44
         Top             =   960
         Width           =   1125
      End
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Command1"
         Default         =   -1  'True
         Height          =   255
         Left            =   7740
         TabIndex        =   43
         Top             =   3090
         Width           =   375
      End
      Begin VB.TextBox txtBairroContrib 
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
         Left            =   8970
         TabIndex        =   42
         Top             =   630
         Width           =   2055
      End
      Begin VB.ComboBox cboTipoLogrContrib 
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
         ItemData        =   "TCIU301.frx":0000
         Left            =   1470
         List            =   "TCIU301.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   570
         Width           =   1365
      End
      Begin VB.TextBox txtOcupante 
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
         Left            =   3060
         TabIndex        =   40
         Top             =   1350
         Width           =   4965
      End
      Begin VB.TextBox txtCpfOcupante 
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
         Left            =   8970
         MaxLength       =   20
         TabIndex        =   39
         Top             =   1350
         Width           =   2055
      End
      Begin VB.TextBox txtCpfCgc 
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
         Left            =   8970
         MaxLength       =   20
         TabIndex        =   38
         Top             =   210
         Width           =   2055
      End
      Begin VB.ComboBox cboUF 
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
         ItemData        =   "TCIU301.frx":002E
         Left            =   10245
         List            =   "TCIU301.frx":003B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   990
         Width           =   795
      End
      Begin VB.TextBox txtNumeroContrib 
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
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   36
         Top             =   600
         Width           =   525
      End
      Begin VB.TextBox txtMunic 
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
         Left            =   5880
         TabIndex        =   35
         Top             =   990
         Width           =   4335
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   8
         Left            =   150
         TabIndex        =   49
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
         Index           =   12
         Left            =   390
         TabIndex        =   50
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
         Index           =   14
         Left            =   2640
         TabIndex        =   51
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
         TabIndex        =   52
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
         Left            =   5580
         TabIndex        =   53
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
         Index           =   17
         Left            =   8370
         TabIndex        =   54
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
         Left            =   6510
         TabIndex        =   55
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
         Left            =   8085
         TabIndex        =   56
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
         TabIndex        =   57
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
         TabIndex        =   58
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
   End
   Begin VB.TextBox txtFatorFixo 
      Height          =   285
      Left            =   8640
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   3030
      Width           =   375
   End
   Begin Threed.SSFrame fra 
      Height          =   1875
      Index           =   0
      Left            =   30
      TabIndex        =   7
      Top             =   660
      Width           =   11115
      _ExtentX        =   19606
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
         Left            =   1470
         TabIndex        =   0
         Top             =   270
         Width           =   1905
      End
      Begin VB.TextBox txtLote 
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
         Left            =   3270
         MaxLength       =   5
         TabIndex        =   21
         Top             =   1410
         Width           =   495
      End
      Begin VB.TextBox txtQuadra 
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
         Left            =   2190
         MaxLength       =   5
         TabIndex        =   20
         Top             =   1380
         Width           =   495
      End
      Begin VB.TextBox txtLoteamento 
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
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   19
         Top             =   1380
         Width           =   465
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
         TabIndex        =   18
         Top             =   1050
         Width           =   1425
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
         TabIndex        =   17
         Top             =   1050
         Width           =   555
      End
      Begin VB.TextBox txtCepIm 
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
         Left            =   8610
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1410
         Width           =   915
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
         ItemData        =   "TCIU301.frx":005C
         Left            =   9600
         List            =   "TCIU301.frx":0066
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   600
         Width           =   1455
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
         Left            =   6240
         TabIndex        =   14
         Top             =   630
         Width           =   1575
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
         Left            =   3270
         MaxLength       =   3
         TabIndex        =   5
         Tag             =   "Unidade"
         Top             =   630
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
         Left            =   2700
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "Lote"
         Top             =   630
         Width           =   525
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
         Left            =   2190
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "Quadra"
         Top             =   630
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
         Left            =   1830
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "Setor"
         Top             =   630
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
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "Distrito"
         Top             =   630
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
         Left            =   1200
         TabIndex        =   13
         Top             =   1020
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
         Left            =   5190
         TabIndex        =   12
         Top             =   1380
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
         Left            =   4050
         TabIndex        =   11
         Top             =   1020
         Width           =   3255
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
         Left            =   2970
         MaxLength       =   11
         TabIndex        =   10
         Top             =   1020
         Width           =   1035
      End
      Begin VB.TextBox txtCodBairro 
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
         Height          =   330
         Left            =   4530
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox txtCodMens 
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
         Left            =   10680
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1410
         Width           =   315
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   8910
         TabIndex        =   22
         Top             =   1110
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
         TabIndex        =   23
         Top             =   1110
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
         Left            =   3930
         TabIndex        =   24
         Top             =   1440
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
         TabIndex        =   25
         Top             =   1410
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
         Left            =   1800
         TabIndex        =   26
         Top             =   1440
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
         Left            =   2790
         TabIndex        =   27
         Top             =   1440
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
         Height          =   225
         Index           =   49
         Left            =   8190
         TabIndex        =   28
         Top             =   1440
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
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   7
         Left            =   9150
         TabIndex        =   29
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
         Height          =   180
         Index           =   72
         Left            =   4980
         TabIndex        =   30
         Top             =   660
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
         Index           =   75
         Left            =   120
         TabIndex        =   31
         Top             =   660
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
         Caption         =   "Insc. Imobiliária:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   82
         Left            =   90
         TabIndex        =   32
         Top             =   1080
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
         Left            =   9600
         TabIndex        =   33
         Top             =   1470
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
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   62
         Top             =   315
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
         Caption         =   "Cad Imobiliário:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   2
      Left            =   10020
      TabIndex        =   59
      Top             =   4380
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   1
      Left            =   8850
      TabIndex        =   60
      Top             =   4380
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Excluir"
      Acao            =   2
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   61
      Top             =   4380
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Limpar"
      Acao            =   6
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   1138
      Icone           =   "TCIU301.frx":0080
   End
End
Attribute VB_Name = "TCIU301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim cadastro As VSImposto
Dim NovoContrib As Boolean
Dim Sql As String
Private Boletim As TipoBoletim
Dim Consultando As Boolean

Function TotalProva(Valor As String) As String
    Static Total As Double
    If Trim(Valor) = "" Then Valor = "0"
    Total = CDbl(Valor) + Total
    TotalProva = Total
End Function

Public Sub HabilitaCaixa(Status As Boolean)
    txtIM.Enabled = Not Status
    txtNomeContrib.Enabled = Status
    cboTipoLogrContrib.Enabled = Status
    txtNomeLogrContrib.Enabled = Status
    txtNumeroContrib.Enabled = Status
    txtCompContrib.Enabled = Status
    txtBairroContrib.Enabled = Status
    txtCep.Enabled = Status
    txtMunic.Enabled = Status
    cboUF.Enabled = Status
    txtIM = ""
    txtNomeContrib = ""
    cboTipoLogrContrib.ListIndex = -1
    txtNomeLogrContrib = ""
    txtNumeroContrib = ""
    txtCompContrib = ""
    txtBairroContrib = ""
    txtCep = ""
    txtMunic = ""
    cboUF.ListIndex = -1
End Sub

Public Sub AtualizaUF(Combo As ComboBox)
   Combo.Clear
   Combo.AddItem "MA"
   Combo.AddItem "AC"
   Combo.AddItem "AM"
   Combo.AddItem "AP"
   Combo.AddItem "AL"
   Combo.AddItem "BA"
   Combo.AddItem "CE"
   Combo.AddItem "DF"
   Combo.AddItem "ES"
   Combo.AddItem "GO"
   Combo.AddItem "MG"
   Combo.AddItem "MS"
   Combo.AddItem "MT"
   Combo.AddItem "PA"
   Combo.AddItem "PB"
   Combo.AddItem "PE"
   Combo.AddItem "PI"
   Combo.AddItem "PR"
   Combo.AddItem "SC"
   Combo.AddItem "SE"
   Combo.AddItem "SP"
   Combo.AddItem "RJ"
   Combo.AddItem "RN"
   Combo.AddItem "RO"
   Combo.AddItem "RR"
   Combo.AddItem "RS"
   Combo.AddItem "TO"
End Sub

Private Sub cmd_Click(Index As Integer)
On Error Resume Next
    Dim Valores As String
    Dim Campos As String
    Dim DataReab As Date
    Dim RsAux As VSRecordset
    Dim Rs As VSRecordset
    Dim InscricaoMunicipal As String
    Dim InscricaoCadastral As String
    Dim CodLogr As Long
    Dim DtVenc As String
    Dim SitCadastral As String
    Static Unidades As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Motivo  As String
    Dim Lote As New BCI
    Select Case cmd(Index).Caption
        Case "&Excluir"
          Do
            Motivo = Edita.TiraPic(Edita.TiraPic(Trim(Util.Entrada("Informe o motivo da exclusão.", "Justificativa.")), Chr(13)), Chr(13))
          Loop While Trim(Motivo) = ""
'                For i = 0 To 4
'                    If Trim(txtic(i)) = "" Then
'                        Avisa "Informe a Inscrição do imóvel."
'                        txtic(i).SetFocus
'                        Exit Sub
'                    End If
'                Next
                If Confirma("Deseja excluir realmente o imovel.") Then
                    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                        txtFatorFixo.Tag = "1000"
                        Dim Inscricoes As String
                        Screen.MousePointer = 11
                        CodLogr = txtCodLogr
                        InscricaoCadastral = txtCodReduzido
                        Lote.GravaHistorico InscricaoCadastral, "REGISTRO EXCLUÍDO  -  " & Motivo
                        InscricaoMunicipal = txtIM
                        Sql = "Select tim_ic from tab_imovel where tim_ic ='" & InscricaoCadastral & "' or tim_ic_condominio ='" & InscricaoCadastral & "'"
                        If Bdados.AbreTabela(Sql) Then
                            Inscricoes = ""
                            Bdados.Tabela.MoveFirst
                            Do
                                Inscricoes = Inscricoes & "'" & Bdados.Tabela(0) & "',"
                                Bdados.Tabela.MoveNext
                            Loop While Not Bdados.Tabela.EOF
                            Inscricoes = Left(Inscricoes, Len(Inscricoes) - 1)
                        Else
                            Avisa "Registro não encontrado."
                            Screen.MousePointer = 0
                            Exit Sub
                        End If
                        Bdados.DeletaDados "TAB_DETALHE_IMOVEL", "TDI_TIM_IC in (" & Inscricoes & ")"
                        Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "TGT_INSCRICAO in (" & Inscricoes & ")"
                        Bdados.DeletaDados "TAB_OBRIGACAO_CONTRIBUINTE", "TOC_INSCRICAO in (" & Inscricoes & ")"
                        Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "TCC_INSCRICAO in (" & Inscricoes & ")"
                        Bdados.DeletaDados "TAB_IMOVEL", "TIM_IC in (" & Inscricoes & ")"
                        Informa "Registro eliminado com sucesso."
                        Edita.LimpaCampos Me
                        For i = 0 To 4
                            txtic(i).Enabled = True
                        Next
                        txtCodReduzido.SetFocus
                        Screen.MousePointer = 0
                    Else
                        txtFatorFixo.Tag = "1000"
                        Screen.MousePointer = 11
                        CodLogr = txtCodLogr
                        InscricaoCadastral = txtCodReduzido
                        InscricaoMunicipal = txtIM
                        Bdados.DeletaDados "TAB_IMOVEL", "TIM_IC ='" & InscricaoCadastral & "'" '& IIf(CInt(txtic(4)) >= 1, " and tim_unidade = " & Bdados.Converte(txtic(4), TCInteiro), "")
                        Bdados.DeletaDados "TAB_DETALHE_IMOVEL", "TDI_TIM_IC ='" & InscricaoCadastral & "'" '& IIf(CInt(txtic(4)) >= 1, " and tdi_tim_ic_unidade = " & Bdados.Converte(txtic(4), TCInteiro), "")
                        Informa "Registro eliminado com sucesso."
                        Edita.LimpaCampos Me
                        txtic(0).SetFocus
                        For i = 0 To 4
                            txtic(i).Enabled = True
                        Next
                        txtic(0).SetFocus
                        Screen.MousePointer = 0
                        
                    End If
                End If
        Case "&Limpar"
            Call Edita.LimpaCampos(Me)
            For i = 0 To 4
                txtic(i).Enabled = True
            Next
            txtic(0).SetFocus
        Case "Sai&r"
            Unload Me
    End Select
End Sub

Private Sub cmdEnter_Click()
        SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    
    Dim Controle As Control
    Dim i As Byte
    Dim Rs As VSRecordset
    Set cadastro = New VSImposto
    Call Edita.AtualizaCombo(Bdados, cboTipoLogrContrib, "Select ttl_nome From Tab_Tipo_Logr")
    Call AtualizaUF(cboUF)
    
    For Each Controle In Controls
        If IsNumeric(Controle.Tag) Then
            If Val(Controle.Tag) < 20 Then Call Edita.AtualizaCombo(Bdados, Controle, "Select tco_descricao_componente From Tab_Componente_Avancado Where tco_grupo = " & Controle.Tag & " order by tco_cod_componente asc")
        End If
    Next
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    NovoContrib = True
    Bdados.FechaTabela Rs
    Boletim = tbo_Territorial
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

Private Sub txtBairroContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCep_LostFocus()
    If IsNumeric(txtCep) Then txtCepIm = Edita.FormataTexto(txtCep, CEP)
End Sub

Private Sub txtCepIm_LostFocus()
    If IsNumeric(txtCepIm) Then txtCepIm = Edita.FormataTexto(txtCepIm, CEP)
End Sub

Private Sub txtCodBairro_LostFocus()
    Dim Rs As VSRecordset
    Dim Sql As String
    On Error Resume Next
    If Trim(txtCodBairro) <> "" Then
        Sql = " select TBA_NOME from TAB_BAIRRO where tba_cod_bairro=" & txtCodBairro
        If Bdados.AbreTabela(Sql, Rs) Then
            txtBairroBt = Rs(0)
        Else
            Avisa "Bairro inexistente."
            txtCodBairro.SetFocus
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtCodComponente_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodLogr_LostFocus()
    Dim Query As String
    Dim Rs As VSRecordset
    If Trim(txtCodLogr) = "" Then Exit Sub
    Query = "SELECT TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome, " & _
        " TAB_BAIRRO.TBA_NOME FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
        " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
         " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & txtCodLogr & "'"
    If Bdados.AbreTabela(Query, Rs) Then
        txtTipoLogrBt = Rs(0)
        txtLogrBt = Rs(1)
    Else
        Avisa "Código de logradouro inválido."
        txtCodLogr.Enabled = True
        txtCodLogr = ""
        
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtCodLogrBc_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select "
End Sub

Private Sub txtCodMens_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtCodReduzido_Validate(Cancel As Boolean)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim i As Byte
    
    If Trim(txtCodReduzido) = "" Then Exit Sub
    
'    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        Sql = "Select * from tab_imovel where tIM_ic ='" & txtCodReduzido & "'"
'    Else
'        Util.Avisa "Função não implementada."
'        txtCodReduzido.SetFocus
'        Exit Sub
'    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        txtICAnterior = "" & IIf(Rs!tim_ic_anterior = 0, "", Rs!tim_ic_anterior)
        cboTipoImovel.ListIndex = Rs!tim_tipo_imovel - 1
        txtCodLogr = "" & Rs!tim_tlg_cod_logradouro
        txtCodLogr_LostFocus
        txtNumero = "" & Rs!tim_numero
        txtComplemento = "" & Rs!tim_complemento
        txtLoteamento = "" & Rs!tim_loteamento
        txtQuadra = "" & Rs!tim_QUADRA
        txtLote = "" & Rs!tim_lote
        txtCepIm = "" & Rs!tim_cep
        txtOcupante = "" & Rs!tim_ocupante
        txtCPFOcupante = "" & Rs!tim_cgc_cpf_ocupante
        txtCodBairro = "" & Rs!tim_TBA_COD_BAIRRO
        txtCodBairro_LostFocus
        txtICAnterior = "" & Rs!tim_ic_anterior
        txtCodMens = "" & Rs!tim_COD_MENSAGEM
        txtic(0) = "" & Left(Trim(Rs.Fields("tim_ic_auxiliar")), 2)
        txtic(1) = "" & Mid(Trim(Rs.Fields("tim_ic_auxiliar")), 3, 2)
        txtic(2) = "" & Mid(Trim(Rs.Fields("tim_ic_auxiliar")), 5, 3)
        txtic(3) = "" & Mid(Trim(Rs.Fields("tim_ic_auxiliar")), 8, 4)
        txtic(4) = "" & Right(Trim(Rs.Fields("tim_ic_auxiliar")), 3)
        'VOU PEGAR O CONTRIBUINTE
        txtIM = Rs!tim_tci_im
        txtIm_LostFocus
        For i = 0 To 4
            txtic(i).Enabled = False
        Next
        Bdados.FechaTabela Rs
    Else
        Informa "Imóvel não cadastrado."
        txtic(4).SetFocus
    End If
    Bdados.FechaTabela Rs
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
    If Len(txtCpfCgc) = 11 Then
        If Not Util.ValidaCpf(txtCpfCgc) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCpfCgc.SetFocus
            Exit Sub
        End If
        txtCpfCgc = Edita.FormataTexto(txtCpfCgc, Cpf)
    ElseIf Len(txtCpfCgc) = 14 And Mid(txtCpfCgc, 4, 1) <> "." Then
        txtCpfCgc.MaxLength = 20
        txtCpfCgc = Edita.FormataTexto(txtCpfCgc, Cgc)
    ElseIf Trim(txtCpfCgc) <> "" And Len(txtCpfCgc) <> 18 And Mid(txtCpfCgc, 4, 1) <> "." Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
        txtCpfCgc.SetFocus
    End If
End Sub

Private Sub txtIc_Change(Index As Integer)
    If Len(txtic(Index)) = txtic(Index).MaxLength Then
       SendKeys "{ENTER}"
    End If
End Sub

Private Sub txtic_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_LostFocus(Index As Integer)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim i As Byte
    
    If Trim(txtic(4)) = "" Then Exit Sub
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        Sql = "Select * from tab_imovel where tIM_ic ='" & txtCodReduzido & "'"
    Else
        Sql = "Select * from tab_imovel where tIM_ic ='" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & txtic(4) & "'" & IIf(CInt(txtic(4)) >= 1, " and tim_unidade = " & Bdados.Converte(txtic(4), TCInteiro), "")
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        txtICAnterior = "" & IIf(Rs!tim_ic_anterior = 0, "", Rs!tim_ic_anterior)
        cboTipoImovel.ListIndex = Rs!tim_tipo_imovel - 1
        txtCodLogr = "" & Rs!tim_tlg_cod_logradouro
        txtCodLogr_LostFocus
        txtNumero = "" & Rs!tim_numero
        txtComplemento = "" & Rs!tim_complemento
        txtLoteamento = "" & Rs!tim_loteamento
        txtQuadra = "" & Rs!tim_QUADRA
        txtLote = "" & Rs!tim_lote
        txtCepIm = "" & Rs!tim_cep
        txtOcupante = "" & Rs!tim_ocupante
        txtCPFOcupante = "" & Rs!tim_cgc_cpf_ocupante
        txtCodBairro = "" & Rs!tim_TBA_COD_BAIRRO
        txtCodBairro_LostFocus
        txtICAnterior = "" & Rs!tim_ic_anterior
        txtCodMens = "" & Rs!tim_COD_MENSAGEM
        'VOU PEGAR O CONTRIBUINTE
        txtIM = Rs!tim_tci_im
        txtIm_LostFocus
        For i = 0 To 4
            txtic(i).Enabled = False
        Next
        Bdados.FechaTabela Rs
    Else
        Informa "Imóvel não cadastrado."
        txtic(4).SetFocus
    End If
    Bdados.FechaTabela Rs
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
    On Error Resume Next
    Dim Rs As VSRecordset
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIM) <> "" Then
        If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
            txtIM = cadastro.FormataInscricao(txtIM, InscContrib)
        End If
        Sql = "Select tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, " & _
        " tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF from Tab_Contribuinte where tci_im = '" & txtIM & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            txtNomeContrib = "" & Rs(0) 'Rs!tci_Nome
            cboTipoLogrContrib.ListIndex = "" & cadastro.BuscaCodLogr(Nvl("" & Rs(1), 0)) - 1
            txtNomeLogrContrib = "" & Rs(2) '!tci_nome_logradouro
            txtNumeroContrib = "" & Rs(3) '!tci_numero
            txtCompContrib = "" & Rs(4) '!tci_complemento
            txtBairroContrib = "" & Rs(5) '!tci_bairro
            txtCep = Rs(6) '!tci_cep
            txtMunic = "" & Rs(7)
            If Trim("" & Rs(8)) <> "" Then cboUF = "" & Rs(8) '!tci_UF
            txtCpfCgc = "" & Rs(9)
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtIM.Enabled = True
            txtIM.SetFocus
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtLoteamento_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtMunic_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMunic_LostFocus()
    If Trim(txtMunic) = "" Then txtMunic = Aplicacoes.municipio
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


