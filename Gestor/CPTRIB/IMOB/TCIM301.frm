VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.0#0"; "VTControles.ocx"
Begin VB.Form TCIM301 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fra 
      Height          =   1785
      Index           =   1
      Left            =   30
      TabIndex        =   33
      Top             =   2100
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
         TabIndex        =   47
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
         MaxLength       =   11
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   960
         Width           =   1125
      End
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Command1"
         Default         =   -1  'True
         Height          =   255
         Left            =   7740
         TabIndex        =   42
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
         TabIndex        =   41
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
         ItemData        =   "TCIM301.frx":0000
         Left            =   1470
         List            =   "TCIM301.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         ItemData        =   "TCIM301.frx":002E
         Left            =   10245
         List            =   "TCIM301.frx":003B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   990
         Width           =   4335
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   8
         Left            =   150
         TabIndex        =   48
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
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
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
         TabIndex        =   54
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
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   57
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
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   3000
      Width           =   375
   End
   Begin Threed.SSFrame fra 
      Height          =   1425
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   660
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   2514
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
         TabIndex        =   20
         Top             =   1020
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
         TabIndex        =   19
         Top             =   990
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
         TabIndex        =   18
         Top             =   990
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
         TabIndex        =   17
         Top             =   660
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
         TabIndex        =   16
         Top             =   660
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
         TabIndex        =   15
         Top             =   1020
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
         ItemData        =   "TCIM301.frx":005C
         Left            =   9600
         List            =   "TCIM301.frx":0066
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   210
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
         TabIndex        =   13
         Top             =   240
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
         Left            =   3000
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
         Left            =   2430
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "Lote"
         Top             =   240
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
         Left            =   1920
         MaxLength       =   3
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
         Left            =   1560
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
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "Distrito"
         Top             =   240
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
         TabIndex        =   12
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
         Left            =   5190
         TabIndex        =   11
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
         Left            =   4050
         TabIndex        =   10
         Top             =   630
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
         TabIndex        =   9
         Top             =   630
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
         TabIndex        =   8
         Top             =   990
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
         TabIndex        =   7
         Top             =   1020
         Width           =   315
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   8910
         TabIndex        =   21
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
         TabIndex        =   22
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
         Left            =   3930
         TabIndex        =   23
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
         TabIndex        =   24
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
         Left            =   1800
         TabIndex        =   25
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
         Left            =   2790
         TabIndex        =   26
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
         Height          =   225
         Index           =   49
         Left            =   8190
         TabIndex        =   27
         Top             =   1050
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
         TabIndex        =   28
         Top             =   270
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
         TabIndex        =   29
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
         Height          =   225
         Index           =   75
         Left            =   120
         TabIndex        =   30
         Top             =   270
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
         Caption         =   "Inscriçao:"
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
         TabIndex        =   31
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
         Left            =   9600
         TabIndex        =   32
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
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   1138
      Icone           =   "TCIM301.frx":0080
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   2
      Left            =   10020
      TabIndex        =   59
      Top             =   3930
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
      Left            =   8850
      TabIndex        =   60
      Top             =   3930
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Excluir"
      Acao            =   2
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   61
      Top             =   3930
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Limpar"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCIM301"
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
    Dim rs As VSRecordset
    Dim InscricaoMunicipal As String
    Dim InscricaoCadastral As String
    Dim CodLogr As Long
    Dim DtVenc As String
    Dim SitCadastral As String
    Static Unidades As Integer
    Dim i As Integer
    Dim j As Integer
    Select Case cmd(Index).Caption
        Case "&Excluir"
                For i = 0 To 4
                    If Trim(txtIc(i)) = "" Then
                        Avisa "Informe a Inscrição do imóvel."
                        txtIc(i).SetFocus
                        Exit Sub
                    End If
                Next
                If Confirma("Deseja excluir realmente o imovel.") Then
                    txtFatorFixo.Tag = "1000"
                    Screen.MousePointer = 11
                    CodLogr = txtCodLogr
                    InscricaoCadastral = txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & txtIc(4)
                    InscricaoMunicipal = txtIM
                    Bdados.DeletaDados "TAB_IMOVEL", "TIM_IC ='" & InscricaoCadastral & "'" & IIf(CInt(txtIc(4)) >= 1, " and tim_unidade = " & Bdados.Converte(txtIc(4), TCInteiro), "")
                    Bdados.DeletaDados "TAB_DETALHE_IMOVEL", "TDI_TIM_IC ='" & InscricaoCadastral & "'" & IIf(CInt(txtIc(4)) >= 1, " and tdi_tim_ic_unidade = " & Bdados.Converte(txtIc(4), TCInteiro), "")
                    Informa "Registro eliminado com sucesso."
                    Dim Capa As New cCapa
                    Capa.FechaLote txtIc(0), txtIc(1), txtIc(2)
                    Set Capa = Nothing

                    Edita.LimpaCampos Me
                    txtIc(0).SetFocus
                    For i = 0 To 4
                        txtIc(i).Enabled = True
                    Next
                    txtIc(0).SetFocus
                    Screen.MousePointer = 0
                End If
        Case "&Limpar"
            Call Edita.LimpaCampos(Me)
            For i = 0 To 4
                txtIc(i).Enabled = True
            Next
            txtIc(0).SetFocus
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
    Dim rs As VSRecordset
    Set Cadastro = New VSImposto
    Call Edita.AtualizaCombo(Bdados, cboTipoLogrContrib, "Select ttl_nome From Tab_Tipo_Logr")
    Call AtualizaUF(cboUF)
    
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
    Dim rs As VSRecordset
    Dim Sql As String
    On Error Resume Next
    If Trim(txtCodBairro) <> "" Then
        Sql = " select TBA_NOME from TAB_BAIRRO where tba_cod_bairro=" & txtCodBairro & " and tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
        If Bdados.AbreTabela(Sql, rs) Then
            txtBairroBt = rs(0)
        Else
            Avisa "Bairro inexistente."
            txtCodBairro.SetFocus
        End If
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtCodComponente_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodLogr_LostFocus()
    Dim Lote As New BCI
    Lote.BuscaLogradouro txtCodLogr, txtTipoLogrBt, txtLogrBt
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
    If Len(txtIc(Index)) = txtIc(Index).MaxLength Then
       SendKeys "{ENTER}"
    End If
End Sub

Private Sub txtic_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_LostFocus(Index As Integer)
    Dim Sql As String
    Dim rs As VSRecordset
    Dim i As Byte
    
    If Trim(txtIc(4)) = "" Then Exit Sub
    
    Sql = "Select * from tab_imovel where tIM_ic ='" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & txtIc(4) & "'" & IIf(CInt(txtIc(4)) >= 1, " and tim_unidade = " & Bdados.Converte(txtIc(4), TCInteiro), "")
    If Bdados.AbreTabela(Sql, rs) Then
        txtIcAnterior = "" & IIf(rs!tim_ic_anterior = 0, "", rs!tim_ic_anterior)
        cboTipoImovel.ListIndex = rs!tim_tipo_imovel - 1
        txtCodLogr = "" & rs!TIM_tlg_cod_logradouro
        txtCodLogr_LostFocus
        txtNumero = "" & rs!tim_numero
        txtComplemento = "" & rs!tim_complemento
        txtLoteamento = "" & rs!Tim_loteamento
        txtQuadra = "" & rs!tim_quadra
        txtLote = "" & rs!tim_Lote
        txtCepIm = "" & rs!tim_cep
        txtOcupante = "" & rs!tim_ocupante
        txtCpfOcupante = "" & rs!tim_cgc_cpf_ocupante
        txtCodBairro = "" & rs!TIM_TBA_COD_BAIRRO
        txtCodBairro_LostFocus
        txtIcAnterior = "" & rs!tim_ic_anterior
        txtCodMens = "" & rs!TIM_COD_MENSAGEM
        'VOU PEGAR O CONTRIBUINTE
        txtIM = rs!tim_tci_im
        txtIm_LostFocus
        For i = 0 To 4
            txtIc(i).Enabled = False
        Next
        Bdados.FechaTabela rs
    Else
        Informa "Imóvel não cadastrado."
        txtIc(4).SetFocus
    End If
    Bdados.FechaTabela rs
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
            cboTipoLogrContrib.ListIndex = Cadastro.BuscaCodLogr(rs(1)) - 1
            txtNomeLogrContrib = rs(2) '!tci_nome_logradouro
            txtNumeroContrib = rs(3) '!tci_numero
            txtCompContrib = rs(4) '!tci_complemento
            txtBairroContrib = rs(5) '!tci_bairro
            txtCep = rs(6) '!tci_cep
            txtMunic = rs(7)
            cboUF = rs(8) '!tci_UF
            txtCpfCgc = "" & rs(9)
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtIM.Enabled = True
            txtIM.SetFocus
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


