VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU102"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fra 
      Height          =   2325
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   4101
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
      Begin VB.TextBox txtxNovoCadastro 
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
         Left            =   5220
         TabIndex        =   1
         Top             =   270
         Width           =   1965
      End
      Begin VB.TextBox txtInscImob 
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
         Left            =   1620
         TabIndex        =   2
         Top             =   660
         Width           =   1965
      End
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
         Left            =   1620
         TabIndex        =   0
         Top             =   270
         Width           =   1965
      End
      Begin VB.ComboBox cboSit 
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
         ItemData        =   "TCIU102.frx":0000
         Left            =   1620
         List            =   "TCIU102.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "Tipo Imovel"
         Top             =   1800
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.TextBox txtZona 
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
         Left            =   8280
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Zona"
         Top             =   675
         Visible         =   0   'False
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
         Left            =   6150
         TabIndex        =   3
         Top             =   675
         Width           =   1575
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
         Left            =   10710
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Cod Mensagem"
         Top             =   1440
         Visible         =   0   'False
         Width           =   315
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
         Left            =   6030
         MaxLength       =   50
         TabIndex        =   20
         Tag             =   "Bairro"
         Top             =   1425
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
         Left            =   3150
         MaxLength       =   11
         TabIndex        =   19
         Top             =   1050
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
         Left            =   4230
         TabIndex        =   18
         Tag             =   "Nome Contribuinte"
         Top             =   1050
         Width           =   3495
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
         Left            =   6630
         TabIndex        =   17
         Tag             =   "Nome Contribuinte"
         Top             =   1440
         Width           =   2865
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
         Left            =   1620
         TabIndex        =   6
         Tag             =   "Logradouro"
         Top             =   1050
         Width           =   1485
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
         ItemData        =   "TCIU102.frx":0023
         Left            =   9600
         List            =   "TCIU102.frx":002D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo Imovel"
         Top             =   660
         Visible         =   0   'False
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
         TabIndex        =   16
         Top             =   1050
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
         TabIndex        =   15
         Top             =   1050
         Width           =   1425
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
         Left            =   1620
         MaxLength       =   5
         TabIndex        =   14
         Top             =   1440
         Width           =   405
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
         Left            =   2550
         MaxLength       =   5
         TabIndex        =   13
         Top             =   1440
         Width           =   555
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
         Left            =   3690
         MaxLength       =   5
         TabIndex        =   12
         Top             =   1440
         Width           =   615
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   8910
         TabIndex        =   22
         Top             =   1095
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
         Left            =   5430
         TabIndex        =   24
         Top             =   1455
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
         Left            =   510
         TabIndex        =   25
         Top             =   1500
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
         Left            =   2130
         TabIndex        =   26
         Top             =   1500
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
         Left            =   3240
         TabIndex        =   27
         Top             =   1500
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
         Left            =   9150
         TabIndex        =   28
         Top             =   690
         Visible         =   0   'False
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
         Left            =   690
         TabIndex        =   29
         Top             =   1095
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
         TabIndex        =   30
         Top             =   1485
         Visible         =   0   'False
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
         Left            =   4890
         TabIndex        =   31
         Top             =   735
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
         Left            =   180
         TabIndex        =   32
         Top             =   735
         Width           =   1395
         _ExtentX        =   2461
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
         Caption         =   "Insc. Imobiliária :"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   76
         Left            =   7800
         TabIndex        =   33
         Top             =   735
         Visible         =   0   'False
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
         Height          =   270
         Index           =   0
         Left            =   810
         TabIndex        =   34
         Top             =   1860
         Visible         =   0   'False
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
         Caption         =   "Situação:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   8
         Left            =   270
         TabIndex        =   35
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
         Caption         =   "Cad. Imobiliário:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   9
         Left            =   3870
         TabIndex        =   36
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
         Caption         =   "Novo Cadastro:"
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
      TabIndex        =   10
      Top             =   3120
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
      TabIndex        =   8
      Top             =   3120
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
      TabIndex        =   9
      Top             =   3120
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
      TabIndex        =   37
      Top             =   0
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   1138
      Icone           =   "TCIU102.frx":0047
   End
End
Attribute VB_Name = "TCIU102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InscricaoAntiga As String

Sub AtualizaRegistroPagamentos(ImovelNovo As String, ImovelAntigo As String)
    Bdados.AtualizaDados "TAB_GERACAO_TRIBUTO", Bdados.PreparaValor(Bdados.Converte(ImovelNovo, tctexto)), "TGT_INSCRICAO", "TGT_INSCRICAO='" & ImovelAntigo & "'"
    Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(Bdados.Converte(ImovelNovo, tctexto)), "TOC_INSCRICAO", "TOC_INSCRICAO='" & ImovelAntigo & "'"
    Bdados.AtualizaDados "TAB_CONTA_CONTRIBUINTE", Bdados.PreparaValor(Bdados.Converte(ImovelNovo, tctexto)), "TCC_INSCRICAO", "TCC_INSCRICAO='" & ImovelAntigo & "'"
    Bdados.AtualizaDados "TAB_DARM_RECEBIDO", Bdados.PreparaValor(Bdados.Converte(ImovelNovo, tctexto)), "TDR_INSCRICAO", "TDR_INSCRICAO='" & ImovelAntigo & "'"
    Bdados.AtualizaDados "TAB_PARCELAMENTO", Bdados.PreparaValor(Bdados.Converte(ImovelNovo, tctexto)), "TPA_TIM_IC", "TPA_TIM_IC ='" & ImovelAntigo & "'"
    
End Sub

Function InsereTerritorio() As Boolean
    Dim Campos As String
    Dim Valores As String, InscricaoCadastral As String
    Dim rs As VSRecordset
    Dim Sql As String
    
    
    InscricaoCadastral = Trim(txtxNovoCadastro)
    InscricaoAntiga = Trim(txtCodReduzido)
    'ALTERANDO TAB_IMOVEL
    Campos = "TIM_IC" ',TIM_SITUACAO_LOTE":
    Valores = Bdados.PreparaValor(Bdados.Converte(InscricaoCadastral, tctexto)) ', cboSit.ListIndex)
    Bdados.AtualizaDados "TAB_IMOVEL", Valores, Campos, "tim_ic= '" & InscricaoAntiga & "'"
    
    'ALTERANDO TAB_DETALHE_IMOVEL
    Campos = "TDI_TIM_IC": Valores = Bdados.PreparaValor(Bdados.Converte(InscricaoCadastral, tctexto))
    Bdados.AtualizaDados "TAB_DETALHE_IMOVEL", Valores, Campos, "TDI_TIM_IC='" & InscricaoAntiga & "'"
    
    'ALTERANDO TAB_IMOVEL_HISTORICO
    Campos = "TIM_IC" ',TIM_SITUACAO_LOTE":
    Valores = Bdados.PreparaValor(Bdados.Converte(InscricaoCadastral, tctexto)) ', cboSit.ListIndex)
    Bdados.AtualizaDados "TAB_IMOVEL_HISTORICO", Valores, Campos, "tim_ic='" & InscricaoAntiga & "'"
    
    'ALTERANDO TAB_DETALHE_IMOVEL_HISTORICO
    Campos = "TDI_TIM_IC": Valores = Bdados.PreparaValor(Bdados.Converte(InscricaoCadastral, tctexto))
    Bdados.AtualizaDados "TAB_DETALHE_IMOVEL_HISTORICO", Valores, Campos, "TDI_TIM_IC='" & InscricaoAntiga & "'"
    
    'ALTERANDO TABELAS DE PAGAMENTO
    AtualizaRegistroPagamentos InscricaoCadastral, InscricaoAntiga
    Informa "Inscrição Alterada."
    Bdados.FechaTabela rs
End Function

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
    Dim InscricaoMunicipal As String
    Dim SitCadastral As String
    Dim InscricaoCadastral As String
    
    If txtxNovoCadastro = "" Then
        Util.Avisa "Informe novo cadastro."
        txtxNovoCadastro.SetFocus
        Exit Sub
    End If
    InsereTerritorio
    InscricaoAntiga = ""
    cmdLimpar_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    cboSit.Visible = False
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
    If txtCodReduzido = "" Then Exit Sub
    
    Sql = "Select * from tab_imovel where tIM_ic ='" & Trim(txtCodReduzido) & "'"
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
        txtCPFOcupante = "" & rs!tim_cgc_cpf_ocupante
'        cboSit.ListIndex = IIf(IsNull(Rs!TIM_SITUACAO_LOTE), 0, "" & Rs!TIM_SITUACAO_LOTE) - 1
        txtIcAnterior = "" & rs!tim_ic_anterior
        txtCodMens = "" & rs!tim_COD_MENSAGEM
        txtZona = "" & rs!tim_ZONA
        txtCodBairro = "" & rs!tim_TBA_COD_BAIRRO
        txtCodBairro_LostFocus
        InscricaoAntiga = txtCodReduzido
        Screen.MousePointer = 0
        txtInscImob = "" & rs!tim_ic_auxiliar
    End If
End Sub

Private Sub txtInscImob_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Tem As String
    Dim aux As String
    
    If txtInscImob = "" Then Exit Sub
    
    'If Trim(InscricaoAntiga) <> "" Then
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            Sql = "Select * from tab_imovel where tim_ic_auxiliar ='" & txtInscImob & "'"
        Else
            Sql = "Select * from tab_imovel where tIM_ic ='" & txtInscImob & "'"
        End If
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
            txtCodReduzido = "" & rs!TIM_IC
            txtCPFOcupante = "" & rs!tim_cgc_cpf_ocupante
'            cboSit.ListIndex = IIf(IsNull(Rs!TIM_SITUACAO_LOTE), 0, "" & Rs!TIM_SITUACAO_LOTE)
            txtIcAnterior = "" & rs!tim_ic_anterior
            txtCodMens = "" & rs!tim_COD_MENSAGEM
            txtZona = "" & rs!tim_ZONA
            txtCodBairro = "" & rs!tim_TBA_COD_BAIRRO
            txtCodBairro_LostFocus
            InscricaoAntiga = txtCodReduzido
            Screen.MousePointer = 0
        End If
'    Else
'        If Not Confirma("Alterando lote " & InscricaoAntiga & ". Confirma?") Then
'            cmdLimpar_Click
'        End If
    'End If
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

