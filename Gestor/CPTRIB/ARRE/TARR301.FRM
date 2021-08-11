VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TARR301 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TARR301"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   ControlBox      =   0   'False
   Icon            =   "TARR301.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1138
      Icone           =   "TARR301.frx":030A
   End
   Begin VB.TextBox txtMotivo 
      Appearance      =   0  'Flat
      Height          =   960
      Left            =   1005
      TabIndex        =   1
      Tag             =   "Motivo"
      Top             =   5325
      Width           =   9060
   End
   Begin Threed.SSFrame fra 
      Height          =   1365
      Index           =   2
      Left            =   45
      TabIndex        =   19
      Top             =   2460
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   2408
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
      Caption         =   "Imposto"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox TxtDtPagamento 
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
         MaxLength       =   14
         TabIndex        =   10
         Tag             =   "Data Pagamento"
         Top             =   930
         Width           =   1305
      End
      Begin VB.TextBox txtInscCadastral 
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
         MaxLength       =   20
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox txtParcela 
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
         MaxLength       =   14
         TabIndex        =   12
         Top             =   930
         Width           =   495
      End
      Begin VB.TextBox txtDtVencimento 
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
         MaxLength       =   14
         TabIndex        =   9
         Top             =   570
         Width           =   1305
      End
      Begin VB.TextBox txtPeriodo 
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
         MaxLength       =   14
         TabIndex        =   8
         Top             =   570
         Width           =   1305
      End
      Begin VB.TextBox txtCodImposto 
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
         MaxLength       =   14
         TabIndex        =   6
         Top             =   210
         Width           =   915
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
         Left            =   3150
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   210
         Width           =   2265
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   450
         TabIndex        =   28
         Top             =   255
         Width           =   1545
         _ExtentX        =   2725
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
         Caption         =   "Código do Tributo:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   360
         TabIndex        =   29
         Top             =   615
         Width           =   1635
         _ExtentX        =   2884
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
         Caption         =   "Período Referência:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   6
         Left            =   6690
         TabIndex        =   30
         Top             =   615
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "Data Vencimento:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   7
         Left            =   7260
         TabIndex        =   31
         Top             =   975
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
         Caption         =   "Nº Parcela:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   12
         Left            =   6540
         TabIndex        =   35
         Top             =   255
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
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
         Caption         =   "Inscrição Cadastral:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   13
         Left            =   615
         TabIndex        =   36
         Top             =   975
         Width           =   1380
         _ExtentX        =   2434
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
         Caption         =   "Data Pagamento:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   1755
      Index           =   1
      Left            =   30
      TabIndex        =   21
      Top             =   645
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   3096
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
      Caption         =   "Contribuinte"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtEndereco 
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1290
         Width           =   7710
      End
      Begin VB.TextBox txtDAM 
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
         Left            =   2190
         MaxLength       =   14
         TabIndex        =   0
         Tag             =   "NO. DAM"
         Top             =   210
         Width           =   2235
      End
      Begin VB.TextBox txtIm 
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
         MaxLength       =   14
         TabIndex        =   5
         Top             =   570
         Width           =   2235
      End
      Begin VB.TextBox txtrazao 
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   930
         Width           =   7710
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   960
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
         Caption         =   "Nome ou Razão Social:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   5
         Left            =   330
         TabIndex        =   24
         Top             =   600
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
         Caption         =   "Inscrição Municipal:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   1230
         TabIndex        =   37
         Top             =   240
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
         Caption         =   "N°. DAM:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   16
         Left            =   1110
         TabIndex        =   41
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
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
         Caption         =   "Endereco:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   345
      Left            =   3780
      TabIndex        =   20
      Top             =   1110
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4020
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   990
      Width           =   1185
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7440
      Top             =   990
   End
   Begin Threed.SSFrame fra 
      Height          =   1365
      Index           =   0
      Left            =   45
      TabIndex        =   25
      Top             =   3900
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   2408
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
      Caption         =   "Informações do Pagamento"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtValorPago 
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
         Left            =   8070
         TabIndex        =   11
         Top             =   930
         Width           =   1845
      End
      Begin VB.TextBox txtTotalDarm 
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
         Left            =   8070
         TabIndex        =   17
         Top             =   570
         Width           =   1845
      End
      Begin VB.TextBox txtTaxas 
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
         Left            =   8070
         TabIndex        =   16
         Top             =   195
         Width           =   1845
      End
      Begin VB.TextBox txtValorJuro 
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
         Left            =   2220
         TabIndex        =   14
         Top             =   570
         Width           =   1845
      End
      Begin VB.TextBox txtValorMulta 
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
         Left            =   2220
         TabIndex        =   15
         Top             =   930
         Width           =   1845
      End
      Begin VB.TextBox txtValorOriginal 
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
         Left            =   2220
         TabIndex        =   13
         Top             =   195
         Width           =   1845
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   1
         Left            =   90
         TabIndex        =   26
         Top             =   210
         Width           =   2100
         _ExtentX        =   3704
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
         Caption         =   "Valor Original do Tributo:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   945
         Width           =   1950
         _ExtentX        =   3440
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
         Caption         =   "Valor de Multas:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   9
         Left            =   210
         TabIndex        =   32
         Top             =   585
         Width           =   1950
         _ExtentX        =   3440
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
         Caption         =   "Valor de Juros:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   10
         Left            =   6030
         TabIndex        =   33
         Top             =   210
         Width           =   1950
         _ExtentX        =   3440
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
         Caption         =   "Taxas:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   11
         Left            =   6030
         TabIndex        =   34
         Top             =   585
         Width           =   1950
         _ExtentX        =   3440
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
         Caption         =   "Total do DAM:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   14
         Left            =   6030
         TabIndex        =   38
         Top             =   945
         Width           =   1950
         _ExtentX        =   3440
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
         Caption         =   "Valor Pago(R$):"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSPanel lbl 
      Height          =   270
      Index           =   15
      Left            =   180
      TabIndex        =   39
      Top             =   5325
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
      Caption         =   "Motivo:"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   4
      RoundedCorners  =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   8940
      TabIndex        =   4
      Top             =   6420
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Sair"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdEstorno 
      Height          =   375
      Left            =   6510
      TabIndex        =   2
      Top             =   6420
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Estornar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdCancelar 
      Height          =   375
      Left            =   7740
      TabIndex        =   3
      Top             =   6420
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TARR301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Vez As Byte
Private DadosImposto(1 To 6) As String
Private Pagamento(1 To 5) As Double
Private Imposto As New VSImposto
Private Function VerificaDigitacao() As Boolean
    On Error GoTo trata
    Dim Sql As String
    Dim RsArr As VSRecordset
    Dim i As Byte
    Dim Imposto As Byte
    Dim cLSImposto As New VSImposto
    VerificaDigitacao = False
    Sql = "SELECT tdr_im FROM Tab_Darm_Recebido WHERE " & _
    " tdr_im = '" & txtIm & "' AND tdr_sit_pago <> 2 and tdr_tip_cod_imposto = '" & txtCodImposto & _
    "' AND tdr_periodo = " & IIf(Len(txtPeriodo) = 4, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2))
    
    If UCase(txtNomeImposto) = cLSImposto.NomeTributo(ttr_IPTU) Then
        Sql = Sql & " and tdr_tim_ic ='" & txtInscCadastral & "'"
    End If
    If Bdados.AbreTabela(Sql, RsArr) Then
        Util.Avisa "Pagamento de DARM já informado anteriormente."
        Call cmdCancelar_Click
        Bdados.FechaTabela RsArr
        Exit Function
    End If
    If UCase(txtNomeImposto) = cLSImposto.NomeTributo(ttr_IPTU) Then
        Sql = "Select tim_tci_im From Tab_Imovel where tim_ic='" & txtInscCadastral & "'"
        If Bdados.AbreTabela(Sql, RsArr) Then
            If Trim(txtIm) <> RsArr(0) Then
                Call Util.Avisa("Este imóvel não está cadastrado por este contribuinte.")
                txtInscCadastral.SetFocus
                Bdados.FechaTabela RsArr
                Exit Function
            End If
            If Len(txtPeriodo) <> 4 Then
                txtPeriodo.SetFocus
                Util.Avisa "Período inválido."
                Exit Function
            End If
        Else
            Call Util.Avisa("Inscrição Cadastral inválida.")
            txtInscCadastral.SetFocus
            Bdados.FechaTabela RsArr
            Exit Function
        End If
    End If
    Bdados.FechaTabela RsArr
    If CSng(txtTotalDarm) <> CSng(txtValorOriginal) + CSng(txtValorMulta) + CSng(txtValorJuro) Then
        Util.Avisa "INFORMAÇÕES DE PAGAMENTO INVÁLIDAS:" & Chr(13) & _
        txtValorOriginal & "  +  " & txtValorMulta & "  +  " & txtValorJuro & "  <>  " & txtTotalDarm
        txtTotalDarm.SetFocus
        Exit Function
    End If
    If Len(Trim(txtPeriodo)) = 4 Then
        If CSng(Trim(txtPeriodo)) > Year(Date) Then
            Util.Avisa "Período inválido."
            txtPeriodo.SetFocus
            Exit Function
        End If
    ElseIf Len(Trim(txtPeriodo)) = 7 Then
        If CByte(Left(Trim(txtPeriodo), 2)) > 12 Or CByte(Left(Trim(txtPeriodo), 2)) < 1 Then
            Util.Avisa "Mês inválido."
            txtPeriodo.SetFocus
            Exit Function
        End If
        If CSng(Mid(Trim(txtPeriodo), 4)) > Year(Date) Then
            Util.Avisa "Ano inválido."
            txtPeriodo.SetFocus
            Exit Function
        End If
    Else
        Util.Avisa "Período inválido."
        txtPeriodo.SetFocus
        Exit Function
    End If
    If IsDate(Trim(txtDtVencimento)) Then
        If Right(Trim(txtDtVencimento), 4) < Right(Trim(txtPeriodo), 4) Then
            Util.Avisa "Ano da data de vencimento não pode ser menor que ano do período de referência."
            txtDtVencimento.SetFocus
            Exit Function
        Else
            If Len(Trim(txtPeriodo)) = 7 Then
                If Mid(Trim(txtDtVencimento), 4, 2) < Left(Trim(txtPeriodo), 2) And Right(Trim(txtDtVencimento), 4) > Right(Trim(txtPeriodo), 4) Then
                    Util.Avisa "Mês da data de vencimento não pode ser menor que mês do período de referência."
                    txtDtVencimento.SetFocus
                    Exit Function
                End If
            End If
        End If
    Else
        Util.Avisa "Data de  vencimento inválida."
        txtDtVencimento.SetFocus
        Exit Function
    End If
    If IsDate(TxtDtPagamento) Then
        If Right(Trim(TxtDtPagamento), 4) < Right(Trim(txtPeriodo), 4) Then
            Util.Avisa "Ano da data de pagamento não pode ser menor que ano do período de referência."
            txtDtVencimento.SetFocus
            Exit Function
        Else
            If Len(Trim(txtPeriodo)) = 7 Then
                If Mid(Trim(TxtDtPagamento), 4, 2) < Left(Trim(txtPeriodo), 2) And Right(Trim(TxtDtPagamento), 4) > Right(Trim(txtPeriodo), 4) Then
                    Util.Avisa "Mês da data de pagamento não pode ser menor que mês do período de referência."
                    TxtDtPagamento.SetFocus
                    Exit Function
                End If
            End If
        End If
    Else
        Util.Avisa "Data de  pagamento inválida."
        TxtDtPagamento.SetFocus
        Exit Function
    End If
    VerificaDigitacao = True
    Exit Function
trata:
    If Err.Number = 13 Then
        Util.Avisa "Dados inválidos nas Informações de Pagamento."
    ElseIf Err.Number <> 0 And Err.Number <> 13 Then
        Util.Avisa Err.Number & " - " & Err.Description & "."
    End If
End Function

Private Function GuardaValores() As Boolean
    On Error GoTo trata
    GuardaValores = True
    DadosImposto(1) = txtCodImposto
    DadosImposto(2) = txtInscCadastral
    DadosImposto(3) = txtDtVencimento
    DadosImposto(4) = txtPeriodo
    DadosImposto(5) = TxtDtPagamento
    
    Pagamento(1) = txtValorOriginal
    Pagamento(2) = txtValorJuro
    Pagamento(3) = txtValorMulta
    Pagamento(4) = "" 'txtCorrecao
    Pagamento(5) = txtTotalDarm
    Exit Function
trata:
    GuardaValores = False
End Function

Sub LimpaValores()
    txtCodImposto = ""
    txtInscCadastral = ""
    txtDtVencimento = ""
    txtPeriodo = ""
    TxtDtPagamento = ""
    txtParcela = ""
    txtValorOriginal = ""
    txtValorJuro = ""
    txtValorMulta = ""
    txtTotalDarm = ""
    txtNomeImposto = ""
End Sub

Private Function BateValores() As Boolean
    Dim Controle As Control
    Dim i As Byte
    BateValores = False
    
    For i = 1 To 6
        For Each Controle In Controls
            If Controle.Tag = i Then
                If DadosImposto(i) <> Controle.Text Then
                    Call Util.Avisa("Informações do campo " & Controle.Tag & " não bate com a da primeira transcrição:" & DadosImposto(i) & " <> " & Controle.Text)
                    Exit Function
                End If
            End If
        Next
    Next
    BateValores = True
End Function

Private Sub cmdImprimir_Click()
    If Edita.CriticaCampos(Me) Then Exit Sub
End Sub

Private Sub cmdCancelar_Click()
    Edita.LimpaCampos Me
    txtInscCadastral.TabStop = False
    txtInscCadastral.Visible = False
    lbl(12).Visible = False
    txtIm.Locked = False
    txtDAM.SetFocus
    Vez = 0
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdEstorno_Click()
    Dim Conta As New ContaCorrente
    Dim Campos As String
    Dim Valores As String
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Obrig As New obrigacao
    TxtDtPagamento.Tag = ""
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    Campos = "tdr_motivo_cancelamento,tdr_data_cancelamento,tdr_tus_cod_usuario_cancelamento,tdr_sit_pago"
    Valores = Bdados.PreparaValor(txtMotivo, Bdados.Converte(Date, TCDataHora), Aplicacoes.Usuario, 2)
    Bdados.DeletaDados "TAB_DARM_RECEBIDO", "tdr_tgt_cod_pagamento=" & txtDAM
    Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "tcc_codigo_conta = " & txtDAM
    Conta.CriaContaContribuinte txtDAM
    
    Obrig.TrocaSitObrigacao txtDAM, etsCreditoOriginalAberto
    If CDbl(Nvl(Trim(txtTaxas), 0)) > 0 Then
        Sql = "Select TGT_COD_PAGAMENTO,TGT_TIP_COD_IMPOSTO,TGT_VALOR_TRIBUTO from tab_geracao_tributo where TGT_COD_PAGAMENTO_VINCULADO=" & txtDAM & " AND TGT_COD_PAGAMENTO <> " & txtDAM
        If Bdados.AbreTabela(Sql, Rs) Then
            Rs.MoveFirst
            Do
                Bdados.DeletaDados "TAB_DARM_RECEBIDO", "tdr_tgt_cod_pagamento=" & Rs!tgt_cod_Pagamento
                Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "tcc_codigo_conta = " & Rs!tgt_cod_Pagamento
                Conta.CriaContaContribuinte Rs!tgt_cod_Pagamento
                Rs.MoveNext
                DoEvents
            Loop While Not Rs.EOF
        End If
    End If
    Util.Informa "Dados eliminados com Segurança."
    Call cmdCancelar_Click
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub txtCodImposto_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtCodImposto_LostFocus()
    Dim RsImp As VSRecordset
    Dim Sql As String
    If Trim(txtCodImposto) = "" Then Exit Sub
    If Me.ActiveControl.Name = "cmdCancelar" Or Me.ActiveControl.Name = "cmdSair" Then Exit Sub
    Sql = "Select tip_sigla_imposto from Tab_Imposto where tip_cod_imposto ='" & txtCodImposto & "'"
    If Bdados.AbreTabela(Sql, RsImp) Then
        txtNomeImposto = RsImp(0)
        If Trim(txtNomeImposto) = Imposto.NomeTributo(ttr_IPTU) Then
            txtInscCadastral.TabStop = True
            txtInscCadastral.Visible = True
            lbl(12).Visible = True
        Else
            txtInscCadastral.TabStop = False
            txtInscCadastral.Visible = False
            lbl(12).Visible = False
            txtMotivo.SetFocus
        End If
    Else
        Call Util.Avisa("Código de Imposto inválido!")
    End If
    Bdados.FechaTabela RsImp
End Sub

Private Sub txtCorrecao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtDAM_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsParcela As VSRecordset
    Dim Conta As New ContaCorrente
    
    If Trim(txtDAM) <> "" Then
        
        Sql = "Select tdr_data_pagamento,TDR_PARCELA," _
            & " tdr_valor_total,tdr_valor_real_pago," _
            & " tdr_valor_real_juros,tdr_valor_real_multa," _
            & " tdr_inscricao,tdr_tip_cod_imposto," _
            & " tdr_tim_ic,tdr_periodo," _
            & " tdr_data_vencimento,tdr_valor_original," _
            & " tdr_taxa_expidiente from tab_darm_recebido" _
            & " where tdr_tgt_cod_pagamento=" & txtDAM & " and tdr_sit_pago <> 2"
            
        If Bdados.AbreTabela(Sql, Rs) Then
            If Rs(0) <> Null Then
                TxtDtPagamento = Rs(0)
            End If
            'ANDRE MUDA ISSO: USA O CAMPO TGT_COD_PAGAMENTO_VINCULADO
            txtValorPago = Format(CDbl(Nvl("" & Rs!tdr_valor_total, 0)) + CDbl(Nvl("" & Rs!tdr_taxa_expidiente, 0)), Const_Monetario)
            txtValorPago = Format(Rs!tdr_valor_real_pago, Const_Monetario)
            txtTotalDarm = Format(Rs!tdr_valor_total, Const_Monetario)
            txtValorJuro = Format(Rs!tdr_valor_real_juros, Const_Monetario)
            txtValorMulta = Format(Rs!tdr_valor_real_multa, Const_Monetario)
            
            txtIm = Rs!tdr_inscricao
            Call txtIm_LostFocus
            txtCodImposto = Rs!tdr_tip_cod_imposto
            txtParcela = Nvl("" & Rs!TDR_PARCELA, 0)
            Call txtCodImposto_LostFocus
            txtInscCadastral = "" & Rs!tdr_tim_ic
            If Not IsNull(Rs!tdr_periodo) Then
                txtPeriodo = IIf(Len(Rs!tdr_periodo) = 4, Rs!tdr_periodo, Right(Rs!tdr_periodo, 2) & "/" & Left(Rs!tdr_periodo, 4))
                If Left(txtPeriodo, 2) > 12 And Len(txtPeriodo) <> 4 Then txtPeriodo = Right(txtPeriodo, 4)
            Else
                Sql = "Select tpa_periodo from tab_parcelamento where tpa_num_parcelamento =" & Nvl("" & Rs!tgt_tpa_num_parcelamento, 0)
                If Bdados.AbreTabela(Sql, RsParcela) Then
                    txtPeriodo = IIf(Len(RsParcela!tpa_periodo) = 4, RsParcela!tpa_periodo, Right(RsParcela!tpa_periodo, 2) & "/" & Left(RsParcela!tpa_periodo, 4))
                ElseIf Not IsNull(Rs!tdr_periodo) Then
                    txtPeriodo = Rs!tdr_periodo
                Else
                    txtPeriodo = Year(Date)
                End If
            End If
            txtDtVencimento = Rs!tdr_data_vencimento
            txtValorOriginal = Format(Rs!tdr_valor_original, Const_Monetario)
            txtTaxas = Format(Nvl("" & Rs!tdr_taxa_expidiente, 0), Const_Monetario)
        Else
            Avisa "DAM sem baixa registrada."
            txtDAM.SetFocus
            Screen.MousePointer = 0
            Bdados.FechaTabela Rs
            Exit Sub
        End If
    End If
    Bdados.FechaTabela Rs
    Bdados.FechaTabela RsParcela
End Sub

Private Sub TxtDtPagamento_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub TxtDtPagamento_LostFocus()
    Dim Conta As New ContaCorrente
    TxtDtPagamento = Edita.FormataTexto(TxtDtPagamento, Data)
    If Trim(TxtDtPagamento) = "" Then Exit Sub
    If DateDiff("d", Date, TxtDtPagamento) <= 0 Then
        txtValorJuro = Format(Conta.CalculaValoresJurosAvulsos(txtCodImposto, IIf(Len(txtPeriodo) = 4, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2)), EtcCreditoTributario, TxtDtPagamento, txtDtVencimento, txtValorOriginal), Const_Monetario)
        txtValorMulta = Format(Conta.CalculaValoresMultaAvulsos(txtCodImposto, IIf(Len(txtPeriodo) = 4, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2)), EtcCreditoTributario, TxtDtPagamento, txtDtVencimento, txtValorOriginal), Const_Monetario)
        If DateDiff("d", txtDtVencimento, TxtDtPagamento) <= 0 Then
            txtValorJuro = "0,00"
            txtValorMulta = "0,00"
        End If
        txtTotalDarm = CDbl(txtValorOriginal) + CDbl(txtValorJuro) + CDbl(txtValorMulta)
    Else
        Avisa "Data de Pagamento inválido."
        TxtDtPagamento.SetFocus
    End If
End Sub

Private Sub txtDtVencimento_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtDtVencimento_LostFocus()
    txtDtVencimento = Edita.FormataTexto(txtDtVencimento, Data)
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtIm_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    If Trim(txtIm) = "" Then Exit Sub
    If Me.ActiveControl.Name = "cmdCancelar" Or Me.ActiveControl.Name = "cmdSair" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtrazao, txtEndereco)
    Bdados.FechaTabela Rs
    Vez = 1
End Sub

Private Sub txtInscCadastral_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtInscCadastral_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    txtInscCadastral = Imposto.FormataInscricao(txtInscCadastral, InscImovel)
    
    Sql = "SELECT tgt_im from tab_geracao_tributo where tgt_tim_ic='" & txtInscCadastral & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If Rs(0) <> txtIm Then
            Call Avisa("Imóvel não cadastrado para este contribuinte.")
            txtInscCadastral.SetFocus
        End If
    Else
        Call Avisa(" Imóvel não cadastrado.")
        txtInscCadastral.SetFocus
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtParcela_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtPeriodo_LostFocus()
    Dim Periodo As String
    Dim Sql As String
    Dim Rs As VSRecordset
    If Me.ActiveControl.Name = "cmd" Then Exit Sub
    If Len(txtPeriodo) = 6 Then
        Periodo = Mid(txtPeriodo, 3) & Left(txtPeriodo, 2)
        txtPeriodo = Left(txtPeriodo, 2) & "/" & Mid(txtPeriodo, 3)
    Else
        Periodo = txtPeriodo
    End If
    If Trim(txtPeriodo) <> "" Then
        Sql = "SELECT tpi_tipo_tributo from Tab_Parametro_Imposto where tpi_tip_cod_imposto='" & txtCodImposto & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            If Rs(0) = 1 Then
                Sql = "SELECT tgt_data_vencimento FROM tab_geracao_tributo where tgt_tip_cod_imposto='" & _
                    txtCodImposto & "' and tgt_periodo=" & Periodo
                If Trim(txtInscCadastral) <> "" Then
                    Sql = Sql & " and tgt_tim_ic='" & txtInscCadastral & "'"
                Else
                    Sql = Sql & " and tgt_im='" & txtIm & "'"
                End If
                If Bdados.AbreTabela(Sql, Rs) Then
                    txtDtVencimento.Text = Format(Rs(0), "dd/mm/yyyy")
                    txtDtVencimento.Enabled = False
                Else
                    Avisa "Período sem Crédito Tributário para este contribuinte."
                    txtPeriodo.SetFocus
                    Exit Sub
                End If
            Else
                txtDtVencimento.Enabled = True
                DoEvents
            End If
        Else
            Avisa "Imposto Inválido."
            txtPeriodo.SetFocus
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtTotalDarm_Change()
    txtValorOriginal = Edita.FormataTexto(txtValorOriginal, Monetario, True)
End Sub

Private Sub txtTotalDarm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtTotalDarm_LostFocus()
    txtTotalDarm = Edita.FormataTexto(txtTotalDarm, Monetario, True)
End Sub

Private Sub txtValorJuro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtValorJuro_LostFocus()
    txtValorJuro = Edita.FormataTexto(txtValorJuro, Monetario, True)
End Sub

Private Sub txtValorMulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtValorMulta_LostFocus()
    txtValorMulta = Edita.FormataTexto(txtValorMulta, Monetario, True)
End Sub

Private Sub txtValorOriginal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtValorOriginal_LostFocus()
    txtValorOriginal = Edita.FormataTexto(txtValorOriginal, Monetario, True)
End Sub

Private Sub txtValorPago_Change()
    txtValorOriginal = Edita.FormataTexto(txtValorOriginal, Monetario, True)
End Sub

Private Sub txtValorPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub
