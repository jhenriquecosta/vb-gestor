VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.0#0"; "VTControles.ocx"
Begin VB.Form TCOB301 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VS"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMotivo 
      Appearance      =   0  'Flat
      Height          =   1260
      Left            =   990
      TabIndex        =   1
      Tag             =   "Motivo"
      Top             =   5055
      Width           =   9045
   End
   Begin Threed.SSFrame fra 
      Height          =   1395
      Index           =   2
      Left            =   30
      TabIndex        =   17
      Top             =   2160
      Width           =   10035
      _ExtentX        =   17701
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
      Caption         =   "Imposto"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtDamSendoEstornado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   960
         Left            =   3150
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "NÚMERO DO DAM SENDO ESTORNADO"
         Top             =   240
         Visible         =   0   'False
         Width           =   3915
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
         Left            =   8550
         MaxLength       =   14
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   8550
         MaxLength       =   14
         TabIndex        =   10
         Top             =   960
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
         Left            =   8550
         MaxLength       =   14
         TabIndex        =   9
         Top             =   600
         Width           =   1335
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
         Top             =   600
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
         Top             =   240
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   2265
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   27
         Top             =   277
         Width           =   1620
         _ExtentX        =   2858
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
         Caption         =   "Código do Tributo:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   28
         Top             =   667
         Width           =   1620
         _ExtentX        =   2858
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
         Caption         =   "Período Referência:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   6
         Left            =   6960
         TabIndex        =   29
         Top             =   667
         Width           =   1620
         _ExtentX        =   2858
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
         Caption         =   "Data Vencimento:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   7
         Left            =   7530
         TabIndex        =   30
         Top             =   990
         Width           =   1020
         _ExtentX        =   1799
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
         Caption         =   "Nº Parcela:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   12
         Left            =   6810
         TabIndex        =   34
         Top             =   277
         Visible         =   0   'False
         Width           =   1620
         _ExtentX        =   2858
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
         Caption         =   "Inscrição Cadastral:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   1455
      Index           =   1
      Left            =   30
      TabIndex        =   20
      Top             =   660
      Width           =   10035
      _ExtentX        =   17701
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
      Caption         =   "Contribuinte"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtDistrito 
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
         Left            =   6690
         MaxLength       =   2
         TabIndex        =   38
         ToolTipText     =   "DISTRITO"
         Top             =   210
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtPeriodoEstorno 
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
         Left            =   6690
         MaxLength       =   4
         TabIndex        =   41
         ToolTipText     =   "PERÍODO (ANO)"
         Top             =   540
         Visible         =   0   'False
         Width           =   1005
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
         Left            =   7710
         MaxLength       =   4
         TabIndex        =   40
         ToolTipText     =   "QUADRA"
         Top             =   210
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSetor 
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
         Left            =   7200
         MaxLength       =   2
         TabIndex        =   39
         ToolTipText     =   "SETOR"
         Top             =   210
         Visible         =   0   'False
         Width           =   495
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
         Top             =   270
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
         Top             =   630
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   990
         Width           =   7710
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   1020
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
         TabIndex        =   23
         Top             =   660
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
         TabIndex        =   35
         Top             =   300
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
         Caption         =   "NO. DAM:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdEstornaTudo 
         Height          =   375
         Left            =   8250
         TabIndex        =   43
         Top             =   210
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   661
         Caption         =   "Estornar Tudo"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   345
      Left            =   3780
      TabIndex        =   19
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
      Height          =   1395
      Index           =   0
      Left            =   30
      TabIndex        =   24
      Top             =   3600
      Width           =   10035
      _ExtentX        =   17701
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
      Caption         =   "Informações do Pagamento"
      Alignment       =   2
      ShadowStyle     =   1
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
         TabIndex        =   15
         Top             =   600
         Width           =   1845
      End
      Begin VB.TextBox txtCorrecao 
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
         TabIndex        =   14
         Top             =   240
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
         TabIndex        =   12
         Top             =   600
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
         TabIndex        =   13
         Top             =   960
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
         TabIndex        =   11
         Top             =   240
         Width           =   1845
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   262
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
         Caption         =   "Valor Original do DAM:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   180
         TabIndex        =   26
         Top             =   975
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
         Left            =   120
         TabIndex        =   31
         Top             =   615
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
         TabIndex        =   32
         Top             =   262
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
         TabIndex        =   33
         Top             =   660
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
   End
   Begin Threed.SSPanel lbl 
      Height          =   270
      Index           =   15
      Left            =   150
      TabIndex        =   36
      Top             =   5070
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
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1138
      Icone           =   "TCOB301.frx":0000
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   8910
      TabIndex        =   4
      Top             =   6390
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   375
      Left            =   6630
      TabIndex        =   2
      Top             =   6390
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Estornar"
      Acao            =   2
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdCancelar 
      Height          =   375
      Left            =   7770
      TabIndex        =   3
      Top             =   6390
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin Threed.SSCommand cmdAjuda 
      Height          =   435
      Left            =   60
      TabIndex        =   18
      ToolTipText     =   "Ajuda"
      Top             =   4500
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      MousePointer    =   14
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TCOB301.frx":031A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TCOB301.frx":0336
      Caption         =   "?"
      ButtonStyle     =   3
      PictureAlignment=   6
   End
End
Attribute VB_Name = "TCOB301"
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
    Dim sql As String
    Dim RsArr As VSRecordset
    Dim i As Byte
    Dim Imposto As Byte
    Dim cLSImposto As New VSImposto
    VerificaDigitacao = False
    sql = "SELECT tdr_im FROM Tab_Darm_Recebido WHERE " & _
    " tdr_im = '" & txtIM & "' AND tdr_sit_pago <> 2 and  tdr_tip_cod_imposto = '" & txtCodImposto & _
    "' AND tdr_periodo = " & IIf(Len(txtPeriodo) = 4, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2))
    
    If UCase(txtNomeImposto) = cLSImposto.NomeTributo(ttr_IPTU) Then
        sql = sql & " and tdr_tim_ic ='" & txtInscCadastral & "'"
    End If
    If Bdados.AbreTabela(sql, RsArr) Then
        Util.Avisa "Pagamento de DARM já informado anteriormente."
        Call cmdCancelar_Click
        Bdados.FechaTabela RsArr
        Exit Function
    End If
    If UCase(txtNomeImposto) = cLSImposto.NomeTributo(ttr_IPTU) Then
        sql = "Select tim_tci_im From Tab_Imovel where tim_ic='" & txtInscCadastral & "'"
        If Bdados.AbreTabela(sql, RsArr) Then
            If Trim(txtIM) <> RsArr(0) Then
                Call Util.Avisa("Este imóvel não está cadastrado por este contribuinte.")
                Bdados.FechaTabela RsArr
                Exit Function
            End If
            If Len(txtPeriodo) <> 4 Then
                Util.Avisa "Período inválido."
                Exit Function
            End If
        Else
            Call Util.Avisa("Inscrição Cadastral inválida.")
            Bdados.FechaTabela RsArr
            Exit Function
        End If
    End If
    Bdados.FechaTabela RsArr
    If CSng(txtTotalDarm) <> CSng(txtValorOriginal) + CSng(txtValorMulta) + CSng(txtValorJuro) + CSng(txtCorrecao) Then
        Util.Avisa "INFORMAÇÕES DE PAGAMENTO INVÁLIDAS:" & Chr(13) & _
        txtValorOriginal & "  +  " & txtValorMulta & "  +  " & txtValorJuro & " + " & txtCorrecao & "  <>  " & txtTotalDarm
        Exit Function
    End If
    If Len(Trim(txtPeriodo)) = 4 Then
        If CSng(Trim(txtPeriodo)) > Year(Date) Then
            Util.Avisa "Período inválido."
            Exit Function
        End If
    ElseIf Len(Trim(txtPeriodo)) = 7 Then
        If CByte(Left(Trim(txtPeriodo), 2)) > 12 Or CByte(Left(Trim(txtPeriodo), 2)) < 1 Then
            Util.Avisa "Mês inválido."
            Exit Function
        End If
        If CSng(Mid(Trim(txtPeriodo), 4)) > Year(Date) Then
            Util.Avisa "Ano inválido."
            Exit Function
        End If
    Else
        Util.Avisa "Período inválido."
        Exit Function
    End If
    If IsDate(Trim(txtDtVencimento)) Then
        If Right(Trim(txtDtVencimento), 4) < Right(Trim(txtPeriodo), 4) Then
            Util.Avisa "Ano da data de vencimento não pode ser menor que ano do período de referência."
            Exit Function
        Else
            If Len(Trim(txtPeriodo)) = 7 Then
                If Mid(Trim(txtDtVencimento), 4, 2) < Left(Trim(txtPeriodo), 2) And Right(Trim(txtDtVencimento), 4) > Right(Trim(txtPeriodo), 4) Then
                    Util.Avisa "Mês da data de vencimento não pode ser menor que mês do período de referência."
                    Exit Function
                End If
            End If
        End If
    Else
        Util.Avisa "Data de  vencimento inválida."
        Exit Function
    End If
    Bdados.FechaTabela RsArr
    
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
    
    Pagamento(1) = txtValorOriginal
    Pagamento(2) = txtValorJuro
    Pagamento(3) = txtValorMulta
    Pagamento(4) = txtCorrecao
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
    txtParcela = ""
    txtValorOriginal = ""
    txtValorJuro = ""
    txtValorMulta = ""
    txtCorrecao = ""
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
    txtIM.Locked = False
    txtDAM.SetFocus
    Vez = 0
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdEstornaTudo_Click()
' ESTORNA TODOS OS DAMS NÃO PAGOS ENCONTRADOS
Dim sql As String
Dim RsRegistros As VSRecordset
Dim rs As VSRecordset '

If Not Util.Confirma("O PROCESSO DE ESTORNO DOS DOCUMENTOS PARA O SETOR, QUADRA E PERÍODO INFORMADOS É PERIGOSO, CONTINUAR?") Then Exit Sub
'    If Temp.PegaParametro(Bdados, "ESTORNO GLOBAL DE DAM") <> "AUTORIZADO" Then
'        Util.Informa "Parâmetro Faltando. Não será efetuado o estorno."
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
    Screen.MousePointer = 11
    sql = "Select tgt_cod_pagamento from  Tab_Geracao_Tributo INNER JOIN Tab_Imposto ON " _
        & " Tab_Geracao_Tributo.tgt_tip_cod_imposto = Tab_Imposto.tip_cod_imposto LEFT OUTER JOIN " _
        & " Tab_Darm_Recebido ON Tab_Geracao_Tributo.tgt_cod_pagamento = Tab_Darm_Recebido.tdr_tgt_cod_pagamento  " _
        & " where  tgt_tip_cod_imposto=tip_cod_imposto and (tgt_ativo =0 or tgt_ativo is null) and " _
        & " tgt_periodo >= " & txtPeriodoEstorno & " and tgt_periodo <= " & txtPeriodoEstorno & " and " _
        & " tgt_tim_ic LIKE '" & Trim(txtDistrito) & Trim(txtSetor) & Trim(txtQuadra) & "%' and " _
        & " tgt_tip_cod_imposto not in ('NOTIFICA','EXTRATO') and Tab_Darm_Recebido.tdr_tgt_cod_pagamento is null " _
        & " AND TIP_COD_IMPOSTO not in ('NOTIFICA','EXTRATO')"
        
        If Bdados.AbreTabela(sql, RsRegistros) Then
            sql = ""
            While Not RsRegistros.EOF
                txtDamSendoEstornado = RsRegistros!tgt_cod_pagamento
                    DoEvents
                    Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "tgt_cod_pagamento =" & RsRegistros!tgt_cod_pagamento
                    Bdados.DeletaDados "TAB_GERACAO_TRIBUTO_PARCELA", "tgt_cod_pagamento =" & RsRegistros!tgt_cod_pagamento
                    Bdados.DeletaDados "Tab_Detalhe_Dam", "tdd_tgt_cod_pagamento =" & RsRegistros!tgt_cod_pagamento
                    Bdados.DeletaDados "Tab_Nota_Avulsa", "tna_cod_pagamento = " & RsRegistros!tgt_cod_pagamento
                    Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "tcc_codigo_conta = " & RsRegistros!tgt_cod_pagamento
                    DoEvents
                    sql = "Select TTD_TGT_COD_PAGAMENTO_TAXA,TDD_TIP_COD_IMPOSTO from tab_taxa_dam where TTD_TGT_COD_PAGAMENTO =" & RsRegistros!tgt_cod_pagamento
                    If Bdados.AbreTabela(sql, rs) Then
                        rs.MoveFirst
                        Do
                            Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "tgt_cod_pagamento=" & rs!TTD_TGT_COD_PAGAMENTO_TAXA
                            Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "tcc_codigo_conta = " & rs!TTD_TGT_COD_PAGAMENTO_TAXA
                            rs.MoveNext
                        Loop While Not rs.EOF
                    End If
                    Bdados.DeletaDados "TAB_TAXA_DAM", "TTD_TGT_COD_PAGAMENTO =" & RsRegistros!tgt_cod_pagamento
                    DoEvents
                    RsRegistros.MoveNext
            Wend
            Bdados.FechaTabela RsRegistros
            Bdados.FechaTabela rs
            Screen.MousePointer = 0
            Call cmdCancelar_Click
            Util.Informa "Todos os registros solicitados foram apagados."
        End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Conta As New ContaCorrente
    Dim Campos As String
    Dim Valores As String
    Dim sql As String '
    Dim rs As VSRecordset '
    If Confirma("Deseja estornar o DAM nº " & txtDAM & "?") Then
        If Not Edita.CriticaCampos(Me) Then Exit Sub
        Screen.MousePointer = 11
        
        Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "tgt_cod_pagamento =" & txtDAM
        Bdados.DeletaDados "TAB_GERACAO_TRIBUTO_PARCELA", "tgt_cod_pagamento =" & txtDAM
        Bdados.DeletaDados "Tab_Detalhe_Dam", "tdd_tgt_cod_pagamento =" & txtDAM
        Bdados.DeletaDados "Tab_Nota_Avulsa", "tna_cod_pagamento =" & txtDAM
        Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "tcc_codigo_conta = " & txtDAM
        '
        sql = "Select TTD_TGT_COD_PAGAMENTO_TAXA,TDD_TIP_COD_IMPOSTO from tab_taxa_dam where TTD_TGT_COD_PAGAMENTO =" & txtDAM
        If Bdados.AbreTabela(sql, rs) Then
            rs.MoveFirst
            Do
                Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "tgt_cod_pagamento=" & rs!TTD_TGT_COD_PAGAMENTO_TAXA
                Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "tcc_codigo_conta = " & rs!TTD_TGT_COD_PAGAMENTO_TAXA
                rs.MoveNext
            Loop While Not rs.EOF
        End If
        Bdados.DeletaDados "TAB_TAXA_DAM", "TTD_TGT_COD_PAGAMENTO =" & txtDAM
        Util.Informa "DAM " & txtDAM & " pertencente a IM n° " & txtIM & " estornado com sucesso."
        Call cmdCancelar_Click
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    
    ' PERMITE APARECER OS CAMPOS PARA O ESTORNO POR SETOR, QUADRA E PERIODO
    If Temp.PegaParametro(Bdados, "ESTORNO GLOBAL DE DAM") = "AUTORIZADO" Then
        If UCase(Aplicacoes.Usuario) = "ANDRE" Or UCase(Aplicacoes.Usuario) = "BOSING" Then
            txtDistrito.Visible = True
            txtSetor.Visible = True
            txtQuadra.Visible = True
            txtPeriodoEstorno.Visible = True
            txtDamSendoEstornado.Visible = True
            cmdEstornaTudo.Visible = True
        End If
    End If
End Sub

Private Sub txtCodImposto_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodImposto_LostFocus()
    Dim RsImp As VSRecordset
    Dim sql As String
    If Trim(txtCodImposto) = "" Then Exit Sub
    If Me.ActiveControl.Name = "cmdCancelar" Or Me.ActiveControl.Name = "cmdSair" Then Exit Sub
    sql = "Select tip_sigla_imposto from Tab_Imposto where tip_cod_imposto ='" & txtCodImposto & "'"
    If Bdados.AbreTabela(sql, RsImp) Then
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
    End If
    Bdados.FechaTabela RsImp
End Sub

Private Sub txtCorrecao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCorrecao_LostFocus()
    txtCorrecao = Edita.FormataTexto(txtCorrecao, Monetario, True)
End Sub

Private Sub txtDAM_LostFocus()
    Dim sql As String
    Dim rs As VSRecordset
    Dim RsParcela As VSRecordset
    Dim Conta As New ContaCorrente
    If Trim(txtDAM) <> "" Then
        sql = "Select tdr_data_pagamento from tab_darm_recebido where tdr_tgt_cod_pagamento=" & txtDAM & " and tdr_sit_pago <> 2"
        If Bdados.AbreTabela(sql, rs) Then
            Informa "DAM já baixado."
            txtDAM.SetFocus
            txtDAM.SelStart = 0
            txtDAM.SelLength = Len(Trim(txtDAM))
            Exit Sub
        End If
        Set rs = Conta.BuscaDam(txtDAM)
        If Not rs.EOF Then
            txtIM = "" & rs!TGT_im
            Call txtIm_LostFocus
            txtCodImposto = "" & rs!tgt_tip_cod_imposto
            Call txtCodImposto_LostFocus
            txtInscCadastral = "" & rs!TGT_tim_ic
            If Nvl("" & rs!tgt_tpa_num_parcelamento, 0) = 0 Then
                txtPeriodo = IIf(Len("" & rs!TGT_periodo) = 4, "" & rs!TGT_periodo, Right("" & rs!TGT_periodo, 2) & "/" & Left("" & rs!TGT_periodo, 4))
                If Left(txtPeriodo, 2) > 12 And Len(txtPeriodo) <> 4 Then txtPeriodo = Right(txtPeriodo, 4)
            Else
                sql = "Select tpa_periodo from tab_parcelamento where tpa_num_parcelamento =" & Nvl("" & rs!tgt_tpa_num_parcelamento, 0)
                If Bdados.AbreTabela(sql, RsParcela) Then
                    txtPeriodo = IIf(Len("" & RsParcela!tpa_periodo) = 4, "" & RsParcela!tpa_periodo, Right("" & RsParcela!tpa_periodo, 2) & "/" & Left("" & RsParcela!tpa_periodo, 4))
                    If Trim(txtPeriodo) = "" Then txtPeriodo = Year(Date)
                ElseIf Not IsNull(rs!TGT_periodo) Then
                    txtPeriodo = "" & rs!TGT_periodo
                Else
                    txtPeriodo = Year(Date)
                End If
            End If
            txtDtVencimento = "" & rs!tgt_data_vencimento
            txtValorOriginal = CDbl(Nvl("" & rs!TGT_VALOR_TRIBUTO, 0)) + CDbl(Nvl("" & rs!tgt_taxa_expediente, 0))
            txtValorOriginal = Edita.FormataTexto(txtValorOriginal, Monetario, True)
            txtValorMulta = Conta.ItemConta(txtDAM, EtdMulta)
            txtValorMulta = Edita.FormataTexto(txtValorMulta, Monetario, True)
            txtValorJuro = Conta.ItemConta(txtDAM, EtdJuros)
            txtValorJuro = Edita.FormataTexto(txtValorJuro, Monetario, True)
            txtCorrecao = Format(Nvl("" & rs!tgt_taxa_expediente, 0), Const_Monetario)
            txtTotalDarm = CDbl(txtValorMulta) + CDbl(txtValorJuro) + CDbl(Nvl("" & rs!TGT_VALOR_TRIBUTO, 0)) + CDbl(Nvl("" & rs!tgt_taxa_expediente, 0))
            txtTotalDarm = Edita.FormataTexto(txtTotalDarm, Monetario, True)
            txtParcela = Nvl("" & rs!TGT_PARCELA, 0)
        Else
            sql = "select * from tab_nota_avulsa where tna_cod_pagamento =" & txtDAM
            If Bdados.AbreTabela(sql, rs) Then
                txtIM = "" & rs!tna_tca_identidade_remetente
                Call txtIm_LostFocus
                txtCodImposto = BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto='" & Imposto.NomeTributo(ttr_ISSQN) & "'")
                txtNomeImposto = Imposto.NomeTributo(ttr_ISSQN)
                txtPeriodo = IIf(Len(rs!Tna_periodo) = 4, rs!Tna_periodo, Right(rs!Tna_periodo, 2) & "/" & Left(rs!Tna_periodo, 4))
                If Left(txtPeriodo, 2) > 12 And Len(txtPeriodo) <> 4 Then txtPeriodo = Right(txtPeriodo, 4)
                txtDtVencimento = UltimoDiaDoMes("" & rs!tna_data_emissao)
                txtValorOriginal = Format(rs!tna_valor_imposto, Const_Monetario)
                txtValorMulta = "0,00"
                txtValorJuro = "0,00"
                txtTotalDarm = CDbl(txtValorOriginal)
                txtTotalDarm = Edita.FormataTexto(txtTotalDarm, Monetario, True)
                txtCorrecao = "0,00"
                txtParcela = "0"
            Else
                Util.Avisa "Nº de Pagamento não encontrado."
                LimpaCampos Me
                txtDAM.SetFocus
            End If
        End If
    End If
    Bdados.FechaTabela rs
    Bdados.FechaTabela RsParcela
End Sub

Private Sub TxtDtPagamento_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtDistrito_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtDtVencimento_LostFocus()
    txtDtVencimento = Edita.FormataTexto(txtDtVencimento, Data)
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    Dim sql As String
    Dim rs As VSRecordset
    If Trim(txtIM) = "" Then Exit Sub
    If Me.ActiveControl.Name = "cmdCancelar" Or Me.ActiveControl.Name = "cmdSair" Then Exit Sub
    txtIM = Imposto.FormataInscricao(txtIM, InscContrib)
    sql = "select tci_nome from Tab_Contribuinte where tci_im ='" & txtIM & "'"
    If Bdados.AbreTabela(sql, rs) Then
        txtrazao = rs(0)
    Else
        sql = "select tci_nome from Tab_Contribuinte where tci_im ='" & txtIM & "'"
        If Bdados.AbreTabela(sql, rs) Then
            txtrazao = rs(0)
        Else
            If Trim(txtIM) <> Const_ImAvulso Then
                If Bdados.AbreTabela("Select tca_nome from tab_contribuinte_avulso where tca_identidade='" & txtIM & "'", rs) Then
                    txtrazao = "" & rs(0)
                End If
            End If
            Bdados.FechaTabela rs
            Exit Sub
        End If
    End If
    Bdados.FechaTabela rs
    Vez = 1
End Sub

Private Sub txtInscCadastral_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtInscCadastral_LostFocus()
    Dim sql As String
    Dim rs As VSRecordset
    txtInscCadastral = Imposto.FormataInscricao(txtInscCadastral, InscImovel)
    
    sql = "SELECT tgt_im from tab_geracao_tributo where tgt_tim_ic='" & txtInscCadastral & "'"
    If Bdados.AbreTabela(sql, rs) Then
        If rs(0) <> txtIM Then
            Call Avisa("Imóvel não cadastrado para este contribuinte.")
        End If
    Else
        Call Avisa(" Imóvel não cadastrado.")
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtParcela_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPeriodo_LostFocus()
    Dim Periodo As String
    Dim sql As String
    Dim rs As VSRecordset
    If Me.ActiveControl.Name = "cmdSair" Then Exit Sub
    If Len(txtPeriodo) = 6 Then
        Periodo = Mid(txtPeriodo, 3) & Left(txtPeriodo, 2)
        txtPeriodo = Left(txtPeriodo, 2) & "/" & Mid(txtPeriodo, 3)
    Else
        Periodo = txtPeriodo
    End If
    If Trim(txtPeriodo) <> "" Then
        sql = "SELECT tpi_tipo_tributo from Tab_Parametro_Imposto where tpi_tip_cod_imposto='" & txtCodImposto & "'"
        If Bdados.AbreTabela(sql, rs) Then
            If rs(0) = 1 Then
                sql = "SELECT tgt_data_vencimento FROM tab_geracao_tributo where tgt_tip_cod_imposto='" & _
                    txtCodImposto & "' and tgt_periodo=" & Periodo
                If Trim(txtInscCadastral) <> "" Then
                    sql = sql & " and tgt_tim_ic='" & txtInscCadastral & "'"
                Else
                    sql = sql & " and tgt_im='" & txtIM & "'"
                End If
                If Bdados.AbreTabela(sql, rs) Then
                    txtDtVencimento.Text = Format(rs(0), "dd/mm/yyyy")
                    txtDtVencimento.Enabled = False
                Else
                    Avisa "Período sem Crédito Tributário para este contribuinte."
                    Exit Sub
                End If
            Else
                txtDtVencimento.Enabled = True
                DoEvents
            End If
        Else
            Avisa "Imposto Inválido."
        End If
        Bdados.FechaTabela rs
    End If
End Sub

Private Sub txtQuadra_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtSetor_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtTotalDarm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtTotalDarm_LostFocus()
    txtTotalDarm = Edita.FormataTexto(txtTotalDarm, Monetario, True)
End Sub

Private Sub txtValorJuro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValorJuro_LostFocus()
    txtValorJuro = Edita.FormataTexto(txtValorJuro, Monetario, True)
End Sub

Private Sub txtValorMulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValorMulta_LostFocus()
    txtValorMulta = Edita.FormataTexto(txtValorMulta, Monetario, True)
End Sub

Private Sub txtValorOriginal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValorPago_Change()

End Sub

Private Sub txtValorPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub
