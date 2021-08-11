VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TOBR105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleMode       =   0  'User
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   66
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TOBR105.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   60
      Top             =   6945
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   5610
         TabIndex        =   27
         Top             =   105
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   4245
         TabIndex        =   26
         Top             =   105
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Emitir DAM"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   6720
         TabIndex        =   28
         Top             =   105
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   90
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   58
      Top             =   4800
      Visible         =   0   'False
      Width           =   795
   End
   Begin Threed.SSFrame fra 
      Height          =   2280
      Index           =   2
      Left            =   90
      TabIndex        =   36
      Top             =   4665
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   4022
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
      Caption         =   "Detalhes"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VTOcx.txtVISUAL txtdam 
         Height          =   300
         Left            =   6360
         TabIndex        =   21
         Top             =   555
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Caption         =   ""
         Text            =   ""
         Formato         =   5
         AlinhamentoTexto=   1
      End
      Begin VB.TextBox txtAliquota 
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
         Left            =   2910
         TabIndex        =   16
         Text            =   "0"
         Top             =   870
         Width           =   1275
      End
      Begin VB.TextBox txtTotalImposto 
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
         Left            =   6360
         TabIndex        =   25
         Tag             =   " "
         Text            =   "0"
         Top             =   1860
         Width           =   1275
      End
      Begin VB.TextBox txtJuros 
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
         Left            =   6360
         TabIndex        =   24
         Text            =   "0"
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txtImposto 
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
         Left            =   6360
         TabIndex        =   22
         Text            =   "0"
         Top             =   870
         Width           =   1275
      End
      Begin VB.TextBox txtSaldo 
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
         Left            =   6360
         TabIndex        =   20
         Text            =   "0"
         Top             =   210
         Width           =   1275
      End
      Begin VB.TextBox txtMaterial 
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
         Left            =   2910
         TabIndex        =   18
         Text            =   "0"
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txtTotalNotas 
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
         Left            =   2910
         TabIndex        =   17
         Text            =   "0"
         Top             =   1200
         Width           =   1275
      End
      Begin VB.TextBox txtNotaFinal 
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
         Left            =   2910
         TabIndex        =   15
         Text            =   "0"
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox txtNotaInicial 
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
         Left            =   2910
         TabIndex        =   14
         Text            =   "0"
         Top             =   210
         Width           =   1275
      End
      Begin VB.TextBox txtMulta 
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
         Left            =   6360
         TabIndex        =   23
         Text            =   "0"
         Top             =   1200
         Width           =   1275
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   11
         Left            =   1785
         TabIndex        =   37
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "Nº Nota Incial:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   18
         Left            =   5835
         TabIndex        =   38
         Top             =   1275
         Width           =   510
         _ExtentX        =   900
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
         Caption         =   "Multa:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   2
         Left            =   1830
         TabIndex        =   44
         Top             =   630
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
         Caption         =   "Nº Nota Final:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   5
         Left            =   1320
         TabIndex        =   30
         Top             =   1290
         Width           =   1620
         _ExtentX        =   2858
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
         Caption         =   "Total em Notas R$ :"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   1575
         Width           =   2760
         _ExtentX        =   4868
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
         Caption         =   "Valor de material sujeito ao ICMS:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   10
         Left            =   4950
         TabIndex        =   45
         Top             =   270
         Width           =   1380
         _ExtentX        =   2434
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
         Caption         =   "Saldo Tributável:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   13
         Left            =   4680
         TabIndex        =   46
         Top             =   945
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
         Caption         =   "Imposto a Recolher:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   14
         Left            =   5820
         TabIndex        =   47
         Top             =   1620
         Width           =   510
         _ExtentX        =   900
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
         Caption         =   "Juros:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   15
         Left            =   4950
         TabIndex        =   48
         Top             =   1980
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
         Caption         =   "Total a Recolher:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   24
         Left            =   2025
         TabIndex        =   29
         Top             =   945
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
         Caption         =   "Aliquota %"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   25
         Left            =   5580
         TabIndex        =   64
         Top             =   600
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
         Caption         =   "TXDAM:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   26
         Left            =   2385
         TabIndex        =   65
         Top             =   1920
         Width           =   525
         _ExtentX        =   926
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
         Caption         =   "Título:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   315
         Left            =   2910
         TabIndex        =   19
         Top             =   1875
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   3285
      Index           =   0
      Left            =   90
      TabIndex        =   32
      Top             =   660
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   5794
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
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.ComboBox cboConta 
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
         ItemData        =   "TOBR105.frx":2123
         Left            =   5940
         List            =   "TOBR105.frx":2130
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   -405
         Width           =   825
      End
      Begin VTOcx.txtVISUAL txtDtVenc 
         Height          =   315
         Left            =   1425
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         Formato         =   0
      End
      Begin VB.ComboBox cboTaxa 
         DataField       =   "ttl_nome"
         DataSource      =   "dtTipLogr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         ItemData        =   "TOBR105.frx":213D
         Left            =   2775
         List            =   "TOBR105.frx":213F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   4890
      End
      Begin VB.ComboBox cboImposto 
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
         ItemData        =   "TOBR105.frx":2141
         Left            =   60
         List            =   "TOBR105.frx":2143
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Imposto"
         Top             =   225
         Width           =   7575
      End
      Begin VTOcx.cboVISUAL CboItem 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   615
         Visible         =   0   'False
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VB.TextBox txtFatorAleatorio 
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
         Left            =   5010
         TabIndex        =   5
         Top             =   1155
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtFator 
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
         Left            =   3810
         TabIndex        =   4
         Top             =   1155
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtParcela 
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
         Left            =   6870
         TabIndex        =   6
         Top             =   1155
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Frame fraIm 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   705
         Left            =   60
         TabIndex        =   49
         Top             =   2550
         Visible         =   0   'False
         Width           =   7635
         Begin VB.TextBox txtImovel 
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
            Left            =   1860
            TabIndex        =   12
            Top             =   315
            Width           =   5715
         End
         Begin VB.ComboBox cboIC 
            BeginProperty Font 
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
            ItemData        =   "TOBR105.frx":2145
            Left            =   30
            List            =   "TOBR105.frx":2147
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   300
            Width           =   1815
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   6
            Left            =   60
            TabIndex        =   62
            Top             =   45
            Width           =   1500
            _ExtentX        =   2646
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
            Caption         =   "Insc. Cadastral:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin VB.TextBox txtContribuinte 
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
         Left            =   75
         TabIndex        =   10
         Tag             =   "Contribuinte"
         Top             =   2220
         Width           =   7545
      End
      Begin VB.TextBox txtPeriodo 
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
         TabIndex        =   7
         Tag             =   "Exercicio"
         Top             =   1680
         Width           =   1335
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
         Left            =   90
         TabIndex        =   2
         Top             =   1155
         Width           =   1065
      End
      Begin VB.TextBox txtCgc 
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
         Left            =   1710
         TabIndex        =   3
         Top             =   1155
         Width           =   2025
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   90
         TabIndex        =   33
         Top             =   30
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
         Caption         =   "Tributo:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   4
         Left            =   1710
         TabIndex        =   34
         Top             =   945
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
         Caption         =   "CNPJ/CPF:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   7
         Left            =   5460
         TabIndex        =   35
         Top             =   2280
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
         Caption         =   "Seção:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   9
         Left            =   90
         TabIndex        =   40
         Top             =   945
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "Insc. Municipal:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   12
         Left            =   90
         TabIndex        =   41
         Top             =   1980
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
         Caption         =   "Contribuinte:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   42
         Top             =   1455
         Width           =   1170
         _ExtentX        =   2064
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
         Caption         =   "Exercício:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   1
         Left            =   1425
         TabIndex        =   43
         Top             =   1455
         Width           =   1170
         _ExtentX        =   2064
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
         Caption         =   "Vencimento"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   16
         Left            =   5940
         TabIndex        =   50
         Top             =   -645
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
         Caption         =   "Conta:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   17
         Left            =   6870
         TabIndex        =   51
         Top             =   915
         Visible         =   0   'False
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
         Caption         =   "Parcela:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   20
         Left            =   2775
         TabIndex        =   54
         Top             =   1455
         Visible         =   0   'False
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
         Caption         =   "Taxa:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   21
         Left            =   3810
         TabIndex        =   55
         Top             =   945
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Multiplicador:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   22
         Left            =   5040
         TabIndex        =   57
         Top             =   945
         Visible         =   0   'False
         Width           =   480
         _ExtentX        =   847
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
         Caption         =   "Fator:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesq 
         Height          =   300
         Index           =   0
         Left            =   1170
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1155
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   23
         Left            =   135
         TabIndex        =   63
         Top             =   645
         Visible         =   0   'False
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
         Caption         =   "Sub Movimento:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.Timer tmr 
      Interval        =   10
      Left            =   2430
      Top             =   4875
   End
   Begin Threed.SSFrame fra 
      Height          =   735
      Index           =   1
      Left            =   90
      TabIndex        =   52
      Top             =   3930
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   1296
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
      Caption         =   "Detalhes"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtObservacao 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   1350
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   6255
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   19
         Left            =   120
         TabIndex        =   53
         Top             =   240
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
         Caption         =   "Observações :"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.TextBox txtEnderecoContrib 
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
      Left            =   660
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   1230
      Width           =   5895
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   1138
      Icone           =   "TOBR105.frx":2149
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1200
      TabIndex        =   39
      Top             =   180
      Width           =   375
   End
   Begin VTOcx.grdVISUAL Grdtaxas 
      Height          =   6555
      Left            =   7890
      TabIndex        =   67
      Top             =   675
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   11562
      Caption         =   "Taxas"
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
      Ordenavel       =   0   'False
   End
End
Attribute VB_Name = "TOBR105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Imposto As New VSImposto
Dim CodImposto As String
Dim Exercicio As String
Dim Conta As New ContaCorrente
Dim Aliquota As Double
Dim CodTaxa As String
Dim Incidencia As Integer
'Variaveis para o Report
Dim InscMuni As String
Dim RazaoSocial As String
Dim Documento As String
Dim Localizacao As String
Dim Data_Vencimento As String
Dim Codigo_Imovel As String
Dim Valor_Imposto As String
Dim CPFCNPJ As String
Dim Endereco As String
Dim Bairro As String
Dim Cod_Atividade As String
Dim Cod_Cidade As String
Dim Cep As String
Dim Uf As String
Dim Cod_Tributo As String
Dim Juro As String
Dim Multa As String
Dim TotalImposto As String
Dim TaxaServico As Double
Dim BaseDeCalculo As String
Dim VetLinhas(0 To 5) As String
Dim Linhas As Byte
Dim ObsAux As String
Dim NomeImposto As String
Dim TributoTaxa As Boolean
Dim TributoTaxaFixa As Double
Dim Tributo As Double
Dim Alvara As Double
Dim PosTraco As Byte
Dim TSU As Double
Dim AreaConstruida As Double
Dim AreaTotal As Double
Dim ValorTerreno As Double
Dim Valoredific As Double
Dim Zona As Integer
Dim ValorMetro As Double
Dim TaxaParcela As Double
Dim Desconto As String
Dim Reducao As String
Dim CodPagamento As String
Dim CodTaxaEmissaoDam As String
Dim blnConsultaIM As Boolean
Dim String_Taxas As String
Dim Total_Taxas As Double

'Dim Admin As VSAFuncoes.VSACalculos
Function BuscaValorTaxa(NomeTaxa As String, ByRef CodTributo As String) As Double
    Dim Rs                                      As VSRecordset
    Dim Sql                                      As String
    Sql = " select tpi_tip_cod_imposto,tpi_valor_taxa_fixa  from Tab_Parametro_Imposto " & _
    " WHERE  tpi_tip_cod_imposto = (SELECT TIP_COD_IMPOSTO FROM TAB_IMPOSTO WHERE TIP_NOME_IMPOSTO ='" & NomeTaxa & "')"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaValorTaxa = IIf(IsNull(Rs!tpi_valor_taxa_fixa), 0, Rs!tpi_valor_taxa_fixa)
        CodTributo = Rs!tpi_tip_cod_imposto
    Else
        BuscaValorTaxa = 0
    End If
    Bdados.FechaTabela Rs
End Function
Sub GeraIptu()
        Dim Reducao As Double
        Dim Sql As String
        Dim RsCob As VSRecordset
        Dim Rs As VSRecordset
        Dim Correcao As Double
        Sql = "SELECT * FROM tab_imovel,Tab_Contribuinte where "

        If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 1) = 1 Then
            If CInt(Right(cboIC, 3)) <> 200 Then
                Sql = Sql & " tim_ic ='" & cboIC & "' and tim_tci_im=tci_im AND TIM_SITUACAO_LOTE <> 1"
            Else
                Sql = Sql & " tim_ic >'" & cboIC & "' and  tim_ic  <= '" & Left(cboIC, 12) & "300' AND tim_tci_im=tci_im "
            End If
        Else
            Sql = Sql & " tim_ic ='" & cboIC & "' and tim_tci_im=tci_im "
        End If
        If Bdados.AbreTabela(Sql, RsCob, Registros) Then
            If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 1) = 1 Then
                Dim CalculoIptu As New VSIptu
                If CInt(txtPeriodo) >= CInt(Nvl(Temp.PegaParametro(Bdados, "ANO PGV"), 9999)) Then
                    CalculoIptu.AnoLancamento = CInt(txtPeriodo)
                    CalculoIptu.CarregaDetalheLote Trim(cboIC)
                    CalculoIptu.CalculaValorIptu
                Else
                    Call Imposto.GeraIptu(cip_Balsas, RsCob, CInt(txtPeriodo), CInt(txtPeriodo), tgi_SemParcelas)
                End If
            Else
               txtImposto = Imposto.CriaIptu(RsCob, txtPeriodo, txtPeriodo)
            End If
            If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 1) = 1 Then Exit Sub

            Reducao = Imposto.BuscaReducao(CodImposto, Right(txtPeriodo, 4))
            Sql = "select tim_valor from tab_imovel where tim_ic='" & cboIC & "'"
            If Bdados.AbreTabela(Sql, Rs) Then
                BaseDeCalculo = CDbl(Nvl(Rs(0), 0)) - (CDbl(Nvl(Rs(0), 0)) * Reducao)
            End If

            Sql = "select tgt_taxa_expediente,TGT_COD_PAGAMENTO,tgt_parcela,TGT_VALOR_TRIBUTO,TGT_VALOR_JUROS,TGT_VALOR_MULTA,TGT_TIP_COD_IMPOSTO,TGT_PERIODO,TGT_DATA_VENCIMENTO from Tab_Geracao_Tributo where tgt_inscricao='" & txtIM & "' and tgt_tip_cod_imposto='" & CodImposto & "' and tgt_periodo=" & txtPeriodo & " AND  TGT_PARCELA = " & IIf(cboConta = 1, 0, 1) & " AND TGT_TIM_IC ='" & cboIC & "'"
            If Bdados.AbreTabela(Sql, Rs) Then
                TaxaParcela = Nvl("" & Rs!tgt_taxa_expediente, 0)
                CodPagamento = Rs(1)
                txtParcela = Nvl("" & Rs!TGT_PARCELA, 0)
                txtTotalImposto = Rs!TGT_VALOR_TRIBUTO
                If NomeImposto = Imposto.NomeTributo(ttr_ITU) Or NomeImposto = Imposto.NomeTributo(ttr_IPTU) And txtPeriodo < Year(Date) Then
                    Correcao = Conta.CalculaValoresCorrecaoAvulso(Rs!TGT_TIP_COD_IMPOSTO, Rs!TGT_PERIODO, Rs!TGT_DATA_VENCIMENTO, Date, Rs!TGT_VALOR_TRIBUTO)
                    txtMulta = Conta.CalculaValoresMultaAvulsos(Rs!TGT_TIP_COD_IMPOSTO, Rs!TGT_PERIODO, EtcCreditoTributario, Date, Rs!TGT_DATA_VENCIMENTO, Rs!TGT_VALOR_TRIBUTO + Correcao)
                    txtJuros = Conta.CalculaValoresJurosAvulsos(Rs!TGT_TIP_COD_IMPOSTO, Rs!TGT_PERIODO, EtcCreditoTributario, Date, Rs!TGT_DATA_VENCIMENTO, Rs!TGT_VALOR_TRIBUTO + Correcao)
                Else
                    txtMulta = 0
                    txtJuros = 0
                End If
                txtJuros = Edita.FormataTexto(txtJuros, Monetario, True)
                txtMulta = Edita.FormataTexto(txtMulta, Monetario, True)
                Valor_Imposto = txtTotalImposto
                txtTotalImposto = Format(CDbl(Nvl(txtTotalImposto, 0)) + CDbl(Nvl(txtMulta, 0)) + CDbl(Nvl(txtJuros, 0)), Const_Monetario)
                Juro = txtJuros
                Multa = txtMulta
                TotalImposto = txtTotalImposto
                txtImposto = txtTotalImposto
                TaxaServico = Nvl("" & Rs!tgt_taxa_expediente, 0)
            End If
            Sql = "Select TGE_NOME from tab_geral where TGE_TIPO = 755 and TGE_CODIGO > 0"
            If Bdados.AbreTabela(Sql, Rs) Then
                Desconto = Nvl("" & Rs(0), 0)
            End If

            Sql = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & cboIC & _
                "' and (tdi_tco_cod_componente=110 or tdi_tco_cod_componente=108)"
            If Bdados.AbreTabela(Sql, Rs) Then
                Rs.MoveFirst
                Do While Not Rs.EOF
                    If Rs(0) = 110 Then
                        AreaTotal = Nvl("" & Rs(1), 0)
                    ElseIf Rs(0) = 108 Then
                        AreaConstruida = Nvl(Rs(1), 0)
                    End If
                    Rs.MoveNext
                Loop
            End If
            Bdados.FechaTabela Rs
            Sql = "select tvl_valor from tab_valor_terreno where tvl_tlg_cod_logradouro=(" & _
                " select tim_tlg_cod_logradouro from tab_imovel where tim_ic='" & cboIC & "')"
            If Bdados.AbreTabela(Sql, Rs) Then
                ValorMetro = Format(Rs(0), Const_Monetario)
            End If
        Else
            Informa "IPTU não foi gerado."
        End If
        Bdados.FechaTabela Rs
End Sub

Sub GeraAlvara()
    Dim DataVenc As String
    Dim Valores As String
    Dim Campos As String
    Dim Correcao As Double
    DataVenc = txtDtVenc
    Alvara = Imposto.CalculaAlvara(IIf(Trim(txtIM) = "", txtCgc, txtIM), txtPeriodo, TaxaServico, Trim(Mid(cboImposto, PosTraco + 2)), DataVenc, CodPagamento, CodTaxa)
    txtTotalImposto = Alvara
    If Alvara = 0 Then
        Informa "Faltando informaçãoes necessárias para cálculo do ALVARA do contribuinte "
        Exit Sub
    End If
    If IsNumeric(DataVenc) Then
        txtDtVenc = Format(Right(DataVenc, 2) & "/" & Mid(DataVenc, 5, 2) & "/" & Left(DataVenc, 4), "dd/mm/yyyy")
    Else
        txtDtVenc = DataVenc
    End If
    Data_Vencimento = DataVenc
    DoEvents
    txtImposto = Alvara
    Valor_Imposto = txtTotalImposto
'    Call Conta.CriaContaContribuinte(CodPagamento)
'    Call Conta.MovimentaContaContribuinte(CodPagamento)
    Correcao = Conta.CalculaValoresCorrecaoAvulso(CodImposto, txtPeriodo, txtDtVenc, Format(Date, "dd/mm/yyyy"), CDbl(txtTotalImposto))
    txtMulta = Conta.CalculaValoresMultaAvulsos(CodImposto, txtPeriodo, cboConta, Format(Date, "dd/mm/yyyy"), txtDtVenc, CDbl(txtTotalImposto) + Correcao)
    txtJuros = Conta.CalculaValoresJurosAvulsos(CodImposto, txtPeriodo, cboConta, Format(Date, "dd/mm/yyyy"), txtDtVenc, CDbl(txtTotalImposto) + Correcao)
    txtTotalImposto = CDbl(Nvl(txtTotalImposto, 0)) + CDbl(Nvl(txtMulta, 0)) + CDbl(Nvl(txtJuros, 0))
    TotalImposto = CDbl(txtTotalImposto)
    Juro = txtJuros
    Multa = txtMulta
End Sub

Sub GeraIssqn()
    Dim Campos As String
    Dim Valores As String
    Dim Novo As Boolean
    Dim cLSImposto As New VSImposto
    Dim DAMOriginal As String
    Dim Obrig As New Obrigacao
    Dim Pos As Integer
    If CDbl(txtNotaInicial) > CDbl(txtNotaFinal) Then
        Informa "Nº de nota inicial não pode ser maior que nº de nota final"
        Screen.MousePointer = 0
        txtNotaInicial.SetFocus
    Exit Sub
    End If
    Incidencia = 0
    If NomeImposto <> Imposto.NomeTributo(ttr_ISSQN) Then
        Incidencia = cLSImposto.BuscaNumeroIncidencia(txtIM, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), CodImposto)
    End If
    Pos = InStr(cboImposto.Text, " # ")
    CodPagamento = Conta.GeraCodPagamento(54)  ' DAM EM BRANCO...
    DAMOriginal = CodPagamento
    If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNEST) Then
        If Nvl(CStr(cboTipo.Coluna(1).Valor), 1) = 1 Then 'DAM BRANCO
            Conta.GeraPagamento txtIM, "", CodImposto, Left(txtPeriodo, 2) & Right(txtPeriodo, 4), txtDtVenc, 0, 0, 0, CodPagamento, 0, Nvl(txtParcela, 0), TaxaServico, CodTaxa, , , Incidencia, , etgDAM, etdBranco
            Obrig.CriaObrigacao Imposto.BuscaCodImposto(NomeImposto), Left(txtPeriodo, 2) & Right(txtPeriodo, 4), Right(txtPeriodo, 4) & Left(txtPeriodo, 2), txtIM, CDbl(txtTotalImposto) - CDbl(Nvl(txtJuros, 0)) - CDbl(Nvl(txtMulta, 0)), etsCreditoNaoLancado, etsCriaNova, txtDtVenc, , , , , , , , 0
            txtTotalImposto = CDbl(txtImposto) + CDbl(txtMulta) + CDbl(txtJuros)
        Else
            'DAM Normal
            Obrig.CriaObrigacao Imposto.BuscaCodImposto(NomeImposto), Left(txtPeriodo, 2) & Right(txtPeriodo, 4), Right(txtPeriodo, 4) & Left(txtPeriodo, 2), txtIM, CDbl(txtTotalImposto) - CDbl(Nvl(txtJuros, 0)) - CDbl(Nvl(txtMulta, 0)), etsCreditoNaoLancado, etsCriaNova, txtDtVenc, , , , , , , , 0
            txtTotalImposto = CDbl(txtImposto) + CDbl(txtMulta) + CDbl(txtJuros) + CDbl(txtDAM)
        End If
    End If
    Bdados.DeletaDados "TAB_DETALHE_DAM", "tdd_tgt_cod_pagamento = " & DAMOriginal
    'Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "tgt_cod_pagamento = " & DAMOriginal
    Valores = Bdados.PreparaValor(CodPagamento, Nvl(txtNotaInicial, 0), Nvl(txtNotaFinal, 0), Bdados.Converte(Nvl(txtTotalNotas, 0), TCDuplo), Bdados.Converte(Nvl(txtMaterial, 0), TCDuplo), txtObservacao.Text, txtAliquota)
    Campos = "tdd_tgt_cod_pagamento,tdd_num_nota_inicial,tdd_num_nota_final,tdd_total_nota,tdd_total_material_reducao,tdd_obs,tdd_Aliquota"
    Bdados.GravaDados "TAB_DETALHE_DAM", Valores, Campos, "tdd_tgt_cod_pagamento=" & CodPagamento

    Valor_Imposto = txtTotalImposto
    Juro = txtJuros
    Multa = txtMulta
    TotalImposto = txtTotalImposto
    BaseDeCalculo = txtSaldo
End Sub

Sub GeraItbi()
        Dim Campos As String
        Dim Valores As String
        Dim Sql As String
        Dim Rs As VSRecordset
        Dim Novo As Boolean
        CodPagamento = Conta.NumPagamento(txtIM, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), CodImposto, cboIC.Text, Nvl(txtParcela, 0), Novo, 0)
        Sql = "select tim_valor from tab_imovel where tim_ic='" & cboIC & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            BaseDeCalculo = Rs(0)
            txtTotalImposto = Rs(0) * Aliquota
        End If
        TotalImposto = txtTotalImposto
        txtImposto = TotalImposto
        Valor_Imposto = txtTotalImposto
        Bdados.FechaTabela Rs
        Conta.GeraPagamento txtIM, cboIC, CodImposto, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), txtDtVenc, txtTotalImposto, 0, 0, CodPagamento, 0, Nvl(txtParcela, 0), TaxaServico, CodTaxa
End Sub

Sub GeraParcelamento()
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim sQL2 As String


    Sql = "Select tgt_cod_pagamento,tgt_periodo,tgt_data_vencimento from tab_geracao_tributo,tab_parcelamento where " & _
                " TPA_TCI_IM ='" & txtIM & "' and TPA_TIP_COD_IMPOSTO= '" & CodImposto & _
                "' and tpa_periodo=" & IIf(Len(txtPeriodo) = 4, txtPeriodo, IIf(Len(txtPeriodo) = 4, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2))) & _
                " and TPA_NUM_PARCELAMENTO = tgt_tpa_num_parcelamento and tgt_parcela=" & txtParcela
    If Trim(cboIC) <> "" Then
        Sql = Sql & " and TPA_TIM_IC='" & cboIC & "'"
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        CodPagamento = Rs(0)
        Data_Vencimento = IIf(CDbl(Format(Rs(2), "yyyymmdd")) < CDbl(Format(Date, "yyyymmdd")), Format(Date, "dd/mm/yyyy"), Format(Rs(2), "dd/mm/yyyy"))
        txtDtVenc = Data_Vencimento
    Else
        Informa "Não existe parcelamento para este contribuinte neste período."
        txtIM.SetFocus
        Bdados.FechaTabela Rs
        Screen.MousePointer = 0
        Exit Sub
    End If
    Bdados.FechaTabela Rs
    Sql = "Select tgt_valor_tributo,tgt_Valor_multa,tgt_valor_juros from tab_geracao_tributo where tgt_cod_pagamento =" & CodPagamento
    If Trim(cboIC) <> "" Then
        Sql = Sql & " and tgt_tim_ic='" & cboIC & "' and tgt_parcela=" & txtParcela

        sQL2 = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & cboIC & _
                "' and (tdi_tco_cod_componente=110 or tdi_tco_cod_componente=108)"
            If Bdados.AbreTabela(sQL2, Rs) Then
                Rs.MoveFirst
                Do While Not Rs.EOF
                    If Rs(0) = 110 Then
                        AreaTotal = Rs(1)
                    ElseIf Rs(0) = 105 Then
                        AreaConstruida = Rs(1)
                    End If
                    Rs.MoveNext
                Loop
            End If
            Bdados.FechaTabela Rs
            sQL2 = "select tvl_valor from tab_valor_terreno where tvl_tlg_cod_logradouro=(" & _
                " select tim_tlg_cod_logradouro from tab_imovel where tim_ic='" & cboIC & "')"
            If Bdados.AbreTabela(sQL2, Rs) Then
                ValorMetro = Format(Rs(0), Const_Monetario)
            End If

            Reducao = Imposto.BuscaReducao(CodImposto, Right(txtPeriodo, 4))
            sQL2 = "select tim_valor from tab_imovel where tim_ic='" & cboIC & "'"
            If Bdados.AbreTabela(sQL2, Rs) Then
                BaseDeCalculo = Rs(0) - (Rs(0) * Reducao)
            End If
            Bdados.FechaTabela Rs
            sQL2 = "select tgt_taxa_expediente from Tab_Geracao_Tributo where tgt_inscricao='" & txtIM & "' and tgt_tip_cod_imposto='" & CodImposto & "' and tgt_periodo=" & txtPeriodo
            If Bdados.AbreTabela(sQL2, Rs) Then
                TaxaParcela = Rs!tgt_taxa_expediente
            End If
            Bdados.FechaTabela Rs
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        txtMulta = Rs!TGT_VALOR_MULTA
        txtJuros = Rs!tgt_valor_juros
        Valor_Imposto = Rs!TGT_VALOR_TRIBUTO
        txtTotalImposto = CDbl(Valor_Imposto) + CDbl(txtMulta) + CDbl(txtJuros)
        Juro = txtJuros
        Multa = txtMulta
        TotalImposto = txtTotalImposto
        txtImposto = Valor_Imposto
        Exercicio = txtPeriodo
    End If
    Bdados.FechaTabela Rs
End Sub

Sub GeraImpostoQualquer()
    Dim Novo As Boolean
    Dim cLSImposto As VSImposto

    Set cLSImposto = New VSImposto
    Incidencia = cLSImposto.BuscaNumeroIncidencia(txtIM, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), CodImposto)
    CodPagamento = Conta.NumPagamento(txtIM, txtPeriodo, CodImposto, cboIC.Text, Nvl(txtParcela, 0), Novo, Incidencia)
    Valor_Imposto = txtTotalImposto
    txtImposto = txtTotalImposto
    Conta.GeraPagamento IIf(Trim(txtIM) = "", Const_ImAvulso, txtIM), cboIC, CodImposto, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), txtDtVenc, txtTotalImposto, 0, 0, CodPagamento, 0, Nvl(txtParcela, 0), TaxaServico, CodTaxa
    txtTotalImposto = CDbl(Nvl(txtTotalImposto, 0)) + CDbl(Nvl(txtMulta, 0)) + CDbl(Nvl(txtJuros, 0))
    TotalImposto = CDbl(Nvl(txtTotalImposto, 0))
    Juro = Nvl(txtJuros, 0)
    Multa = Nvl(txtMulta, 0)
End Sub

Private Function GeraParcelamentoIptu(Cidade As CidadeIptu, Ic As String, Periodo As Integer, Tipo As TipoGeracaoImposto) As Boolean
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Pagamentos As String
    Sql = "Select tgt_cod_pagamento from tab_geracao_tributo where tgt_tim_ic ='" & Ic & "' and tgt_parcela > 0 order by tgt_cod_pagamento "
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do
            Pagamentos = Pagamentos & Rs(0) & "; "
            Rs.MoveNext
        Loop While Not Rs.EOF
        Informa "Parcelamento já existente. Pagamentos nº: " & Mid(Pagamentos, 1, Len(Pagamentos) - 2) & "."
        Screen.MousePointer = 0
        GeraParcelamentoIptu = False
        Bdados.FechaTabela Rs
        Exit Function
    End If

    Sql = "SELECT * FROM tab_imovel,Tab_Contribuinte where tim_ic ='" & Mid(Ic, 1, 12) & "' and tim_tci_im=tci_im "
    If Bdados.AbreTabela(Sql, Rs, Dinamico) Then
        Call Imposto.GeraIptu(cip_Balsas, Rs, Periodo, Periodo, Tipo)
        GeraParcelamentoIptu = True
    End If
    Bdados.FechaTabela Rs
End Function

Sub GeraTaxas(Valor As Double)
        Dim Campos As String
        Dim Valores As String
        Dim Novo As Boolean
        Dim Conta As New ContaCorrente
        txtTotalImposto = Valor
        Valor_Imposto = Valor
        TotalImposto = Valor
        txtImposto = Valor
        BaseDeCalculo = ""
        CodPagamento = Conta.GeraCodPagamento(CodImposto)
        Conta.GeraPagamento IIf(Trim(txtIM) = "", Const_ImAvulso, txtIM), cboIC, CodImposto, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), txtDtVenc, txtTotalImposto, 0, 0, CodPagamento, 0, 0, TaxaServico, CodTaxa
        txtJuros = Format(Conta.CalculaValoresJurosAvulsos(CodImposto, IIf(Len(txtPeriodo) = 4, txtPeriodo, CLng(Right(txtPeriodo, 4) & Left(txtPeriodo, 2))), Nvl(cboConta, 1), Format(Date, "dd/mm/yyyy"), txtDtVenc, txtImposto), Const_Monetario)
        txtMulta = Format(Conta.CalculaValoresMultaAvulsos(CodImposto, IIf(Len(txtPeriodo) = 4, txtPeriodo, CLng(Right(txtPeriodo, 4) & Left(txtPeriodo, 2))), Nvl(cboConta, 1), Format(Date, "dd/mm/yyyy"), txtDtVenc, txtImposto), Const_Monetario)
End Sub

Sub ImprimeBoletoIptu()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Venc As String
    Dim Cobranca As New VSCobranca
    Dim DtGeracao As String
    Sql = "Select tgt_cod_pagamento,tgt_data_vencimento,tgt_valor_tributo,tgt_tim_ic,tgt_data_geracao,tgt_taxa_expediente," & _
    " tgt_periodo,tgt_tim_ic,tgt_parcela  from tab_geracao_tributo where tgt_tim_ic='" & cboIC & _
    "' and tgt_tip_cod_imposto='" & CodImposto & "' and tgt_periodo=" & txtPeriodo & " and tgt_parcela " & IIf(cboConta = "1", " =0", "> 0")
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do
            cboIC_LostFocus
            CodPagamento = Rs!tgt_cod_pagamento
            Venc = Rs!TGT_DATA_VENCIMENTO
            txtImposto = Nvl("" & Rs!TGT_VALOR_TRIBUTO, 0)
            Juro = Conta.CalculaValoresJurosAvulsos(CodImposto, Rs!TGT_PERIODO, IIf(CInt(Nvl(txtParcela, 0)) = 0, EtcCreditoTributario, EtcParcelamento), Format(Date, "dd/mm/yyyy"), Venc, Nvl("" & Rs!TGT_VALOR_TRIBUTO, 0))
            Multa = Conta.CalculaValoresMultaAvulsos(CodImposto, Rs!TGT_PERIODO, IIf(CInt(Nvl(txtParcela, 0)) = 0, EtcCreditoTributario, EtcParcelamento), Format(Date, "dd/mm/yyyy"), Venc, Nvl("" & Rs!TGT_VALOR_TRIBUTO, 0))
            'cboIC = "" & Rs!tgt_tim_ic
            DtGeracao = "" & Rs!tgt_data_geracao
            TaxaServico = Rs!tgt_taxa_expediente
            txtImposto = Edita.FormataTexto(txtImposto, Monetario, True)
            txtJuros = Format(Juro, Const_Monetario)
            txtMulta = Format(Multa, Const_Monetario)
            TaxaServico = TaxaServico + TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
            DoEvents
            Cobranca.ImprimeDam Rpt, CodPagamento, txtIM, txtContribuinte, "", txtEnderecoContrib, cboIC, txtImovel, _
                CodImposto, Trim(Left(NomeImposto, PosTraco - 2)), Trim(Mid(NomeImposto, PosTraco + 2)), txtPeriodo, Nvl("" & Rs!TGT_PARCELA, 0), IIf(txtParcela = "0", 1, 3), Venc, BaseDeCalculo, txtImposto, _
                txtMulta, txtJuros, TaxaServico, Desconto, Cod_Atividade, txtObservacao, PicBarra, txtNotaInicial, txtNotaFinal, txtMaterial, ValorMetro, _
                TaxaParcela, AreaTotal, AreaConstruida, ValorTerreno, Valoredific, Zona, , , String_Taxas
            Rs.MoveNext
        Loop While Not Rs.EOF
'        Informa "DAM Emitido e Gravaçao efetuada com sucesso."
    End If
    Screen.MousePointer = 0
    Bdados.FechaTabela Rs
End Sub

Private Sub cboConta_Click()
    On Error Resume Next
    If cboConta.Text = "3" Then
        If NomeImposto = Imposto.NomeTributo(ttr_IPTU) Then
            txtParcela = "1"
        Else
            fra(2).Visible = False
            lbl(17).Visible = True
            txtParcela.Visible = True
            txtParcela.Tag = "Parcela"
            txtParcela.SetFocus
        End If
    Else
        'If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_IRPJ) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Then
        If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Then
            fra(2).Visible = True
        Else
            fra(2).Visible = False
        End If
        fra(2).Visible = True
        lbl(17).Visible = False
        txtParcela.Visible = False
        txtParcela.Tag = ""
        txtParcela = ""
        'txtPeriodo.SetFocus
    End If
End Sub

Sub CarregaImovel(InscIc As String)
    Dim Sql As String
    Dim Rs As VSRecordset

    Sql = "select ttl_nome,tlg_nome,tba_nome,tim_numero,TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC,Tim_Zona," & _
    " tim_valor,tim_tci_im ,TIM_SITUACAO_LOTE   from TAB_IMOVEL,TAB_BAIRRO," & _
    " TAB_LOGRADOURO,TAB_TIPO_LOGR " & _
    " where tim_ic='" & InscIc & _
    "' AND tim_tlg_cod_logradouro = " & _
    " TAB_LOGRADOURO.tlg_cod_logradouro AND tlg_ttl_cod_tip_logr = ttl_cod_tip_logr AND " & _
    " tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio & " AND TBA_TMU_COD_MUNICIPIO =" & _
    Aplicacoes.Codigo_Municipio & "  AND TIM_TBA_COD_BAIRRO = TBA_COD_BAIRRO"
    If Trim(cboIC) = "" Then Exit Sub
    If Bdados.AbreTabela(Sql, Rs) Then
        If "" & Rs!TIM_SITUACAO_LOTE = 1 Then
            Informa "Imóvel desativado."
            cboIC.SetFocus
            Exit Sub
        End If
        txtImovel = Rs(0) & " " & Rs(1) & " " & Rs(2) & " " & Rs(3)
        ValorTerreno = Nvl("" & Rs!TIM_VALOR_TERRENO, 0)
        Valoredific = Nvl("" & Rs!TIM_VALOR_EDIFIC, 0)
        Zona = Nvl("" & Rs!tim_zona, Mid(cboIC, 4, 1))
        BaseDeCalculo = "" & Rs!tim_valor
        txtIM = Nvl("" & Rs!tim_tci_im, 0)
        txtIm_LostFocus
    Else
        Informa "Imovel não cadastrado."
    End If
    Bdados.FechaTabela Rs
End Sub
Private Sub cboIC_LostFocus()
    CarregaImovel cboIC
End Sub


Private Sub cboImposto_LostFocus()
'    Dim Sql As String
'    Dim rs As VSRecordset
'    If Trim(cboImposto) = "" Then Exit Sub
'    Dim i As Byte
'
'    If blnConsultaIM Then Exit Sub
'
'    txtIM.Enabled = True
'    txtCgc.Enabled = True
'    txtContribuinte = ""
'    txtDtVenc = ""
'    txtPeriodo = ""
'    txtFator.Visible = False
'    If Trim(cboImposto) <> "" Then
'        i = InStr(1, cboImposto, "#")
'        NomeImposto = Left(cboImposto.Text, i - 2)
'    End If
'    txtIM.SetFocus
'    txtMaterial.Enabled = True
'    If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNEST) Then
'        fra(2).Visible = True
'        txtMaterial.Enabled = IIf(NomeImposto = Imposto.NomeTributo(ttr_IRPJ), False, True)
'        cboTipo.SetarLinha 2, 1
'        cboTipo_Click
'        txtPeriodo.MaxLength = 6
'    ElseIf NomeImposto = Imposto.NomeTributo(ttr_IPTU) Or NomeImposto = Imposto.NomeTributo(ttr_ITBI) Or NomeImposto = Imposto.NomeTributo(ttr_2VIA) Or NomeImposto = Imposto.NomeTributo(ttr_AFORO) Or NomeImposto = Imposto.NomeTributo(ttr_VISTORIA) Or NomeImposto = Imposto.NomeTributo(ttr_DESMEMBRAMENTO) Or NomeImposto = Imposto.NomeTributo(ttr_CONSTRUCAO) Or NomeImposto = Imposto.NomeTributo(ttr_HABIT) Or NomeImposto = Imposto.NomeTributo(ttr_RECONSTRUCAO) Or NomeImposto = Imposto.NomeTributo(ttr_REMEMBRA) Then
'        fra(2).Visible = False
'
'        txtPeriodo.MaxLength = 4
'    ElseIf NomeImposto = Imposto.NomeTributo(ttr_ALVARA) Then
'        fra(2).Visible = False
'
'        txtPeriodo.MaxLength = 4
'    End If
    NomeImposto = Trim(ParseString(Trim(cboImposto.Text), "#", 2))
    CodImposto = BuscaCodigo("SELECT TIP_COD_IMPOSTO FROM TAB_IMPOSTO WHERE TIP_NOME_IMPOSTO = '" & NomeImposto & "'")
'    'Ver se é Taxa
'    Sql = "Select tpi_tipo_tributo,tpi_valor_taxa_fixa,tpi_tipo_inscricao,tpi_tipo_ic from tab_parametro_imposto where tpi_tip_cod_imposto ='" & CodImposto & "'"
'    If Bdados.AbreTabela(Sql, rs) Then
'           TributoTaxa = IIf(rs!tpi_tipo_tributo = 2, True, False)
'           If TributoTaxa Then
'                lbl(16).Visible = False
'                cboConta.Visible = False
'                cboConta.Tag = ""
'           Else
'                lbl(16).Visible = True
'                cboConta.Visible = True
'                cboConta.Tag = "Conta"
'
'                txtFator.Visible = False
'                lbl(21).Visible = False
'
'                lbl(22).Visible = False
'                txtFatorAleatorio.Visible = False
'           End If
'           If Nvl("" & rs!tpi_tipo_ic, 0) = 1 Then
'                fraIm.Visible = True
'           Else
'                fraIm.Visible = False
'           End If
'           If TributoTaxa Then
'                txtContribuinte.Enabled = True
'                fra(2).Visible = False
'                txtPeriodo.MaxLength = 7
'                txtPeriodo = Format(Month(Date), "00") & Format(Year(Date), "0000")
'                txtPeriodo_LostFocus
'                txtPeriodo.MaxLength = 6
'                Exercicio = txtPeriodo
'                cboConta.Tag = ""
'                txtIM = ""
'                txtCgc = ""
'                txtFator.Visible = True
'                lbl(21).Visible = True
'                If TributoTaxaFixa = 0 Then
'                    lbl(22).Visible = True
'                    txtFatorAleatorio.Visible = True
'                Else
'                    lbl(22).Visible = False
'                    txtFatorAleatorio.Visible = False
'                End If
'
'           ElseIf rs!tpi_tipo_inscricao = 2 Then
'                txtContribuinte.Enabled = False
'                txtIM.Enabled = True
'                txtCgc.Enabled = True
'                txtFator.Visible = False
'                lbl(21).Visible = False
'
'                lbl(22).Visible = False
'                txtFatorAleatorio.Visible = False
'           End If
'    Else
'            lbl(21).Visible = False
'            txtFator.Visible = False
'            lbl(22).Visible = False
'            txtFatorAleatorio.Visible = False
'
'    End If
'    'Se for publicidade monto acombo com os items da publiciade
'    If NomeImposto = "PUBL" Then
'        If Bdados.AbreTabela("SELECT TPD_TIP_COD_IMPOSTO,tpd_descricao ,tpd_item  From TAB_PARAMETRO_DETALHE where  TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(CodImposto, tctexto)) Then
'            CboItem.Visible = True
'            lbl(23).Visible = True
'            CboItem.Preencher Bdados, "SELECT TPD_TIP_COD_IMPOSTO,tpd_descricao ,tpd_item  From TAB_PARAMETRO_DETALHE where  TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(CodImposto, tctexto), 1
'        Else
'            CboItem.Visible = False
'            lbl(23).Visible = False
'        End If
'    Else
'        CboItem.Visible = False
'        lbl(23).Visible = False
'    End If
'    Bdados.FechaTabela rs
    
End Sub

Private Sub cboTipo_Click()
    If cboTipo.Coluna(1).Valor = 1 Then
        txtTotalImposto = "0,00"
        txtDAM = "0,00"
        txtJuros = "0,00"
        txtMulta = "0,00"
        txtImposto = "0,00"
    Else
        txtTotalNotas_Change
        txtDAM = TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim a As Integer
    Dim Valores As String
    Dim Campos As String
    Dim ValorImposto As Double
    Dim RsCob As VSRecordset
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim SqlParc As String
    Dim Cobranca As New VSCobranca


    Select Case cmd(Index).Caption

        Case "&Emitir DAM"
            If Not Util.Confirma("Confirma a emissão do DAM?") Then
                Exit Sub
            End If
            a = InStr(cboImposto, " # ")
            NomeImposto = Trim(Left(cboImposto.Text, a - 1))
            txtContribuinte.Enabled = True
            Documento = CodPagamento
            Data_Vencimento = txtDtVenc
            Cod_Tributo = NomeImposto
            CPFCNPJ = txtCgc
            InscMuni = txtIM
            RazaoSocial = txtContribuinte
            Juro = 0
            Multa = 0
            TotalImposto = 0
            Linhas = 0
            Screen.MousePointer = 11

            PosTraco = InStr(1, cboImposto, "#")

            If Not Edita.CriticaCampos(Me) Then
                txtContribuinte.Enabled = True
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If Not Left(cboImposto, 4) = "PUBL" Then
                If Trim(cboIC) = "" And cboIC.Visible = True Then
                    Informa "Informe a inscrição cadastral."
                    cboIC.SetFocus
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If

            If TributoTaxa Then 'Taxas Diversas
                'Para não modificar o código, pego somente as taxas de publicidades...
                If Left(cboImposto.Text, 4) = "PUBL" Then
                   'procedimento que vai me retornar o codigo da taxa atraves do nome da mesma...
                    Dim Pos
                    Dim Codigo_I As String
                    Dim I_P As String
                    Pos = InStr(cboImposto, "#") + 1
                    Codigo_I = Right(cboImposto, Len(cboImposto) - Pos)
                    BuscaValorTaxa Codigo_I, I_P
                    'Checo se existem taxas para esse contribuinte...
                    Sql = "Select * from tab_Anuncio where TAN_TCI_IM = " & Bdados.Converte(txtIM, tctexto) & " and  TAN_MOVIMENTO = " & Bdados.Converte(I_P, tctexto)
                    If Bdados.AbreTabela(Sql, Rs) Then
                        If CboItem.Visible Then
                            Do Until Rs.EOF
                                If Trim(Rs.Fields("TAN_MOVIMENTO")) = Trim(I_P) And Trim(Rs.Fields("TAN_TPD_ITEM")) = Trim(CStr(CboItem.Coluna(2).Valor)) Then
                                    GeraTaxas Rs.Fields("TAN_VALOR_APAGAR")
                                    Exit Do
                                End If
                                Rs.MoveNext
                            Loop
                        Else
                            Do Until Rs.EOF
                                If Trim(Rs.Fields("TAN_MOVIMENTO")) = Trim(I_P) Then
                                    GeraTaxas Rs.Fields("TAN_VALOR_APAGAR")
                                    Exit Do
                                End If
                                Rs.MoveNext
                            Loop
                        End If
                    Else
                        Util.Avisa "Não existem anuncios para este contribuinte."
                          Screen.MousePointer = 0
                        Exit Sub
                    End If

                Else
                    If TributoTaxaFixa > 0 Then
                        GeraTaxas IIf(txtFator.Visible = True, CDbl(Nvl(CStr(TributoTaxaFixa), Nvl(txtFatorAleatorio, 1))) * CDbl(Nvl(txtFator, 1)), TributoTaxaFixa)
                    Else
                        GeraTaxas CDbl(Nvl(txtFator, 0)) * CDbl(Nvl(txtFatorAleatorio, 0))
                    End If
                End If
            Else 'Imposto Comum
                If cboConta.Text <> "3" Then 'Credito Tributario / Auto Infracao
                    Sql = "Select tgt_valor_tributo,tgt_data_geracao,tgt_cod_pagamento from tab_geracao_tributo where tgt_tip_cod_imposto='" & CodImposto & "' and tgt_periodo=" & IIf(Len(txtPeriodo) = 4, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2)) & " and tgt_parcela =" & Nvl(txtParcela, 0)
                    If Trim(cboIC) <> "" Then
                        Sql = Sql & " AND tgt_inscricao ='" & cboIC & "'"
                    Else
                        Sql = Sql & " AND tgt_inscricao='" & txtIM & "'"
                    End If
                    If Bdados.AbreTabela(Sql, Rs) Then
                        If NomeImposto <> Imposto.NomeTributo(ttr_ISSQNRET) And NomeImposto <> Imposto.NomeTributo(ttr_ISSQNSUBST) And NomeImposto <> Imposto.NomeTributo(ttr_ISSQNCOMP) And NomeImposto <> Imposto.NomeTributo(ttr_ISSQNEST) Then
                            If Not Confirma("DAM já emitido no valor de R$ " & Format(Nvl(Rs(0), 0), Const_Monetario) & ", em " & IIf(Not IsNull(Rs(1)), Rs(1), "") & ". NÚMERO DO PAGAMENTO: " & Rs!tgt_cod_pagamento & ". Deseja continuar?") Then
                                Screen.MousePointer = 0
                                Bdados.FechaTabela Rs
                                Exit Sub
                            End If
                        End If
                        Bdados.FechaTabela Rs
                    End If
                    If NomeImposto = Imposto.NomeTributo(ttr_IPTU) Then
                        GeraIptu
                        If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
                            ImprimeBoletoIptu
                            Exit Sub
                        End If
                    ElseIf NomeImposto = Imposto.NomeTributo(ttr_ALVARA) Then
                        GeraAlvara
                    ElseIf NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNEST) Then
                        GeraIssqn
                    ElseIf NomeImposto = Imposto.NomeTributo(ttr_ITBI) Then
                        GeraItbi
                    Else
                        GeraImpostoQualquer
                    End If
                Else ' PARCELAMENTO DE DIVIDA
                    If NomeImposto = Imposto.NomeTributo(ttr_IPTU) Then
                        If GeraParcelamentoIptu(cip_Balsas, cboIC, txtPeriodo, tgi_SoParcelas) Then
                            'GeraIptu
                            If Temp.PegaParametro(Bdados, "TIPO IPTU") = 1 Then
                                ImprimeBoletoIptu
                                Exit Sub
                            End If
                        Else
                            Exit Sub
                        End If
                    Else
                        GeraParcelamento
                    End If
                End If
            End If
            'Pego as taxas
            Call Pega_taxas
            TaxaServico = TaxaServico + Total_Taxas
            If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNEST) Then
                If cboTipo.Coluna(1).Valor = 2 Then
'                    If TotalImposto > 0 Then
                        PosTraco = InStr(1, cboImposto, "#")
                        Cobranca.ImprimeDam Rpt, CodPagamento, txtIM, txtContribuinte, "" & txtCgc, txtEnderecoContrib, cboIC, txtImovel, _
                        CodImposto, Trim(Left(cboImposto, PosTraco - 2)), Trim(Mid(cboImposto, PosTraco + 2)), TiraTudo(txtPeriodo), txtParcela, Nvl(cboConta, 1), txtDtVenc, BaseDeCalculo, 0, _
                        Multa, Juro, CStr(Temp.PegaParametro(Bdados, "TXDAM")), "", Cod_Atividade, txtObservacao, PicBarra, txtNotaInicial, txtNotaFinal, txtMaterial, ValorMetro, _
                        TaxaParcela, AreaTotal, AreaConstruida, , , , , etdBranco, String_Taxas
                        Desconto = 8
'                    Else8
'                        Informa "Total do imposto igual zero. DAM não será emitido."
'                        Bdados.DeletaDados "tab_geracao_tributo", "tgt_cod_pagamento=" & CodPagamento
'                        Bdados.DeletaDados "tab_conta_contribuinte", "tcc_codigo_conta=" & CodPagamento
'                    End If
                Else
                        PosTraco = InStr(1, cboImposto, "#")
                        Cobranca.ImprimeDam Rpt, CodPagamento, txtIM, txtContribuinte, "" & txtCgc, txtEnderecoContrib, cboIC, txtImovel, _
                        Imposto.BuscaCodImposto(Trim(Left(cboImposto, PosTraco - 2))), Trim(Left(cboImposto, PosTraco - 2)), Trim(Mid(cboImposto, PosTraco + 2)), txtPeriodo, txtParcela, Nvl(cboConta, 1), txtDtVenc, BaseDeCalculo, txtImposto, _
                        Multa, Juro, TaxaServico, "", Cod_Atividade, txtObservacao, PicBarra, txtNotaInicial, txtNotaFinal, txtMaterial, ValorMetro, _
                        TaxaParcela, AreaTotal, AreaConstruida, , , , , etdBranco, String_Taxas, txtDtVenc
                        Desconto = 0
                End If
            Else
                If TotalImposto > 0 Then
                    PosTraco = InStr(1, cboImposto, "#")
                    Cobranca.ImprimeDam Rpt, CodPagamento, txtIM, txtContribuinte, "" & txtCgc, txtEnderecoContrib, cboIC, txtImovel, _
                    CodImposto, Trim(Left(cboImposto, PosTraco - 2)), Trim(Mid(cboImposto, PosTraco + 2)), txtPeriodo, txtParcela, Nvl(cboConta, 1), txtDtVenc, BaseDeCalculo, txtImposto, _
                     Multa, Juro, TaxaServico, "", Cod_Atividade, txtObservacao, PicBarra, txtNotaInicial, txtNotaFinal, txtMaterial, ValorMetro, _
                    TaxaParcela, AreaTotal, AreaConstruida, , , , , , String_Taxas
                    Desconto = 0
    '                Informa "DAM Emitido e Gravaçao efetuada com sucesso."
                Else
                    Informa "Total do imposto igual zero. DAM não será emitido."
                    Bdados.DeletaDados "tab_geracao_tributo", "tgt_cod_pagamento=" & CodPagamento
                    Bdados.DeletaDados "tab_conta_contribuinte", "tcc_codigo_conta=" & CodPagamento
                End If
            End If
            'Call cmdLimpar_Click
        Case "Sai&r"
           Unload Me
    End Select
    Screen.MousePointer = 0
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    cboTaxa.ListIndex = 0
    txtDAM = TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
    txtDAM.Enabled = False
    cboImposto.SetFocus
End Sub

Private Sub cmdPesq_Click(Index As Integer)
    blnConsultaIM = True
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIM
    blnConsultaIM = False
End Sub

Private Sub Form_Load()

    Dim Controle As Control
    Dim i As Byte
    Call Edita.AtualizaCombo(Bdados, cboImposto, "Select TIP_sigla_IMPOSTO   " & _
        Bdados.Concatena & " ' # '" & Bdados.Concatena & _
        "  tip_nome_imposto  from tab_imposto where tip_sigla_imposto like 'ISS%'")
    GrdTaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
    CodTaxaEmissaoDam = Temp.PegaParametro(Bdados, "TAXA EMISSAO DAM")
    If CodTaxaEmissaoDam <> "" Then
        Call Edita.AtualizaCombo(Bdados, cboTaxa, "select TIP_NOME_IMPOSTO from tab_imposto" & _
        " WHERE  tip_cod_imposto = '" & CodTaxaEmissaoDam & "'")
        cboTaxa.ListIndex = 0
    Else
        Call Edita.AtualizaCombo(Bdados, cboTaxa, "select distinct(TIP_NOME_IMPOSTO) from tab_imposto" & _
        " WHERE  tip_cod_imposto in (SELECT tpi_tip_cod_imposto FROM Tab_Parametro_Imposto where tpi_tipo_tributo = 2 and tpi_valor_taxa_fixa > 0)")
        cboTaxa.AddItem " "
    End If
    cboTipo.PreencherGeral Bdados, "TIPO DAM"
    cboTipo.Enabled = True
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    txtDAM = TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
    txtDAM.Enabled = False
    cboConta.ListIndex = 0
    cboTipo.ListIndex = 0
    cboTipo_Click
    cboTipo.Enabled = False
    txtNotaInicial.Enabled = False
    txtNotaFinal.Enabled = False
    txtAliquota.Enabled = False
    txtTotalImposto.Enabled = False
    txtTotalNotas.Enabled = False
    txtMaterial.Enabled = False
        
End Sub

Private Sub txtAliquota_Change()
    If txtAliquota = "" Then
        txtAliquota = 3
    End If
    txtTotalNotas_Change
End Sub

Private Sub txtcgc_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtcgc_LostFocus()

    Dim Sql As String
    Dim Rs As VSRecordset


    If Trim(txtCgc) = "" Then Exit Sub
    If Len(txtCgc) <> 18 Then
        If Len(txtCgc) = 11 Then
            txtCgc = Edita.FormataTexto(txtCgc, Cpf)
        ElseIf Len(txtCgc) = 14 And Mid(txtCgc, 4, 1) <> "." Then
            txtCgc = Edita.FormataTexto(txtCgc, Cgc)
        End If
    End If
    Sql = "select tci_nome,tci_im,tci_logradouro,tci_nome_logradouro," & _
    " tci_numero,tci_complemento,tci_bairro,tci_cidade,tci_UF from Tab_Contribuinte" & _
    " where tci_cgc_cpf='" & txtCgc & "' and tci_tsc_cod_sit_cad =1"
    If Bdados.AbreTabela(Sql, Rs) Then
        txtContribuinte = Rs(0)
        txtIM = Rs!TCI_IM
        Endereco = Rs!tci_logradouro & " " & Rs!tci_nome_logradouro & " " & Rs!tci_NUMERO & " " & Rs!tci_COMPLEMENTO
        Bairro = Rs!tci_BAIRRO
        Cod_Cidade = Rs!tci_cidade
        Uf = Rs!tci_UF
        Call txtIm_LostFocus

        If NomeImposto = Imposto.NomeTributo(ttr_IPTU) Or NomeImposto = Imposto.NomeTributo(ttr_ITBI) Then
            cboIC.Enabled = True
            cboIC.Visible = True
        End If
    ElseIf Not TributoTaxa Then
        Util.Informa "Cgc/Cpf não cadastrado."
        txtCgc.SetFocus
    Else
        Cod_Atividade = ""
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtContribuinte_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtFator_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtFator_LostFocus()
    If Trim(txtFator) = "" Then Exit Sub
    If CDbl(txtFator) <= 0 Then txtFator = "1"
End Sub

Private Sub txtFatorAleatorio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    On Error GoTo trata
    'txtIM = Edita.TiraTudo(txtIM)
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        txtIM = Imposto.FormataInscricao(txtIM, InscContrib)
    End If
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsIptu As VSRecordset
    If Trim(txtIM) = "" Then Exit Sub
    Sql = "select * from Tab_Contribuinte where tci_im='" & txtIM & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        txtContribuinte = Rs!tci_nome
        txtEnderecoContrib = Rs!tci_logradouro & " " & Rs!tci_nome_logradouro & " " & Rs!tci_NUMERO & " " & Rs!tci_BAIRRO & ", CEP: " & Rs!tci_cep & ", " & Rs!tci_cidade & "-" & Rs!tci_UF
        If (NomeImposto = Imposto.NomeTributo(ttr_IPTU) Or NomeImposto = Imposto.NomeTributo(ttr_ITBI)) And Trim(txtImovel) = "" Then
            Sql = "select tim_ic,TIM_UNIDADE from tab_imovel where tim_tci_im='" & txtIM & "'"
            If Bdados.AbreTabela(Sql, RsIptu) Then
                RsIptu.MoveFirst
                cboIC.Clear
                Do While Not RsIptu.EOF
                    cboIC.AddItem RsIptu(0) & Format(RsIptu!TIM_UNIDADE, "000")
                    RsIptu.MoveNext
                Loop
            End If
            'cboIC.SetFocus
        End If
        If IsNull(Rs("tci_tae_cae")) And (NomeImposto <> Imposto.NomeTributo(ttr_IPTU) And NomeImposto <> Imposto.NomeTributo(ttr_ITBI)) Then
            If Not TributoTaxa Then
                Informa "Contribuinte não sujeito à cobrança de " & NomeImposto & "."
                txtIM.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        txtContribuinte = "" & Rs("tci_nome") & ""
        txtCgc = "" & Rs("tci_cgc_cpf") & ""

        Cod_Atividade = IIf(Nvl("" & Rs("tci_tae_cae"), 0) = 0, "", Rs("tci_tae_cae")) & Imposto.BuscaNomeCAE(Rs("tci_tae_cae") & "") & ""

        Endereco = "" & Rs("tci_logradouro") & "  " & Rs("tci_nome_logradouro") & "," & Rs("tci_numero") & " " & Rs("tci_complemento")
        Bairro = "" & Rs("tci_bairro") & ""
        Cod_Cidade = "" & Rs("tci_cidade") & ""
        Cep = "" & Rs("tci_cep") & ""
        Uf = "" & Rs("tci_uf") & ""
        If Not TributoTaxa Then
            'If ((RS!tci_tipo_recolhimento_iss = 1 Or RS!tci_tipo_recolhimento_iss = 3) And NomeImposto = Imposto.NomeTributo(ttr_ISSQN)) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_IRPJ) Then
            If (Rs!tci_tipo_recolhimento_iss = 1 Or Rs!tci_tipo_recolhimento_iss = 3) Then
                txtPeriodo.MaxLength = 6
            ElseIf Rs!tci_tipo_recolhimento_iss = 2 Then
                txtPeriodo.MaxLength = 4
            Else
                If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Then
                    Informa "Contribuinte sem obrigação de recolhimento de ISSQN"
                    txtCgc = ""
                    txtIM = ""
                    txtIM.SetFocus
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        Else
            txtPeriodo.MaxLength = 6
        End If
    Else
        Util.Informa "Inscrição não cadastrada."
        txtIM = ""
        Cod_Atividade = ""
        txtIM.SetFocus
    End If
    Bdados.FechaTabela Rs
    Bdados.FechaTabela RsIptu
    txtPeriodo.SetFocus
    txtPeriodo.MaxLength = 0
    Exit Sub
trata:
    If Err.Number = 3265 Then
        Resume Next
        'Util.Erro Err.Description
    Else
        Util.Erro Err.Description
    End If
End Sub

Private Sub txtImposto_Change()
    txtTotalImposto = txtImposto
End Sub

Private Sub txtMaterial_Change()
    On Error Resume Next
    Dim dblTotalNotas As Double
    Dim dblValorMaterial As Double

    If IsNumeric(txtTotalNotas) Then
        dblTotalNotas = txtTotalNotas
    End If
    If IsNumeric(txtMaterial) Then
        dblValorMaterial = txtMaterial
    End If
    txtSaldo = dblTotalNotas - dblValorMaterial

    txtSaldo = Edita.FormataTexto(txtSaldo, Monetario, True)
    txtImposto = (txtAliquota * txtTotalNotas / 100)
    txtImposto = Edita.FormataTexto(txtImposto, Monetario, True)
    If Trim(txtImposto) <> "" Then
        txtJuros = Format(Conta.CalculaValoresJurosAvulsos(CodImposto, IIf(Len(txtPeriodo) = 4, txtPeriodo, CLng(Right(txtPeriodo, 4) & Left(txtPeriodo, 2))), cboConta, Format(Date, "dd/mm/yyyy"), txtDtVenc, txtImposto), Const_Monetario)
        txtMulta = Format(Conta.CalculaValoresMultaAvulsos(CodImposto, IIf(Len(txtPeriodo) = 4, txtPeriodo, CLng(Right(txtPeriodo, 4) & Left(txtPeriodo, 2))), cboConta, Format(Date, "dd/mm/yyyy"), txtDtVenc, txtImposto), Const_Monetario)
        txtTotalImposto = Format(CDbl(Nvl(txtImposto, 0)) + CDbl(Nvl(txtJuros, 0)) + CDbl(Nvl(txtMulta, 0)), Const_Monetario)
    End If
End Sub

Private Sub txtMaterial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtMaterial_LostFocus()
    txtMaterial = Edita.FormataTexto(txtMaterial, Monetario, True)
    If CDbl(Nvl(txtTotalNotas, 0)) < CDbl(Nvl(txtMaterial, 0)) Then
        Informa "Valor não pode ser maior que Total em notas."
        txtMaterial = "0,00"
        txtMaterial.SetFocus
    End If
End Sub

Private Sub txtNotaFinal_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtNotaInicial_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtObservacao_KeyDown(KeyCode As Integer, Shift As Integer)
    Shift = 0
End Sub

Private Sub txtObservacao_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtParcela_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtParcela_LostFocus()
    If Trim(txtParcela) <> "" Then
        If txtParcela < 1 Then
            Informa "Número de cota inválido."
            txtParcela.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPeriodo_LostFocus()
    
    If Len(txtPeriodo) = 6 Then
        txtPeriodo.MaxLength = 7
        txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
        txtPeriodo.MaxLength = 6
    End If
    
    txtDtVenc = Imposto.BuscaDataVencimento(CodImposto, Left(txtPeriodo, 2) & Right(txtPeriodo, 4))
End Sub

Private Sub txtTotalNotas_Change()
    On Error Resume Next
    Dim Conta As New ContaCorrente
    Dim dblTotalNotas As Double
    Dim dblValorMaterial As Double

    If IsNumeric(txtTotalNotas) Then
        dblTotalNotas = txtTotalNotas
    End If
    If IsNumeric(txtMaterial) Then
        dblValorMaterial = txtMaterial
    End If
    txtSaldo = dblTotalNotas + dblValorMaterial
    txtSaldo = Edita.FormataTexto(txtSaldo, Monetario, True)
    If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNEST) Then
        'txtImposto = Aliquota * CDbl(Nvl(txtSaldo, 0))
        txtImposto = (txtAliquota * Val(txtTotalNotas) / 100)
    End If
    txtImposto = Edita.FormataTexto(txtImposto, Monetario, True)
    If Trim(txtImposto) <> "" Then
        txtJuros = Format(Conta.CalculaValoresJurosAvulsos(CodImposto, IIf(Len(txtPeriodo) = 4, txtPeriodo, CLng(Right(txtPeriodo, 4) & Left(txtPeriodo, 2))), cboConta, Format(Date, "DD/MM/YYYY"), txtDtVenc, CDbl(txtImposto)), Const_Monetario)
        txtMulta = Format(Conta.CalculaValoresMultaAvulsos(CodImposto, IIf(Len(txtPeriodo) = 4, txtPeriodo, CLng(Right(txtPeriodo, 4) & Left(txtPeriodo, 2))), cboConta, Format(Date, "DD/MM/YYYY"), txtDtVenc, CDbl(txtImposto)), Const_Monetario)
        txtTotalImposto = Format(CDbl(txtImposto) + CDbl(txtJuros) + CDbl(txtMulta), Const_Monetario)
    End If
End Sub

Private Sub txtTotalNotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtTotalNotas_LostFocus()
    txtTotalNotas = Edita.FormataTexto(txtTotalNotas, Monetario, True)
End Sub
Private Sub Pega_taxas()
    Dim i As Integer
    Dim Pos As Integer
    String_Taxas = ""
    Total_Taxas = 0
    For i = 1 To GrdTaxas.ListItems.Count
        If GrdTaxas.ListItems(i).Checked Then
            Pos = InStr(GrdTaxas.ListItems(i).SubItems(1), "-") - 1
            If String_Taxas = "" Then
                String_Taxas = String_Taxas & " [ " & Left(GrdTaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(GrdTaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            Else
                String_Taxas = String_Taxas & ", [ " & Left(GrdTaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(GrdTaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            End If
            Total_Taxas = Total_Taxas + CCur(GrdTaxas.ListItems(i).SubItems(2))
        End If
    Next
End Sub
