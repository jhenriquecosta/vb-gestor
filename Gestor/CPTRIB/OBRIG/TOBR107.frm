VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TOBR107 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleMode       =   0  'User
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   60
      Top             =   6945
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   5610
         TabIndex        =   28
         Top             =   120
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
         CorFoco         =   16761024
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   4245
         TabIndex        =   27
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Emitir DAM"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
         CorFoco         =   16761024
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   6720
         TabIndex        =   29
         Top             =   120
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
         CorFoco         =   16761024
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
      TabIndex        =   59
      Top             =   4800
      Visible         =   0   'False
      Width           =   795
   End
   Begin Threed.SSFrame fra 
      Height          =   2280
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   4680
      Visible         =   0   'False
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
         TabIndex        =   22
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
         TabIndex        =   17
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
         TabIndex        =   26
         Tag             =   " "
         Top             =   1860
         Width           =   1275
      End
      Begin VB.TextBox txtJuros 
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
         Left            =   6360
         TabIndex        =   25
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txtImposto 
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
         Left            =   6360
         TabIndex        =   23
         Top             =   870
         Width           =   1275
      End
      Begin VB.TextBox txtSaldo 
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
         Left            =   6360
         TabIndex        =   21
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   210
         Width           =   1275
      End
      Begin VB.TextBox txtMulta 
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
         Left            =   6360
         TabIndex        =   24
         Top             =   1200
         Width           =   1275
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   11
         Left            =   1785
         TabIndex        =   38
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
         TabIndex        =   39
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
         TabIndex        =   45
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
         TabIndex        =   31
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
         TabIndex        =   32
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
         TabIndex        =   46
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
         TabIndex        =   47
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
         TabIndex        =   48
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
         TabIndex        =   49
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
         TabIndex        =   30
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
         TabIndex        =   20
         Top             =   1875
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
         Enabled         =   0   'False
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   3285
      Index           =   0
      Left            =   90
      TabIndex        =   33
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
      Begin VTOcx.txtVISUAL txtDtVenc 
         Height          =   315
         Left            =   1425
         TabIndex        =   9
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
         ItemData        =   "TOBR107.frx":0000
         Left            =   2775
         List            =   "TOBR107.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
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
         ItemData        =   "TOBR107.frx":0004
         Left            =   60
         List            =   "TOBR107.frx":0006
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
         TabIndex        =   7
         Top             =   1155
         Visible         =   0   'False
         Width           =   765
      End
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
         ItemData        =   "TOBR107.frx":0008
         Left            =   5940
         List            =   "TOBR107.frx":0015
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Conta"
         Top             =   1155
         Width           =   825
      End
      Begin VB.Frame fraIm 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   705
         Left            =   60
         TabIndex        =   50
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
            TabIndex        =   13
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
            ItemData        =   "TOBR107.frx":0022
            Left            =   30
            List            =   "TOBR107.frx":0024
            Sorted          =   -1  'True
            TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   8
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
         TabIndex        =   34
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
         TabIndex        =   35
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
         TabIndex        =   36
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
         TabIndex        =   41
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
         TabIndex        =   42
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
         TabIndex        =   44
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
         TabIndex        =   51
         Top             =   915
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
         TabIndex        =   52
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
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   58
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
      TabIndex        =   53
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
         TabIndex        =   14
         Top             =   240
         Width           =   6255
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   19
         Left            =   120
         TabIndex        =   54
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
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1230
      Width           =   5895
   End
   Begin VTOcx.grdVISUAL Grdtaxas 
      Height          =   6555
      Left            =   7890
      TabIndex        =   66
      Top             =   675
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   11562
      Caption         =   "Taxas"
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
      Ordenavel       =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   1138
      Icone           =   "TOBR107.frx":0026
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1200
      TabIndex        =   40
      Top             =   180
      Width           =   375
   End
End
Attribute VB_Name = "TOBR107"
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
    Dim rs                                      As VSRecordset
    Dim Sql                                      As String
    Sql = " select tpi_tip_cod_imposto,tpi_valor_taxa_fixa  from Tab_Parametro_Imposto " & _
    " WHERE  tpi_tip_cod_imposto = (SELECT TIP_COD_IMPOSTO FROM TAB_IMPOSTO WHERE TIP_NOME_IMPOSTO ='" & NomeTaxa & "')"
    If Bdados.AbreTabela(Sql, rs) Then
        BuscaValorTaxa = IIf(IsNull(rs!tpi_valor_taxa_fixa), 0, rs!tpi_valor_taxa_fixa)
        CodTributo = rs!tpi_tip_cod_imposto
    Else
        BuscaValorTaxa = 0
    End If
    Bdados.FechaTabela rs
End Function
Sub GeraIptu()
        Dim Reducao As Double
        Dim Sql As String
        Dim RsCob As VSRecordset
        Dim rs As VSRecordset
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
            If Bdados.AbreTabela(Sql, rs) Then
                BaseDeCalculo = CDbl(Nvl(rs(0), 0)) - (CDbl(Nvl(rs(0), 0)) * Reducao)
            End If
            
            Sql = "select tgt_taxa_expediente,TGT_COD_PAGAMENTO,tgt_parcela,TGT_VALOR_TRIBUTO,TGT_VALOR_JUROS,TGT_VALOR_MULTA,TGT_TIP_COD_IMPOSTO,TGT_PERIODO,TGT_DATA_VENCIMENTO from Tab_Geracao_Tributo where tgt_im='" & txtIM & "' and tgt_tip_cod_imposto='" & CodImposto & "' and tgt_periodo=" & txtPeriodo & " AND  TGT_PARCELA = " & IIf(cboConta = 1, 0, 1) & " AND TGT_TIM_IC ='" & cboIC & "'"
            If Bdados.AbreTabela(Sql, rs) Then
                TaxaParcela = Nvl("" & rs!tgt_taxa_expediente, 0)
                CodPagamento = rs(1)
                txtParcela = Nvl("" & rs!TGT_PARCELA, 0)
                txtTotalImposto = rs!TGT_VALOR_TRIBUTO
                If NomeImposto = Imposto.NomeTributo(ttr_ITU) Or NomeImposto = Imposto.NomeTributo(ttr_IPTU) And txtPeriodo < Year(Date) Then
                    txtMulta = Conta.CalculaValoresMultaAvulsos(rs!TGT_TIP_COD_IMPOSTO, rs!TGT_PERIODO, EtcCreditoTributario, Date, rs!TGT_DATA_VENCIMENTO, rs!TGT_VALOR_TRIBUTO)
                    txtJuros = Conta.CalculaValoresJurosAvulsos(rs!TGT_TIP_COD_IMPOSTO, rs!TGT_PERIODO, EtcCreditoTributario, Date, rs!TGT_DATA_VENCIMENTO, rs!TGT_VALOR_TRIBUTO)
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
                TaxaServico = Nvl("" & rs!tgt_taxa_expediente, 0)
            End If
            Sql = "Select TGE_NOME from tab_geral where TGE_TIPO = 755 and TGE_CODIGO > 0"
            If Bdados.AbreTabela(Sql, rs) Then
                Desconto = Nvl("" & rs(0), 0)
            End If
            
            Sql = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & cboIC & _
                "' and (tdi_tco_cod_componente=110 or tdi_tco_cod_componente=108)"
            If Bdados.AbreTabela(Sql, rs) Then
                rs.MoveFirst
                Do While Not rs.EOF
                    If rs(0) = 110 Then
                        AreaTotal = Nvl("" & rs(1), 0)
                    ElseIf rs(0) = 108 Then
                        AreaConstruida = Nvl(rs(1), 0)
                    End If
                    rs.MoveNext
                Loop
            End If
            Bdados.FechaTabela rs
            Sql = "select tvl_valor from tab_valor_terreno where tvl_tlg_cod_logradouro=(" & _
                " select tim_tlg_cod_logradouro from tab_imovel where tim_ic='" & cboIC & "')"
            If Bdados.AbreTabela(Sql, rs) Then
                ValorMetro = Format(rs(0), Const_Monetario)
            End If
        Else
            Informa "IPTU não foi gerado."
        End If
        Bdados.FechaTabela rs
End Sub

Sub GeraAlvara()
    Dim DataVenc As String
    Dim Valores As String
    Dim Campos As String
    
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
    txtMulta = Conta.CalculaValoresMultaAvulsos(CodImposto, txtPeriodo, cboConta, Format(Date, "dd/mm/yyyy"), txtDtVenc, CDbl(txtTotalImposto))
    txtJuros = Conta.CalculaValoresJurosAvulsos(CodImposto, txtPeriodo, cboConta, Format(Date, "dd/mm/yyyy"), txtDtVenc, CDbl(txtTotalImposto))
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
    If Trim(txtIM) = "" Then txtIM = Const_ImAvulso
    
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
    CodPagamento = Conta.NumPagamento(txtIM, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), CodImposto, cboIC.Text, Nvl(txtParcela, 0), Novo, Incidencia)
    DAMOriginal = CodPagamento
    If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNEST) Then
        If cboTipo.Coluna(1).Valor = 1 Then 'DAM BRANCO
            Conta.GeraPagamento txtIM, "", CodImposto, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), txtDtVenc, CDbl(txtTotalImposto) - CDbl(Nvl(txtJuros, 0)) - CDbl(Nvl(txtMulta, 0)), 0, 0, CodPagamento, 0, Nvl(txtParcela, 0), TaxaServico, CodTaxa, , , Incidencia, , etgDAM, etdBranco
        Else
            'DAM Normal
            Conta.GeraPagamento txtIM, "", CodImposto, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), txtDtVenc, CDbl(txtTotalImposto) - CDbl(Nvl(txtJuros, 0)) - CDbl(Nvl(txtMulta, 0)), 0, 0, CodPagamento, 0, Nvl(txtParcela, 0), TaxaServico, CodTaxa, , , Incidencia, , etgDAM, etdNormal
        End If
    End If
'    Call Conta.CriaContaContribuinte(CodPagamento)
'    Call Conta.MovimentaContaContribuinte(CodPagamento)
    Bdados.DeletaDados "TAB_DETALHE_DAM", "tdd_tgt_cod_pagamento = " & DAMOriginal
    'Bdados.DeletaDados "TAB_GERACAO_TRIBUTO", "tgt_cod_pagamento = " & DAMOriginal
    Valores = Bdados.PreparaValor(CodPagamento, Nvl(txtNotaInicial, 0), Nvl(txtNotaFinal, 0), Bdados.Converte(Nvl(txtTotalNotas, 0), TCDuplo), Bdados.Converte(Nvl(txtMaterial, 0), TCDuplo), txtObservacao.Text, txtAliquota)
    Campos = "tdd_tgt_cod_pagamento,tdd_num_nota_inicial,tdd_num_nota_final,tdd_total_nota,tdd_total_material_reducao,tdd_obs,tdd_Aliquota"
    Bdados.GravaDados "TAB_DETALHE_DAM", Valores, Campos, "tdd_tgt_cod_pagamento=" & CodPagamento
    
    Valor_Imposto = txtTotalImposto
    txtTotalImposto = CDbl(txtImposto) + CDbl(txtMulta) + CDbl(txtJuros) + CDbl(Nvl(txtDAM, 0))
    Juro = txtJuros
    Multa = txtMulta
    TotalImposto = txtTotalImposto
    BaseDeCalculo = txtSaldo
End Sub

Sub GeraItbi()
        Dim Campos As String
        Dim Valores As String
        Dim Sql As String
        Dim rs As VSRecordset
        Dim Novo As Boolean
        CodPagamento = Conta.NumPagamento(txtIM, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), CodImposto, cboIC.Text, Nvl(txtParcela, 0), Novo, 0)
        Sql = "select tim_valor from tab_imovel where tim_ic='" & cboIC & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            BaseDeCalculo = rs(0)
            txtTotalImposto = rs(0) * Aliquota
        End If
        TotalImposto = txtTotalImposto
        txtImposto = TotalImposto
        Valor_Imposto = txtTotalImposto
        Bdados.FechaTabela rs
        Conta.GeraPagamento txtIM, cboIC, CodImposto, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), txtDtVenc, txtTotalImposto, 0, 0, CodPagamento, 0, Nvl(txtParcela, 0), TaxaServico, CodTaxa
End Sub

Sub GeraParcelamento()
    Dim rs As VSRecordset
    Dim Sql As String
    Dim sQL2 As String
    
    
    Sql = "Select tgt_cod_pagamento,tgt_periodo,tgt_data_vencimento from tab_geracao_tributo,tab_parcelamento where " & _
                " TPA_TCI_IM ='" & txtIM & "' and TPA_TIP_COD_IMPOSTO= '" & CodImposto & _
                "' and tpa_periodo=" & IIf(Len(txtPeriodo) = 4, txtPeriodo, IIf(Len(txtPeriodo) = 4, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2))) & _
                " and TPA_NUM_PARCELAMENTO = tgt_tpa_num_parcelamento and tgt_parcela=" & txtParcela
    If Trim(cboIC) <> "" Then
        Sql = Sql & " and TPA_TIM_IC='" & cboIC & "'"
    End If
    If Bdados.AbreTabela(Sql, rs) Then
        CodPagamento = rs(0)
        Data_Vencimento = IIf(CDbl(Format(rs(2), "yyyymmdd")) < CDbl(Format(Date, "yyyymmdd")), Format(Date, "dd/mm/yyyy"), Format(rs(2), "dd/mm/yyyy"))
        txtDtVenc = Data_Vencimento
    Else
        Informa "Não existe parcelamento para este contribuinte neste período."
        txtIM.SetFocus
        Bdados.FechaTabela rs
        Screen.MousePointer = 0
        Exit Sub
    End If
    Bdados.FechaTabela rs
    Sql = "Select tgt_valor_tributo,tgt_Valor_multa,tgt_valor_juros from tab_geracao_tributo where tgt_cod_pagamento =" & CodPagamento
    If Trim(cboIC) <> "" Then
        Sql = Sql & " and tgt_tim_ic='" & cboIC & "' and tgt_parcela=" & txtParcela
        
        sQL2 = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & cboIC & _
                "' and (tdi_tco_cod_componente=110 or tdi_tco_cod_componente=108)"
            If Bdados.AbreTabela(sQL2, rs) Then
                rs.MoveFirst
                Do While Not rs.EOF
                    If rs(0) = 110 Then
                        AreaTotal = rs(1)
                    ElseIf rs(0) = 105 Then
                        AreaConstruida = rs(1)
                    End If
                    rs.MoveNext
                Loop
            End If
            Bdados.FechaTabela rs
            sQL2 = "select tvl_valor from tab_valor_terreno where tvl_tlg_cod_logradouro=(" & _
                " select tim_tlg_cod_logradouro from tab_imovel where tim_ic='" & cboIC & "')"
            If Bdados.AbreTabela(sQL2, rs) Then
                ValorMetro = Format(rs(0), Const_Monetario)
            End If
            
            Reducao = Imposto.BuscaReducao(CodImposto, Right(txtPeriodo, 4))
            sQL2 = "select tim_valor from tab_imovel where tim_ic='" & cboIC & "'"
            If Bdados.AbreTabela(sQL2, rs) Then
                BaseDeCalculo = rs(0) - (rs(0) * Reducao)
            End If
            Bdados.FechaTabela rs
            sQL2 = "select tgt_taxa_expediente from Tab_Geracao_Tributo where tgt_im='" & txtIM & "' and tgt_tip_cod_imposto='" & CodImposto & "' and tgt_periodo=" & txtPeriodo
            If Bdados.AbreTabela(sQL2, rs) Then
                TaxaParcela = rs!tgt_taxa_expediente
            End If
            Bdados.FechaTabela rs
    End If
    If Bdados.AbreTabela(Sql, rs) Then
        txtMulta = rs!TGT_VALOR_MULTA
        txtJuros = rs!tgt_valor_juros
        Valor_Imposto = rs!TGT_VALOR_TRIBUTO
        txtTotalImposto = CDbl(Valor_Imposto) + CDbl(txtMulta) + CDbl(txtJuros)
        Juro = txtJuros
        Multa = txtMulta
        TotalImposto = txtTotalImposto
        txtImposto = Valor_Imposto
        Exercicio = txtPeriodo
    End If
    Bdados.FechaTabela rs
End Sub

Sub GeraImpostoQualquer()
    Dim Novo As Boolean
    Dim cLSImposto As VSImposto
    
    Set cLSImposto = New VSImposto
    Incidencia = cLSImposto.BuscaNumeroIncidencia(txtIM, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), CodImposto)
    CodPagamento = Conta.NumPagamento(txtIM, txtPeriodo, CodImposto, cboIC, Nvl(txtParcela, 0), Novo, Incidencia)
    Valor_Imposto = txtTotalImposto
    txtImposto = txtTotalImposto
    Conta.GeraPagamento IIf(Trim(txtIM) = "", Const_ImAvulso, txtIM), cboIC, CodImposto, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), txtDtVenc, txtTotalImposto, 0, 0, CodPagamento, 0, Nvl(txtParcela, 0), TaxaServico, CodTaxa
    txtTotalImposto = CDbl(Nvl(txtTotalImposto, 0)) + CDbl(Nvl(txtMulta, 0)) + CDbl(Nvl(txtJuros, 0))
    TotalImposto = CDbl(Nvl(txtTotalImposto, 0))
    Juro = Nvl(txtJuros, 0)
    Multa = Nvl(txtMulta, 0)
End Sub

Private Function GeraParcelamentoIptu(Cidade As CidadeIptu, Ic As String, Periodo As Integer, Tipo As TipoGeracaoImposto) As Boolean
    Dim rs As VSRecordset
    Dim Sql As String
    Dim Pagamentos As String
    Sql = "Select tgt_cod_pagamento from tab_geracao_tributo where tgt_tim_ic ='" & Ic & "' and tgt_parcela > 0 order by tgt_cod_pagamento "
    If Bdados.AbreTabela(Sql, rs) Then
        rs.MoveFirst
        Do
            Pagamentos = Pagamentos & rs(0) & "; "
            rs.MoveNext
        Loop While Not rs.EOF
        Informa "Parcelamento já existente. Pagamentos nº: " & Mid(Pagamentos, 1, Len(Pagamentos) - 2) & "."
        Screen.MousePointer = 0
        GeraParcelamentoIptu = False
        Bdados.FechaTabela rs
        Exit Function
    End If
    
    Sql = "SELECT * FROM tab_imovel,Tab_Contribuinte where tim_ic ='" & Mid(Ic, 1, 12) & "' and tim_tci_im=tci_im "
    If Bdados.AbreTabela(Sql, rs, Dinamico) Then
        Call Imposto.GeraIptu(cip_Balsas, rs, Periodo, Periodo, Tipo)
        GeraParcelamentoIptu = True
    End If
    Bdados.FechaTabela rs
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
    Dim rs As VSRecordset
    Dim Venc As String
    Dim Cobranca As New VSCobranca
    Dim DtGeracao As String
    Sql = "Select tgt_cod_pagamento,tgt_data_vencimento,tgt_valor_tributo,tgt_tim_ic,tgt_data_geracao,tgt_taxa_expediente," & _
    " tgt_periodo,tgt_tim_ic,tgt_parcela  from tab_geracao_tributo where tgt_tim_ic='" & cboIC & _
    "' and tgt_tip_cod_imposto='" & CodImposto & "' and tgt_periodo=" & txtPeriodo & " and tgt_parcela " & IIf(cboConta = "1", " =0", "> 0")
    If Bdados.AbreTabela(Sql, rs) Then
        rs.MoveFirst
        Do
            cboIC_LostFocus
            CodPagamento = rs!tgt_cod_pagamento
            Venc = rs!TGT_DATA_VENCIMENTO
            txtImposto = Nvl("" & rs!TGT_VALOR_TRIBUTO, 0)
            Juro = Conta.CalculaValoresJurosAvulsos(CodImposto, rs!TGT_PERIODO, IIf(CInt(Nvl(txtParcela, 0)) = 0, EtcCreditoTributario, EtcParcelamento), Format(Date, "dd/mm/yyyy"), Venc, Nvl("" & rs!TGT_VALOR_TRIBUTO, 0))
            Multa = Conta.CalculaValoresMultaAvulsos(CodImposto, rs!TGT_PERIODO, IIf(CInt(Nvl(txtParcela, 0)) = 0, EtcCreditoTributario, EtcParcelamento), Format(Date, "dd/mm/yyyy"), Venc, Nvl("" & rs!TGT_VALOR_TRIBUTO, 0))
            'cboIC = "" & Rs!tgt_tim_ic
            DtGeracao = "" & rs!tgt_data_geracao
            TaxaServico = rs!tgt_taxa_expediente
            txtImposto = Edita.FormataTexto(txtImposto, Monetario, True)
            txtJuros = Format(Juro, Const_Monetario)
            txtMulta = Format(Multa, Const_Monetario)
            TaxaServico = TaxaServico + TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
            DoEvents
            Cobranca.ImprimeDam Rpt, CodPagamento, txtIM, txtContribuinte, "", txtEnderecoContrib, cboIC, txtImovel, _
                CodImposto, Trim(Left(NomeImposto, PosTraco - 2)), Trim(Mid(NomeImposto, PosTraco + 2)), txtPeriodo, Nvl("" & rs!TGT_PARCELA, 0), IIf(txtParcela = "0", 1, 3), Venc, BaseDeCalculo, txtImposto, _
                txtMulta, txtJuros, TaxaServico, Desconto, Cod_Atividade, txtObservacao, PicBarra, txtNotaInicial, txtNotaFinal, txtMaterial, ValorMetro, _
                TaxaParcela, AreaTotal, AreaConstruida, ValorTerreno, Valoredific, Zona, , , String_Taxas
            rs.MoveNext
        Loop While Not rs.EOF
'        Informa "DAM Emitido e Gravaçao efetuada com sucesso."
    End If
    Screen.MousePointer = 0
    Bdados.FechaTabela rs
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
        lbl(17).Visible = False
        txtParcela.Visible = False
        txtParcela.Tag = ""
        txtParcela = ""
        'txtPeriodo.SetFocus
    End If
End Sub

Sub CarregaImovel(InscIc As String)
    Dim Sql As String
    Dim rs As VSRecordset
    
    Sql = "select ttl_nome,tlg_nome,tba_nome,tim_numero,TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC,Tim_Zona," & _
    " tim_valor,tim_tci_im ,TIM_SITUACAO_LOTE   from TAB_IMOVEL,TAB_BAIRRO," & _
    " TAB_LOGRADOURO,TAB_TIPO_LOGR " & _
    " where tim_ic='" & InscIc & _
    "' AND tim_tlg_cod_logradouro = " & _
    " TAB_LOGRADOURO.tlg_cod_logradouro AND tlg_ttl_cod_tip_logr = ttl_cod_tip_logr AND " & _
    " tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio & " AND TBA_TMU_COD_MUNICIPIO =" & _
    Aplicacoes.Codigo_Municipio & "  AND TIM_TBA_COD_BAIRRO = TBA_COD_BAIRRO"
    If Trim(cboIC) = "" Then Exit Sub
    If Bdados.AbreTabela(Sql, rs) Then
        If "" & rs!TIM_SITUACAO_LOTE = 1 Then
            Informa "Imóvel desativado."
            cboIC.SetFocus
            Exit Sub
        End If
        txtImovel = rs(0) & " " & rs(1) & " " & rs(2) & " " & rs(3)
        ValorTerreno = Nvl("" & rs!TIM_VALOR_TERRENO, 0)
        Valoredific = Nvl("" & rs!TIM_VALOR_EDIFIC, 0)
        Zona = Nvl("" & rs!tim_zona, Mid(cboIC, 4, 1))
        BaseDeCalculo = "" & rs!tim_valor
        txtIM = Nvl("" & rs!tim_tci_im, 0)
        txtIm_LostFocus
    Else
        Informa "Imovel não cadastrado."
    End If
    Bdados.FechaTabela rs
End Sub
Private Sub cboIC_LostFocus()
    CarregaImovel cboIC
End Sub


Private Sub cboImposto_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(cboImposto) = "" Then Exit Sub
    Dim i As Byte
    
    If blnConsultaIM Then Exit Sub
    
    txtIM.Enabled = True
    txtCgc.Enabled = True
    txtContribuinte = ""
    txtDtVenc = ""
    txtPeriodo = ""
    txtFator.Visible = False
    If Trim(cboImposto) <> "" Then
        i = InStr(1, cboImposto, "#")
        NomeImposto = Left(cboImposto.Text, i - 2)
    End If
    If NomeImposto = "PUBL" Then
        NomeImposto = "TSU"
    End If
    txtIM.SetFocus
    txtMaterial.Enabled = True
    If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNEST) Then
        fra(2).Visible = True
        txtMaterial.Enabled = IIf(NomeImposto = Imposto.NomeTributo(ttr_IRPJ), False, True)
        cboTipo.SetarLinha 2, 1
        cboTipo_Click
        txtPeriodo.MaxLength = 6
    ElseIf NomeImposto = Imposto.NomeTributo(ttr_IPTU) Or NomeImposto = Imposto.NomeTributo(ttr_ITBI) Or NomeImposto = Imposto.NomeTributo(ttr_2VIA) Or NomeImposto = Imposto.NomeTributo(ttr_AFORO) Or NomeImposto = Imposto.NomeTributo(ttr_VISTORIA) Or NomeImposto = Imposto.NomeTributo(ttr_DESMEMBRAMENTO) Or NomeImposto = Imposto.NomeTributo(ttr_CONSTRUCAO) Or NomeImposto = Imposto.NomeTributo(ttr_HABIT) Or NomeImposto = Imposto.NomeTributo(ttr_RECONSTRUCAO) Or NomeImposto = Imposto.NomeTributo(ttr_REMEMBRA) Then
        fra(2).Visible = False
        
        txtPeriodo.MaxLength = 4
    ElseIf NomeImposto = Imposto.NomeTributo(ttr_ALVARA) Then
        fra(2).Visible = False
        
        txtPeriodo.MaxLength = 4
    End If
    NomeImposto = Trim(Mid(cboImposto.Text, i + 2))
    CodImposto = BuscaCodigo("SELECT TIP_COD_IMPOSTO FROM TAB_IMPOSTO WHERE TIP_NOME_IMPOSTO = '" & NomeImposto & "'")
    NomeImposto = Left(cboImposto.Text, i - 2)
    'Gambi : Motivo fazer com que a taxa de publicidade fique mas pratica...
    If NomeImposto = "PUBL" Then
        NomeImposto = "TSU"
        CodImposto = "11220100"
    End If
    'Ver se é Taxa
    Sql = "Select tpi_tipo_tributo,tpi_valor_taxa_fixa,tpi_tipo_inscricao,tpi_tipo_ic from tab_parametro_imposto where tpi_tip_cod_imposto ='" & CodImposto & "'"
    If Bdados.AbreTabela(Sql, rs) Then
           TributoTaxa = IIf(rs!tpi_tipo_tributo = 2, True, False)
           If TributoTaxa Then
                lbl(16).Visible = False
                cboConta.Visible = False
                cboConta.Tag = ""
           Else
                lbl(16).Visible = True
                cboConta.Visible = True
                cboConta.Tag = "Conta"
                
                txtFator.Visible = False
                lbl(21).Visible = False
                
                lbl(22).Visible = False
                txtFatorAleatorio.Visible = False

           End If
           If Nvl("" & rs!tpi_tipo_ic, 0) = 1 Then
                fraIm.Visible = True
           Else
                fraIm.Visible = False
           End If
           If TributoTaxa Then
                txtContribuinte.Enabled = True
                fra(2).Visible = False
                txtPeriodo.MaxLength = 7
                txtPeriodo = Format(Month(Date), "00") & Format(Year(Date), "0000")
                txtPeriodo_LostFocus
                txtPeriodo.MaxLength = 6
                Exercicio = txtPeriodo
                cboConta.Tag = ""
                txtIM = ""
                txtCgc = ""
                txtFator.Visible = True
                lbl(21).Visible = True
                If TributoTaxaFixa = 0 Then
                    lbl(22).Visible = True
                    txtFatorAleatorio.Visible = True
                Else
                    lbl(22).Visible = False
                    txtFatorAleatorio.Visible = False
                End If
                
           ElseIf rs!tpi_tipo_inscricao = 2 Then
                txtContribuinte.Enabled = False
                txtIM.Enabled = True
                txtCgc.Enabled = True
                txtFator.Visible = False
                lbl(21).Visible = False
                
                lbl(22).Visible = False
                txtFatorAleatorio.Visible = False
           End If
    Else
            lbl(21).Visible = False
            txtFator.Visible = False
            lbl(22).Visible = False
            txtFatorAleatorio.Visible = False
        
    End If
    'Se for publicidade monto acombo com os items da publiciade
    If NomeImposto = "PUBL" Then
        If Bdados.AbreTabela("SELECT TPD_TIP_COD_IMPOSTO,tpd_descricao ,tpd_item  From TAB_PARAMETRO_DETALHE where  TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(CodImposto, tctexto)) Then
            CboItem.Visible = True
            lbl(23).Visible = True
            CboItem.Preencher Bdados, "SELECT TPD_TIP_COD_IMPOSTO,tpd_descricao ,tpd_item  From TAB_PARAMETRO_DETALHE where  TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(CodImposto, tctexto), 1
        Else
            CboItem.Visible = False
            lbl(23).Visible = False
        End If
    Else
        CboItem.Visible = False
        lbl(23).Visible = False
    End If
    Bdados.FechaTabela rs
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
    Dim rs As VSRecordset
    Dim Sql As String
    Dim SqlParc As String
    Dim Cobranca As New VSCobranca
    
    txtPeriodo = TiraTudo(txtPeriodo)
        
    Select Case cmd(Index).Caption
        
        Case "&Emitir DAM"
            If Not Util.Confirma("Confirma a emissão do DAM") Then
                Exit Sub
            End If
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
            'If cboTaxa.ListIndex = -1 Then
            'Else
            '    TaxaServico = BuscaValorTaxa(cboTaxa.Text, CodTaxa) + txtDam
            'End If
            
            'If IsNumeric(cboIC) And Len(cboIC) <> 15 Then
            '    Informa "Informe a inscrição cadastral completa do imóvel."
            '    cboIC.SetFocus
            '    Screen.MousePointer = 0
            '    Exit Sub
            'End If
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
                    GoTo Gambi 'Motivo Fazercom que  as taxa de Publicidade fiquem pratica...
                   'procedimento que vai me retornar o codigo da taxa atraves do nome da mesma...
                    Dim Pos
                    Dim Codigo_I As String
                    Dim I_P As String
                    Pos = InStr(cboImposto, "#") + 1
                    Codigo_I = Right(cboImposto, Len(cboImposto) - Pos)
                    BuscaValorTaxa Codigo_I, I_P
                    'Checo se existem taxas para esse contribuinte...
                    Sql = "Select * from tab_Anuncio where TAN_TCI_IM = " & Bdados.Converte(txtIM, tctexto) & " and  TAN_MOVIMENTO = " & Bdados.Converte(I_P, tctexto)
                    If Bdados.AbreTabela(Sql, rs) Then
                        If CboItem.Visible Then
                            Do Until rs.EOF
                                If Trim(rs.Fields("TAN_MOVIMENTO")) = Trim(I_P) And Trim(rs.Fields("TAN_TPD_ITEM")) = Trim(CStr(CboItem.Coluna(2).Valor)) Then
                                    GeraTaxas rs.Fields("TAN_VALOR_APAGAR")
                                    Exit Do
                                End If
                                rs.MoveNext
                            Loop
                        Else
                            Do Until rs.EOF
                                If Trim(rs.Fields("TAN_MOVIMENTO")) = Trim(I_P) Then
                                    GeraTaxas rs.Fields("TAN_VALOR_APAGAR")
                                    Exit Do
                                End If
                                rs.MoveNext
                            Loop
                        End If
                    Else
                        If Confirma("Não existem anuncios para este contribuinte.") = False Then
                          Screen.MousePointer = 0
                            Exit Sub
                        End If
                    End If
                    
                Else
                    If TributoTaxaFixa > 0 Then
                        GeraTaxas IIf(txtFator.Visible = True, CDbl(Nvl(CStr(TributoTaxaFixa), Nvl(txtFatorAleatorio, 1))) * CDbl(Nvl(txtFator, 1)), TributoTaxaFixa)
                    Else
Gambi:
                        GeraTaxas CDbl(Nvl(txtFator, 0)) * CDbl(Nvl(txtFatorAleatorio, 0))
                    End If
                End If
            Else 'Imposto Comum
                If cboConta.Text <> "3" Then 'Credito Tributario / Auto Infracao
                    Sql = "Select tgt_valor_tributo,tgt_data_geracao,tgt_cod_pagamento from tab_geracao_tributo where tgt_tip_cod_imposto='" & CodImposto & "' and tgt_periodo=" & IIf(Len(txtPeriodo) = 4, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2)) & " and tgt_parcela =" & Nvl(txtParcela, 0)
                    If Trim(cboIC) <> "" Then
                        Sql = Sql & " AND tgt_tim_ic ='" & cboIC & "'"
                    Else
                        Sql = Sql & " AND tgt_im='" & txtIM & "'"
                    End If
                    If Bdados.AbreTabela(Sql, rs) Then
                        If NomeImposto <> Imposto.NomeTributo(ttr_ISSQNRET) And NomeImposto <> Imposto.NomeTributo(ttr_ISSQNSUBST) And NomeImposto <> Imposto.NomeTributo(ttr_ISSQNCOMP) And NomeImposto <> Imposto.NomeTributo(ttr_ISSQNEST) Then
                            Call Informa("DAM já emitido no valor de R$ " & Format(rs(0), Const_Monetario) & ", em " & IIf(Not IsNull(rs(1)), rs(1), "") & ". NÚMERO DO PAGAMENTO: " & rs!tgt_cod_pagamento & ".")
                            Screen.MousePointer = 0
                            Bdados.FechaTabela rs
                            Exit Sub
                        End If
                        Bdados.FechaTabela rs
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
                    If TotalImposto > 0 Then
                        PosTraco = InStr(1, cboImposto, "#")
                        Cobranca.ImprimeDam Rpt, CodPagamento, txtIM, txtContribuinte, "" & txtCgc, txtEnderecoContrib, cboIC, txtImovel, _
                        CodImposto, Trim(Left(cboImposto, PosTraco - 2)), Trim(Mid(cboImposto, PosTraco + 2)), txtPeriodo, txtParcela, Nvl(cboConta, 1), txtDtVenc, BaseDeCalculo, txtImposto, _
                        Multa, Juro, TaxaServico, "", Cod_Atividade, txtObservacao, PicBarra, txtNotaInicial, txtNotaFinal, txtMaterial, ValorMetro, _
                        TaxaParcela, AreaTotal, AreaConstruida, , , , , , String_Taxas
                        Desconto = 0
                    Else
                        Informa "Total do imposto igual zero. DAM não será emitido."
                        Bdados.DeletaDados "tab_geracao_tributo", "tgt_cod_pagamento=" & CodPagamento
                        Bdados.DeletaDados "tab_conta_contribuinte", "tcc_codigo_conta=" & CodPagamento
                    End If
                Else
                        PosTraco = InStr(1, cboImposto, "#")
                        Cobranca.ImprimeDam Rpt, CodPagamento, txtIM, txtContribuinte, "" & txtCgc, txtEnderecoContrib, cboIC, txtImovel, _
                        CodImposto, Trim(Left(cboImposto, PosTraco - 2)), Trim(Mid(cboImposto, PosTraco + 2)), txtPeriodo, txtParcela, Nvl(cboConta, 1), txtDtVenc, BaseDeCalculo, txtImposto, _
                        Multa, Juro, TaxaServico, "", Cod_Atividade, txtObservacao, PicBarra, txtNotaInicial, txtNotaFinal, txtMaterial, ValorMetro, _
                        TaxaParcela, AreaTotal, AreaConstruida, , , , , etdBranco, String_Taxas
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
    Call Edita.AtualizaCombo(Bdados, cboImposto, "Select  TIP_sigla_IMPOSTO  " & Bdados.Concatena & " ' # ' " & Bdados.Concatena & " tip_nome_imposto From TAB_IMPOSTO " _
        & " WHERE  tip_sigla_imposto like 'ISS%' or tip_cod_imposto in (SELECT tpi_tip_cod_imposto FROM Tab_Parametro_Imposto where TPI_GERA_OBRIGACAO = 2 OR TPI_GERA_OBRIGACAO  IS NULL)")
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
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    txtDAM = TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
    txtDAM.Enabled = False
    cboConta.ListIndex = 0
    
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
    Dim rs As VSRecordset
    
    If Trim(txtIM) <> "" Then Exit Sub
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
    If Bdados.AbreTabela(Sql, rs) Then
        txtContribuinte = rs(0)
        txtIM = "" & rs!TCI_IM
        Endereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_COMPLEMENTO
        Bairro = "" & rs!tci_BAIRRO
        Cod_Cidade = "" & rs!tci_cidade
        Uf = "" & rs!tci_UF
        Call txtIm_LostFocus
        
        If NomeImposto = Imposto.NomeTributo(ttr_IPTU) Or NomeImposto = Imposto.NomeTributo(ttr_ITBI) Then
            cboIC.Enabled = True
            cboIC.Visible = True
        End If
    ElseIf Not TributoTaxa Then
        Util.Informa "Cgc/Cpf não cadastrado."
        txtContribuinte.Enabled = True
    Else
        Cod_Atividade = ""
        txtContribuinte.Enabled = True
        
    End If
    Bdados.FechaTabela rs
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
    On Error Resume Next
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
    Dim rs As VSRecordset
    Dim RsIptu As VSRecordset
    If Trim(txtIM) = "" Then Exit Sub
    Sql = "select * from Tab_Contribuinte where tci_im='" & txtIM & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtContribuinte = rs!tci_nome
        txtEnderecoContrib = rs!tci_logradouro & " " & rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_BAIRRO & ", CEP: " & rs!tci_cep & ", " & rs!tci_cidade & "-" & rs!tci_UF
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
        If IsNull(rs("tci_tae_cae")) And (NomeImposto <> Imposto.NomeTributo(ttr_IPTU) And NomeImposto <> Imposto.NomeTributo(ttr_ITBI)) Then
            If Not TributoTaxa Then
                Informa "Contribuinte não sujeito à cobrança de " & NomeImposto & "."
                txtIM.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        txtContribuinte = "" & rs("tci_nome") & ""
        txtCgc = "" & rs("tci_cgc_cpf") & ""
        
        Cod_Atividade = IIf(Nvl("" & rs("tci_tae_cae"), 0) = 0, "", rs("tci_tae_cae")) & Imposto.BuscaNomeCAE(rs("tci_tae_cae") & "") & ""
        
        Endereco = "" & rs("tci_logradouro") & "  " & rs("tci_nome_logradouro") & "," & rs("tci_numero") & " " & rs("tci_complemento")
        Bairro = "" & rs("tci_bairro") & ""
        Cod_Cidade = "" & rs("tci_cidade") & ""
        Cep = "" & rs("tci_cep") & ""
        Uf = "" & rs("tci_uf") & ""
        If Not TributoTaxa Then
            'If ((RS!tci_tipo_recolhimento_iss = 1 Or RS!tci_tipo_recolhimento_iss = 3) And NomeImposto = Imposto.NomeTributo(ttr_ISSQN)) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_IRPJ) Then
            If ((rs!tci_tipo_recolhimento_iss = 1 Or rs!tci_tipo_recolhimento_iss = 3) And NomeImposto = Imposto.NomeTributo(ttr_ISSQN)) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Then
                txtPeriodo.MaxLength = 6
            ElseIf rs!tci_tipo_recolhimento_iss = 2 Then
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
    Bdados.FechaTabela rs
    Bdados.FechaTabela RsIptu
    txtPeriodo.SetFocus
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
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Venc As String
    Dim Hoje As String
    Dim SqlParc As String
    Dim cAtividade As Atividade
    
    
    Dim ValorAtividade As Double
    txtPeriodo = Edita.TiraTudo(txtPeriodo)
    If Not IsNumeric(txtPeriodo) Then Exit Sub
    If Trim(cboImposto) <> "" And Trim(txtPeriodo) <> "" Then
        If Len(txtPeriodo) < 4 Then
            Informa "Período inválido."
            txtPeriodo.SetFocus
            Exit Sub
        End If
        If cboConta <> "3" Then
           'Ver se é Taxa com valor fixo
           Sql = "Select tpi_tipo_tributo,tpi_valor_taxa_fixa from tab_parametro_imposto where tpi_tip_cod_imposto ='" & CodImposto & "'"
           If CInt(Right(txtPeriodo, 4)) <> CInt(Year(Date)) Then Sql = Sql & " and tpi_ano_imposto= '" & Imposto.BuscaAnoImposto(CodImposto, Right(txtPeriodo, 4)) & "'"
           If Bdados.AbreTabela(Sql, rs) Then
                  If rs!tpi_tipo_tributo = 2 Then
                      TributoTaxa = True
                      TributoTaxaFixa = IIf(IsNull(rs!tpi_valor_taxa_fixa), 0, rs!tpi_valor_taxa_fixa)
                      Venc = Imposto.BuscaDataVencimento(CodImposto, IIf(TributoTaxa, "", Left(txtPeriodo, 2) & Right(txtPeriodo, 4)))
                      txtPeriodo.MaxLength = 7
                      txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
                      txtDtVenc = Venc
                      txtPeriodo.MaxLength = 6
                      Exercicio = txtPeriodo
                      Bdados.FechaTabela rs
                      Exit Sub
                  End If
           End If
           Bdados.FechaTabela rs
        Else
            If NomeImposto = Imposto.NomeTributo(ttr_IPTU) Then  'SO PARA IPTU
                If CInt(txtPeriodo) = 2001 Then
                    Sql = "SELECT TGE_NOME,TGE_CODIGO FROM TAB_GERAL WHERE TGE_TIPO = 710 AND TGE_CODIGO =" & txtParcela
                    If Bdados.AbreTabela(Sql, rs) Then
                        Venc = rs(0)
                        txtDtVenc = rs(0)
                        Bdados.FechaTabela rs
                        Exit Sub
                    Else
                        Informa "Número de parcela inválido."
                        txtParcela.SetFocus
                        Exit Sub
                    End If
                    Bdados.FechaTabela rs
                End If
            Else 'ANALISAR DEPOIS
                Sql = "Select tgt_data_vencimento from tab_geracao_tributo where tgt_TPA_NUM_PARCELAMENTO in " & _
                "  (select TPA_NUM_PARCELAMENTO from tab_parcelamento where TPA_TCI_IM ='" & txtIM & _
                "' and TPA_TIP_COD_IMPOSTO='" & CodImposto & "' and TPA_PERIODO=" & IIf(Len(txtPeriodo) = 4, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2)) & ") and tgt_parcela=" & txtParcela
                If Bdados.AbreTabela(Sql, rs) Then
                    txtDtVenc = rs(0)
                    Venc = rs(0)
                End If
                Bdados.FechaTabela rs
                If Trim(txtPeriodo) <> "" Then Aliquota = Imposto.BuscaAliquota(CodImposto, Right(txtPeriodo, 4))
                
                txtPeriodo.MaxLength = txtPeriodo.MaxLength + 1
                If Len(txtPeriodo) = 6 Then txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
                txtPeriodo.MaxLength = txtPeriodo.MaxLength - 1
                Exit Sub
            End If
        End If
        If Trim(txtPeriodo) <> "" Then Aliquota = CDbl(Imposto.BuscaAliquota(CodImposto, Right(txtPeriodo, 4)))
        If Len(txtPeriodo) <> 4 And Len(txtPeriodo) <> 6 And Len(txtPeriodo) <> 7 Then
            Informa "Período inválido."
            txtPeriodo = ""
            txtPeriodo.SetFocus
            Exit Sub
        End If
        'buscar o valor fixo ou alíquota correspondente
        Set cAtividade = New Atividade
        'Resume
        If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNEST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Then
            ValorAtividade = 0
            Aliquota = cAtividade.BuscaAliquotaAtividade(Bdados, Nvl("" & txtIM, 0), txtPeriodo, ValorAtividade)
            txtDtVenc.Enabled = True
        End If
        ' *********************
        Set cAtividade = Nothing
        If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Then
            If Len(txtPeriodo) <> 6 And Len(txtPeriodo) <> 4 Then
                Informa "Período inválido."
                txtPeriodo = ""
                txtPeriodo.SetFocus
                Exit Sub
            End If
            txtDtVenc.Enabled = True
        End If
       
        If NomeImposto = Imposto.NomeTributo(ttr_ITBI) Then
            If Len(txtPeriodo) <> 4 Then
                Informa "Período inválido."
                txtPeriodo = ""
                txtPeriodo.SetFocus
                Exit Sub
            Else
                txtDtVenc = Imposto.BuscaDataVencimento(CodImposto, IIf(TributoTaxa, "", Left(txtPeriodo, 2) & Right(txtPeriodo, 4)))
                Exit Sub
            End If
        End If
        
        If cboConta.Text = "1" Or cboConta.Text = "4" Then
            If NomeImposto = Imposto.NomeTributo(ttr_ITU) Or NomeImposto = Imposto.NomeTributo(ttr_IPTU) And txtPeriodo < Year(Date) Then
                Venc = Imposto.BuscaDataVencimento(CodImposto, IIf(TributoTaxa, "", Left(txtPeriodo, 2) & Right(txtPeriodo, 4)))
            Else
                Venc = Imposto.BuscaDataVencimento(CodImposto, txtPeriodo)
            End If
        ElseIf cboConta.Text = "3" Then
            Sql = "Select tgt_data_vencimento,TPA_NUM_COTAS from Tab_Geracao_Tributo,TAB_PARCELAMENTO " & _
                " where tgt_tip_cod_imposto='" & CodImposto & "' " & _
                " " & _
                "' AND tgt_im=TPA_TCI_IM AND tgt_tip_cod_imposto= TPA_TIP_COD_IMPOSTO AND tgt_tpa_num_parcelamento=TPA_NUM_PARCELAMENTO"
            If Trim(cboIC.Text) <> "" Then
                Sql = Sql & " and tgt_tim_ic='" & cboIC.Text & "' AND tgt_tim_ic=TPA_TIM_IC "
            Else
                Sql = Sql & " and tgt_im='" & txtIM & "'"
            End If
            Sql = Sql & " ORDER BY TGT_PERIODO asc"
            If Bdados.AbreTabela(Sql, rs) Then
                If Trim(txtParcela) = "" Then
                    Informa "Informe nº da parcela."
                    txtParcela.SetFocus
                    Bdados.FechaTabela rs
                    Exit Sub
                End If
                
                If rs(1) < txtParcela Then
                    Util.Informa "Total de cotas deste parcelamento igual à: " & rs(1)
                    txtParcela.SetFocus
                    Exit Sub
                End If
                rs.MoveLast
                rs.MoveFirst
                rs.Move CInt(Trim(txtParcela)) - 1
                Venc = rs(0)
            Else
                Informa "Não existe parcelamento para este contribuinte neste período."
                txtPeriodo.SetFocus
                txtDtVenc = ""
                Exit Sub
            End If
        ElseIf txtContribuinte.Enabled = False Then
            Informa "Informe a conta."
            cboConta.SetFocus
        End If
        If Venc = "0" Then
            Informa "Não existe formulário definido."
            'txtPeriodo.SetFocus
            txtDtVenc = ""
            Exit Sub
        End If
        Venc = Format(Venc, "dd/mm/yyyy")
        Exercicio = txtPeriodo
        txtDtVenc = Venc
    End If
    txtPeriodo.MaxLength = txtPeriodo.MaxLength + 1
    If Len(txtPeriodo) = 6 Then txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
    txtPeriodo.MaxLength = txtPeriodo.MaxLength - 1
    If Trim(txtImposto) <> "" Then
        If NomeImposto = Imposto.NomeTributo(ttr_ITU) Or NomeImposto = Imposto.NomeTributo(ttr_IPTU) And txtPeriodo < Year(Date) Then
            txtJuros = 0
            txtMulta = 0
        Else
            txtJuros = Format(Conta.CalculaValoresJurosAvulsos(CodImposto, IIf(Len(txtPeriodo) = 4, txtPeriodo, CLng(Right(txtPeriodo, 4) & Left(txtPeriodo, 2))), Nvl(cboConta, 1), Format(Date, "dd/mm/yyyy"), txtDtVenc, txtImposto), Const_Monetario)
            txtMulta = Format(Conta.CalculaValoresMultaAvulsos(CodImposto, IIf(Len(txtPeriodo) = 4, txtPeriodo, CLng(Right(txtPeriodo, 4) & Left(txtPeriodo, 2))), Nvl(cboConta, 1), Format(Date, "dd/mm/yyyy"), txtDtVenc, txtImposto), Const_Monetario)
        End If
        txtTotalImposto = Format(CDbl(Nvl(txtImposto, 0)) + CDbl(Nvl(txtJuros, 0)) + CDbl(Nvl(txtMulta, 0)), Const_Monetario)
    End If
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
