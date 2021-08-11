VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCOB204 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
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
      TabIndex        =   54
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCOB204.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   46
      Top             =   7260
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   5565
         TabIndex        =   48
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   4140
         TabIndex        =   47
         Top             =   90
         Width           =   1380
         _ExtentX        =   2434
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
         TabIndex        =   49
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   3465
      Index           =   0
      Left            =   53
      TabIndex        =   18
      Top             =   660
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6112
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
         Height          =   510
         Left            =   1650
         TabIndex        =   5
         Top             =   1695
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   900
         Caption         =   "Vencimento"
         Text            =   ""
         Formato         =   0
         AlinhamentoRotulo=   1
         RetirarMascara  =   0   'False
      End
      Begin VB.TextBox txtNomeImposto 
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   44
         Tag             =   "NO. DAM"
         Top             =   750
         Width           =   7575
      End
      Begin VB.TextBox txtDAM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5430
         MaxLength       =   14
         TabIndex        =   0
         Tag             =   "NO. DAM"
         Top             =   300
         Width           =   2235
      End
      Begin VB.TextBox txtParcela 
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
         Left            =   6660
         TabIndex        =   3
         Top             =   1350
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Frame fraIm 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   690
         Left            =   60
         TabIndex        =   35
         Top             =   2730
         Visible         =   0   'False
         Width           =   7725
         Begin VB.ComboBox cboIC 
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
            ItemData        =   "TCOB204.frx":2123
            Left            =   30
            List            =   "TCOB204.frx":2125
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   330
            Width           =   1635
         End
         Begin VB.TextBox txtImovel 
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
            TabIndex        =   36
            Top             =   330
            Width           =   5895
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   6
            Left            =   60
            TabIndex        =   37
            Top             =   60
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
         Left            =   90
         TabIndex        =   7
         Top             =   2415
         Width           =   7545
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
         Left            =   60
         TabIndex        =   4
         Tag             =   "Exercicio"
         Top             =   1890
         Width           =   1335
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
         Left            =   60
         TabIndex        =   1
         Top             =   1350
         Width           =   1485
      End
      Begin VB.TextBox txtCgc 
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
         Left            =   1680
         TabIndex        =   2
         Top             =   1350
         Width           =   1815
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   60
         TabIndex        =   19
         Top             =   540
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
         Left            =   1680
         TabIndex        =   20
         Top             =   1140
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
         Height          =   270
         Index           =   9
         Left            =   60
         TabIndex        =   25
         Top             =   1140
         Width           =   1410
         _ExtentX        =   2487
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
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   12
         Left            =   120
         TabIndex        =   26
         Top             =   2220
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
         TabIndex        =   27
         Top             =   1680
         Width           =   1410
         _ExtentX        =   2487
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
         Index           =   17
         Left            =   6660
         TabIndex        =   38
         Top             =   1110
         Visible         =   0   'False
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
         Caption         =   "Parcela:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   7
         Left            =   5460
         TabIndex        =   43
         Top             =   60
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
         Caption         =   "Nº DAM:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   2370
      Index           =   2
      Left            =   45
      TabIndex        =   21
      Top             =   4890
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4180
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
      Begin VB.TextBox txtAliquota 
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
         Left            =   2955
         TabIndex        =   51
         Top             =   870
         Width           =   1275
      End
      Begin VB.TextBox txtTotalImposto 
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
         TabIndex        =   17
         Tag             =   " "
         Top             =   1935
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
         TabIndex        =   16
         Top             =   1605
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
         TabIndex        =   14
         Top             =   945
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
         TabIndex        =   13
         Top             =   615
         Width           =   1275
      End
      Begin VB.TextBox txtMaterial 
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
         Left            =   2955
         TabIndex        =   12
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txtTotalNotas 
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
         Left            =   2955
         TabIndex        =   11
         Top             =   1200
         Width           =   1275
      End
      Begin VB.TextBox txtNotaFinal 
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
         Left            =   2955
         TabIndex        =   10
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox txtNotaInicial 
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
         Left            =   2955
         TabIndex        =   9
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
         TabIndex        =   15
         Top             =   1275
         Width           =   1275
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   11
         Left            =   1800
         TabIndex        =   22
         Top             =   270
         Width           =   1230
         _ExtentX        =   2170
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
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   18
         Left            =   5835
         TabIndex        =   23
         Top             =   1335
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
         Left            =   1845
         TabIndex        =   28
         Top             =   630
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
         Caption         =   "Nº Nota Final:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   5
         Left            =   1635
         TabIndex        =   29
         Top             =   1260
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
         Caption         =   "Total em Notas:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   8
         Left            =   135
         TabIndex        =   30
         Top             =   1605
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
         Left            =   4965
         TabIndex        =   31
         Top             =   675
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
         TabIndex        =   32
         Top             =   1005
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
         TabIndex        =   33
         Top             =   1680
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
         Left            =   4980
         TabIndex        =   34
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
         Left            =   1980
         TabIndex        =   50
         Top             =   945
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "Aliquota %:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   26
         Left            =   2400
         TabIndex        =   52
         Top             =   1935
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
         Left            =   2955
         TabIndex        =   53
         Top             =   1875
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
         Enabled         =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1200
      TabIndex        =   24
      Top             =   960
      Width           =   375
   End
   Begin VB.Timer tmr 
      Interval        =   10
      Left            =   2430
      Top             =   4890
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
      Left            =   683
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1230
      Width           =   5895
   End
   Begin Threed.SSFrame fra 
      Height          =   765
      Index           =   1
      Left            =   53
      TabIndex        =   40
      Top             =   4110
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1349
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
         Left            =   1380
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   6255
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   19
         Left            =   120
         TabIndex        =   41
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
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4560
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   42
      Top             =   690
      Visible         =   0   'False
      Width           =   795
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   1138
      Icone           =   "TCOB204.frx":2127
   End
   Begin VTOcx.grdVISUAL Grdtaxas 
      Height          =   6840
      Left            =   7890
      TabIndex        =   55
      Top             =   675
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   12065
      Caption         =   "Taxas"
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
      Ordenavel       =   0   'False
   End
End
Attribute VB_Name = "TCOB204"
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
Dim DamPago  As Boolean
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
Dim DtGeracao As String
Dim CodPagamento As String
Dim JurosReal As Double
Dim MultaReal As Double
Dim CodImpostoOriginal As String
Dim Mudou As Boolean
Dim DataOriginalVence As String
Dim CodTaxa As String
Dim String_Taxas As String
Dim Total_Taxas As Double

Private Sub cboIC_Click()
    If Trim(cboIC) <> "" Then CarregaImovel cboIC
End Sub
Sub CarregaImovel(InscIc As String)
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Logradouro As String
    Sql = "select ttl_nome,tlg_nome,tba_nome,tim_numero,TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC,Tim_Zona,tim_valor ,TAB_IMOVEL.TIM_tlg_cod_logradouro   from TAB_IMOVEL,VIS_BVT " & _
    " where tim_ic='" & cboIC & "' AND tim_tlg_cod_logradouro = " & " VIS_BVT.tlg_cod_logradouro "
    If Bdados.AbreTabela(Sql, rs) Then
        txtImovel = rs(0) & " " & rs(1) & " " & rs(2) & " " & rs(3)
        Logradouro = "" & rs!TIM_tlg_cod_logradouro
        BaseDeCalculo = Nvl("" & rs!tim_valor, 0)
        If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
            ValorTerreno = Nvl("" & rs!TIM_VALOR_TERRENO, 0)
            Valoredific = Nvl("" & rs!TIM_VALOR_EDIFIC, 0)
            Zona = Nvl("" & rs!tim_zona, 1)
        Else
            Sql = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & cboIC & _
                    "' and (tdi_tco_cod_componente=110 or tdi_tco_cod_componente=108)"
                If Bdados.AbreTabela(Sql, rs) Then
                    rs.MoveFirst
                    Do While Not rs.EOF
                        If rs(0) = 110 Then
                            AreaTotal = rs(1)
                        ElseIf rs(0) = 108 Then
                            AreaConstruida = rs(1)
                        End If
                        rs.MoveNext
                    Loop
                End If
                Bdados.FechaTabela rs
                
                Sql = "select tvl_valor  as ValorMetro from TAB_VALOR_TERRENO where tvl_tlg_cod_logradouro='" & Logradouro & "'"
                If Bdados.AbreTabela(Sql, rs) Then
                    ValorMetro = rs!ValorMetro
                End If
                Bdados.FechaTabela rs
        End If
    End If
    Bdados.FechaTabela rs
End Sub
Private Sub cboIC_LostFocus()
    cmd(1).SetFocus
End Sub


Private Sub cboTipo_Click()
    If cboTipo.Coluna(1).Valor = 1 Then
        txtTotalImposto = "0,00"
        txtDam = "0,00"
        txtJuros = "0,00"
        txtMulta = "0,00"
        txtImposto = "0,00"
    Else
        'txtTotalNotas_Change
        txtDam = TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXTDAM"), 0), ".", ",")
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    On Error GoTo TRATA
    Dim a As Integer
    Dim Valores As String
    Dim Campos As String
    Dim ValorImposto As Double
    Dim RsCob As VSRecordset
    Dim rs As VSRecordset
    Dim Sql As String
    Dim SqlParc As String
    Dim Cobranca As New VSCobranca
    Dim ImpostoOriginal As String
    Documento = CodPagamento
    Data_Vencimento = txtDtVenc
    Cod_Tributo = NomeImposto
    CPFCNPJ = txtcgc
    InscMuni = txtIm
    RazaoSocial = txtContribuinte
    Juro = 0
    Multa = 0
    TotalImposto = 0
    Linhas = 0
    Screen.MousePointer = 11
    TaxaServico = 0
    'Pego as taxas
    Call Pega_taxas
    TaxaServico = TaxaServico + Total_Taxas
    Select Case cmd(Index).Caption
        
        Case "&Emitir DAM"
            If Trim$(txtDtVenc) = "" Then txtDtVenc = Format(Now, "dd/mm/yyyy")
            txtContribuinte.Enabled = True
            
            If Not Edita.CriticaCampos(Me) Then
                txtContribuinte.Enabled = True
                Screen.MousePointer = 0
                Exit Sub
            End If
            PosTraco = InStr(1, txtNomeImposto, "-")
            If Trim(Left(txtNomeImposto, PosTraco - 2)) = Imposto.NomeTributo(ttr_DATIVA) And Trim(CodImpostoOriginal) <> "" Then
                Sql = "Select tip_sigla_imposto from tab_imposto where tip_cod_imposto ='" & CodImpostoOriginal & "'"
                If Bdados.AbreTabela(Sql, rs) Then
                    ImpostoOriginal = " - " & rs(0)
                End If
                Bdados.FechaTabela rs
            End If
            Dim pos As Integer
            Dim Sigla As String
            pos = InStr(txtNomeImposto, "-")
            Sigla = Trim(Left(txtNomeImposto, pos - 1))
            If Sigla = Imposto.NomeTributo(ttr_ISSQN) Or Sigla = Imposto.NomeTributo(ttr_ISSQNRET) Or Sigla = Imposto.NomeTributo(ttr_ISSQNSUBST) Or Sigla = Imposto.NomeTributo(ttr_ISSQNCOMP) Or Sigla = Imposto.NomeTributo(ttr_ISSQNEST) Then
                If cboTipo.Coluna(1).Valor = 1 Then
                    Cobranca.ImprimeDam Rpt, CodPagamento, txtIm, txtContribuinte, txtcgc, txtEnderecoContrib, cboIC, txtImovel, _
                    CodImposto, Trim(Left(txtNomeImposto, PosTraco - 2)), Trim(Mid(txtNomeImposto, PosTraco + 2)) & ImpostoOriginal, txtPeriodo, txtParcela, IIf(txtParcela = "0", 1, 3), txtDtVenc, BaseDeCalculo, txtImposto, _
                    txtMulta, txtJuros, TaxaServico, Desconto, Cod_Atividade, txtObservacao, PicBarra, txtNotaInicial, txtNotaFinal, txtMaterial, ValorMetro, _
                    TaxaParcela, AreaTotal, AreaConstruida, ValorTerreno, Valoredific, Zona, , etdBranco, String_Taxas
                    Desconto = 0
                Else
                    Cobranca.ImprimeDam Rpt, CodPagamento, txtIm, txtContribuinte, txtcgc, txtEnderecoContrib, cboIC, txtImovel, _
                    CodImposto, Trim(Left(txtNomeImposto, PosTraco - 2)), Trim(Mid(txtNomeImposto, PosTraco + 2)) & ImpostoOriginal, txtPeriodo, txtParcela, IIf(txtParcela = "0", 1, 3), txtDtVenc, BaseDeCalculo, txtImposto, _
                    txtMulta, txtJuros, TaxaServico, Desconto, Cod_Atividade, txtObservacao, PicBarra, txtNotaInicial, txtNotaFinal, txtMaterial, ValorMetro, _
                    TaxaParcela, AreaTotal, AreaConstruida, ValorTerreno, Valoredific, Zona, , etdNormal, String_Taxas
                    Desconto = 0
                End If
            Else
                Cobranca.ImprimeDam Rpt, CodPagamento, txtIm, txtContribuinte, txtcgc, txtEnderecoContrib, cboIC, txtImovel, _
                CodImposto, Trim(Left(txtNomeImposto, PosTraco - 2)), Trim(Mid(txtNomeImposto, PosTraco + 2)) & ImpostoOriginal, txtPeriodo, txtParcela, IIf(txtParcela = "0", 1, 3), txtDtVenc, BaseDeCalculo, txtImposto, _
                txtMulta, txtJuros, TaxaServico, Desconto, Cod_Atividade, txtObservacao, PicBarra, txtNotaInicial, txtNotaFinal, txtMaterial, ValorMetro, _
                TaxaParcela, AreaTotal, AreaConstruida, ValorTerreno, Valoredific, Zona, , etdNormal, String_Taxas
                Desconto = 0
            End If
            fraIm.Visible = False
'            Informa "DAM Reemitido com sucesso."
        Case "Sai&r"
           Unload Me
    End Select
    Screen.MousePointer = 0
    Exit Sub
TRATA:
    Erro Err.Description
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtDam.SetFocus
End Sub

Private Sub Form_Activate()
    Mudou = False
End Sub

Private Sub Form_Load()
    Dim Controle As Control
    
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    cboTipo.PreencherGeral Bdados, "TIPO DAM"
    Grdtaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
End Sub


Private Sub txtcgc_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtcgc_LostFocus()
    
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtcgc) = "" Then Exit Sub
    If Len(txtcgc) = 18 Then Exit Sub
    If Len(txtcgc) = 11 Then
        txtcgc = Edita.FormataTexto(txtcgc, Cpf)
    ElseIf Len(txtcgc) = 14 And Mid(txtcgc, 4, 1) <> "." Then
        txtcgc = Edita.FormataTexto(txtcgc, Cgc)
    End If
    Sql = "select tci_nome,tci_im,tci_logradouro,tci_nome_logradouro," & _
    " tci_numero,tci_complemento,tci_bairro,tci_cidade,tci_UF from Tab_Contribuinte" & _
    " where tci_cgc_cpf='" & txtcgc & "' and tci_tsc_cod_sit_cad =1"
    If Bdados.AbreTabela(Sql, rs) Then
        txtContribuinte = rs(0)
        txtIm = rs!TCI_IM
        Endereco = rs!tci_logradouro & " " & rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_COMPLEMENTO
        Bairro = rs!tci_BAIRRO
        Cod_Cidade = rs!tci_cidade
        Uf = rs!tci_UF
        Call txtIm_LostFocus
        If NomeImposto = Imposto.NomeTributo(ttr_IPTU) Or NomeImposto = Imposto.NomeTributo(ttr_ITBI) Then cboIC.SetFocus
    ElseIf Not TributoTaxa Then
        Util.Avisa "Inscrição não cadastrada."
        txtcgc.SetFocus
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtContribuinte_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtDAM_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDAM_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim RsAuxa As VSRecordset
    Dim i As Byte
    Dim JaPago As Boolean
    
    If Trim(txtDam) = "" Then Exit Sub
    
    Sql = "SELECT tdr_data_pagamento,tdr_valor_real_juros,tdr_valor_real_multa FROM TAB_DARM_RECEBIDO WHERE tdr_tgt_cod_pagamento =" & txtDam
    If Bdados.AbreTabela(Sql, rs) Then
        JurosReal = "" & rs!tdr_valor_real_juros
        MultaReal = "" & rs!tdr_valor_real_multa
        DamPago = True
        If Not Confirma("DAM já pago em " & rs(0) & ". Deseja reimprimir?") Then
            Bdados.FechaTabela rs
            txtDam.SetFocus
            Exit Sub
        End If
        JaPago = True
    Else
        DamPago = False
        JurosReal = 0
        MultaReal = 0
        JaPago = False
    End If
    
    Set rs = Conta.BuscaDam(txtDam)
    If Not rs.EOF Then
        txtIm = "" & rs!Inscricao
        txtDam = "" & rs!Documento
        cboTipo.SetarLinha rs.Fields("tgt_tipo"), 1
        If rs.Fields("tgt_tipo") = 1 Then
            txtJuros.Enabled = False
            txtMulta.Enabled = False
            txtTotalImposto.Enabled = False
        Else
            txtJuros.Enabled = True
            txtMulta.Enabled = True
            txtTotalImposto.Enabled = True
        End If
        Sql = "Select tip_cod_imposto,tip_sigla_imposto ,tip_nome_imposto from tab_imposto where tip_cod_imposto = '" & rs!Imposto & "'"
        Dim RsAux As VSRecordset
        If Bdados.AbreTabela(Sql, RsAux) Then
            txtNomeImposto = RsAux!TIP_sigla_IMPOSTO & " - " & RsAux!tip_nome_imposto
            CodImposto = RsAux!tip_cod_imposto
        End If
        CodImpostoOriginal = "" & rs!tgt_tip_cod_imposto_original
        i = InStr(1, txtNomeImposto, "-")
        NomeImposto = Trim(Left(txtNomeImposto, i - 2))
        
        If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNEST) Then
            fra(2).Visible = True
            fraIm.Visible = False
            Sql = "SELECT * FROM TAB_DETALHE_DAM WHERE TDD_TGT_COD_PAGAMENTO = " & txtDam
            If Bdados.AbreTabela(Sql) Then
                txtNotaInicial = "" & Bdados.Tabela(1)
                txtNotaFinal = "" & Bdados.Tabela(2)
                txtTotalNotas = Format("" & Bdados.Tabela(3), "STANDARD")
                txtMaterial = "" & Bdados.Tabela(4)
                txtAliquota = "" & Bdados.Tabela("tdd_Aliquota")
            End If
            
        ElseIf NomeImposto = Imposto.NomeTributo(ttr_IPTU) Or NomeImposto = Imposto.NomeTributo(ttr_ITBI) Then
            fra(2).Visible = False
            fraIm.Visible = True
            txtPeriodo.MaxLength = 4
        ElseIf NomeImposto = Imposto.NomeTributo(ttr_ALVARA) Then
            fra(2).Visible = False
            fraIm.Visible = False
            txtPeriodo.MaxLength = 4
        End If
        CodImposto = RsAux!tip_cod_imposto
        txtParcela = Nvl("" & rs!Parcela, 0)
        
        txtPeriodo.MaxLength = 0
        txtPeriodo = IIf(Len("" & rs!Periodo) = 4, "" & rs!Periodo, Right("" & rs!Periodo, 2) & Left("" & rs!Periodo, 4))
        txtIm_LostFocus
        DoEvents
        txtPeriodo_LostFocus
        DoEvents
        NomeImposto = Trim(Mid(txtNomeImposto, i + 2))
    Else
        Avisa "DAM não encontrado."
        Edita.LimpaCampos Me
        txtDam = ""
        txtDam.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    DoEvents
    Bdados.FechaTabela rs
    Mudou = False
End Sub

Private Sub txtDtVenc_Change()
    Mudou = True
End Sub

Private Sub txtDtVenc_LostFocus()
    On Error Resume Next
    Call txtPeriodo_LostFocus
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    On Error GoTo TRATA
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIm = Imposto.FormataInscricao(txtIm, InscContrib)
    End If
    Dim Sql As String
    Dim rs As VSRecordset
    Dim RsIptu As VSRecordset
    If Trim(txtIm) = "" Then Exit Sub
    Sql = "select * from Tab_Contribuinte where tci_im='" & txtIm & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtContribuinte = "" & rs!tci_nome
        txtEnderecoContrib = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_BAIRRO & ", CEP: " & rs!tci_cep & ", " & rs!tci_cidade & "-" & rs!tci_UF
        If IsNull(rs("tci_tae_cae")) And (NomeImposto <> Imposto.NomeTributo(ttr_IPTU) And NomeImposto <> Imposto.NomeTributo(ttr_ITBI)) Then
            If Not TributoTaxa Then
                Avisa "Contribuinte não sujeito à cobrança de " & NomeImposto & "."
                txtIm.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        txtContribuinte = "" & rs("tci_nome") & ""
        txtcgc = "" & rs("tci_cgc_cpf") & ""
        
        Cod_Atividade = "" & IIf(rs("tci_tae_cae") > 0 And rs!tci_tipo_contribuinte > 1, rs("tci_tae_cae") & Imposto.BuscaNomeCAE(rs("tci_tae_cae") & "") & "", "")
        Endereco = "" & rs("tci_logradouro") & "  " & rs("tci_nome_logradouro") & "," & rs("tci_numero") & " " & rs("tci_complemento")
        Bairro = "" & rs("tci_bairro") & ""
        Cod_Cidade = "" & rs("tci_cidade") & ""
        Cep = "" & rs("tci_cep") & ""
        Uf = "" & rs("tci_uf") & ""
        If Not TributoTaxa Then
            If rs!tci_tipo_recolhimento_iss = 1 And (NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP)) Then
                txtPeriodo.MaxLength = 6
            ElseIf rs!tci_tipo_recolhimento_iss = 2 Then
                txtPeriodo.MaxLength = 4
            End If
        Else
            txtPeriodo.MaxLength = 6
        End If
    End If
    Screen.MousePointer = 0
    Bdados.FechaTabela rs
    Exit Sub
TRATA:
    If Err.Number = 3265 Then
        Resume Next
    End If
End Sub

Private Sub txtImposto_Change()
    txtTotalImposto = txtImposto
End Sub

Private Sub txtParcela_LostFocus()
    If Trim(txtParcela) <> "" Then
        If txtParcela < 1 Then
            Informa "Número de cota inválido."
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
    Dim Conta As New ContaCorrente
    'If Not IsNumeric(txtPeriodo) Then Exit Sub
    If Len(txtPeriodo) = 0 Then Exit Sub
    If Len(txtPeriodo) < 4 Then
        Informa "Período inválido."
        Exit Sub
    End If
    If Len(txtPeriodo) <> 4 And Len(txtPeriodo) <> 6 And Len(txtPeriodo) <> 7 Then
        Avisa "Período inválido."
        txtPeriodo = ""
        Exit Sub
    End If
    
    If NomeImposto = Imposto.NomeTributo(ttr_ISSQN) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNRET) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNSUBST) Or NomeImposto = Imposto.NomeTributo(ttr_ISSQNCOMP) Then
        If Len(txtPeriodo) <> 6 And Len(txtPeriodo) <> 4 Then
            Avisa "Período inválido."
            txtPeriodo = ""
            Exit Sub
        End If
    End If
    If NomeImposto = Imposto.NomeTributo(ttr_ITBI) Then
        If Len(txtPeriodo) <> 6 Then
            Avisa "Período inválido."
            txtPeriodo = ""
            Exit Sub
        End If
    End If
    
    
    If Venc = "0" Then
        Avisa "Não existe formulário definido."
        txtDtVenc = ""
        Exit Sub
    End If
        
    txtPeriodo.MaxLength = 7
    If Len(txtPeriodo) = 6 Then txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
    Sql = "Select tgt_cod_pagamento,tgt_data_vencimento,tgt_valor_tributo,tgt_tim_ic,tgt_valor_juros,tgt_data_geracao,tgt_taxa_expediente," & _
    " tgt_periodo,tgt_tim_ic,tgt_parcela from tab_geracao_tributo where tgt_cod_pagamento=" & txtDam
    Juro = 0
    Multa = 0
    Set rs = Conta.BuscaDam(txtDam)
    If Not rs.EOF Then
        CodPagamento = rs!Documento
        Venc = rs!vencimento
        DataOriginalVence = rs!vencimento
        txtImposto = CDbl(Nvl("" & rs!ValorTributo, 0))
        If Nvl("" & rs!Parcela, 0) <> 0 Then
            Juro = Nvl("" & rs!tgt_valor_juros, 0)
        Else
            Juro = 0
        End If
        If DamPago Then
            Juro = JurosReal
            Multa = MultaReal
        Else
            Juro = Juro + Conta.CalculaValoresJurosAvulsos(CodImposto, rs!Periodo, IIf(CInt(txtParcela) = 0, EtcCreditoTributario, EtcParcelamento), Format(txtDtVenc, "dd/mm/yyyy"), DataOriginalVence, CDbl(Nvl("" & rs!ValorTributo, 0)) + CDbl(Nvl("" & rs!taxa, 0)))
            Multa = Conta.CalculaValoresMultaAvulsos(CodImposto, rs!Periodo, IIf(CInt(txtParcela) = 0, EtcCreditoTributario, EtcParcelamento), Format(txtDtVenc, "dd/mm/yyyy"), DataOriginalVence, CDbl(Nvl("" & rs!ValorTributo, 0)) + CDbl(Nvl("" & rs!taxa, 0)))
        End If
        cboIC = "" & rs!TGT_tim_ic
        If Not IsNull(rs!TGT_tim_ic) Then cboIC_Click
        DtGeracao = "" & rs!tgt_data_geracao
        TaxaServico = Nvl("" & rs!taxa, 0)
        'CodTaxa =
        txtImposto = Edita.FormataTexto(txtImposto, Monetario, True)
        txtJuros = Format(Juro, Const_Monetario)
        txtMulta = Format(Multa, Const_Monetario)
        DoEvents
        Sql = "Select tdd_num_nota_inicial,tdd_num_nota_final,tdd_total_nota,tdd_total_material_reducao,tdd_obs from Tab_Detalhe_Dam where tdd_tgt_cod_pagamento=" & Nvl(txtDam, 0)
        If Bdados.AbreTabela(Sql, rs) Then
                txtNotaInicial = rs!tdd_num_nota_inicial
                txtNotaFinal = rs!tdd_num_nota_final
                txtTotalNotas = Format(rs!tdd_total_nota, Const_Monetario)
                BaseDeCalculo = txtTotalNotas
                txtMaterial = Format(rs!tdd_total_material_reducao, Const_Monetario)
                txtSaldo = Format(CDbl(txtTotalNotas) - CDbl(txtMaterial), Const_Monetario)
                txtObservacao = "" & rs!tdd_obs
        End If
        txtTotalImposto = CDbl(Nvl(txtImposto, 0)) + CDbl(Nvl(txtJuros, 0)) + CDbl(Nvl(txtMulta, 0))
    Else
        Avisa ("DAM ainda não emitido. Use o formulário de Impressão de DAM.")
        Screen.MousePointer = 0
        txtDam.SetFocus
        Exit Sub
    End If
    Venc = Format(Venc, "dd/mm/yyyy")
    Exercicio = txtPeriodo
    If DamPago Then
        txtDtVenc = Venc
    ElseIf Not Mudou Then
        txtDtVenc = ""
    End If
    Bdados.FechaTabela rs
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
    Dim pos As Integer
    String_Taxas = ""
    Total_Taxas = 0
    For i = 1 To Grdtaxas.ListItems.Count
        If Grdtaxas.ListItems(i).Checked Then
            pos = InStr(Grdtaxas.ListItems(i).SubItems(1), "-") - 1
            If String_Taxas = "" Then
                String_Taxas = String_Taxas & " [ " & Left(Grdtaxas.ListItems(i).SubItems(1), pos) & " ]" & " - " & Format(Grdtaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            Else
                String_Taxas = String_Taxas & ", [ " & Left(Grdtaxas.ListItems(i).SubItems(1), pos) & " ]" & " - " & Format(Grdtaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            End If
            Total_Taxas = Total_Taxas + CCur(Grdtaxas.ListItems(i).SubItems(2))
        End If
    Next
End Sub
