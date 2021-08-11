VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TARR201 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VS"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   46
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TARR201.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   1138
      Icone           =   "TARR201.frx":2123
   End
   Begin VB.CheckBox chkNovoLanca 
      Caption         =   "Criar novo lancamento"
      Height          =   225
      Left            =   90
      MaskColor       =   &H8000000F&
      TabIndex        =   17
      Top             =   6000
      Width           =   2115
   End
   Begin Threed.SSFrame fra 
      Height          =   825
      Index           =   3
      Left            =   60
      TabIndex        =   30
      Top             =   720
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   1455
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
      Begin VB.TextBox txtNumLote 
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
         Left            =   2220
         TabIndex        =   0
         Tag             =   "Nº do Lote"
         Top             =   330
         Width           =   1965
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   15
         Left            =   375
         TabIndex        =   31
         Top             =   330
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   476
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
         Caption         =   "NÚMERO DO LOTE"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSFrame fra 
         Height          =   660
         Index           =   4
         Left            =   4665
         TabIndex        =   32
         Top             =   60
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   1164
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
         Begin VB.TextBox txtValorAberto 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   3945
            TabIndex        =   37
            Tag             =   "Nº do Lote"
            Top             =   255
            Width           =   1275
         End
         Begin VB.TextBox txtValorDig 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   2130
            TabIndex        =   35
            Tag             =   "Nº do Lote"
            Top             =   255
            Width           =   1305
         End
         Begin VB.TextBox txtValoLote 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   240
            TabIndex        =   33
            Tag             =   "Nº do Lote"
            Top             =   255
            Width           =   1305
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   16
            Left            =   225
            TabIndex        =   34
            Top             =   45
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Valor do Lote"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   17
            Left            =   2100
            TabIndex        =   36
            Top             =   45
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   397
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
            Caption         =   "Valor Digitado"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   240
            Index           =   18
            Left            =   3930
            TabIndex        =   38
            Top             =   30
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   423
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
            Caption         =   "Valor Aberto"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
   End
   Begin Threed.SSFrame fraContrib 
      Height          =   1680
      Left            =   60
      TabIndex        =   23
      Top             =   1575
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   2963
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
         Left            =   5160
         MaxLength       =   20
         TabIndex        =   3
         Top             =   600
         Width           =   2265
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
         Left            =   2220
         MaxLength       =   15
         TabIndex        =   1
         Top             =   210
         Width           =   2265
      End
      Begin VB.TextBox txtSeq 
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
         Left            =   9270
         MaxLength       =   14
         TabIndex        =   39
         Top             =   600
         Width           =   660
      End
      Begin VB.TextBox txtInscricao 
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
         Left            =   2220
         MaxLength       =   15
         TabIndex        =   2
         Top             =   600
         Width           =   1905
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   5
         Left            =   1290
         TabIndex        =   24
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Contribuinte"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   10
         Left            =   7890
         TabIndex        =   40
         Top             =   645
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   397
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
         Caption         =   "Sequência no Lote"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   1650
         TabIndex        =   43
         Top             =   960
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   556
         Caption         =   "Razão"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   315
         Left            =   1380
         TabIndex        =   44
         Top             =   1290
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   556
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   1110
         TabIndex        =   45
         Top             =   255
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   397
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
         Caption         =   "No Documento"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   4650
         TabIndex        =   47
         Top             =   645
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Imóvel"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   4155
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   615
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   7455
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
   End
   Begin Threed.SSFrame fraImposto 
      Height          =   1650
      Left            =   60
      TabIndex        =   21
      Top             =   3270
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   2910
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
      Begin VTOcx.cboVISUAL cboTributo 
         Height          =   315
         Left            =   570
         TabIndex        =   5
         Top             =   540
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
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
         Left            =   2220
         MaxLength       =   14
         TabIndex        =   20
         Tag             =   "Data Pagamento"
         Top             =   1260
         Width           =   1245
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
         Left            =   8520
         MaxLength       =   14
         TabIndex        =   9
         Top             =   1260
         Width           =   495
      End
      Begin VB.TextBox txtDtVencimento 
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
         Left            =   8520
         MaxLength       =   14
         TabIndex        =   8
         Top             =   885
         Width           =   1305
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
         Left            =   2220
         MaxLength       =   14
         TabIndex        =   7
         Top             =   885
         Width           =   1245
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   570
         TabIndex        =   26
         Top             =   930
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   397
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
         Caption         =   "Período Referência"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   6
         Left            =   7110
         TabIndex        =   27
         Top             =   930
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   397
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
         Caption         =   "Data Vencimento"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   7
         Left            =   7365
         TabIndex        =   28
         Top             =   1305
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   397
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
         Caption         =   "Nº da Parcela"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   13
         Left            =   825
         TabIndex        =   29
         Top             =   1305
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Data Pagamento"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtOrigem 
         Height          =   315
         Left            =   7170
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   510
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         Caption         =   "Código Origem"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtCodTributo 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Tag             =   "Valor de Deduções"
         Top             =   180
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   529
         Caption         =   "Cod. Tributo"
         Text            =   ""
         Restricao       =   2
         AutoTAB         =   -1  'True
      End
   End
   Begin Threed.SSFrame fraPag 
      Height          =   1005
      Left            =   60
      TabIndex        =   25
      Top             =   4935
      Width           =   10080
      _ExtentX        =   17780
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
      Caption         =   "Informações do Pagamento"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VTOcx.txtVISUAL txtValorOriginal 
         Height          =   300
         Left            =   285
         TabIndex        =   10
         Tag             =   "Valor Original do DAM"
         Top             =   210
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   529
         Caption         =   "Valor Original do DAM"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtValorPago 
         Height          =   300
         Left            =   7095
         TabIndex        =   15
         Tag             =   "Valor Total Pago"
         Top             =   600
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   529
         Caption         =   "Valor Total Pago"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtDeducao 
         Height          =   300
         Left            =   6885
         TabIndex        =   14
         Tag             =   "Valor de Deduções"
         Top             =   255
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   529
         Caption         =   "Valor de Deduções"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtTaxa 
         Height          =   300
         Left            =   3795
         TabIndex        =   13
         Tag             =   "Valor de Taxa"
         Top             =   585
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   529
         Caption         =   "Valor de Taxa"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtValorPagoMulta 
         Height          =   300
         Left            =   3675
         TabIndex        =   12
         Tag             =   "Valor de Multas"
         Top             =   240
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         Caption         =   "Valor de Multas"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtValorPagoJuros 
         Height          =   300
         Left            =   945
         TabIndex        =   11
         Tag             =   "Valor de Juros"
         Top             =   585
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   529
         Caption         =   "Valor de Juros"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AutoTAB         =   -1  'True
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7440
      Top             =   60
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   345
      Left            =   4470
      TabIndex        =   22
      Top             =   750
      Width           =   855
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   9000
      TabIndex        =   19
      Top             =   5985
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
      Left            =   7845
      TabIndex        =   16
      Top             =   5985
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdCancelar 
      Height          =   375
      Left            =   6690
      TabIndex        =   18
      Top             =   5985
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.grdVISUAL lstDocs 
      Height          =   2130
      Left            =   45
      TabIndex        =   42
      Top             =   6420
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   3757
      Caption         =   "Documentos do Lote"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
   Begin VB.Menu munArr 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuEstornoPago 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "TARR201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Vez As Byte
Private DadosImposto(1 To 6) As String
Private Pagamento(1 To 5) As Double
Private Imposto As New VSImposto
Private DtGeracao  As String
Dim ValorOriginalImpostoLancado As Double
Dim ValorOriginalTaxaLancado As Double
Dim Inscricao As String
Dim bCotasObrigacao As Boolean
Dim TipoInscricao As Integer
Dim Tipo  As TipoInscricaoObrigacao
Option Explicit

Private Function ValidaDAM(DAM) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsAux As VSRecordset
    Dim ValImposto As Double
    Sql = "SElect tdr_data_pagamento,tdr_valor_real_pago,tdr_im,tci_nome from tab_darm_recebido,tab_contribuinte where tdr_tgt_cod_pagamento=" & DAM & " and tdr_sit_pago <> 2 and tdr_im = tci_im"
    If Bdados.AbreTabela(Sql, Rs) Then
        Sql = "Select sum(tdr_valor_real_pago) from tab_darm_recebido where tdr_tgt_cod_pagamento in ( select TTD_TGT_COD_PAGAMENTO_TAXA from tab_taxa_dam where TTD_TGT_COD_PAGAMENTO = " & DAM & ")"
        If Bdados.AbreTabela(Sql, RsAux) Then
            ValImposto = ValImposto + CDbl(Nvl("" & RsAux(0), 0))
        End If
        Bdados.FechaTabela RsAux
        Informa "DAM já recebida em " & Rs(0) & " no valor de R$" & Format(ValImposto, Const_Monetario) & ". Contribuinte " & Rs!TDR_im & " - " & Rs!tci_nome
        Bdados.FechaTabela Rs
        Exit Function
    End If
End Function

Private Sub AtualizaLote(ByRef Cancela As Boolean)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RSDam As VSRecordset
    Cancela = False
    If Trim(txtNumLote) = "" Then
        lstDocs.Preencher Bdados, ""
        Exit Sub
    End If
    Sql = "Select TLP_VALOR_ARRECADADO,TLP_SITUACAO_LOTE,TLP_DATA_ARRECADACAO " & _
        "from TAB_LOTE_PAGAMENTO  where TLP_COD_LOTE =" & txtNumLote
    If Bdados.AbreTabela(Sql, Rs) Then
        txtValoLote = Format(Rs!TLP_VALOR_ARRECADADO, Const_Monetario)
        Sql = "select sum(tdr_valor_real_pago) as Digitado from Tab_Darm_Recebido where  TDR_TLP_COD_LOTE =" & txtNumLote & " and tdr_sit_pago <> 2    " 'and tdr_tgt_cod_pagamento_vinculado =tdr_tgt_cod_pagamento"
        If Bdados.AbreTabela(Sql, RSDam) Then
            txtValorDig = Format(IIf(IsNull(RSDam!Digitado), 0, RSDam!Digitado), Const_Monetario)
            txtValorAberto = Format(CDbl(txtValoLote) - IIf(IsNull(RSDam!Digitado), 0, RSDam!Digitado), Const_Monetario)
            TxtDtPagamento = Rs!TLP_DATA_ARRECADACAO
        End If
    Else
        Informa "Lote inexistente"
        Cancela = True
        txtNumLote.SetFocus
        Bdados.FechaTabela Rs
        Exit Sub
    End If
    If CDbl(Nvl(txtValorAberto, 0)) <= 0 Then
        Informa "Lote fechado."
        
        'txtNumLote.SetFocus
        txtNumLote.SelStart = 0
        txtNumLote.SelLength = Len(Trim(txtNumLote))
        If Rs!TLP_SITUACAO_LOTE <> 1 Then
            Bdados.AtualizaDados "TAB_LOTE_PAGAMENTO", Bdados.PreparaValor(1), "TLP_SITUACAO_LOTE", "TLP_COD_LOTE =" & txtNumLote
        End If
        lstDocs.Mensagem = ""
        DoEvents
    Else
        If Rs!TLP_SITUACAO_LOTE <> 0 Then
            Bdados.AtualizaDados "TAB_LOTE_PAGAMENTO", Bdados.PreparaValor(0), "TLP_SITUACAO_LOTE", "TLP_COD_LOTE =" & txtNumLote
        End If
    End If
    'BUSCANDO VALORES RESGATADO(PAGOS) - 2ª grade
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "SELECT tdr_tgt_cod_pagamento as [Conta],TDR_INSCRICAO as Contribuinte, " & _
                "tip_sigla_imposto as Tributo," & _
                Bdados.Converte("tdr_valor_real_pago", TCDuplo) & " as [Vl Pago],tdr_periodo as Periodo, TDR_PARCELA AS Parcela," & _
                "TDR_SEQUENCIA_DAM_LOTE Seq from tab_darm_recebido,tab_imposto" & _
                " where  tdr_tip_cod_imposto = tip_cod_imposto and " & _
                "tdr_sit_pago <> 2 and tdr_tlp_cod_lote = " & txtNumLote & _
                " order by TDR_SEQUENCIA_DAM_LOTE desc"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "SELECT tdr_tgt_cod_pagamento as Conta,TDR_INSCRICAO as Contribuinte, " & _
                "tip_sigla_imposto as Tributo," & _
                " to_number(tdr_valor_real_pago,'999999.99') as Vl_Pago,tdr_periodo as Periodo, TDR_PARCELA AS Parcela," & _
                "TDR_SEQUENCIA_DAM_LOTE Seq from tab_darm_recebido,tab_imposto" & _
                " where  tdr_tip_cod_imposto = tip_cod_imposto and " & _
                "tdr_sit_pago <> 2 and tdr_tlp_cod_lote = " & txtNumLote & _
                " order by TDR_SEQUENCIA_DAM_LOTE desc"
    End If
    lstDocs.Preencher Bdados, Sql, 900, 2000, 1200, 1200, 900, 600
    If lstDocs.ListItems.Count > 0 Then lstDocs.Mensagem = "Total Digitado: R$ " & Format(lstDocs.Colunas(4).Soma, Const_Monetario)
    Bdados.FechaTabela Rs
End Sub

Private Function GravaDados() As Boolean
    On Error GoTo trata
    Dim Arrec As New Arrecadacao
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Inscricao As String
    Dim Lote As String
    Dim Conta As New ContaCorrente
    'NUM CORRELATIVO
    
'    txtSeq = Conta.GeraCodPagamento("LOTE")
    
    Sql = "Select max(TDR_SEQUENCIA_DAM_LOTE) + 1 from tab_darm_recebido where " & _
            " TDR_TLP_COD_LOTE=" & txtNumLote
    If Bdados.AbreTabela(Sql, Rs) Then
        txtSeq = Format(Nvl("" & Rs(0), 1), "000000")
    Else
        txtSeq = "000001"
    End If

    Inscricao = IIf(Trim(txtInscricao) = "", txtImovel, txtInscricao)
'    If Nvl(Trim(txtParcela), 0) = 0 Then
        txtPeriodo = IIf(Len(txtPeriodo) = 4, txtPeriodo, IIf(Len(txtPeriodo) = 6, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2)))
        txtInscricao = Trim(txtInscricao)
        If Trim(txtInscricao) = "" Then
            TipoInscricao = 1
        Else
            TipoInscricao = 2
        End If
        
        GravaDados = Arrec.GravaPagamento(Inscricao, TipoInscricao, cboTributo.Coluna(0).Valor, txtPeriodo, txtDtVencimento, _
                              TxtDtPagamento, CDbl(txtValorOriginal), CDbl(txtTaxa), CDbl(txtDeducao), _
                 CDbl(txtValorPagoJuros), 0, CDbl(txtValorPagoMulta), CDbl(txtNumLote), txtSeq, Nvl(txtParcela, 0), txtDAM, IIf(chkNovoLanca.Value = 1, etsCriaNova, etsNaoSubstitui), txtValorPago.Text, Nvl(Trim(txtOrigem), 0), etsCreditoPago)
'    Else
'        GravaDados = Arrec.GravaPagamento_Cotas_Obrigacao(Inscricao, cboTributo.Coluna(0).VALOR, txtPeriodo, txtDtVencimento, _
'                            TxtDtPagamento, txtValorOriginal, txtTaxa, txtDeducao, _
'                 txtValorPagoJuros, txtValorPagoMulta, txtNumLote, txtSeq, CInt(txtParcela), txtDAM, IIf(chkNovoLanca.Value = 1, etsCriaNova, etsNaoSubstitui), txtValorPago.Text, Trim(txtOrigem), etsCreditoPago)
'    End If
    Exit Function
trata:
    Avisa Err.Description
    Exit Function
    Resume
End Function

Private Sub cboTributo_Click()
    txtCodTributo = cboTributo.Coluna(0).Valor
End Sub

Private Sub cmdCancelar_Click()
    Dim Lote As Double
    Lote = Nvl(txtNumLote, 0)
    Edita.LimpaCampos Me
    txtNumLote = ""
    fraContrib.Enabled = True
    fraContrib.Enabled = True
    cboTributo.Enabled = True
    txtPeriodo.Enabled = True
    txtDtVencimento.Enabled = True
    txtParcela.Enabled = True
    Call txtNumLote_Validate(True)
    If CDbl(Nvl(txtValorAberto, 0)) > 0 Then
        txtInscricao.SetFocus
    Else
        txtNumLote.SetFocus
    End If
    Vez = 0
    TipoInscricao = 0
    txtInscricao.Enabled = True
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscricao
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Lote As String
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    If CLng(CDbl(Nvl(Trim(txtValorPago), 0)) / CDbl(CDbl(Nvl(Trim(txtValorOriginal), 0)) + CDbl(Nvl(Trim(txtValorPagoJuros), 0)) + CDbl(Nvl(Trim(txtValorPagoMulta), 0)) + CDbl(Nvl(Trim(txtTaxa), 0)) - CDbl(Nvl(Trim(txtDeducao), 0)))) <> 1 Then
        If Not Confirma("Valor total não confere. Confirmar dados e salvar? " & CDbl(Nvl(Trim(txtValorPago), 0)) & " diferente de " & CDbl(CDbl(Nvl(Trim(txtValorOriginal), 0)) + CDbl(Nvl(Trim(txtValorPagoJuros), 0)) + CDbl(Nvl(Trim(txtValorPagoMulta), 0)) + CDbl(Nvl(Trim(txtTaxa), 0)) - CDbl(Nvl(Trim(txtDeducao), 0)))) Then
            txtValorPago.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If GravaDados Then
         Util.Informa "Dados Gravados com Segurança."
        Lote = txtNumLote
        Call cmdCancelar_Click
        txtNumLote = Lote
        Call txtNumLote_Validate(True)
        If CDbl(Nvl(Trim(txtValorAberto), 0)) > 0 Then SendKeys "{TAB}"
    Else
        Util.Informa "Dados não foram Gravados "
    End If
    TipoInscricao = 0
    Screen.MousePointer = 0
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

'Private Sub cmdVISUAL1_Click()
'    Dim sql As String
'    Dim rs As VSRecordset
'    Dim Rs_IC As VSRecordset
'    Dim ca As String
'    Dim va As String
'    Dim con As String
'
'    txtInscricao = 1
'    sql = "SELECT * "
'    sql = sql & " From TAB_DARM_RECEBIDO"
'    sql = sql & " WHERE   tdr_tip_cod_imposto = '11120200' and len(tdr_inscricao) = 11"
'    If Bdados.AbreTabela(sql, rs) Then
'        Do Until rs.EOF
'            If rs.Fields("tdr_im") <> "" Then
'                va = Bdados.PreparaValor(Trim(rs.Fields("tdr_IM")))
'                ca = "tdr_inscricao"
'                con = "tdr_im = " & Bdados.Converte(rs.Fields("tdr_im"), tctexto)
'                DoEvents
'                DoEvents
'                DoEvents
'                Bdados.GravaDados "TAB_DARM_RECEBIDO", va, ca, con
'                DoEvents
'                DoEvents
'                DoEvents
'                txtInscricao = txtInscricao + 1
'            End If
'            rs.MoveNext
'        Loop
'    End If
'End Sub

Private Sub Form_Activate()
    txtNumLote.SetFocus
    AtualizaLote False
End Sub

Private Sub lstDocs_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not (lstDocs.SelectedItem Is Nothing) Then
    If Button = 2 Then
        mnuEstornoPago.Caption = "Estornar DAM " & lstDocs.SelectedItem
        Me.PopupMenu munArr
    End If
End If
End Sub

Private Sub mnuEstornoPago_Click()
    TARR301.txtDAM = lstDocs.SelectedItem
    SendKeys "{tab}"
    TARR301.Show
    TARR301.txtMotivo.SetFocus
End Sub

Private Sub Form_Load()
    Dim Obrig As New obrigacao
    cabVisual.Exibir Bdados, Me.Name, App.Path
    Obrig.PreencheComboTributo cboTributo, False
End Sub

Private Sub txtCodTributo_LostFocus()
    cboTributo.SetarLinha txtCodTributo, 0
End Sub

Private Sub txtDAM_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtDAM_LostFocus()
    Dim Conta As New ContaCorrente
    Dim Sql As New ContaCorrente
    Dim Rs As VSRecordset
    If Trim(txtDAM) = "" Then Exit Sub
    If Not IsNumeric(Trim(txtDAM)) Then Exit Sub
    Set Rs = Conta.BuscaDam(txtDAM)
    If Not Rs.EOF Then
        bCotasObrigacao = False
'        Sql = "Select toc_inscricao as Inscricao,toc_tip_cod_imposto as Imposto, toc_periodo as Periodo," & _
            "toc_data_vencimento as Vencimento, toc_parcela as Parcela, toc_valor_tributo as ValorTributo," & _
            "toc_valor_juros as Juros, toc_valor_multa as Multa, toc_taxa_expediente as Taxa " & _
            "from tab_obrigacao_contribuinte where toc_cod_obrigacao=" & Documento

        If Trim("" & Rs!Inscricao) = "" Then
            Avisa "DAM corrompido. Digite-o manualmente"
            Screen.MousePointer = 0
            Exit Sub
        End If
        TipoInscricao = Nvl("" & Rs.Fields("TipoInscricao"), 0)
        If TipoInscricao = 2 Then
            txtInscricao = Trim("" & Rs!Inscricao)
        Else
            txtImovel = Trim("" & Rs!Inscricao)
        End If
        txtInscricao_LostFocus
        TipoInscricao = Nvl("" & Rs!TipoInscricao, 1)
        If Trim(txtImovel) = "" And Trim(txtInscricao) = "" Then
            TipoInscricao = 2
            txtInscricao = Trim("" & Rs!Inscricao)
            txtInscricao_LostFocus
        End If
        txtPeriodo = Nvl("" & Rs!Periodo, 0)
        cboTributo.SetarLinha Nvl("" & Rs!Imposto, 0), 0
        txtDtVencimento = Nvl("" & Rs!Vencimento, 0)
        txtParcela = Nvl("" & Rs!Parcela, 0)
        txtValorOriginal = Nvl("" & Rs!ValorTributo, 0)
        txtValorPagoJuros = Nvl("" & Rs!Juros, 0)
        txtValorPagoMulta = Nvl("" & Rs!Multa, 0)
        txtCodTributo = "" & Rs!Imposto
        If Temp.PegaParametro(Bdados, "MUNICIPIO") Then
            txtTaxa = CDbl(Nvl("" & Rs!Taxa, 0))
        End If
        
        txtValorPago = CDbl(txtValorOriginal) + CDbl(txtValorPagoJuros) + CDbl(txtValorPagoMulta) + CDbl(txtTaxa)
        txtDeducao = "0,00"
        fraContrib.Enabled = False
        cboTributo.Enabled = False
        txtPeriodo.Enabled = False
        txtDtVencimento.Enabled = False
        txtParcela.Enabled = False
        txtValorOriginal.SetFocus
    Else
        Avisa "Número inválido."
        txtDAM = ""
    End If
End Sub


Private Sub TxtDtPagamento_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtDtVencimento_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtDtVencimento_LostFocus()
    txtDtVencimento = Edita.FormataTexto(txtDtVencimento, Data)
    If Trim(txtDtVencimento) = "" Then Exit Sub
    If txtDtVencimento.Enabled = True Then
        If Not IsDate(txtDtVencimento) Then
            Avisa "Data inválida."
            txtDtVencimento.SetFocus
        End If
    End If
End Sub

Private Sub txtImovel_LostFocus()
    If Trim(txtImovel) = "" Then Exit Sub
    TipoInscricao = 1
    txtInscricao_LostFocus
End Sub

Private Sub txtInscricao_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtInscricao_LostFocus()
    If TipoInscricao = 2 Or TipoInscricao = 0 Then
        txtInscricao = BuscaContribuinte(txtInscricao, txtrazao, txtEndereco, , etiContribuinte)
    Else
        txtImovel = BuscaContribuinte(txtImovel, txtrazao, txtEndereco, , etiImovel)
    End If
End Sub

Private Sub txtNumLote_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtNumLote_Validate(Cancel As Boolean)
    If Trim(txtNumLote) <> "" Then
        Screen.MousePointer = 11
        Call AtualizaLote(Cancel)
        Screen.MousePointer = 0
    Else
        txtValoLote = ""
        txtValorAberto = ""
        txtValorDig = ""
        lstDocs.Preencher Bdados, ""
    End If
End Sub

Private Sub txtParcela_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtTotalDarm_KeyPress(KeyAscii As Integer)
    txtValorOriginal = Edita.FormataTexto(txtValorOriginal, Monetario, True)
End Sub

Private Sub txtPeriodo_LostFocus()
    Dim Obrig As New obrigacao
    If Trim(txtPeriodo) = "" Then Exit Sub
    
    txtPeriodo.MaxLength = txtPeriodo.MaxLength + 1
    If Len(txtPeriodo) = 6 Then txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
    txtPeriodo.MaxLength = txtPeriodo.MaxLength - 1
End Sub

Private Sub txtValorPago_LostFocus()
    On Error Resume Next
    txtValorPago = Edita.FormataTexto(txtValorPago, Monetario, True)

    If CDbl(Nvl(txtValorAberto, 0)) < CDbl(Nvl(txtValorPago, 0)) Then
        Informa "Valor pago maior que valor em aberto do lote."
        txtValorPago.SetFocus
        Exit Sub
    End If
    
End Sub
