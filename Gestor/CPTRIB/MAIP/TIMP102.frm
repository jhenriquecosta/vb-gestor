VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TIMP102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   99
      Top             =   7455
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   8490
         TabIndex        =   39
         Top             =   135
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   9555
         TabIndex        =   40
         Top             =   135
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin ActiveTabs.SSActiveTabs TabTributo 
      Height          =   6735
      Left            =   15
      TabIndex        =   42
      Top             =   645
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   11880
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "TIMP102.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   45
         Top             =   30
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   11192
         _Version        =   131082
         TabGuid         =   "TIMP102.frx":00C4
         Begin Threed.SSFrame fra 
            Height          =   615
            Index           =   1
            Left            =   75
            TabIndex        =   46
            Top             =   30
            Width           =   10500
            _ExtentX        =   18521
            _ExtentY        =   1085
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
            Caption         =   "Período"
            ShadowStyle     =   1
            Begin VB.TextBox txtNomeImposto 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   315
               Left            =   2685
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   210
               Width           =   5895
            End
            Begin VB.TextBox txtAnoImposto 
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
               Left            =   9960
               MaxLength       =   4
               TabIndex        =   1
               Top             =   210
               Width           =   465
            End
            Begin VB.TextBox txtCodImposto 
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
               MaxLength       =   8
               TabIndex        =   0
               Tag             =   "Código"
               Top             =   210
               Width           =   975
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   9
               Left            =   90
               TabIndex        =   47
               Top             =   240
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
               Caption         =   "Código do Tributo:"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   3
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   12
               Left            =   8670
               TabIndex        =   48
               Top             =   240
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
               Caption         =   "Ano Imposto:"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   3
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame fra 
            Height          =   5550
            Index           =   0
            Left            =   75
            TabIndex        =   50
            Top             =   675
            Width           =   10500
            _ExtentX        =   18521
            _ExtentY        =   9790
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
            Begin VB.Frame Frame1 
               Caption         =   "Descontos À vista"
               Height          =   825
               Left            =   30
               TabIndex        =   56
               Top             =   4560
               Width           =   6345
               Begin VB.TextBox txtDesconto 
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
                  Left            =   2235
                  MaxLength       =   8
                  TabIndex        =   31
                  Tag             =   "Redutor"
                  Top             =   420
                  Width           =   765
               End
               Begin VB.TextBox txtDescontoMulta 
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
                  Left            =   3195
                  MaxLength       =   8
                  TabIndex        =   32
                  Tag             =   "Redutor"
                  Top             =   435
                  Width           =   765
               End
               Begin VB.TextBox txtDescontoJuros 
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
                  Left            =   4125
                  MaxLength       =   8
                  TabIndex        =   33
                  Tag             =   "Redutor"
                  Top             =   420
                  Width           =   765
               End
               Begin VB.TextBox txtDescontoCorrecao 
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
                  Left            =   5040
                  MaxLength       =   8
                  TabIndex        =   34
                  Tag             =   "Redutor"
                  Top             =   420
                  Width           =   765
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   2
                  Left            =   2265
                  TabIndex        =   57
                  Top             =   180
                  Width           =   615
                  _ExtentX        =   1085
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
                  Caption         =   "Tributo"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   5
                  Left            =   3240
                  TabIndex        =   58
                  Top             =   180
                  Width           =   465
                  _ExtentX        =   820
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
                  Caption         =   "Multa"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   22
                  Left            =   4125
                  TabIndex        =   59
                  Top             =   180
                  Width           =   495
                  _ExtentX        =   873
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
                  Caption         =   "Juros"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   23
                  Left            =   5040
                  TabIndex        =   60
                  Top             =   180
                  Width           =   945
                  _ExtentX        =   1667
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
                  Caption         =   "Atualizacão"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin VTOcx.cboVISUAL cboDescontoAteVencimento 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   30
                  Top             =   435
                  Width           =   2115
                  _ExtentX        =   3731
                  _ExtentY        =   556
                  Caption         =   ""
                  Text            =   ""
                  AutoFocaliza    =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   8
                  Left            =   90
                  TabIndex        =   61
                  Top             =   195
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
                  Caption         =   "Prazo"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Descontos em Parcelamento"
               Height          =   825
               Left            =   6570
               TabIndex        =   51
               Top             =   4560
               Width           =   3810
               Begin VB.TextBox txtDescontoCorrecaoParc 
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
                  Left            =   2490
                  MaxLength       =   8
                  TabIndex        =   38
                  Tag             =   "Redutor"
                  Top             =   420
                  Width           =   765
               End
               Begin VB.TextBox txtDescontoJurosParc 
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
                  MaxLength       =   8
                  TabIndex        =   37
                  Tag             =   "Redutor"
                  Top             =   420
                  Width           =   765
               End
               Begin VB.TextBox txtDescontoMultaParc 
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
                  Left            =   870
                  MaxLength       =   8
                  TabIndex        =   36
                  Tag             =   "Redutor"
                  Top             =   420
                  Width           =   765
               End
               Begin VB.TextBox txtDescontoTributoParc 
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
                  MaxLength       =   8
                  TabIndex        =   35
                  Tag             =   "Redutor"
                  Top             =   420
                  Width           =   765
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   24
                  Left            =   90
                  TabIndex        =   52
                  Top             =   180
                  Width           =   615
                  _ExtentX        =   1085
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
                  Caption         =   "Tributo"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   25
                  Left            =   900
                  TabIndex        =   53
                  Top             =   180
                  Width           =   465
                  _ExtentX        =   820
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
                  Caption         =   "Multa"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   26
                  Left            =   1680
                  TabIndex        =   54
                  Top             =   180
                  Width           =   495
                  _ExtentX        =   873
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
                  Caption         =   "Juros"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   27
                  Left            =   2490
                  TabIndex        =   55
                  Top             =   180
                  Width           =   945
                  _ExtentX        =   1667
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
                  Caption         =   "Atualizacão"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin VTOcx.cmdVISUAL cmdBuscar 
                  Height          =   300
                  Left            =   3330
                  TabIndex        =   101
                  ToolTipText     =   "Definir Faixas"
                  Top             =   450
                  Width           =   390
                  _ExtentX        =   688
                  _ExtentY        =   529
                  Caption         =   "..."
                  CorBorda        =   8421504
                  CorFrente       =   16384
               End
            End
            Begin VB.ComboBox cboTaxaEmissao 
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
               ItemData        =   "TIMP102.frx":00EC
               Left            =   8730
               List            =   "TIMP102.frx":00F6
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Tag             =   "Taxa Emissão"
               Top             =   450
               Width           =   1755
            End
            Begin VB.TextBox txtLei 
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
               Left            =   45
               MaxLength       =   200
               TabIndex        =   29
               Top             =   4215
               Width           =   10230
            End
            Begin VTOcx.cmdVISUAL cmdImpostosRelacionados 
               Height          =   240
               Left            =   6195
               TabIndex        =   62
               Top             =   5610
               Visible         =   0   'False
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   423
               Caption         =   "Tributos Relacionados"
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin Threed.SSFrame SSFrame1 
               Height          =   795
               Left            =   5430
               TabIndex        =   63
               Top             =   2610
               Width           =   4995
               _ExtentX        =   8811
               _ExtentY        =   1402
               _Version        =   196610
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Receitas diversas"
               Begin VTOcx.cboVISUAL cboReceitaAMais 
                  Height          =   510
                  Left            =   90
                  TabIndex        =   25
                  Top             =   180
                  Width           =   2205
                  _ExtentX        =   3889
                  _ExtentY        =   900
                  Caption         =   "Valor a Mais"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboReceitaAMenos 
                  Height          =   510
                  Left            =   2535
                  TabIndex        =   26
                  Top             =   180
                  Width           =   2445
                  _ExtentX        =   4313
                  _ExtentY        =   900
                  Caption         =   "Valor a Menos"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
            End
            Begin Threed.SSFrame fra 
               Height          =   825
               Index           =   2
               Left            =   60
               TabIndex        =   64
               Top             =   870
               Width           =   10380
               _ExtentX        =   18309
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
               Caption         =   "Tipo"
               ShadowStyle     =   1
               Begin VB.ComboBox cboTipoTributo 
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
                  ItemData        =   "TIMP102.frx":0104
                  Left            =   1740
                  List            =   "TIMP102.frx":0106
                  Style           =   2  'Dropdown List
                  TabIndex        =   9
                  Tag             =   "Logradouro"
                  Top             =   390
                  Width           =   1830
               End
               Begin VB.ComboBox cboTipoContrib 
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
                  ItemData        =   "TIMP102.frx":0108
                  Left            =   75
                  List            =   "TIMP102.frx":010A
                  Style           =   2  'Dropdown List
                  TabIndex        =   8
                  Tag             =   "Logradouro"
                  Top             =   390
                  Width           =   1650
               End
               Begin VB.ComboBox cboTipoInsc 
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
                  ItemData        =   "TIMP102.frx":010C
                  Left            =   3570
                  List            =   "TIMP102.frx":010E
                  Style           =   2  'Dropdown List
                  TabIndex        =   10
                  Tag             =   "Logradouro"
                  Top             =   390
                  Width           =   2160
               End
               Begin VB.ComboBox cboTipoIC 
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
                  ItemData        =   "TIMP102.frx":0110
                  Left            =   5745
                  List            =   "TIMP102.frx":0112
                  Style           =   2  'Dropdown List
                  TabIndex        =   11
                  Tag             =   "Logradouro"
                  Top             =   390
                  Width           =   2190
               End
               Begin VTOcx.cboVISUAL cboObrigacao 
                  Height          =   510
                  Left            =   7950
                  TabIndex        =   12
                  Top             =   180
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   900
                  Caption         =   "Obrigacão"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   0
                  Left            =   90
                  TabIndex        =   65
                  Top             =   180
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   476
                  _Version        =   196610
                  ForeColor       =   0
                  Windowless      =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Contribuinte"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   1
                  Left            =   1740
                  TabIndex        =   66
                  Top             =   180
                  Width           =   600
                  _ExtentX        =   1058
                  _ExtentY        =   476
                  _Version        =   196610
                  ForeColor       =   0
                  Windowless      =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Tributo"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   10
                  Left            =   3570
                  TabIndex        =   67
                  Top             =   180
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   476
                  _Version        =   196610
                  ForeColor       =   0
                  Windowless      =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Inscrição Municipal"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   21
                  Left            =   5745
                  TabIndex        =   75
                  Top             =   150
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   476
                  _Version        =   196610
                  ForeColor       =   0
                  Windowless      =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Inscrição Cadastral"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
            End
            Begin Threed.SSFrame fra 
               Height          =   810
               Index           =   4
               Left            =   60
               TabIndex        =   76
               Top             =   2580
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   1429
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
               Caption         =   "Valor da Multa Mora (%)"
               ShadowStyle     =   1
               Begin VB.TextBox txtValMinMulta 
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
                  Left            =   120
                  MaxLength       =   8
                  TabIndex        =   19
                  Tag             =   "Valor Mínimo"
                  Top             =   420
                  Width           =   615
               End
               Begin VB.TextBox txtValMaxMulta 
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
                  Left            =   840
                  MaxLength       =   8
                  TabIndex        =   20
                  Tag             =   "Valor Máximo"
                  Top             =   420
                  Width           =   645
               End
               Begin VB.TextBox txtVariacaoMulta 
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
                  Left            =   1590
                  MaxLength       =   8
                  TabIndex        =   21
                  Tag             =   "Valor Máximo"
                  Top             =   420
                  Width           =   645
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   11
                  Left            =   120
                  TabIndex        =   77
                  Top             =   210
                  Width           =   630
                  _ExtentX        =   1111
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
                  Caption         =   "Mínimo"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   13
                  Left            =   840
                  TabIndex        =   78
                  Top             =   210
                  Width           =   660
                  _ExtentX        =   1164
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
                  Caption         =   "Máximo"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   7
                  Left            =   1590
                  TabIndex        =   79
                  Top             =   210
                  Width           =   945
                  _ExtentX        =   1667
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
                  Caption         =   "Inc. Mensal"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
            End
            Begin Threed.SSFrame fra 
               Height          =   885
               Index           =   5
               Left            =   2490
               TabIndex        =   80
               Top             =   1680
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   1561
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
               Caption         =   "Valor Mínimo:"
               ShadowStyle     =   1
               Begin VB.TextBox txtValMinImposto 
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
                  MaxLength       =   8
                  TabIndex        =   16
                  Tag             =   "Imposto"
                  Top             =   510
                  Width           =   1185
               End
               Begin VB.TextBox txtValMinBase 
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
                  Left            =   120
                  MaxLength       =   8
                  TabIndex        =   15
                  Tag             =   "Base de Cálculo"
                  Top             =   510
                  Width           =   1245
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   16
                  Left            =   120
                  TabIndex        =   81
                  Top             =   300
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
                  Caption         =   "Base de Cálculo"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   17
                  Left            =   1470
                  TabIndex        =   82
                  Top             =   300
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
                  Caption         =   "Imposto"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
            End
            Begin Threed.SSFrame fraTaxa 
               Height          =   840
               Left            =   3945
               TabIndex        =   83
               Top             =   2550
               Visible         =   0   'False
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   1482
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
               Caption         =   "Taxa"
               ShadowStyle     =   1
               Begin VB.TextBox txtValorTaxa 
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
                  Left            =   525
                  MaxLength       =   8
                  TabIndex        =   24
                  Top             =   405
                  Width           =   795
               End
               Begin Threed.SSCheck chkTaxa 
                  Height          =   195
                  Left            =   60
                  TabIndex        =   23
                  Top             =   210
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   344
                  _Version        =   196610
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Fixa"
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   19
                  Left            =   60
                  TabIndex        =   84
                  Top             =   435
                  Width           =   450
                  _ExtentX        =   794
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
                  Caption         =   "Valor"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
            End
            Begin Threed.SSFrame fra 
               Height          =   855
               Index           =   3
               Left            =   4875
               TabIndex        =   85
               Top             =   30
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   1508
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
               Caption         =   "Periodicidade"
               ShadowStyle     =   1
               Begin VB.ComboBox cboPeriodoDecl 
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
                  ItemData        =   "TIMP102.frx":0114
                  Left            =   90
                  List            =   "TIMP102.frx":0116
                  Style           =   2  'Dropdown List
                  TabIndex        =   5
                  Tag             =   "Declaracao"
                  Top             =   480
                  Width           =   1605
               End
               Begin VB.ComboBox cboPeriodoCalc 
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
                  ItemData        =   "TIMP102.frx":0118
                  Left            =   1890
                  List            =   "TIMP102.frx":011A
                  Style           =   2  'Dropdown List
                  TabIndex        =   6
                  Tag             =   "Calculo"
                  Top             =   480
                  Width           =   1815
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   3
                  Left            =   60
                  TabIndex        =   86
                  Top             =   240
                  Width           =   945
                  _ExtentX        =   1667
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
                  Caption         =   " Declaração"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   4
                  Left            =   1890
                  TabIndex        =   87
                  Top             =   240
                  Width           =   615
                  _ExtentX        =   1085
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
                  Caption         =   "Cálculo"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
            End
            Begin Threed.SSFrame SSFrame2 
               Height          =   885
               Left            =   60
               TabIndex        =   88
               Top             =   1680
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   1561
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
               Caption         =   "Cálculo do Tributo(%)"
               ShadowStyle     =   1
               Begin VB.TextBox txtRedutor 
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
                  MaxLength       =   8
                  TabIndex        =   14
                  Tag             =   "Redutor"
                  Top             =   480
                  Width           =   885
               End
               Begin VB.TextBox txtAliquota 
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
                  MaxLength       =   8
                  TabIndex        =   13
                  Tag             =   "Aliquota"
                  Top             =   480
                  Width           =   885
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   14
                  Left            =   90
                  TabIndex        =   89
                  Top             =   240
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
                  Caption         =   "Aliquota"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   18
                  Left            =   1170
                  TabIndex        =   90
                  Top             =   240
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
                  Caption         =   "Reducão Base"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
            End
            Begin Threed.SSFrame SSFrame3 
               Height          =   810
               Left            =   2670
               TabIndex        =   91
               Top             =   2580
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   1429
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
               Caption         =   "Valor Juros(%)"
               ShadowStyle     =   1
               Begin VB.TextBox txtJuros 
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
                  MaxLength       =   8
                  TabIndex        =   22
                  Tag             =   "Juros"
                  Top             =   420
                  Width           =   825
               End
               Begin Threed.SSPanel lbl 
                  Height          =   270
                  Index           =   15
                  Left            =   90
                  TabIndex        =   92
                  Top             =   210
                  Width           =   450
                  _ExtentX        =   794
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
                  Caption         =   "Valor"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   1
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
            End
            Begin Threed.SSFrame SSFrame4 
               Height          =   855
               Left            =   5430
               TabIndex        =   93
               Top             =   1710
               Width           =   5025
               _ExtentX        =   8864
               _ExtentY        =   1508
               _Version        =   196610
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Receitas diversas"
               Begin VB.ComboBox CboJuros 
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
                  ItemData        =   "TIMP102.frx":011C
                  Left            =   60
                  List            =   "TIMP102.frx":0126
                  Style           =   2  'Dropdown List
                  TabIndex        =   17
                  Tag             =   "Tipo de Juros"
                  Top             =   390
                  Width           =   1695
               End
               Begin Threed.SSPanel lbl 
                  Height          =   225
                  Index           =   20
                  Left            =   60
                  TabIndex        =   94
                  Top             =   180
                  Width           =   1260
                  _ExtentX        =   2223
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
                  Caption         =   "Tipo de Juros"
                  BorderWidth     =   1
                  BevelOuter      =   0
                  AutoSize        =   2
                  Alignment       =   0
                  RoundedCorners  =   0   'False
               End
               Begin VTOcx.cboVISUAL cboCorrecao 
                  Height          =   510
                  Left            =   1770
                  TabIndex        =   18
                  Top             =   180
                  Width           =   3225
                  _ExtentX        =   5689
                  _ExtentY        =   900
                  Caption         =   "Atualização Monetária"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
            End
            Begin Threed.SSFrame SSFrame5 
               Height          =   825
               Left            =   90
               TabIndex        =   95
               Top             =   30
               Width           =   4740
               _ExtentX        =   8361
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
               Caption         =   "Vencimentos"
               ShadowStyle     =   1
               Begin VTOcx.cboVISUAL cboPeriodVenc 
                  Height          =   510
                  Left            =   60
                  TabIndex        =   2
                  Top             =   240
                  Width           =   3315
                  _ExtentX        =   5847
                  _ExtentY        =   900
                  Caption         =   "Periodo Vencimento"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboDia 
                  Height          =   510
                  Left            =   3360
                  TabIndex        =   3
                  Top             =   240
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   900
                  Caption         =   "Dia"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
                  Enabled         =   0   'False
               End
               Begin VTOcx.cboVISUAL cboMes 
                  Height          =   510
                  Left            =   4005
                  TabIndex        =   4
                  Top             =   240
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   900
                  Caption         =   "Mês"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
                  Enabled         =   0   'False
               End
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   28
               Left            =   8760
               TabIndex        =   96
               Top             =   180
               Width           =   1695
               _ExtentX        =   2990
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
               Caption         =   "Incluir Taxa Emissão"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   6
               Left            =   45
               TabIndex        =   97
               Top             =   3975
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
               Caption         =   "Base Legal"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   3
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin VTOcx.cboVISUAL CboFormaCalculo 
               Height          =   315
               Left            =   45
               TabIndex        =   27
               Top             =   3630
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   29
               Left            =   75
               TabIndex        =   98
               Top             =   3390
               Width           =   1710
               _ExtentX        =   3016
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
               Caption         =   "Cálculo Proporcional"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboGerarObrigacaoZero 
               Height          =   315
               Left            =   6900
               TabIndex        =   28
               Top             =   3630
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   30
               Left            =   4080
               TabIndex        =   100
               Top             =   3660
               Width           =   2820
               _ExtentX        =   4974
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
               Caption         =   "Gerar Obrigação com o Valor Zero"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6345
         Left            =   30
         TabIndex        =   43
         Top             =   30
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   11192
         _Version        =   131082
         TabGuid         =   "TIMP102.frx":0141
         Begin VTOcx.grdVISUAL lstImposto 
            Height          =   6480
            Left            =   60
            TabIndex        =   44
            Top             =   60
            Width           =   10500
            _ExtentX        =   18521
            _ExtentY        =   11430
            CorTitulo       =   16711680
            OcultarRodape   =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   6345
         Left            =   -99969
         TabIndex        =   102
         Top             =   30
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   11192
         _Version        =   131082
         TabGuid         =   "TIMP102.frx":0169
         Begin VB.Frame Frame3 
            Caption         =   "Descontos(%)"
            Height          =   825
            Left            =   0
            TabIndex        =   103
            Top             =   420
            Width           =   5730
            Begin VTOcx.cmdVISUAL cmdAdicionar 
               Height          =   360
               Left            =   4320
               TabIndex        =   74
               ToolTipText     =   "Definir Faixas"
               Top             =   360
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   635
               Caption         =   "Adicionar"
               Acao            =   1
               CorBorda        =   16711680
               CorFrente       =   0
               CorFundo        =   16777088
            End
            Begin VTOcx.txtVISUAL txtDescontoTributoParcFaixas 
               Height          =   495
               Left            =   60
               TabIndex        =   70
               Top             =   210
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   873
               Caption         =   "Tributo"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtDescontoMultaParcFaixas 
               Height          =   495
               Left            =   1080
               TabIndex        =   71
               Top             =   210
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   873
               Caption         =   "Multa"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtDescontoJurosParcFaixas 
               Height          =   495
               Left            =   2100
               TabIndex        =   72
               Top             =   210
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   873
               Caption         =   "Juros"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtDescontoCorrecaoParcFaixas 
               Height          =   495
               Left            =   3120
               TabIndex        =   73
               Top             =   210
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   873
               Caption         =   "Atualização"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
         End
         Begin VTOcx.txtVISUAL txtLimiteInferior 
            Height          =   285
            Left            =   30
            TabIndex        =   68
            Top             =   120
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   503
            Caption         =   "Limite Inferior"
            Text            =   ""
            Restricao       =   2
         End
         Begin VTOcx.txtVISUAL txtLimiteSuperior 
            Height          =   285
            Left            =   3105
            TabIndex        =   69
            Top             =   120
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   503
            Caption         =   "Limite Superior"
            Text            =   ""
            Restricao       =   2
         End
         Begin VTOcx.grdVISUAL grdFaixas 
            Height          =   3390
            Left            =   30
            TabIndex        =   104
            Top             =   1320
            Width           =   10440
            _ExtentX        =   18415
            _ExtentY        =   5980
            Caption         =   "Tipos"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   192
         End
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   195
      Left            =   1890
      TabIndex        =   41
      Top             =   840
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6780
      Top             =   945
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   105
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   1138
      Icone           =   "TIMP102.frx":0191
   End
End
Attribute VB_Name = "TIMP102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub AtualizaGrig()
    Dim Sql As String
    Sql = " SELECT tpi_tip_cod_imposto as Código, tip_sigla_imposto as Sigla,"
    Sql = Sql & " tip_nome_imposto as Imposto, tpi_ano_imposto as Ano"
    Sql = Sql & " From Tab_Parametro_Imposto, Tab_Imposto"
    Sql = Sql & " where tpi_tip_cod_imposto = tip_cod_imposto "
    lstImposto.Preencher Bdados, Sql, 1000, 1400, 6500, 700
End Sub
Public Sub GravaFaixasREFIS(CodTributo As String, Ano As String, Lista As Object)
    Dim Valores As String
    Dim Campos As String
    Dim Ilaco As Integer
    If tabTributo.Tabs(3).Visible = False Then Exit Sub
    Campos = "TPI_TIP_COD_IMPOSTO,TPI_ANO_IMPOSTO,TPI_LIMITE_INFERIOR,TPI_LIMITE_SUPERIOR,TPI_DESCONTO_TRIBUTO_PARC," & _
        "TPI_DESCONTO_JUROS_PARC,TPI_DESCONTO_MULTA_PARC,TPI_DESCONTO_ATUALIZACAO_PARC"
    Bdados.DeletaDados "TAB_PARAMETRO_IMPOSTO_REFIS", "TPI_TIP_COD_IMPOSTO ='" & CodTributo & "'" & IIf(cboTipoTributo = "TAXA", "", " and tpi_ano_imposto='" & Ano & "'")
    If Lista.ListItems.Count >= 1 Then
        For Ilaco = 1 To Lista.ListItems.Count
            Valores = Bdados.PreparaValor(CodTributo, Ano, Lista.ListItems(Ilaco), Lista.ListItems(Ilaco).SubItems(1), _
                Lista.ListItems(Ilaco).SubItems(2), _
                Lista.ListItems(Ilaco).SubItems(3), Lista.ListItems(Ilaco).SubItems(4), Lista.ListItems(Ilaco).SubItems(5))
             Bdados.InsereDados "TAB_PARAMETRO_IMPOSTO_REFIS", Valores, Campos
        Next
    End If
End Sub
Sub CarregaCombos()
    Call Edita.AtualizaCombo(Bdados, cboTipoContrib, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =1 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboTipoTributo, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =2 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboPeriodoDecl, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =3 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboPeriodoCalc, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =4 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboTipoInsc, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =6 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboTipoIC, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =6 ORDER BY TGE_CODIGO ASC")
End Sub

Private Sub cboPeriodVenc_Click()
    Select Case cboPeriodVenc.Coluna(1).Valor
        Case 1
            cboDia.Enabled = False
            cboMes.Enabled = False
        Case 2
            cboDia.Enabled = True
            cboMes.Enabled = False
        Case 3
            cboDia.Enabled = True
            cboMes.Enabled = False
        Case 4
            cboDia.Enabled = True
            cboMes.Enabled = True
        Case 6
            cboDia.Enabled = True
            cboMes.Enabled = False
    End Select
End Sub

Private Sub cboTipoTributo_Click()
    If cboTipoTributo = "TAXA" Then
        fraTaxa.Visible = True
    Else
        fraTaxa.Visible = False
    End If
End Sub

Private Sub chkTaxa_Click(Value As Integer)
    On Error Resume Next
    If Value Then
        fraTaxa.Visible = True
        txtAliquota = 0
        txtAliquota.Enabled = False
        txtValorTaxa.Enabled = True
        txtValorTaxa.SetFocus
    Else
        txtAliquota.Enabled = True
        txtValorTaxa.Enabled = False
        txtAliquota.SetFocus
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Select Case Index
        Case 0
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Valores = Bdados.PreparaValor(txtCodImposto, Bdados.Converte(txtAnoImposto, tctexto), cboPeriodVenc.Coluna(1).Valor, IIf(cboDia.Enabled, cboDia, 0), _
                    IIf(cboMes.Enabled, cboMes, 0), cboPeriodoDecl.ListIndex + 1, cboPeriodoCalc.ListIndex + 1, cboTipoContrib.ListIndex + 1, cboTipoTributo.ListIndex + 1, _
                    cboTipoInsc.ListIndex + 1, Bdados.Converte(txtValMinMulta, TCDuplo), Bdados.Converte(txtValMaxMulta, TCDuplo), Bdados.Converte(txtValMinBase, TCDuplo), Bdados.Converte(txtValMinImposto, TCDuplo), _
                    Bdados.Converte(txtAliquota, TCDuplo), Bdados.Converte(txtJuros, TCDuplo), _
                    Bdados.Converte(txtRedutor, TCDuplo), Bdados.Converte(IIf(Trim(txtValorTaxa) = "", 0, txtValorTaxa), _
                    TCDuplo), CboJuros.ListIndex, cboTipoIC.ListIndex + 1, cboObrigacao.Coluna(1).Valor, _
                    Nvl("" & cboReceitaAMais.Coluna(1).Valor, 0), Nvl("" & cboReceitaAMenos.Coluna(1).Valor, 0), _
                    Nvl("" & cboCorrecao.Coluna(1).Valor, 0), Bdados.Converte(Nvl(txtDesconto, 0), TCDuplo), _
                    Bdados.Converte(Nvl(txtDescontoJuros, 0), TCDuplo), Bdados.Converte(Nvl(txtDescontoMulta, 0), TCDuplo), _
                    Bdados.Converte(Nvl(txtDescontoCorrecao, 0), TCDuplo), Bdados.Converte(Nvl(txtDescontoTributoParc, 0), TCDuplo), _
                    Bdados.Converte(Nvl(txtDescontoJurosParc, 0), TCDuplo), Bdados.Converte(Nvl(txtDescontoMultaParc, 0), TCDuplo), _
                    Bdados.Converte(Nvl(txtDescontoCorrecaoParc, 0), TCDuplo), cboTaxaEmissao.ListIndex + 1, txtLei, Bdados.Converte(Nvl(txtVariacaoMulta, 0), _
                    TCMonetario), Bdados.Converte(cboDescontoAteVencimento.Coluna(1).Valor, tctexto), _
                    Bdados.Converte(CboFormaCalculo.Coluna(1).Valor, tctexto), cboGerarObrigacaoZero.Coluna(1).Valor, IIf(tabTributo.Tabs(3).Visible = False, 1, 2))
                    
            Campos = "tpi_tip_cod_imposto,tpi_ano_imposto,TPI_PERIODO_VENCIMENTO,TPI_DIA_BASE,TPI_MES_BASE,"
            Campos = Campos & "TPI_PERIODIC_DECLARA,TPI_PERIODIC_CALCULO,tpi_tipo_contribuinte,tpi_tipo_tributo,"
            Campos = Campos & "tpi_tipo_inscricao,tpi_valor_min_multa,tpi_valor_max_multa,"
            Campos = Campos & "tpi_valor_min_base_calc,tpi_valor_min_imposto,"
            Campos = Campos & "tpi_aliquota,tpi_valor_juros,tpi_reducao,tpi_valor_taxa_fixa," & _
                            "TPI_JUROS_CAPTALIZADOS,tpi_tipo_ic,TPI_GERA_OBRIGACAO,TPI_RECEITA_A_MAIS," & _
                            "TPI_RECEITA_A_MENOS,TPI_AUTORIZACAO_CORRECAO,TPI_DESCONTO_TRIBUTO," & _
                            "TPI_DESCONTO_JUROS,TPI_DESCONTO_MULTA,TPI_DESCONTO_ATUALIZACAO," & _
                            "TPI_DESCONTO_TRIBUTO_PARC,TPI_DESCONTO_JUROS_PARC,TPI_DESCONTO_MULTA_PARC," & _
                            "TPI_DESCONTO_ATUALIZACAO_PARC,TPI_GERA_TAXA_IMPRESSAO,tpi_lei," & _
                            "TPI_INCREMENTO_MENSAL_MULTA,TPI_DESCONTO_ATE_VENCIMENTO,TPI_FORMA_CALCULO," & _
                            "TPI_GERAR_OBRIGACAO_ZERADA,TPI_PARAMETRO_REFIS"
                            
            If Bdados.GravaDados("Tab_Parametro_Imposto", Valores, Campos, "tpi_tip_cod_imposto='" & txtCodImposto & "'" & IIf(cboTipoTributo = "TAXA", "", " and tpi_ano_imposto='" & txtAnoImposto & "'")) Then
                GravaFaixasREFIS txtCodImposto, txtAnoImposto, grdFaixas
                Call Util.Informa("Transação Completada.")
                tabTributo.Tabs(1).Selected = True
                tabTributo.Tabs(3).Visible = False
                grdFaixas.Preencher Bdados, ""
            End If
            Edita.LimpaCampos Me
            AtualizaGrig
            tabTributo.Tabs(1).Selected = True
            
        Case 1
            Unload Me
    End Select
End Sub

Private Sub cmdAdicionar_Click()
    Dim Indice As Integer
    If Trim(txtLimiteInferior) = "" Or Trim(txtLimiteSuperior) = "" Then Exit Sub
    Indice = grdFaixas.ListItems.Count + 1
    grdFaixas.ListItems.Add Indice, , txtLimiteInferior
    grdFaixas.ListItems(Indice).SubItems(1) = txtLimiteSuperior
    grdFaixas.ListItems(Indice).SubItems(2) = txtDescontoTributoParcFaixas
    grdFaixas.ListItems(Indice).SubItems(3) = txtDescontoMultaParcFaixas
    grdFaixas.ListItems(Indice).SubItems(4) = txtDescontoJurosParcFaixas
    grdFaixas.ListItems(Indice).SubItems(5) = txtDescontoCorrecaoParcFaixas
    txtLimiteInferior.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    tabTributo.Tabs(3).Visible = True
    tabTributo.Tabs(3).Selected = True
    txtLimiteInferior.SetFocus
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImpostosRelacionados_Click()
    If Trim(txtCodImposto) = "" Then Exit Sub
    TIMP105.CarregaTributos txtCodImposto, txtAnoImposto
    TIMP105.Show 1
End Sub

Private Sub Form_Load()
    Dim Sql As String
    Dim i As Byte
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    Call CarregaCombos
    AtualizaGrig
    AtualizaCabecalho lstImposto
    cboObrigacao.PreencherGeral Bdados, "LANCAMENTO OBRIGACAO"
    cboReceitaAMais.PreencherGeral Bdados, "RECEITA A MAIS"
    cboReceitaAMenos.PreencherGeral Bdados, "RECEITA A MENOS"
    cboCorrecao.PreencherGeral Bdados, "AUTORIZACAO CORRECAO"
    cboPeriodVenc.PreencherGeral Bdados, "PERIODO VENCIMENTO"
    cboDescontoAteVencimento.PreencherGeral Bdados, "PRAZO"
    CboFormaCalculo.PreencherGeral Bdados, "FORMA DE CALCULO"
    cboGerarObrigacaoZero.PreencherGeral Bdados, "GERAR TRIBUTO COM VALOR ZERO"
    For i = 1 To 31
        'dias
        cboDia.AddItem Format(i, "00")
    Next
    For i = 1 To 12
        'Meses
        cboMes.AddItem Format(i, "00")
    Next
    grdFaixas.ColumnHeaders.Clear
    grdFaixas.ColumnHeaders.Add , , "Limite Inferior", 1500
    grdFaixas.ColumnHeaders.Add , , "Limite Superior", 1500
    grdFaixas.ColumnHeaders.Add , , "Tributo", 1000
    grdFaixas.ColumnHeaders.Add , , "Multa", 1000
    grdFaixas.ColumnHeaders.Add , , "Juros", 1000
    grdFaixas.ColumnHeaders.Add , , "Atualização", 1000
    tabTributo.Tabs(3).Visible = False
End Sub

Private Sub lstImposto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstImposto, ColumnHeader
End Sub

Private Sub grdFaixas_Click()
    Dim i As Integer
    If grdFaixas.ListItems.Count >= 1 Then
        grdFaixas.ListItems.Remove grdFaixas.SelectedItem.Index
    End If
End Sub

Private Sub lstImposto_DblClick()
    Dim Sql As String
    txtCodImposto = lstImposto.SelectedItem
    txtCodImposto_LostFocus
    txtAnoImposto = lstImposto.SelectedItem.SubItems(3)
    txtAnoImposto_LostFocus
    tabTributo.Tabs(2).Selected = True
    tabTributo.Tabs(3).Visible = False
    Sql = "SELECT TPI_LIMITE_INFERIOR [Limite Inferior],TPI_LIMITE_SUPERIOR [Limite Superior],TPI_DESCONTO_TRIBUTO_PARC [Tributo],TPI_DESCONTO_MULTA_PARC Multa,TPI_DESCONTO_JUROS_PARC Juros,TPI_DESCONTO_ATUALIZACAO_PARC atualização FROM TAB_PARAMETRO_IMPOSTO_REFIS where tpi_tip_cod_imposto ='" & txtCodImposto & "' and (tpi_ano_imposto='" & txtAnoImposto & "'" & IIf(Trim(txtAnoImposto) <> "", ")", " OR tpi_ano_imposto IS NULL)")
    If grdFaixas.Preencher(Bdados, Sql, 1500, 1500, 1000, 1000, 1000, 1000) Then
        tabTributo.Tabs(3).Visible = True
    End If
End Sub

Private Sub lstImposto_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    Dim Sql As String
    
    If Confirma("Deseja excluir o tributo " & lstImposto.SelectedItem.SubItems(1) & " ano " & lstImposto.SelectedItem.SubItems(3) & "?") Then
        Bdados.DeletaDados "Tab_Parametro_Imposto", " tpi_tip_cod_imposto =  '" & lstImposto.SelectedItem & "' and tpi_ano_imposto ='" & lstImposto.SelectedItem.SubItems(3) & "'"
        Avisa "Parâmetro excluído com sucesso! "
        txtCodImposto.Enabled = True
        Edita.LimpaCampos Me
        AtualizaGrig
    End If
End Sub

Private Sub txtAliquota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtAnoImposto_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtAnoImposto_LostFocus()
    
    
    Dim Sql As String
    Dim rs As VSRecordset
    On Error Resume Next
    If Trim(txtCodImposto) = "" Then Exit Sub
    If Trim(Len(Trim(txtAnoImposto))) = 1 Then
        Avisa "Formato do Ano = 'AA'"
        txtAnoImposto.SetFocus
        Exit Sub
    End If
    txtLei = ""
    Sql = "SELECT * from Tab_Parametro_Imposto where tpi_tip_cod_imposto ='" & txtCodImposto
    Sql = Sql & "'  and (tpi_ano_imposto='" & txtAnoImposto & "'" & IIf(Trim(txtAnoImposto) <> "", ")", " OR tpi_ano_imposto IS NULL)")
    If Bdados.AbreTabela(Sql, rs) Then
        cboPeriodVenc.SetarLinha Nvl("" & rs!TPI_PERIODO_VENCIMENTO, 0), 1
        If Nvl("" & rs!TPI_DIA_BASE, 0) <> 0 Then
            cboDia.Enabled = True
            cboDia = Nvl("" & rs!TPI_DIA_BASE, 0)
        Else
            cboDia.Enabled = False
            cboDia = ""
        End If
        If Nvl("" & rs!TPI_MES_BASE, 0) <> 0 Then
             cboMes.Enabled = True
             cboMes = Format(Nvl("" & rs!TPI_MES_BASE, 0), "00")
        Else
            cboMes.Enabled = False
            cboMes = ""
        End If
        txtLei = "" & rs.Fields("tpi_lei")
        cboGerarObrigacaoZero.SetarLinha "" & rs.Fields("TPI_GERAR_OBRIGACAO_ZERADA"), 1
        cboPeriodoDecl.ListIndex = rs!TPI_PERIODIC_DECLARA - 1
        cboPeriodoCalc.ListIndex = rs!TPI_PERIODIC_CALCULO - 1
        cboTipoContrib.ListIndex = rs!tpi_tipo_contribuinte - 1
        cboTipoTributo.ListIndex = rs!tpi_tipo_tributo - 1
        cboTipoInsc.ListIndex = rs!tpi_tipo_inscricao - 1
        cboTipoIC.ListIndex = Nvl("" & rs!tpi_tipo_ic - 1, -1)
        CboFormaCalculo.SetarLinha Nvl("" & rs.Fields("TPI_FORMA_CALCULO"), 0), 1
        cboTaxaEmissao.ListIndex = Nvl("" & rs!TPI_GERA_TAXA_IMPRESSAO - 1, -1)
        txtValMinMulta = Format(rs!tpi_valor_min_multa, Const_Monetario)
        cboDescontoAteVencimento.SetarLinha Nvl("" & rs!TPI_DESCONTO_ATE_VENCIMENTO, 0), 1
        txtValMaxMulta = Format(rs!tpi_valor_max_multa, Const_Monetario)
        txtVariacaoMulta = Format(rs!TPI_INCREMENTO_MENSAL_MULTA, Const_Monetario)
        txtValMinBase = Format(rs!tpi_valor_min_base_calc, Const_Monetario)
        txtValMinImposto = Format(rs!tpi_valor_min_imposto, Const_Monetario)
        txtAliquota = Format("" & rs!tpi_aliquota, Const_Monetario)
        txtJuros = Format("" & rs!tpi_valor_juros, Const_Monetario)
        txtRedutor = Format("" & rs!tpi_reducao, Const_Monetario)
        txtDesconto = Format(Nvl("" & rs!tpi_desconto, 0), Const_Monetario)
        cboObrigacao.SetarLinha Nvl("" & rs!TPI_GERA_OBRIGACAO, 0), 1
        cboReceitaAMais.SetarLinha Nvl("" & rs!TPI_RECEITA_A_MAIS, 0), 1
        cboReceitaAMenos.SetarLinha Nvl("" & rs!TPI_RECEITA_A_MENOS, 0), 1
        cboCorrecao.SetarLinha Nvl("" & rs!TPI_AUTORIZACAO_CORRECAO, 0), 1
        
        txtDesconto = Format(Nvl("" & rs!TPI_DESCONTO_TRIBUTO, 0), Const_Monetario)
        txtDescontoCorrecao = Format(Nvl("" & rs!TPI_DESCONTO_ATUALIZACAO, 0), Const_Monetario)
        txtDescontoJuros = Format(Nvl("" & rs!TPI_DESCONTO_JUROS, 0), Const_Monetario)
        txtDescontoMulta = Format(Nvl("" & rs!TPI_DESCONTO_MULTA, 0), Const_Monetario)
        
        txtDescontoCorrecaoParc = Format(Nvl("" & rs!TPI_DESCONTO_ATUALIZACAO_PARC, 0), Const_Monetario)
        txtDescontoJurosParc = Format(Nvl("" & rs!TPI_DESCONTO_JUROS_PARC, 0), Const_Monetario)
        txtDescontoMultaParc = Format(Nvl("" & rs!TPI_DESCONTO_MULTA_PARC, 0), Const_Monetario)
        txtDescontoTributoParc = Format(Nvl("" & rs!TPI_DESCONTO_TRIBUTO_PARC, 0), Const_Monetario)
        fraTaxa.Visible = IIf(rs!tpi_tipo_tributo = 2, True, False)
        CboJuros.ListIndex = IIf(IsNull(rs!TPI_JUROS_CAPTALIZADOS), 0, rs!TPI_JUROS_CAPTALIZADOS)
        cboTaxaEmissao.ListIndex = CInt(Nvl("" & rs!TPI_GERA_TAXA_IMPRESSAO, -1)) - 1
        If Not IsNull(rs!tpi_valor_taxa_fixa) Then
            If rs!tpi_valor_taxa_fixa > 0 Then
                txtValorTaxa = rs!tpi_valor_taxa_fixa
                txtValorTaxa = Edita.FormataTexto(txtValorTaxa, Monetario, True)
                chkTaxa.Value = ssCBChecked
            Else
                 fraTaxa.Visible = False
                txtValorTaxa = ""
                chkTaxa.Value = ssCBUnchecked
            End If
        Else
            txtValorTaxa = ""
            chkTaxa.Value = ssCBUnchecked
        End If
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtCodImposto_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtCodImposto) = "" Then Exit Sub
    If Me.ActiveControl.Name = "cmd" Then Exit Sub
    
    Sql = "Select tip_nome_imposto from Tab_Imposto Where tip_cod_imposto='" & txtCodImposto & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtNomeImposto = rs(0)
    Else
        txtNomeImposto = ""
        Avisa "Código de Tributo inválido."
        txtCodImposto.SetFocus
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtDiasDecl_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDiasPag_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDtInicio_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtJuros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtRedutor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValMaxMulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValMaxMulta_LostFocus()
    txtValMaxMulta = Edita.FormataTexto(txtValMaxMulta, Monetario, True)
End Sub

Private Sub txtValMinBase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValMinBase_LostFocus()
    txtValMinBase = Edita.FormataTexto(txtValMinBase, Monetario, True)
End Sub

Private Sub txtValMinImposto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValMinImposto_LostFocus()
    txtValMinImposto = Edita.FormataTexto(txtValMinImposto, Monetario, True)
End Sub

Private Sub txtValMinMulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValMinMulta_LostFocus()
    'txtValMinMulta = Edita.FormataTexto(txtValMinMulta, Monetario,true)
End Sub

Private Sub txtValorTaxa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub


Private Sub txtValorTaxa_LostFocus()
    txtValorTaxa = Edita.FormataTexto(txtValorTaxa, Monetario, True)
End Sub

Private Sub txtVariacaoMulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub
