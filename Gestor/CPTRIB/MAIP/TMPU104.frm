VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TMPU104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   6330
      Index           =   2
      Left            =   0
      TabIndex        =   36
      Top             =   600
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   11165
      _Version        =   196610
      Font3D          =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin TabDlg.SSTab SSTab1 
         Height          =   3825
         Left            =   30
         TabIndex        =   38
         Top             =   2430
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   6747
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Infra - Estrutura / Serviços"
         TabPicture(0)   =   "TMPU104.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fra(5)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Item do Boletim Cadastral / Sistema Viário"
         TabPicture(1)   =   "TMPU104.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra(1)"
         Tab(1).Control(1)=   "fra(0)"
         Tab(1).ControlCount=   2
         Begin Threed.SSFrame fra 
            Height          =   1395
            Index           =   0
            Left            =   -74910
            TabIndex        =   39
            Top             =   450
            Width           =   11085
            _ExtentX        =   19553
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
            Caption         =   "Item do Boletim Cadastral"
            Alignment       =   2
            ShadowStyle     =   1
            Begin VB.TextBox txtCodComponente 
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
               Index           =   16
               Left            =   7230
               MaxLength       =   3
               TabIndex        =   21
               Top             =   630
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   15
               Left            =   7230
               MaxLength       =   3
               TabIndex        =   20
               Top             =   270
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   14
               Left            =   1650
               MaxLength       =   3
               TabIndex        =   19
               Top             =   990
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   13
               Left            =   1650
               MaxLength       =   3
               TabIndex        =   18
               Top             =   630
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   12
               Left            =   1650
               MaxLength       =   3
               TabIndex        =   17
               Top             =   270
               Width           =   375
            End
            Begin VB.ComboBox cboEstacionamento 
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
               ItemData        =   "TMPU104.frx":0038
               Left            =   2070
               List            =   "TMPU104.frx":003A
               Style           =   2  'Dropdown List
               TabIndex        =   46
               TabStop         =   0   'False
               Tag             =   "13"
               Top             =   270
               Width           =   3375
            End
            Begin VB.ComboBox cboLimpeza 
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
               ItemData        =   "TMPU104.frx":003C
               Left            =   2070
               List            =   "TMPU104.frx":003E
               Style           =   2  'Dropdown List
               TabIndex        =   45
               TabStop         =   0   'False
               Tag             =   "14"
               Top             =   630
               Width           =   3375
            End
            Begin VB.ComboBox cboConservacao 
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
               ItemData        =   "TMPU104.frx":0040
               Left            =   2070
               List            =   "TMPU104.frx":0042
               Style           =   2  'Dropdown List
               TabIndex        =   44
               TabStop         =   0   'False
               Tag             =   "15"
               Top             =   990
               Width           =   3375
            End
            Begin VB.ComboBox cboPonto 
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
               ItemData        =   "TMPU104.frx":0044
               Left            =   7650
               List            =   "TMPU104.frx":0046
               Style           =   2  'Dropdown List
               TabIndex        =   43
               TabStop         =   0   'False
               Tag             =   "16"
               Top             =   270
               Width           =   3375
            End
            Begin VB.ComboBox cboHidrante 
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
               ItemData        =   "TMPU104.frx":0048
               Left            =   7650
               List            =   "TMPU104.frx":004A
               Style           =   2  'Dropdown List
               TabIndex        =   42
               TabStop         =   0   'False
               Tag             =   "17"
               Top             =   630
               Width           =   3375
            End
            Begin VB.ComboBox Combo2 
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
               ItemData        =   "TMPU104.frx":004C
               Left            =   7410
               List            =   "TMPU104.frx":004E
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   2310
               Width           =   2535
            End
            Begin VB.ComboBox cboInstSanit17 
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
               ItemData        =   "TMPU104.frx":0050
               Left            =   6930
               List            =   "TMPU104.frx":0052
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   1950
               Width           =   3015
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   9
               Left            =   6390
               TabIndex        =   47
               Top             =   660
               Width           =   690
               _ExtentX        =   1217
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
               Caption         =   "Hidrante"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   10
               Left            =   5685
               TabIndex        =   48
               Top             =   330
               Width           =   1395
               _ExtentX        =   2461
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
               Caption         =   "Pontos  Serviços"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   11
               Left            =   420
               TabIndex        =   49
               Top             =   1020
               Width           =   1155
               _ExtentX        =   2037
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
               Caption         =   "Conserv. Vias"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   12
               Left            =   255
               TabIndex        =   50
               Top             =   330
               Width           =   1320
               _ExtentX        =   2328
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
               Caption         =   "Estacionamento"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   180
               Index           =   13
               Left            =   240
               TabIndex        =   51
               Top             =   660
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   318
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
               Caption         =   "Limpeza Pública"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   31
               Left            =   5190
               TabIndex        =   52
               Top             =   2010
               Width           =   1680
               _ExtentX        =   2963
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
               Caption         =   "Instalação Sanitária:"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   3
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   14
               Left            =   5700
               TabIndex        =   53
               Top             =   2370
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
               Caption         =   "Instalação Elétrica:"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   3
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame fra 
            Height          =   1815
            Index           =   1
            Left            =   -74940
            TabIndex        =   54
            Top             =   1860
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   3201
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
            Caption         =   "Sistema Viário"
            Alignment       =   2
            ShadowStyle     =   1
            Begin VB.ComboBox cboTipoDeLogr 
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
               ItemData        =   "TMPU104.frx":0054
               Left            =   7650
               List            =   "TMPU104.frx":0056
               Style           =   2  'Dropdown List
               TabIndex        =   100
               TabStop         =   0   'False
               Tag             =   "20"
               Top             =   1350
               Width           =   3375
            End
            Begin VB.TextBox Text1 
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
               Left            =   7230
               MaxLength       =   3
               TabIndex        =   28
               Tag             =   "25"
               Top             =   990
               Width           =   1365
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   19
               Left            =   7230
               MaxLength       =   3
               TabIndex        =   29
               Top             =   1350
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   18
               Left            =   7230
               MaxLength       =   3
               TabIndex        =   27
               Top             =   630
               Width           =   375
            End
            Begin VB.ComboBox cboCiclovia 
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
               ItemData        =   "TMPU104.frx":0058
               Left            =   7650
               List            =   "TMPU104.frx":005A
               Style           =   2  'Dropdown List
               TabIndex        =   93
               TabStop         =   0   'False
               Tag             =   "19"
               Top             =   630
               Width           =   3375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   17
               Left            =   7230
               MaxLength       =   3
               TabIndex        =   26
               Top             =   270
               Width           =   375
            End
            Begin VB.TextBox txtCanteiro 
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
               Left            =   2115
               MaxLength       =   3
               TabIndex        =   25
               Tag             =   "24"
               Top             =   1350
               Width           =   1365
            End
            Begin VB.TextBox txtCalcada 
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
               Left            =   2115
               MaxLength       =   3
               TabIndex        =   24
               Tag             =   "23"
               Top             =   990
               Width           =   1365
            End
            Begin VB.TextBox txtViaProj 
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
               Left            =   2115
               MaxLength       =   3
               TabIndex        =   23
               Tag             =   "22"
               Top             =   630
               Width           =   1365
            End
            Begin VB.TextBox txtVia 
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
               Left            =   2115
               MaxLength       =   3
               TabIndex        =   22
               Tag             =   "21"
               Top             =   270
               Width           =   1365
            End
            Begin VB.ComboBox cboVia 
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
               ItemData        =   "TMPU104.frx":005C
               Left            =   7650
               List            =   "TMPU104.frx":005E
               Style           =   2  'Dropdown List
               TabIndex        =   57
               TabStop         =   0   'False
               Tag             =   "18"
               Top             =   270
               Width           =   3375
            End
            Begin VB.ComboBox Combo3 
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
               ItemData        =   "TMPU104.frx":0060
               Left            =   7410
               List            =   "TMPU104.frx":0062
               Style           =   2  'Dropdown List
               TabIndex        =   56
               Top             =   2310
               Width           =   2535
            End
            Begin VB.ComboBox Combo4 
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
               ItemData        =   "TMPU104.frx":0064
               Left            =   6930
               List            =   "TMPU104.frx":0066
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Top             =   1950
               Width           =   3015
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   16
               Left            =   360
               TabIndex        =   58
               Top             =   1380
               Width           =   1650
               _ExtentX        =   2910
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
               Caption         =   "Largura do Canteiro"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   17
               Left            =   450
               TabIndex        =   59
               Top             =   1020
               Width           =   1560
               _ExtentX        =   2752
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
               Caption         =   "Largura da Calçada"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   18
               Left            =   825
               TabIndex        =   60
               Top             =   300
               Width           =   1185
               _ExtentX        =   2090
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
               Caption         =   "Largura da Via"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   180
               Index           =   19
               Left            =   570
               TabIndex        =   61
               Top             =   660
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   318
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
               Caption         =   "Largura Via(Proj.)"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   23
               Left            =   5190
               TabIndex        =   62
               Top             =   2010
               Width           =   1680
               _ExtentX        =   2963
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
               Caption         =   "Instalação Sanitária:"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   3
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   24
               Left            =   5700
               TabIndex        =   63
               Top             =   2370
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
               Caption         =   "Instalação Elétrica:"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   3
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   15
               Left            =   6210
               TabIndex        =   94
               Top             =   300
               Width           =   915
               _ExtentX        =   1614
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
               Caption         =   "Tipo de Via"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   25
               Left            =   5490
               TabIndex        =   95
               Top             =   1410
               Width           =   1635
               _ExtentX        =   2884
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
               Caption         =   "Tipo de Logradouro"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   26
               Left            =   6150
               TabIndex        =   96
               Top             =   1050
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "Uso do Solo"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   180
               Index           =   27
               Left            =   6495
               TabIndex        =   97
               Top             =   690
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   318
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
               Caption         =   "Ciclovia"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame fra 
            Height          =   2175
            Index           =   5
            Left            =   90
            TabIndex        =   64
            Top             =   450
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   3836
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
            Caption         =   "Infra - Estrutura"
            Alignment       =   2
            ShadowStyle     =   1
            Begin VB.TextBox txtCodComponente 
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
               Left            =   1575
               MaxLength       =   3
               TabIndex        =   9
               Top             =   1740
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Left            =   1575
               MaxLength       =   3
               TabIndex        =   8
               Top             =   1380
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Left            =   1575
               MaxLength       =   3
               TabIndex        =   7
               Top             =   1020
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Left            =   1575
               MaxLength       =   3
               TabIndex        =   6
               Top             =   660
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Left            =   1575
               MaxLength       =   3
               TabIndex        =   5
               Top             =   300
               Width           =   375
            End
            Begin VB.ComboBox cboAgua 
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
               ItemData        =   "TMPU104.frx":0068
               Left            =   1995
               List            =   "TMPU104.frx":006A
               Style           =   2  'Dropdown List
               TabIndex        =   75
               TabStop         =   0   'False
               Tag             =   "1"
               Top             =   300
               Width           =   3375
            End
            Begin VB.ComboBox cboEsgoto 
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
               ItemData        =   "TMPU104.frx":006C
               Left            =   1995
               List            =   "TMPU104.frx":006E
               Style           =   2  'Dropdown List
               TabIndex        =   74
               TabStop         =   0   'False
               Tag             =   "2"
               Top             =   660
               Width           =   3375
            End
            Begin VB.ComboBox cboEletrica 
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
               ItemData        =   "TMPU104.frx":0070
               Left            =   1995
               List            =   "TMPU104.frx":0072
               Style           =   2  'Dropdown List
               TabIndex        =   73
               TabStop         =   0   'False
               Tag             =   "3"
               Top             =   1020
               Width           =   3375
            End
            Begin VB.ComboBox cboTelefonica 
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
               ItemData        =   "TMPU104.frx":0074
               Left            =   1995
               List            =   "TMPU104.frx":0076
               Style           =   2  'Dropdown List
               TabIndex        =   72
               TabStop         =   0   'False
               Tag             =   "4"
               Top             =   1380
               Width           =   3375
            End
            Begin VB.ComboBox cboCalcada 
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
               ItemData        =   "TMPU104.frx":0078
               Left            =   1995
               List            =   "TMPU104.frx":007A
               Style           =   2  'Dropdown List
               TabIndex        =   71
               TabStop         =   0   'False
               Tag             =   "5"
               Top             =   1740
               Width           =   3375
            End
            Begin VB.ComboBox cboInstElet18 
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
               ItemData        =   "TMPU104.frx":007C
               Left            =   7410
               List            =   "TMPU104.frx":007E
               Style           =   2  'Dropdown List
               TabIndex        =   70
               Top             =   2550
               Width           =   2535
            End
            Begin VB.ComboBox cboIluminacao 
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
               ItemData        =   "TMPU104.frx":0080
               Left            =   7590
               List            =   "TMPU104.frx":0082
               Style           =   2  'Dropdown List
               TabIndex        =   69
               TabStop         =   0   'False
               Tag             =   "7"
               Top             =   660
               Width           =   3375
            End
            Begin VB.ComboBox cboDrenagem 
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
               ItemData        =   "TMPU104.frx":0084
               Left            =   7590
               List            =   "TMPU104.frx":0086
               Style           =   2  'Dropdown List
               TabIndex        =   68
               TabStop         =   0   'False
               Tag             =   "6"
               Top             =   300
               Width           =   3375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   5
               Left            =   7170
               MaxLength       =   3
               TabIndex        =   10
               Top             =   300
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   6
               Left            =   7170
               MaxLength       =   3
               TabIndex        =   11
               Top             =   660
               Width           =   375
            End
            Begin VB.ComboBox cboPavimentacao 
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
               ItemData        =   "TMPU104.frx":0088
               Left            =   7590
               List            =   "TMPU104.frx":008A
               Style           =   2  'Dropdown List
               TabIndex        =   67
               TabStop         =   0   'False
               Tag             =   "8"
               Top             =   1020
               Width           =   3375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   7
               Left            =   7170
               MaxLength       =   3
               TabIndex        =   12
               Top             =   1020
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   9
               Left            =   7170
               MaxLength       =   3
               TabIndex        =   14
               Top             =   1740
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   8
               Left            =   7170
               MaxLength       =   3
               TabIndex        =   13
               Top             =   1380
               Width           =   375
            End
            Begin VB.ComboBox cboArborizacao 
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
               ItemData        =   "TMPU104.frx":008C
               Left            =   7575
               List            =   "TMPU104.frx":008E
               Style           =   2  'Dropdown List
               TabIndex        =   66
               TabStop         =   0   'False
               Tag             =   "10"
               Top             =   1740
               Width           =   3405
            End
            Begin VB.ComboBox cboSarjeta 
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
               ItemData        =   "TMPU104.frx":0090
               Left            =   7575
               List            =   "TMPU104.frx":0092
               Style           =   2  'Dropdown List
               TabIndex        =   65
               TabStop         =   0   'False
               Tag             =   "9"
               Top             =   1380
               Width           =   3405
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   30
               Left            =   825
               TabIndex        =   76
               Top             =   1770
               Width           =   630
               _ExtentX        =   1111
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
               Caption         =   "Calçada"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   32
               Left            =   135
               TabIndex        =   77
               Top             =   1440
               Width           =   1320
               _ExtentX        =   2328
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
               Caption         =   "Rede Telefônica"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   37
               Left            =   5700
               TabIndex        =   78
               Top             =   2370
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
               Caption         =   "Instalação Elétrica:"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   3
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   225
               Index           =   33
               Left            =   390
               TabIndex        =   79
               Top             =   1050
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   397
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
               Caption         =   "Rede Elétrica"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   225
               Index           =   34
               Left            =   315
               TabIndex        =   80
               Top             =   330
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   397
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
               Caption         =   "Rede de Água"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   225
               Index           =   35
               Left            =   165
               TabIndex        =   81
               Top             =   690
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   397
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
               Caption         =   "Rede de Esgoto"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   210
               Index           =   20
               Left            =   5520
               TabIndex        =   82
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   370
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
               Caption         =   "Iluminação Pública"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   2
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   210
               Index           =   21
               Left            =   6180
               TabIndex        =   83
               Top             =   330
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   370
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
               Caption         =   "Drenagem"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   2
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   210
               Index           =   22
               Left            =   5895
               TabIndex        =   84
               Top             =   1080
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   370
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
               Caption         =   "Pavimentação"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   2
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   210
               Index           =   36
               Left            =   6075
               TabIndex        =   85
               Top             =   1800
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   370
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
               Caption         =   "Arborização"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   2
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   210
               Index           =   38
               Left            =   5640
               TabIndex        =   86
               Top             =   1440
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   370
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
               Caption         =   "Sarjeta / Meio Fio"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   2
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
         End
         Begin Threed.SSFrame fra 
            Height          =   675
            Index           =   3
            Left            =   90
            TabIndex        =   87
            Top             =   2610
            Width           =   11025
            _ExtentX        =   19447
            _ExtentY        =   1191
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
            Caption         =   "Serviços"
            Alignment       =   2
            ShadowStyle     =   1
            Begin VB.TextBox txtCodComponente 
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
               Index           =   11
               Left            =   7200
               MaxLength       =   3
               TabIndex        =   16
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtCodComponente 
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
               Index           =   10
               Left            =   1650
               MaxLength       =   3
               TabIndex        =   15
               Top             =   240
               Width           =   375
            End
            Begin VB.ComboBox cboTopogr 
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
               ItemData        =   "TMPU104.frx":0094
               Left            =   2070
               List            =   "TMPU104.frx":0096
               Style           =   2  'Dropdown List
               TabIndex        =   89
               TabStop         =   0   'False
               Tag             =   "11"
               Top             =   240
               Width           =   3375
            End
            Begin VB.ComboBox cboTransporte 
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
               ItemData        =   "TMPU104.frx":0098
               Left            =   7620
               List            =   "TMPU104.frx":009A
               Style           =   2  'Dropdown List
               TabIndex        =   88
               TabStop         =   0   'False
               Tag             =   "12"
               Top             =   240
               Width           =   3375
            End
            Begin Threed.SSPanel lbl 
               Height          =   315
               Index           =   7
               Left            =   5640
               TabIndex        =   90
               Top             =   300
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   556
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
               Caption         =   "Transp. Coletivo"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin Threed.SSPanel lbl 
               Height          =   270
               Index           =   8
               Left            =   270
               TabIndex        =   91
               Top             =   300
               Width           =   1170
               _ExtentX        =   2064
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
               Caption         =   "Coleta de Lixo"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
         End
      End
      Begin MSComctlLib.ListView lstBvt 
         Height          =   1245
         Left            =   90
         TabIndex        =   98
         Top             =   1140
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   2196
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Bairro"
            Text            =   "Trecho"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Logr"
            Text            =   "DT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Nome Logr"
            Text            =   "QD"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "Lado"
            Text            =   "BA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "Valor"
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
      End
      Begin Threed.SSFrame fra 
         Height          =   1035
         Index           =   2
         Left            =   30
         TabIndex        =   92
         Top             =   45
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   1826
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
         Caption         =   "GeoLocalização"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.ComboBox cboTipoLogr 
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
            ItemData        =   "TMPU104.frx":009C
            Left            =   4410
            List            =   "TMPU104.frx":00A9
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Tag             =   "Tipo Logradouro"
            Top             =   217
            Width           =   1305
         End
         Begin VB.TextBox txtLogr 
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
            Left            =   6600
            MaxLength       =   50
            TabIndex        =   35
            Top             =   217
            Width           =   4395
         End
         Begin VB.ComboBox cboBairro 
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
            Height          =   330
            Left            =   6600
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Tag             =   "Bairro"
            Top             =   585
            Width           =   4425
         End
         Begin VB.TextBox txtCodLogr 
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
            Height          =   330
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   0
            Top             =   217
            Width           =   2025
         End
         Begin VB.TextBox txtTrecho 
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
            Height          =   330
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   1
            Top             =   585
            Width           =   765
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
            Height          =   330
            Left            =   2550
            MaxLength       =   4
            TabIndex        =   2
            Top             =   585
            Width           =   555
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
            Height          =   330
            Left            =   4425
            MaxLength       =   4
            TabIndex        =   3
            Top             =   585
            Width           =   465
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   2
            Left            =   3345
            TabIndex        =   101
            Top             =   270
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
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
            Caption         =   "Logradouro"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   1
            Left            =   6015
            TabIndex        =   102
            Top             =   645
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   397
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
            Caption         =   "Bairro"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   0
            Left            =   195
            TabIndex        =   103
            Top             =   270
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   397
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
            Caption         =   "Cod. Logr"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   4
            Left            =   165
            TabIndex        =   104
            Top             =   645
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   397
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
            Caption         =   "N°. Trecho"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   5
            Left            =   2025
            TabIndex        =   105
            Top             =   645
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   397
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
            Caption         =   "Setor"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   6
            Left            =   3735
            TabIndex        =   106
            Top             =   645
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   397
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
            Caption         =   "Quadra"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   3
            Left            =   6030
            TabIndex        =   107
            Top             =   270
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   397
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
            Caption         =   "Nome"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   525
      Index           =   1
      Left            =   0
      TabIndex        =   37
      Top             =   6945
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   926
      _Version        =   196610
      Font3D          =   3
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cboRelatorio 
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
         ItemData        =   "TMPU104.frx":00CA
         Left            =   1860
         List            =   "TMPU104.frx":00D4
         Style           =   2  'Dropdown List
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   90
         Width           =   3375
      End
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   6555
         TabIndex        =   32
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10080
         TabIndex        =   34
         Top             =   90
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
         Left            =   8895
         TabIndex        =   30
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   7710
         TabIndex        =   31
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   5310
         TabIndex        =   33
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   109
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1138
      Icone           =   "TMPU104.frx":00ED
   End
End
Attribute VB_Name = "TMPU104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodigoBairro As Long
Dim Bairro As String
Dim CodigoTipLogr As Long
Dim CodigoLogr As Long
Dim Logr As String

Private Sub AtualizaLista(Logr As String)
    Dim Sql As String
    Dim rs As VSRecordset
        
        'Tab_Imposto ON " _
        & " Tab_Geracao_Tributo.tgt_tip_cod_imposto = Tab_Imposto.tip_cod_imposto LEFT OUTER JOIN " _
        & " Tab_Darm_Recebido ON Tab_Geracao_Tributo.tgt_cod_pagamento = Tab_Darm_Recebido.tdr_tgt_cod_pagamento  " _

    Sql = "Select TTC_COD_TRECHO as Trecho,TTC_DISTRITO as DT,TTC_SETOR as ST,TTC_QUADRA as " & _
        " QD,TTC_TBA_COD_BAIRRO as BA,TTC_VALOR as Valor from " & _
        " tab_trecho where TTC_TLG_COD_LOGRADOURO = '" & txtCodLogr & "'"
    Call MontaGrid(Bdados, lstBvt, Sql, 800, 800, 800, 800, 1000)
End Sub

Public Function BuscaCodItemLogr(NomeItem As String, Grupo As String) As String
    Dim Sql As String
    Dim RsItem As VSRecordset
    Sql = "SELECT tcl_cod_componente FROM tab_componente_logradouro WHERE tcl_descricao_componente = '" & NomeItem & "'"
    Sql = Sql & " and tcl_grupo = " & Grupo
    If Bdados.AbreTabela(Sql, RsItem) Then
        BuscaCodItemLogr = RsItem(0)
    End If
    Bdados.FechaTabela RsItem
End Function

Public Sub GravaDetalheLogradouro(CodLogradouro As String, Trecho As String)
    Dim Controle As Control
    Dim rs As VSRecordset
    Dim CodItem As String
    Dim ValorItem As String
    Dim Valores As String
    Dim Campos As String
    Dim ValorGerado As Single
    
    Bdados.DeletaDados "Tab_Detalhe_Logradouro", "tdl_tlg_cod_logradouro='" & CodLogradouro & "' and tdl_num_trecho ='" & Trecho & "'"
    For Each Controle In Controls
        If IsNumeric(Controle.Tag) Then
            If (Val(Controle.Tag)) <= 20 Then
                    CodItem = BuscaCodItemLogr(Controle.Text, Controle.Tag)
                    ValorItem = Controle.ListIndex
            Else
                    CodItem = Controle.Tag
                    ValorItem = Controle.Text
            End If
            If Trim(CodItem) <> "" Then
                Valores = Bdados.PreparaValor(CodLogradouro, CodItem, Bdados.Converte(ValorItem, TCDuplo), Controle.Tag, Trecho, CodigoBairro)
                Campos = "tdl_tlg_cod_logradouro,tdl_tcl_cod_componente,tdl_valor_item,tdl_tgl_cod_grupo,tdl_num_trecho,TDL_TBA_COD_BAIRRO"
                Call Bdados.InsereDados("Tab_Detalhe_Logradouro", Valores, Campos)
                Bdados.FechaTabela rs
            End If
        End If
    Next
End Sub

Private Sub cboTipoLogr_Change()
    If cboTipoLogr <> "" Then
        CodigoTipLogr = BuscaCodigo("select ttl_cod_tip_logr from tab_tipo_logr where ttl_nome = '" & cboTipoLogr & "'")
    End If
End Sub

Private Sub cmdExcluir_Click()
    Dim ExcluiTrecho As Boolean
    Dim TempCodLog As Long
    Screen.MousePointer = 11
    ExcluiTrecho = False
    If MsgBox("Excluir o trecho " & txtTrecho & " selecionado?", vbQuestion + vbOKCancel, "Exclusão de trechos") = vbOK Then
        If Trim(txtSetor) <> "" And Trim(txtQuadra) <> "" Then
            If Bdados.DeletaDados("Tab_Detalhe_Logradouro", "tdl_tlg_cod_logradouro='" & Trim(txtCodLogr) & "' and tdl_num_trecho ='" & Trim(txtTrecho) & "'") Then
                If Not Bdados.DeletaDados("Tab_Trecho", "ttc_tlg_cod_logradouro='" & Trim(txtCodLogr) & "' and ttc_cod_trecho ='" & Trim(txtTrecho) & "' and TTC_SETOR ='" & txtSetor & "' and TTC_QUADRA = '" & txtQuadra & "'") Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            Else
                Screen.MousePointer = 0
                Exit Sub
            End If
        Else
            If Not Bdados.DeletaDados("Tab_Trecho", "ttc_tlg_cod_logradouro='" & Trim(txtCodLogr) & "' and ttc_cod_trecho ='" & Trim(txtTrecho) & "' and TTC_SETOR is null and TTC_QUADRA is null") Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        Util.Informa "Trecho " & txtTrecho & " excluido."
        TempCodLog = txtCodLogr
        Edita.LimpaCampos Me
        Util.MontaGrid Bdados, lstBvt, ""
        txtCodLogr = TempCodLog
        Call txtCodLogr_LostFocus
        txtTrecho.SetFocus
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo trata
    Dim Filtro As String
    
    Filtro = ""
    With Rpt
        Select Case cboRelatorio.ListIndex
            Case 0
                If Not .DefinirArquivo(Bdados, App.Path & "\TMPU104.rpt") Then Exit Sub
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                .Titulo = "Componentes do Cadastro Imobiliário"
                If Trim(txtCodLogr) <> "" Then .Selecao = "{TAB_LOGRADOURO.tlg_cod_logradouro} = '" & txtCodLogr & "'"
            
            Case 1
                If Not .DefinirArquivo(Bdados, App.Path & "\TBoletimInfra.rpt") Then Exit Sub
                .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name, Horizontal
                .Titulo = "Boletim de Infra-Estrutura"
                
                If Trim$(txtSetor) <> "" Then Filtro = "cdbl({VIS_INFRA.TTC_SETOR}) = cdbl(" & txtSetor & ")"
                If Trim(txtCodLogr) <> "" Then Filtro = Filtro & IIf(Filtro = "", "", " and ") & "{VIS_INFRA.tlg_cod_logradouro} ='" & txtCodLogr & "'"
            
            Case Else
                Erro "Informe o relatório."
                Set Rpt = Nothing
                Exit Sub
        End Select
        
        If Filtro <> "" Then .Selecao = Filtro
        .Visualizar
        DoEvents
    End With
    Set Rpt = Nothing
    Exit Sub
    
trata:
    'Erro Err.Description
End Sub

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    txtCodLogr.SetFocus
    SSTab1.Tab = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Sql As String
    Dim Campos As String
    Dim rs As VSRecordset
    If Edita.CriticaCampos(Me) Then
        CodigoBairro = BuscaCodigo("select tba_cod_bairro from tab_bairro where tba_nome = '" & cboBairro & "' and tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio)
        CodigoTipLogr = BuscaCodigo("select ttl_cod_tip_Logr from tab_tipo_logr where ttl_nome = '" & Me.cboTipoLogr & "'")
        
        'Campos = "tlg_tmu_cod_municipio,tlg_cod_logradouro,tlg_tba_cod_bairro,tlg_ttl_cod_tip_logr,tlg_nome,tlg_secao,tlg_ttr_cod_trecho,tlg_quadra"
        'Valores = Bdados.PreparaValor(Aplicacoes.Codigo_Municipio, txtCodLogr, CodigoBairro, CodigoTipLogr, txtLogr, txtSetor, txtTrecho, txtQuadra)
        'Call Bdados.GravaDados( "Tab_Logradouro", Valores, Campos, "tlg_tba_cod_bairro = " & CodigoBairro & " and tlg_ttl_cod_tip_logr = " & CodigoTipLogr)
        GravaDetalheLogradouro txtCodLogr, txtTrecho
        Call Util.Informa("Transação completada.")
        Dim Aux As String
        Aux = txtCodLogr
        txtCodLogr = Aux
        txtCodLogr_LostFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Dim Controle As Control
    Dim i As Byte
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    Call Edita.AtualizaCombo(Bdados, cboTipoLogr, "Select ttl_nome From Tab_Tipo_Logr")
    Call Edita.AtualizaCombo(Bdados, cboBairro, "Select tba_nome From Tab_Bairro where tba_tmu_cod_municipio=" & Temp.PegaParametro(Bdados, "MUNICIPIO"))
    For Each Controle In Controls
        If IsNumeric(Controle.Tag) Then
            If Val(Controle.Tag) <= 20 Then Call Edita.AtualizaCombo(Bdados, Controle, "Select tcL_descricao_componente From Tab_Componente_LOGRADOURO Where tcL_grupo = " & Controle.Tag & " order by tcl_cod_componente asc")
        End If
    Next
    Screen.MousePointer = 0
End Sub

Private Sub lstBvt_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstBvt, ColumnHeader
End Sub

Private Sub lstBvt_DblClick()
    On Error Resume Next
    txtTrecho = lstBvt.SelectedItem
    txtSetor = lstBvt.SelectedItem.SubItems(2)
    txtQuadra = lstBvt.SelectedItem.SubItems(3)
    txtTrecho_LostFocus
End Sub

Private Sub tmrLog_Timer()
    On Error Resume Next
    
End Sub

Private Sub txtCodComponente_Change(Index As Integer)
    Dim Controle As Control
    On Error GoTo trata
    
    For Each Controle In Controls
        If Controle.Tag = Index + 1 Then
            Controle.ListIndex = Util.Nvl(txtCodComponente(Index).Text, 0) - 1
            Exit For
        End If
    Next
trata:
    If Err.Number = 380 Then
        Avisa "Valor inválido."
        txtCodComponente(Index).SetFocus
    End If
End Sub

Private Sub txtCodLogr_LostFocus()
    Dim rs As VSRecordset
    Dim Sql As String
    
    If Trim(txtCodLogr) <> "" Then

        Sql = "Select TTL_NOME,tlg_nome from tab_logradouro, tab_tipo_logr " & _
            " where tlg_cod_logradouro='" & txtCodLogr & "' and tlg_ttl_cod_tip_logr = " & _
            " TTL_COD_TIP_LOGR "
        If Bdados.AbreTabela(Sql, rs) Then
            cboTipoLogr.Text = rs(0)
            txtLogr = rs(1)
        Else
            Avisa "Código de logradouro inexistente."
            txtCodLogr.SetFocus
            Exit Sub
        End If
        AtualizaLista txtCodLogr
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtLogr_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSetor_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtTrecho_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTrecho_LostFocus()
    Dim rs As VSRecordset
    Dim Sql As String
    Dim CodItem  As String
    Dim Controle As Control
    If Trim(txtTrecho) <> "" Then
        Sql = "Select TTC_COD_TRECHO ,TTC_DISTRITO ,tba_nome,ttc_setor,ttc_quadra from tab_bairro,tab_trecho where TTC_TLG_COD_LOGRADOURO ='" & txtCodLogr & "' and TTC_COD_TRECHO='" & txtTrecho & "' and TTC_TBA_COD_BAIRRO =tba_cod_bairro and tba_tmu_cod_municipio=" & Temp.PegaParametro(Bdados, "MUNICIPIO")
        If Bdados.AbreTabela(Sql, rs) Then
            cboBairro.Text = "" & rs!TBA_NOME
'            txtSetor = "" & Rs!TTC_SETOR
'            txtQuadra = "" & Rs!TTC_QUADRA
            Sql = "Select * from tab_detalhe_logradouro where tdl_tlg_cod_logradouro='" & txtCodLogr & "' and tdl_num_trecho = '" & txtTrecho & "'"
            If Bdados.AbreTabela(Sql, rs) Then
                rs.MoveFirst
                Do
                    For Each Controle In Controls
                        If IsNumeric(Controle.Tag) Then
                            If (Val(Controle.Tag)) <= 20 Then
                                    If CDbl(Controle.Tag) = rs!tdl_tgl_cod_grupo Then
                                        txtCodComponente(CInt(Controle.Tag) - 1).Text = rs!tdl_tcl_cod_componente
                                    End If
                            Else
                                    If CDbl(Controle.Tag) = rs!tdl_tgl_cod_grupo Then
                                        Controle.Text = rs!tdl_valor_item
                                    End If
                            End If
                        End If
                    Next
                    rs.MoveNext
                Loop While Not rs.EOF
            End If
        Else
            Avisa "Trecho não definido."
            'txtTrecho.SetFocus
            DoEvents
        End If
    End If
    Bdados.FechaTabela rs
End Sub
