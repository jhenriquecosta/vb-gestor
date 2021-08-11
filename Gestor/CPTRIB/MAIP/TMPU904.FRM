VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TMPU904 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TMPU904"
   ClientHeight    =   3135
   ClientLeft      =   1125
   ClientTop       =   870
   ClientWidth     =   6855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog DLG 
      Left            =   2280
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton optOpcao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Infra-estruturas (INFRA)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1590
      Width           =   6075
   End
   Begin VTOcx.txtVISUAL txtCaminho 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   2220
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   556
      Caption         =   "Caminho"
      Text            =   ""
      RetirarMascara  =   0   'False
   End
   Begin VB.OptionButton optOpcao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Distritos (SIMPRO)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   1290
      Width           =   6075
   End
   Begin VB.OptionButton optOpcao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Trechos (CADLOGTRECHO)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   1005
      Width           =   6075
   End
   Begin VB.OptionButton optOpcao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Logradouros (CADLOGNOME/CADLOTNOME)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Value           =   -1  'True
      Width           =   6075
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   7
      Top             =   2610
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   926
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Top             =   90
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
         CorFoco         =   -2147483626
      End
      Begin VTOcx.cmdVISUAL cmdImportar 
         Height          =   375
         Left            =   4890
         TabIndex        =   6
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "&Importar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
         CorFoco         =   -2147483626
      End
   End
   Begin VTOcx.cmdVISUAL cmdBuscar 
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   2190
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      Acao            =   5
      CorBorda        =   -2147483626
      CorFrente       =   -2147483626
      CorFoco         =   -2147483626
   End
   Begin VB.PictureBox SSTab1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   6615
      TabIndex        =   9
      Top             =   2700
      Width           =   6675
      Begin VB.TextBox txtOutros 
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
         Left            =   -70110
         MaxLength       =   3
         TabIndex        =   57
         Tag             =   "28"
         Top             =   1830
         Width           =   645
      End
      Begin VB.TextBox txtUrbanizacao 
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
         Left            =   -70800
         MaxLength       =   3
         TabIndex        =   56
         Tag             =   "27"
         Top             =   1830
         Width           =   645
      End
      Begin VB.TextBox txtInundacao 
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
         Left            =   -71490
         MaxLength       =   3
         TabIndex        =   55
         Tag             =   "26"
         Top             =   1830
         Width           =   645
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
         ItemData        =   "TMPU904.frx":0000
         Left            =   -74520
         List            =   "TMPU904.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   54
         TabStop         =   0   'False
         Tag             =   "18"
         Top             =   1440
         Width           =   1965
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
         Left            =   -74940
         MaxLength       =   3
         TabIndex        =   53
         Top             =   1440
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
         ItemData        =   "TMPU904.frx":0004
         Left            =   -72090
         List            =   "TMPU904.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   52
         TabStop         =   0   'False
         Tag             =   "19"
         Top             =   1080
         Width           =   1965
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
         Left            =   -72510
         MaxLength       =   3
         TabIndex        =   51
         Top             =   1080
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
         Index           =   19
         Left            =   -72510
         MaxLength       =   3
         TabIndex        =   50
         Top             =   1440
         Width           =   375
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
         Left            =   -72870
         MaxLength       =   3
         TabIndex        =   49
         Tag             =   "25"
         Top             =   1830
         Width           =   645
      End
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
         ItemData        =   "TMPU904.frx":0008
         Left            =   -72090
         List            =   "TMPU904.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   48
         TabStop         =   0   'False
         Tag             =   "20"
         Top             =   1440
         Width           =   1965
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
         Left            =   -74910
         MaxLength       =   3
         TabIndex        =   47
         Tag             =   "21"
         Top             =   1830
         Width           =   645
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
         Left            =   -73545
         MaxLength       =   3
         TabIndex        =   46
         Tag             =   "22"
         Top             =   1830
         Width           =   645
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
         Left            =   -74235
         MaxLength       =   3
         TabIndex        =   45
         Tag             =   "23"
         Top             =   1830
         Width           =   645
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
         Left            =   -72180
         MaxLength       =   3
         TabIndex        =   44
         Tag             =   "24"
         Top             =   1830
         Width           =   645
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
         ItemData        =   "TMPU904.frx":000C
         Left            =   -72090
         List            =   "TMPU904.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   43
         TabStop         =   0   'False
         Tag             =   "17"
         Top             =   720
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":0010
         Left            =   -72090
         List            =   "TMPU904.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "16"
         Top             =   360
         Width           =   1965
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
         Left            =   -72510
         MaxLength       =   3
         TabIndex        =   41
         Top             =   360
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
         Index           =   16
         Left            =   -72510
         MaxLength       =   3
         TabIndex        =   40
         Top             =   720
         Width           =   375
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
         ItemData        =   "TMPU904.frx":0014
         Left            =   -74520
         List            =   "TMPU904.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "15"
         Top             =   1080
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":0018
         Left            =   -74520
         List            =   "TMPU904.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   38
         TabStop         =   0   'False
         Tag             =   "14"
         Top             =   720
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":001C
         Left            =   -74520
         List            =   "TMPU904.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   37
         TabStop         =   0   'False
         Tag             =   "13"
         Top             =   360
         Width           =   1965
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
         Left            =   -74940
         MaxLength       =   3
         TabIndex        =   36
         Top             =   360
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
         Left            =   -74940
         MaxLength       =   3
         TabIndex        =   35
         Top             =   720
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
         Left            =   -74940
         MaxLength       =   3
         TabIndex        =   34
         Top             =   1080
         Width           =   375
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
         ItemData        =   "TMPU904.frx":0020
         Left            =   2940
         List            =   "TMPU904.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "12"
         Top             =   2115
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":0024
         Left            =   480
         List            =   "TMPU904.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "11"
         Top             =   2115
         Width           =   1965
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
         Left            =   60
         MaxLength       =   3
         TabIndex        =   31
         Top             =   2130
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
         Index           =   11
         Left            =   2550
         MaxLength       =   3
         TabIndex        =   30
         Top             =   2130
         Width           =   375
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
         ItemData        =   "TMPU904.frx":0028
         Left            =   2955
         List            =   "TMPU904.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "9"
         Top             =   1410
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":002C
         Left            =   2955
         List            =   "TMPU904.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "10"
         Top             =   1770
         Width           =   1965
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
         Left            =   2550
         MaxLength       =   3
         TabIndex        =   27
         Top             =   1410
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
         Left            =   2550
         MaxLength       =   3
         TabIndex        =   26
         Top             =   1770
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
         Index           =   7
         Left            =   2550
         MaxLength       =   3
         TabIndex        =   25
         Top             =   1050
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
         ItemData        =   "TMPU904.frx":0030
         Left            =   2970
         List            =   "TMPU904.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "8"
         Top             =   1050
         Width           =   1965
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
         Left            =   2550
         MaxLength       =   3
         TabIndex        =   23
         Top             =   690
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
         Index           =   5
         Left            =   2550
         MaxLength       =   3
         TabIndex        =   22
         Top             =   330
         Width           =   375
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
         ItemData        =   "TMPU904.frx":0034
         Left            =   2970
         List            =   "TMPU904.frx":0036
         Style           =   2  'Dropdown List
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "6"
         Top             =   330
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":0038
         Left            =   2970
         List            =   "TMPU904.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "7"
         Top             =   690
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":003C
         Left            =   480
         List            =   "TMPU904.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "5"
         Top             =   1770
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":0040
         Left            =   480
         List            =   "TMPU904.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "4"
         Top             =   1410
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":0044
         Left            =   480
         List            =   "TMPU904.frx":0046
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "3"
         Top             =   1050
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":0048
         Left            =   480
         List            =   "TMPU904.frx":004A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   690
         Width           =   1965
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
         ItemData        =   "TMPU904.frx":004C
         Left            =   480
         List            =   "TMPU904.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   330
         Width           =   1965
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
         Left            =   60
         MaxLength       =   3
         TabIndex        =   14
         Top             =   330
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
         Left            =   60
         MaxLength       =   3
         TabIndex        =   13
         Top             =   690
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
         Left            =   60
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1050
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
         Left            =   60
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1410
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
         Index           =   4
         Left            =   60
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1770
         Width           =   375
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1138
      Icone           =   "TMPU904.frx":0050
   End
End
Attribute VB_Name = "TMPU904"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscar_Click()
    'dlg.FileName = ""
    DLG.Filter = "BD Access (*.mdb)|*.mdb"
    DLG.InitDir = App.Path
    DLG.ShowOpen
    txtCaminho = DLG.FileName
    cmdImportar.SetFocus
End Sub

Private Sub cmdImportar_Click()
    Dim Opcao As OptionButton
    Dim Qtd As Integer
    Dim retorno As Boolean
    
    If Trim$(txtCaminho) = "" Then
        Erro "Informe o caminho do banco."
        Exit Sub
    End If
    If Confirma("Importar dados de " & DLG.FileTitle & " ?") Then
        For Each Opcao In optOpcao
            If Opcao.Value Then
                Select Case Opcao.Index
                    Case 0, 1 'Logradouros/Loteamentos
                        retorno = ImportarLogradouros(txtCaminho, Qtd)
                    Case 2 'Trechos
                        retorno = ImportarTrechos(txtCaminho, Qtd)
                    Case 3 'Distritos
                        retorno = ImportarDistritos(txtCaminho, Qtd)
                    Case 4 'Infra
                        retorno = ImportarInfra(txtCaminho, Qtd)
                End Select
                If retorno Then
                    Avisa Qtd & " " & Util.ParseString(Opcao.Caption, "(", 1) & " importados com sucesso."
                Else
                    Erro "erro."
                End If
                Exit For
            End If
        Next
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Controle As Object
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    For Each Controle In Controls
        If IsNumeric(Controle.Tag) Then
            If Val(Controle.Tag) <= 20 Then Call Edita.AtualizaCombo(Bdados, Controle, "Select tcL_descricao_componente From Tab_Componente_LOGRADOURO Where tcL_grupo = " & Controle.Tag & " order by tcl_cod_componente asc")
        End If
    Next
End Sub



Private Function ImportarLogradouros(CaminhoBD As String, ByRef Qtd As Integer) As Boolean
'ROTEIRO
'1. Grava as informacoes do tipo de logradouro
'2. Grava as informacoes do bairro
'3. Grava as informacoes do logradouro
    On Error GoTo Erro
    Dim BD As New VSDados
    Dim rs As VSRecordset
    
    Dim Valores As String
    Dim Campos As String
    Dim CodTipLogr As String
    Dim iQtd As Integer
    
    iQtd = 0
    Screen.MousePointer = vbArrowHourglass
    If BD.AbreBanco(Access, CaminhoBD, "admin") Then
        If BD.AbreTabela("SELECT * FROM [CADLOG]", rs) Then
            Do While Not rs.EOF
                '1.
                CodTipLogr = Bdados.BuscaCodigo("SELECT TTL_COD_TIP_LOGR FROM TAB_TIPO_LOGR WHERE TTL_NOME='" & TiraPic(rs!TIPOLOG, "'") & "'")
                If CodTipLogr = "0" Then
                    CodTipLogr = Bdados.BuscaCodigo("SELECT MAX(TTL_COD_TIP_LOGR)+1 FROM TAB_TIPO_LOGR")
                    If CodTipLogr = "0" Then CodTipLogr = "1"
                    Campos = "TTL_COD_TIP_LOGR,TTL_NOME"
                    Valores = Bdados.PreparaValor(CodTipLogr, rs!TIPOLOG)
                    Bdados.GravaDados "TAB_TIPO_LOGR", Valores, Campos, "TTL_COD_TIP_LOGR=" & CodTipLogr
                End If
            
                '2.
                Campos = "TBA_TMU_COD_MUNICIPIO,TBA_COD_BAIRRO,TBA_NOME"
                Valores = Bdados.PreparaValor(Aplicacoes.Codigo_Municipio, rs!CODBAILOG, rs!NOMEBAILOG)
                Bdados.GravaDados "TAB_BAIRRO", Valores, Campos, "TBA_TMU_COD_MUNICIPIO=" & Aplicacoes.Codigo_Municipio & " AND TBA_COD_BAIRRO=" & rs!CODBAILOG
                
                '3.
                Campos = "tlg_tmu_cod_municipio,tlg_tba_cod_bairro,tlg_ttl_cod_tip_logr,tlg_cod_logradouro,tlg_nome,tlg_cod_logradouro_inicial,tlg_cod_logradouro_final,tlg_cod_bairro_inicial,tlg_cod_bairro_final"
                Valores = Bdados.PreparaValor(Aplicacoes.Codigo_Municipio, rs!CODBAILOG, CodTipLogr, rs!CODLOG, TiraPic(rs!NOMELOG, "'"), rs!CODLOGINI, rs!CODLOGFIM, rs!CODBAIINI, rs!CODBAIFIM)
                Bdados.GravaDados "TAB_LOGRADOURO", Valores, Campos, "tlg_cod_logradouro='" & TiraPic(rs!CODLOG, "'") & "'"
                iQtd = iQtd + 1
                cmdImportar.Caption = iQtd & "   . . .  "
                rs.MoveNext
            Loop
            ImportarLogradouros = True
            cmdImportar.Caption = "&Importar"
        End If
        BD.FechaTabela rs
    End If
    BD.FechaBanco
    Screen.MousePointer = vbNormal
    Qtd = iQtd
    Exit Function
Erro:
    Avisa Err.Description
    'Resume
End Function

Private Function ImportarTrechos(Caminho As String, ByRef Qtd As Integer) As Boolean
'ROTEIRO
'1. Grava as informacoes do trecho
    On Error GoTo Erro
    Dim BD As New VSDados
    Dim rs As VSRecordset
    
    Dim Valores As String
    Dim Campos As String
    Dim iQtd As Integer
    
    iQtd = 0
    Screen.MousePointer = vbArrowHourglass
    If BD.AbreBanco(Access, Caminho, "admin") Then
        If BD.AbreTabela("SELECT * FROM [CADLOGTRECHO] ORDER BY [CDLOG], [NºT]", rs) Then
            Do While Not rs.EOF
                '1.
                Campos = "TTC_TLG_COD_LOGRADOURO, TTC_COD_TRECHO, TTC_SETOR, TTC_QUADRA, TTC_TBA_COD_BAIRRO, TTC_LOGR_INICIAL, TTC_LOGR_FINAL, TTC_SEQ_TRECHO"
                Valores = Bdados.PreparaValor(rs!CDLOG, rs![NºT] & rs!LD, rs!ST, rs!QD, rs!CDBAI, rs!CDINI, rs!CDFIM, rs![NºT])
                Bdados.GravaDados "TAB_TRECHO", Valores, Campos, "TTC_TLG_COD_LOGRADOURO='" & rs!CDLOG & "' AND TTC_COD_TRECHO='" & rs![NºT] & "'"
                
                iQtd = iQtd + 1
                cmdImportar.Caption = iQtd & "   . . .  "
                rs.MoveNext
            Loop
            ImportarTrechos = True
            cmdImportar.Caption = "&Importar"
        End If
        BD.FechaTabela rs
    End If
    BD.FechaBanco
    Screen.MousePointer = vbNormal
    Qtd = iQtd
    Exit Function
Erro:
    Avisa Err.Description
End Function


Private Function ImportarDistritos(Caminho As String, ByRef Qtd As Integer) As Boolean
'ROTEIRO
'1. Atualiza as informacoes do distrito
    On Error GoTo Erro
    Dim BD As New VSDados
    Dim rs As VSRecordset
    
    Dim Valores As String
    Dim Campos As String
    Dim iQtd As Integer
    
    iQtd = 0
    Screen.MousePointer = vbArrowHourglass
    If BD.AbreBanco(Access, Caminho, "admin") Then
        If BD.AbreTabela("SELECT * FROM [SIMPRO]", rs) Then
            Do While Not rs.EOF
                '1.
                Campos = "TTC_DISTRITO"
                Valores = Bdados.PreparaValor(rs![Distrito Novo])
                Bdados.AtualizaDados "TAB_TRECHO", Valores, Campos, "cast(TTC_QUADRA as smallint)=" & Val(rs!QuadraNova) & " AND CAST(TTC_SETOR AS smallint)=" & Val(rs!SetorNovo)
                
                iQtd = iQtd + 1
                cmdImportar.Caption = iQtd & "   . . .  "
                rs.MoveNext
            Loop
            ImportarDistritos = True
            cmdImportar.Caption = "&Importar"
        End If
        BD.FechaTabela rs
    End If
    BD.FechaBanco
    Screen.MousePointer = vbNormal
    Qtd = iQtd
    Exit Function
Erro:
    Avisa Err.Description
End Function


Private Function ImportarInfra(Caminho As String, ByRef Qtd As Integer) As Boolean
'ROTEIRO
'1. Exibir informacoes nos controles
'2. Atualizar as informacoes de infra-estrutura dos logradouros
    On Error GoTo Erro
    Dim BD As New VSDados
    Dim rs As VSRecordset
    
    Dim Valores As String
    Dim Campos As String
    Dim iQtd As Integer
    
    iQtd = 0
    Screen.MousePointer = vbArrowHourglass
    If BD.AbreBanco(Access, Caminho, "admin") Then
        If BD.AbreTabela("SELECT * FROM [INFRA]", rs) Then
            Do While Not rs.EOF
                '1.
                ExibirInfra rs
                
                '2.
                GravaDetalheLogradouro rs!CÓDLOG, rs![Nº TRECHO] & rs![LADO E/D], "" & rs!CODBAI
                iQtd = iQtd + 1
                cmdImportar.Caption = iQtd & "   . . .  "
                rs.MoveNext
            Loop
            ImportarInfra = True
            cmdImportar.Caption = "&Importar"
        End If
        BD.FechaTabela rs
    End If
    BD.FechaBanco
    Screen.MousePointer = vbNormal
    Qtd = iQtd
    Exit Function
Erro:
    Avisa Err.Description
    Resume
End Function



Private Sub optOpcao_Click(Index As Integer)
    Dim Opcao As OptionButton
    
    For Each Opcao In optOpcao
        Opcao.Font.Bold = False
        Opcao.Font.Size = 7
    Next
    optOpcao(Index).Font.Bold = True
    optOpcao(Index).Font.Size = 8
    cmdBuscar.SetFocus
End Sub

Private Sub ExibirInfra(rs As VSRecordset)
    On Error GoTo trata
    
    Dim CodGrupo As Integer
    'SSTab1.Tab = 0
    '1. Agua
    CodGrupo = 0
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Rede de Água], 0)
    
    '2. Esgoto
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Rede de Esgoto], 0)
    
    '3. Eletrica
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Rede Elétrica], 0)
    
    '4. Telefone
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Rede Telefônica], 0)
    
    '5. Calcada
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Calçada], 0)
    
    '6. Drenagem
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Drenagem], 0)
    
    '7. Iluminacao
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Ilumin Pública], 0)
    
    '8. Pavimentacao
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Pavimentação], 0)
    
    '9. Sarjeta
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Sarjeta/Meio-fio], 0)
    
    '10. Arborizacao
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Arborização], 0)
    
    '11. Lixo
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Coleta de lixo], 0)
    
    '12. Coletivo
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Transp Coletivo], 0)
    
    'SSTab1.Tab = 1
    '13. Estacionamento
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Estacionamento], 0)
    
    '14. Limpeza
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Limp Pública], 0)
    
    '15. Conservacao
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Conservação de vias], 0)
    
    '16. Servicos
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Pontos de serviços], 0)
    
    '17. Hidrante
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Hidrante], 0)
    
    '18. Ciclovia
    CodGrupo = CodGrupo + 1
    
    '19. Via
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Tipo de via], 0)
    
    '20. Logradouro
    CodGrupo = CodGrupo + 1
    txtCodComponente(CodGrupo) = Nvl("" & rs.Fields![Tipo logradouro], 0)
    
    '21. Larg via
    txtVia = Nvl("" & rs.Fields![Largura da via], 0)
    
    '22. Larg via (proj)
    txtViaProj = 0
    
    '23. Larg calcada
    txtCalcada = Nvl("" & rs![Largura da calçada], 0)
    
    '24. Larg canteiro
    txtCanteiro = Nvl("" & rs![Largura canteiro], 0)
    
    '25. Uso solo
    Text1 = Nvl("" & rs![Uso do solo], 0)

    '26. Cota inundacao
    txtInundacao = Nvl("" & rs![Cota inundação], 0)

    '27. Urbanizacao
    txtUrbanizacao = Nvl("" & rs![Urbanização], 0)

    '28. Outros
    txtOutros = Nvl("" & rs![Outros], 0)
    
    Exit Sub
trata:
    Erro Err.Description
    'Resume
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
        'Avisa "Valor inválido."
        'txtCodComponente(Index).SetFocus
    End If
End Sub

Public Sub GravaDetalheLogradouro(CodLogradouro As String, Trecho As String, CodigoBairro As String)
    Dim Controle As Control
    
    Dim CodItem As String
    Dim ValorItem As String
    
    Dim Valores As String
    Dim Campos As String
    
    If CodLogradouro <> "" And Trecho <> "" And CodigoBairro <> "" Then
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
                If Trim(CodItem) = "" Then
                    CodItem = 0
                    ValorItem = 0
                End If
                Campos = "tdl_tlg_cod_logradouro,tdl_tcl_cod_componente,tdl_valor_item,tdl_tgl_cod_grupo,tdl_num_trecho,TDL_TBA_COD_BAIRRO"
                Valores = Bdados.PreparaValor(CodLogradouro, CodItem, Bdados.Converte(ValorItem, TCDuplo), Controle.Tag, Trecho, CodigoBairro)
                Bdados.InsereDados "Tab_Detalhe_Logradouro", Valores, Campos
            End If
        Next
    End If
End Sub

Public Function BuscaCodItemLogr(NomeItem As String, Grupo As String) As String
    Dim Sql As String
    Dim RsItem As VSRecordset
    Sql = "SELECT tcl_cod_componente FROM tab_componente_logradouro WHERE tcl_descricao_componente = '" & TiraPic(NomeItem, "'") & "'"
    Sql = Sql & " and tcl_grupo = " & Grupo
    If Bdados.AbreTabela(Sql, RsItem) Then
        BuscaCodItemLogr = RsItem(0)
    End If
    Bdados.FechaTabela RsItem
End Function

