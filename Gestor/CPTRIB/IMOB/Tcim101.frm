VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Begin VB.Form TCIM101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabCad 
      Height          =   6015
      Left            =   45
      TabIndex        =   103
      Top             =   690
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "BT (Boletim Territorial)"
      TabPicture(0)   =   "Tcim101.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstPesq"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "BT (cont.)"
      TabPicture(1)   =   "Tcim101.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra(4)"
      Tab(1).Control(1)=   "fra(3)"
      Tab(1).Control(2)=   "fra(5)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "BP (Boletim Predial)"
      TabPicture(2)   =   "Tcim101.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra(6)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "BC (Boletim de Condomínio)"
      TabPicture(3)   =   "Tcim101.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra(7)"
      Tab(3).Control(1)=   "cmdAdCond"
      Tab(3).Control(2)=   "lstCond"
      Tab(3).Control(3)=   "fra(8)"
      Tab(3).Control(4)=   "fra(2)"
      Tab(3).ControlCount=   5
      Begin Threed.SSFrame fra 
         Height          =   1755
         Index           =   0
         Left            =   120
         TabIndex        =   104
         Top             =   540
         Width           =   11085
         _ExtentX        =   19553
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
         Caption         =   "01 - Referência Cadastral / 02 - Localização do Imóvel"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtDescMens 
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
            Left            =   1740
            TabIndex        =   247
            Tag             =   "Nome Contribuinte"
            Top             =   1380
            Width           =   4215
         End
         Begin VB.TextBox txtZona 
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
            Left            =   7890
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Zona"
            Top             =   240
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
            Left            =   4830
            TabIndex        =   5
            Top             =   240
            Width           =   2505
         End
         Begin VB.TextBox txtIc 
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
            Index           =   4
            Left            =   2760
            MaxLength       =   3
            TabIndex        =   4
            Tag             =   "Unidade"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtIc 
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
            Index           =   3
            Left            =   2250
            MaxLength       =   4
            TabIndex        =   3
            Tag             =   "Lote"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtIc 
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
            Index           =   2
            Left            =   1830
            MaxLength       =   3
            TabIndex        =   2
            Tag             =   "Quadra"
            Top             =   240
            Width           =   405
         End
         Begin VB.TextBox txtIc 
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
            Index           =   1
            Left            =   1500
            MaxLength       =   2
            TabIndex        =   1
            Tag             =   "Setor"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtIc 
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
            Index           =   0
            Left            =   1170
            MaxLength       =   2
            TabIndex        =   0
            Tag             =   "Distrito"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtCodMens 
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
            Left            =   1155
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "Cod Mensagem"
            Top             =   1380
            Width           =   525
         End
         Begin VB.TextBox txtCodBairro 
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
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   13
            Tag             =   "Bairro"
            Top             =   1012
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
            Left            =   2850
            MaxLength       =   11
            TabIndex        =   9
            Top             =   630
            Width           =   1140
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
            Left            =   4050
            TabIndex        =   10
            Tag             =   "Nome Contribuinte"
            Top             =   630
            Width           =   3765
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
            Left            =   1740
            TabIndex        =   14
            Tag             =   "Nome Contribuinte"
            Top             =   1020
            Width           =   4215
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
            Height          =   315
            Left            =   1170
            TabIndex        =   8
            Tag             =   "Logradouro"
            Top             =   630
            Width           =   1185
         End
         Begin VB.ComboBox cboTipoImovel 
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
            ItemData        =   "Tcim101.frx":0070
            Left            =   9060
            List            =   "Tcim101.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "Tipo Imovel"
            Top             =   232
            Width           =   1935
         End
         Begin VB.TextBox txtNumero 
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
            Left            =   10440
            MaxLength       =   10
            TabIndex        =   12
            Top             =   630
            Width           =   555
         End
         Begin VB.TextBox txtComplemento 
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
            Left            =   8580
            TabIndex        =   11
            Top             =   630
            Width           =   1425
         End
         Begin VB.TextBox txtLoteamento 
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
            Left            =   7080
            MaxLength       =   5
            TabIndex        =   15
            Top             =   1020
            Width           =   1095
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
            Left            =   8580
            MaxLength       =   5
            TabIndex        =   16
            Top             =   1020
            Width           =   975
         End
         Begin VB.TextBox txtLote 
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
            Left            =   10080
            MaxLength       =   5
            TabIndex        =   17
            Top             =   1020
            Width           =   915
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   2
            Left            =   7890
            TabIndex        =   105
            Top             =   675
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   1
            Left            =   10140
            TabIndex        =   106
            Top             =   675
            Width           =   390
            _ExtentX        =   688
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
            Caption         =   "N.º:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   3
            Left            =   570
            TabIndex        =   107
            Top             =   1065
            Width           =   705
            _ExtentX        =   1244
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
            Caption         =   "Bairro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   4
            Left            =   6015
            TabIndex        =   108
            Top             =   1065
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Loteamento:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   5
            Left            =   8250
            TabIndex        =   109
            Top             =   1065
            Width           =   660
            _ExtentX        =   1164
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
            Caption         =   "Qd.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   6
            Left            =   9630
            TabIndex        =   110
            Top             =   1065
            Width           =   435
            _ExtentX        =   767
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
            Caption         =   "Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   7
            Left            =   8580
            TabIndex        =   120
            Top             =   285
            Width           =   450
            _ExtentX        =   794
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
            Caption         =   "Tipo:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   82
            Left            =   270
            TabIndex        =   124
            Top             =   675
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
            Caption         =   "Cód. Logr:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   50
            Left            =   135
            TabIndex        =   125
            Top             =   1425
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   72
            Left            =   3570
            TabIndex        =   126
            Top             =   285
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Insc. Anterior:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   75
            Left            =   285
            TabIndex        =   127
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Insc. Cad.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   76
            Left            =   7410
            TabIndex        =   128
            Top             =   285
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Zona:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   345
            Index           =   2
            Left            =   2370
            TabIndex        =   230
            Top             =   600
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1785
         Index           =   1
         Left            =   90
         TabIndex        =   111
         Top             =   2940
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   3149
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
         Caption         =   "03 - Dados do Proprietário"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtCodBairroContrib 
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
            Left            =   720
            TabIndex        =   36
            Tag             =   "Logradouro"
            Top             =   975
            Width           =   585
         End
         Begin VB.TextBox txtUf 
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
            Left            =   10350
            MaxLength       =   50
            TabIndex        =   40
            Tag             =   "Bairro"
            Top             =   970
            Width           =   615
         End
         Begin VB.TextBox txtNomeTipoLogrContrib 
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
            Left            =   1650
            MaxLength       =   11
            TabIndex        =   32
            Top             =   585
            Width           =   1035
         End
         Begin VB.TextBox txtCodLogrContrib 
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
            Left            =   705
            TabIndex        =   31
            Tag             =   "Logradouro"
            Top             =   585
            Width           =   915
         End
         Begin VB.TextBox txtMunic 
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
            Left            =   5910
            TabIndex        =   39
            Tag             =   "Município"
            Top             =   970
            Width           =   4395
         End
         Begin VB.TextBox txtNumeroContrib 
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
            Left            =   6255
            MaxLength       =   10
            TabIndex        =   34
            Top             =   585
            Width           =   465
         End
         Begin VB.TextBox txtCpfCgc 
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
            Left            =   9210
            MaxLength       =   20
            TabIndex        =   30
            Top             =   210
            Width           =   1755
         End
         Begin VB.TextBox txtCpfOcupante 
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
            Left            =   8940
            MaxLength       =   20
            TabIndex        =   42
            Top             =   1350
            Width           =   2025
         End
         Begin VB.TextBox txtOcupante 
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
            Left            =   3330
            TabIndex        =   41
            Top             =   1350
            Width           =   4575
         End
         Begin VB.TextBox txtBairroContrib 
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
            Left            =   1350
            TabIndex        =   37
            Tag             =   "Bairro"
            Top             =   975
            Width           =   2235
         End
         Begin VB.CommandButton cmdEnter 
            Caption         =   "Command1"
            Default         =   -1  'True
            Height          =   255
            Left            =   7740
            TabIndex        =   117
            Top             =   3090
            Width           =   375
         End
         Begin VB.TextBox txtCep 
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
            Left            =   4050
            MaxLength       =   10
            TabIndex        =   38
            Top             =   970
            Width           =   885
         End
         Begin VB.TextBox txtNomeLogrContrib 
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
            Left            =   2700
            TabIndex        =   33
            Tag             =   "Nome Logradouro"
            Top             =   585
            Width           =   3195
         End
         Begin VB.TextBox txtNomeContrib 
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
            Left            =   690
            TabIndex        =   26
            Tag             =   "Nome Contribuinte"
            Top             =   210
            Width           =   4665
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
            Left            =   6570
            MaxLength       =   11
            TabIndex        =   29
            Top             =   210
            Width           =   1305
         End
         Begin VB.TextBox txtCompContrib 
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
            Left            =   7530
            TabIndex        =   35
            Top             =   585
            Width           =   3435
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   8
            Left            =   105
            TabIndex        =   112
            Top             =   255
            Width           =   525
            _ExtentX        =   926
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
            Caption         =   "Nome:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   14
            Left            =   3630
            TabIndex        =   113
            Top             =   1020
            Width           =   375
            _ExtentX        =   661
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
            Caption         =   "CEP:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   15
            Left            =   5025
            TabIndex        =   114
            Top             =   1020
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Municipio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   13
            Left            =   5970
            TabIndex        =   115
            Top             =   600
            Width           =   270
            _ExtentX        =   476
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
            Caption         =   "N.º:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   17
            Left            =   120
            TabIndex        =   116
            Top             =   1020
            Width           =   555
            _ExtentX        =   979
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
            Caption         =   "Bairro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   16
            Left            =   6870
            TabIndex        =   119
            Top             =   600
            Width           =   1290
            _ExtentX        =   2275
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   29
            Left            =   8355
            TabIndex        =   121
            Top             =   255
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   11
            Left            =   2445
            TabIndex        =   122
            Top             =   1395
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "Ocupante:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   18
            Left            =   8085
            TabIndex        =   123
            Top             =   1425
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   228
            Index           =   49
            Left            =   180
            TabIndex        =   225
            Top             =   600
            Width           =   444
            _ExtentX        =   794
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
            Caption         =   "Logr:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   345
            Index           =   1
            Left            =   5820
            TabIndex        =   28
            Top             =   195
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            Caption         =   ""
            Acao            =   6
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   345
            Index           =   0
            Left            =   5400
            TabIndex        =   27
            Top             =   195
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   90
            Left            =   6270
            TabIndex        =   233
            Top             =   270
            Width           =   225
            _ExtentX        =   397
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
            Caption         =   "IM:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin MSComctlLib.ListView lstPesq 
         Height          =   1215
         Left            =   90
         TabIndex        =   129
         Top             =   4740
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   2143
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   2540
         EndProperty
      End
      Begin Threed.SSFrame fra 
         Height          =   1395
         Index           =   5
         Left            =   -74880
         TabIndex        =   130
         Top             =   540
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
         Caption         =   "Características do Imóvel:"
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
            Left            =   7200
            MaxLength       =   3
            TabIndex        =   47
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
            Index           =   3
            Left            =   7200
            MaxLength       =   3
            TabIndex        =   46
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
            Index           =   2
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   45
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
            Index           =   1
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   44
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
            Index           =   0
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   43
            Top             =   270
            Width           =   375
         End
         Begin VB.ComboBox cboOcupLote 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":0094
            Left            =   2070
            List            =   "Tcim101.frx":0096
            Style           =   2  'Dropdown List
            TabIndex        =   137
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   270
            Width           =   3375
         End
         Begin VB.ComboBox cboPatrimonio 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":0098
            Left            =   2070
            List            =   "Tcim101.frx":009A
            Style           =   2  'Dropdown List
            TabIndex        =   136
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   630
            Width           =   3375
         End
         Begin VB.ComboBox cboCobranca 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":009C
            Left            =   2070
            List            =   "Tcim101.frx":009E
            Style           =   2  'Dropdown List
            TabIndex        =   135
            TabStop         =   0   'False
            Tag             =   "3"
            Top             =   990
            Width           =   3375
         End
         Begin VB.ComboBox cboLimites 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":00A0
            Left            =   7620
            List            =   "Tcim101.frx":00A2
            Style           =   2  'Dropdown List
            TabIndex        =   134
            TabStop         =   0   'False
            Tag             =   "4"
            Top             =   270
            Width           =   3375
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
            ItemData        =   "Tcim101.frx":00A4
            Left            =   7620
            List            =   "Tcim101.frx":00A6
            Style           =   2  'Dropdown List
            TabIndex        =   133
            TabStop         =   0   'False
            Tag             =   "5"
            Top             =   630
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
            ItemData        =   "Tcim101.frx":00A8
            Left            =   7410
            List            =   "Tcim101.frx":00AA
            Style           =   2  'Dropdown List
            TabIndex        =   132
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
            ItemData        =   "Tcim101.frx":00AC
            Left            =   6930
            List            =   "Tcim101.frx":00AE
            Style           =   2  'Dropdown List
            TabIndex        =   131
            Top             =   1950
            Width           =   3015
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   30
            Left            =   6090
            TabIndex        =   138
            Top             =   600
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
            Caption         =   "Arborização:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   276
            Index           =   32
            Left            =   6420
            TabIndex        =   139
            Top             =   300
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
            Caption         =   "Limites:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   276
            Index           =   33
            Left            =   336
            TabIndex        =   140
            Top             =   1020
            Width           =   1296
            _ExtentX        =   2302
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
            Caption         =   "Cod. Cobrança:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   276
            Index           =   34
            Left            =   60
            TabIndex        =   141
            Top             =   300
            Width           =   1584
            _ExtentX        =   2805
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
            Caption         =   "Ocupação do Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   35
            Left            =   630
            TabIndex        =   142
            Top             =   660
            Width           =   960
            _ExtentX        =   1693
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
            Caption         =   "Patrimônio:"
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
            TabIndex        =   143
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
            Index           =   37
            Left            =   5700
            TabIndex        =   144
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
         Height          =   975
         Index           =   3
         Left            =   -74880
         TabIndex        =   145
         Top             =   1950
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   1720
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
         Caption         =   "Características do Terreno:"
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
            Index           =   6
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   49
            Top             =   570
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
            Left            =   7200
            MaxLength       =   3
            TabIndex        =   50
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
            Index           =   5
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   48
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
            ItemData        =   "Tcim101.frx":00B0
            Left            =   2070
            List            =   "Tcim101.frx":00B2
            Style           =   2  'Dropdown List
            TabIndex        =   148
            TabStop         =   0   'False
            Tag             =   "6"
            Top             =   240
            Width           =   3375
         End
         Begin VB.ComboBox cboPedol 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":00B4
            Left            =   7590
            List            =   "Tcim101.frx":00B6
            Style           =   2  'Dropdown List
            TabIndex        =   147
            TabStop         =   0   'False
            Tag             =   "8"
            Top             =   240
            Width           =   3405
         End
         Begin VB.ComboBox cboSit 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":00B8
            Left            =   2070
            List            =   "Tcim101.frx":00BA
            Style           =   2  'Dropdown List
            TabIndex        =   146
            TabStop         =   0   'False
            Tag             =   "7"
            Top             =   600
            Width           =   3375
         End
         Begin Threed.SSPanel lbl 
            Height          =   312
            Index           =   20
            Left            =   696
            TabIndex        =   149
            Top             =   600
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Caption         =   "Topografia:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   264
            Index           =   21
            Left            =   876
            TabIndex        =   150
            Top             =   216
            Width           =   780
            _ExtentX        =   1376
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
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   240
            Index           =   22
            Left            =   6270
            TabIndex        =   151
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
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
            Caption         =   "Pedologia:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1515
         Index           =   4
         Left            =   -74880
         TabIndex        =   152
         Top             =   2940
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   2672
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
         Caption         =   "Dimensões do Terreno (m²)"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtAreaEdifTotalLote 
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
            Left            =   9030
            TabIndex        =   59
            Top             =   480
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox txtTestadaPrin 
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
            Left            =   360
            TabIndex        =   51
            Tag             =   "100"
            Top             =   480
            Width           =   1125
         End
         Begin VB.TextBox txtTestada2 
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
            Left            =   360
            TabIndex        =   52
            Tag             =   "101"
            Top             =   1020
            Width           =   1125
         End
         Begin VB.TextBox txtTrechoLogr2 
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
            Left            =   2550
            TabIndex        =   53
            Tag             =   "102"
            Top             =   510
            Width           =   1125
         End
         Begin VB.TextBox txtTestada3 
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
            Left            =   2550
            TabIndex        =   54
            Tag             =   "103"
            Top             =   1050
            Width           =   1125
         End
         Begin VB.TextBox txtTrechoLogr4 
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
            Left            =   7020
            TabIndex        =   57
            Tag             =   "106"
            Top             =   480
            Width           =   1125
         End
         Begin VB.TextBox txtTrechoLogr3 
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
            Left            =   4680
            TabIndex        =   55
            Tag             =   "104"
            Top             =   480
            Width           =   1125
         End
         Begin VB.TextBox txtTestada4 
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
            Left            =   4680
            TabIndex        =   56
            Tag             =   "105"
            Top             =   1020
            Width           =   1125
         End
         Begin VB.TextBox txtAreaLote 
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
            Left            =   9030
            TabIndex        =   60
            Tag             =   "108"
            Top             =   1080
            Width           =   1125
         End
         Begin VB.TextBox txtTestadaCampo 
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
            Left            =   7020
            TabIndex        =   58
            Tag             =   "107"
            Top             =   1080
            Width           =   1125
         End
         Begin Threed.SSPanel lbl 
            Height          =   276
            Index           =   23
            Left            =   4596
            TabIndex        =   153
            Top             =   240
            Width           =   1224
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
            Caption         =   "Trecho Logr. 3"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   276
            Index           =   24
            Left            =   6936
            TabIndex        =   154
            Top             =   240
            Width           =   1224
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
            Caption         =   "Trecho Logr. 4"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   330
            Index           =   25
            Left            =   2880
            TabIndex        =   155
            Top             =   840
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   582
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
            Caption         =   "Testada 3"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   264
            Index           =   26
            Left            =   2496
            TabIndex        =   156
            Top             =   276
            Width           =   1224
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
            Caption         =   "Trecho Logr. 2"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   276
            Index           =   27
            Left            =   120
            TabIndex        =   157
            Top             =   240
            Width           =   1452
            _ExtentX        =   2566
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
            Caption         =   "Testada Principal"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   28
            Left            =   660
            TabIndex        =   158
            Top             =   810
            Width           =   810
            _ExtentX        =   1429
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
            Caption         =   "Testada 2"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   264
            Index           =   19
            Left            =   4956
            TabIndex        =   159
            Top             =   816
            Width           =   816
            _ExtentX        =   1455
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
            Caption         =   "Testada 4"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   276
            Index           =   45
            Left            =   6780
            TabIndex        =   160
            Top             =   840
            Width           =   1368
            _ExtentX        =   2408
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
            Caption         =   "Testada(Campo)"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   51
            Left            =   9090
            TabIndex        =   161
            Top             =   840
            Width           =   1080
            _ExtentX        =   1905
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
            Caption         =   "Área do Lote"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   73
            Left            =   8940
            TabIndex        =   235
            Top             =   225
            Visible         =   0   'False
            Width           =   1500
            _ExtentX        =   2646
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
            Caption         =   "Área Edif. Total"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   5325
         Index           =   6
         Left            =   -74880
         TabIndex        =   162
         Top             =   540
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   9393
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
         Caption         =   "Características das Edificações"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtIc 
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
            Index           =   14
            Left            =   3870
            MaxLength       =   1
            TabIndex        =   62
            Tag             =   "Unidade"
            Top             =   210
            Width           =   285
         End
         Begin VB.TextBox txtIc 
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
            Index           =   13
            Left            =   2820
            MaxLength       =   4
            TabIndex        =   174
            Top             =   210
            Width           =   495
         End
         Begin VB.TextBox txtIc 
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
            Index           =   12
            Left            =   2310
            MaxLength       =   4
            TabIndex        =   173
            Top             =   210
            Width           =   495
         End
         Begin VB.TextBox txtIc 
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
            Index           =   11
            Left            =   1980
            MaxLength       =   2
            TabIndex        =   172
            Top             =   210
            Width           =   315
         End
         Begin VB.TextBox txtIc 
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
            Index           =   10
            Left            =   1650
            MaxLength       =   2
            TabIndex        =   171
            Top             =   210
            Width           =   315
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
            Left            =   1620
            MaxLength       =   3
            TabIndex        =   68
            Top             =   2655
            Width           =   375
         End
         Begin VB.TextBox txtAnoConst 
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
            Left            =   6990
            TabIndex        =   72
            Tag             =   "111"
            Top             =   1500
            Width           =   1185
         End
         Begin VB.TextBox txtAreaEdifTotal 
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
            Left            =   6990
            TabIndex        =   74
            Tag             =   "113"
            Top             =   2258
            Width           =   1155
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
            Left            =   6990
            MaxLength       =   3
            TabIndex        =   70
            Top             =   728
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
            Left            =   6990
            MaxLength       =   3
            TabIndex        =   71
            Top             =   1121
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
            Left            =   6990
            MaxLength       =   3
            TabIndex        =   69
            Top             =   390
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
            Left            =   1620
            MaxLength       =   3
            TabIndex        =   67
            Top             =   2265
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
            TabIndex        =   63
            Top             =   728
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
            TabIndex        =   64
            Top             =   1121
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
            Left            =   1620
            MaxLength       =   3
            TabIndex        =   66
            Top             =   1875
            Width           =   375
         End
         Begin VB.TextBox txtInscImobiliaria 
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
            Left            =   3330
            TabIndex        =   61
            Top             =   210
            Width           =   525
         End
         Begin VB.TextBox txtPavimento 
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
            Left            =   1650
            TabIndex        =   65
            Tag             =   "110"
            Top             =   1500
            Width           =   825
         End
         Begin VB.ComboBox cboPredio 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":00BC
            Left            =   2040
            List            =   "Tcim101.frx":00BE
            Style           =   2  'Dropdown List
            TabIndex        =   170
            TabStop         =   0   'False
            Tag             =   "15"
            Top             =   1113
            Width           =   3615
         End
         Begin VB.ComboBox cboUso 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":00C0
            Left            =   2010
            List            =   "Tcim101.frx":00C2
            Style           =   2  'Dropdown List
            TabIndex        =   169
            TabStop         =   0   'False
            Tag             =   "16"
            Top             =   1860
            Width           =   3615
         End
         Begin VB.ComboBox cboSentido 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":00C4
            Left            =   2040
            List            =   "Tcim101.frx":00C6
            Style           =   2  'Dropdown List
            TabIndex        =   168
            TabStop         =   0   'False
            Tag             =   "14"
            Top             =   720
            Width           =   3615
         End
         Begin VB.TextBox txtFracaoEdif 
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
            Left            =   6990
            TabIndex        =   75
            Tag             =   "114"
            Top             =   2648
            Width           =   1155
         End
         Begin VB.TextBox txtAreaEdif 
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
            Left            =   6990
            TabIndex        =   73
            Tag             =   "112"
            Top             =   1868
            Width           =   1185
         End
         Begin VB.ComboBox cboEstrutura 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":00C8
            Left            =   2010
            List            =   "Tcim101.frx":00CA
            Style           =   2  'Dropdown List
            TabIndex        =   167
            TabStop         =   0   'False
            Tag             =   "10"
            Top             =   2640
            Width           =   3615
         End
         Begin VB.ComboBox cboDestinacao 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":00CC
            Left            =   7380
            List            =   "Tcim101.frx":00CE
            Style           =   2  'Dropdown List
            TabIndex        =   166
            TabStop         =   0   'False
            Tag             =   "11"
            Top             =   375
            Width           =   3615
         End
         Begin VB.ComboBox cboTipologia 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":00D0
            Left            =   2010
            List            =   "Tcim101.frx":00D2
            Style           =   2  'Dropdown List
            TabIndex        =   165
            TabStop         =   0   'False
            Tag             =   "9"
            Top             =   2250
            Width           =   3615
         End
         Begin VB.ComboBox cboPadrao 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":00D4
            Left            =   7380
            List            =   "Tcim101.frx":00D6
            Style           =   2  'Dropdown List
            TabIndex        =   164
            TabStop         =   0   'False
            Tag             =   "12"
            Top             =   720
            Width           =   3615
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
            ItemData        =   "Tcim101.frx":00D8
            Left            =   7380
            List            =   "Tcim101.frx":00DA
            Style           =   2  'Dropdown List
            TabIndex        =   163
            TabStop         =   0   'False
            Tag             =   "13"
            Top             =   1113
            Width           =   3615
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   36
            Left            =   732
            TabIndex        =   175
            Top             =   2700
            Width           =   828
            _ExtentX        =   1455
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
            Caption         =   "Estrutura:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   38
            Left            =   756
            TabIndex        =   176
            Top             =   2316
            Width           =   804
            _ExtentX        =   1429
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
            Caption         =   "Tipologia:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   39
            Left            =   5940
            TabIndex        =   177
            Top             =   432
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   344
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
            Caption         =   "Destinação:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   40
            Left            =   6276
            TabIndex        =   178
            Top             =   768
            Width           =   624
            _ExtentX        =   1111
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
            Caption         =   "Padrão:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   41
            Left            =   5760
            TabIndex        =   179
            Top             =   1164
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   344
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
            Caption         =   "Conservação:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   42
            Left            =   5856
            TabIndex        =   180
            Top             =   1548
            Width           =   1044
            _ExtentX        =   1852
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
            Caption         =   "Ano Constr.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   44
            Left            =   5616
            TabIndex        =   181
            Top             =   2304
            Width           =   1284
            _ExtentX        =   2275
            _ExtentY        =   344
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
            Caption         =   "Área Edif. Total:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   47
            Left            =   876
            TabIndex        =   182
            Top             =   780
            Width           =   684
            _ExtentX        =   1217
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
            Caption         =   "Sentido:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   48
            Left            =   1188
            TabIndex        =   183
            Top             =   1920
            Width           =   372
            _ExtentX        =   661
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
            Caption         =   "Uso:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   52
            Left            =   960
            TabIndex        =   184
            Top             =   1176
            Width           =   600
            _ExtentX        =   1058
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
            Caption         =   "Predio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   53
            Left            =   528
            TabIndex        =   185
            Top             =   1560
            Width           =   1032
            _ExtentX        =   1826
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
            Caption         =   "Pavimentos:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin MSComctlLib.ListView lstEdific 
            Height          =   2175
            Left            =   90
            TabIndex        =   186
            Top             =   3030
            Width           =   10890
            _ExtentX        =   19209
            _ExtentY        =   3836
            View            =   3
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   15
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Insc. Imobiliária"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "14"
               Text            =   "Sentido"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "15"
               Text            =   "Prédio"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "110"
               Text            =   "Pavimentos"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "16"
               Text            =   "Uso"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "9"
               Text            =   "TipoLogia"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "10"
               Text            =   "Estrutura"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "11"
               Text            =   "Destinação"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "12"
               Text            =   "Padrão"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Tag             =   "13"
               Text            =   "Conservação"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Object.Tag             =   "111"
               Text            =   "Área Constr."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Object.Tag             =   "112"
               Text            =   "Área Edificada"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Object.Tag             =   "113"
               Text            =   "Área Edificada Total"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Object.Tag             =   "114"
               Text            =   "Fração Ideal"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "Sub Unidade"
               Object.Width           =   2540
            EndProperty
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   9
            Left            =   216
            TabIndex        =   187
            Top             =   252
            Width           =   1344
            _ExtentX        =   2381
            _ExtentY        =   344
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
            Caption         =   "Insc. Imobiliária:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   43
            Left            =   6060
            TabIndex        =   188
            Top             =   1908
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "Área Edif.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   204
            Index           =   46
            Left            =   5868
            TabIndex        =   189
            Top             =   2688
            Width           =   1032
            _ExtentX        =   1826
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
            Caption         =   "Fração Ideal:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdAdEdif 
            Height          =   375
            Left            =   8670
            TabIndex        =   76
            Top             =   2580
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   661
            Caption         =   "&Adicionar Edificação"
            Acao            =   1
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1455
         Index           =   2
         Left            =   -74880
         TabIndex        =   190
         Top             =   540
         Width           =   11085
         _ExtentX        =   19553
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
         Caption         =   "Referência Cadastral / Localização do Imóvel"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtIc 
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
            Index           =   15
            Left            =   3330
            MaxLength       =   1
            TabIndex        =   78
            Tag             =   "Unidade"
            Top             =   285
            Width           =   285
         End
         Begin VB.ComboBox cboTipoImovelBc 
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
            ItemData        =   "Tcim101.frx":00DC
            Left            =   9450
            List            =   "Tcim101.frx":00E6
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Tag             =   "Logradouro"
            Top             =   277
            Width           =   1545
         End
         Begin VB.TextBox txtCepImBc 
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
            MaxLength       =   10
            TabIndex        =   201
            Top             =   1050
            Width           =   1035
         End
         Begin VB.TextBox txtNumeroBc 
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
            Left            =   7470
            MaxLength       =   10
            TabIndex        =   81
            Top             =   667
            Width           =   525
         End
         Begin VB.TextBox txtComplementoBc 
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
            Left            =   8790
            TabIndex        =   82
            Top             =   667
            Width           =   2205
         End
         Begin VB.TextBox txtLoteamentoBc 
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
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   200
            Top             =   1050
            Width           =   705
         End
         Begin VB.TextBox txtQuadraBc 
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
            Left            =   2700
            MaxLength       =   5
            TabIndex        =   199
            Top             =   1050
            Width           =   705
         End
         Begin VB.TextBox txtLoteBc 
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
            Left            =   4110
            MaxLength       =   5
            TabIndex        =   198
            Top             =   1050
            Width           =   765
         End
         Begin VB.TextBox txtIc 
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
            Index           =   5
            Left            =   1260
            MaxLength       =   2
            TabIndex        =   197
            Tag             =   "Distrito"
            Top             =   285
            Width           =   315
         End
         Begin VB.TextBox txtIc 
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
            Index           =   6
            Left            =   1590
            MaxLength       =   2
            TabIndex        =   196
            Tag             =   "Setor"
            Top             =   285
            Width           =   315
         End
         Begin VB.TextBox txtIc 
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
            Index           =   7
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   195
            Tag             =   "Quadra"
            Top             =   285
            Width           =   495
         End
         Begin VB.TextBox txtIc 
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
            Index           =   8
            Left            =   2430
            MaxLength       =   4
            TabIndex        =   194
            Tag             =   "Lote"
            Top             =   285
            Width           =   495
         End
         Begin VB.TextBox txtIc 
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
            Index           =   9
            Left            =   2940
            MaxLength       =   3
            TabIndex        =   77
            Tag             =   "Unidade"
            Top             =   285
            Width           =   375
         End
         Begin VB.TextBox txtInscAnteriorBC 
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
            Left            =   7170
            TabIndex        =   79
            Top             =   285
            Width           =   1665
         End
         Begin VB.TextBox txtCodLogrBc 
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
            Left            =   1260
            TabIndex        =   102
            Top             =   667
            Width           =   1485
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
            Height          =   315
            Left            =   2790
            MaxLength       =   11
            TabIndex        =   193
            Top             =   667
            Width           =   1035
         End
         Begin VB.TextBox txtNomeLogr 
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
            Left            =   3870
            TabIndex        =   192
            Tag             =   "Nome Contribuinte"
            Top             =   660
            Width           =   3255
         End
         Begin VB.TextBox txtBairro 
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
            Left            =   5610
            TabIndex        =   191
            Tag             =   "Nome Contribuinte"
            Top             =   1050
            Width           =   3675
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   10
            Left            =   90
            TabIndex        =   202
            Top             =   712
            Width           =   1080
            _ExtentX        =   1905
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
            Caption         =   "Cod. Logr:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   54
            Left            =   8100
            TabIndex        =   203
            Top             =   712
            Width           =   660
            _ExtentX        =   1164
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
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   55
            Left            =   7200
            TabIndex        =   204
            Top             =   712
            Width           =   390
            _ExtentX        =   688
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
            Caption         =   "N.º:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   56
            Left            =   5010
            TabIndex        =   205
            Top             =   1095
            Width           =   705
            _ExtentX        =   1244
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
            Caption         =   "Bairro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   57
            Left            =   90
            TabIndex        =   206
            Top             =   1095
            Width           =   1170
            _ExtentX        =   2064
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
            Caption         =   "Loteamento:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   58
            Left            =   2040
            TabIndex        =   207
            Top             =   1095
            Width           =   660
            _ExtentX        =   1164
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
            Caption         =   "Quadra:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   59
            Left            =   3600
            TabIndex        =   208
            Top             =   1095
            Width           =   750
            _ExtentX        =   1323
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
            Caption         =   "Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   60
            Left            =   9540
            TabIndex        =   209
            Top             =   1095
            Width           =   360
            _ExtentX        =   635
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
            Caption         =   "CEP:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   61
            Left            =   8970
            TabIndex        =   210
            Top             =   330
            Width           =   420
            _ExtentX        =   741
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
            Caption         =   "Tipo:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   86
            Left            =   5940
            TabIndex        =   211
            Top             =   330
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Insc. Anterior:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   234
            Top             =   330
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Insc. Cad.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   675
         Index           =   8
         Left            =   -74880
         TabIndex        =   212
         Top             =   3795
         Width           =   8565
         _ExtentX        =   15108
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
         Caption         =   "Características do Imóvel:"
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
            Height          =   330
            Index           =   20
            Left            =   1530
            MaxLength       =   3
            TabIndex        =   99
            Top             =   225
            Width           =   420
         End
         Begin VB.ComboBox cboCobrancaBc 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":0100
            Left            =   2010
            List            =   "Tcim101.frx":0102
            Style           =   2  'Dropdown List
            TabIndex        =   100
            TabStop         =   0   'False
            Tag             =   "3"
            Top             =   225
            Width           =   6450
         End
         Begin VB.ComboBox Combo12 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":0104
            Left            =   7410
            List            =   "Tcim101.frx":0106
            Style           =   2  'Dropdown List
            TabIndex        =   214
            Top             =   2310
            Width           =   2535
         End
         Begin VB.ComboBox Combo13 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":0108
            Left            =   6930
            List            =   "Tcim101.frx":010A
            Style           =   2  'Dropdown List
            TabIndex        =   213
            Top             =   1950
            Width           =   3015
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   74
            Left            =   195
            TabIndex        =   215
            Top             =   300
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
            Caption         =   "Cod. Cobrança:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   77
            Left            =   5190
            TabIndex        =   216
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
            Index           =   78
            Left            =   5700
            TabIndex        =   217
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
         Height          =   645
         Index           =   9
         Left            =   120
         TabIndex        =   218
         Top             =   2310
         Visible         =   0   'False
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   1138
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
         Caption         =   "Aforamento"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtDtRegistro 
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
            Left            =   9990
            MaxLength       =   10
            TabIndex        =   25
            Top             =   255
            Width           =   990
         End
         Begin VB.TextBox txtRegistro 
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
            Left            =   8010
            MaxLength       =   50
            TabIndex        =   24
            Top             =   247
            Width           =   585
         End
         Begin VB.TextBox txtFolhaAforamento 
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
            Left            =   4695
            MaxLength       =   50
            TabIndex        =   22
            Top             =   247
            Width           =   585
         End
         Begin VB.TextBox txtDataAforamento 
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
            Left            =   5835
            MaxLength       =   10
            TabIndex        =   23
            Top             =   255
            Width           =   1215
         End
         Begin VB.TextBox txtNumAforamento 
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
            Left            =   390
            MaxLength       =   5
            TabIndex        =   19
            Top             =   255
            Width           =   840
         End
         Begin VB.TextBox txtFichaAforamento 
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
            Left            =   1935
            MaxLength       =   5
            TabIndex        =   20
            Top             =   255
            Width           =   705
         End
         Begin VB.TextBox txtLivroAforamento 
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
            Left            =   3375
            MaxLength       =   5
            TabIndex        =   21
            Top             =   255
            Width           =   615
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   79
            Left            =   4095
            TabIndex        =   219
            Top             =   300
            Width           =   495
            _ExtentX        =   873
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
            Caption         =   "Folha:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   80
            Left            =   90
            TabIndex        =   220
            Top             =   300
            Width           =   225
            _ExtentX        =   397
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
            Caption         =   "Nº:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   81
            Left            =   1455
            TabIndex        =   221
            Top             =   300
            Width           =   480
            _ExtentX        =   847
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
            Caption         =   "Ficha:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   87
            Left            =   2865
            TabIndex        =   222
            Top             =   300
            Width           =   480
            _ExtentX        =   847
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
            Caption         =   "Livro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   88
            Left            =   5355
            TabIndex        =   223
            Top             =   300
            Width           =   405
            _ExtentX        =   714
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
            Caption         =   "Data:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   12
            Left            =   7230
            TabIndex        =   231
            Top             =   300
            Width           =   735
            _ExtentX        =   1296
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
            Caption         =   "Registro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   89
            Left            =   8820
            TabIndex        =   232
            Top             =   300
            Width           =   1155
            _ExtentX        =   2037
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
            Caption         =   "Data Registro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin MSComctlLib.ListView lstCond 
         Height          =   1380
         Left            =   -74910
         TabIndex        =   224
         Top             =   4530
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   2434
         View            =   3
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   23
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "14"
            Text            =   "IC Anterior"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "15"
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "110"
            Text            =   "Nº"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "16"
            Text            =   "Complemento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "9"
            Text            =   "IM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "10"
            Text            =   "Contribuinte"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "11"
            Text            =   "CPF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "12"
            Text            =   "Tipo Logr"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "13"
            Text            =   "Logradouro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "111"
            Text            =   "Nº"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "112"
            Text            =   "Complemento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Tag             =   "113"
            Text            =   "Bairro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Tag             =   "114"
            Text            =   "CEP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Municipio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "UF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Ocupante"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Cpf Ocupante"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Cod. Cobrança"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Sub Unidade"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Cod Logr"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Cod Bairro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Cod Logr Und"
            Object.Width           =   2540
         EndProperty
      End
      Begin VTOcx.cmdVISUAL cmdAdCond 
         Height          =   375
         Left            =   -66120
         TabIndex        =   101
         Top             =   3990
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "Adicionar &Condomínio"
         Acao            =   1
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin Threed.SSFrame fra 
         Height          =   1785
         Index           =   7
         Left            =   -74880
         TabIndex        =   236
         Top             =   1950
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   3149
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
         Caption         =   "Dados do Proprietário"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtMunicBc 
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
            Left            =   6720
            TabIndex        =   95
            Tag             =   "Município"
            Top             =   960
            Width           =   3465
         End
         Begin VB.TextBox txtNumeroContribBc 
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
            Left            =   7500
            MaxLength       =   10
            TabIndex        =   90
            Top             =   585
            Width           =   525
         End
         Begin VB.ComboBox cboUFBc 
            BeginProperty Font 
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
            ItemData        =   "Tcim101.frx":010C
            Left            =   10215
            List            =   "Tcim101.frx":0119
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Tag             =   "UF"
            Top             =   945
            Width           =   795
         End
         Begin VB.TextBox txtCpfCgcBc 
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
            Left            =   8955
            MaxLength       =   20
            TabIndex        =   86
            Top             =   210
            Width           =   2040
         End
         Begin VB.TextBox txtCpfOcupanteBc 
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
            Left            =   8955
            MaxLength       =   20
            TabIndex        =   98
            Top             =   1335
            Width           =   2040
         End
         Begin VB.TextBox txtOcupanteBc 
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
            Left            =   3060
            TabIndex        =   97
            Top             =   1335
            Width           =   4965
         End
         Begin VB.TextBox txtBairroContribBc 
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
            Left            =   1965
            TabIndex        =   93
            Tag             =   "Bairro"
            Top             =   960
            Width           =   2250
         End
         Begin VB.TextBox txtCepBc 
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
            Left            =   4650
            MaxLength       =   10
            TabIndex        =   94
            Top             =   960
            Width           =   1065
         End
         Begin VB.TextBox txtNomeLogrContribBc 
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
            Left            =   3060
            TabIndex        =   89
            Tag             =   "Nome Logradouro"
            Top             =   585
            Width           =   4065
         End
         Begin VB.TextBox txtNomeContribBc 
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
            Left            =   3060
            TabIndex        =   85
            Tag             =   "Nome Contribuinte"
            Top             =   210
            Width           =   4965
         End
         Begin VB.TextBox txtIMBc 
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
            Left            =   1440
            MaxLength       =   11
            TabIndex        =   83
            Top             =   210
            Width           =   1215
         End
         Begin VB.TextBox txtCompContribBc 
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
            Left            =   9270
            TabIndex        =   91
            Top             =   570
            Width           =   1725
         End
         Begin VB.TextBox txtCodLogrContribBc 
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
            Left            =   1425
            TabIndex        =   87
            Tag             =   "Logradouro"
            Top             =   585
            Width           =   525
         End
         Begin VB.TextBox txtTipoLogrContribBc 
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
            Left            =   1980
            MaxLength       =   11
            TabIndex        =   88
            Top             =   585
            Width           =   1035
         End
         Begin VB.TextBox txtCodBairroContribBc 
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
            Left            =   1410
            TabIndex        =   92
            Tag             =   "Logradouro"
            Top             =   960
            Width           =   525
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   62
            Left            =   150
            TabIndex        =   237
            Top             =   255
            Width           =   1275
            _ExtentX        =   2249
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
            Caption         =   "Insc. Municipal:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   64
            Left            =   4260
            TabIndex        =   238
            Top             =   1005
            Width           =   375
            _ExtentX        =   661
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
            Caption         =   "CEP:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   65
            Left            =   5835
            TabIndex        =   239
            Top             =   1005
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Municipio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   66
            Left            =   7200
            TabIndex        =   240
            Top             =   630
            Width           =   270
            _ExtentX        =   476
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
            Caption         =   "N.º:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   67
            Left            =   810
            TabIndex        =   241
            Top             =   1020
            Width           =   555
            _ExtentX        =   979
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
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   68
            Left            =   8490
            TabIndex        =   242
            Top             =   630
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
            Index           =   69
            Left            =   8085
            TabIndex        =   243
            Top             =   277
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   70
            Left            =   2175
            TabIndex        =   244
            Top             =   1380
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "Ocupante:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   71
            Left            =   8085
            TabIndex        =   245
            Top             =   1380
            Width           =   795
            _ExtentX        =   1402
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
            Caption         =   "CPF/CNPJ"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   63
            Left            =   420
            TabIndex        =   246
            Top             =   615
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
            Caption         =   "Logradouro"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   345
            Index           =   3
            Left            =   2670
            TabIndex        =   84
            Top             =   210
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
   End
   Begin VB.TextBox txtFatorFixo 
      Height          =   285
      Left            =   8640
      TabIndex        =   118
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   4560
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   226
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1138
      Descricao       =   "Cadastro"
      Icone           =   "Tcim101.frx":013A
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   2
      Left            =   10200
      TabIndex        =   227
      Top             =   6750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   1
      Left            =   9060
      TabIndex        =   228
      Top             =   6750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   4
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   229
      Top             =   6750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCIM101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Cadastro As VSImposto
Dim NovoContrib As Boolean
Dim sql As String
Dim Lote As New BCI
Private Boletim As TipoBoletim

Function TotalProva(Valor As String) As String
    Static Total As Double
    If Trim(Valor) = "" Then Valor = "0"
    Total = CDbl(Valor) + Total
    TotalProva = Total
End Function

Public Sub HabilitaCaixa(Status As Boolean)
    txtIM.Enabled = Not Status
    txtNomeContrib.Enabled = Status
    txtNomeTipoLogrContrib.Enabled = Status
    txtNomeLogrContrib.Enabled = Status
    txtNumeroContrib.Enabled = Status
    txtCompContrib.Enabled = Status
    txtBairroContrib.Enabled = Status
    txtCep.Enabled = Status
    txtMunic.Enabled = Status
    txtUf.Enabled = Status
    txtIM = ""
    txtNomeContrib = ""
    txtNomeTipoLogrContrib = ""
    txtNomeLogrContrib = ""
    txtNumeroContrib = ""
    txtCompContrib = ""
    txtBairroContrib = ""
    txtCep = ""
    txtMunic = ""
    txtUf = ""
End Sub

Private Sub cboTipoImovel_Click()
    If cboTipoImovel = "PREDIAL" Then
        Boletim = tbo_Predial
    Else
        Boletim = tbo_Territorial
    End If

End Sub

Private Sub cmd_Click(Index As Integer)
    On Error Resume Next
    Dim Valores As String
    Dim Campos As String
    Dim DataReab As Date
    Dim RsAux As VSRecordset
    Dim Rs As VSRecordset
    Dim InscricaoMunicipal  As String
    Dim InscricaoCadastral As String
    Dim CodLogr As Long
    Dim DtVenc As String
    Dim SitCadastral As String
    Static Unidades As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Cadastro As New VSImposto
    Select Case cmd(Index).Caption
        Case "&Salvar"
                If Trim(txtCodBairro) = "" Then
                    Util.Informa "Falta a definição do bairro."
                    txtCodBairro.SetFocus
                    tabCad.Tab = 0
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                txtFatorFixo.Tag = "1000"
                CodLogr = txtCodLogr
                InscricaoCadastral = txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & txtIc(4)
                InscricaoMunicipal = txtIM
                Screen.MousePointer = 11
                If Not Lote.VerificaFechamentoAreas(lstEdific) Then Exit Sub
                'GRAVANDO BT
                Lote.CarregaDadosContribuinte InscricaoMunicipal, txtNomeContrib, txtCpfCgc, txtCodLogrContrib, txtNomeTipoLogrContrib, txtNomeLogrContrib, _
                        txtNumeroContrib, txtCompContrib, txtCodBairroContrib, txtBairroContrib, txtCep, txtMunic, txtUf
                If Not Lote.InsereContribuinte(NovoContrib) Then Exit Sub
                txtIcAnterior = Edita.TiraPic(txtIcAnterior, ".")
                Lote.CarregaDadosImovel InscricaoCadastral, txtIcAnterior, txtIc(4), "0", "", "", CStr(CodLogr), txtCodBairro, _
                     txtNumero, txtComplemento, txtLote, txtQuadra, txtLoteamento, Boletim, txtOcupante, txtCpfOcupante, _
                     txtCodMens, txtZona, txtNumAforamento, txtFichaAforamento, txtLivroAforamento, txtFolhaAforamento, txtRegistro, txtDataAforamento, txtDtRegistro
                
                If Not Lote.InsereTerritorio() Then Exit Sub
                Call Lote.GravaComponentes(InscricaoCadastral, Me, 1, 8, False, txtIc(4), 0)
                Call Lote.GravaComponentes(InscricaoCadastral, Me, 100, 109, True, txtIc(4), 0)

                'VERIFCANDO BP'S
                If cboTipoImovel = "PREDIAL" Then
                    If Not Lote.VerificaDigitacaoBP(lstEdific, txtCodBairro, tabCad) Then
                        txtCodComponente(13).SetFocus
                        Exit Sub
                    End If
                End If
                'GRAVANDO BP
                Lote.GravaBP lstEdific, txtCodMens, txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3), txtIc(4)
                'GRAVANDO BC'S
                If CInt(Nvl(Trim(txtIc(4)), 0)) >= 200 Then
                    cboCobrancaBc.Tag = "3"
                    cboCobranca.Tag = ""
                    If lstCond.ListItems.Count > 0 Then
                        For j = 1 To lstCond.ListItems.Count 'Para cada edificacao
                            lstCond.ListItems(j).Selected = True
                            InscricaoMunicipal = txtIMBc
                            CodLogr = txtCodLogr
                            InscricaoCadastral = txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & lstCond.SelectedItem
                            'INSERE CONTRIBUINTE
                            If lstCond.SelectedItem.ListSubItems(5) = "" Then
                                InscricaoMunicipal = Cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
                            Else
                                InscricaoMunicipal = lstCond.SelectedItem.ListSubItems(5)
                            End If
                            
                            Lote.CarregaDadosContribuinte InscricaoMunicipal, lstCond.SelectedItem.ListSubItems(6), _
                                    "", lstCond.SelectedItem.ListSubItems(20), lstCond.SelectedItem.ListSubItems(8), _
                                    lstCond.SelectedItem.ListSubItems(9), lstCond.SelectedItem.ListSubItems(10), _
                                     lstCond.SelectedItem.ListSubItems(11), lstCond.SelectedItem.ListSubItems(21), _
                                    lstCond.SelectedItem.ListSubItems(12), lstCond.SelectedItem.ListSubItems(13), _
                                    lstCond.SelectedItem.ListSubItems(14), lstCond.SelectedItem.ListSubItems(15)
                            Lote.InsereContribuinte
                            
                            'INSERE IMOVEL
                            Lote.CarregaDadosImovel InscricaoCadastral, "", lstCond.SelectedItem, lstCond.SelectedItem.ListSubItems(19), _
                                    InscricaoCadastral & txtIc(4), "", lstCond.SelectedItem.SubItems(22), txtCodBairro, _
                                    lstCond.SelectedItem.ListSubItems(3), lstCond.SelectedItem.ListSubItems(4), _
                                    Trim(txtLoteBc), Trim(txtQuadraBc), Trim(txtLoteamentoBc), lstCond.SelectedItem.ListSubItems(2), _
                                    lstCond.SelectedItem.ListSubItems(16), lstCond.SelectedItem.ListSubItems(17), _
                                    Nvl(lstCond.SelectedItem.ListSubItems(18), 0), Nvl(txtZona, 1)
                            Lote.InsereTerritorio
                            
                            'INSERE COD. COBRANÇA
                            Call Lote.GravaComponente(InscricaoCadastral, lstCond.SelectedItem, lstCond.SelectedItem.ListSubItems(18), 3, CInt(Nvl(lstCond.SelectedItem.ListSubItems(19), 0)))
                        Next
                    End If
                End If
                'LIMPA TELA
                Informa "Registro gravado com sucesso."
                Dim Capa As New cCapa
                If Capa.FechaLote(txtIc(0), txtIc(1), txtIc(2)) Then Avisa "Lote fechado."
                Set Capa = Nothing
                'If Lote.FechaLote(txtIc(0), txtIc(1), txtIc(2)) Then Avisa "Lote fechado."
                 Call cmd_Click(0)
                DoEvents
        Case "&Novo"
            Call Edita.LimpaCampos(Me)
            NovoContrib = True
            cboCobrancaBc.Tag = ""
            cboCobranca.Tag = "3"
            lstEdific.ListItems.Clear
            lstCond.ListItems.Clear
            tabCad.Tab = 0
            Unidades = 0
            Screen.MousePointer = 0
            txtIc(0).SetFocus
        Case "Sai&r"
            NovoContrib = True
            Unload Me
    End Select
End Sub

Private Sub cmdAdCond_Click()
     'NOVIDADE
    Dim ItmX As Object
    Dim i As Byte
    
    On Error Resume Next
    If Trim(txtIc(9)) = "" Then
        Avisa "Informe a unidade."
        txtIc(9).SetFocus
        Exit Sub
    End If
    Set ItmX = lstCond.ListItems.Add(, , txtIc(9))
    ItmX.SubItems(1) = txtInscAnteriorBC
    ItmX.SubItems(2) = cboTipoImovelBc.ListIndex + 1
    ItmX.SubItems(3) = txtNumeroBc
    ItmX.SubItems(4) = txtComplementoBc
    ItmX.SubItems(5) = IIf(Trim(txtIMBc) = "", "", txtIMBc)
    ItmX.SubItems(6) = txtNomeContribBc
    ItmX.SubItems(7) = txtCpfCgcBc
    ItmX.SubItems(8) = txtTipoLogrContribBc
    ItmX.SubItems(9) = txtNomeLogrContribBc
    ItmX.SubItems(10) = txtNumeroContribBc
    ItmX.SubItems(11) = txtCompContribBc
    ItmX.SubItems(12) = txtBairroContribBc
    ItmX.SubItems(13) = txtCepImBc
    ItmX.SubItems(14) = txtMunicBc
    ItmX.SubItems(15) = cboUFBc
    ItmX.SubItems(16) = txtOcupanteBc
    ItmX.SubItems(17) = txtCpfOcupanteBc
    ItmX.SubItems(18) = txtCodComponente(20)
    ItmX.SubItems(19) = txtIc(15)
    ItmX.SubItems(20) = txtCodLogrContribBc
    ItmX.SubItems(21) = txtCodBairroContribBc
    ItmX.SubItems(22) = txtCodLogrBc '(SQz:As unidades condominiais podem ter enderecos diferentes)
    txtNumeroBc = ""
    txtComplementoBc = ""
    txtIMBc = ""
    txtNomeContribBc.Text = ""
    txtCpfCgcBc = ""
    txtTipoLogrContribBc = ""
    txtNomeLogrContribBc = ""
    txtNumeroContribBc = ""
    txtCompContribBc = ""
    txtNumeroBc = ""
    txtComplementoBc = ""
    txtBairroContribBc = ""
    txtCepImBc = ""
    txtMunicBc = ""
    cboUFBc.ListIndex = -1
    txtOcupanteBc = ""
    txtCpfOcupanteBc = ""
    txtCodComponente(20) = ""
    txtIc(15) = ""
    txtIc(9) = Format(CInt(txtIc(9)) + 1, "000")

    txtIc(9).SetFocus
End Sub

Private Sub cmdAdEdif_Click()
    On Error GoTo Trata
    Dim ItmX As Object
    Dim i As Byte
    If Trim(txtInscImobiliaria) = "" Then
        Informa "Informe a unidade."
        txtInscImobiliaria.SetFocus
        Exit Sub
    End If
    
    atualizarAreaEdifTotal txtAreaEdifTotal
    Set ItmX = lstEdific.ListItems.Add(, , txtInscImobiliaria)
    ItmX.SubItems(1) = txtCodComponente(13)
    ItmX.SubItems(2) = txtCodComponente(14)
    ItmX.SubItems(3) = txtPavimento
    ItmX.SubItems(4) = txtCodComponente(15)
    For i = 8 To 13
        ItmX.SubItems(i - 3) = txtCodComponente(i)
    Next
    ItmX.SubItems(10) = txtAnoConst
    ItmX.SubItems(11) = IIf(Trim(txtAreaEdif) = "", 0, txtAreaEdif)
    ItmX.SubItems(12) = IIf(Trim(txtAreaEdifTotal) = "", 0, txtAreaEdifTotal)
    ItmX.SubItems(13) = IIf(Trim(txtFracaoEdif) = "", 0, txtFracaoEdif)
    ItmX.SubItems(14) = IIf(Trim(txtIc(14)) = "", 0, Trim(txtIc(14)))
    
    For i = 8 To 15
        txtCodComponente(i) = ""
    Next
    txtAnoConst = ""
    txtAreaEdif = ""
    'txtAreaEdifTotal = txtAreaEdifTotalLote
    txtFracaoEdif = ""
    txtPavimento = ""
    txtInscImobiliaria = Format(CInt(txtInscImobiliaria) + 1, "000")
    txtCodComponente(13).SetFocus
Trata:
End Sub

Private Sub atualizarAreaEdifTotal(Area As String)
    Dim Edificacao As ListItem
    
    For Each Edificacao In lstEdific.ListItems
        Edificacao.SubItems(12) = Area
    Next
End Sub
Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub cmdOpcao_Click(Index As Integer)
    Dim Rs As VSRecordset
    Dim sql As String
    Select Case Index
        Case 0
            sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where tci_nome like '" & txtNomeContrib & "%' or tci_nome like '% " & txtNomeContrib & "%'"
            sql = sql & " and tci_tsc_cod_sit_cad =1 ORDER BY tci_nome"
            If Not Bdados.AbreTabela(sql, Rs) Then
                Call Util.Avisa("Nenhum contribuinte encontrado.")
                SendKeys "{tab}"
            End If
            Bdados.FechaTabela Rs
            MontaGrid Bdados, lstPesq, sql, 1400
            If lstPesq.ListItems.Count > 0 Then lstPesq.SortKey = 1
        Case 1
            NovoContrib = True
            txtIM = ""
            Call HabilitaCaixa(True)
            txtCpfCgc = ""
            txtCpfOcupante = ""
            txtOcupante = ""
            txtNomeContrib.SetFocus
        Case 3
           AplicacoesVTFuncoes.BuscaNoEconomico TcoFisica, txtIMBc
           txtIMBc_LostFocus
    End Select
End Sub

Private Sub Form_Load()
    
    Dim Controle As Control
    Dim i As Byte
    Dim Rs As VSRecordset
    Set Cadastro = New VSImposto
    
    Call AtualizaUF(cboUFBc)
    
    For Each Controle In Controls
        If IsNumeric(Controle.Tag) Then
            If Val(Controle.Tag) < 20 Then Call Edita.AtualizaCombo(Bdados, Controle, "Select tco_descricao_componente From Tab_Componente_Avancado Where tco_grupo = " & Controle.Tag & " order by tco_cod_componente asc")
        End If
    Next
    HabilitaCaixa False
    txtNomeContrib.Enabled = True
    fra(9).Visible = Nvl(Temp.PegaParametro(Bdados, "AFORAMENTO TCIU101"), 1)
        
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    NovoContrib = True
    Bdados.FechaTabela Rs
    Boletim = tbo_Territorial
End Sub

Private Sub lstCond_DblClick()
    If lstCond.SelectedItem Is Nothing Then Exit Sub
    Dim ItmX As Object
    Dim i As Byte
    
    On Error Resume Next
    txtIc(9) = lstCond.SelectedItem
    txtInscAnteriorBC = lstCond.SelectedItem.SubItems(1)
    cboTipoImovelBc.ListIndex = lstCond.SelectedItem.SubItems(2) - 1
    txtNumeroBc = lstCond.SelectedItem.SubItems(3)
    txtComplementoBc = lstCond.SelectedItem.SubItems(4)
    txtIMBc = lstCond.SelectedItem.SubItems(5)
    txtNomeContribBc = lstCond.SelectedItem.SubItems(6)
    txtCpfCgcBc = lstCond.SelectedItem.SubItems(7)
    txtTipoLogrContribBc = lstCond.SelectedItem.SubItems(8)
    txtNomeLogrContribBc = lstCond.SelectedItem.SubItems(9)
    txtNumeroContribBc = lstCond.SelectedItem.SubItems(10)
    txtComplementoBc = lstCond.SelectedItem.SubItems(11)
    txtBairroContribBc = lstCond.SelectedItem.SubItems(12)
    txtCepImBc = lstCond.SelectedItem.SubItems(13)
    txtMunicBc = lstCond.SelectedItem.SubItems(14)
    cboUFBc = lstCond.SelectedItem.SubItems(15)
    txtOcupanteBc = lstCond.SelectedItem.SubItems(16)
    txtCpfOcupanteBc = lstCond.SelectedItem.SubItems(17)
    txtCodComponente(20) = lstCond.SelectedItem.SubItems(18)
    txtCodLogrContribBc = lstCond.SelectedItem.SubItems(20)
    txtCodBairroContribBc = lstCond.SelectedItem.SubItems(21)
    lstCond.ListItems.Remove lstCond.SelectedItem.Index
End Sub

Private Sub lstEdific_DblClick()
    Dim i As Integer
    If lstEdific.SelectedItem Is Nothing Then Exit Sub
    If Trim(txtInscImobiliaria) <> "" Then
        If Confirma("Existe uma unidade edificada em aberto. Deseja exclui-la?") Then
            lstEdific.ListItems.Remove lstEdific.SelectedItem.Index
        Else
            txtInscImobiliaria = Right(lstEdific.SelectedItem, 3)
            txtCodComponente(13) = lstEdific.SelectedItem.SubItems(1)
            txtCodComponente(14) = lstEdific.SelectedItem.SubItems(2)
            txtPavimento = lstEdific.SelectedItem.SubItems(3)
            txtCodComponente(15) = lstEdific.SelectedItem.SubItems(4)
            For i = 8 To 12
                txtCodComponente(i) = lstEdific.SelectedItem.SubItems(i - 3)
            Next
            txtAnoConst = lstEdific.SelectedItem.SubItems(10)
            txtAreaEdif = lstEdific.SelectedItem.SubItems(11)
            txtAreaEdifTotal = lstEdific.SelectedItem.SubItems(12)
            txtAreaEdifTotalLote = txtAreaEdifTotal
            txtFracaoEdif = lstEdific.SelectedItem.SubItems(13)
            txtIc(14) = lstEdific.SelectedItem.SubItems(14)
            lstEdific.ListItems.Remove lstEdific.SelectedItem.Index
        End If
    End If
End Sub

Private Sub lstPesq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaGrid lstPesq, ColumnHeader
End Sub

Private Sub lstPesq_DblClick()
    txtIM = lstPesq.SelectedItem
    Call txtIm_LostFocus
End Sub

Private Sub tabCad_Click(PreviousTab As Integer)
    'NOVIDADE
    If tabCad.Tab = 2 Then
        If Trim(txtInscImobiliaria) = "" And Trim(txtIc(4)) <> "" Then
            txtInscImobiliaria.Enabled = True
            txtIc(10) = txtIc(0)
            txtIc(11) = txtIc(1)
            txtIc(12) = txtIc(2)
            txtIc(13) = txtIc(3)
            txtInscImobiliaria.SetFocus
        End If
    ElseIf tabCad.Tab = 3 Then
        If Trim(txtIc(4)) <> "" Then
            txtIc(5) = txtIc(0)
            txtIc(6) = txtIc(1)
            txtIc(7) = txtIc(2)
            txtIc(8) = txtIc(3)
            txtIc(9).Enabled = True
            
            cboTipoImovelBc.ListIndex = cboTipoImovel.ListIndex
            txtCodLogrBc = txtCodLogr
            txtLogr = txtTipoLogrBt
            txtNomeLogr = txtLogrBt
            txtNumeroBc = txtNumero
            txtLoteamentoBc = txtLoteamento
            txtQuadraBc = txtQuadra
            txtLoteBc = txtLote
            txtBairro = txtBairroBt
            txtCepImBc = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
        End If
    End If
End Sub



Private Sub txtAconstr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtAnoAq_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtAreaNao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtAreaEdif_Change()
On Error Resume Next
If Trim(txtAreaEdif) <> "" Then
    If Not IsNumeric(txtAreaEdif) Then Exit Sub
   txtFracaoEdif = Format(CDbl(Nvl(txtAreaEdif, 1)) / CDbl(Nvl(txtAreaEdifTotal, 1)), "#0.000,0000")
Else
    txtFracaoEdif = ""
End If
End Sub

Private Sub txtAreaEdifTotal_Change()
    Call txtAreaEdif_Change
End Sub

Private Sub txtAreaEdifTotalLote_Change()
    If CDbl(Nvl(Trim(txtAreaEdifTotalLote), 0)) > 0 Then
        txtAreaEdifTotal = txtAreaEdifTotalLote
        txtFracaoEdif = ""
    Else
        txtAreaEdif = ""
    End If
End Sub

Private Sub txtAreaEdifTotalLote_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtAreaLote_LostFocus()
    tabCad.Tab = 2
    DoEvents
    txtInscImobiliaria.Enabled = True
    txtInscImobiliaria.SetFocus
End Sub

Private Sub txtBairroContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBairroContribBc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCep_LostFocus()
    If Trim$(txtCep) = "" Then
        txtCep = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
    Else
        If IsNumeric(txtCep) Then txtCepImBc = Edita.FormataTexto(txtCep, CEP)
    End If
End Sub

Private Sub txtCepImBc_LostFocus()
    If IsNumeric(txtCepImBc) Then txtCepImBc = Edita.FormataTexto(txtCepImBc, CEP)
End Sub

Private Sub txtCodBairro_LostFocus()
    Lote.BuscaBairro txtCodBairro, txtBairroBt
    
End Sub

Private Sub txtCodBairroContrib_LostFocus()
    Lote.BuscaBairro txtCodBairroContrib, txtBairroContrib
End Sub

Private Sub txtCodBairroContribBc_LostFocus()
    Lote.BuscaBairro txtCodBairroContribBc, txtBairroContribBc
End Sub

Private Sub txtCodComponente_Change(Index As Integer)
    Dim Controle As Control
    On Error GoTo Trata
     If Index = 20 Then
        cboCobrancaBc.ListIndex = Nvl(txtCodComponente(Index).Text, 0) - 1
        Exit Sub
    End If
    For Each Controle In Controls
        If Controle.Tag = Index + 1 Then
            Controle.ListIndex = Util.Nvl(txtCodComponente(Index).Text, 0) - 1
            Exit For
        End If
    Next
Trata:
    If Err.Number = 380 Then
        txtCodComponente(Index).SetFocus
    End If
End Sub

Private Sub txtCodComponente_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodLogr_LostFocus()
    Lote.BuscaLogradouro txtCodLogr, txtTipoLogrBt, txtLogrBt
End Sub

Private Sub txtCodLogrBc_LostFocus()
    Lote.BuscaLogradouro txtCodLogrBc, txtLogr, txtNomeLogr, , txtCepImBc
End Sub

Private Sub txtCodLogrContrib_LostFocus()
    Lote.BuscaLogradouro txtCodLogrContrib, txtNomeTipoLogrContrib, txtNomeLogrContrib, txtMunic, txtCep, txtUf
End Sub

Private Sub txtCodLogrContribBc_LostFocus()
    Lote.BuscaLogradouro txtCodLogrContribBc, txtNomeLogrContribBc, txtNomeLogrContribBc, txtMunicBc, txtCepBc, cboUFBc
End Sub

Private Sub txtCodMens_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtCodMens_LostFocus()
    txtDescMens = exibirMensagem(txtCodMens)
End Sub

Private Function exibirMensagem(ByRef Codigo As Object) As String
    If Trim$(Codigo) <> "" Then
        If Bdados.AbreTabela("SELECT TCM_MENSAGEM FROM TAB_COD_MENSAGEM WHERE TCM_CODIGO=" & Trim$(Codigo)) Then
            exibirMensagem = Bdados.Tabela(0).Value
        Else
            Erro "Código de mensagem inválido."
            Codigo = ""
        End If
        Bdados.FechaTabela
    End If
End Function
Private Sub txtCompContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFracao_Change()

End Sub

Private Sub txtCpfCgc_LostFocus()
    If Len(txtCpfCgc) = 11 Then
        If Not Util.ValidaCpf(txtCpfCgc) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCpfCgc = ""
            txtCpfCgc.SetFocus
            Exit Sub
        End If
        txtCpfCgc = Edita.FormataTexto(txtCpfCgc, Cpf)
        'txtCpfOcupante = txtCpfCgc
    ElseIf Len(txtCpfCgc) = 14 And Mid(txtCpfCgc, 4, 1) <> "." Then
        txtCpfCgc.MaxLength = 20
        txtCpfCgc = Edita.FormataTexto(txtCpfCgc, Cgc)
        'txtCpfOcupante = txtCpfCgc
    ElseIf Trim(txtCpfCgc) <> "" And Len(txtCpfCgc) <> 18 And Mid(txtCpfCgc, 4, 1) <> "." Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
        txtCpfCgc = ""
        txtCpfCgc.SetFocus
    End If
End Sub

Private Sub txtCpfCgcBc_LostFocus()
    If Len(txtCpfCgcBc) = 11 Then
        If Not Util.ValidaCpf(txtCpfCgcBc) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCpfCgcBc = ""
             txtCpfCgcBc.SetFocus
            Exit Sub
        End If
        txtCpfCgcBc = Edita.FormataTexto(txtCpfCgcBc, Cpf)
    ElseIf Len(txtCpfCgcBc) = 14 And Mid(txtCpfCgcBc, 4, 1) <> "." Then
        txtCpfCgcBc.MaxLength = 20
        txtCpfCgcBc = Edita.FormataTexto(txtCpfCgcBc, Cgc)
    ElseIf Trim(txtCpfCgcBc) <> "" And Len(txtCpfCgcBc) <> 18 And Mid(txtCpfCgcBc, 4, 1) <> "." Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
        txtCpfCgcBc = ""
        txtCpfCgcBc.SetFocus
    End If
End Sub

Private Sub txtCpfOcupante_LostFocus()
    If Len(txtCpfOcupante) = 11 Then
        If Not Util.ValidaCpf(txtCpfOcupante) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCpfOcupante = ""
            txtCpfOcupante.SetFocus
            Exit Sub
        End If
        txtCpfOcupante = Edita.FormataTexto(txtCpfOcupante, Cpf)
    ElseIf Len(txtCpfOcupante) = 14 And Mid(txtCpfOcupante, 4, 1) <> "." Then
        txtCpfOcupante.MaxLength = 20
        txtCpfOcupante = Edita.FormataTexto(txtCpfOcupante, Cgc)
    ElseIf Trim(txtCpfOcupante) <> "" And Len(txtCpfOcupante) <> 18 And Mid(txtCpfOcupante, 4, 1) <> "." Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
        txtCpfOcupante = ""
        txtCpfOcupante.SetFocus
    End If
    tabCad.Tab = 1
    DoEvents
    txtCodComponente(0).SetFocus
End Sub

Private Sub txtCpfOcupanteBc_LostFocus()
    If Len(txtCpfOcupanteBc) = 11 Then
        If Not Util.ValidaCpf(txtCpfOcupanteBc) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCpfOcupante = ""
            txtCpfOcupanteBc.SetFocus
            Exit Sub
        End If
        txtCpfOcupanteBc = Edita.FormataTexto(txtCpfOcupanteBc, Cpf)
    End If
End Sub

Private Sub txtDataAforamento_LostFocus()
    txtDataAforamento = Edita.FormataTexto(txtDataAforamento, Data)
End Sub

Private Sub txtDtRegistro_Validate(Cancel As Boolean)
    txtDtRegistro = Edita.FormataTexto(txtDtRegistro, Data)
End Sub

Private Sub txtIc_Change(Index As Integer)
' Função - Filtrar as informações pré-cadastradas na tela TCIM103 a fim de montar lotes de digitação e consistências nos mesmo
' Autor - Éderson - Imperatriz 30/01/2003
' Alteração

    If Len(txtIc(Index)) = txtIc(Index).MaxLength Then
        If txtIc(Index).Tag = "Quadra" Then
            If Not Lote.LoteCadastrado(txtIc(0), txtIc(1), txtIc(2)) Then
                Util.Informa "O Lote informado: " & txtIc(0) & "." & txtIc(1) & "." & txtIc(2) & ", não foi encontrado."
                Edita.LimpaCampos Me
                txtIc(0).SetFocus
                Exit Sub
            Else
                If Lote.LoteFechado(txtIc(0), txtIc(1), txtIc(2)) Then
                    Util.Informa "O Lote informado: " & txtIc(0) & "." & txtIc(1) & "." & txtIc(2) & ", já está fechado."
                    Edita.LimpaCampos Me
                    txtIc(0).SetFocus
                    Exit Sub
                End If
            End If
        End If
    SendKeys "{ENTER}"
    End If
End Sub

Private Sub txtic_LostFocus(Index As Integer)
    Dim Rs As VSRecordset
    Dim sql As String
    If Index = 4 Then
        If Trim(txtIc(Index)) = "" Then Exit Sub
        If CInt(Trim(txtIc(4))) < 200 Then txtIc(4) = "000"
        sql = "Select * from tab_imovel where (tIM_ic ='" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & _
            IIf(CInt(txtIc(4)) < 200, "000", txtIc(4)) & "'" & IIf(CInt(txtIc(4)) >= 200, _
            " AND TIM_UNIDADE =" & txtIc(4), "") & ") "
        If Bdados.AbreTabela(sql, Rs) Then
            Avisa "Imóvel já cadastrado."
            txtIc(4).SetFocus
        End If
    End If
End Sub

Private Sub txtIcAnterior_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCpfCgc_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIcAnterior_LostFocus()
    If Trim$(txtIcAnterior) <> "" Then
        txtIcAnterior = Edita.TiraPic(txtIcAnterior, ".")
        txtIcAnterior = Edita.BotaPic(txtIcAnterior, ".", 2)
        txtIcAnterior = Edita.BotaPic(txtIcAnterior, ".", 5)
        txtIcAnterior = Edita.BotaPic(txtIcAnterior, ".", 9)
        txtIcAnterior = Edita.BotaPic(txtIcAnterior, ".", 14)
        txtIcAnterior = Edita.BotaPic(txtIcAnterior, ".", 18)
    End If
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    Dim Rs As VSRecordset
    NovoContrib = False
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIM) <> "" Then
        txtIM = Cadastro.FormataInscricao(txtIM, InscContrib)
        sql = "Select tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, " & _
        " tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF,TCI_COD_LOGRADOURO,TCI_COD_BAIRRO from Tab_Contribuinte where tci_im = '" & txtIM & "'"
        If Bdados.AbreTabela(sql, Rs) Then
            txtNomeContrib = "" & Rs(0) 'Rs!tci_Nome
            'txtOcupante = txtNomeContrib
            txtNomeTipoLogrContrib = CStr("" & Rs(1))
            txtNomeLogrContrib = "" & Rs(2) '!tci_nome_logradouro
            txtNumeroContrib = "" & Rs(3) '!tci_numero
            txtCompContrib = "" & Rs(4) '!tci_complemento
            txtBairroContrib = "" & Rs(5) '!tci_bairro
            txtCep = "" & Rs(6) '!tci_cep
            txtMunic = Rs(7)
            txtUf = Rs(8) '!tci_UF
            txtCpfCgc = "" & Rs(9)
            txtCodLogrContrib = "" & Rs!tci_cod_logradouro
            txtCodBairroContrib = "" & Rs!tci_cod_bairro
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtIM.Enabled = True
            txtIM.SetFocus
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtIMBc_LostFocus()
    'NOVIDADE
    Dim Rs As VSRecordset
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIMBc) <> "" Then
        txtIMBc = Cadastro.FormataInscricao(txtIMBc, InscContrib)
        sql = "Select tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF,tci_cod_logradouro,tci_cod_bairro from Tab_Contribuinte where tci_im = '" & txtIMBc & "'"
        If Bdados.AbreTabela(sql, Rs) Then
            txtNomeContribBc = Rs(0)  'Rs!tci_Nome
            txtTipoLogrContribBc = Rs(1)
            txtNomeLogrContribBc = Rs(2)  '!tci_nome_logradouro
            txtNumeroContribBc = Rs(3)  '!tci_numero
            txtCompContribBc = Rs(4)  '!tci_complemento
            txtBairroContribBc = Rs(5)  '!tci_bairro
            txtCepBc = Rs(6)  '!tci_cep
            txtMunicBc = Rs(7)
            cboUFBc = Rs(8)  '!tci_UF
            txtCpfCgcBc = "" & Rs(9)
            txtCodLogrContribBc = "" & Rs!tci_cod_logradouro
            txtCodBairroContribBc = "" & Rs!tci_cod_bairro
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtIMBc.Enabled = True
            txtIMBc.SetFocus
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtInscImobiliaria_LostFocus()
    txtInscImobiliaria = Format(txtInscImobiliaria, "000")
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
'    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtLoteamento_KeyPress(KeyAscii As Integer)
'    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtMunic_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMunic_LostFocus()
    If Trim(txtMunic) = "" Then txtMunic = Aplicacoes.Municipio
End Sub

Private Sub txtMunicBc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMunicBc_LostFocus()
    If Trim(txtMunicBc) = "" Then txtMunicBc = Aplicacoes.Municipio
End Sub

Private Sub txtNomeContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomeContrib_LostFocus()
    'txtOcupante = txtNomeContrib
End Sub

Private Sub txtNomeContribBc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomeLogrContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomeLogrContribBc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumeroContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtOcupante_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtOcupanteBc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtQuadra_KeyPress(KeyAscii As Integer)
'    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub
