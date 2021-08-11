VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Begin VB.Form TCIU201 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU201"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabCad 
      Height          =   6015
      Left            =   30
      TabIndex        =   93
      Top             =   690
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Boletim Territorial"
      TabPicture(0)   =   "TCIU201.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(49)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstPesq"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtMotivo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra(9)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Detalhe BT"
      TabPicture(1)   =   "TCIU201.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra(5)"
      Tab(1).Control(1)=   "fra(3)"
      Tab(1).Control(2)=   "fra(4)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Boletim Predial"
      TabPicture(2)   =   "TCIU201.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra(6)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Boletim de Condomínio"
      TabPicture(3)   =   "TCIU201.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra(7)"
      Tab(3).Control(1)=   "fra(8)"
      Tab(3).Control(2)=   "lstCond"
      Tab(3).Control(3)=   "cmdAdCond"
      Tab(3).Control(4)=   "fra(2)"
      Tab(3).ControlCount=   5
      Begin Threed.SSFrame fra 
         Height          =   585
         Index           =   9
         Left            =   120
         TabIndex        =   209
         Top             =   2100
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   1032
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
            TabIndex        =   20
            Top             =   188
            Width           =   615
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
            TabIndex        =   19
            Top             =   188
            Width           =   855
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
            Left            =   360
            MaxLength       =   5
            TabIndex        =   18
            Top             =   188
            Width           =   840
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
            TabIndex        =   22
            Top             =   188
            Width           =   1215
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
            TabIndex        =   21
            Top             =   180
            Width           =   585
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
            TabIndex        =   23
            Top             =   180
            Width           =   585
         End
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
            TabIndex        =   24
            Top             =   195
            Width           =   990
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   79
            Left            =   4095
            TabIndex        =   210
            Top             =   240
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
            TabIndex        =   211
            Top             =   233
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
            TabIndex        =   200
            Top             =   240
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
            TabIndex        =   201
            Top             =   240
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
            TabIndex        =   208
            Top             =   240
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
            Index           =   89
            Left            =   7230
            TabIndex        =   198
            Top             =   225
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
            Index           =   90
            Left            =   8820
            TabIndex        =   199
            Top             =   240
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
      Begin VB.TextBox txtMotivo 
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
         Left            =   2025
         TabIndex        =   31
         Tag             =   "Motivo"
         Top             =   5595
         Width           =   9150
      End
      Begin Threed.SSFrame fra 
         Height          =   1695
         Index           =   0
         Left            =   120
         TabIndex        =   94
         Top             =   390
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   2990
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
         Begin VB.TextBox txtSecao 
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
            MaxLength       =   5
            TabIndex        =   13
            Top             =   990
            Width           =   645
         End
         Begin VB.TextBox txtCodReduzido 
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
            Left            =   5700
            TabIndex        =   5
            Top             =   240
            Width           =   1545
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
            Left            =   6720
            MaxLength       =   5
            TabIndex        =   17
            Top             =   1320
            Width           =   615
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
            Left            =   5670
            MaxLength       =   5
            TabIndex        =   16
            Top             =   1320
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
            Left            =   9555
            TabIndex        =   11
            Top             =   660
            Width           =   1425
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
            Left            =   8265
            MaxLength       =   10
            TabIndex        =   10
            Top             =   660
            Width           =   555
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
            ItemData        =   "TCIU201.frx":0070
            Left            =   1185
            List            =   "TCIU201.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Tag             =   "Tipo Imovel"
            Top             =   960
            Width           =   1455
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
            Left            =   1200
            TabIndex        =   9
            Tag             =   "Logradouro"
            Top             =   630
            Width           =   1485
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
            Left            =   3870
            TabIndex        =   34
            Tag             =   "Nome Contribuinte"
            Top             =   990
            Width           =   4425
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
            Left            =   3780
            TabIndex        =   33
            Tag             =   "Nome Contribuinte"
            Top             =   630
            Width           =   3525
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
            Left            =   2730
            MaxLength       =   11
            TabIndex        =   32
            Top             =   630
            Width           =   1035
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
            Left            =   3300
            MaxLength       =   50
            TabIndex        =   12
            Tag             =   "Bairro"
            Top             =   990
            Width           =   525
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
            Left            =   10665
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Cod Mensagem"
            Top             =   990
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
            Left            =   2100
            MaxLength       =   2
            TabIndex        =   0
            Tag             =   "Distrito"
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
            Index           =   1
            Left            =   2430
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
            Index           =   2
            Left            =   2760
            MaxLength       =   4
            TabIndex        =   2
            Tag             =   "Quadra"
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
            Index           =   3
            Left            =   3270
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
            Index           =   4
            Left            =   3780
            MaxLength       =   3
            TabIndex        =   4
            Tag             =   "Unidade"
            Top             =   240
            Width           =   375
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
            Left            =   8550
            TabIndex        =   6
            Top             =   240
            Width           =   2415
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
            Left            =   10440
            MaxLength       =   10
            TabIndex        =   7
            TabStop         =   0   'False
            Tag             =   "Zona"
            Top             =   -360
            Width           =   555
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   2
            Left            =   8865
            TabIndex        =   95
            Top             =   720
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
            Index           =   1
            Left            =   7935
            TabIndex        =   96
            Top             =   720
            Width           =   390
            _ExtentX        =   688
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
            Caption         =   "N.º:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   3
            Left            =   2700
            TabIndex        =   97
            Top             =   1050
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
            Caption         =   "Bairro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   5
            Left            =   5310
            TabIndex        =   98
            Top             =   1380
            Width           =   660
            _ExtentX        =   1164
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
            Caption         =   "Qd.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   6
            Left            =   6270
            TabIndex        =   99
            Top             =   1350
            Width           =   435
            _ExtentX        =   767
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
            Caption         =   "Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   7
            Left            =   615
            TabIndex        =   100
            Top             =   1020
            Width           =   420
            _ExtentX        =   741
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
            Caption         =   "Tipo:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   210
            Index           =   82
            Left            =   270
            TabIndex        =   101
            Top             =   690
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
            Caption         =   "Cód. Logr:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   50
            Left            =   9645
            TabIndex        =   102
            Top             =   1050
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
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   72
            Left            =   7320
            TabIndex        =   103
            Top             =   270
            Width           =   1185
            _ExtentX        =   2090
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
            Caption         =   "Insc. Anterior:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   75
            Left            =   750
            TabIndex        =   104
            Top             =   285
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "Insc. Imobiliária:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   76
            Left            =   9960
            TabIndex        =   105
            Top             =   -300
            Width           =   390
            _ExtentX        =   688
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
            Caption         =   "Zona:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cboVISUAL cboLoteamento 
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   1350
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   556
            Caption         =   "Loteamento"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   4
            Left            =   4350
            TabIndex        =   244
            Top             =   285
            Width           =   1425
            _ExtentX        =   2514
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
            Caption         =   "Cad. Imobiliário:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   91
            Left            =   8340
            TabIndex        =   245
            Top             =   1035
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
            Caption         =   "Secão:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1785
         Index           =   1
         Left            =   120
         TabIndex        =   106
         Top             =   2715
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
            Height          =   330
            Left            =   10425
            MaxLength       =   50
            TabIndex        =   42
            Tag             =   "Bairro"
            Top             =   945
            Width           =   585
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
            Left            =   1020
            TabIndex        =   28
            Tag             =   "Logradouro"
            Top             =   585
            Width           =   645
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
            Left            =   2085
            MaxLength       =   11
            TabIndex        =   35
            Top             =   585
            Width           =   1035
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
            Left            =   7560
            TabIndex        =   38
            Top             =   585
            Width           =   735
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
            Left            =   1470
            TabIndex        =   25
            Top             =   210
            Width           =   1305
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
            Left            =   3660
            TabIndex        =   26
            Tag             =   "Nome Contribuinte"
            Top             =   210
            Width           =   4665
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
            Left            =   3225
            TabIndex        =   36
            Tag             =   "Nome Logradouro"
            Top             =   585
            Width           =   2370
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
            Left            =   3060
            MaxLength       =   10
            TabIndex        =   40
            Top             =   960
            Width           =   1125
         End
         Begin VB.CommandButton cmdEnter 
            Caption         =   "Command1"
            Default         =   -1  'True
            Height          =   255
            Left            =   7740
            TabIndex        =   107
            Top             =   3090
            Width           =   375
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
            Left            =   8970
            TabIndex        =   39
            Tag             =   "Bairro"
            Top             =   585
            Width           =   2040
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
            Left            =   3060
            TabIndex        =   29
            Top             =   1335
            Width           =   4965
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
            Left            =   8970
            MaxLength       =   20
            TabIndex        =   30
            Top             =   1335
            Width           =   2040
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
            Left            =   9240
            MaxLength       =   20
            TabIndex        =   27
            Top             =   210
            Width           =   1770
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
            Left            =   6180
            MaxLength       =   10
            TabIndex        =   37
            Top             =   585
            Width           =   525
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
            Left            =   5880
            TabIndex        =   41
            Tag             =   "Município"
            Top             =   960
            Width           =   4335
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   8
            Left            =   150
            TabIndex        =   108
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
            Index           =   14
            Left            =   2640
            TabIndex        =   109
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
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   15
            Left            =   4995
            TabIndex        =   110
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
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   13
            Left            =   5805
            TabIndex        =   111
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
            Index           =   17
            Left            =   8370
            TabIndex        =   112
            Top             =   600
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
            Index           =   16
            Left            =   6810
            TabIndex        =   113
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
            Index           =   29
            Left            =   8385
            TabIndex        =   114
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
            Index           =   11
            Left            =   2175
            TabIndex        =   115
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
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   330
            Index           =   18
            Left            =   8085
            TabIndex        =   116
            Top             =   1320
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "CPF/CNPJ:"
            BorderWidth     =   1
            BevelOuter      =   0
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   12
            Left            =   105
            TabIndex        =   197
            Top             =   630
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
         Begin VTOcx.cmdVISUAL cmdNovo 
            Height          =   315
            Left            =   3165
            TabIndex        =   204
            Top             =   210
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   ""
            Acao            =   1
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   315
            Index           =   0
            Left            =   2790
            TabIndex        =   205
            Top             =   210
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   ""
            Acao            =   5
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   315
            Index           =   3
            Left            =   1680
            TabIndex        =   206
            Top             =   585
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   ""
            Acao            =   5
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1395
         Index           =   5
         Left            =   -74880
         TabIndex        =   117
         Top             =   480
         Width           =   11115
         _ExtentX        =   19606
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
            ItemData        =   "TCIU201.frx":0094
            Left            =   6930
            List            =   "TCIU201.frx":0096
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   1950
            Width           =   3015
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
            ItemData        =   "TCIU201.frx":0098
            Left            =   7410
            List            =   "TCIU201.frx":009A
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   2310
            Width           =   2535
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
            ItemData        =   "TCIU201.frx":009C
            Left            =   7680
            List            =   "TCIU201.frx":009E
            Style           =   2  'Dropdown List
            TabIndex        =   122
            TabStop         =   0   'False
            Tag             =   "5"
            Top             =   630
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
            ItemData        =   "TCIU201.frx":00A0
            Left            =   7680
            List            =   "TCIU201.frx":00A2
            Style           =   2  'Dropdown List
            TabIndex        =   121
            TabStop         =   0   'False
            Tag             =   "4"
            Top             =   270
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
            ItemData        =   "TCIU201.frx":00A4
            Left            =   2070
            List            =   "TCIU201.frx":00A6
            Style           =   2  'Dropdown List
            TabIndex        =   120
            TabStop         =   0   'False
            Tag             =   "3"
            Top             =   990
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
            ItemData        =   "TCIU201.frx":00A8
            Left            =   2070
            List            =   "TCIU201.frx":00AA
            Style           =   2  'Dropdown List
            TabIndex        =   119
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   630
            Width           =   3375
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
            ItemData        =   "TCIU201.frx":00AC
            Left            =   2070
            List            =   "TCIU201.frx":00AE
            Style           =   2  'Dropdown List
            TabIndex        =   118
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   270
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
            Index           =   0
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   43
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
            Index           =   3
            Left            =   7260
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
            Index           =   4
            Left            =   7260
            MaxLength       =   3
            TabIndex        =   47
            Top             =   630
            Width           =   375
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   30
            Left            =   6150
            TabIndex        =   125
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
            Height          =   270
            Index           =   32
            Left            =   6480
            TabIndex        =   126
            Top             =   300
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
            Caption         =   "Limites:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   33
            Left            =   330
            TabIndex        =   127
            Top             =   1020
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
            Index           =   34
            Left            =   60
            TabIndex        =   128
            Top             =   300
            Width           =   1530
            _ExtentX        =   2699
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
            TabIndex        =   129
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
            TabIndex        =   130
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
            TabIndex        =   131
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
         Left            =   -74910
         TabIndex        =   132
         Top             =   1860
         Width           =   11115
         _ExtentX        =   19606
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
            ItemData        =   "TCIU201.frx":00B0
            Left            =   2070
            List            =   "TCIU201.frx":00B2
            Style           =   2  'Dropdown List
            TabIndex        =   135
            TabStop         =   0   'False
            Tag             =   "7"
            Top             =   600
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
            ItemData        =   "TCIU201.frx":00B4
            Left            =   7650
            List            =   "TCIU201.frx":00B6
            Style           =   2  'Dropdown List
            TabIndex        =   134
            TabStop         =   0   'False
            Tag             =   "8"
            Top             =   240
            Width           =   3405
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
            ItemData        =   "TCIU201.frx":00B8
            Left            =   2070
            List            =   "TCIU201.frx":00BA
            Style           =   2  'Dropdown List
            TabIndex        =   133
            TabStop         =   0   'False
            Tag             =   "6"
            Top             =   240
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
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   48
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
            Index           =   7
            Left            =   7260
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
            Index           =   6
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   49
            Top             =   570
            Width           =   375
         End
         Begin Threed.SSPanel lbl 
            Height          =   315
            Index           =   20
            Left            =   690
            TabIndex        =   136
            Top             =   600
            Width           =   945
            _ExtentX        =   1667
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
            Height          =   270
            Index           =   21
            Left            =   870
            TabIndex        =   137
            Top             =   210
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
            Left            =   6330
            TabIndex        =   138
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
         Height          =   1005
         Index           =   4
         Left            =   -74910
         TabIndex        =   139
         Top             =   2820
         Width           =   11145
         _ExtentX        =   19659
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
         Caption         =   "Dimensões do Terreno (m²)"
         Alignment       =   2
         ShadowStyle     =   1
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
            Left            =   8280
            TabIndex        =   58
            Tag             =   "107"
            Top             =   510
            Width           =   735
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
            Left            =   10350
            TabIndex        =   59
            Tag             =   "108"
            Top             =   210
            Width           =   735
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
            Left            =   5880
            TabIndex        =   56
            Tag             =   "105"
            Top             =   600
            Width           =   735
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
            Left            =   5880
            TabIndex        =   55
            Tag             =   "104"
            Top             =   240
            Width           =   735
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
            Left            =   8280
            TabIndex        =   57
            Tag             =   "106"
            Top             =   180
            Width           =   735
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
            Left            =   3780
            TabIndex        =   54
            Tag             =   "103"
            Top             =   540
            Width           =   735
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
            Left            =   3780
            TabIndex        =   53
            Tag             =   "102"
            Top             =   210
            Width           =   735
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
            Left            =   1650
            TabIndex        =   52
            Tag             =   "101"
            Top             =   570
            Width           =   735
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
            Left            =   1650
            TabIndex        =   51
            Tag             =   "100"
            Top             =   240
            Width           =   735
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   23
            Left            =   4590
            TabIndex        =   140
            Top             =   240
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
            Caption         =   "Trecho Logr. 3:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   24
            Left            =   6930
            TabIndex        =   141
            Top             =   240
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
            Caption         =   "Trecho Logr. 4:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   330
            Index           =   25
            Left            =   2910
            TabIndex        =   142
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Testada 3:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   26
            Left            =   2490
            TabIndex        =   143
            Top             =   270
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
            Caption         =   "Trecho Logr. 2:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   27
            Left            =   120
            TabIndex        =   144
            Top             =   240
            Width           =   1470
            _ExtentX        =   2593
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
            Caption         =   "Testada Principal:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   28
            Left            =   720
            TabIndex        =   145
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Testada 2:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   19
            Left            =   4950
            TabIndex        =   146
            Top             =   570
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "Testada 4:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   45
            Left            =   6780
            TabIndex        =   147
            Top             =   600
            Width           =   1425
            _ExtentX        =   2514
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
            Caption         =   "Testada(Campo):"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   51
            Left            =   9210
            TabIndex        =   148
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
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
            Caption         =   "Área do Lote:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   5325
         Index           =   6
         Left            =   -74880
         TabIndex        =   149
         Top             =   480
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
            ItemData        =   "TCIU201.frx":00BC
            Left            =   2040
            List            =   "TCIU201.frx":00BE
            Style           =   2  'Dropdown List
            TabIndex        =   161
            TabStop         =   0   'False
            Tag             =   "13"
            Top             =   2580
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
            ItemData        =   "TCIU201.frx":00C0
            Left            =   2040
            List            =   "TCIU201.frx":00C2
            Style           =   2  'Dropdown List
            TabIndex        =   160
            TabStop         =   0   'False
            Tag             =   "12"
            Top             =   2220
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
            ItemData        =   "TCIU201.frx":00C4
            Left            =   7440
            List            =   "TCIU201.frx":00C6
            Style           =   2  'Dropdown List
            TabIndex        =   159
            TabStop         =   0   'False
            Tag             =   "9"
            Top             =   1080
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
            ItemData        =   "TCIU201.frx":00C8
            Left            =   2040
            List            =   "TCIU201.frx":00CA
            Style           =   2  'Dropdown List
            TabIndex        =   158
            TabStop         =   0   'False
            Tag             =   "11"
            Top             =   1860
            Width           =   3615
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
            ItemData        =   "TCIU201.frx":00CC
            Left            =   7440
            List            =   "TCIU201.frx":00CE
            Style           =   2  'Dropdown List
            TabIndex        =   157
            TabStop         =   0   'False
            Tag             =   "10"
            Top             =   1440
            Width           =   3615
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
            Left            =   7050
            TabIndex        =   72
            Tag             =   "112"
            Top             =   2280
            Width           =   1185
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
            Left            =   9870
            TabIndex        =   73
            Tag             =   "114"
            Top             =   2280
            Width           =   1155
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
            ItemData        =   "TCIU201.frx":00D0
            Left            =   2040
            List            =   "TCIU201.frx":00D2
            Style           =   2  'Dropdown List
            TabIndex        =   156
            TabStop         =   0   'False
            Tag             =   "14"
            Top             =   720
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
            ItemData        =   "TCIU201.frx":00D4
            Left            =   7440
            List            =   "TCIU201.frx":00D6
            Style           =   2  'Dropdown List
            TabIndex        =   155
            TabStop         =   0   'False
            Tag             =   "16"
            Top             =   720
            Width           =   3615
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
            ItemData        =   "TCIU201.frx":00D8
            Left            =   2040
            List            =   "TCIU201.frx":00DA
            Style           =   2  'Dropdown List
            TabIndex        =   154
            TabStop         =   0   'False
            Tag             =   "15"
            Top             =   1080
            Width           =   3615
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
            TabIndex        =   63
            Tag             =   "110"
            Top             =   1440
            Width           =   825
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
            TabIndex        =   60
            Top             =   210
            Width           =   735
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
            Left            =   7050
            MaxLength       =   3
            TabIndex        =   64
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
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   62
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
            Index           =   13
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   61
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
            Index           =   8
            Left            =   7050
            MaxLength       =   3
            TabIndex        =   65
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
            Index           =   10
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   67
            Top             =   1860
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
            TabIndex        =   69
            Top             =   2580
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
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   68
            Top             =   2220
            Width           =   375
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
            Left            =   9870
            TabIndex        =   71
            Tag             =   "113"
            Top             =   1860
            Width           =   1155
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
            Left            =   7050
            TabIndex        =   70
            Tag             =   "111"
            Top             =   1860
            Width           =   1185
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
            Left            =   7050
            MaxLength       =   3
            TabIndex        =   66
            Top             =   1440
            Width           =   375
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
            TabIndex        =   153
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
            Index           =   11
            Left            =   1980
            MaxLength       =   2
            TabIndex        =   152
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
            Index           =   12
            Left            =   2310
            MaxLength       =   4
            TabIndex        =   151
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
            Index           =   13
            Left            =   2820
            MaxLength       =   4
            TabIndex        =   150
            Top             =   210
            Width           =   495
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   36
            Left            =   6045
            TabIndex        =   162
            Top             =   1500
            Width           =   825
            _ExtentX        =   1455
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
            Caption         =   "Estrutura:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   38
            Left            =   6060
            TabIndex        =   163
            Top             =   1080
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
            Caption         =   "Tipologia:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   240
            Index           =   39
            Left            =   510
            TabIndex        =   164
            Top             =   1920
            Width           =   960
            _ExtentX        =   1693
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
            Caption         =   "Destinação:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   40
            Left            =   840
            TabIndex        =   165
            Top             =   2280
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
            Caption         =   "Padrão:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   240
            Index           =   41
            Left            =   330
            TabIndex        =   166
            Top             =   2670
            Width           =   1140
            _ExtentX        =   2011
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
            Caption         =   "Conservação:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   42
            Left            =   5820
            TabIndex        =   167
            Top             =   1905
            Width           =   1050
            _ExtentX        =   1852
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
            Caption         =   "Ano Constr.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   44
            Left            =   8505
            TabIndex        =   168
            Top             =   1905
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
            Caption         =   "Área Edif. Total:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   47
            Left            =   780
            TabIndex        =   169
            Top             =   720
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
            Caption         =   "Sentido:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   48
            Left            =   6495
            TabIndex        =   170
            Top             =   750
            Width           =   375
            _ExtentX        =   661
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
            Caption         =   "Uso:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   52
            Left            =   870
            TabIndex        =   171
            Top             =   1080
            Width           =   600
            _ExtentX        =   1058
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
            Caption         =   "Predio:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   53
            Left            =   435
            TabIndex        =   172
            Top             =   1470
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
            Caption         =   "Pavimentos:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin MSComctlLib.ListView lstEdific 
            Height          =   1815
            Left            =   90
            TabIndex        =   173
            Top             =   3420
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   3201
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
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   14
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
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   9
            Left            =   120
            TabIndex        =   174
            Top             =   270
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
            Caption         =   "Insc. Imobiliária:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   43
            Left            =   6030
            TabIndex        =   175
            Top             =   2302
            Width           =   840
            _ExtentX        =   1482
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
            Caption         =   "Área Edif.:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   46
            Left            =   8760
            TabIndex        =   176
            Top             =   2302
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
            Caption         =   "Fração Ideal:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdAdEdif 
            Height          =   375
            Left            =   120
            TabIndex        =   207
            Top             =   3000
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   661
            Caption         =   "&Adicionar Edificação"
            Acao            =   1
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1785
         Index           =   7
         Left            =   -74880
         TabIndex        =   177
         Top             =   2100
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
            Left            =   7290
            TabIndex        =   84
            Top             =   585
            Width           =   735
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
            Left            =   1470
            MaxLength       =   14
            TabIndex        =   78
            Top             =   210
            Width           =   1305
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
            TabIndex        =   79
            Tag             =   "Nome Contribuinte"
            Top             =   210
            Width           =   4965
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
            TabIndex        =   82
            Tag             =   "Nome Logradouro"
            Top             =   585
            Width           =   2415
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
            Left            =   3060
            MaxLength       =   10
            TabIndex        =   86
            Top             =   960
            Width           =   1125
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
            Left            =   8955
            TabIndex        =   85
            Tag             =   "Bairro"
            Top             =   585
            Width           =   2040
         End
         Begin VB.ComboBox cboTipoLogrContribBc 
            BeginProperty Font 
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
            ItemData        =   "TCIU201.frx":00DC
            Left            =   1470
            List            =   "TCIU201.frx":00E9
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Tag             =   "Logradouro"
            Top             =   570
            Width           =   1365
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
            TabIndex        =   89
            Top             =   1335
            Width           =   4965
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
            TabIndex        =   90
            Top             =   1335
            Width           =   2040
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
            TabIndex        =   80
            Top             =   210
            Width           =   2040
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
            ItemData        =   "TCIU201.frx":010A
            Left            =   10215
            List            =   "TCIU201.frx":0117
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Tag             =   "UF"
            Top             =   945
            Width           =   795
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
            Left            =   5880
            MaxLength       =   10
            TabIndex        =   83
            Top             =   585
            Width           =   525
         End
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
            Left            =   5880
            TabIndex        =   87
            Tag             =   "Município"
            Top             =   960
            Width           =   4275
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   62
            Left            =   150
            TabIndex        =   178
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
            Index           =   63
            Left            =   390
            TabIndex        =   179
            Top             =   630
            Width           =   1035
            _ExtentX        =   1826
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
            Caption         =   "Logradouro:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   64
            Left            =   2640
            TabIndex        =   180
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
            Left            =   4995
            TabIndex        =   181
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
            Left            =   5580
            TabIndex        =   182
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
            Left            =   8370
            TabIndex        =   183
            Top             =   600
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
            Left            =   6510
            TabIndex        =   184
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
            TabIndex        =   185
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
            TabIndex        =   186
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
            TabIndex        =   187
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
      End
      Begin Threed.SSFrame fra 
         Height          =   615
         Index           =   8
         Left            =   -74880
         TabIndex        =   188
         Top             =   3840
         Width           =   8625
         _ExtentX        =   15214
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
         Caption         =   "Características do Imóvel:"
         Alignment       =   2
         ShadowStyle     =   1
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
            ItemData        =   "TCIU201.frx":0138
            Left            =   6930
            List            =   "TCIU201.frx":013A
            Style           =   2  'Dropdown List
            TabIndex        =   192
            Top             =   1950
            Width           =   3015
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
            ItemData        =   "TCIU201.frx":013C
            Left            =   7410
            List            =   "TCIU201.frx":013E
            Style           =   2  'Dropdown List
            TabIndex        =   191
            Top             =   2310
            Width           =   2535
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
            ItemData        =   "TCIU201.frx":0140
            Left            =   2010
            List            =   "TCIU201.frx":0142
            Style           =   2  'Dropdown List
            TabIndex        =   189
            TabStop         =   0   'False
            Tag             =   "3"
            Top             =   195
            Width           =   6525
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
            Index           =   20
            Left            =   1455
            MaxLength       =   3
            TabIndex        =   91
            Top             =   210
            Width           =   495
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   74
            Left            =   105
            TabIndex        =   193
            Top             =   270
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
            TabIndex        =   194
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
            TabIndex        =   195
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
      Begin MSComctlLib.ListView lstCond 
         Height          =   1410
         Left            =   -74880
         TabIndex        =   196
         Top             =   4500
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   2487
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
         NumItems        =   19
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
      End
      Begin MSComctlLib.ListView lstPesq 
         Height          =   1005
         Left            =   105
         TabIndex        =   202
         Top             =   4575
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   1773
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
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
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   49
         Left            =   120
         TabIndex        =   203
         Top             =   5625
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
         Caption         =   "Motivo da Modificação:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdAdCond 
         Height          =   375
         Left            =   -66165
         TabIndex        =   190
         Top             =   3975
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "Adicionar &Condomínio"
         Acao            =   1
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin Threed.SSFrame fra 
         Height          =   1815
         Index           =   2
         Left            =   -74880
         TabIndex        =   215
         Top             =   270
         Width           =   11085
         _ExtentX        =   19553
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
         Caption         =   "Referência Cadastral / Localização do Imóvel"
         Alignment       =   2
         ShadowStyle     =   1
         Begin VB.TextBox txtBairro 
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
            Left            =   5610
            TabIndex        =   228
            Tag             =   "Nome Contribuinte"
            Top             =   1410
            Width           =   3525
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
            TabIndex        =   227
            Tag             =   "Nome Contribuinte"
            Top             =   667
            Width           =   3255
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
            TabIndex        =   226
            Top             =   667
            Width           =   1035
         End
         Begin VB.TextBox txtCodLogrBc 
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
            Left            =   1260
            TabIndex        =   225
            Top             =   667
            Width           =   1485
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
            Left            =   6120
            TabIndex        =   75
            Top             =   285
            Width           =   1665
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
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   74
            Tag             =   "Unidade"
            Top             =   285
            Width           =   375
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
            Left            =   3300
            MaxLength       =   4
            TabIndex        =   224
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
            Index           =   7
            Left            =   2310
            MaxLength       =   4
            TabIndex        =   223
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
            Index           =   6
            Left            =   1260
            MaxLength       =   2
            TabIndex        =   222
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
            Index           =   5
            Left            =   540
            MaxLength       =   2
            TabIndex        =   221
            Tag             =   "Distrito"
            Top             =   285
            Width           =   315
         End
         Begin VB.TextBox txtLoteBc 
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
            Left            =   4110
            MaxLength       =   5
            TabIndex        =   220
            Top             =   1410
            Width           =   765
         End
         Begin VB.TextBox txtQuadraBc 
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
            Left            =   2700
            MaxLength       =   5
            TabIndex        =   219
            Top             =   1410
            Width           =   705
         End
         Begin VB.TextBox txtLoteamentoBc 
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
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   218
            Top             =   1410
            Width           =   705
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
            TabIndex        =   77
            Top             =   667
            Width           =   2205
         End
         Begin VB.TextBox txtNumeroBc 
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
            Left            =   7470
            MaxLength       =   10
            TabIndex        =   217
            Top             =   667
            Width           =   525
         End
         Begin VB.TextBox txtCepImBc 
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
            Left            =   9960
            MaxLength       =   10
            TabIndex        =   216
            Top             =   1410
            Width           =   1035
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
            ItemData        =   "TCIU201.frx":0144
            Left            =   9450
            List            =   "TCIU201.frx":014E
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Tag             =   "Logradouro"
            Top             =   277
            Width           =   1545
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   10
            Left            =   90
            TabIndex        =   229
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
            TabIndex        =   230
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
            TabIndex        =   231
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
            TabIndex        =   232
            Top             =   1455
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
            TabIndex        =   233
            Top             =   1455
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
            TabIndex        =   234
            Top             =   1455
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
            TabIndex        =   235
            Top             =   1455
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
            TabIndex        =   236
            Top             =   1455
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
            TabIndex        =   237
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
            Index           =   0
            Left            =   150
            TabIndex        =   238
            Top             =   330
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
            Caption         =   "Dist:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   73
            Left            =   930
            TabIndex        =   239
            Top             =   330
            Width           =   330
            _ExtentX        =   582
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
            Caption         =   "Set:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   83
            Left            =   1650
            TabIndex        =   240
            Top             =   330
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
            Caption         =   "Quadra:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   84
            Left            =   2820
            TabIndex        =   241
            Top             =   330
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
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   85
            Left            =   3900
            TabIndex        =   242
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
            Caption         =   "Unid:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   4
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   86
            Left            =   4890
            TabIndex        =   243
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
         Begin VTOcx.cboVISUAL cboEdificio 
            Height          =   315
            Left            =   630
            TabIndex        =   246
            Top             =   1020
            Width           =   6525
            _ExtentX        =   11509
            _ExtentY        =   556
            Caption         =   "Edificio"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VB.Shape Shape2 
            Height          =   405
            Left            =   90
            Top             =   240
            Width           =   4695
         End
      End
   End
   Begin VB.TextBox txtFatorFixo 
      Height          =   285
      Left            =   8640
      TabIndex        =   92
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   4560
      Width           =   375
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   2
      Left            =   10230
      TabIndex        =   212
      Top             =   6750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   213
      Top             =   6750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   0
      Left            =   9075
      TabIndex        =   214
      Top             =   6750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Limpar"
      Acao            =   6
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   247
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   1138
      Icone           =   "TCIU201.frx":0168
   End
End
Attribute VB_Name = "TCIU201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim cadastro As VSImposto
Dim NovoContrib As Boolean
Dim sql As String
Private Boletim As TipoBoletim
Dim Consultando As Boolean
Dim Lote As New BCI

Private Sub AtualizaCodComponente(Combo As ComboBox)
    On Error Resume Next
    With Combo
        If .ListIndex + 1 = 0 Then
            txtCodComponente(Val(.Tag) - 1).Text = ""
        Else
            txtCodComponente(Val(.Tag) - 1).Text = .ListIndex + 1
        End If
    End With
    
End Sub

Function TotalProva(Valor As String) As String
    Static Total As Double
    If Trim(Valor) = "" Then Valor = "0"
    Total = CDbl(Valor) + Total
    TotalProva = Total
End Function

Public Sub HabilitaCaixa(Status As Boolean)
    txtIM.Enabled = Not Status
    txtNomeContrib.Enabled = Status
    
    txtNomeLogrContrib.Enabled = Status
    txtNumeroContrib.Enabled = Status
    txtCompContrib.Enabled = Status
    txtBairroContrib.Enabled = Status
    txtCEP.Enabled = Status
    txtMunic.Enabled = Status
    txtIM = ""
    txtNomeContrib = ""
        
    txtNomeLogrContrib = ""
    txtNumeroContrib = ""
    txtCompContrib = ""
    txtBairroContrib = ""
    txtCEP = ""
    txtMunic = ""
    txtCpfCgc = ""
    txtOcupante = ""
    
    If Status Then txtNomeContrib.SetFocus
End Sub

Private Sub cboArborizacao_Click()
'    AtualizaCodComponente cboArborizacao
End Sub

Private Sub cboArborizacao_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboArborizacao
End Sub

Private Sub cboArborizacao_Scroll()
Dim a
a = a
End Sub

Private Sub cboCobranca_Click()
'    AtualizaCodComponente cboCobranca
End Sub

Private Sub cboCobranca_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboCobranca
End Sub

Private Sub cboConservacao_Click()
'    AtualizaCodComponente cboConservacao
End Sub

Private Sub cboConservacao_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboConservacao
End Sub

Private Sub cboDestinacao_Click()
'    AtualizaCodComponente cboDestinacao
End Sub

Private Sub cboDestinacao_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboDestinacao
End Sub

Private Sub cboEstrutura_Click()
'    AtualizaCodComponente cboEstrutura
End Sub

Private Sub cboEstrutura_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboEstrutura
End Sub

Private Sub cboLimites_Click()
'    AtualizaCodComponente cboLimites
End Sub

Private Sub cboLimites_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboLimites
End Sub

Private Sub cboOcupLote_Click()
'    AtualizaCodComponente cboOcupLote
End Sub

Private Sub cboOcupLote_KeyPress(KeyAscii As Integer)
'    cboOcupLote_Click
End Sub

Private Sub cboPadrao_Click()
'    AtualizaCodComponente cboPadrao
End Sub

Private Sub cboPadrao_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboPadrao
End Sub

Private Sub cboPatrimonio_Click()
'    AtualizaCodComponente cboPatrimonio
End Sub

Private Sub cboPatrimonio_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboPatrimonio
End Sub

Private Sub cboPedol_Click()
'    AtualizaCodComponente cboPedol
End Sub

Private Sub cboPedol_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboPedol
End Sub

Private Sub cboPredio_Click()
'    AtualizaCodComponente cboPredio
End Sub

Private Sub cboPredio_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboPredio
End Sub

Private Sub cboSentido_Click()
'    AtualizaCodComponente cboSentido
End Sub

Private Sub cboSentido_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboSentido
End Sub

Private Sub cboSit_Click()
'    AtualizaCodComponente cboSit
End Sub

Private Sub cboSit_KeyPress(KeyAscii As Integer)
'    AtualizaCodComponente cboSit
End Sub

Private Sub cboTipoImovel_Click()
    If cboTipoImovel = "PREDIAL" Then
        tabCad.TabEnabled(0) = True
        tabCad.TabEnabled(1) = True
        tabCad.TabEnabled(2) = True
        tabCad.TabEnabled(3) = True
        Boletim = tbo_Predial
    Else
        tabCad.TabEnabled(2) = False
        tabCad.TabEnabled(3) = False
        Boletim = tbo_Territorial
    End If
End Sub

Private Sub cmdAdCond_Click()
     'NOVIDADE
    Dim ItmX As Object
    Dim i As Byte
    
    On Error Resume Next
    Set ItmX = lstCond.ListItems.Add(, , txtic(9))
    ItmX.SubItems(1) = txtInscAnteriorBC
    ItmX.SubItems(2) = cboTipoImovelBc.ListIndex + 1
    ItmX.SubItems(3) = txtNumeroBc
    ItmX.SubItems(4) = txtComplementoBc
    ItmX.SubItems(5) = IIf(Trim(txtIMBc) = "", "", txtIMBc)
    ItmX.SubItems(6) = txtNomeContribBc
    ItmX.SubItems(7) = txtCpfCgcBc
    ItmX.SubItems(8) = cboTipoLogrContribBc
    ItmX.SubItems(9) = txtNomeLogrContribBc
    ItmX.SubItems(10) = txtNumeroBc
    ItmX.SubItems(11) = txtComplementoBc
    ItmX.SubItems(12) = txtBairroContribBc
    ItmX.SubItems(13) = txtCepImBc
    ItmX.SubItems(14) = txtMunicBc
    ItmX.SubItems(15) = cboUFBc
    ItmX.SubItems(16) = txtOcupanteBc
    ItmX.SubItems(17) = txtCpfOcupanteBc
    ItmX.SubItems(18) = txtCodComponente(20)
    txtic(9).SetFocus
End Sub

Private Sub cmdAdEdif_Click()
    Dim ItmX As Object
    Dim i As Byte
    If Trim(txtInscImobiliaria) = "" Then
        Avisa "Informe a unidade."
        txtInscImobiliaria.SetFocus
        Exit Sub
    End If
    For i = 8 To 15
        If Trim(txtCodComponente(i)) = "" Then
            Avisa "Informe todos os campos."
            txtCodComponente(i).SetFocus
            Exit Sub
        End If
    Next
    
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
    
    For i = 8 To 15
        txtCodComponente(i) = ""
    Next
    txtAnoConst = ""
    txtAreaEdif = ""
    txtAreaEdifTotal = ""
    txtFracaoEdif = ""
    txtPavimento = ""
    txtInscImobiliaria = ""
    txtInscImobiliaria.SetFocus
    
End Sub

Private Sub cmdEnter_Click()
        SendKeys "{Tab}"
End Sub

Private Sub cmdImprime_Click()
    If Me.Tag <> "" Then
        With Rpt
            If Not .DefinirArquivo(Bdados, App.Path & "\TCIU201.rpt") Then Exit Sub
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
            .Selecao = "{TAB_IMOVEL.tim_ic} = '" & Me.Tag & "'"
            .Titulo = "Ficha Cadastral"
            .Arvore = False
            .Visualizar
            DoEvents
        End With
        Set Rpt = Nothing
    End If
End Sub

Private Sub cmdNovo_Click()
    Static Status As Boolean
    Status = Not Status
    HabilitaCaixa Status
End Sub

Private Sub cmdOpcao_Click(Index As Integer)
    Dim Rs As VSRecordset
    Select Case Index
        Case 0
            NovoContrib = False
            sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where tci_nome like '" & txtNomeContrib & "%' or tci_nome like '% " & txtNomeContrib & "%'"
            sql = sql & " and tci_tsc_cod_sit_cad =1"
            If Not Bdados.AbreTabela(sql, Rs) Then
                Call Util.Avisa("Nenhum contribuinte encontrado.")
            End If
            Bdados.FechaTabela Rs
            MontaGrid Bdados, lstPesq, sql, 1400
        Case 1
            NovoContrib = True
            txtIM = ""
            Call HabilitaCaixa(True)
            txtCEP = Temp.PegaParametro(Bdados, "CEP")
            txtNomeContrib.SetFocus
    End Select
End Sub

Private Sub Form_Activate()
    
    Dim i As Byte
    Dim tam_quadra As Byte
    
    If Me.Tag <> "" Then
        Consultando = True
        For i = 0 To 8
            fra(i).Enabled = False
        Next
        tam_quadra = CInt(Nvl(Temp.PegaParametro(Bdados, "CAMPO QUADRA"), 4))
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            txtCodReduzido = Mid(Trim(Me.Tag), 2)
            txtCodReduzido_LostFocus
            fra(9).Visible = False
        Else
            txtic(0) = Left(Me.Tag, 2)
            txtic(1) = Mid(Me.Tag, 3, 2)
            txtic(2) = Mid(Me.Tag, 5, tam_quadra)
            txtic(3) = Mid(Me.Tag, 5 + tam_quadra, 4)
            txtic(4) = Right(Trim(Me.Tag), 3) 'IIf(Right(Trim(Me.Tag), 3) < 200, "000", IIf(Right(Trim(Me.Tag), 3) < 600, "200", "600"))
            Call txtic_LostFocus(4)
            tabCad.Tab = 0
            cmd(0).Enabled = False
            cmd(1).Enabled = False
        End If
    End If
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        fra(9).Visible = False
    End If
    DoEvents
    Consultando = False
End Sub

Private Sub Form_Load()
    
    Dim Controle As Control
    Dim i As Byte
    Dim Rs As VSRecordset
    Set cadastro = New VSImposto
    
    Call Edita.AtualizaCombo(Bdados, cboTipoLogrContribBc, "Select ttl_nome From Tab_Tipo_Logr")
    Call AtualizaUF(cboUFBc)
    txtic(2).MaxLength = Nvl(Temp.PegaParametro(Bdados, "CAMPO QUADRA"), 4)
    For Each Controle In Controls
        If IsNumeric(Controle.Tag) Then
            If Val(Controle.Tag) < 20 Then Call Edita.AtualizaCombo(Bdados, Controle, "Select convert(varchar,tco_cod_componente) " & Bdados.Concatena & "'-'" & Bdados.Concatena & " tco_descricao_componente From Tab_Componente_Avancado Where tco_grupo = " & Controle.Tag & " order by tco_cod_componente asc")
        End If
    Next
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    NovoContrib = True
    Bdados.FechaTabela Rs
    cboLoteamento.Preencher Bdados, "Select TLO_COD_LOTEAMENTO,TLO_DESCRICAO from TAB_LOTEAMENTO ORDER BY TLO_DESCRICAO", 1
    cboEdificio.Preencher Bdados, "Select TED_COD_EDIFICIO,TED_DESCRICAO from TAB_EDIFICIO ORDER BY TED_DESCRICAO", 1

    Boletim = tbo_Territorial
End Sub

Private Sub lstCond_DblClick()
    If lstCond.SelectedItem Is Nothing Then Exit Sub
    Dim ItmX As Object
    Dim i As Byte
    
    On Error Resume Next
    txtic(9) = lstCond.SelectedItem
    txtInscAnteriorBC = lstCond.SelectedItem.SubItems(1)
    cboTipoImovelBc.ListIndex = lstCond.SelectedItem.SubItems(2) - 1
    txtNumeroBc = lstCond.SelectedItem.SubItems(3)
    txtComplementoBc = lstCond.SelectedItem.SubItems(4)
    txtIMBc = lstCond.SelectedItem.SubItems(5)
    txtNomeContribBc = lstCond.SelectedItem.SubItems(6)
    txtCpfCgcBc = lstCond.SelectedItem.SubItems(7)
    cboTipoLogrContribBc = lstCond.SelectedItem.SubItems(8)
    txtNomeLogrContribBc = lstCond.SelectedItem.SubItems(9)
    txtNumeroBc = lstCond.SelectedItem.SubItems(10)
    txtComplementoBc = lstCond.SelectedItem.SubItems(11)
    txtBairroContribBc = lstCond.SelectedItem.SubItems(12)
    txtCepImBc = lstCond.SelectedItem.SubItems(13)
    txtMunicBc = lstCond.SelectedItem.SubItems(14)
    cboUFBc = lstCond.SelectedItem.SubItems(15)
    txtOcupanteBc = lstCond.SelectedItem.SubItems(16)
    txtCpfOcupanteBc = lstCond.SelectedItem.SubItems(17)
    txtCodComponente(20) = lstCond.SelectedItem.SubItems(18)
    lstCond.ListItems.Remove lstCond.SelectedItem.Index
    DoEvents
End Sub


Private Sub lstEdific_Click()
    Dim i As Byte
    Dim sql As String
    Dim Rs As VSRecordset
    On Error Resume Next
    If lstEdific.SelectedItem Is Nothing Then Exit Sub
    If Trim(txtCodComponente(13)) <> "" Then
        If Not Confirma("Existe uma unidade edificada em aberto. Deseja exclui-la?") Then
            Exit Sub
        End If
    End If
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
    txtFracaoEdif = lstEdific.SelectedItem.SubItems(13)
    
    'ElseIf CInt(Nvl(txtInscImobiliaria, 0)) >= 200 Then
        'CONSULTA BC
    If lstEdific.SelectedItem >= 200 Then
        'tabCad.TabEnabled(2) = True
        txtic(5) = txtic(0)
        txtic(6) = txtic(1)
        txtic(7) = txtic(2)
        txtic(8) = txtic(3)
        txtCodLogrBc = txtCodLogr
        txtLogr = txtTipoLogrBt
        txtNomeLogr = txtLogrBt
        txtNumeroBc = txtNumero
        txtLoteamentoBc = cboLoteamento.Coluna(0).Valor
        txtQuadraBc = txtQuadra
        txtLoteBc = txtLote
        txtBairro = txtBairroBt
        txtCepImBc = txtCEP
        txtic(9) = lstEdific.SelectedItem
        txtic(9).Enabled = True
        
        sql = "SELECT TIM_complemento,tim_tci_im,tim_ocupante,tim_cgc_cpf_ocupante,tim_tipo_imovel from tab_imovel where tim_ic ='" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & "' AND TIM_UNIDADE=" & lstEdific.SelectedItem
        If Bdados.AbreTabela(sql, Rs) Then
            txtComplementoBc = "" & Rs(0)
            txtIMBc = "" & Rs(1)
            txtIMBc_LostFocus
            txtOcupanteBc = "" & Rs(2)
            txtCpfOcupanteBc = "" & Rs(3)
            cboTipoImovelBc.ListIndex = Rs(4) - 1
            DoEvents
            sql = "select tdi_tco_cod_componente from tab_detalhe_imovel where tdi_tgc_cod_grupo = 3 and tdi_tim_ic_unidade = " & txtic(9) & " and tdi_tim_ic ='" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & "'"
            If Bdados.AbreTabela(sql, Rs) Then
                txtCodComponente(20) = Rs(0)
            End If
            Bdados.FechaTabela Rs
        End If
    End If
    If Me.Tag = "" Then
        lstEdific.ListItems.Remove lstEdific.SelectedItem.Index
    End If
    If Trim(txtFracaoEdif) = "" Then txtFracaoEdif = "1,00"
    'FIM CONSULTA BC
End Sub

Private Sub lstPesq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaGrid lstPesq, ColumnHeader
End Sub

Private Sub lstPesq_DblClick()
    On Error Resume Next
    txtIM = lstPesq.SelectedItem
    Call txtIm_LostFocus
End Sub

Private Sub tabCad_Click(PreviousTab As Integer)
    'NOVIDADE
    If tabCad.Tab = 2 Then
        If Trim(txtInscImobiliaria) = "" And Trim(txtic(4)) <> "" Then
            txtInscImobiliaria.Enabled = True
            txtic(10) = txtic(0)
            txtic(11) = txtic(1)
            txtic(12) = txtic(2)
            txtic(13) = txtic(3)
            If txtInscImobiliaria.Enabled Then txtInscImobiliaria.SetFocus
        End If
    ElseIf tabCad.Tab = 3 Then
        If Trim(txtic(4)) <> "" Then
            txtic(5) = txtic(0)
            txtic(6) = txtic(1)
            txtic(7) = txtic(2)
            txtic(8) = txtic(3)
            txtic(9).Enabled = True
            
            cboTipoImovelBc.ListIndex = cboTipoImovel.ListIndex
            txtCodLogrBc = txtCodLogr
            txtLogr = txtTipoLogrBt
            txtNomeLogr = txtLogrBt
            txtNumeroBc = txtNumero
            txtLoteamentoBc = cboLoteamento.Coluna(0).Valor
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

Private Sub txtAreaLote_LostFocus()
    tabCad.Tab = 2
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") <> "REDUZIDA" Then txtInscImobiliaria.SetFocus
End Sub

Private Sub txtBairroContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCep_LostFocus()
    txtCEP = Format(txtCEP, "00000-000")
End Sub

Private Sub txtCepImBc_LostFocus()
    txtCepImBc = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
End Sub

Private Sub txtCodBairro_LostFocus()
    Dim Rs As VSRecordset
    Dim sql As String
    If Trim(txtCodBairro) <> "" Then
        sql = " select TBA_NOME from TAB_BAIRRO where tba_cod_bairro=" & txtCodBairro & " and tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
        If Bdados.AbreTabela(sql, Rs) Then
            txtBairroBt = Rs(0)
        Else
            Avisa "Bairro inexistente."
            txtCodBairro.SetFocus
        End If
        Bdados.FechaTabela Rs
    Else
        txtBairroBt = ""
    End If
    
End Sub

Private Sub txtCodComponente_Change(Index As Integer)
    Dim Controle As Control
    On Error GoTo trata
     If Index = 20 Then
        'cboCobrancaBc.ListIndex = Nvl(txtCodComponente(Index).Text, 0) - 1
        cboCobrancaBc.ListIndex = ListIndexDe(cboCobrancaBc, txtCodComponente(Index).Text & "-" & DescricaoComponente(cboCobrancaBc.Tag, txtCodComponente(Index).Text))
        Exit Sub
    End If
    For Each Controle In Controls
        If Controle.Tag = Index + 1 Then
            'Controle.ListIndex = Util.Nvl(txtCodComponente(Index).Text, 0) - 1
            Controle.ListIndex = ListIndexDe(Controle, txtCodComponente(Index).Text & "-" & DescricaoComponente(Controle.Tag, txtCodComponente(Index).Text))
            Exit For
        End If
    Next
trata:
    If Err.Number = 380 Then
'        txtCodComponente(Index).SetFocus
    End If
End Sub

Private Sub txtCodComponente_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodLogr_LostFocus()
    Dim Query As String
    Dim Rs As VSRecordset
    If Trim(txtCodLogr) = "" Then Exit Sub
    Query = "SELECT TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome, " & _
        " TAB_BAIRRO.TBA_NOME FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
        " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
         " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & txtCodLogr & "' and tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio & " and tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
    If Bdados.AbreTabela(Query, Rs) Then
        txtTipoLogrBt = Rs(0)
        txtLogrBt = Rs(1)
    Else
        Avisa "Código de logradouro inválido."
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtCodLogrBc_LostFocus()
    Dim sql As String
    Dim Rs As VSRecordset
    
    sql = "Select "
End Sub

Private Sub txtCodLogrContrib_LostFocus()
    Dim Query As String
    Dim Rs As VSRecordset
    If Trim(txtCodLogrContrib) <> "" Then
        If Trim(txtCodLogrContrib) = "" Then Exit Sub
        Query = "SELECT TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome, " & _
            " TAB_BAIRRO.TBA_NOME,tlg_cep FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
            " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
             " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & txtCodLogrContrib & "' and tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
        If Bdados.AbreTabela(Query, Rs) Then
            txtNomeTipoLogrContrib = Rs(0)
            txtNomeLogrContrib = Rs(1)
            txtBairroContrib = Rs(2)
            txtMunic = Aplicacoes.Municipio
            txtCEP = Rs.Fields("tlg_cep")
        Else
            Avisa "Código de logradouro inválido."
            txtCodLogr.SetFocus
        End If
        Bdados.FechaTabela Rs
        If txtCEP = "" Then
            txtCEP = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
        End If
        txtUF = Temp.PegaParametro(Bdados, "ESTADO CLIENTE")
        
    End If
End Sub

Private Sub txtCodMens_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtCodReduzido_LostFocus()
    txtic_LostFocus 4
End Sub


Private Sub txtCompContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFracao_Change()

End Sub

Private Sub txtCpfCgc_LostFocus()

    If txtCpfCgc = "99999999999" Or txtCpfCgc = "999.999.999-99" Or txtCpfCgc = "00000000000" Or txtCpfCgc = "000.000.000-00" Then
        Util.Avisa "CPF inválido."
        txtCpfCgc.SetFocus
    End If
    txtCpfCgc = Edita.TiraPic(txtCpfCgc, ".")
    txtCpfCgc = Edita.TiraPic(txtCpfCgc, "-")
    txtCpfCgc = Edita.TiraPic(txtCpfCgc, "/")
    If Len(Trim(txtCpfCgc)) = 11 Then
        If Not Util.ValidaCpf(Trim(txtCpfCgc)) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCpfCgc.SetFocus
            Exit Sub
        End If
        txtCpfCgc = Edita.FormataTexto(txtCpfCgc, Cpf)
    ElseIf Len(Trim(txtCpfCgc)) = 14 And Mid(Trim(txtCpfCgc), 4, 1) <> "." Then
        txtCpfCgc.MaxLength = 20
        txtCpfCgc = Trim(txtCpfCgc)
        txtCpfCgc = Edita.FormataTexto(txtCpfCgc, Cgc)
    ElseIf Trim(txtCpfCgc) <> "" And Len(Trim(txtCpfCgc)) <> 18 And Mid(Trim(txtCpfCgc), 4, 1) <> "." Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
        txtCpfCgc.SetFocus
    End If
End Sub

Private Sub txtCpfCgcBc_LostFocus()
    If Len(txtCpfCgcBc) = 11 Then
        If Not Util.ValidaCpf(txtCpfCgcBc) Then
            Call Util.Informa("Número de CPF inválido.")
             txtCpfCgcBc.SetFocus
            Exit Sub
        End If
        txtCpfCgcBc = Edita.FormataTexto(txtCpfCgcBc, Cpf)
    ElseIf Len(txtCpfCgcBc) = 14 And Mid(txtCpfCgcBc, 4, 1) <> "." Then
        txtCpfCgcBc.MaxLength = 20
        txtCpfCgcBc = Edita.FormataTexto(txtCpfCgcBc, Cgc)
    ElseIf Trim(txtCpfCgcBc) <> "" And Len(txtCpfCgcBc) <> 18 And Mid(txtCpfCgcBc, 4, 1) <> "." Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
        txtCpfCgcBc.SetFocus
    End If
End Sub

Private Sub txtCpfOcupante_LostFocus()
    If Len(txtCPFOcupante) = 11 Then
        If Not Util.ValidaCpf(txtCPFOcupante) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCPFOcupante.SetFocus
            Exit Sub
        End If
        txtCPFOcupante = Edita.FormataTexto(txtCPFOcupante, Cpf)
    End If
    tabCad.Tab = 1
    DoEvents
    txtCodComponente(0).SetFocus
End Sub

Private Sub txtCpfOcupanteBc_LostFocus()
    If Len(txtCpfOcupanteBc) = 11 Then
        If Not Util.ValidaCpf(txtCpfOcupanteBc) Then
            Call Util.Informa("Número de CPF inválido.")
            txtCpfOcupanteBc.SetFocus
            Exit Sub
        End If
        txtCpfOcupanteBc = Edita.FormataTexto(txtCpfOcupanteBc, Cpf)
    End If
End Sub

Private Sub txtDataAforamento_LostFocus()
    txtDataAforamento = Edita.FormataTexto(txtDataAforamento, Data)
End Sub

Private Sub txtDtRegistro_LostFocus()
    txtDtRegistro = Edita.FormataTexto(txtDtRegistro, Data)
End Sub

Private Sub txtFracaoEdif_LostFocus()
    If CInt(Nvl(Trim(txtic(4)), 0)) >= 200 Then
        tabCad.Tab = 3
        DoEvents
        txtic(9).SetFocus
    End If
End Sub

Private Sub txtIc_Change(Index As Integer)
    If Len(txtic(Index)) = txtic(Index).MaxLength Then
       SendKeys "{ENTER}"
    End If
End Sub

Private Sub txtic_LostFocus(Index As Integer)
    Dim sql As String
    Dim Rs As VSRecordset
    Dim Tem As String
    Dim Temporario As String
    
     If Index = 9 Then
        
        If Trim(txtic(9)) = "" Then Exit Sub
        Screen.MousePointer = 11
        
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            sql = "Select * from tab_imovel where tim_ic_auxiliar ='" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & "'"
        Else
            sql = "Select * from tab_imovel where (tIM_ic ='" & _
            txtic(0) & txtic(1) & txtic(2) & txtic(3) & "' AND TIM_UNIDADE =" & txtic(9) & ")" & Temporario
        End If
        If Bdados.AbreTabela(sql, Rs) Then
            txtIMBc = Rs!tim_tci_im
            txtIMBc_LostFocus
            cboTipoImovelBc.ListIndex = Rs!tim_tipo_imovel - 1
            txtInscAnteriorBC = "" & Rs!tim_ic_anterior
            txtOcupanteBc = "" & Rs!tim_ocupante
            txtCpfCgcBc = "" & Rs!tim_cgc_cpf_ocupante
            sql = "Select * from tab_detalhe_imovel where TDI_TIM_IC ='" & _
            txtic(0) & txtic(1) & txtic(2) & txtic(3) & "' AND tdi_tim_ic_unidade = " & CInt(Nvl(txtic(9), 0)) & " and tdi_tgc_cod_grupo = 3" ' order by tdi_tco_cod_componente asc"
            If Bdados.AbreTabela(sql, Rs) Then
                txtCodComponente(20) = Rs!TDI_VALOR_ITEM
            End If
        End If
        Bdados.FechaTabela Rs
    ElseIf Index = 4 Then
        If Trim(txtic(4)) = "" And (txtCodReduzido) = "" Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        
            If txtic(4).Enabled = False Or txtCodReduzido.Enabled = False Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        txtCodReduzido = Trim(txtCodReduzido)
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            If txtCodReduzido = "" Then
                sql = "Select * from tab_imovel where tim_ic_auxiliar ='" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & txtic(4) & "' " & Temporario
            Else
                sql = "Select * from tab_imovel where tim_ic ='" & txtCodReduzido & "'"
            End If
            
        Else
            sql = "Select * from tab_imovel where (tIM_ic ='" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & txtic(4) & "'" & IIf(CInt(txtic(4)) >= 200, " AND TIM_UNIDADE =" & txtic(4), "") & ") " & Temp
        End If
        
        If Bdados.AbreTabela(sql, Rs) Then
            txtICAnterior = "" & IIf(Rs!tim_ic_anterior = 0, "", Rs!tim_ic_anterior)
            cboTipoImovel.ListIndex = Rs!tim_tipo_imovel - 1
            
            txtCodLogr = "" & Rs!tim_tlg_cod_logradouro
            txtCodLogr_LostFocus
            txtNumero = "" & Rs!tim_numero
            txtComplemento = "" & Rs!tim_complemento
            cboLoteamento.SetarLinha "" & Rs!tim_loteamento, 0
            txtQuadra = "" & Rs!tim_QUADRA
            txtLote = "" & Rs!tim_lote
            txtOcupante = "" & Rs!tim_ocupante
            txtCPFOcupante = "" & Rs!tim_cgc_cpf_ocupante
            
            txtICAnterior = "" & Rs!tim_ic_anterior
            txtCodMens = "" & Rs!tim_COD_MENSAGEM
            txtNumAforamento = "" & Rs!tim_AFORAMENTO_NUMERO
            txtFichaAforamento = "" & Rs!tim_AFORAMENTO_FICHA
            txtLivroAforamento = "" & Rs!tim_AFORAMENTO_LIVRO
            txtFolhaAforamento = "" & Rs!tim_AFORAMENTO_FOLHA
            txtDataAforamento = "" & Rs!tim_AFORAMENTO_DATA
            txtRegistro = "" & Rs!tim_AFORAMENTO_REGISTRO
            txtDtRegistro = "" & Rs!tim_DATA_REGISTRO
            txtSecao = "" & Rs!tim_secao
            
            If txtCodReduzido <> "" Then
                txtic(0) = "" & Left(Trim(Rs.Fields("tim_ic_auxiliar")), 2)
                txtic(1) = "" & Mid(Trim(Rs.Fields("tim_ic_auxiliar")), 3, 2)
                txtic(2) = "" & Mid(Trim(Rs.Fields("tim_ic_auxiliar")), 5, 3)
                txtic(3) = "" & Mid(Trim(Rs.Fields("tim_ic_auxiliar")), 8, 4)
                txtic(4) = "" & Right(Trim(Rs.Fields("tim_ic_auxiliar")), 3)
            End If
            'VOU PEGAR O CONTRIBUINTE
            txtIM = "" & Rs!tim_tci_im
            txtCodReduzido = Rs!TIM_IC
            txtIm_LostFocus
            txtZona = "" & Rs!tim_ZONA
            txtCodBairro = "" & Rs!tim_TBA_COD_BAIRRO
            
            txtCodBairro_LostFocus
            'VOU PEGAR OS DETALHES
            'Temporario = " or (TDI_TIM_IC ='" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & "' AND tdi_tim_ic_unidade =" & CInt(txtic(4)) & ")"
            If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                sql = "Select * from TAB_DETALHE_IMOVEL where TDI_TIM_IC ='" & txtCodReduzido & "'order by tdi_tco_cod_componente asc"
            Else
                sql = "Select * from TAB_DETALHE_IMOVEL where (TDI_TIM_IC ='" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & txtic(4) & "' and (tdi_tgc_cod_grupo < 9 or tdi_tgc_cod_grupo >=100))" ' AND tdi_tim_ic_unidade = " & CInt(txtic(4)) & ") " & Temporario '& " order by tdi_tco_cod_componente asc"
            End If
            
            If Bdados.AbreTabela(sql, Rs) Then
                Rs.MoveFirst
                Do While Not Rs.EOF
                    If Rs!tdi_tgc_cod_grupo <= 8 Then
                        On Error Resume Next
                        'txtCodComponente(Rs!tdi_tgc_cod_grupo - 1) = Rs!TDI_VALOR_ITEM + 1
                        'txtCodComponente(rs!tdi_tgc_cod_grupo - 1) = rs!TDI_VALOR_ITEM
                        txtCodComponente(Rs!tdi_tgc_cod_grupo - 1) = Rs!tdi_tco_cod_componente
                        On Error GoTo 0
                    Else
                        Dim Controle As Control
                        On Error Resume Next
                        For Each Controle In Controls
                            If IsNumeric(Controle.Tag) Then
                                If CInt(Controle.Tag) = Rs!tdi_tgc_cod_grupo Then
                                    Controle.Text = Rs!TDI_VALOR_ITEM
                                    Exit For
                                End If
                            End If
                        Next
                        On Error GoTo 0
                    End If
                    Rs.MoveNext
                Loop
            End If
            'Vou pegar as construcões
            Dim i As Byte
            Dim ItmX As Object
            If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                Temporario = "  TDI_TIM_IC ='" & txtCodReduzido & "' and tdi_tim_ic_unidade > 0"
            Else
                If CInt(txtic(4)) <> 200 Then
                    Temporario = "  TDI_TIM_IC ='" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & txtic(4) & "' and tdi_tim_ic_unidade > 0"
                Else
                    Temporario = " ( TDI_TIM_IC >'" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & txtic(4) & "' AND TDI_TIM_IC <'" & txtic(0) & txtic(1) & txtic(2) & txtic(3) & "300'" & ")"
                End If
            End If
            sql = "Select * from tab_detalhe_imovel where " & Temporario & " order by tdi_tim_ic_unidade asc, tdi_tgc_cod_grupo asc"
            lstEdific.ListItems.Clear
            If Bdados.AbreTabela(sql, Rs) Then
                Rs.MoveFirst
                Set ItmX = lstEdific.ListItems.Add(, , Format(Rs!tdi_tim_ic_unidade, "000"))
                Dim Conta As Byte
                Conta = 1
                Do While Not Rs.EOF
                    If Format(Nvl(Rs!tdi_tim_ic_unidade, 0), "000") <> ItmX Then
                        Set ItmX = lstEdific.ListItems.Add(, , Format(Rs!tdi_tim_ic_unidade, "000"))
                        Conta = Conta + 1
                    End If
                    'If Rs!tdi_tgc_cod_grupo >= 14 And Rs!tdi_tgc_cod_grupo <= 15 Then
                    For i = 2 To 14
                        If CInt(lstEdific.ColumnHeaders(i).Tag) = CInt(Rs!tdi_tgc_cod_grupo) Then
                            ItmX.SubItems(i - 1) = Rs!TDI_VALOR_ITEM
                            Screen.MousePointer = 0
                            Exit For
                        End If
                    Next
                    Rs.MoveNext
                Loop
                txtic(10) = txtic(0)
                txtic(11) = txtic(1)
                txtic(12) = txtic(2)
                txtic(13) = txtic(3)
            End If
            If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") <> "REDUZIDA" Then
                For i = 0 To 4
                    txtic(i).Enabled = False
                Next
            Else
                txtCodReduzido.Enabled = False
            End If
            txtAnoConst = ""
            txtAreaEdif = ""
            txtAreaEdifTotal = ""
            txtFracaoEdif = ""
            txtPavimento = ""
            'Vou pegar os condominios
            If Trim(txtic(4)) <> "" Then
                If CInt(Trim(txtic(4))) < 200 Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            Dim Campos As String
            Dim Dados As String
            Dim Dados_Extras As String
            Dim Base As String
            
            Campos = "TIM_UNIDADE, tim_ic_anterior,tim_tipo_imovel,tim_numero,  tim_complemento, " _
                    & "tim_tci_im , " _
                    & "  tci_nome,tci_cgc_cpf,tci_logradouro," _
                    & " tci_nome_logradouro, tci_numero," _
                    & "tci_complemento, tci_bairro,tci_cep,tci_cidade,tci_UF," _
                    & "tim_ocupante,tim_cgc_cpf_ocupante,tdi_valor_item"
            
            'Sql = "Select " & Campos & " from tab_imovel,tab_contribuinte " & _
                " where (TIM_IC >'" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & txtIc(4) & _
                "' AND TIM_IC <'" & txtIc(0) & txtIc(1) & txtIc(2) & txtIc(3) & "300'" & ")" & _
                " and tim_tci_im = tci_im  order by tim_unidade asc"
            If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                Dados = txtCodReduzido
                Base = 1
            Else
                Base = ""
                Dados = txtic(0) & txtic(1) & txtic(2) & txtic(3) & txtic(4)
                Dados_Extras = txtic(0) & txtic(1) & txtic(2) & txtic(3) & "300'"
            End If
            sql = "Select " & Campos & " from tab_imovel,tab_contribuinte,tab_detalhe_imovel " & _
                " where (TIM_IC >'" & Dados & _
                " ' AND TIM_IC <'" & IIf(Base <> "", Dados, Dados_Extras) & "')" & _
                " and tim_tci_im = tci_im  and tim_ic = tdi_tim_ic and tdi_tgc_cod_grupo = 3 order by tim_unidade asc"
            lstCond.ListItems.Clear
            If Bdados.AbreTabela(sql, Rs) Then
                Rs.MoveFirst
                Conta = 1
                Do While Not Rs.EOF
                    Set ItmX = lstCond.ListItems.Add(, , Format(Rs!tim_unidade, "000"))
                    For i = 1 To 18
                        On Error Resume Next
                        ItmX.SubItems(i) = CStr("" & Rs(CInt(i)))
                    Next
                    Rs.MoveNext
                Loop
                txtic(10) = txtic(0)
                txtic(11) = txtic(1)
                txtic(12) = txtic(2)
                txtic(13) = txtic(3)
            End If
            
            If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") <> "REDUZIDA" Then
                For i = 0 To 4
                    txtic(i).Enabled = False
                Next
            Else
                txtCodReduzido.Enabled = False
            End If
          
        Else
            Avisa "Imóvel não cadastrado."
            DoEvents
            ' LIMPA A TELA E DEVOLVE O FOCO PARA O CAMPO IC
            cmd_Click 0
            
        End If
        Bdados.FechaTabela Rs
    End If
    Screen.MousePointer = 0
    NovoContrib = False
End Sub

Private Sub txtIcAnterior_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCpfCgc_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    Dim Rs As VSRecordset
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIM) <> "" Then
        If InStr(txtIM, " - ") <> 0 Then
            txtIM = cadastro.FormataInscricao(txtIM, InscContrib)
        End If
        sql = "Select tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, " & _
        " tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF,TCI_COD_LOGRADOURO from Tab_Contribuinte where tci_im = '" & txtIM & "'"
        If Bdados.AbreTabela(sql, Rs) Then
            txtNomeContrib = "" & Rs(0) 'Rs!tci_Nome
            txtNomeTipoLogrContrib = "" & Rs(1)
            txtNomeLogrContrib = "" & Rs(2) '!tci_nome_logradouro
            txtNumeroContrib = "" & Rs(3) '!tci_numero
            txtCompContrib = "" & Rs(4) '!tci_complemento
            txtBairroContrib = "" & Rs(5) '!tci_bairro
            txtCEP = "" & Rs(6) '!tci_cep
            txtMunic = "" & Rs(7)
            txtUF = "" & Rs(8) '!tci_UF
            txtCpfCgc = "" & Rs(9)
            txtCodLogrContrib = "" & Rs.Fields("TCI_COD_LOGRADOURO")
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtIM.Enabled = True
            txtIM.SetFocus
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtIMBc_LostFocus()
    Dim Rs As VSRecordset
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIMBc) <> "" Then
        txtIMBc = cadastro.FormataInscricao(txtIMBc, InscContrib)
        sql = "Select tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF from Tab_Contribuinte where tci_im = '" & txtIMBc & "'"
        If Bdados.AbreTabela(sql, Rs) Then
            txtNomeContribBc = "" & Rs(0)  'Rs!tci_Nome
            cboTipoLogrContribBc.ListIndex = cadastro.BuscaCodLogr(Rs(1)) - 1
            txtNomeLogrContribBc = "" & Rs(2)  '!tci_nome_logradouro
            txtNumeroContribBc = "" & Rs(3)  '!tci_numero
            txtCompContribBc = "" & Rs(4)  '!tci_complemento
            txtBairroContribBc = "" & Rs(5)  '!tci_bairro
            txtCepBc = "" & Rs(6)  '!tci_cep
            txtMunicBc = "" & Rs(7)
            cboUFBc = "" & Rs(8)  '!tci_UF
            txtCpfCgcBc = "" & Rs(9)
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
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.Maiuscula(KeyAscii)
End Sub

Private Sub txtMunic_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMunic_LostFocus()
    If Trim(txtMunic) = "" Then txtMunic = Aplicacoes.Municipio
End Sub

Private Sub txtNomeContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNomeLogrContrib_KeyPress(KeyAscii As Integer)
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

Private Sub txtQuadra_KeyPress(KeyAscii As Integer)
'    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtSecao_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.Maiuscula(KeyAscii)
End Sub

Private Sub txtServCalc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtServIlum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtServLimp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtTtotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtUnids_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtZona_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
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
    Dim InscReduzida As String
    Dim CodLogr As Long
    Dim DtVenc As String
    Dim SitCadastral As String
    Static Unidades As Integer
    Dim i As Integer
    Dim j As Integer
    Dim cadastro As New VSImposto
    Dim Conta As New ContaCorrente
    
    Select Case cmd(Index).Caption
        Case "&Salvar"
            If Not Util.Confirma("Deseja salvar as alterações?", "Alteração de Cadastro") Then Exit Sub
                If Trim(txtCodBairro) = "" Then
                    Util.Informa "Falta a definição do bairro."
                    txtCodBairro.SetFocus
                    tabCad.Tab = 0
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                
                'VERIFCANDO BP'S
                If cboTipoImovel = "PREDIAL" Then
                    If Not Lote.VerificaDigitacaoBP(lstEdific, txtCodBairro, tabCad) Then
                        txtCodComponente(13).SetFocus
                        Exit Sub
                    End If
                End If
                
                txtFatorFixo.Tag = "1000"
                CodLogr = txtCodLogr
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                    InscReduzida = Trim(txtCodReduzido)
                    InscricaoCadastral = txtic(0) & txtic(1) & txtic(2) & txtic(3) & txtic(4)
                Else
                    InscricaoCadastral = txtic(0) & txtic(1) & txtic(2) & txtic(3) & txtic(4)
                End If
                InscricaoMunicipal = txtIM
                Screen.MousePointer = 11
                If Not Lote.VerificaFechamentoAreas(lstEdific) Then: Screen.MousePointer = 0: Exit Sub
                'GRAVANDO BT
           
                Lote.CarregaDadosContribuinte InscricaoMunicipal, txtNomeContrib, txtCpfCgc, "", txtNomeTipoLogrContrib, txtNomeLogrContrib, _
                        txtNumeroContrib, txtCompContrib, "", txtBairroContrib, txtCEP, txtMunic, txtUF
                If Not Lote.InsereContribuinte(NovoContrib) Then Exit Sub
                'SQz (BLS, 15/04/2003): Ao apagar perde-se a informação do valor venal (!)
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "PETROLINA" Then
                    Lote.ApagaImovel InscReduzida
                Else
                    Lote.ApagaImovel InscricaoCadastral
                End If
                
                Lote.CarregaDadosImovel InscricaoCadastral, txtICAnterior, txtic(4), "0", "", "", CStr(CodLogr), txtCodBairro, _
                     txtNumero, txtComplemento, txtLote, txtQuadra, CStr(cboLoteamento.Coluna(0).Valor), Boletim, txtOcupante, txtCPFOcupante, _
                     txtCodMens, txtZona, txtNumAforamento, txtFichaAforamento, txtLivroAforamento, txtFolhaAforamento, txtRegistro, txtDataAforamento, txtDtRegistro, , InscReduzida, txtSecao, CStr(cboEdificio.Coluna(0).Valor)
                
                If Not Lote.InsereTerritorio() Then Exit Sub
                'alteração do contribuinte
                
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                    Lote.ApagaDetalheImovel InscReduzida
                     Call Lote.GravaComponentes(InscReduzida, Me, 1, 8, False, 0, 0)
                     Call Lote.GravaComponentes(InscReduzida, Me, 100, 109, True, 0, 0)
                Else
                    Lote.ApagaDetalheImovel InscricaoCadastral
                    Call Lote.GravaComponentes(InscricaoCadastral, Me, 1, 8, False, Nvl(txtic(4), 0), 0)
                    Call Lote.GravaComponentes(InscricaoCadastral, Me, 100, 109, True, Nvl(txtic(4), 0), 0)
                End If
                
                'GRAVANDO BP
                If txtCodMens = 98 Or txtCodMens = 99 Then Util.Informa "Código de Mensagem " & txtCodMens & ". BP não será gravado."
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                    Lote.GravaBP lstEdific, txtCodMens, InscReduzida, 0
                Else
                    Lote.GravaBP lstEdific, txtCodMens, txtic(0) & txtic(1) & txtic(2) & txtic(3), txtic(4)
                End If
                'GRAVANDO BC'S
                If CInt(Nvl(Trim(txtic(4)), 0)) >= 200 Then
                    cboCobrancaBc.Tag = "3"
                    If lstCond.ListItems.Count > 0 Then
                        For j = 1 To lstCond.ListItems.Count 'Para cada edificacao
                            lstCond.ListItems(j).Selected = True
                            InscricaoMunicipal = txtIMBc
                            CodLogr = txtCodLogr
                            InscricaoCadastral = txtic(0) & txtic(1) & txtic(2) & txtic(3) & lstCond.SelectedItem
                            'INSERE CONTRIBUINTE
                            If lstCond.SelectedItem.ListSubItems(5) = "" Then
                                InscricaoMunicipal = cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
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
                                    InscricaoCadastral & txtic(4), "", txtCodLogrBc, txtCodBairro, _
                                    lstCond.SelectedItem.ListSubItems(3), lstCond.SelectedItem.ListSubItems(4), _
                                    Trim(txtLoteBc), Trim(txtQuadraBc), Trim(txtLoteamentoBc), lstCond.SelectedItem.ListSubItems(2), _
                                    lstCond.SelectedItem.ListSubItems(16), lstCond.SelectedItem.ListSubItems(17), _
                                    Nvl(txtCodMens, 0), Nvl(txtZona, 1), , , , , , , , , InscReduzida, txtSecao, CStr(cboEdificio.Coluna(0).Valor)
                            Lote.InsereTerritorio
'alteração do contribuinte
                            'INSERE COD. COBRANÇA
                            Call Lote.GravaComponente(InscricaoCadastral, lstCond.SelectedItem, lstCond.SelectedItem.ListSubItems(18), 3, 0)
                        Next
                    End If
                End If
                atualizarEnderecoContribuinte InscricaoCadastral, txtTipoLogrBt, txtLogrBt, txtNumero, txtComplemento, txtBairroBt
                atualizarContribuinte InscricaoCadastral, InscricaoMunicipal
                'LIMPA TELA
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") <> "REDUZIDA" Then
                    Util.Informa "Imóvel " & InscricaoCadastral & " gravado com sucesso."
                Else
                    Util.Informa "Imóvel " & InscReduzida & " gravado com sucesso."
                End If
                Call cmd_Click(0)
                DoEvents
        Case "&Limpar"
            NovoContrib = True
            Call Edita.LimpaCampos(Me)
            cboCobrancaBc.Tag = ""
            cboCobranca.Tag = "3"
            lstEdific.ListItems.Clear
            lstCond.ListItems.Clear
            tabCad.Tab = 0
            Unidades = 0
            Screen.MousePointer = 0
            If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") <> "REDUZIDA" Then
                For i = 0 To 4
                    txtic(i).Enabled = True
                Next
            Else
                txtCodReduzido.Enabled = True
            End If
            txtic(0).SetFocus
        Case "Sai&r"
            NovoContrib = True
            Unload Me
    End Select
End Sub


Private Function atualizarEnderecoContribuinte(Ic As String, Tipo As String, Logradouro As String, Numero As String, Complemento As String, Bairro As String) As Boolean
    Dim sql As String
    
    sql = "UPDATE TAB_CONTRIBUINTE " & _
            " SET tci_logradouro='" & Tipo & "', " & _
                " tci_nome_logradouro='" & Logradouro & "', " & _
                " tci_numero='" & Numero & "', " & _
                " tci_complemento='" & Complemento & "', " & _
                " tci_bairro='" & Bairro & "'" & _
            " WHERE tci_tim_ic='" & Ic & "'"
    atualizarEnderecoContribuinte = Bdados.Executa(sql)
End Function

Private Function atualizarContribuinte(Ic As String, Im As String) As Boolean
    Dim sql As String
    
    sql = "UPDATE TAB_GERACAO_TRIBUTO" & _
            " SET tgt_im='" & Im & "'" & _
            " WHERE tgt_tim_ic='" & Ic & "'"
    atualizarContribuinte = Bdados.Executa(sql)
    
    sql = "UPDATE TAB_DARM_RECEBIDO" & _
            " SET tdr_im='" & Im & "'" & _
            " WHERE tdr_tim_ic='" & Ic & "'"
    atualizarContribuinte = atualizarContribuinte And Bdados.Executa(sql)
End Function
