VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form dec001 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SIAT - Declaração Eletrônica: Cadastro"
   ClientHeight    =   6585
   ClientLeft      =   870
   ClientTop       =   1575
   ClientWidth     =   9600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   HelpContextID   =   1
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6585
   ScaleMode       =   0  'User
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Begin VB.Data data_contrib 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   8580
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5430
      Visible         =   0   'False
      Width           =   1140
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6555
      HelpContextID   =   2
      Left            =   30
      TabIndex        =   31
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11562
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      ForeColor       =   16711680
      TabCaption(0)   =   "&Contribuintes"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame_contador"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frame_contrib"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Co&ntadores"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame frame_contrib 
         Height          =   6075
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   9435
         Begin VB.Frame Frame12 
            ForeColor       =   &H00FF0000&
            Height          =   1155
            Left            =   350
            TabIndex        =   105
            Top             =   3780
            Width           =   8985
            Begin VB.ComboBox cmb_localizacao 
               Height          =   315
               ItemData        =   "DEC001.frx":0000
               Left            =   4935
               List            =   "DEC001.frx":0002
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   330
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VB.ComboBox cmb_tipo 
               Height          =   315
               ItemData        =   "DEC001.frx":0004
               Left            =   6540
               List            =   "DEC001.frx":0006
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   330
               Width           =   1545
            End
            Begin VB.TextBox label1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   330
               Width           =   4335
            End
            Begin MSMask.MaskEdBox Mask_dt_inicio_op 
               Height          =   315
               Left            =   3390
               TabIndex        =   20
               Top             =   750
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "@@/@@/@@@@"
               PromptChar      =   "_"
            End
            Begin VB.PictureBox cmdAtiEcon 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4440
               ScaleHeight     =   285
               ScaleWidth      =   225
               TabIndex        =   109
               Top             =   315
               Width           =   285
            End
            Begin VB.Label Label14 
               Caption         =   "Data de Início de Atividades: (ddmmaaaa)"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   120
               TabIndex        =   110
               Top             =   750
               Visible         =   0   'False
               Width           =   3165
            End
            Begin VB.Label Label15 
               Caption         =   "Livros Fiscais:"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   4965
               TabIndex        =   108
               Top             =   130
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VB.Label Label23 
               Caption         =   "Regime de Pagamento"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   6570
               TabIndex        =   107
               Top             =   130
               Width           =   1845
            End
            Begin VB.Label Label24 
               Caption         =   "Atividade Econômica:"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   120
               TabIndex        =   106
               Top             =   130
               Width           =   1665
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Contador"
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   350
            TabIndex        =   78
            Top             =   4920
            Width           =   3705
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   2520
               TabIndex        =   104
               Text            =   "Text2"
               Top             =   120
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.ComboBox cmbCrc 
               Height          =   315
               Left            =   150
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   360
               Width           =   3435
            End
            Begin VB.ComboBox cmb_relacao 
               Height          =   315
               Left            =   6375
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   360
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.TextBox txt_num_crc 
               Height          =   315
               Left            =   840
               MaxLength       =   7
               TabIndex        =   24
               Top             =   60
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.PictureBox cmd_buscar_crc 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   -600
               ScaleHeight     =   315
               ScaleWidth      =   855
               TabIndex        =   103
               Tag             =   "buscar"
               Top             =   600
               Visible         =   0   'False
               Width           =   915
            End
            Begin MSMask.MaskEdBox mskdtinicio 
               Height          =   315
               Left            =   3990
               TabIndex        =   22
               Top             =   360
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "@@/@@/@@@@"
               PromptChar      =   "_"
            End
            Begin VB.Label Label17 
               Caption         =   "Vínculo com empresa:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   6390
               TabIndex        =   80
               Top             =   120
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.Label Label16 
               Caption         =   "Data Início: (ddmmaaaa)"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3960
               TabIndex        =   79
               Top             =   120
               Visible         =   0   'False
               Width           =   1815
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Contribuinte"
            ForeColor       =   &H00FF0000&
            Height          =   1245
            Left            =   350
            TabIndex        =   73
            Top             =   510
            Width           =   8985
            Begin VB.TextBox msk_inscricao 
               Height          =   315
               Left            =   90
               MaxLength       =   12
               TabIndex        =   0
               Top             =   360
               Width           =   1545
            End
            Begin VB.TextBox txt_fantasia 
               Height          =   315
               Left            =   1800
               MaxLength       =   40
               TabIndex        =   3
               Top             =   870
               Width           =   5340
            End
            Begin VB.TextBox txt_razao 
               Height          =   315
               Left            =   1830
               MaxLength       =   60
               TabIndex        =   1
               Top             =   360
               Width           =   6600
            End
            Begin MSMask.MaskEdBox msk_cgc 
               Height          =   315
               Left            =   60
               TabIndex        =   2
               Top             =   870
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   18
               Format          =   "@@.@@@.@@@/@@@@-@@"
               PromptChar      =   "_"
            End
            Begin VB.Label Label41 
               Caption         =   "CGC/CPF"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   60
               TabIndex        =   77
               Top             =   675
               Width           =   1455
            End
            Begin VB.Label Label4 
               Caption         =   "Nome Fantasia:"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   1800
               TabIndex        =   76
               Top             =   680
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "Nome Empresarial"
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   1830
               TabIndex        =   75
               Top             =   165
               Width           =   1875
            End
            Begin VB.Label Label2 
               Caption         =   "Inscrição Estadual:"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   90
               TabIndex        =   74
               Top             =   160
               Width           =   1455
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Correio Eletrônico"
            ForeColor       =   &H00FF0000&
            Height          =   795
            Left            =   5820
            TabIndex        =   72
            Top             =   2985
            Width           =   3525
            Begin VB.TextBox txt_corr_eletronico 
               Height          =   315
               Left            =   90
               MaxLength       =   30
               TabIndex        =   16
               Top             =   420
               Width           =   3255
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "FAX"
            ForeColor       =   &H00FF0000&
            Height          =   795
            Left            =   3090
            TabIndex        =   69
            Top             =   2985
            Width           =   2295
            Begin VB.TextBox mskfax 
               Height          =   315
               Left            =   840
               MaxLength       =   10
               TabIndex        =   15
               Top             =   420
               Width           =   1095
            End
            Begin VB.TextBox mskddd_fax 
               Height          =   315
               Left            =   120
               MaxLength       =   5
               TabIndex        =   14
               Top             =   420
               Width           =   550
            End
            Begin VB.Label Label19 
               Caption         =   "Número:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   840
               TabIndex        =   71
               Top             =   180
               Width           =   735
            End
            Begin VB.Label Label18 
               Caption         =   "DDD:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   120
               TabIndex        =   70
               Top             =   180
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Telefone"
            ForeColor       =   &H00FF0000&
            Height          =   780
            Left            =   350
            TabIndex        =   66
            Top             =   2985
            Width           =   2295
            Begin VB.TextBox msktel 
               Height          =   315
               Left            =   840
               MaxLength       =   10
               TabIndex        =   13
               Top             =   420
               Width           =   1215
            End
            Begin VB.TextBox mskddd 
               Height          =   315
               Left            =   120
               MaxLength       =   5
               TabIndex        =   12
               Top             =   420
               Width           =   550
            End
            Begin VB.Label Label13 
               Caption         =   "Número:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   840
               TabIndex        =   68
               Top             =   180
               Width           =   735
            End
            Begin VB.Label Label12 
               Caption         =   "DDD:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   120
               TabIndex        =   67
               Top             =   180
               Width           =   495
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Endereço"
            ForeColor       =   &H00FF0000&
            Height          =   1245
            Left            =   350
            TabIndex        =   33
            Top             =   1740
            Width           =   8985
            Begin VB.ComboBox cmb_MUN 
               Height          =   315
               Left            =   3075
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Tag             =   "municipio"
               Top             =   870
               Width           =   3255
            End
            Begin VB.ComboBox CMB_UF 
               Height          =   315
               Left            =   6705
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   870
               Width           =   800
            End
            Begin VB.TextBox txt_bairro 
               Height          =   315
               Left            =   120
               MaxLength       =   20
               TabIndex        =   8
               Top             =   870
               Width           =   2535
            End
            Begin VB.TextBox txt_num 
               Height          =   315
               Left            =   5100
               MaxLength       =   5
               TabIndex        =   6
               Top             =   360
               Width           =   780
            End
            Begin VB.TextBox txt_comp 
               Height          =   315
               Left            =   6000
               MaxLength       =   20
               TabIndex        =   7
               Top             =   360
               Width           =   2835
            End
            Begin VB.TextBox txt_log 
               Height          =   315
               Left            =   1440
               MaxLength       =   30
               TabIndex        =   5
               Top             =   360
               Width           =   3525
            End
            Begin VB.ComboBox cmb_tipo_logradouro 
               Height          =   315
               ItemData        =   "DEC001.frx":0008
               Left            =   120
               List            =   "DEC001.frx":000A
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   360
               Width           =   1155
            End
            Begin MSMask.MaskEdBox mskcep 
               Height          =   315
               Left            =   7850
               TabIndex        =   11
               Top             =   870
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "@@.@@@-@@@"
               PromptChar      =   "_"
            End
            Begin VB.Label Label11 
               Caption         =   "CEP:"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   7850
               TabIndex        =   65
               Top             =   675
               Width           =   375
            End
            Begin VB.Label Label10 
               Caption         =   "UF:"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   6705
               TabIndex        =   64
               Top             =   675
               Width           =   855
            End
            Begin VB.Label Label9 
               Caption         =   "Município:"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   3075
               TabIndex        =   63
               Top             =   675
               Width           =   855
            End
            Begin VB.Label Label8 
               Caption         =   "Bairro:"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   120
               TabIndex        =   62
               Top             =   675
               Width           =   855
            End
            Begin VB.Label Label7 
               Caption         =   "Número:"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   5100
               TabIndex        =   61
               Top             =   160
               Width           =   735
            End
            Begin VB.Label Label6 
               Caption         =   "Complemento:"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   6000
               TabIndex        =   60
               Top             =   160
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Logradouro:"
               ForeColor       =   &H00FF0000&
               Height          =   200
               Left            =   120
               TabIndex        =   34
               Top             =   160
               Width           =   855
            End
         End
         Begin VB.PictureBox outAtEcon 
            BackColor       =   &H00FFFFFF&
            Height          =   2745
            Left            =   630
            ScaleHeight     =   2685
            ScaleWidth      =   5955
            TabIndex        =   81
            Top             =   1320
            Visible         =   0   'False
            Width           =   6015
         End
         Begin VB.PictureBox cmd_iniciar_contr 
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   870
            ScaleHeight     =   285
            ScaleWidth      =   1155
            TabIndex        =   25
            Tag             =   "iniciar"
            Top             =   150
            Width           =   1215
         End
         Begin VB.PictureBox cmd_buscar_contr 
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2160
            ScaleHeight     =   285
            ScaleWidth      =   1155
            TabIndex        =   26
            Tag             =   "buscar"
            Top             =   150
            Width           =   1215
         End
         Begin VB.PictureBox cmd_salvar_contr 
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   3480
            ScaleHeight     =   285
            ScaleWidth      =   1155
            TabIndex        =   27
            Top             =   150
            Width           =   1215
         End
         Begin VB.PictureBox cmd_excluir_contr 
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   4800
            ScaleHeight     =   285
            ScaleWidth      =   1155
            TabIndex        =   28
            Top             =   150
            Width           =   1215
         End
         Begin VB.PictureBox cmd_sair_contr 
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   7500
            ScaleHeight     =   285
            ScaleWidth      =   1155
            TabIndex        =   30
            Tag             =   "sair"
            Top             =   150
            Width           =   1215
         End
         Begin VB.PictureBox cmd_imprime_contr 
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   6180
            ScaleHeight     =   285
            ScaleWidth      =   1155
            TabIndex        =   29
            Top             =   150
            Width           =   1215
         End
      End
      Begin VB.Frame frame_contador 
         Height          =   5535
         Left            =   135
         TabIndex        =   82
         Top             =   315
         Width           =   9285
         Begin VB.Frame Frame13 
            Caption         =   "Dados da(o)  Esposa(o)"
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   120
            TabIndex        =   111
            Top             =   4440
            Visible         =   0   'False
            Width           =   9015
            Begin VB.TextBox txt_nome_esp 
               Height          =   285
               Left            =   120
               MaxLength       =   50
               TabIndex        =   52
               Top             =   495
               Width           =   5535
            End
            Begin MSMask.MaskEdBox Msk_cpf_esp 
               Height          =   315
               Left            =   5880
               TabIndex        =   53
               Top             =   450
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   14
               Format          =   "@@@.@@@.@@@-@@"
               PromptChar      =   "_"
            End
            Begin VB.Label Label20 
               Caption         =   "CPF:"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   5880
               TabIndex        =   113
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label21 
               Caption         =   "Nome "
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   120
               TabIndex        =   112
               Top             =   285
               Width           =   735
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Correio Eletrônico"
            ForeColor       =   &H00FF0000&
            Height          =   855
            Left            =   5370
            TabIndex        =   102
            Top             =   3480
            Width           =   3750
            Begin VB.TextBox txt_corr_eletr_conta 
               Height          =   315
               Left            =   120
               MaxLength       =   30
               TabIndex        =   51
               Top             =   480
               Width           =   3255
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "FAX"
            ForeColor       =   &H00FF0000&
            Height          =   855
            Left            =   2730
            TabIndex        =   99
            Top             =   3480
            Width           =   2295
            Begin VB.TextBox mskdddfax_conta 
               Height          =   315
               Left            =   120
               MaxLength       =   5
               TabIndex        =   49
               Top             =   480
               Width           =   550
            End
            Begin VB.TextBox mskfax_conta 
               Height          =   315
               Left            =   840
               MaxLength       =   10
               TabIndex        =   50
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label39 
               Caption         =   "DDD:"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   120
               TabIndex        =   101
               Top             =   280
               Width           =   495
            End
            Begin VB.Label Label38 
               Caption         =   "Número:"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   840
               TabIndex        =   100
               Top             =   280
               Width           =   735
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Telefone"
            ForeColor       =   &H00FF0000&
            Height          =   855
            Left            =   90
            TabIndex        =   96
            Top             =   3480
            Width           =   2295
            Begin VB.TextBox mskddd_conta 
               Height          =   315
               Left            =   120
               MaxLength       =   5
               TabIndex        =   47
               Top             =   480
               Width           =   550
            End
            Begin VB.TextBox msktel_conta 
               Height          =   315
               Left            =   840
               MaxLength       =   10
               TabIndex        =   48
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label37 
               Caption         =   "DDD:"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   120
               TabIndex        =   98
               Top             =   280
               Width           =   495
            End
            Begin VB.Label Label36 
               Caption         =   "Número:"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   840
               TabIndex        =   97
               Top             =   280
               Width           =   735
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Contador"
            ForeColor       =   &H00FF0000&
            Height          =   855
            Left            =   210
            TabIndex        =   91
            Top             =   690
            Width           =   9015
            Begin VB.TextBox txt_num_crc_conta 
               Height          =   315
               Left            =   120
               MaxLength       =   7
               TabIndex        =   35
               Tag             =   "num_crc_conta"
               Top             =   480
               Width           =   855
            End
            Begin VB.ComboBox cmb_uf_crc_conta 
               Height          =   315
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   480
               Width           =   800
            End
            Begin VB.TextBox txt_nome 
               Height          =   315
               Left            =   2310
               MaxLength       =   32
               TabIndex        =   37
               Top             =   480
               Width           =   4215
            End
            Begin MSMask.MaskEdBox msk_cpf 
               Height          =   315
               Left            =   6840
               TabIndex        =   38
               Top             =   480
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   14
               Format          =   "@@@.@@@.@@@-@@"
               PromptChar      =   "_"
            End
            Begin VB.Label Label35 
               Caption         =   "CPF:"
               ForeColor       =   &H00C00000&
               Height          =   200
               Left            =   6840
               TabIndex        =   95
               Top             =   270
               Width           =   615
            End
            Begin VB.Label Label34 
               Caption         =   "CRC:"
               ForeColor       =   &H00C00000&
               Height          =   200
               Left            =   120
               TabIndex        =   94
               Top             =   270
               Width           =   495
            End
            Begin VB.Label Label33 
               Caption         =   "UF CRC:"
               ForeColor       =   &H00C00000&
               Height          =   200
               Left            =   1080
               TabIndex        =   93
               Top             =   270
               Width           =   735
            End
            Begin VB.Label Label32 
               Caption         =   "Nome:"
               ForeColor       =   &H00C00000&
               Height          =   200
               Left            =   2280
               TabIndex        =   92
               Top             =   270
               Width           =   615
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Endereço Profissional"
            ForeColor       =   &H00FF0000&
            Height          =   1635
            Left            =   90
            TabIndex        =   83
            Top             =   1800
            Width           =   9015
            Begin VB.ComboBox cmb_tipo_logra_conta 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   510
               Width           =   1155
            End
            Begin VB.TextBox txt_log_conta 
               Height          =   315
               Left            =   1350
               MaxLength       =   30
               TabIndex        =   40
               Top             =   510
               Width           =   3525
            End
            Begin VB.TextBox txt_comp_conta 
               Height          =   315
               Left            =   6030
               MaxLength       =   20
               TabIndex        =   42
               Top             =   510
               Width           =   2835
            End
            Begin VB.TextBox txt_num_conta 
               Height          =   315
               Left            =   5040
               MaxLength       =   5
               TabIndex        =   41
               Top             =   510
               Width           =   780
            End
            Begin VB.TextBox txt_bairro_conta 
               Height          =   315
               Left            =   120
               MaxLength       =   20
               TabIndex        =   43
               Top             =   1160
               Width           =   2535
            End
            Begin VB.ComboBox cmb_uf_conta 
               Height          =   315
               Left            =   6630
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   1160
               Width           =   800
            End
            Begin VB.ComboBox cmb_mun_conta 
               Height          =   315
               Left            =   2880
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Tag             =   "municipio"
               Top             =   1160
               Width           =   3255
            End
            Begin MSMask.MaskEdBox mskcep_conta 
               Height          =   315
               Left            =   7875
               TabIndex        =   46
               Top             =   1160
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "@@.@@@-@@@"
               PromptChar      =   "_"
            End
            Begin VB.Label Label31 
               Caption         =   "Logradouro:"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label30 
               Caption         =   "Complemento:"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   6030
               TabIndex        =   89
               Top             =   300
               Width           =   1095
            End
            Begin VB.Label Label29 
               Caption         =   "Número:"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   5040
               TabIndex        =   88
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label28 
               Caption         =   "Bairro:"
               ForeColor       =   &H00C00000&
               Height          =   200
               Left            =   120
               TabIndex        =   87
               Top             =   950
               Width           =   855
            End
            Begin VB.Label Label27 
               Caption         =   "Município:"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   2880
               TabIndex        =   86
               Top             =   950
               Width           =   855
            End
            Begin VB.Label Label26 
               Caption         =   "UF:"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   6630
               TabIndex        =   85
               Top             =   950
               Width           =   855
            End
            Begin VB.Label Label25 
               Caption         =   "CEP:"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   7875
               TabIndex        =   84
               Top             =   950
               Width           =   375
            End
         End
         Begin VB.PictureBox cmd_imprime_conta 
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   5400
            ScaleHeight     =   315
            ScaleWidth      =   1155
            TabIndex        =   58
            Top             =   240
            Width           =   1215
         End
         Begin VB.PictureBox cmd_sair_conta 
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   7050
            ScaleHeight     =   315
            ScaleWidth      =   1155
            TabIndex        =   59
            Tag             =   "sair"
            Top             =   240
            Width           =   1215
         End
         Begin VB.PictureBox cmd_excluir_conta 
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   4080
            ScaleHeight     =   315
            ScaleWidth      =   1155
            TabIndex        =   57
            Top             =   240
            Width           =   1215
         End
         Begin VB.PictureBox cmd_salvar_conta 
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   2760
            ScaleHeight     =   315
            ScaleWidth      =   1155
            TabIndex        =   56
            Top             =   240
            Width           =   1215
         End
         Begin VB.PictureBox cmd_buscar_conta 
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1440
            ScaleHeight     =   315
            ScaleWidth      =   1155
            TabIndex        =   55
            Tag             =   "buscar"
            Top             =   240
            Width           =   1215
         End
         Begin VB.PictureBox cmd_iniciar_conta 
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   1155
            TabIndex        =   54
            Tag             =   "iniciar"
            Top             =   240
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "dec001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************Decalação de variáveis*********************************************
Option Explicit
Dim ja_estou_indo_cadastrar As Integer
Dim cRetorno As String
Dim mycgc As String
Dim mycep As String
Dim ret_tel As Integer
Dim msg As String
Dim MyInscricao As String
Dim mensagem As String
Dim Maria As Integer
Dim Joao As Integer
Dim flag_entrando_tab As Integer
Dim gravacao_ok As Integer
Dim vou_cadastrar_contador As Integer
Dim so_tenho_uma_atividade As Integer
Dim erro_geral As Integer
'*********************************************************************************************

'*********************************************************************************************
Function critica_contador(mensagem As String, Campo As Integer) As Integer
'****** Obrigatórios
    critica_contador = True
    Campo = 0
    If Trim(txt_num_crc_conta) = "" Then
      mensagem = "Número de CRC obrigatório."
      Campo = 1
      critica_contador = False
      Exit Function
    End If
    reta = CRIT_CRC(txt_num_crc_conta)
    If reta <> 0 Then
       mensagem = "Número de CRC Inválido."
       Campo = 1
       critica_contador = False
       Exit Function
    End If
    If Trim(cmb_uf_crc_conta.Text) = "" Then
       mensagem = "UF do CRC obrigatória."
       Campo = 2
       critica_contador = False
       Exit Function
    End If
    retorno = CritUF(cmb_uf_crc_conta.Text)
    If Len(retorno) > 0 Then
       mensagem = "UF do CRC inválida."
       Campo = 2
       critica_contador = False
       Exit Function
    End If
    If Trim(txt_nome) = "" Then
      mensagem = "Nome do Contador Obrigatório."
      Campo = 3
      critica_contador = False
      Exit Function
    End If
    cRetorno = CritString(txt_nome)
    If Len(cRetorno) > 0 Then
      mensagem = "Nome do Contador " & cRetorno
      Campo = 3
      critica_contador = False
      Exit Function
    End If
    If Trim(msk_cpf) = "" Then
       mensagem = "Número de CPF Obrigatório."
       critica_contador = False
       Campo = 4
       Exit Function
    End If
    reta = Rotcpf(msk_cpf)
    If reta = 0 Then
       mensagem = "Número de CPF Inválido."
       critica_contador = False
       Campo = 4
       Exit Function
    End If
    If Trim(Mid(cmb_tipo_logra_conta.Text, 1, 3)) = "" Then
        mensagem = "Tipo de Logradouro Inválido."
        critica_contador = False
        Campo = 5
        Exit Function
    End If
    If Not critica_tipo_log(Trim(Right(cmb_tipo_logra_conta.Text, 3))) Then
        mensagem = "Tipo de Logradouro Inválido."
        critica_contador = False
        Campo = 5
        Exit Function
    End If
    cRetorno = CritString(txt_log_conta)
    If Len(cRetorno) > 0 Then
        mensagem = "Logradouro " & cRetorno
        critica_contador = False
        Campo = 6
        Exit Function
    End If
    If Len(Trim(cmb_mun_conta.Text)) = 0 Then
        mensagem = "Município inválido."
        critica_contador = False
        Campo = 7
        Exit Function
    End If
    retorno = ""
    retorno = CritUF(cmb_uf_conta.Text)
    If Len(retorno) > 0 Then
       mensagem = "UF inválida."
       Campo = 8
       critica_contador = False
       Exit Function
    End If
    If cmb_mun_conta.ListIndex = 216 And cmb_uf_conta.ListIndex = 9 Then
       mensagem = "UF Não pode ser Maranhão."
       Campo = 8
       critica_contador = False
       Exit Function
    End If
     If Len(Trim(MSKcep_conta)) = 0 Then
        mensagem = "CEP campo obrigatório."
        critica_contador = False
        Campo = 9
        Exit Function
     End If
     ret_tel = 0
     ret_tel = critica_telefone(mskddd_conta, msktel_conta)
     If ret_tel <> 0 Then
      If ret_tel = 1 Then
         mensagem = "DDD informado e Telefone não informado."
         critica_contador = False
         Campo = 17
         Exit Function
      ElseIf ret_tel = 2 Then
         mensagem = "Telefone informado e DDD não informado."
         critica_contador = False
         Campo = 17
         Exit Function
     Else
         mensagem = "DDD Inválido."
         critica_contador = False
         Campo = 17
         Exit Function
      End If
    End If
    If cmb_mun_conta.ListIndex = 216 And Trim(mskddd_conta) = "098" Then
        mensagem = "DDD Inválido."
        critica_contador = False
        Campo = 17
        Exit Function
    End If
    ret_tel = 0
    ret_tel = critica_telefone(mskdddfax_conta, mskfax_conta)
    If ret_tel <> 0 Then
        If ret_tel = 1 Then
            mensagem = "DDD do Fax informado e Fax  não informado."
            critica_contador = False
            Campo = 18
            Exit Function
        ElseIf ret_tel = 2 Then
            mensagem = "Fax informado e DDD não informado."
            critica_contador = False
            Campo = 18
            Exit Function
        Else
            mensagem = "DDD Inválido."
            critica_contador = False
            Campo = 18
            Exit Function
         End If
    End If
    If cmb_mun_conta.ListIndex = 216 And Trim(mskdddfax_conta) = "098" Then
       mensagem = "DDD Inválido."
       critica_contador = False
       Campo = 17
       Exit Function
    End If
    If Trim(txt_corr_eletr_conta) <> "" Then
        If Not critica_correio_eletronico(txt_corr_eletr_conta) Then
             mensagem = "Correio Eletrônico Inválido."
             critica_contador = False
             Campo = 20
             Exit Function
         End If
   End If
' TIRAR LINHAS ABAIXO
'****** NÃO Obrigatórios
'If Len(Trim(txt_comp_conta)) > 0 Then
'   cRetorno = CritString(txt_comp_conta)
'  If Len(cRetorno) > 0 Then
'     mensagem = "Complemento " & cRetorno
'     critica_contador = False
'     Exit Function
'  End If
'End If

'If Len(Trim(txt_bairro_conta)) > 0 Then
'  cRetorno = CritString(txt_bairro_conta)
'  If Len(cRetorno) > 0 Then
'     mensagem = "Bairro " & cRetorno
'     critica_contador = False
'     Exit Function
'  End If
'End If
End Function
'*********************************************************************************************

'*********************************************************************************************
Function critica_contrib(mensagem As String, Campo As Integer) As Integer
Dim datasis As Double
Dim dataCont  As Double
Dim data_inicio_op As Double
critica_contrib = True
MyInscricao = msk_inscricao.Text
MyInscricao = tirapictu(MyInscricao, ".")
MyInscricao = tirapictu(MyInscricao, "-")
    Campo = 0
    If Trim(MyInscricao) = "" Then
       critica_contrib = False
       Campo = 1
       mensagem = "Número de Inscrição Inválido."
       Exit Function
    End If
    If Val(Mid(Trim(MyInscricao), 1, 2)) <> 12 Then
       Campo = 1
       critica_contrib = False
       mensagem = "Número de Inscrição Inválido."
       Exit Function
    End If
    If GenDV(CStr(Mid(MyInscricao, 1, 8)), BASE9) <> Val(Mid(MyInscricao, 9, 1)) Then
       Campo = 1
       critica_contrib = False
       mensagem = "Número de Inscrição Inválido."
       Exit Function
    End If
    cRetorno = CritString(txt_razao) 'critica razao social
    If Len(cRetorno) > 0 Then
       critica_contrib = False
       Campo = 2
       mensagem = "Nome Empresarial " & cRetorno
       Exit Function
    End If
    mycgc = msk_cgc.Text
    mycgc = tirapictu(mycgc, ".")
    mycgc = tirapictu(mycgc, "/")
    mycgc = tirapictu(mycgc, "-")
    Dim msg As String ' VAR DEFINIDA EM LOCAL INCORRETO
    If Len(mycgc) = 14 Then
       If Not Validar_Cgc_c(mycgc, mensagem) Then
          critica_contrib = False
          Campo = 3
          Exit Function
       End If
    ElseIf Len(mycgc) = 11 Then
         reta = Rotcpf(mycgc)
         If reta = 0 Then
            critica_contrib = False
            Campo = 3
            Exit Function
         End If
     End If
    If Not critica_tipo_log(Right(cmb_tipo_logradouro.Text, 3)) Then
      mensagem = "Tipo de Logradouro Inválido."
      critica_contrib = False
      Campo = 4
      Exit Function
    End If
    ' TIRAR LINHAS ABAIXO
   'If Len(Trim(txt_log)) = 0 Then
   '     mensagem = "Logradouro deve ser informado."
   '     critica_contrib = False
   '     Campo = 5
   '     Exit Function
   'End If
    cRetorno = CritString(txt_log)
    If Len(cRetorno) > 0 Then
       mensagem = "Logradouro " & cRetorno
       critica_contrib = False
       Campo = 5
       Exit Function
    End If
    cRetorno = CritString(txt_bairro)
    If Len(cRetorno) > 0 Then
       mensagem = "Bairro " & cRetorno
       critica_contrib = False
       Campo = 6
       Exit Function
    End If
    If Len(Trim(cmb_MUN.Text)) = 0 Then
       mensagem = "Município inválido."
       critica_contrib = False
       Campo = 7
       Exit Function
    End If
    retorno = ""
    retorno = CritUF(CMB_UF.Text)
    If Len(retorno) > 0 Then
      mensagem = "UF inválida."
      critica_contrib = False
      Campo = 8
      Exit Function
    End If
    If cmb_MUN.ListIndex = 216 And CMB_UF.ListIndex = 9 Then
        mensagem = "UF Não pode ser Maranhão."
        critica_contrib = False
        Campo = 8
       Exit Function
    End If
    If Len(Trim(MSKcep)) = 0 Then
       mensagem = "CEP campo obrigatório."
       critica_contrib = False
       Campo = 9
       Exit Function
    End If
    If Len(Trim(MSKcep)) < 8 Then
       mensagem = "Número do CEP Inválido."
       critica_contrib = False
       Campo = 9
       Exit Function
    End If
    If Trim(Label1.Text) = "" Then
        mensagem = "Atividade Economica Obtrigatória."
        critica_contrib = False
        Campo = 10
        Exit Function
   End If
   reta = CritATIVID1(Label1.Text)
   If reta <> 0 Then
     mensagem = "Atividade Economica Inválida."
     critica_contrib = False
     Campo = 10
     Exit Function
   End If
   reta = Crit_relacao(cmb_localizacao.ListIndex + 1)
   If Not reta Then
       mensagem = "Localização do Livro Fiscal inválido."
       critica_contrib = False
       Campo = 11
       Exit Function
   End If
   If Len(Trim(cmb_tipo.Text)) = 0 Then
       mensagem = "Condição do Contribuinte não preenchida."
       critica_contrib = False
       Campo = 12
       Exit Function
    End If
    If (cmb_localizacao.ListIndex + 1) = 2 Then   'a localizacao do livro está no contador
       If Len(Trim(txt_num_crc)) = 0 Then
          mensagem = "CRC do Contador obrigatório quando Localização do Livro Fiscal for no Contador."
          critica_contrib = False
          Campo = 19
          Exit Function
        End If
    End If
    If Trim(cmb_tipo.Text) = "Normal" And Trim(cmbCrc.Text) = "" And Len(msk_cgc.ClipText) <> 11 Then
          mensagem = "É obrigatório informar Contador."
          critica_contrib = False
          Campo = 19
          Exit Function
    End If
    If Trim(Mask_dt_inicio_op) = "" Then
          mensagem = "Data de Início de Operações Obrigatoria."
          critica_contrib = False
          Campo = 21
          Exit Function
    End If
    reta = CRIT_Data(Mask_dt_inicio_op)
    If reta <> 0 Then
       mensagem = "Data Inválida."
       critica_contrib = False
       Campo = 21
       Exit Function
    End If
    datasis = Val(Year(Date)) & Mid$(Date, 4, 2) & Mid$(Date, 1, 2)
    dataCont = Val(Mid$(Mask_dt_inicio_op, 5, 4) & Mid$(Mask_dt_inicio_op, 3, 2) & Mid$(Mask_dt_inicio_op, 1, 2))
    If dataCont > datasis Then
      mensagem = "Data de Início de Operações não pode ser maior que data atual."
      critica_contrib = False
      Campo = 21
      Exit Function
    End If
    If Len(Trim(txt_num_crc)) <> 0 Then
      reta = CRIT_CRC(txt_num_crc)
      If reta <> 0 Then
         mensagem = "CRC inválido."
         critica_contrib = False
         Campo = 15
         Exit Function
      End If
      'TIRARA LINHAS ABAIXO
     ' If Len(Trim(mskdtinicio)) = 0 Or Len(Trim(cmb_relacao.Text)) = 0 Then
     '    mensagem = "Dados do Contador incompletos."
     '    critica_contrib = False
     '    Campo = 15
     '    Exit Function
     ' End If
    '  reta = CRIT_Data(mskdtinicio)
    '  If reta <> 0 Then
    '    mensagem = "Data Inválida."
    '    critica_contrib = False
    '    Campo = 15
    '    Exit Function
    '  End If
    '  datasis = Val(Year(Date)) & Mid$(Date, 4, 2) & Mid$(Date, 1, 2)
    'dataCont = Val(Mid$(mskdtinicio, 5, 4) & Mid$(mskdtinicio, 3, 2) & Mid$(mskdtinicio, 1, 2))
    'If dataCont > datasis Then
    '  mensagem = "Data de início não pode ser maior que data atual."
    '  critica_contrib = False
    '  Campo = 15
    '  Exit Function
    'End If
    'data_inicio_op = Val(Mid$(Mask_dt_inicio_op, 5, 4) & Mid$(Mask_dt_inicio_op, 3, 2) & Mid$(Mask_dt_inicio_op, 1, 2))
    'If dataCont < data_inicio_op Then
    '   mensagem = "Data de início não pode ser menor que data inicio de atividades."
    '   critica_contrib = False
    '   Campo = 15
    '   Exit Function
    'End If

      
    '  reta = Crit_relacao(cmb_relacao.ListIndex + 1)
    '  If Not reta Then
    '    mensagem = "Vínculo com Empresa inválido."
    '    critica_contrib = False
    '    Campo = 16
    '    Exit Function
    '  End If
    
    'ElseIf Len(Trim(cmb_relacao.Text)) <> 0 Or Len(Trim(mskdtinicio)) <> 0 Then
    '     mensagem = "Dados do Contador incompletos."
    '     critica_contrib = False
    '     Campo = 13
    '     Exit Function
         
    End If
    ret_tel = 0
    ret_tel = critica_telefone(mskddd, msktel)
    If ret_tel <> 0 Then
        If ret_tel = 1 Then
           mensagem = "DDD informado e Telefone não informado."
           critica_contrib = False
           Campo = 17
           Exit Function
        ElseIf ret_tel = 2 Then
           mensagem = "Telefone informado e DDD não informado."
           critica_contrib = False
           Campo = 17
           Exit Function
        Else
           mensagem = "DDD Inválido."
           critica_contrib = False
           Campo = 17
           Exit Function
        
        End If
    End If

    If cmb_MUN.ListIndex = 216 And Trim(mskddd) = "098" Then
         mensagem = "DDD Inválido."
         critica_contrib = False
         Campo = 17
         Exit Function
    End If

    ret_tel = 0
    ret_tel = critica_telefone(mskddd_fax, mskfax)
    If ret_tel <> 0 Then
        If ret_tel = 1 Then
           mensagem = "DDD do Fax informado e Fax  não informado."
           critica_contrib = False
           Campo = 18
           Exit Function
        ElseIf ret_tel = 2 Then
           mensagem = "Fax informado e DDD não informado."
           critica_contrib = False
           Campo = 18
           Exit Function
        Else
           mensagem = "DDD Inválido."
           critica_contrib = False
           Campo = 18
           Exit Function
        End If
    End If
    If cmb_MUN.ListIndex = 216 And Trim(mskddd_fax) = "098" Then
       mensagem = "DDD Inválido."
       critica_contrib = False
       Campo = 17
       Exit Function
    End If
    If Trim(txt_corr_eletronico) <> "" Then
        If Not critica_correio_eletronico(txt_corr_eletronico) Then
             mensagem = "Correio Eletrônico Inválido."
             critica_contrib = False
             Campo = 20
             Exit Function
         End If
    End If
End Function
'*********************************************************************************************

'*********************************************************************************************
 Sub grava_contador()
      Screen.MousePointer = 11
      Dim mycpf As String
      Dim mycep As String
      Dim mycpfconj As String
      mycep = MSKcep_conta.Text
      mycep = tirapictu(mycep, ".")
      mycep = tirapictu(mycep, "/")
      mycep = tirapictu(mycep, "-")
      mycpf = msk_cpf.Text
      mycpf = tirapictu(mycpf, ".")
      mycpf = tirapictu(mycpf, "/")
      mycpf = tirapictu(mycpf, "-")
      mycpfconj = Msk_cpf_esp.Text
      mycpfconj = tirapictu(Msk_cpf_esp, ".")
      mycpfconj = tirapictu(Msk_cpf_esp, "/")
      mycpfconj = tirapictu(Msk_cpf_esp, "-")
      query = "select * from tab_cad_contadores where tco_uf_crc = '" & Trim(UCase(cmb_uf_crc_conta.Text)) & "'"
      query = query & " and tco_num_crc = '" & txt_num_crc_conta & "'"
      Set Contadores = dbgim.OpenRecordset(query, dbOpenDynaset)
      If Contadores.EOF And Contadores.BOF Then
         'TIRAR LINHAS ABAIXO
         'query = "insert into tab_cad_contadores"
         'query = query & "(tco_uf_crc, tco_num_crc,tco_num_cpf,"
         'query = query & "tco_nome_contador,tco_ti_logradouro,tco_logradouro,"
         'query = query & "tco_numero,tco_complemento,tco_bairro, tco_municipio,"
         'query = query & "tco_cod_uf,tco_cep,tco_ddd_fone,tco_num_fone,"
         'query = query & "tco_ddd_fax,tco_num_fax,tco_corr_eletronico, "
         'query = query & "tco_Estado,tco_nome_conj,tco_cpf_conj) "
         'query = query & "values('" & Trim(UCase(cmb_uf_crc_conta.Text)) & "','" & txt_num_crc_conta & "','" & mycpf & "','"  'colocar o cpf
         'query = query & Trim(UCase(txt_nome)) & "','" & UCase(Trim(Right(cmb_tipo_logra_conta.Text, 3))) & "','" & Trim(UCase(txt_log_conta)) & "','"
         'query = query & Trim(UCase(txt_num_conta)) & "','" & Trim(UCase(txt_comp_conta)) & "','" & Trim(UCase(txt_bairro_conta))
         'query = query & "','" & UCase(Trim(Right(cmb_mun_conta.Text, 5))) & "','"
         'query = query & Trim(UCase(cmb_uf_conta.Text)) & "','" & mycep & "','" & mskddd_conta & "','" & msktel_conta & "','"
         'query = query & mskdddfax_conta & "','" & mskfax_conta & "','" & Trim(txt_corr_eletr_conta) & "','"
         'query = query & "A','" & txt_nome_esp & "','" & mycpfconj & "')"
         'tab_cad_contadores"
         Contadores.AddNew
         Contadores("tco_uf_crc") = Trim(UCase(cmb_uf_crc_conta.Text))
         Contadores("tco_num_crc") = txt_num_crc_conta
         Contadores("tco_num_cpf") = mycpf
         Contadores("tco_nome_contador") = Trim(UCase(txt_nome))
         Contadores("tco_ti_logradouro") = UCase(Trim(Right(cmb_tipo_logra_conta.Text, 3)))
         Contadores("tco_logradouro") = Trim(UCase(txt_log_conta))
         Contadores("tco_numero") = Trim(UCase(txt_num_conta))
         Contadores("tco_complemento") = Trim(UCase(txt_comp_conta))
         Contadores("tco_bairro") = Trim(UCase(txt_bairro_conta))
         Contadores("tco_municipio") = UCase(Trim(Right(cmb_mun_conta.Text, 5)))
         Contadores("tco_cod_uf") = Trim(UCase(cmb_uf_conta.Text))
         Contadores("tco_cep") = Mid(mycep, 1, 8)
         Contadores("tco_ddd_fone") = mskddd_conta
         Contadores("tco_num_fone") = msktel_conta
         Contadores("tco_ddd_fax") = mskdddfax_conta
         Contadores("tco_num_fax") = mskfax_conta
         Contadores("tco_corr_eletronico") = Trim(txt_corr_eletr_conta)
         Contadores("tco_Estado") = "A"
         Contadores("tco_nome_conj") = txt_nome_esp
         Contadores("tco_cpf_conj") = mycpfconj
         Contadores.Update
      Else
         query = "update tab_cad_contadores set "
         If Contadores.Fields("tco_num_cpf") <> mycpf Then query = query & "tco_num_cpf = '" & mycpf & "',"
         If Contadores.Fields("tco_nome_contador") <> UCase(Trim(txt_nome)) Then query = query & "tco_nome_contador = '" & UCase(Trim(txt_nome)) & "' ,"
         If Contadores.Fields("tco_ti_logradouro") <> UCase(Trim(Right(cmb_tipo_logra_conta.Text, 3))) Then query = query & "tco_ti_logradouro = '" & UCase(Trim(Right(cmb_tipo_logra_conta.Text, 3))) & "' ,"
         If Contadores.Fields("tco_logradouro") <> UCase(Trim(txt_log_conta)) Then query = query & "tco_logradouro = '" & UCase(Trim(txt_log_conta)) & "' ,"
         If Contadores.Fields("tco_numero") <> UCase(Trim(txt_num_conta)) Then query = query & "tco_numero = '" & UCase(Trim(txt_num_conta)) & "' ,"
         If Contadores.Fields("tco_complemento") <> UCase(Trim(txt_comp_conta)) Then query = query & "tco_complemento = '" & UCase(Trim(txt_comp_conta)) & "' ,"
         If Contadores.Fields("tco_bairro") <> UCase(Trim(txt_bairro_conta)) Then query = query & "tco_bairro = '" & UCase(Trim(txt_bairro_conta)) & "' ,"
         If Contadores.Fields("tco_municipio") <> Trim(Right(cmb_mun_conta.Text, 5)) Then query = query & "tco_municipio = '" & Trim(Right(cmb_mun_conta.Text, 5)) & "' ,"
         If Contadores.Fields("tco_cod_uf") <> UCase(Trim(cmb_uf_conta.Text)) Then query = query & "tco_cod_uf = '" & UCase(Trim(cmb_uf_conta.Text)) & "' ,"
         If Contadores.Fields("tco_cep") <> mycep Then query = query & "tco_cep = '" & mycep & "' ,"
         If Contadores.Fields("tco_ddd_fone") <> mskddd_conta Then query = query & "tco_ddd_fone = '" & mskddd_conta & "' ,"
         If Contadores.Fields("tco_num_fone") <> msktel_conta Then query = query & "tco_num_fone = '" & msktel_conta & "' ,"
         If Contadores.Fields("tco_ddd_fax") <> mskdddfax_conta Then query = query & "tco_ddd_fax = '" & mskdddfax_conta & "' ,"
         If Contadores.Fields("tco_num_fax") <> mskfax_conta Then query = query & "tco_num_fax = '" & mskfax_conta & "' ,"
         If Contadores.Fields("tco_corr_eletronico") <> Trim(txt_corr_eletr_conta) Then query = query & "tco_corr_eletronico = '" & Trim(txt_corr_eletr_conta) & "' ,"
         query = query & "tco_estado = 'A' ,tco_nome_conj = '" & txt_nome_esp & "', tco_cpf_conj = '" & mycpfconj & "'"
         query = query & "where tco_num_crc = '" & txt_num_crc_conta & "'"
         query = query & " and tco_uf_crc = '" & Trim(UCase(cmb_uf_crc_conta.Text)) & "'"
         If Not EjecutarQuery(query) Then
              Screen.MousePointer = 0
         End If
      End If
      Contadores.Close
      Screen.MousePointer = 0
End Sub
'*********************************************************************************************
Sub grava_contrib()
'*********************************************************************************************
'
    
    Dim tipo_contribuinte As Byte
    Dim QueryDel As String
    Dim ret As Integer
    Dim localizacao As Byte
    Dim relacao As Byte
    
    Screen.MousePointer = 11
    MyInscricao = msk_inscricao.Text
    MyInscricao = tirapictu(MyInscricao, ".")
    MyInscricao = tirapictu(MyInscricao, "-")
    mycgc = msk_cgc.Text
    mycgc = tirapictu(mycgc, ".")
    mycgc = tirapictu(mycgc, "/")
    mycgc = tirapictu(mycgc, "-")
    mycep = MSKcep.Text
    mycep = tirapictu(mycep, ".")
    mycep = tirapictu(mycep, "/")
    mycep = tirapictu(mycep, "-")
    query = "select * from tab_cad_contribuintes where tcc_nu_inscricao = '" & Trim(MyInscricao) & "'"
    'Set dyn_dare = dbgim.OpenRecordset(query, dbOpenDynaset)
    Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
    ' dbgim.BeginTrans
    tipo_contribuinte = cmb_tipo.ListIndex + 1
    localizacao = cmb_localizacao.ListIndex + 1
    relacao = cmb_relacao.ListIndex + 1
    If restauracontr.EOF And restauracontr.BOF Then
        query = "insert into tab_cad_contribuintes"
        query = query & "(tcc_nu_inscricao, tcc_cad_geral_contrib,"
        query = query & "tcc_razao_social,tcc_no_fantasia,tcc_ti_logradouro,"
        query = query & "tcc_no_logradouro,tcc_nu_logradouro,tcc_complemento, tcc_bairro,"
        query = query & "tcc_municipio,tcc_cod_uf,tcc_cep,"
        query = query & "tcc_ddd_fone,tcc_num_fone, "
        query = query & "tcc_ddd_fax,tcc_num_fax,tcc_corr_eletronico, "
        query = query & "tcc_ativ_economica, "
        query = query & "tcc_uf_crc_contador,tcc_num_crc_contador, "
        query = query & "tcc_da_inicio_contador,tcc_relacao_contador, tcc_loc_livro_fiscal, tcc_tipo_contribuinte,"
        query = query & "tcc_Estado ,tcc_da_inicio_op) "
        query = query & "values('" & Trim(MyInscricao) & "','" & mycgc & "','"
        query = query & UCase(Trim(txt_razao)) & "','" & UCase(Trim(txt_fantasia)) & "','"
        query = query & UCase(Trim(Right(cmb_tipo_logradouro.Text, 3))) & "','"
        query = query & UCase(Trim(txt_log)) & "','" & UCase(Trim(txt_num)) & "','" & UCase(Trim(txt_comp)) & "','" & UCase(txt_bairro) & "','"
        query = query & Trim(Right(cmb_MUN.Text, 5)) & "','" & UCase(Trim(CMB_UF.Text)) & "','" & mycep & "','"
    '     query = query & Trim(Mid(cmb_MUN.Text, 1, 5)) & "','" & UCase(Trim(CMB_UF.Text)) & "','" & mycep & "','"
        query = query & mskddd & "','" & msktel & "','"
        query = query & mskddd_fax & "','" & mskfax & "','" & Trim(txt_corr_eletronico) & "','"
        query = query & Mid(Trim(Label1.Text), 1, 5) & "','"
        query = query & UCase(Trim(Text2)) & "','" & Trim(Right(cmbCrc, 8)) & "','"
    '    query = query & mskdtinicio & "','" & Mid(cmb_relacao.Text, 1, 1) & "','" & Mid(cmb_localizacao.Text, 1, 1) & "'," & tipo_contribuinte & ","
        query = query & mskdtinicio & "','" & relacao & "','" & localizacao & "'," & tipo_contribuinte & ","
    
        query = query & "'A','" & Mask_dt_inicio_op & "' )"
    Else
        query = "update tab_cad_contribuintes set "
        If restauracontr.Fields("tcc_tipo_contribuinte") <> tipo_contribuinte Then
           ' Ret = MsgBox("A alteração da condição do contribuinte, acarretará a eliminação das declarações já informadas, se for o caso." & Chr(10) & "Confirma a alteração?", 20, "ATENÇÃO")
           'If Ret = 6 Then
           '   QueryDel = "Delete * from tab_declaracao where tde_num_inscricao = '" & myinscricao & "'"
           '     If Not EjecutarQuery(QueryDel) Then
           '     End If
           '   QueryDel = "Delete * from tab_detalhe_declaracao where tdd_num_inscricao = '" & myinscricao & "'"
           '     If Not EjecutarQuery(QueryDel) Then
           '     End If
              query = query & "tcc_tipo_contribuinte = '" & tipo_contribuinte & "' ,"
           'Else
           '   restauracontr.Close
           '   Screen.MousePointer = 0
           '   Exit Sub
           'End If
        End If
        If restauracontr.Fields("tcc_cad_geral_contrib") <> mycgc Then query = query & "tcc_cad_geral_contrib = '" & mycgc & "',"
        
        If restauracontr.Fields("tcc_razao_social") <> UCase(Trim(txt_razao)) Then query = query & "tcc_razao_social = '" & UCase(Trim(txt_razao)) & "' ,"
        
        If restauracontr.Fields("tcc_no_fantasia") <> UCase(Trim(txt_fantasia)) Then query = query & "tcc_no_fantasia = '" & UCase(Trim(txt_fantasia)) & "' ,"
        
        If restauracontr.Fields("tcc_ti_logradouro") <> UCase(Trim(Right(cmb_tipo_logradouro.Text, 3))) Then query = query & "tcc_ti_logradouro = '" & UCase(Trim(Right(cmb_tipo_logradouro.Text, 3))) & "' ,"
        
        If restauracontr.Fields("tcc_no_logradouro") <> UCase(Trim(txt_log)) Then query = query & "tcc_no_logradouro = '" & UCase(Trim(txt_log)) & "' ,"
        
        If restauracontr.Fields("tcc_nu_logradouro") <> UCase(Trim(txt_num)) Then query = query & "tcc_nu_logradouro = '" & UCase(Trim(txt_num)) & "' ,"
    
        If restauracontr.Fields("tcc_complemento") <> UCase(Trim(txt_comp)) Then query = query & "tcc_complemento = '" & UCase(Trim(txt_comp)) & "' ,"
        
        If restauracontr.Fields("tcc_bairro") <> UCase(Trim(txt_bairro)) Then query = query & "tcc_bairro = '" & UCase(Trim(txt_bairro)) & "' ,"
    
        If restauracontr.Fields("tcc_municipio") <> UCase(Trim(Right(cmb_MUN.Text, 5))) Then query = query & "tcc_municipio = '" & UCase(Trim(Right(cmb_MUN.Text, 5))) & "' ,"
    
        If restauracontr.Fields("tcc_cod_uf") <> UCase(Trim(CMB_UF.Text)) Then query = query & "tcc_cod_uf = '" & UCase(Trim(CMB_UF.Text)) & "' ,"
    
        If restauracontr.Fields("tcc_cep") <> mycep Then query = query & "tcc_cep = '" & mycep & "' ,"
    
        If restauracontr.Fields("tcc_ddd_fone") <> mskddd Then query = query & "tcc_ddd_fone = '" & mskddd & "' ,"
        
        If restauracontr.Fields("tcc_num_fone") <> msktel Then query = query & "tcc_num_fone = '" & msktel & "' ,"
    
        If restauracontr.Fields("tcc_ddd_fax") <> mskddd_fax Then query = query & "tcc_ddd_fax = '" & mskddd_fax & "' ,"
    
        If restauracontr.Fields("tcc_num_fax") <> mskfax Then query = query & "tcc_num_fax = '" & mskfax & "' ,"
    
        If restauracontr.Fields("tcc_corr_eletronico") <> Trim(txt_corr_eletronico) Then query = query & "tcc_corr_eletronico = '" & Trim(txt_corr_eletronico) & "' ,"
    
        If restauracontr.Fields("tcc_ativ_economica") <> Mid(Trim(Label1.Text), 1, 5) Then query = query & "tcc_ativ_economica = '" & Mid(Trim(Label1.Text), 1, 5) & "' ,"
    
        If restauracontr.Fields("tcc_uf_crc_contador") <> UCase(Trim(Text2)) Then query = query & "tcc_uf_crc_contador = '" & UCase(Trim(Text2)) & "' ,"
    
        If restauracontr.Fields("tcc_num_crc_contador") <> Trim(Right(cmbCrc, 8)) Then query = query & "tcc_num_crc_contador = '" & Trim(Right(cmbCrc, 8)) & "' ,"
    
        If restauracontr.Fields("tcc_da_inicio_contador") <> mskdtinicio Then query = query & "tcc_da_inicio_contador = '" & mskdtinicio & "' ,"
    
        If restauracontr.Fields("tcc_relacao_contador") <> Mid(cmb_relacao.Text, 1, 1) Then query = query & "tcc_relacao_contador = '" & relacao & "' ,"
    
        If restauracontr.Fields("tcc_loc_livro_fiscal") <> Mid(cmb_localizacao.Text, 1, 1) Then query = query & "tcc_loc_livro_fiscal = '" & localizacao & "' ,"
        If restauracontr.Fields("tcc_da_inicio_op") <> Mask_dt_inicio_op Then query = query & "tcc_da_inicio_op = '" & Mask_dt_inicio_op & "' ,"
    '    Mask_dt_inicio_op
        query = query & "tcc_estado = 'A' "
        query = query & "where tcc_nu_inscricao = '" & Trim(MyInscricao) & "'"
    End If
'   query = query & "A')"
    restauracontr.Close
    If Not EjecutarQuery(query) Then
       Screen.MousePointer = 0
       gravacao_ok = 1
       Exit Sub
    Else
       gravacao_ok = 0
       Screen.MousePointer = 0
    End If
 'dbgim.CommitTrans
End Sub
'*********************************************************************************************
Private Function Indent_1() As Integer
'*********************************************************************************************
'
    Dim reg As String
    Dim RegAux1 As String, RegAux2 As String
    Dim SubClass As Recordset
    Dim ret As String
    Dim SubClass1 As String
    Dim NumUltReg As Integer

    RegAux1 = outAtEcon.List(outAtEcon.ListIndex)
    RegAux2 = outAtEcon.List(outAtEcon.ListCount - 1)
    Do While outAtEcon.List(0) <> RegAux1
        outAtEcon.RemoveItem 0
    Loop
    If outAtEcon.ListCount <> 1 Then
        Do While outAtEcon.List(1) <> RegAux2
            outAtEcon.RemoveItem 1
        Loop
        outAtEcon.RemoveItem 1
    End If
    outAtEcon.Refresh
    SubClass1 = "SELECT * FROM AtividadeEconomica2 where mid$(trim$(cstr(AtividadeEconomica2.IT_CO_ATIVIDADE)),1,1) = '" & Mid$(Trim$(RegAux1), 1, 1) & "'"
    Set SubClass = dbgim.OpenRecordset(SubClass1, dbOpenSnapshot)
    Do While Not SubClass.EOF
        outAtEcon.AddItem SubClass("IT_CO_ATIVIDADE") & "   " & SubClass("IT_NO_ATIVIDADE")
      '  outAtEcon.Indent(outAtEcon.ListCount - 1) = 2
        SubClass.MoveNext
    Loop
    SubClass.Close
    NumUltReg = 0
   ' NumUltReg = outAtEcon.ListCount - 1
   ' outAtEcon.RemoveItem NumUltReg
    outAtEcon.Expand(0) = True
    Indent_1 = True
End Function
'*********************************************************************************************
Private Function Indent_2() As Integer
'*********************************************************************************************
'
    Dim ret As String
    Dim reg As String
    Dim RegAux1 As String, RegAux2 As String
    Dim Class As Recordset
    Dim Class1 As String
    Dim NumUltReg As Integer

    RegAux1 = outAtEcon.List(outAtEcon.ListIndex)
    RegAux2 = outAtEcon.List(outAtEcon.ListCount - 1)
    Do While outAtEcon.List(1) <> RegAux1
        outAtEcon.RemoveItem 1
    Loop
    If outAtEcon.ListCount > 2 Then
        Do While outAtEcon.List(2) <> RegAux2
           outAtEcon.RemoveItem 2
        Loop
        outAtEcon.RemoveItem 2
    End If
    outAtEcon.Refresh
    Class1 = "SELECT * FROM AtividadeEconomica3 where mid$(trim$(cstr(AtividadeEconomica3.IT_CO_ATIVIDADE)),1,3) = '" & Mid$(Trim$(RegAux1), 1, 3) & "'"
    Set Class = dbgim.OpenRecordset(Class1, dbOpenSnapshot)
    so_tenho_uma_atividade = 0
    Do While Not Class.EOF
        so_tenho_uma_atividade = so_tenho_uma_atividade + 1
        outAtEcon.AddItem Class("IT_CO_ATIVIDADE") & "   " & Class("IT_NO_ATIVIDADE")
     '   outAtEcon.Indent(outAtEcon.ListCount - 1) = 2
        Class.MoveNext
    Loop
    NumUltReg = 0
 '   NumUltReg = outAtEcon.ListCount - 1
  '  outAtEcon.RemoveItem NumUltReg  'por bellyrio em 21/3/98
    Class.Close
    outAtEcon.Expand(1) = True
    Indent_2 = True
End Function
'*********************************************************************************************
Private Sub Indent_3()
'*********************************************************************************************
'
    Dim atiecon As String
    atiecon = Mid$(outAtEcon.List(outAtEcon.ListIndex), 1, 2) + Mid$(outAtEcon.List(outAtEcon.ListIndex), 4, 2) + Mid$(outAtEcon.List(outAtEcon.ListIndex), 7, 1)
    If Len(outAtEcon.List(outAtEcon.ListIndex)) > 99 Then
        Label1.Text = " " + Mid$(outAtEcon.List(outAtEcon.ListIndex), 1, 90)
    Else
        Label1.Text = " " + Mid$(outAtEcon.List(outAtEcon.ListIndex), 1, Len(outAtEcon.List(outAtEcon.ListIndex)) - 9)
    End If
    out3 = True
    Label1.Refresh
End Sub
'*********************************************************************************************
Sub limpa_tela()
'*********************************************************************************************
'
    msk_inscricao = ""
    msk_cgc = ""
    txt_razao = ""
    txt_fantasia = ""
    txt_log = ""
    txt_num = ""
    txt_comp = ""
    txt_bairro = ""
    txt_comp = ""
    txt_bairro = ""
    Label1.Text = ""
    MSKcep = ""
    mskddd = ""
    msktel = ""
    mskddd_fax = ""
    mskfax = ""
    txt_corr_eletronico = ""
    txt_num_crc = ""
    mskdtinicio = ""
    cmbCrc.ListIndex = -1
    'Mask_dt_inicio_op = ""
    outAtEcon.Visible = False
    monta_combos_contrib
    teve_alteracao = False
    cmd_imprime_contr.Enabled = False
    cmd_salvar_contr.Enabled = False
    cmd_excluir_contr.Enabled = False
End Sub
'*********************************************************************************************
Sub limpa_tela_contador()
'*********************************************************************************************
'
   txt_num_crc_conta = ""
   txt_nome = ""
   msk_cpf = ""
   Msk_cpf_esp = ""
   txt_nome_esp = ""
   txt_log_conta = ""
   txt_num_conta = ""
   txt_comp_conta = ""
   txt_bairro_conta = ""
   txt_comp_conta = ""
   MSKcep_conta = ""
   mskddd_conta = ""
   msktel_conta = ""
   mskfax_conta = ""
   mskdddfax_conta = ""
   txt_corr_eletr_conta = ""
   monta_combos_contador
   teve_alteracao_contador = False
   cmd_imprime_conta.Enabled = False
   cmd_salvar_conta.Enabled = False
   cmd_excluir_conta.Enabled = False
End Sub
'*********************************************************************************************
Sub Monta_Combo_Crc()
'*********************************************************************************************
'
    cmbCrc.Clear
    strSQLA = "Select * from tab_cad_contadores "
    Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
    Do While Not PESQ_SNAP.EOF
        cmbCrc.AddItem PESQ_SNAP!tco_nome_contador & Space(100) & PESQ_SNAP!tco_num_crc
        PESQ_SNAP.MoveNext
    Loop
    cmbCrc.AddItem ""
End Sub
'*********************************************************************************************
Sub Monta_Combo_crc_click()
'*********************************************************************************************
'
    Dim i As Integer

    cmbCrc.Clear
    strSQLA = "Select * from tab_cad_contadores "
    Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
    Do While Not PESQ_SNAP.EOF
        cmbCrc.AddItem PESQ_SNAP!tco_nome_contador & Space(100) & PESQ_SNAP!tco_num_crc
        PESQ_SNAP.MoveNext
    Loop
    cmbCrc.AddItem ""
    For i = 0 To dec001.cmbCrc.ListCount - 1
        If dec001.txt_num_crc = Trim(Right(dec001.cmbCrc.List(i), 8)) Then
           dec001.cmbCrc.ListIndex = i 'data_contrib.Recordset("tcc_uf_crc_contador") & ""
           Exit For
        End If
    Next i
End Sub
'*********************************************************************************************
Sub monta_combos_contador()
'*********************************************************************************************
'
    Dim munic As String
    Dim tabMunic As Recordset

    munic = "SELECT * FROM tab_cod_munic " 'SELECIONA TUDO DA TABELA Municipios
    Set tabMunic = dbgim.OpenRecordset(munic, dbOpenSnapshot)
    cmb_mun_conta.Clear 'Limpa o combo antes de jogar os dados
    Screen.MousePointer = 11
    Do Until tabMunic.EOF 'PARA LEITURA DE  todos os registros
'         cmb_mun_conta.AddItem Format$(tabMunic("Cod_Munic"), "00000") & "-" & tabMunic("Nome_Munic")   'adiciona o codigo e a descrição ao combo
       cmb_mun_conta.AddItem tabMunic("Nome_Munic") & Space(130) & Format$(tabMunic("Cod_Munic"), "00000")   'adiciona só a descrição ao combo
       tabMunic.MoveNext ' Lê o próximo registro
    Loop
    cmb_mun_conta.AddItem "MUNICIPIOS DE OUTROS ESTADOS" & Space(130) & "99999"
    tabMunic.Close
    carComboLogra dec001.cmb_tipo_logra_conta
    carComboUF dec001.cmb_uf_conta
    carComboUF dec001.cmb_uf_crc_conta
    Screen.MousePointer = 0
End Sub
'*********************************************************************************************
Sub monta_combos_contrib()
'*********************************************************************************************
'
    Dim snativ As Recordset
    Dim ativ As String
      
    ativ = "SELECT * FROM tab_cod_munic " 'SELECIONA TUDO DA TABELA Municipios
    Set snativ = dbgim.OpenRecordset(ativ, dbOpenSnapshot)
    cmb_MUN.Clear 'Limpa o combo antes de jogar os dados
    Screen.MousePointer = 11
    Do Until snativ.EOF 'PARA LEITURA DE  todos os registros
'       cmb_MUN.AddItem Format$(snativ("Cod_Munic"), "00000") & "-" & snativ("Nome_Munic")   'adiciona o codigo e a descrição ao combo
        cmb_MUN.AddItem snativ("Nome_Munic") & Space(130) & Format$(snativ("Cod_Munic"), "00000")   'adiciona o codigo e a descrição ao combo
       snativ.MoveNext ' Lê o próximo registro
    Loop
    cmb_MUN.AddItem "MUNICIPIOS DE OUTROS ESTADOS" & Space(130) & "99999"
    snativ.Close
    carComboLogra dec001.cmb_tipo_logradouro
    cmb_relacao.Clear
'   cmb_relacao.AddItem "1 - Empregado"
'   cmb_relacao.AddItem "2 - Autônomo "
    cmb_relacao.AddItem "Empregado"
    cmb_relacao.AddItem "Autônomo "
    cmb_relacao.AddItem ""
    cmb_localizacao.Clear
'   cmb_localizacao.AddItem "1 - Na Empresa "
'   cmb_localizacao.AddItem "2 - No Contador"
    cmb_localizacao.AddItem "Na Empresa "
    cmb_localizacao.AddItem "No Contador"
    cmb_localizacao.ListIndex = 0
    cmb_tipo.Clear
'   cmb_tipo.AddItem "1 - Normal"
'   cmb_tipo.AddItem "2 - Pequeno Porte"
'   cmb_tipo.AddItem "3 - Microempresa"
    cmb_tipo.AddItem "Normal"
    cmb_tipo.AddItem "Microempresa"
    cmb_tipo.AddItem "Pequeno Porte"
    carComboUF dec001.CMB_UF
    '  carComboUF dec001.cmb_uf_crc
    Screen.MousePointer = 0
End Sub
'*********************************************************************************************
Sub pesquisa_contador()
'*********************************************************************************************
'
    query = "select * from tab_cad_contadores where tco_num_crc = '" & Trim(txt_num_crc_conta) & "'"
    query = query & " and tco_uf_crc = '" & Trim(cmb_uf_crc_conta.Text) & "'"
    Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
    If Not restauracontr.EOF And Not restauracontr.BOF Then
       data_contrib.RecordSource = query
       data_contrib.DatabaseName = App.Path & "\Dief.MDB"
       data_contrib.Refresh
       cmd_excluir_conta.Enabled = True
       cmd_imprime_conta.Enabled = True
      Call monta_tela_contador(data_contrib)
    End If
End Sub
'*********************************************************************************************
Sub pesquisa_contrib()
'*********************************************************************************************
'
    MyInscricao = msk_inscricao.Text
    MyInscricao = tirapictu(MyInscricao, ".")
    MyInscricao = tirapictu(MyInscricao, "-")
   
    query = "select * from tab_cad_contribuintes where tcc_nu_inscricao = '" & Trim(MyInscricao) & "'"
    Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
    If Not restauracontr.EOF And Not restauracontr.BOF Then
       data_contrib.RecordSource = query
       data_contrib.DatabaseName = App.Path & "\Dief.MDB"
       data_contrib.Refresh
       cmd_excluir_contr.Enabled = True
       cmd_imprime_contr.Enabled = True
      Call monta_tela_contribuinte(data_contrib)
    End If
End Sub
'*********************************************************************************************
Private Sub MontaListaAE()
'*********************************************************************************************
'
    Dim ret As String
    Dim reg As String
    Dim Classes As Recordset
    Dim classes1 As String
    Dim NumUltReg As Integer
    
    classes1 = "Select *  From AtividadeEconomica1"
    outAtEcon.Clear
    Set Classes = dbgim.OpenRecordset(classes1, dbOpenSnapshot)
    Do While Not Classes.EOF
       outAtEcon.AddItem Classes("IT_CO_ATIVIDADE") & "   " & Classes("IT_NO_ATIVIDADE")
       outAtEcon.Indent(outAtEcon.ListCount - 1) = 1
       Classes.MoveNext
    Loop
    NumUltReg = outAtEcon.ListCount - 1
 '  outAtEcon.RemoveItem NumUltReg
    outAtEcon.ListIndex = 0
    outAtEcon.Refresh
    Classes.Close
End Sub
'*********************************************************************************************
Sub monta_tab()
'*********************************************************************************************
'
   frame_contrib.Left = 40
   frame_contrib.Visible = True
   frame_contrib.Enabled = True
   Frame7.Left = 0
   Frame7.Visible = True
   Frame7.Enabled = True
End Sub
'*********************************************************************************************
Private Sub cmb_localizacao_Change()
'*********************************************************************************************
'
   teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_localizacao_Click()
'*********************************************************************************************
'
   teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_localizacao_GotFocus()
'*********************************************************************************************
'
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub cmb_localizacao_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
End Sub
'*********************************************************************************************
Private Sub cmb_localizacao_LostFocus()
'*********************************************************************************************
'
    If Mid(cmb_localizacao, 1, 1) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 0 Then
        reta = Crit_relacao(cmb_localizacao.ListIndex + 1)
        If Not reta Then
           MsgBox "Livro fiscal inválido.", 48, NOME_SISTEMA_DEFAULT
           cmb_localizacao.SetFocus
           Exit Sub
        End If
        cmd_salvar_contr.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub cmb_MUN_Change()
'*********************************************************************************************
'
    teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_MUN_Click()
'*********************************************************************************************
'
    CMB_UF.Enabled = True
    teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_mun_conta_Change()
'*********************************************************************************************
'
    teve_alteracao_contador = True
End Sub
'*********************************************************************************************
Private Sub cmb_mun_conta_Click()
'*********************************************************************************************
'
    teve_alteracao_contador = True
    cmb_uf_conta.Enabled = True
End Sub
'*********************************************************************************************

Private Sub cmb_mun_conta_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
End Sub
'*********************************************************************************************
Private Sub cmb_mun_conta_LostFocus()
'*********************************************************************************************
'
    If Trim(cmb_mun_conta) = "" Then
       Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        If cmb_mun_conta.ListIndex <> 216 Then
            cmb_uf_conta.ListIndex = 9
            cmb_uf_conta.Enabled = False
        Else
            cmb_uf_conta.Enabled = True
            cmb_uf_conta.Refresh
            cmb_uf_conta.SetFocus
        End If
        cmd_salvar_conta.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub cmb_MUN_GotFocus()
'*********************************************************************************************
'
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub cmb_MUN_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
End Sub
'*********************************************************************************************
Private Sub cmb_MUN_LostFocus()
'*********************************************************************************************
'
    If SSTab1.Tab = 0 Then
        If cmb_MUN.ListIndex <> 216 Then
            CMB_UF.ListIndex = 9
            CMB_UF.Enabled = False
        Else
            CMB_UF.Enabled = True
            CMB_UF.Refresh
            If SSTab1.Tab = 0 Then
               CMB_UF.SetFocus
            End If
        End If
       cmd_salvar_contr.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub cmb_relacao_Change()
'*********************************************************************************************
'
    teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_relacao_Click()
'*********************************************************************************************
'
    'teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_relacao_GotFocus()
    'outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub cmb_relacao_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    'Call PulaCampo(KeyAscii)
End Sub
'*********************************************************************************************
Private Sub cmb_relacao_LostFocus()
'*********************************************************************************************
'
    If Mid(cmb_relacao.Text, 1, 1) = "A" Then
        Exit Sub
    End If
    If SSTab1.Tab = 110 Then
        reta = Crit_relacao(cmb_relacao.ListIndex + 1)
        If Not reta Then
            MsgBox "Vínculo com Empresa inválido.", 48, NOME_SISTEMA_DEFAULT
            cmb_relacao.SetFocus
            Exit Sub
        End If
        cmd_salvar_contr.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_Change()
'*********************************************************************************************
'
    teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_Click()
'*********************************************************************************************
'
   teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_GotFocus()
'*********************************************************************************************
'
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_logra_conta_Change()
'*********************************************************************************************
'
    teve_alteracao_contador = True
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_logra_conta_Click()
'*********************************************************************************************
'
    teve_alteracao_contador = True
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_logra_conta_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
End Sub

Private Sub cmb_tipo_logra_conta_LostFocus()
'*********************************************************************************************
'
    If Trim(Right(cmb_tipo_logra_conta.Text, 3)) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        If Not critica_tipo_log(Trim(Right(cmb_tipo_logra_conta.Text, 3))) Then
            MsgBox "Tipo de Logradouro Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            If SSTab1.Tab = 1 Then
                cmb_tipo_logra_conta.SetFocus
            End If
            Exit Sub
        End If
            cmd_salvar_conta.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_logradouro_Change()
'*********************************************************************************************
'
    teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_logradouro_Click()
'*********************************************************************************************
'
    teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_logradouro_GotFocus()
'*********************************************************************************************
'
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_logradouro_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
End Sub
'*********************************************************************************************
Private Sub cmb_tipo_logradouro_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
       If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
          Exit Sub
       End If
    End If
    If Trim(Right(cmb_tipo_logradouro.Text, 3)) = "" Then
      Exit Sub
    End If
    If SSTab1.Tab = 0 Then
        If Not critica_tipo_log(Trim(Right(cmb_tipo_logradouro.Text, 3))) Then
           MsgBox "Tipo de Logradouro Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           If SSTab1.Tab = 0 Then
              cmb_tipo_logradouro.SetFocus
           End If
           Exit Sub
        End If
         cmd_salvar_contr.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub CMB_UF_Change()
'*********************************************************************************************
'
    teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub CMB_UF_Click()
'*********************************************************************************************
'
    teve_alteracao = True
End Sub
'*********************************************************************************************
Private Sub cmb_uf_conta_Change()
'*********************************************************************************************
'
    teve_alteracao_contador = True
End Sub
'*********************************************************************************************
Private Sub cmb_uf_conta_Click()
'*********************************************************************************************
'
    teve_alteracao_contador = True
End Sub
'*********************************************************************************************
Private Sub cmb_uf_conta_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
End Sub
'*********************************************************************************************
Private Sub cmb_uf_conta_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
           Exit Sub
        End If
    End If
 
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 1 Then
           Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is ComboBox Then
        If Me.ActiveControl.Tag = "municipio" Then
           Exit Sub
        End If
    End If
    retorno = ""
    If Trim(cmb_uf_conta.Text) = "" Then
       Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        retorno = CritUF(cmb_uf_conta.Text)
        If Len(retorno) > 0 Then
           MsgBox "UF inválida.", 48, NOME_SISTEMA_DEFAULT
           cmb_uf_conta.SetFocus
           Exit Sub
        End If
        If cmb_mun_conta.ListIndex = 216 And cmb_uf_conta.ListIndex = 9 Then
            cmb_uf_conta.Enabled = True
            cmb_uf_conta.Refresh
            MsgBox "UF Não pode ser Maranhão", 48, NOME_SISTEMA_DEFAULT
          '  cmb_MUN.ListIndex = -1
            If SSTab1.Tab = 1 Then
               cmb_uf_conta.SetFocus
            End If
            Exit Sub
        End If
       cmd_salvar_conta.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub cmb_uf_crc_conta_Change()
'*********************************************************************************************
'
    teve_alteracao_contador = True
End Sub
'*********************************************************************************************
Private Sub cmb_uf_crc_conta_Click()
'*********************************************************************************************
'
    teve_alteracao_contador = True
End Sub
'*********************************************************************************************
Private Sub cmb_uf_crc_conta_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
End Sub
'*********************************************************************************************
Private Sub cmb_uf_crc_conta_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is TextBox Then
        If Me.ActiveControl.Tag = "num_crc_conta" Then
            Exit Sub
        End If
    End If
    retorno = ""
    If Trim(cmb_uf_crc_conta.Text) = "" Then
       Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        retorno = CritUF(cmb_uf_crc_conta.Text)
        If Len(retorno) > 0 Then
           MsgBox "UF do CRC inválida.", 48, NOME_SISTEMA_DEFAULT
           If SSTab1.Tab = 1 Then
             cmb_uf_crc_conta.SetFocus
           End If
           Exit Sub
        End If
       ' cmb_uf_crc_conta.Text = UCase(cmb_uf_crc_conta.Text)
        pesquisa_contador
        '    txt_num_crc_conta.Enabled = False
        txt_num_crc_conta.Locked = True
        cmb_uf_crc_conta.Enabled = False
        cmd_salvar_conta.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub CMB_UF_GotFocus()
'*********************************************************************************************
'
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub CMB_UF_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
End Sub
'*********************************************************************************************
Private Sub CMB_UF_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 0 Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is ComboBox Then
        If Me.ActiveControl.Tag = "municipio" Then
            Exit Sub
        End If
    End If
    retorno = ""
    If Trim(CMB_UF.Text) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 0 Then
        retorno = CritUF(CMB_UF.Text)
        If Len(retorno) > 0 Then
           MsgBox "UF inválida.", 48, NOME_SISTEMA_DEFAULT
           CMB_UF.SetFocus
           Exit Sub
        End If
        If cmb_MUN.ListIndex = 216 And CMB_UF.ListIndex = 9 Then
            CMB_UF.Enabled = True
            CMB_UF.Refresh
            MsgBox "UF Não pode ser Maranhão.", 48, NOME_SISTEMA_DEFAULT
          '  cmb_MUN.ListIndex = -1
            If SSTab1.Tab = 0 Then
               CMB_UF.SetFocus
            End If
            Exit Sub
        End If
     cmd_salvar_contr.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub cmbCrc_GotFocus()
'*********************************************************************************************
'
    If cmb_tipo.Text <> "Microempresa" Then
        query = "select * from tab_cad_contadores "
        Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
        If restauracontr.EOF And restauracontr.BOF Then
            resposta = enviar_mensagem("Não existe Contador cadastrado. Cadastra agora?")
            If resposta = 1 Then
               ja_estou_indo_cadastrar = True
               SSTab1.Tab = 1
               txt_num_crc_conta.SetFocus
               vou_cadastrar_contador = True
               Exit Sub
            End If
        End If
    End If
End Sub
'*********************************************************************************************
Private Sub cmbCrc_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
'*********************************************************************************************
Private Sub cmbCrc_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 0 Then
            Exit Sub
        End If
    End If
    If ja_estou_indo_cadastrar = True Then
        ja_estou_indo_cadastrar = False
        Exit Sub
    End If
    If cmb_tipo.Text <> "Microempresa" Then
       If Trim(cmbCrc.Text) = "" Then
          resposta = enviar_mensagem("Quer cadastrar novo Contador?")
          If resposta = 1 Then
             SSTab1.Tab = 1
             txt_num_crc_conta.SetFocus
             vou_cadastrar_contador = True
             Exit Sub
          End If
       End If
    End If
    strSQLA = "Select * from tab_cad_contadores Where tco_num_crc='" & Trim(Right(cmbCrc, 8)) & "'"
    Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
    If PESQ_SNAP.RecordCount > 0 Then
        txt_num_crc = PESQ_SNAP!tco_num_crc
        Text2 = PESQ_SNAP!tco_uf_crc
    Else
     txt_num_crc = ""
        Text2 = ""
    End If
End Sub
'*********************************************************************************************
Private Sub cmd_buscar_conta_Click()
'*********************************************************************************************
'
On Error Resume Next
    BuscaContad = True
    busca_total_contador = True
    Me.Hide
   'Unload Me
    dec006.Show 1
    Me.Show  '1
    If Trim(g_nro_contador) = "" Then
        cmd_salvar_conta.Enabled = False
        cmd_excluir_conta.Enabled = False
        cmd_imprime_conta.Enabled = False
        '  BuscaContad = False
    Else
        teve_alteracao_contador = False
        cmd_salvar_conta.Enabled = True
        cmd_excluir_conta.Enabled = True
        cmd_imprime_conta.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub cmd_buscar_contr_Click()
 '*********************************************************************************************
 '
On Error Resume Next
    quem_me_chamou = ""
    BuscaContr = True
    Me.Hide
    dec005.Show 1
    Me.Show
 
    If Trim(g_nro_inscricao) = "" Then
        If teve_alteracao = True Then
            cmd_salvar_contr.Enabled = True
            cmd_excluir_contr.Enabled = True
            cmd_imprime_contr.Enabled = True
       Else
            cmd_salvar_contr.Enabled = False
            cmd_excluir_contr.Enabled = False
            cmd_imprime_contr.Enabled = False
            BuscaContr = False
       End If
    Else
       teve_alteracao = False
       cmd_salvar_contr.Enabled = True
       cmd_excluir_contr.Enabled = True
       cmd_imprime_contr.Enabled = True
    '  msk_inscricao.Enabled = False
       msk_inscricao.Locked = True
       txt_razao.SetFocus
    End If
End Sub
'*********************************************************************************************
Private Sub cmd_buscar_crc_Click()
'*********************************************************************************************
'
On Error Resume Next
    busca_total_contador = False
    Me.Hide
    dec006.Show 1
    Me.Show  '1
    If txt_num_crc <> "" Then
       cmbCrc.SetFocus
    End If
End Sub
'*********************************************************************************************
Private Sub cmd_excluir_conta_Click()
'*********************************************************************************************
'
On Error GoTo err_excluir_conta:
    query = "select * from tab_cad_contribuintes where tcc_num_crc_contador = '" & txt_num_crc_conta & "'"
    query = query & " and tcc_uf_crc_contador = '" & Trim(UCase(cmb_uf_crc_conta.Text)) & "'"
    Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
    If Not restauracontr.EOF And Not restauracontr.BOF Then
        MsgBox "Número de CRC com Contribuinte vinculado.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
        
        txt_nome.SetFocus
        Exit Sub
    End If
    resposta = enviar_mensagem("Confirma a exlusão?")
    If resposta = 1 Then
        query = "Delete from tab_cad_contadores"
        query = query & " where tco_num_crc = '" & txt_num_crc_conta & "'"
        query = query & " and tco_uf_crc = '" & Trim(UCase(cmb_uf_crc_conta.Text)) & "'"
        If Not EjecutarQuery(query) Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        cmd_salvar_conta.Enabled = False
        cmd_excluir_conta.Enabled = False
        cmd_imprime_conta.Enabled = False
        teve_alteracao_contador = False
        limpa_tela_contador
        '    txt_num_crc_conta.Enabled = True
        cmb_uf_crc_conta.Enabled = True
        txt_num_crc_conta.Locked = False
        txt_num_crc_conta.SetFocus
        Exit Sub
    End If
    txt_nome.SetFocus
    Exit Sub
err_excluir_conta:
 
    MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
End Sub
'*********************************************************************************************
Private Sub cmd_excluir_contr_Click()
'*********************************************************************************************
'
 On Error GoTo err_excluir:

    MyInscricao = msk_inscricao.Text
    MyInscricao = tirapictu(MyInscricao, ".")
    MyInscricao = tirapictu(MyInscricao, "-")
    query = "select * from tab_declaracao where tde_num_inscricao = '" & Trim(MyInscricao) & "'"
    Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
    If Not restauracontr.EOF And Not restauracontr.BOF Then
        resposta = enviar_mensagem("Este Contribuinte tem Declaraçao(ões) registradas. Confirma a exlusão ?")
        '    resposta = MsgBox("Este Contribuinte tem Declaraçao(ões) registradas. Confirma a exlusão ?", vbYesNo + vbDefaultButton2, NOME_SISTEMA_DEFAULT)
    Else
        resposta = enviar_mensagem("Confirma a exlusão?")
    End If
    If resposta = 1 Then
        query = "Delete from tab_declaracao"
        query = query & " where tde_num_inscricao = '" & Trim(MyInscricao) & "'"
        If Not EjecutarQuery(query) Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        query = "Delete from tab_detalhe_declaracao"
        query = query & " where tdd_num_inscricao = '" & Trim(MyInscricao) & "'"
        If Not EjecutarQuery(query) Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        query = "Delete from tab_cad_contribuintes"
        query = query & " where tcc_nu_inscricao = '" & Trim(MyInscricao) & "'"
        If Not EjecutarQuery(query) Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        cmd_salvar_contr.Enabled = False
        cmd_excluir_contr.Enabled = False
        cmd_imprime_contr.Enabled = False
        teve_alteracao = False
        limpa_tela
    End If
    'msk_inscricao.Enabled = True
    msk_inscricao.Locked = False
    msk_inscricao.SetFocus
    Exit Sub
err_excluir:
 
     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
End Sub
'*********************************************************************************************
Private Sub cmd_imprime_conta_Click()
'*********************************************************************************************
'
 On Error GoTo err_imprimir_conta:
 
    Screen.MousePointer = 11
    Imprime_contador cmb_uf_crc_conta.Text, txt_num_crc_conta
    Screen.MousePointer = 0
    Exit Sub
err_imprimir_conta:
 
     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
End Sub
'*********************************************************************************************
Private Sub cmd_imprime_contr_Click()
'*********************************************************************************************
'
 On Error GoTo err_imprimir:
  
    MyInscricao = msk_inscricao.Text
    MyInscricao = tirapictu(MyInscricao, ".")
    MyInscricao = tirapictu(MyInscricao, "-")
    Screen.MousePointer = 11
    IMprime_contrib MyInscricao
    Screen.MousePointer = 0
    Exit Sub
    
err_imprimir:
 
     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
End Sub
'*********************************************************************************************
Private Sub cmd_iniciar_conta_Click()
'*********************************************************************************************
'
On Error Resume Next

    If teve_alteracao_contador = True Then
        resposta = enviar_mensagem("Confirma a perda das alterações?")
        If resposta = 1 Then
            limpa_tela_contador
            txt_num_crc_conta.Locked = False
            cmb_uf_crc_conta.Enabled = True
            txt_num_crc_conta.SetFocus
            Exit Sub
        Else
            txt_nome.SetFocus
            Exit Sub
        End If
        teve_alteracao_contador = False
    Else
        limpa_tela_contador
        txt_num_crc_conta.Locked = False
        cmb_uf_crc_conta.Enabled = True
        txt_num_crc_conta.SetFocus
        Exit Sub
   End If
End Sub
'*********************************************************************************************
Private Sub cmd_iniciar_contr_Click()
'*********************************************************************************************
'
On Error Resume Next

    If teve_alteracao = True Then
        resposta = enviar_mensagem("Confirma a perda das alterações?")
        If resposta = 1 Then
            limpa_tela
            msk_inscricao.Locked = False
            msk_inscricao.SetFocus
        Else
            txt_razao.SetFocus
            Exit Sub
        End If
        teve_alteracao = False
   End If
   Monta_Combo_Crc
   limpa_tela
   CMB_UF.Enabled = True
   msk_inscricao.Locked = False
   msk_inscricao.SetFocus
End Sub
'*********************************************************************************************
Private Sub cmd_sair_conta_Click()
'*********************************************************************************************
'
On Error Resume Next

    If teve_alteracao_contador = True Then
        resposta = enviar_mensagem("Confirma a perda das alterações?")
        If resposta = 1 Then
            limpa_tela_contador
        Else
            txt_nome.SetFocus
            Exit Sub
        End If
        teve_alteracao_contador = False
    End If
Unload Me
End Sub
'*********************************************************************************************
Private Sub cmd_sair_contr_Click()
'*********************************************************************************************
'
On Error Resume Next

    If teve_alteracao = True Then
        resposta = enviar_mensagem("Confirma a perda das alterações?")
        If resposta = 1 Then
            limpa_tela
        Else
            txt_razao.SetFocus
            Exit Sub
        End If
        teve_alteracao = False
    End If
    Unload Me
End Sub
'*********************************************************************************************
Private Sub cmd_salvar_conta_Click()
'*********************************************************************************************
'
    Dim Campo As Integer
    On Error GoTo err_salvar_conta:
    Screen.MousePointer = 11
    If Not critica_contador(msg, Campo) Then
        Screen.MousePointer = 0
        If Campo = 1 Then
           MsgBox "Número de CRC Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           txt_num_crc_conta.Locked = False
           cmb_uf_crc_conta.Enabled = True
    
           txt_num_crc_conta.SetFocus
        Else
            resposta = MsgBox(msg & " Gravação cancelada.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT)
            Select Case Campo
                Case Is = 2
                  cmb_uf_crc_conta.SetFocus
                Case Is = 3
                  txt_nome.SetFocus
                Case Is = 4
                  msk_cpf.SetFocus
                Case Is = 5
                  cmb_tipo_logra_conta.SetFocus
                Case Is = 6
                  txt_log_conta.SetFocus
                Case Is = 7
                  cmb_mun_conta.SetFocus
                Case Is = 8
                  cmb_uf_conta.SetFocus
                Case Is = 9
                  MSKcep_conta.SetFocus
                Case Is = 17
                  mskddd_conta.SetFocus
                Case Is = 18
                  mskdddfax_conta.SetFocus
                Case Is = 20
                  txt_corr_eletr_conta.SetFocus
            End Select
            Exit Sub
         End If
    Else
         gravacao_ok = 0
         grava_contador
         If gravacao_ok = 0 Then
            MsgBox "Gravação efetuada com Sucesso.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            Screen.MousePointer = 0
            If vou_cadastrar_contador = True Then
               txt_num_crc = txt_num_crc_conta
               Text2 = cmb_uf_crc_conta
               limpa_tela_contador
               SSTab1.Tab = 0
               cmbCrc.SetFocus
               Exit Sub
            End If
            teve_alteracao_contador = False
            Exit Sub
        End If
    End If
    teve_alteracao_contador = False
    cmd_excluir_conta.Enabled = True
    cmd_imprime_conta.Enabled = True
    Screen.MousePointer = 0
    txt_num_crc_conta.Locked = False
    cmb_uf_crc_conta.Enabled = True
    txt_num_crc_conta.SetFocus
    Exit Sub
err_salvar_conta:
 
     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
End Sub
'*********************************************************************************************
Private Sub cmd_salvar_contr_Click()
'*********************************************************************************************
'
    Dim Campo As Integer
    On Error GoTo err_salvar:
    Screen.MousePointer = 11
    If Not critica_contrib(msg, Campo) Then
        Screen.MousePointer = 0
        If Campo = 1 Then
            MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            msk_inscricao.Locked = False
            msk_inscricao.SetFocus
         Else
            resposta = MsgBox(msg & " Gravação cancelada.", 48, NOME_SISTEMA_DEFAULT)
            Select Case Campo
                Case Is = 2
                    txt_razao.SetFocus
                Case Is = 3
                    msk_cgc.SetFocus
                Case Is = 4
                    cmb_tipo_logradouro.SetFocus
                Case Is = 5
                    txt_log.SetFocus
                Case Is = 6
                    txt_bairro.SetFocus
                Case Is = 7
                    cmb_MUN.SetFocus
                Case Is = 8
                    CMB_UF.SetFocus
                Case Is = 9
                    MSKcep.SetFocus
                Case Is = 10
                    cmdAtiEcon.SetFocus
                Case Is = 12
                    cmb_tipo.SetFocus
                Case Is = 13
                    cmbCrc.SetFocus
                Case Is = 16
                    cmb_relacao.SetFocus
                Case Is = 17
                    mskddd.SetFocus
                Case Is = 18
                    mskddd_fax.SetFocus
                Case Is = 19
                    cmbCrc.SetFocus
                Case Is = 20
                    txt_corr_eletronico.SetFocus
                Case Is = 21
                    Mask_dt_inicio_op.SetFocus
            End Select
            Exit Sub
         End If
    Else
        gravacao_ok = 0
        grava_contrib
        If gravacao_ok = 0 Then
            MsgBox "Gravação efetuada com Sucesso.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            teve_alteracao = False
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    teve_alteracao = False
    Monta_Combo_Crc
    cmd_excluir_contr.Enabled = True
    cmd_imprime_contr.Enabled = True
    Screen.MousePointer = 0
    msk_inscricao.Locked = False
    msk_inscricao.SetFocus
    Exit Sub
err_salvar:
 
    MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
End Sub
'*********************************************************************************************
Private Sub cmdAtiEcon_Click()
'*********************************************************************************************
'
    outAtEcon.Top = 2150
    outAtEcon.Height = 2800
    MontaListaAE
    outAtEcon.Visible = True
    outAtEcon.Refresh
    outAtEcon.SetFocus
    If Maria = True Then
       Maria = False
       outAtEcon.Visible = False
       Call cmdAtiEcon_Click
    End If
    outAtEcon.Visible = True
End Sub
'*********************************************************************************************
Private Sub cmdAtiEcon_KeyDown(KeyCode As Integer, Shift As Integer)
'*********************************************************************************************
'
    If (Shift = 4 And KeyCode = 40) Then
        cmdAtiEcon_Click
    ElseIf (Shift = 0 And KeyCode = 40) Then
        cmdAtiEcon_Click
    End If
End Sub
'*********************************************************************************************
Private Sub cmdAtiEcon_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
End Sub
'*********************************************************************************************
Private Sub Form_Activate()
    cmb_localizacao.ListIndex = 0
    Mask_dt_inicio_op.Text = "01011990"
    CentrarFormu Me
    If BuscaContr = True Then
        BuscaContr = False
        msk_inscricao.Locked = True
        txt_razao.SetFocus
        SSTab1.Tab = 0
        Maria = True
   ElseIf BuscaContad = True Then
        BuscaContad = False
        SSTab1.Tab = 1
        If Trim(g_nro_contador) <> "" Then
           cmb_uf_crc_conta.Enabled = False
           txt_num_crc_conta.Locked = True
           txt_nome.SetFocus
        Else
           If txt_num_crc_conta.Locked = False Then
              txt_num_crc_conta.SetFocus
           Else
               txt_nome.SetFocus
           End If
        End If
    Else
        SSTab1.Tab = 0
        If msk_inscricao.Locked = False Then
           msk_inscricao.SetFocus
        Else
           txt_razao.SetFocus
        End If
              Maria = True
        End If
End Sub
'*********************************************************************************************
Private Sub Form_Load()
'*********************************************************************************************
'
    Me.Top = 1200
    Me.Left = 0
    erro_geral = False
    vou_cadastrar_contador = False
    ja_estou_indo_cadastrar = False
    frame_contrib.Left = 40
    frame_contrib.Top = 360
    frame_contador.Visible = False
    frame_contador.Left = 40
    frame_contador.Top = 405
    monta_combos_contrib
    monta_combos_contador
    Monta_Combo_Crc
    cmd_buscar_contr.Enabled = True
    cmd_salvar_contr.Enabled = False
    cmd_excluir_contr.Enabled = False
    cmd_imprime_contr.Enabled = False
    cmd_buscar_conta.Enabled = True
    cmd_salvar_conta.Enabled = False
    cmd_excluir_conta.Enabled = False
    cmd_imprime_conta.Enabled = False
    MontaListaAE
    teve_alteracao = False  ' entrou e nao teve alteracao ainda
    teve_alteracao_contador = False
End Sub
'*********************************************************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*********************************************************************************************
'
    limpa_baloes
End Sub
'*********************************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*********************************************************************************************
'
    Unload Me
End Sub
'*********************************************************************************************
Private Sub frame_contrib_Click()
'*********************************************************************************************
'
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub Label1_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
'*********************************************************************************************
Private Sub Mask_dt_inicio_op_Change()
'*********************************************************************************************
'
    teve_alteracao = True
    If Len(Trim(Mask_dt_inicio_op)) > 8 Then
        SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub Mask_dt_inicio_op_GotFocus()
'*********************************************************************************************
'
    Mask_dt_inicio_op.SelStart = 0
    Mask_dt_inicio_op.SelLength = Len(Mask_dt_inicio_op)
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub Mask_dt_inicio_op_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub Mask_dt_inicio_op_LostFocus()
'*********************************************************************************************
'
    Dim data_micro As String
    Dim data_inicio As String
    Dim ano_wk As String
    
    If Trim(Mask_dt_inicio_op) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 0 Then
        reta = CRIT_Data(Mask_dt_inicio_op)
        If reta <> 0 Then
           MsgBox "Data Inválida.", 48, NOME_SISTEMA_DEFAULT
           Mask_dt_inicio_op = ""
           Mask_dt_inicio_op.SetFocus
           Exit Sub
        End If
        ano_wk = Val(Mid$(Mask_dt_inicio_op, 5, 4))
        data_inicio = Val(Mid$(Mask_dt_inicio_op, 5, 4) & Mid$(Mask_dt_inicio_op, 3, 2)) & Mid$(Mask_dt_inicio_op, 1, 2)
        data_micro = Val(Year(Date)) & Mid$(Date, 4, 2) & Mid$(Date, 1, 2)
        If ano_wk < 1800 Then
           MsgBox "Data Inválida.", 48, NOME_SISTEMA_DEFAULT
           Mask_dt_inicio_op = ""
           Mask_dt_inicio_op.SetFocus
           Exit Sub
        End If
        If data_inicio > data_micro Then
           MsgBox "Data de início de Atividades não pode ser maior que data atual.", 48, NOME_SISTEMA_DEFAULT
           Mask_dt_inicio_op = ""
           Mask_dt_inicio_op.SetFocus
           Exit Sub
        End If
        cmd_salvar_contr.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub msk_cgc_Change()
'*********************************************************************************************
'
    teve_alteracao = True
        If Len(Trim(msk_cgc)) > 14 Then
            SendKeys "{BackSpace}"
            Exit Sub
        End If
End Sub
'*********************************************************************************************
Private Sub msk_cgc_GotFocus()
'*********************************************************************************************
'
    msk_cgc.SelStart = 0
    msk_cgc.SelLength = Len(msk_cgc)
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub msk_cgc_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub msk_cgc_LostFocus()
'*********************************************************************************************
'
    Dim msg As String
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Then
           Exit Sub
        End If
    End If
    mycgc = msk_cgc.Text
    mycgc = tirapictu(mycgc, ".")
    mycgc = tirapictu(mycgc, "/")
    mycgc = tirapictu(mycgc, "-")
    If SSTab1.Tab = 0 Then
        If Trim(mycgc) <> "" Then
            If Len(mycgc) = 14 Then
                msk_cgc.Format = "@@.@@@.@@@/@@@@-@@"
                If Not Validar_Cgc_c(mycgc, msg) Then
                    msk_cgc.Text = ""
                    MsgBox msg, 48, NOME_SISTEMA_DEFAULT
                    If SSTab1.Tab = 0 Then
                      msk_cgc.SetFocus
                    End If
                    Exit Sub
                    msk_cgc.Format = Format(msk_cgc.Text, "@@.@@@.@@@/@@@@-@@")
                End If
            ElseIf Len(mycgc) = 11 Then
                reta = Rotcpf(mycgc)
                msk_cgc.Format = "@@@.@@@.@@@-@@"
                If reta = 0 Then
                    msk_cgc.Text = ""
                    MsgBox "Número de CPF Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                    If SSTab1.Tab = 1 Then
                       msk_cgc.SetFocus
                    End If
                    Exit Sub
                    msk_cgc.Format = Format(msk_cgc.Text, "@@@.@@@.@@@-@@")
               End If
            End If
        End If
        cmd_salvar_contr.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub msk_cpf_Change()
'*********************************************************************************************
'
    teve_alteracao_contador = True
    If Len(Trim(msk_cpf)) > 11 Then
        SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub Msk_cpf_esp_Change()
'*********************************************************************************************
'
    teve_alteracao_contador = True
    If Len(Trim(Msk_cpf_esp)) > 11 Then
        SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub Msk_cpf_esp_GotFocus()
'*********************************************************************************************
'
    Msk_cpf_esp.SelStart = 0
    Msk_cpf_esp.SelLength = Len(Msk_cpf_esp)
End Sub
'*********************************************************************************************
Private Sub Msk_cpf_esp_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Then
        If Len(Trim(Msk_cpf_esp)) = 0 Then
           MsgBox "Número do CPF Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           Msk_cpf_esp = ""
           Msk_cpf_esp.SetFocus
           Exit Sub
        End If
    End If
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub Msk_cpf_esp_LostFocus()
'*********************************************************************************************
'
    If Trim(Msk_cpf_esp) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        If Val(Trim(Msk_cpf_esp)) = 0 Then
           MsgBox "Número do CPF Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           Msk_cpf_esp = ""
           Msk_cpf_esp.SetFocus
           Exit Sub
        End If
        reta = Rotcpf(Msk_cpf_esp)
        If reta = 0 Then
           Msk_cpf_esp = ""
           MsgBox "Número de CPF Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           If SSTab1.Tab = 1 Then
              Msk_cpf_esp.SetFocus
           End If
           Exit Sub
        End If
        cmd_salvar_conta.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub msk_cpf_GotFocus()
'*********************************************************************************************
'
    msk_cpf.SelStart = 0
    msk_cpf.SelLength = Len(msk_cpf)
End Sub
'*********************************************************************************************
Private Sub msk_cpf_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Then
        If Val(Trim(msk_cpf)) = 0 Then
           MsgBox "Número do CPF Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           msk_cpf = ""
           msk_cpf.SetFocus
           Exit Sub
        End If
    End If
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub msk_cpf_LostFocus()
'*********************************************************************************************
'
    If Trim(msk_cpf) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        If Val(Trim(msk_cpf)) = 0 Then
           MsgBox "Número do CPF Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           msk_cpf = ""
           msk_cpf.SetFocus
           Exit Sub
        End If
        reta = Rotcpf(msk_cpf)
        If reta = 0 Then
           msk_cpf = ""
           MsgBox "Número de CPF Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           If SSTab1.Tab = 1 Then
              msk_cpf.SetFocus
           End If
           Exit Sub
        End If
         cmd_salvar_conta.Enabled = True
End If
End Sub
'*********************************************************************************************
Private Sub msk_inscricao_Change()
'*********************************************************************************************
'
    teve_alteracao = True
    If Len(Trim(msk_inscricao)) > 9 Then
    '   SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub msk_inscricao_GotFocus()
'*********************************************************************************************
'
    msk_inscricao.SelStart = 0
    msk_inscricao.SelLength = Len(msk_inscricao)
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub msk_inscricao_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Then
        If Len(Trim(msk_inscricao)) < 9 Then
            MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            msk_inscricao = ""
            msk_inscricao.SetFocus
            Exit Sub
        End If
     End If
     Call PulaCampo(KeyAscii)
     If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub msk_inscricao_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 0 Then
           Exit Sub
        End If
    End If
    MyInscricao = msk_inscricao.Text
    MyInscricao = tirapictu(MyInscricao, ".")
    MyInscricao = tirapictu(MyInscricao, "-")
 '  If flag_entrando_tab = 0 Then     'entrou na tab do contribuinte
'       flag_entrando_tab = 9   ' entrei e nao critiquei o contribuinte
 '     Exit Sub
 '  End If
    If SSTab1.Tab = 0 Then
        If Trim(MyInscricao) = "" Then
           MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           msk_inscricao.SetFocus
           Exit Sub
        End If
        If Len(Trim(msk_inscricao)) < 9 Then
           MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           msk_inscricao.SetFocus
           Exit Sub
        End If
        If Val(Mid(Trim(MyInscricao), 1, 2)) <> 12 Then
           MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           msk_inscricao.SetFocus
           Exit Sub
        End If
        If GenDV(CStr(Mid(MyInscricao, 1, 8)), BASE9) <> Val(Mid(MyInscricao, 9, 1)) Then
           MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           msk_inscricao.SetFocus
           Exit Sub
        End If
        If Len(msk_inscricao) < 11 Then
            msk_inscricao = Format(msk_inscricao, "@@.@@@.@@@-@")
        End If
        msk_inscricao.Locked = True
        pesquisa_contrib
        '  msk_inscricao.Enabled = False
        cmd_salvar_contr.Enabled = True
   End If
End Sub
'*********************************************************************************************
Private Sub MSKcep_Change()
'*********************************************************************************************
'
    teve_alteracao = True
    If Len(Trim(MSKcep)) > 8 Then
       SendKeys "{BackSpace}"
       Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub MSKcep_conta_Change()
'*********************************************************************************************
'
    teve_alteracao_contador = True
    If Len(Trim(MSKcep_conta)) > 8 Then
        '  SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub MSKcep_conta_GotFocus()
'*********************************************************************************************
'
   MSKcep_conta.SelStart = 0
   MSKcep_conta.SelLength = Len(MSKcep_conta)
End Sub
'*********************************************************************************************
Private Sub MSKcep_conta_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Trim(MSKcep_conta) <> "" Then
            If Len(Trim(MSKcep_conta)) < 8 Then
                MsgBox "Número do CEP Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                MSKcep_conta = ""
                MSKcep_conta.SetFocus
                Exit Sub
            End If
            If Val(Trim(MSKcep_conta)) = 0 Then
                MsgBox "Número do CEP Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                MSKcep_conta = ""
                MSKcep_conta.SetFocus
                Exit Sub
            End If
        End If
    End If
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub MSKcep_conta_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 1 Then
            Exit Sub
        End If
    End If
    If Trim(MSKcep_conta) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        If Len(Trim(MSKcep_conta)) < 8 Then
            MsgBox "Número do CEP Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            MSKcep_conta = ""
            MSKcep_conta.SetFocus
            Exit Sub
        End If
        If Val(Trim(MSKcep_conta)) = 0 Then
            MsgBox "Número do CEP Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            MSKcep_conta = ""
            MSKcep_conta.SetFocus
            Exit Sub
        End If
    End If
    cmd_salvar_conta.Enabled = True
End Sub
'*********************************************************************************************
Private Sub MSKcep_GotFocus()
'*********************************************************************************************
'
    MSKcep.SelStart = 0
    MSKcep.SelLength = Len(MSKcep)
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub MSKcep_KeyPress(KeyAscii As Integer)
 '*********************************************************************************************
 '
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Trim(MSKcep) <> "" Then
            If Len(Trim(MSKcep)) < 8 Then
                MsgBox "Número do CEP Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                MSKcep = ""
                MSKcep.SetFocus
                Exit Sub
            End If
        End If
        If Val(Trim(MSKcep)) = 0 Then
            MsgBox "Número do CEP Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            MSKcep = ""
            MSKcep.SetFocus
            Exit Sub
        End If
    
    End If
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub MSKcep_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 0 Then
             Exit Sub
        End If
    End If
    If Trim(MSKcep) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 0 Then
        If Len(Trim(MSKcep)) < 8 Then
            MsgBox "Número do CEP Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            MSKcep = ""
            MSKcep.SetFocus
            Exit Sub
        End If
        If Val(Trim(MSKcep)) = 0 Then
            MsgBox "Número do CEP Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            MSKcep = ""
            MSKcep.SetFocus
            Exit Sub
        End If
    End If
    cmd_salvar_contr.Enabled = True
End Sub
'*********************************************************************************************
Private Sub mskddd_Change()
'*********************************************************************************************
'
    teve_alteracao = True
    If Len(Trim(mskddd)) > 5 Then
        SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub mskddd_conta_Change()
'*********************************************************************************************
'
   teve_alteracao_contador = True
   If Len(Trim(mskddd_conta)) > 5 Then
      SendKeys "{BackSpace}"
      Exit Sub
   End If
End Sub
'*********************************************************************************************
Private Sub mskddd_conta_GotFocus()
'*********************************************************************************************
'
   mskddd_conta.SelStart = 0
   mskddd_conta.SelLength = Len(mskddd_conta)
End Sub
'*********************************************************************************************
Private Sub mskddd_conta_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Trim(mskddd_conta) = "" Then
            Call PulaCampo(KeyAscii)
            Exit Sub
        End If
        If Len(mskddd_conta) < 3 Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd_conta = ""
            mskddd_conta.SetFocus
            erro_geral = True
            Exit Sub
        End If
        If cmb_mun_conta.ListIndex = 216 And Trim(mskddd_conta) = "098" Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd_conta = ""
            mskddd_conta.SetFocus
            erro_geral = True
            Exit Sub
        End If
    End If
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub mskddd_conta_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
      If SSTab1.Tab = 1 Then
         Exit Sub
      End If
    End If
    If Trim(mskddd_conta) = "" Then
      Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        If Len(mskddd_conta) < 3 Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd_conta = ""
            mskddd_conta.SetFocus
            erro_geral = True
            Exit Sub
        End If
        If cmb_mun_conta.ListIndex = 216 And Trim(mskddd_conta) = "098" Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd_conta = ""
            mskddd_conta.SetFocus
            erro_geral = True
            Exit Sub
        End If
    End If
erro_geral = False
    cmd_salvar_conta.Enabled = True
End Sub
'*********************************************************************************************
Private Sub mskddd_fax_Change()
'*********************************************************************************************
'
    teve_alteracao = True
        If Len(Trim(mskddd_fax)) > 5 Then
            SendKeys "{BackSpace}"
            Exit Sub
        End If
End Sub
'*********************************************************************************************
Private Sub mskddd_fax_GotFocus()
'*********************************************************************************************
'
    mskddd_fax.SelStart = 0
    mskddd_fax.SelLength = Len(mskddd_fax)
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub mskddd_fax_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Trim(mskddd_fax) = "" Then
            Call PulaCampo(KeyAscii)
            Exit Sub
        End If
        If Len(mskddd_fax) < 3 Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd_fax = ""
            mskddd_fax.SetFocus
            erro_geral = True
            Exit Sub
        End If
        If cmb_MUN.ListIndex = 216 And Trim(mskddd_fax) = "098" Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd_fax = ""
            mskddd_fax.SetFocus
            erro_geral = True
            Exit Sub
        End If
    End If
    Call PulaCampo(KeyAscii)
 
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub mskddd_fax_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 0 Then
            Exit Sub
        End If
    End If
    If Trim(mskddd_fax) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 0 Then
        If Len(mskddd_fax) < 3 Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd_fax = ""
            mskddd_fax.SetFocus
            erro_geral = True
            Exit Sub
        End If
        If cmb_MUN.ListIndex = 216 And Trim(mskddd_fax) = "098" Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd_fax = ""
            mskddd_fax.SetFocus
            erro_geral = True
            Exit Sub
        End If
    End If
    erro_geral = False
    cmd_salvar_contr.Enabled = True
End Sub
'*********************************************************************************************
Private Sub mskddd_GotFocus()
'*********************************************************************************************
'
    mskddd.SelStart = 0
    mskddd.SelLength = Len(mskddd)
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub mskddd_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Trim(mskddd) = "" Then
            Call PulaCampo(KeyAscii)
            Exit Sub
        End If
        If Len(mskddd) < 3 Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd = ""
            mskddd.SetFocus
            erro_geral = True
            Exit Sub
        End If
        If cmb_MUN.ListIndex = 216 And Trim(mskddd) = "098" Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd = ""
            mskddd.SetFocus
            erro_geral = True
            Exit Sub
        End If
    End If
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub mskddd_LostFocus()
'*********************************************************************************************
'

    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 0 Then
            Exit Sub
        End If
    End If
    If Trim(mskddd) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 0 Then
        If Len(mskddd) < 3 Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd = ""
            mskddd.SetFocus
            erro_geral = True
            Exit Sub
        End If
        If cmb_MUN.ListIndex = 216 And Trim(mskddd) = "098" Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskddd = ""
            mskddd.SetFocus
            erro_geral = True
            Exit Sub
        End If
    End If
    erro_geral = False
    cmd_salvar_contr.Enabled = True
End Sub
'*********************************************************************************************
Private Sub mskdddfax_conta_Change()
'*********************************************************************************************
'
    teve_alteracao_contador = True

    If Len(Trim(mskdddfax_conta)) > 5 Then
        SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub mskdddfax_conta_GotFocus()
'*********************************************************************************************
'
    mskdddfax_conta.SelStart = 0
    mskdddfax_conta.SelLength = Len(mskdddfax_conta)
End Sub
'*********************************************************************************************
Private Sub mskdddfax_conta_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Trim(mskdddfax_conta) = "" Then
            Call PulaCampo(KeyAscii)
            Exit Sub
        End If
        If Len(mskdddfax_conta) < 3 Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskdddfax_conta = ""
            mskdddfax_conta.SetFocus
            erro_geral = True
            Exit Sub
        End If
        If cmb_mun_conta.ListIndex = 216 And Trim(mskdddfax_conta) = "098" Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskdddfax_conta = ""
            mskdddfax_conta.SetFocus
            erro_geral = True
            Exit Sub
        End If
    End If
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub mskdddfax_conta_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 1 Then
            Exit Sub
        End If
    End If
    If Trim(mskdddfax_conta) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        If Len(mskdddfax_conta) < 3 Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskdddfax_conta = ""
            mskdddfax_conta.SetFocus
            erro_geral = True
            Exit Sub
        End If
        If cmb_mun_conta.ListIndex = 216 And Trim(mskdddfax_conta) = "098" Then
            MsgBox "Número de DDD Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskdddfax_conta = ""
            mskdddfax_conta.SetFocus
            erro_geral = True
            Exit Sub
        End If
    End If
    erro_geral = False
    cmd_salvar_conta.Enabled = True
End Sub
'*********************************************************************************************
Private Sub mskdtinicio_Change()
'*********************************************************************************************
'
    teve_alteracao = True
    If Len(Trim(mskdtinicio)) > 8 Then
        SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub mskdtinicio_GotFocus()
'*********************************************************************************************
'
    mskdtinicio.SelStart = 0
    mskdtinicio.SelLength = Len(mskdtinicio)
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub mskdtinicio_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub mskdtinicio_LostFocus()
'*********************************************************************************************
'
    Dim data_micro As String
    Dim data_inicio As String
    Dim ano_wk As String
    Dim data_inicio_op As String
    
    If Trim(mskdtinicio) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 0 Then
        reta = CRIT_Data(mskdtinicio)
        If reta <> 0 Then
           MsgBox "Data Inválida.", 48, NOME_SISTEMA_DEFAULT
           mskdtinicio = ""
           mskdtinicio.SetFocus
           Exit Sub
        End If
        ano_wk = Val(Mid$(mskdtinicio, 5, 4))
        data_inicio = Val(Mid$(mskdtinicio, 5, 4) & Mid$(mskdtinicio, 3, 2)) & Mid$(mskdtinicio, 1, 2)
        data_inicio_op = Val(Mid$(Mask_dt_inicio_op, 5, 4) & Mid$(Mask_dt_inicio_op, 3, 2)) & Mid$(Mask_dt_inicio_op, 1, 2)
        data_micro = Val(Year(Date)) & Mid$(Date, 4, 2) & Mid$(Date, 1, 2)
        If ano_wk < 1800 Then
           MsgBox "Data Inválida.", 48, NOME_SISTEMA_DEFAULT
           mskdtinicio = ""
           mskdtinicio.SetFocus
           Exit Sub
        End If
        If data_inicio > data_micro Then
           MsgBox "Data de início não pode ser maior que data atual.", 48, NOME_SISTEMA_DEFAULT
           mskdtinicio = ""
           mskdtinicio.SetFocus
           Exit Sub
        End If
        If data_inicio < data_inicio_op Then
           MsgBox "Data de início não pode ser menor que data inicio de atividades.", 48, NOME_SISTEMA_DEFAULT
           mskdtinicio = ""
           mskdtinicio.SetFocus
           Exit Sub
        End If
        cmd_salvar_contr.Enabled = True
    End If
End Sub
'*********************************************************************************************
Private Sub mskfax_Change()
'*********************************************************************************************
'
    teve_alteracao = True
    If Len(Trim(mskfax)) > 10 Then
        SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub mskfax_conta_Change()
'*********************************************************************************************
'
    teve_alteracao_contador = True
    If Len(Trim(mskfax_conta)) > 10 Then
        SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub mskfax_conta_GotFocus()
'*********************************************************************************************
'
   mskfax_conta.SelStart = 0
   mskfax_conta.SelLength = Len(mskfax_conta)
End Sub
'*********************************************************************************************
Private Sub mskfax_conta_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Trim(mskfax_conta) = "" Then
            Call PulaCampo(KeyAscii)
            Exit Sub
        End If
        If Len(mskfax_conta) < 6 Then
           MsgBox "Número de Fax Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           mskfax_conta = ""
           mskfax_conta.SetFocus
           Exit Sub
        End If
    End If
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub mskfax_conta_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
           Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 1 Then
            Exit Sub
        End If
    End If
    If erro_geral = True Then Exit Sub
    If Trim(mskfax_conta) = "" Then
        Exit Sub
    End If
    If Trim(mskfax_conta) = "" And Trim(mskdddfax_conta) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        If Trim(mskfax_conta) <> "" Then
            If Len(mskfax_conta) < 6 Then
               MsgBox "Número de Fax Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
               mskfax_conta = ""
               mskfax_conta.SetFocus
               Exit Sub
            End If
        End If
        ret_tel = 0
        ret_tel = critica_telefone(Trim(mskdddfax_conta), Trim(mskfax_conta))
        If ret_tel <> 0 Then
            If ret_tel = 1 Then
                mensagem = "DDD informado e Fax não informado."
                MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                mskfax_conta.SetFocus
                Exit Sub
            ElseIf ret_tel = 2 Then
                If cmb_uf_conta.ListIndex = 9 Then
                    mskdddfax_conta.Text = "098"
                Else
                    mensagem = "Fax informado e DDD não informado."
                    MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                    mskdddfax_conta.SetFocus
                    Exit Sub
                End If
            Else
                mensagem = "DDD Inválido."
                MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                mskdddfax_conta.SetFocus
                Exit Sub
            End If
        End If
    End If
    cmd_salvar_conta.Enabled = True
End Sub
'*********************************************************************************************
Private Sub mskfax_GotFocus()
'*********************************************************************************************
'
    mskfax.SelStart = 0
    mskfax.SelLength = Len(mskfax)
    outAtEcon.Visible = False
End Sub
'*********************************************************************************************
Private Sub mskfax_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Trim(mskfax) = "" Then
            Call PulaCampo(KeyAscii)
            Exit Sub
        End If
        If Len(mskfax) < 6 Then
            MsgBox "Número de Fax Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            mskfax = ""
            mskfax.SetFocus
            Exit Sub
        End If
    End If
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub mskfax_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
       If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
          Exit Sub
       End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
      If SSTab1.Tab = 0 Then
         Exit Sub
      End If
   End If
   
    If erro_geral = True Then Exit Sub
    If Trim(mskfax) = "" Then
        Exit Sub
    End If
    If Trim(mskfax) = "" And Trim(mskddd_fax) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 0 Then
        If Trim(mskfax) <> "" Then
            If Len(mskfax) < 6 Then
                MsgBox "Número de Fax Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                mskfax = ""
                mskfax.SetFocus
                Exit Sub
            End If
        End If
        ret_tel = 0
        ret_tel = critica_telefone(Trim(mskddd_fax), Trim(mskfax))
        If ret_tel <> 0 Then
            If ret_tel = 1 Then
                mensagem = "DDD informado e Fax não informado."
                MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                mskfax.SetFocus
                Exit Sub
            ElseIf ret_tel = 2 Then
                If CMB_UF.ListIndex = 9 Then
                    mskddd_fax.Text = "098"
                Else
                    mensagem = "Fax informado e DDD não informado."
                    MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                    mskddd_fax.SetFocus
                    Exit Sub
                End If
            Else
                mensagem = "DDD Inválido."
                MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                mskddd_fax.SetFocus
                Exit Sub
            End If
        End If
    End If
    cmd_salvar_contr.Enabled = True
End Sub
'*********************************************************************************************
Private Sub msktel_Change()
'*********************************************************************************************
'
    teve_alteracao = True
    If Len(Trim(msktel)) > 10 Then
        SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub msktel_conta_Change()
'*********************************************************************************************
'
    teve_alteracao_contador = True
    If Len(Trim(msktel_conta)) > 10 Then
        SendKeys "{BackSpace}"
        Exit Sub
    End If
End Sub
'*********************************************************************************************
Private Sub msktel_conta_GotFocus()
'*********************************************************************************************
'
    msktel_conta.SelStart = 0
    msktel_conta.SelLength = Len(msktel_conta)
End Sub
'*********************************************************************************************
Private Sub msktel_conta_KeyPress(KeyAscii As Integer)
'*********************************************************************************************
'
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Trim(msktel_conta) = "" Then
            Call PulaCampo(KeyAscii)
            Exit Sub
        End If
        If Len(msktel_conta) < 6 Then
            MsgBox "Número de Telefone Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            msktel_conta = ""
            msktel_conta.SetFocus
            Exit Sub
        End If
    End If
    Call PulaCampo(KeyAscii)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
End Sub
'*********************************************************************************************
Private Sub msktel_conta_LostFocus()
'*********************************************************************************************
'
    If TypeOf Me.ActiveControl Is SSCommand Then
        If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
            Exit Sub
        End If
    End If
    If TypeOf Me.ActiveControl Is SSTab Then
        If SSTab1.Tab = 1 Then
           Exit Sub
        End If
    End If
    If erro_geral = True Then Exit Sub
    If Trim(msktel_conta) = "" Then
        Exit Sub
    End If
    If Trim(msktel_conta) = "" And Trim(mskddd_conta) = "" Then
        Exit Sub
    End If
    If SSTab1.Tab = 1 Then
        If Trim(msktel_conta) <> "" Then
            If Len(msktel_conta) < 6 Then
                MsgBox "Número de Telefone Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                msktel_conta = ""
                msktel_conta.SetFocus
                Exit Sub
            End If
        End If
        ret_tel = 0
        ret_tel = critica_telefone(Trim(mskddd_conta), Trim(msktel_conta))
        If ret_tel <> 0 Then
            If ret_tel = 1 Then
                mensagem = "DDD informado e Telefone não informado."
                MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                msktel_conta.SetFocus
                Exit Sub
            ElseIf ret_tel = 2 Then
                If cmb_uf_conta.ListIndex = 9 Then
                   mskddd_conta.Text = "098"
                Else
                     mensagem = "Telefone informado e DDD não informado."
                     MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                     mskddd_conta.SetFocus
                     Exit Sub
                End If
            Else
                mensagem = "DDD Inválido."
                MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                mskddd_conta.SetFocus
                Exit Sub
          End If
        End If
    End If
    erro_geral = False
    cmd_salvar_conta.Enabled = True
End Sub
'*********************************************************************************************
                                        'PAREI BEM AQUI

Private Sub msktel_GotFocus()
   msktel.SelStart = 0
   msktel.SelLength = Len(msktel)
outAtEcon.Visible = False
End Sub


Private Sub msktel_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Or KeyAscii = 9 Then
    If Trim(msktel) = "" Then
       Call PulaCampo(KeyAscii)
      Exit Sub
    End If
    If Len(msktel) < 6 Then
       MsgBox "Número de Telefone Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       msktel = ""
       msktel.SetFocus
       Exit Sub
    End If
    
End If
 
 Call PulaCampo(KeyAscii)
 
  If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep

End Sub


Private Sub msktel_LostFocus()

If TypeOf Me.ActiveControl Is SSCommand Then
      If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
         Exit Sub
      End If
   End If
   
    If TypeOf Me.ActiveControl Is SSTab Then
      If SSTab1.Tab = 0 Then
         Exit Sub
      End If
   End If
   
  If erro_geral = True Then Exit Sub

  If Trim(msktel) = "" Then
      Exit Sub
   End If
   If Trim(msktel) = "" And Trim(mskddd) = "" Then
      Exit Sub
   End If
  
 If SSTab1.Tab = 0 Then
     If Trim(msktel) <> "" Then
       If Len(msktel) < 6 Then
          MsgBox "Número de Telefone Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
          msktel = ""
          msktel.SetFocus
          Exit Sub
       End If
     End If
    ret_tel = 0
    ret_tel = critica_telefone(Trim(mskddd), Trim(msktel))
    If ret_tel <> 0 Then
         If ret_tel = 1 Then
            mensagem = "DDD informado e Telefone não informado."
            MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            msktel.SetFocus
            Exit Sub
          ElseIf ret_tel = 2 Then
            If CMB_UF.ListIndex = 9 Then
               mskddd.Text = "098"
            Else
                 mensagem = "Telefone informado e DDD não informado."
                 MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                 mskddd.SetFocus
                Exit Sub
            End If
        Else
           mensagem = "DDD Inválido."
           MsgBox mensagem, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           mskddd.SetFocus
           Exit Sub
            
          End If
    End If
   
End If
 cmd_salvar_contr.Enabled = True

End Sub


Private Sub outAtEcon_Collapse(ListIndex As Integer)

 outAtEcon.ListIndex = ListIndex
 
End Sub


Private Sub outAtEcon_DblClick()
    
 Static Indent1 As Integer, Indent2 As Integer

    If Not outAtEcon.Expand(outAtEcon.ListIndex) Then
        If outAtEcon.Indent(outAtEcon.ListIndex) = 1 Then
            Indent1 = outAtEcon.ListIndex
            If Indent_1() Then
                outAtEcon.ListIndex = 1
            Else
                cmdAtiEcon.SetFocus
                Exit Sub
            End If
        Else
           If outAtEcon.Indent(outAtEcon.ListIndex) = 2 Then
             If so_tenho_uma_atividade < 3 Then
                Indent_3
                teve_alteracao = True
                outAtEcon.Clear
                outAtEcon.Visible = False
                out3 = True
                cmb_tipo.SetFocus
             Else
                 Indent2 = outAtEcon.ListIndex
                 If Indent_2() Then
                     outAtEcon.ListIndex = 2
                 Else
                     cmdAtiEcon.SetFocus
                     Exit Sub
                 End If
              End If
           Else
               If outAtEcon.Indent(outAtEcon.ListIndex) = 3 Then
                    Indent_3
                    teve_alteracao = True
                    outAtEcon.Clear
                    outAtEcon.Visible = False
                    out3 = True
                    cmb_tipo.SetFocus

               End If
          End If
      End If
    ElseIf outAtEcon.Expand(outAtEcon.ListIndex) Then
        If outAtEcon.ListIndex = 1 Then
            MontaListaAE
            outAtEcon.ListIndex = Indent1
            If Indent_1() Then
                outAtEcon.ListIndex = Indent2
            Else
                cmdAtiEcon.SetFocus
                Exit Sub
            End If
        ElseIf outAtEcon.ListIndex = 0 Then
            MontaListaAE
            outAtEcon.ListIndex = 0
        End If
    End If

End Sub


Private Sub outAtEcon_Expand(ListIndex As Integer)

    outAtEcon.ListIndex = ListIndex

End Sub


Private Sub outAtEcon_GotFocus()
    
    Joao = True
so_tenho_uma_atividade = 9
    outAtEcon.Top = 2150
    outAtEcon.Height = 2800
    outAtEcon.Visible = True
    

End Sub


Private Sub outAtEcon_KeyPress(KeyAscii As Integer)
    
  
    If KeyAscii = 27 Then  ' INCLUI EM 4/3/98
        Joao = True
     
      cmb_tipo.SetFocus
         outAtEcon.Visible = False
        'Label1.Text = ""
        Exit Sub
    End If
    
    
    If KeyAscii = 13 Then
        If outAtEcon.ListIndex <> -1 Then
           Call outAtEcon_DblClick
        End If
    ElseIf KeyAscii = 27 Then
        cmdAtiEcon.SetFocus
    End If

End Sub


Private Sub outAtEcon_LostFocus()
    

    
If Joao = True Then
   Joao = False
   Exit Sub
End If


    If outAtEcon.ListIndex <> -1 Then
        If outAtEcon.Indent(outAtEcon.ListIndex) = 3 Then
            If Len(outAtEcon.List(outAtEcon.ListIndex)) > 90 Then
                Label1.Text = " " + Mid$(outAtEcon.List(outAtEcon.ListIndex), 9, 90)
            Else
                Label1.Text = " " + Mid$(outAtEcon.List(outAtEcon.ListIndex), 9, Len(outAtEcon.List(outAtEcon.ListIndex)) - 8)
            End If
            Label1.Refresh
            outAtEcon.Visible = False
            teve_alteracao = True
            
        End If
    End If

    outAtEcon.Visible = False
    

End Sub









Private Sub SSTab1_Click(PreviousTab As Integer)

outAtEcon.Visible = False

If SSTab1.Tab = 0 Then

   Monta_Combo_crc_click
    
   flag_entrando_tab = 0   'entrou na tab do contribuinte
   frame_contador.Visible = False
   frame_contador.Enabled = False
  frame_contrib.Left = 40
   frame_contrib.Top = 360
   frame_contrib.Visible = True
   frame_contrib.Enabled = True
    Frame7.Visible = True
   Frame7.Enabled = True
   SSTab1.HelpContextID = 2
Else
    flag_entrando_tab = 1     'entrou na tab do contador
    frame_contrib.Visible = False
    frame_contrib.Enabled = False
   frame_contador.Left = 40
    frame_contador.Top = 405
    frame_contador.Visible = True
    frame_contador.Enabled = True
       SSTab1.HelpContextID = 7
End If
Dim tempo As Integer
For tempo = 1 To 2000
Next tempo

teve_alteracao = False  ' entrou e nao teve alteracao ainda
teve_alteracao_contador = False
If SSTab1.Tab = 0 Then
    If msk_inscricao.Locked = False Then
      msk_inscricao.SetFocus
    Else
      txt_razao.SetFocus
    End If
Else
   If txt_num_crc_conta.Locked = False Then
       txt_num_crc_conta.SetFocus
   Else
       txt_nome.SetFocus
   End If
End If
  
   
End Sub

Private Sub SSTab1_GotFocus()
outAtEcon.Visible = False
'
If SSTab1.Tab = 0 Then
  If vou_cadastrar_contador = True Then
     vou_cadastrar_contador = False
     cmbCrc.SetFocus
     Exit Sub
  End If
   flag_entrando_tab = 0     'entrou na tab do contribuinte
   frame_contador.Visible = False
   frame_contador.Enabled = False
  frame_contrib.Left = 40
   frame_contrib.Visible = True
   frame_contrib.Enabled = True
    Frame7.Visible = True
   Frame7.Enabled = True
Else
    flag_entrando_tab = 1     'entrou na tab do contador
    frame_contrib.Visible = False
    frame_contrib.Enabled = False
    frame_contador.Left = 40
    frame_contador.Visible = True
    frame_contador.Enabled = True
End If

If SSTab1.Tab = 0 Then
 If msk_inscricao.Locked = False Then
   msk_inscricao.SetFocus
 Else
   txt_razao.SetFocus
 End If
Else
   If txt_num_crc_conta.Locked = False Then
   
       txt_num_crc_conta.SetFocus
   Else
       txt_nome.SetFocus
   End If

End If
End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes

End Sub

Private Sub txt_bairro_Change()
teve_alteracao = True

End Sub


Private Sub txt_bairro_conta_Change()
teve_alteracao_contador = True

End Sub

Private Sub txt_bairro_conta_GotFocus()

   txt_bairro_conta.SelStart = 0
   txt_bairro_conta.SelLength = Len(txt_bairro_conta)

End Sub


Private Sub txt_bairro_conta_KeyPress(KeyAscii As Integer)
 
 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub

Private Sub txt_bairro_conta_LostFocus()

If Trim(txt_bairro_conta) = "" Then
   Exit Sub
End If
If SSTab1.Tab = 1 Then
    cRetorno = CritString(txt_bairro_conta)
      If Len(cRetorno) > 0 Then
         MsgBox "Bairro " & cRetorno, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
         txt_bairro_conta.SetFocus
         Exit Sub
      End If
     cmd_salvar_conta.Enabled = True
    txt_bairro_conta = UCase(txt_bairro_conta)
End If
End Sub


Private Sub txt_bairro_GotFocus()
   txt_bairro.SelStart = 0
   txt_bairro.SelLength = Len(txt_bairro)
outAtEcon.Visible = False
End Sub


Private Sub txt_bairro_KeyPress(KeyAscii As Integer)
 
 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 39 Then
   KeyAscii = 96
End If


End Sub


Private Sub txt_bairro_LostFocus()

If Trim(txt_bairro) = "" Then
   Exit Sub
End If
If SSTab1.Tab = 0 Then
    cRetorno = CritString(txt_bairro)
      If Len(cRetorno) > 0 Then
         MsgBox "Bairro " & cRetorno, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
         txt_bairro.SetFocus
         Exit Sub
      End If
     cmd_salvar_contr.Enabled = True
    txt_bairro = UCase(txt_bairro)
End If
End Sub


Private Sub txt_comp_Change()

teve_alteracao = True

End Sub


Private Sub txt_comp_conta_Change()
teve_alteracao_contador = True
End Sub

Private Sub txt_comp_conta_GotFocus()

   txt_comp_conta.SelStart = 0
   txt_comp_conta.SelLength = Len(txt_comp_conta)


End Sub


Private Sub txt_comp_conta_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_comp_conta_LostFocus()

If Trim(txt_comp_conta) = "" Then
   Exit Sub
End If


cRetorno = CritString(txt_comp_conta)
  If Len(cRetorno) > 0 Then
     MsgBox "Complemento " & cRetorno, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     txt_comp_conta.SetFocus
     Exit Sub
  End If
 cmd_salvar_conta.Enabled = True
txt_comp_conta = UCase(txt_comp_conta)
End Sub

Private Sub txt_comp_GotFocus()
   txt_comp.SelStart = 0
   txt_comp.SelLength = Len(txt_comp)
outAtEcon.Visible = False
End Sub


Private Sub txt_comp_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_comp_LostFocus()

If Trim(txt_comp) = "" Then
   Exit Sub
End If

If SSTab1.Tab = 0 Then
    cRetorno = CritString(txt_comp)
      If Len(cRetorno) > 0 Then
         MsgBox "Complemento " & cRetorno, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
         txt_comp.SetFocus
         Exit Sub
      End If
       cmd_salvar_contr.Enabled = True
    txt_comp = UCase(txt_comp)
End If
End Sub

Private Sub txt_corr_eletr_conta_Change()
teve_alteracao_contador = True
End Sub


Private Sub txt_corr_eletr_conta_GotFocus()

   txt_corr_eletr_conta.SelStart = 0
   txt_corr_eletr_conta.SelLength = Len(txt_corr_eletr_conta)

End Sub


Private Sub txt_corr_eletr_conta_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(LCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_corr_eletr_conta_LostFocus()
 cmd_salvar_conta.Enabled = True

End Sub

Private Sub txt_corr_eletronico_Change()
teve_alteracao = True
End Sub


Private Sub txt_corr_eletronico_GotFocus()
   txt_corr_eletronico.SelStart = 0
   txt_corr_eletronico.SelLength = Len(txt_corr_eletronico)
outAtEcon.Visible = False
End Sub


Private Sub txt_corr_eletronico_KeyPress(KeyAscii As Integer)
 
 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(LCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_corr_eletronico_LostFocus()
 cmd_salvar_contr.Enabled = True

'critica_correio_eletronico

End Sub

Private Sub txt_fantasia_Change()
teve_alteracao = True
End Sub


Private Sub txt_fantasia_GotFocus()
   txt_fantasia.SelStart = 0
   txt_fantasia.SelLength = Len(txt_fantasia)
outAtEcon.Visible = False
End Sub


Private Sub txt_fantasia_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_fantasia_LostFocus()

If Trim(txt_fantasia) = "" Then
  Exit Sub
End If
If SSTab1.Tab = 0 Then

      cRetorno = CritString(txt_fantasia)
      If Len(cRetorno) > 0 Then
         MsgBox "Nome Fantasia " & cRetorno, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
         txt_fantasia.SetFocus
         Exit Sub
      End If
     cmd_salvar_contr.Enabled = True
    txt_fantasia = UCase(txt_fantasia)
End If

End Sub

Private Sub txt_log_Change()
teve_alteracao = True
End Sub


Private Sub txt_log_conta_Change()
teve_alteracao_contador = True
End Sub

Private Sub txt_log_conta_GotFocus()

   txt_log_conta.SelStart = 0
   txt_log_conta.SelLength = Len(txt_log_conta)

End Sub


Private Sub txt_log_conta_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_log_conta_LostFocus()

If Trim(txt_log_conta) = "" Then
  Exit Sub
End If
If SSTab1.Tab = 1 Then
     cRetorno = CritString(txt_log_conta)
      If Len(cRetorno) > 0 Then
         MsgBox "Logradouro " & cRetorno, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
         txt_log_conta.SetFocus
         Exit Sub
      End If
    cmd_salvar_conta.Enabled = True
     txt_log_conta = UCase(txt_log_conta)
End If

End Sub

Private Sub txt_log_GotFocus()
   txt_log.SelStart = 0
   txt_log.SelLength = Len(txt_log)
outAtEcon.Visible = False
End Sub

Private Sub txt_log_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_log_LostFocus()

If Trim(txt_log) = "" Then
  Exit Sub
End If
If SSTab1.Tab = 0 Then
      cRetorno = CritString(txt_log)
      If Len(cRetorno) > 0 Then
         MsgBox "Logradouro " & cRetorno, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
         txt_log.SetFocus
         Exit Sub
      End If
       cmd_salvar_contr.Enabled = True
    txt_log = UCase(txt_log)
End If
End Sub

Private Sub txt_nome_Change()
teve_alteracao = True
End Sub


Private Sub txt_nome_esp_KeyPress(KeyAscii As Integer)
 Call PulaCampo(KeyAscii)

End Sub


Private Sub txt_nome_GotFocus()

   txt_nome.SelStart = 0
   txt_nome.SelLength = Len(txt_nome)
outAtEcon.Visible = False
End Sub


Private Sub txt_nome_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_nome_LostFocus()

   If TypeOf Me.ActiveControl Is SSCommand Then
      If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
         Exit Sub
      End If
   End If
   
    If TypeOf Me.ActiveControl Is SSTab Then
      If SSTab1.Tab = 1 Then
         Exit Sub
      End If
   End If


If Trim(txt_nome) = "" Then
   Exit Sub
End If
If SSTab1.Tab = 1 Then
  cRetorno = CritString(txt_nome)
  If Len(cRetorno) > 0 Then
'    CritRazao = "Razão Social " & cRetorno
     MsgBox "Nome do Contador " & cRetorno, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     txt_nome.SetFocus
     Exit Sub
  End If
 cmd_salvar_conta.Enabled = True
 txt_nome = UCase(txt_nome)
End If
End Sub

Private Sub txt_num_Change()
teve_alteracao = True
End Sub


Private Sub txt_num_conta_Change()
teve_alteracao_contador = True
End Sub

Private Sub txt_num_conta_GotFocus()
   txt_num_conta.SelStart = 0
   txt_num_conta.SelLength = Len(txt_num_conta)

End Sub


Private Sub txt_num_conta_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_num_crc_Change()
teve_alteracao = True
End Sub


Private Sub txt_num_crc_conta_Change()
teve_alteracao_contador = True
End Sub

Private Sub txt_num_crc_conta_GotFocus()
   txt_num_crc_conta.SelStart = 0
   txt_num_crc_conta.SelLength = Len(txt_num_crc_conta)
cmb_uf_conta.Enabled = True
End Sub


Private Sub txt_num_crc_conta_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 
 If Not teclas_especiais(KeyAscii) Then
   KeyAscii = 0
   Exit Sub
End If

 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt_num_crc_conta_LostFocus()

  If TypeOf Me.ActiveControl Is SSCommand Then
       If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "buscar" Then
          Exit Sub
       End If
    End If
 
   If TypeOf Me.ActiveControl Is SSTab Then
      If SSTab1.Tab = 1 Then
         Exit Sub
      End If
   End If
   
 If SSTab1.Tab = 1 Then
    If Trim(txt_num_crc_conta) = "" Then
       MsgBox "Número de CRC Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       txt_num_crc_conta.SetFocus
       Exit Sub
    End If
    reta = CRIT_CRC(Trim(txt_num_crc_conta))
    If reta <> 0 Then
       MsgBox "CRC inválido.", 48, NOME_SISTEMA_DEFAULT
       txt_num_crc_conta.SetFocus
       Exit Sub
    End If

    cmd_salvar_conta.Enabled = True
End If

End Sub

Private Sub txt_num_crc_GotFocus()
 txt_num_crc.SelStart = 0
   txt_num_crc.SelLength = Len(txt_num_crc)
outAtEcon.Visible = False
End Sub

Private Sub txt_num_crc_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_num_crc_LostFocus()

If Trim(txt_num_crc) = "" Then
  Exit Sub
End If
'If SSTab1.Tab = 0 Then
'    reta = CRIT_CRC(txt_num_crc)
'    If reta <> 0 Then
'       MsgBox "CRC inválido.", 48, NOME_SISTEMA_DEFAULT
'       txt_num_crc.SetFocus
'       Exit Sub
'    End If
'     cmd_salvar_contr.Enabled = True
'End If
End Sub

Private Sub txt_num_GotFocus()

   txt_num.SelStart = 0
   txt_num.SelLength = Len(txt_num)
outAtEcon.Visible = False
End Sub

Private Sub txt_num_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_num_LostFocus()
 cmd_salvar_contr.Enabled = True

End Sub

Private Sub txt_razao_Change()
teve_alteracao = True
End Sub


Private Sub txt_razao_GotFocus()
   txt_razao.SelStart = 0
   txt_razao.SelLength = Len(txt_razao)
outAtEcon.Visible = False
End Sub


Private Sub txt_razao_KeyPress(KeyAscii As Integer)

 Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If

End Sub


Private Sub txt_razao_LostFocus()
If Trim(txt_razao) = "" Then
   Exit Sub
End If

If SSTab1.Tab = 0 Then
  cRetorno = CritString(txt_razao)
  If Len(cRetorno) > 0 Then
'    CritRazao = "Razão Social " & cRetorno
     MsgBox "Nome Empresarial " & cRetorno, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     txt_razao.SetFocus
     Exit Sub
  End If
  txt_razao = UCase(txt_razao)
  cmd_salvar_contr.Enabled = True
End If

End Sub


