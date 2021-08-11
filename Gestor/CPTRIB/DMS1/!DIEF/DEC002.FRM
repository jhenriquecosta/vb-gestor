VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form dec002 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SIAT - Declaração Eletrônica: Declaração"
   ClientHeight    =   6840
   ClientLeft      =   315
   ClientTop       =   1050
   ClientWidth     =   9600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   6
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   360
   ScaleMode       =   0  'User
   ScaleWidth      =   2850.607
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraIdentificação 
      Caption         =   "Identificação do Contribuinte"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   885
      Left            =   15
      TabIndex        =   6
      Top             =   420
      Width           =   10440
      Begin VB.TextBox txtCondContr 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   450
         Width           =   1365
      End
      Begin VB.TextBox txt_razao 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   450
         Width           =   3675
      End
      Begin VB.ComboBox cmb_tipo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3100
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "tipo"
         Top             =   450
         Width           =   855
      End
      Begin MSMask.MaskEdBox txt_periodo_inicial 
         Height          =   285
         Left            =   1730
         TabIndex        =   1
         Tag             =   "periodo"
         Top             =   450
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Format          =   "@@/@@@@"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_inscricao 
         Height          =   285
         Left            =   60
         TabIndex        =   0
         Tag             =   "inscricao"
         Top             =   450
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Format          =   "@@.@@@.@@@-@"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_tipo_decl 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   7080
         TabIndex        =   111
         Top             =   120
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label Label10 
         Caption         =   "Condição:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4070
         TabIndex        =   63
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Substitutiva:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3100
         TabIndex        =   10
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Período: (mmaaaa)"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   0
         Left            =   1730
         TabIndex        =   9
         Top             =   210
         Width           =   1470
      End
      Begin VB.Label Label2 
         Caption         =   "Inscrição Estadual:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Nome Empresarial"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5550
         TabIndex        =   7
         Top             =   210
         Width           =   1305
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabHeight       =   500
      ShowFocusRect   =   0   'False
      ForeColor       =   16711680
      TabCaption(0)   =   "Dem.das &Operações(Entradas)"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_pisca_epp"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_epp_entr"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LBL_APUR_PISCA"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_Recolhimento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Dados_Compl"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_EntaNormal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frm_grid_fatu"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra_EntraMicroEPP"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra_Dem_Saidas"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_Apur_ICMS"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Timer1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Dem.das O&perações(Saídas)"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "&Apuração do ICMS"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Recolhimentos no Período"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Dados do Exercício A&nterior"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   8100
         Top             =   4380
      End
      Begin VB.Frame fra_Apur_ICMS 
         Height          =   4785
         Left            =   75
         TabIndex        =   65
         Top             =   600
         Width           =   9465
         Begin VB.Frame frm_tela3 
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
            Height          =   4035
            Left            =   240
            TabIndex        =   76
            Top             =   675
            Width           =   8895
            Begin VB.TextBox txt210 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6870
               TabIndex        =   87
               Top             =   2160
               Width           =   1785
            End
            Begin VB.TextBox txt211 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6870
               TabIndex        =   88
               Top             =   2475
               Width           =   1785
            End
            Begin VB.TextBox txt212 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6870
               TabIndex        =   89
               Top             =   2930
               Width           =   1785
            End
            Begin VB.TextBox txt213 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6870
               TabIndex        =   90
               Top             =   3250
               Width           =   1785
            End
            Begin VB.TextBox txt214 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   6870
               Locked          =   -1  'True
               TabIndex        =   91
               TabStop         =   0   'False
               Top             =   3600
               Width           =   1785
            End
            Begin VB.TextBox txt204 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2325
               Locked          =   -1  'True
               MaxLength       =   21
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   1485
               Width           =   1785
            End
            Begin VB.TextBox txt201 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2325
               Locked          =   -1  'True
               MaxLength       =   21
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   495
               Width           =   1785
            End
            Begin VB.TextBox txt203 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2325
               MaxLength       =   21
               TabIndex        =   80
               Top             =   1155
               Width           =   1785
            End
            Begin VB.TextBox txt202 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2325
               MaxLength       =   21
               TabIndex        =   79
               Top             =   825
               Width           =   1785
            End
            Begin VB.TextBox txt206 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   6870
               Locked          =   -1  'True
               MaxLength       =   21
               TabIndex        =   83
               TabStop         =   0   'False
               Top             =   825
               Width           =   1785
            End
            Begin VB.TextBox txt209 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6870
               MaxLength       =   21
               TabIndex        =   86
               Top             =   1815
               Width           =   1785
            End
            Begin VB.TextBox txt208 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6870
               MaxLength       =   21
               TabIndex        =   85
               Top             =   1485
               Width           =   1785
            End
            Begin VB.TextBox txt207 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6870
               MaxLength       =   21
               TabIndex        =   84
               Top             =   1155
               Width           =   1785
            End
            Begin VB.TextBox txt205 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   6870
               Locked          =   -1  'True
               MaxLength       =   21
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   495
               Width           =   1785
            End
            Begin VB.TextBox txt215 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2325
               Locked          =   -1  'True
               MaxLength       =   21
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   2430
               Width           =   1785
            End
            Begin VB.TextBox txt216 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2325
               MaxLength       =   21
               TabIndex        =   93
               Top             =   2760
               Width           =   1785
            End
            Begin VB.TextBox txt217 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2325
               Locked          =   -1  'True
               MaxLength       =   21
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   3060
               Width           =   1785
            End
            Begin VB.TextBox txt218 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2325
               Locked          =   -1  'True
               MaxLength       =   21
               TabIndex        =   96
               TabStop         =   0   'False
               Top             =   3390
               Width           =   1785
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Restituição"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   18
               Left            =   4440
               TabIndex        =   135
               Top             =   1230
               Width           =   795
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Crédito Presumido"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   4440
               TabIndex        =   134
               Top             =   1560
               Width           =   1275
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Antecipação Parcial Interestadual"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   4440
               TabIndex        =   133
               Top             =   1875
               Width           =   2385
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Antecipação Total "
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   4440
               TabIndex        =   132
               Top             =   2175
               Width           =   1350
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Diferencial de Alíquota/Outros"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   4440
               TabIndex        =   131
               Top             =   2535
               Width           =   2160
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00E0E0E0&
               Height          =   1935
               Left            =   120
               Top             =   1920
               Width           =   4095
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Saídas ou Prestações:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   14
               Left            =   210
               TabIndex        =   127
               Top             =   585
               Width           =   1935
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Total dos Débitos:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   210
               TabIndex        =   126
               Top             =   1575
               Width           =   1575
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Estorno de Créditos:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   210
               TabIndex        =   125
               Top             =   1245
               Width           =   1755
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Outros Débitos (Ver Ajuda)"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   2
               Left            =   180
               TabIndex        =   124
               Top             =   945
               Width           =   1875
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Débito do Imposto: "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Index           =   1
               Left            =   210
               TabIndex        =   123
               Top             =   180
               Width           =   2040
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Crédito do Imposto:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   15
               Left            =   4410
               TabIndex        =   120
               Top             =   180
               Width           =   2085
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Total dos Créditos:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   9
               Left            =   4440
               TabIndex        =   119
               Top             =   3650
               Width           =   1515
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo credor do período anterior:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   8
               Left            =   4440
               TabIndex        =   118
               Top             =   3285
               Width           =   2415
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Estorno de Débitos:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   7
               Left            =   4440
               TabIndex        =   117
               Top             =   2955
               Width           =   1635
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total outros Créditos (Ver Ajuda)"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   6
               Left            =   4440
               TabIndex        =   116
               Top             =   825
               Width           =   2280
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Entradas ou Aquisições:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   4440
               TabIndex        =   115
               Top             =   495
               Width           =   2085
            End
            Begin VB.Label Label1 
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
               Height          =   255
               Index           =   17
               Left            =   2370
               TabIndex        =   101
               Top             =   180
               Width           =   1035
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo Devedor:"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   10
               Left            =   390
               TabIndex        =   100
               Top             =   2430
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Deduções:"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   11
               Left            =   390
               TabIndex        =   99
               Top             =   2760
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Imposto a Recolher:"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   12
               Left            =   390
               TabIndex        =   98
               Top             =   3060
               Width           =   1425
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo credor a transferir:"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   13
               Left            =   390
               TabIndex        =   97
               Top             =   3390
               Width           =   1725
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Apuração do Imposto:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   16
               Left            =   990
               TabIndex        =   94
               Top             =   2010
               Width           =   2415
            End
         End
         Begin VB.Frame Frame4 
            Height          =   510
            Left            =   240
            TabIndex        =   68
            Top             =   120
            Width           =   8880
            Begin VB.TextBox msk_fat_mensal_frm4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2850
               MaxLength       =   21
               TabIndex        =   69
               Top             =   150
               Width           =   1545
            End
            Begin VB.TextBox msk_aliquota_frm4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5610
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   150
               Width           =   495
            End
            Begin VB.TextBox msk_imposto_frm4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   7020
               Locked          =   -1  'True
               MaxLength       =   21
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   150
               Width           =   1545
            End
            Begin VB.Label Label_epp 
               Caption         =   "Faturamento do Mês (Receita Bruta :"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   150
               TabIndex        =   74
               Top             =   150
               Width           =   2655
            End
            Begin VB.Label Label8 
               Caption         =   "Aliquota(%):"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   4740
               TabIndex        =   72
               Top             =   150
               Width           =   825
            End
            Begin VB.Label Label8 
               Caption         =   "Imposto:"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   6390
               TabIndex        =   70
               Top             =   150
               Width           =   615
            End
         End
         Begin VB.Frame Frame_tela3 
            Height          =   510
            Left            =   240
            TabIndex        =   66
            Top             =   120
            Width           =   8850
            Begin VB.TextBox txt_fatura_normal 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3510
               MaxLength       =   16
               TabIndex        =   77
               Top             =   150
               Width           =   1785
            End
            Begin VB.Label Label8 
               Caption         =   "Faturamento do Mês (Receita Bruta):"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   67
               Top             =   150
               Width           =   2910
            End
         End
      End
      Begin VB.Frame fra_Dem_Saidas 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   4065
         Left            =   75
         TabIndex        =   21
         Top             =   645
         Width           =   9315
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   465
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   3480
            Width           =   9075
         End
      End
      Begin VB.Frame fra_EntraMicroEPP 
         Height          =   3945
         Left            =   195
         TabIndex        =   59
         Top             =   735
         Width           =   9255
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   129
            Top             =   3510
            Width           =   4575
         End
      End
      Begin VB.Frame frm_grid_fatu 
         Height          =   2715
         Left            =   405
         TabIndex        =   58
         Top             =   990
         Width           =   7305
         Begin VB.TextBox txt238 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4890
            Locked          =   -1  'True
            MaxLength       =   21
            MultiLine       =   -1  'True
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2145
         End
         Begin VB.TextBox txt237 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   4860
            Locked          =   -1  'True
            MaxLength       =   21
            MultiLine       =   -1  'True
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   930
            Width           =   2145
         End
         Begin VB.TextBox txt236 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4860
            Locked          =   -1  'True
            MaxLength       =   21
            MultiLine       =   -1  'True
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   1493
            Width           =   1065
         End
         Begin VB.TextBox txt229 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   21
            MultiLine       =   -1  'True
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   1500
            Width           =   2175
         End
         Begin VB.TextBox txt242 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4950
            Locked          =   -1  'True
            MaxLength       =   21
            MultiLine       =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   1500
            Width           =   2055
         End
      End
      Begin VB.Frame fra_EntaNormal 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   4125
         Left            =   75
         TabIndex        =   15
         Top             =   600
         Width           =   9255
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   405
            Left            =   100
            Locked          =   -1  'True
            TabIndex        =   122
            Top             =   3500
            Width           =   9045
         End
      End
      Begin VB.Frame fra_Dados_Compl 
         Height          =   4635
         Left            =   165
         TabIndex        =   35
         Top             =   735
         Width           =   9465
         Begin VB.TextBox txt424 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7095
            MaxLength       =   21
            TabIndex        =   38
            Top             =   540
            Width           =   2085
         End
         Begin VB.TextBox txt400 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5880
            MaxLength       =   6
            TabIndex        =   37
            Top             =   540
            Width           =   735
         End
         Begin VB.Frame Frame8 
            Caption         =   "Estoque Final:"
            ForeColor       =   &H00FF0000&
            Height          =   1695
            Left            =   5760
            TabIndex        =   52
            Top             =   2820
            Width           =   3555
            Begin VB.TextBox txt422 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1695
               MaxLength       =   21
               TabIndex        =   45
               Top             =   930
               Width           =   1700
            End
            Begin VB.TextBox txt423 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1695
               Locked          =   -1  'True
               MaxLength       =   21
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   1320
               Width           =   1700
            End
            Begin VB.TextBox txt420 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1695
               MaxLength       =   21
               TabIndex        =   43
               Top             =   210
               Width           =   1700
            End
            Begin VB.TextBox txt421 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1695
               MaxLength       =   21
               TabIndex        =   44
               Top             =   570
               Width           =   1700
            End
            Begin VB.Label Label29 
               Caption         =   "Total"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   56
               Top             =   1320
               Width           =   585
            End
            Begin VB.Label Label27 
               Caption         =   "Tributadas"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   90
               TabIndex        =   55
               Top             =   210
               Width           =   825
            End
            Begin VB.Label Label26 
               Caption         =   "Não Tributadas"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   90
               TabIndex        =   54
               Top             =   570
               Width           =   1155
            End
            Begin VB.Label Label25 
               Caption         =   "Outras"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   90
               TabIndex        =   53
               Top             =   960
               Width           =   735
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Estoque Inicial"
            ForeColor       =   &H00FF0000&
            Height          =   1695
            Left            =   5760
            TabIndex        =   39
            Top             =   1110
            Width           =   3555
            Begin VB.TextBox txt419 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1695
               Locked          =   -1  'True
               MaxLength       =   21
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   1320
               Width           =   1700
            End
            Begin VB.TextBox txt416 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1695
               MaxLength       =   21
               TabIndex        =   40
               Top             =   210
               Width           =   1700
            End
            Begin VB.TextBox txt417 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1695
               MaxLength       =   21
               TabIndex        =   41
               Top             =   570
               Width           =   1700
            End
            Begin VB.TextBox txt418 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1695
               MaxLength       =   21
               TabIndex        =   42
               Top             =   930
               Width           =   1700
            End
            Begin VB.Label Label17 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Total"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   90
               TabIndex        =   51
               Top             =   1320
               Width           =   525
            End
            Begin VB.Label Label31 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Tributadas"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   90
               TabIndex        =   50
               Top             =   210
               Width           =   1185
            End
            Begin VB.Label Label30 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Não Tributadas"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   90
               TabIndex        =   49
               Top             =   570
               Width           =   1245
            End
            Begin VB.Label Label29 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Outras"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   48
               Top             =   930
               Width           =   555
            End
         End
         Begin VB.Frame fraGrid 
            ForeColor       =   &H00FF0000&
            Height          =   4425
            Left            =   90
            TabIndex        =   36
            Top             =   120
            Width           =   5535
            Begin VB.TextBox txt415 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3485
               Locked          =   -1  'True
               MaxLength       =   21
               MultiLine       =   -1  'True
               TabIndex        =   104
               TabStop         =   0   'False
               Top             =   3975
               Width           =   1725
            End
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Caixa e Banco em 31/12:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   7080
            TabIndex        =   107
            Top             =   270
            Width           =   2265
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Num. de empreg:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   5730
            TabIndex        =   64
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resumo do Livro de Registro de Inventário "
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   6105
            TabIndex        =   57
            Top             =   870
            Width           =   3105
         End
      End
      Begin VB.Frame fra_Recolhimento 
         Height          =   4335
         Left            =   200
         TabIndex        =   23
         Top             =   615
         Width           =   9165
         Begin VB.Frame frame7 
            Caption         =   "Recolhimento do Período"
            ForeColor       =   &H00FF0000&
            Height          =   3705
            Left            =   120
            TabIndex        =   33
            Top             =   150
            Width           =   6525
         End
         Begin VB.Frame Frame6 
            Caption         =   "Entradas de Ativo Fixo ou Consumo"
            ForeColor       =   &H00FF0000&
            Height          =   1695
            Left            =   200
            TabIndex        =   24
            Top             =   120
            Width           =   3525
            Begin VB.TextBox txt300 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1650
               MaxLength       =   21
               MultiLine       =   -1  'True
               TabIndex        =   26
               Top             =   270
               Width           =   1785
            End
            Begin VB.TextBox txt303 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1650
               Locked          =   -1  'True
               MaxLength       =   21
               MultiLine       =   -1  'True
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1350
               Width           =   1785
            End
            Begin VB.TextBox txt302 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1650
               MaxLength       =   21
               MultiLine       =   -1  'True
               TabIndex        =   28
               Top             =   960
               Width           =   1785
            End
            Begin VB.TextBox txt301 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1650
               MaxLength       =   21
               MultiLine       =   -1  'True
               TabIndex        =   27
               Top             =   600
               Width           =   1785
            End
            Begin VB.Label Label7 
               Caption         =   "Total"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   13
               Left            =   150
               TabIndex        =   32
               Top             =   1410
               Width           =   2025
            End
            Begin VB.Label Label7 
               Caption         =   "Do Exterior"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   14
               Left            =   150
               TabIndex        =   31
               Top             =   1020
               Width           =   1725
            End
            Begin VB.Label Label7 
               Caption         =   "De Outros Estados"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   15
               Left            =   150
               TabIndex        =   30
               Top             =   630
               Width           =   1575
            End
            Begin VB.Label Label7 
               Caption         =   "Do Estado"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   16
               Left            =   150
               TabIndex        =   29
               Top             =   270
               Width           =   2055
            End
         End
      End
      Begin VB.Label LBL_APUR_PISCA 
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   675
         Left            =   90
         TabIndex        =   114
         Top             =   5430
         Visible         =   0   'False
         Width           =   9735
      End
      Begin VB.Label lbl_epp_entr 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   63000
         TabIndex        =   113
         Top             =   510
         Visible         =   0   'False
         Width           =   10275
      End
      Begin VB.Label lbl_pisca_epp 
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Left            =   120
         TabIndex        =   112
         Top             =   5400
         Visible         =   0   'False
         Width           =   9735
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Incompleta"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   7830
      TabIndex        =   62
      Top             =   510
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Finalizada"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   9060
      TabIndex        =   61
      Top             =   510
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox Situacao 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   7830
      ScaleHeight     =   300
      ScaleWidth      =   1590
      TabIndex        =   128
      Top             =   15
      Width           =   1650
   End
   Begin VB.PictureBox Calcular 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5850
      ScaleHeight     =   315
      ScaleWidth      =   945
      TabIndex        =   130
      TabStop         =   0   'False
      Tag             =   "sair"
      Top             =   0
      Width           =   1005
   End
   Begin VB.PictureBox CmdFinalizar 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5040
      ScaleHeight     =   315
      ScaleWidth      =   735
      TabIndex        =   19
      Tag             =   "sair"
      Top             =   0
      Width           =   795
   End
   Begin VB.PictureBox cmd_buscar_contr 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4170
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   17
      Tag             =   "buscar"
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox cmd_imprimir 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3330
      ScaleHeight     =   315
      ScaleWidth      =   765
      TabIndex        =   16
      Top             =   0
      Width           =   825
   End
   Begin VB.PictureBox cmd_sair 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6870
      ScaleHeight     =   315
      ScaleWidth      =   825
      TabIndex        =   20
      Tag             =   "sair"
      Top             =   0
      Width           =   885
   End
   Begin VB.PictureBox cmd_excluir 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1650
      ScaleHeight     =   315
      ScaleWidth      =   765
      TabIndex        =   13
      Tag             =   "excluir"
      Top             =   0
      Width           =   825
   End
   Begin VB.PictureBox cmd_salvar 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2490
      ScaleHeight     =   315
      ScaleWidth      =   765
      TabIndex        =   14
      Tag             =   "salvar"
      Top             =   0
      Width           =   825
   End
   Begin VB.PictureBox cmd_diretorio 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   840
      ScaleHeight     =   315
      ScaleWidth      =   735
      TabIndex        =   12
      Tag             =   "diretorio"
      Top             =   0
      Width           =   795
   End
   Begin VB.PictureBox cmd_iniciar 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   765
      TabIndex        =   11
      Tag             =   "iniciar"
      Top             =   0
      Width           =   825
   End
End
Attribute VB_Name = "dec002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim soma As Double
Dim AnoSis As Double
Dim Frase As String, Frase3 As String
Dim Pos3 As Integer
Dim erro_inscricao As Integer
Dim PER_INICIO As Double
Dim apenas_troquei_estado As Integer
Dim campopreench As Integer
Dim GridArray() As Variant   'demonstrativo das oper. entradas
Dim grid_saidas() As Variant  'demonstrativo das opr saidas
Dim grid_ApurICMS() As Variant  'Apuração do ICMS
Dim grid_recolim() As Variant  'recolhimento do periodo
Dim grid_fatura() As Variant  'faturamento frame 5
Dim grid_DetDesp() As Variant  'Detalhamento das Despesas
Dim TAB_ANTERIOR As Integer
Dim MaxCol As Long
Dim MaxRow As Long
Dim MaxCol_s As Long
Dim MaxRow_s As Long
Dim MaxCol_ICMS As Long
Dim MaxRow_ICMS As Long
Dim MaxCol_rec As Long
Dim MaxRow_rec As Long
Dim MaxCol_fat As Long
Dim MaxRow_fat As Long
Dim MaxCol_DetDesp As Long
Dim MaxRow_DetDesp As Long

Dim valor_calculo_saidas As Double
Dim valor_contabil_saidas As Double
Dim valor_creditado_saidas As Double
Dim valor_isentas_saidas As Double
Dim valor_outras_saidas As Double
Dim erro_periodo As Integer
Dim valor_a_gravar As String
Dim EstadoDecla As String
Dim APENAS_ENTREI As Integer
'
Dim valor_total_coluna1_DemEntr As Double
Dim valor_total_coluna2_DemEntr As Double
Dim valor_total_coluna3_DemEntr As Double

Dim valor_total_coluna4_DemEntr As Double
Dim valor_total_coluna5_DemEntr As Double

Dim motivo_entradas As String
Dim motivo_percentual As String
Function acumula_grid_apurIcms1(GridArray() As Variant) As Double
Dim indice As Integer
Dim total As Double
total = 0#
acumula_grid_apurIcms1 = 0#
For indice = 0 To 2
 If IsNumeric(Trim(GridArray(1, indice))) Then
  If Trim(GridArray(1, indice)) <> "" And Trim(GridArray(1, indice)) <> "0,00" Then
     total = total + GridArray(1, indice)
  End If
 End If
Next indice
acumula_grid_apurIcms1 = total
End Function

Function acumula_grid_apurIcms2(GridArray() As Variant) As Double
Dim indice As Integer
Dim total As Double
total = 0#
acumula_grid_apurIcms2 = 0#
For indice = 4 To 7
 If IsNumeric(Trim(GridArray(1, indice))) Then
  If Trim(GridArray(1, indice)) <> "" And Trim(GridArray(1, indice)) <> "0,00" Then
    total = total + GridArray(1, indice)
  End If
 End If
Next indice

acumula_grid_apurIcms2 = total

End Function

Function acumula_grid_detalhes(GridArray() As Variant) As Double
Dim indice As Integer
Dim total As Double
total = 0#
acumula_grid_detalhes = 0#
For indice = 0 To 13
 If IsNumeric(Trim(GridArray(1, indice))) Then
  If Trim(GridArray(1, indice)) <> "" And Trim(GridArray(1, indice)) <> "0,00" Then
    total = total + GridArray(1, indice)
  End If
 End If
Next indice


acumula_grid_detalhes = total

End Function

Function acumula_grid_faturamento(GridArray() As Variant, coluna As Integer) As Double
Dim indice As Integer
Dim total As Double
total = 0#
acumula_grid_faturamento = 0#

For indice = 0 To 2

 If IsNumeric(Trim(GridArray(coluna, indice))) Then
 
  If Trim(GridArray(coluna, indice)) <> "" And Trim(GridArray(coluna, indice)) <> "0,00" Then
    total = total + GridArray(coluna, indice)
  End If
 End If
Next indice


acumula_grid_faturamento = total

End Function
Function acumula_grid_recolhimento(GridArray() As Variant, coluna As Integer) As Double
Dim indice As Integer
Dim total As Double

total = 0#
acumula_grid_recolhimento = 0#

Select Case g_condicao

Case Is = 1, 2, 5, 6
    
    grid_recolhimento.Row = 6
    If grid_recolhimento.Columns(coluna).Value <> "" Then
        soma = CCur(grid_recolhimento.Columns(coluna).Value)
    Else
       soma = 0
    End If
    
    For indice = 0 To 5
         If IsNumeric(Trim(grid_recolim(coluna, indice))) Then
          If Trim(grid_recolim(coluna, indice)) <> "" And Trim(grid_recolim(coluna, indice)) <> "0,00" Then
            total = total + grid_recolim(coluna, indice)
          End If
         End If
    Next indice

Case Is = 3, 4, 7
      
      grid_recolhimento.Row = 7
      If grid_recolhimento.Columns(coluna).Value <> "" Then
            soma = CCur(grid_recolhimento.Columns(coluna).Value)
      Else
           soma = 0
      End If
      For indice = 0 To 6
            If IsNumeric(Trim(grid_recolim(coluna, indice))) Then
             If Trim(grid_recolim(coluna, indice)) <> "" And Trim(grid_recolim(coluna, indice)) <> "0,00" Then
               total = total + grid_recolim(coluna, indice)
             End If
            End If
       Next indice
       
End Select

acumula_grid_recolhimento = total

End Function

Function acumula_totais(GridArray() As Variant, coluna As Integer, linha As Integer) As Double
Dim indice As Integer
Dim total As Double
total = 0#
acumula_totais = 0#
For indice = 0 To 10

 If IsNumeric(Trim(GridArray(coluna, indice))) Then
  If Trim(GridArray(coluna, indice)) <> "" And CDbl(GridArray(coluna, indice)) <> 0 Then
    total = total + GridArray(coluna, indice)
  End If
 End If
Next indice


acumula_totais = total
End Function

Sub acumula_total_204()

    grid_ApurICMS(1, 3) = acumula_grid_apurIcms1(grid_ApurICMS)
    grid_ApurICMS(1, 3) = Format(grid_ApurICMS(1, 3), "#,##0.#0")
    txt204.Text = grid_ApurICMS(1, 3)
    
End Sub

Sub acumula_total_214()
txt205.Text = IIf(IsNumeric(txt205), txt205, "0,00")
txt206.Text = IIf(IsNumeric(txt206), txt206, "0,00")
txt212.Text = IIf(IsNumeric(txt212), txt212, "0,00")
txt213.Text = IIf(IsNumeric(txt213), txt213, "0,00")

txt214.Text = Format(CCur(txt205) + CCur(txt206) + CCur(txt212) + CCur(txt213), "standard")
grid_ApurICMS(1, 13) = txt214.Text
End Sub

Sub acumula_total_303()
Dim c300 As Double
Dim c301 As Double
Dim c302 As Double
c300 = 0#
c301 = 0#
c302 = 0#

If Trim(txt300.Text) <> "" And Len(Trim(txt300.Text)) <> 0 Then
   c300 = Trim(txt300.Text)
End If
If Trim(txt301.Text) <> "" And Len(Trim(txt301.Text)) <> 0 Then
   c301 = Trim(txt301.Text)
End If
If Trim(txt302.Text) <> "" And Len(Trim(txt302.Text)) <> 0 Then
   c302 = Trim(txt302.Text)
End If

txt303.Text = c300 + c301 + c302


txt303.Text = Format(txt303.Text, "#,##0.#0")

End Sub
Sub acumula_total_419()

Dim c416 As Double
Dim c417 As Double
Dim c418 As Double
c416 = 0#
c417 = 0#
c418 = 0#

If Trim(txt416.Text) <> "" And Len(Trim(txt416.Text)) <> 0 Then
   c416 = Trim(txt416.Text)
End If
If Trim(txt417.Text) <> "" And Len(Trim(txt417.Text)) <> 0 Then
   c417 = Trim(txt417.Text)
End If
If Trim(txt418.Text) <> "" And Len(Trim(txt418.Text)) <> 0 Then
   c418 = Trim(txt418.Text)
End If

'txt419.Text = CCur(CDbl(c416)) + CCur(CDbl(c417)) + CCur(CDbl(c418))
txt419.Text = c416 + c417 + c418
txt419.Text = Format(txt419.Text, "#,##0.#0")

End Sub

Sub acumula_total_423()
Dim c420 As Double
Dim c421 As Double
Dim c422 As Double
c420 = 0#
c421 = 0#
c422 = 0#

If Trim(txt420.Text) <> "" And Len(Trim(txt420.Text)) <> 0 Then
   c420 = Trim(txt420.Text)
End If
If Trim(txt421.Text) <> "" And Len(Trim(txt421.Text)) <> 0 Then
   c421 = Trim(txt421.Text)
End If
If Trim(txt422.Text) <> "" And Len(Trim(txt422.Text)) <> 0 Then
   c422 = Trim(txt422.Text)
End If

txt423.Text = c420 + c421 + c422


txt423.Text = Format(txt423.Text, "#,##0.#0")

End Sub


Sub ativa_label_pisca_apur()
'LBL_APUR_PISCA.Visible = True
'LBL_APUR_PISCA.Left = 70
'LBL_APUR_PISCA.Top = 4800
'LBL_APUR_PISCA.Width = 9365
'LBL_APUR_PISCA.Refresh
'Pos3 = 0
'Frase3 = Frase + Space(30)
'LBL_APUR_PISCA.Caption = Frase3

End Sub

Sub ativa_label_pisca_entrada()


'lbl_epp_entr.Visible = True
'lbl_epp_entr.Left = 70
'lbl_epp_entr.Top = 4800
'lbl_epp_entr.Refresh
'Pos3 = 0
'Frase3 = Frase + Space(30)
'lbl_epp_entr.Caption = Frase3



End Sub

Sub ativa_label_pisca_pisca()

'lbl_pisca_epp.Left = 70
'lbl_pisca_epp.Top = 4800
'lbl_pisca_epp.Width = 9365



'lbl_pisca_epp.Visible = True
'lbl_pisca_epp.Refresh
'    Frase = " Senhor Contribuinte: Esta informação é muito importante e o senhor deve ler atentamente antes de fazer qualquer besteira"
'    Pos3 = 0
'    Frase3 = Frase + Space(30)
'    lbl_pisca_epp.Caption = Frase3

End Sub
Function atualiza_declaracao(tipo As Integer) As Integer
Dim sit As String
Dim controle As Double
Dim indficha As Integer

On Error GoTo erro_atualiza
valor_total_coluna1_DemEntr = 0#
valor_total_coluna2_DemEntr = 0#
valor_total_coluna3_DemEntr = 0#
valor_total_coluna4_DemEntr = 0#
valor_total_coluna5_DemEntr = 0#
atualiza_declaracao = True

For indficha = 0 To 4
       If Navega_Ficha(0, indficha) <> Navega_Ficha(1, indficha) Then
          Exit For
       End If
Next indficha
   
If indficha <> 5 Then
      sit = "I"
      EstadoDecla = "I"
Else
      sit = "F"
      EstadoDecla = "F"
End If

controle = calcula_controle(Trim(Str(msk_inscricao)), g_periodo_inicial, tipo) * 100
If Err <> 0 Then
   atualiza_declaracao = False
    Exit Function
End If
   
query = "Select * from tab_declaracao where tde_num_inscricao = '" & Trim(Str(msk_inscricao)) & "'"
query = query & " and tde_periodo_inicial =  '" & g_periodo_inicial & "'"
query = query & " and tde_substitutiva =  " & tipo

Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
   
If restauracontr.EOF And restauracontr.BOF Then
    query = "insert into tab_declaracao"
    query = query & "(tde_num_inscricao, tde_periodo_inicial, tde_substitutiva, "
    query = query & "tde_Estado , tde_aliquota_m17, tde_condicao, tde_controle ) "
    
    query = query & "values('" & Trim(Str(msk_inscricao)) & "','" & g_periodo_inicial & "','"
    query = query & tipo & "','"
    query = query & sit & "'," & g_aliquota_Maior17 & "," & g_condicao & "'" & controle & "')"
    If Not EjecutarQuery(query) Then
       atualiza_declaracao = False
       Exit Function
    End If
Else
   query = "update tab_declaracao set "
   query = query & "tde_estado = '" & sit & "'"
   query = query & ", tde_controle = '" & controle & "'"
   query = query & " where tde_num_inscricao = '" & Trim(Str(msk_inscricao)) & "'"
   query = query & " and tde_periodo_inicial = '" & g_periodo_inicial & "'"
   query = query & " and tde_substitutiva = " & Val(g_tipo_decla)
   query = query & " and tde_aliquota_m17 = " & g_aliquota_Maior17
   query = query & " and tde_condicao = " & g_condicao

    If Not EjecutarQuery(query) Then
       atualiza_declaracao = False
       Exit Function
    End If
End If

restauracontr.Close
If EstadoDecla = "I" Then
  apenas_troquei_estado = 1
    Option1(0).Value = True
ElseIf EstadoDecla = "F" Or EstadoDecla = "G" Then
apenas_troquei_estado = 1
    Option1(1).Value = True
End If
Screen.MousePointer = 0
Exit Function

erro_atualiza:

     MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next
    
End Function

Sub AtualizaGrades()
On Error Resume Next

   If GRID_DEM_ENTRADAS.Visible And GRID_DEM_ENTRADAS.Enabled Then
      GRID_DEM_ENTRADAS.Refresh
   End If
   If grid_dem_saidas.Visible And grid_dem_saidas.Enabled Then
      grid_dem_saidas.Refresh
   End If
   If grid_recolhimento.Visible And grid_recolhimento.Enabled Then
      grid_recolhimento.Refresh
   End If

   If grid_Entra_MicroEPP.Visible And grid_Entra_MicroEPP.Enabled Then
      grid_Entra_MicroEPP.Refresh
   End If
   If grdDetDesp.Visible And grdDetDesp.Enabled Then
      grdDetDesp.Refresh
   End If
   If grid_faturamento.Visible And grid_faturamento.Enabled Then
      grid_faturamento.Refresh
   End If
   
   If gformulario = 0 Then
      grid_dem_saidas.MarqueeStyle = 5
      grid_recolhimento.MarqueeStyle = 5
      grid_Entra_MicroEPP.MarqueeStyle = 5
      grdDetDesp.MarqueeStyle = 5
      grid_faturamento.MarqueeStyle = 5
      GRID_DEM_ENTRADAS.MarqueeStyle = 5
   End If
 
   
End Sub



Sub Calcula_206()
txt207.Text = IIf(IsNumeric(txt207), txt207, "0,00")
txt208.Text = IIf(IsNumeric(txt208), txt208, "0,00")
txt209.Text = IIf(IsNumeric(txt209), txt209, "0,00")
txt210.Text = IIf(IsNumeric(txt210), txt210, "0,00")
txt211.Text = IIf(IsNumeric(txt211), txt211, "0,00")
txt206.Text = Format(CCur(txt207) + CCur(txt208) + CCur(txt209) + CCur(txt210) + CCur(txt211), "standard")
grid_ApurICMS(1, 5) = txt206.Text
acumula_total_214
calcula_dados_ApuracaoImposto_chaim
End Sub

Function calcula_apur_imposto_Faturamento_Versus_Aliquota(FATURAMENTO As Variant, ALIQUOTA As Variant) As Double
Dim valor_fat As Double
Dim valor_ali As Double

calcula_apur_imposto_Faturamento_Versus_Aliquota = 0#

If Trim(FATURAMENTO) = "" Then
    valor_fat = 0#
Else
   valor_fat = CDbl(FATURAMENTO)
End If

If Trim(ALIQUOTA) = "" Then
    valor_ali = 0#
Else
   valor_ali = CDbl(ALIQUOTA)
End If

calcula_apur_imposto_Faturamento_Versus_Aliquota = (valor_fat * valor_ali) / 100


End Function

Sub calcula_dados_Apuracao_imposto()


If CDbl(txt204.Text) > CDbl(txt209.Text) Then
  txt210.Text = txt204.Text - txt209.Text
  txt210.Text = Format(txt210.Text, "#,##0.#0")
  grid_ApurICMS(1, 9) = txt210.Text
  txt213.Text = "0,00"
  grid_ApurICMS(1, 12) = txt213.Text
ElseIf CDbl(txt209.Text) > CDbl(txt204.Text) Then
    txt210.Text = "0,00"
    grid_ApurICMS(1, 9) = txt210.Text
    txt213.Text = (txt209.Text - txt204.Text)    '  + txt211.Text
    txt213.Text = Format(txt213.Text, "#,##0.#0")
    grid_ApurICMS(1, 12) = txt213.Text
    txt212.Text = "0,00"
    grid_ApurICMS(1, 11) = txt212.Text
Else
    txt210.Text = "0,00"
    grid_ApurICMS(1, 9) = txt210.Text
    txt213.Text = "0,00"
    grid_ApurICMS(1, 12) = txt213.Text
    
End If


End Sub
Sub calcula_dados_ApuracaoImposto_chaim()

txt214 = IIf(IsNumeric(txt214), txt214, "0,00")
txt204 = IIf(IsNumeric(txt204), txt204, "0,00")
txt215 = IIf(IsNumeric(txt215), txt215, "0,00")
txt216 = IIf(IsNumeric(txt216), txt216, "0,00")

txt215.Text = txt204.Text - txt214.Text
txt217.Text = txt215.Text - txt216.Text

If CDbl(txt217.Text) < 0 Then
    txt218.Text = txt217.Text * -1
    txt218.Text = Format(txt218.Text, "#,##0.#0")
    txt217.Text = "0,00"
Else
    txt217.Text = Format(txt217.Text, "#,##0.#0")
    txt218.Text = "0,00"
End If

If CDbl(txt215.Text) < 0 Then
    txt215.Text = "0,00"
End If

txt215.Text = Format(txt215.Text, "#,##0.#0")
grid_ApurICMS(1, 14) = txt215.Text
grid_ApurICMS(1, 16) = txt217.Text
grid_ApurICMS(1, 17) = txt218.Text


End Sub


Sub calcula_imposto_ARecolher()

   If txt216.Text <> "" And IsNumeric(txt216.Text) Then
       txt217.Text = txt215.Text - txt216.Text
       If CDbl(txt212.Text) < 0 Then
          txt218.Text = txt217.Text * -1
          txt218.Text = Format(txt218.Text, "#,##0.#0")
          txt217.Text = "0,00"
        Else
           txt217.Text = Format(txt217.Text, "#,##0.#0")
           txt218.Text = "0,00"
        End If
   Else
      txt217.Text = txt215.Text      '  isso seria as deducoes opor estar zerada- 0
      txt216.Text = "0,00"
   
   
   End If
   
grid_ApurICMS(1, 16) = txt217.Text
grid_ApurICMS(1, 17) = txt218.Text
End Sub


Function CALCULA_IMPOSTO_EPP(FATURAMENTO As String, ALIQUOTA As String) As Double

Dim valor_fat As Double
Dim valor_ali As Double

CALCULA_IMPOSTO_EPP = 0#

If Trim(FATURAMENTO) = "" Then
    valor_fat = 0#
Else
   valor_fat = CDbl(FATURAMENTO)
End If

If Trim(ALIQUOTA) = "" Then
    valor_ali = 0#
Else
   valor_ali = ALIQUOTA
End If

CALCULA_IMPOSTO_EPP = (valor_fat * valor_ali) / 100


End Function

Sub calcula_Saldo_credor()


If txt204.Text < txt209.Text Then
  txt213.Text = (txt209.Text - txt204.Text) + txt211.Text
   txt213.Text = Format(txt213.Text, "#,##0.#0")
   grid_ApurICMS(1, 12) = txt213.Text
End If


End Sub


Sub cargar_grid_ApurICMS()
On Error Resume Next
 Dim i As Long
   
 query = "Select *  From tab_detalhe_declaracao "
 query = query & "Where tdd_num_inscricao  = '" & g_num_inscricao & "'"
 query = query & " and tdd_periodo_inicial = '" & g_periodo_inicial & "'"
 query = query & " and tdd_substitutiva = " & g_tipo_decla
 query = query & " and tdd_item between 200 and 221 "
  
  
  ReDim grid_ApurICMS(0 To 1, 0 To 17)
  ReDim Preserve grid_ApurICMS(0 To 1, 0 To 17)
    MaxCol_ICMS = 2
    MaxRow_ICMS = 18
    

    Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
    If seleccion.BOF Or seleccion.EOF Then
        seleccion.Close
        monta_cabecalho_ApurICMS
        Exit Sub
    End If
    
    seleccion.MoveFirst
    
    monta_cabecalho_ApurICMS
    i = 0
    While Not seleccion.EOF
      Select Case seleccion("tdd_item")
       Case Is = 200
        txt_fatura_normal = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 201   '  coluna 1
         grid_ApurICMS(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt201 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 202
         grid_ApurICMS(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt202 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 203
         grid_ApurICMS(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt203 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 204
         grid_ApurICMS(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt204.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 205
         grid_ApurICMS(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt205 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 206
         grid_ApurICMS(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt206 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 207
         grid_ApurICMS(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt207 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 208
         grid_ApurICMS(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt208 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 209
         grid_ApurICMS(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt209.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 210
         grid_ApurICMS(1, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt210 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 211
         grid_ApurICMS(1, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt211 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 212
         grid_ApurICMS(1, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt212 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 213
         grid_ApurICMS(1, 12) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt213 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 214
         grid_ApurICMS(1, 13) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 215
         grid_ApurICMS(1, 14) = Format(seleccion("tdd_valor"), "#0.#0")
         txt215 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 216
         grid_ApurICMS(1, 15) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 217
         grid_ApurICMS(1, 16) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 218
         grid_ApurICMS(1, 17) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 219
         msk_fat_mensal_frm4.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 220
         msk_aliquota_frm4 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 221
         msk_imposto_frm4 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      End Select
       seleccion.MoveNext
       i = i + 1
    Wend
    seleccion.Close
    MaxCol_ICMS = 2
    MaxRow_ICMS = 18
    
End Sub
Sub cargar_grid_DetDesp()
On Error Resume Next
 Dim i As Long
   
 query = "Select *  From tab_detalhe_declaracao "
 query = query & "Where tdd_num_inscricao  = '" & g_num_inscricao & "'"
 query = query & " and tdd_periodo_inicial = '" & g_periodo_inicial & "'"
 query = query & " and tdd_substitutiva = " & g_tipo_decla
  
   query = query & " and tdd_item between 400 and 424 "
  
  ReDim grid_DetDesp(0 To 1, 0 To 14)
  ReDim Preserve grid_DetDesp(0 To 1, 0 To 14)
    MaxCol_DetDesp = 2
    MaxRow_DetDesp = 15
    

    Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
    If seleccion.BOF Or seleccion.EOF Then
        seleccion.Close
        monta_cabecalho_DetDesp
         grdDetDesp.Refresh
        Exit Sub
    End If
    
    seleccion.MoveFirst
    
    monta_cabecalho_DetDesp
    i = 0
    While Not seleccion.EOF
      Select Case seleccion("tdd_item")
       Case Is = 400
          txt400.Text = seleccion("tdd_valor")
       Case Is = 401   '  coluna 1
         grid_DetDesp(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 402
         grid_DetDesp(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 403
         grid_DetDesp(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 404
         grid_DetDesp(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 405
         grid_DetDesp(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 406
         grid_DetDesp(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 407
         grid_DetDesp(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 408
         grid_DetDesp(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 409
         grid_DetDesp(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 410
         grid_DetDesp(1, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 411
         grid_DetDesp(1, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 412
         grid_DetDesp(1, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 413
         grid_DetDesp(1, 12) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 414
         grid_DetDesp(1, 13) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 415
         grid_DetDesp(1, 14) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
          txt415.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 416
          txt416.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 417
          txt417.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 418
          txt418.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 419
          txt419.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 420
          txt420.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 421
          txt421.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 422
          txt422.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 423
          txt423.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 424
          txt424.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")

      End Select
       seleccion.MoveNext
        i = i + 1
    Wend
    grdDetDesp.Refresh
    seleccion.Close
    MaxCol_DetDesp = 2
    MaxRow_DetDesp = 15
    
End Sub

Sub cargar_grid_faturamento()
On Error Resume Next
 Dim estado As String
 
  Dim i As Long
   
 query = "Select *  From tab_detalhe_declaracao "
 query = query & "Where tdd_num_inscricao  = '" & g_num_inscricao & "'"
 query = query & " and tdd_periodo_inicial = '" & g_periodo_inicial & "'"
 query = query & " and tdd_substitutiva = " & g_tipo_decla
 query = query & " and tdd_item between 223 and 242 "
   
'      ReDim grid_fatura(0 To 3, 0 To 6)
'      ReDim Preserve grid_fatura(0 To 3, 0 To 6)
       ReDim grid_fatura(0 To 3, 0 To 3)
      ReDim Preserve grid_fatura(0 To 3, 0 To 3)
     
       MaxCol_fat = 4
       MaxRow_fat = 4
       Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
       If seleccion.BOF Or seleccion.EOF Then
           seleccion.Close
           monta_cabecalho_FAt
           grid_faturamento.Refresh
           Exit Sub
       End If
       seleccion.MoveFirst
       
       monta_cabecalho_FAt
       i = 0
       While Not seleccion.EOF
         Select Case seleccion("tdd_item")
   
            'coluna 2
          Case Is = 223
            grid_fatura(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
          Case Is = 224
            grid_fatura(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
          Case Is = 225
            grid_fatura(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     '     Case Is = 226
     '       grid_fatura(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      '    Case Is = 227
      '      grid_fatura(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      '    Case Is = 228
      '      grid_fatura(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
          Case Is = 229
            grid_fatura(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
            txt229.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
        'coluna 3
          Case Is = 230
            grid_fatura(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
          Case Is = 231
            grid_fatura(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
          Case Is = 232
            grid_fatura(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'          Case Is = 233
'            grid_fatura(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'          Case Is = 234
'            grid_fatura(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'          Case Is = 235
'            grid_fatura(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
        'coluna 4
          Case Is = 236
            grid_fatura(3, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
            txt236.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
          Case Is = 237
            grid_fatura(3, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
            txt237.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
          Case Is = 238
            grid_fatura(3, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
            txt238.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
 '         Case Is = 239
 '           grid_fatura(3, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
 '         Case Is = 240
 '           grid_fatura(3, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
 '         Case Is = 241
 '           grid_fatura(3, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
          Case Is = 242
            grid_fatura(3, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
            txt242.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
   
         End Select
           seleccion.MoveNext
           i = i + 1
       Wend
      grid_faturamento.Refresh
        seleccion.Close
      MaxCol_fat = 4
      MaxRow_fat = 4

End Sub
Sub cargar_grid_recolhimento()
  Dim i As Long
On Error Resume Next
 query = "Select *  From tab_detalhe_declaracao "
 query = query & "Where tdd_num_inscricao  = '" & g_num_inscricao & "'"
 query = query & " and tdd_periodo_inicial = '" & g_periodo_inicial & "'"
 query = query & " and tdd_substitutiva = " & g_tipo_decla
   query = query & " and tdd_item between 300 and 325 "
   
    ReDim grid_recolim(0 To 0)
    
    Select Case g_condicao
       Case Is = 3, 4, 7
            ReDim grid_recolim(0 To 2, 0 To 9)
            ReDim Preserve grid_recolim(0 To 2, 0 To 9)
            MaxRow_rec = 8
       Case Else
            ReDim grid_recolim(0 To 2, 0 To 8)
            ReDim Preserve grid_recolim(0 To 2, 0 To 8)
            MaxRow_rec = 7
    End Select
   
   MaxCol_rec = 3
  
  Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
    If seleccion.BOF Or seleccion.EOF Then
       seleccion.Close
        monta_cabecalho_recolhimento
        grid_recolhimento.Refresh
        Exit Sub
    End If
    
    seleccion.MoveFirst
    monta_cabecalho_recolhimento
  
    While Not seleccion.EOF
      If g_condicao = 1 Then
         Select Case seleccion("tdd_item")
             Case Is = 304   '  coluna 1
               grid_recolim(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 307
               grid_recolim(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 308
               grid_recolim(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 309
               grid_recolim(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 310
               grid_recolim(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 311
               grid_recolim(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 312
               grid_recolim(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 313
               grid_recolim(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 314
               grid_recolim(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 315   '  coluna 2
               grid_recolim(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 318
               grid_recolim(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 319
               grid_recolim(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 320
               grid_recolim(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 321
               grid_recolim(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 322
               grid_recolim(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 323
               grid_recolim(2, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 324
               grid_recolim(2, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
             Case Is = 325
               grid_recolim(2, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")

          End Select
       ElseIf g_condicao = 2 Then
               Select Case seleccion("tdd_item")
                Case Is = 306
                  grid_recolim(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 307
                  grid_recolim(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 308
                  grid_recolim(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 309
                  grid_recolim(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 310
                  grid_recolim(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 311
                  grid_recolim(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 312
                  grid_recolim(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 313
                  grid_recolim(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 314
                  grid_recolim(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 317
                  grid_recolim(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 318
                  grid_recolim(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 319
                  grid_recolim(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 320
                  grid_recolim(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 321
                  grid_recolim(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 322
                  grid_recolim(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 323
                  grid_recolim(2, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 324
                  grid_recolim(2, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                Case Is = 325
                  grid_recolim(2, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
            
            End Select
         ElseIf g_condicao = 5 Or g_condicao = 6 Then
                   Select Case seleccion("tdd_item")
                    Case Is = 305
                      grid_recolim(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 307
                      grid_recolim(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 308
                      grid_recolim(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 309
                      grid_recolim(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 310
                      grid_recolim(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 311
                      grid_recolim(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 312
                      grid_recolim(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 313
                      grid_recolim(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 314
                      grid_recolim(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 316
                      grid_recolim(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 318
                      grid_recolim(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 319
                      grid_recolim(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 320
                      grid_recolim(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 321
                      grid_recolim(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 322
                      grid_recolim(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 323
                      grid_recolim(2, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 324
                      grid_recolim(2, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                    Case Is = 325
                      grid_recolim(2, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                End Select
             Else
                Select Case seleccion("tdd_item")
                 Case Is = 304   '  coluna 1
                   grid_recolim(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 305
                   grid_recolim(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 307
                   grid_recolim(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 308
                   grid_recolim(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 309
                   grid_recolim(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 310
                   grid_recolim(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 311
                   grid_recolim(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 312
                   grid_recolim(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 313
                   grid_recolim(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 314
                   grid_recolim(1, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 315   '  coluna 1
                   grid_recolim(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 316
                   grid_recolim(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 318
                   grid_recolim(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 319
                   grid_recolim(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 320
                   grid_recolim(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 321
                   grid_recolim(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 322
                   grid_recolim(2, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 323
                   grid_recolim(2, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 324
                   grid_recolim(2, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
                 Case Is = 325
                   grid_recolim(2, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
               
               End Select
      End If
       Select Case seleccion("tdd_item")
              Case Is = 300
               txt300.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
              Case Is = 301
               txt301.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
              Case Is = 302
               txt302.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
              Case Is = 303
                txt303.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      End Select
        seleccion.MoveNext
        
    Wend
    grid_recolhimento.Refresh
seleccion.Close
    

End Sub
Sub carrega_navega()

         Navega_Ficha(0, 0) = 1
         Navega_Ficha(0, 1) = 1
         Navega_Ficha(0, 2) = 1
         Navega_Ficha(0, 3) = 1
         Navega_Ficha(0, 4) = 1

         Navega_Ficha(1, 0) = 1
         Navega_Ficha(1, 2) = 1
         Navega_Ficha(1, 3) = 1
         
        If g_condicao = 2 Or g_condicao = 5 Or g_condicao = 6 Then
           Navega_Ficha(0, 1) = 0
           Navega_Ficha(1, 1) = 0
        Else
        
            Navega_Ficha(1, 1) = 1
         End If
         
   If Mid(Trim(txt_periodo_inicial), 1, 2) <> "03" Then  'Março
      Navega_Ficha(0, 4) = 0
      Navega_Ficha(1, 4) = 0
   ElseIf g_ano_inicio_operacao < AnoSis Then
         Navega_Ficha(1, 4) = 1
   Else
      Navega_Ficha(1, 4) = 0
       Navega_Ficha(0, 4) = 0
   End If



End Sub

Function converte_em_mes(linha As Integer) As String
converte_em_mes = ""

If Mid$(Trim(g_periodo_inicial), 5, 2) = "03" Then
   Select Case linha
     Case Is = 1
         converte_em_mes = "Janeiro"
     Case Is = 2
         converte_em_mes = "Fevereiro"
     Case Is = 3
        converte_em_mes = "Março"
    End Select
ElseIf Mid$(Trim(g_periodo_inicial), 5, 2) = "06" Then
   Select Case linha
     Case Is = 1
         converte_em_mes = "Abril"
     Case Is = 2
        converte_em_mes = "Maio"
     Case Is = 3
         converte_em_mes = "Junho"
    End Select
    
ElseIf Mid$(Trim(g_periodo_inicial), 5, 2) = "09" Then
       Select Case linha
         Case Is = 1
         converte_em_mes = "Julho"
        Case Is = 2
            converte_em_mes = "Agosto"
        Case Is = 3
            converte_em_mes = "Setembro"
     End Select
ElseIf Mid$(Trim(g_periodo_inicial), 5, 2) = "12" Then
       Select Case linha
        Case Is = 1
            converte_em_mes = "Outubro"
        Case Is = 2
            converte_em_mes = "Novembro"
        Case Is = 3
            converte_em_mes = "Dezembro"
     End Select

 End If
 
End Function

Function critica_apuracao_icms() As Integer
critica_apuracao_icms = True
msgerro = ""
Dim linha As Integer
Dim erro As Integer
Dim mensa1 As String
Dim mensa2 As String
Dim msg_micro1 As String
Dim msg_micro2 As String
Dim msg_micro3 As String
Dim msg_micro4 As String
Dim msg_micro5 As String
Dim msg_micro6 As String
Dim msg_micro7 As String
Dim msg_micro8 As String

msg_micro1 = "I - Para as Microempresas, contribuintes exclusivamente do ICMS, que tenham as seguintes faixas de receita bruta dentro do Ano-Calendário"
msg_micro2 = " a)Até R$ 60.000,00 : 0,6%"
msg_micro3 = " b)de  R$ 60.000,01 a R$ 90.000,00 : 0,8%"
msg_micro4 = " c)de  R$ 90.000,01 a R$ 120.000,00 : 1,0%"
msg_micro5 = "II- Para as Microempresas, contribuintes do ICMS e do ISS, que tenham as seguintes faixas de receita bruta dentro do Ano-Calendário"
msg_micro6 = " a)Até R$ 60.000,00 : 0,3%"
msg_micro7 = " b)de  R$ 60.000,01 a R$ 90.000,00 : 0,4%"
msg_micro8 = " c)de  R$ 90.000,01 a R$ 120.000,00 : 0,5%"

'que tenham as seguintes faixas de receita br

Select Case g_condicao
  Case Is = 1, 7
      retorno_critica = critica_apuracao_icms_normal(txt_fatura_normal, grid_ApurICMS, erro)
   Case Is = 5, 6
      retorno_critica = critica_apuracao_icms_EPP(msk_fat_mensal_frm4, msk_aliquota_frm4, msk_imposto_frm4, erro)
   Case Is = 2
      retorno_critica = critica_apuracao_icms_micro(grid_fatura, linha, erro)
   Case Is = 3, 4
      retorno_critica = critica_apuracao_icms_EPP(msk_fat_mensal_frm4, msk_aliquota_frm4, msk_imposto_frm4, erro)
      If retorno_critica Then
         retorno_critica = critica_apuracao_icms_eppmaior(msk_fat_mensal_frm4, grid_ApurICMS, erro)
      End If
End Select
         
         If Not retorno_critica Then
            Select Case erro
               Case Is = 1
                mensa1 = "Erro na Ficha de Apuração do ICMS o Faturamento do Mês menor que o Total de Campos do Demonstrativo das Saídas."
                mensa2 = "Total dos Campos Vendas para o Estado + Vendas para outros Estados + Vendas para o Exterior = " & Format(g_tot_val_para_faturamento, "###,###,##0.#0")
                msgerro = mensa1 & Chr(13) & mensa2
               Case Is = 2
                 
               Case Is = 3
             
               Case Is = 4
             
                  
               Case Is = 5
           
               Case Is = 6
           
               Case Is = 7
                  msgerro = "Erro - Débito Por Saídas da Apuração do ICMS deve ser maior ou igual ao Total do Imposto Debitado(Saidas)."
               Case Is = 8
                  msgerro = "Erro - Crédito por Entradas da Apuração do ICMS deve ser maior ou igual ao Total do Imposto Creditado(Entradas)."
               Case Is = 9
                  msgerro = "Erro na Alíquota da Apuração do ICMS."
               Case Is = 13
                  msgerro = "Erro na Apuração do ICMS na Alíquota de " & converte_em_mes(linha) & Chr(13) & msg_micro1 & Chr(13) & msg_micro2 & Chr(13) & msg_micro3 & Chr(13) & msg_micro4 & Chr(13) & msg_micro5 & Chr(13) & msg_micro6 & Chr(13) & msg_micro7 & Chr(13) & msg_micro8
                Case Is = 15
                  msgerro = "Erro na Apuração do ICMS Valor do Faturamento Obrigatório em " & converte_em_mes(linha)
               Case Is = 17
                  msgerro = "Erro na Apuração do ICMS Valor do Faturamento Obrigatório."
            End Select
            Navega_Ficha(1, 2) = 0
            critica_apuracao_icms = False
            Exit Function
        End If
         
End Function

Function CRITICA_DEM_ENTRADAS() As Integer
Dim motivo_epp As String
Dim msgtotal As String
msgerro = ""
motivo_percentual = "Deve ser: O imposto Creditado deve ser até 25% da Base de Calculo."

motivo_entradas = "Deve ser: Valores Contábeis = Base de Cálculo + Isentas não Tributadas + Outras."
motivo_epp = "Deve ser: Valores Contábeis >= Base de Cálculo + Isentas não Tributadas + Outras."

On Error GoTo erro_xx1

CRITICA_DEM_ENTRADAS = True
Dim cod_erro As Integer
If g_condicao = 1 Or g_condicao = 3 Or g_condicao = 4 Or g_condicao = 7 Then
        retorno_critica = critica_detalhe_demonstrativos(GridArray, linha_erro, cod_erro, "E", msgtotal)
        If Not retorno_critica Then
          
          Select Case cod_erro
             
          Case Is = 13
                 msgerro = msgtotal
           
          Case Is = 15
                 msgerro = "Erro de crítica no Demonstrativo das Operações (Entradas) - " & converte_linha_texto(linha_erro, "E") & Chr(13) & motivo_percentual
            
          Case Else
                If g_condicao = 1 Or g_condicao = 3 Or g_condicao = 4 Or g_condicao = 7 Then
                   msgerro = "Erro de crítica no Demonstrativo das Operações (Entradas)  - " & converte_linha_texto(linha_erro, "E") & Chr(13) & motivo_epp
                Else
                   msgerro = "Erro de crítica no Demonstrativo das Operações (Entradas)  - " & converte_linha_texto(linha_erro, "E") & Chr(13) & motivo_entradas
                End If
          
          End Select
          CRITICA_DEM_ENTRADAS = False
          Exit Function
        End If
    
Else
       If g_condicao = 2 Or g_condicao = 5 Or g_condicao = 6 Then
          retorno_critica = critica_demonstrativos_EPP_micro(GridArray)
          If Not retorno_critica Then
             msgerro = "Erro no Total no Demonstrativo das Operações (Entradas)."
             CRITICA_DEM_ENTRADAS = False
             Exit Function
          End If
       End If
End If


erro_xx1:
If Err <> 0 Then
     MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
    Resume Next
End If
      
End Function

Function CRITICA_DEM_saidas() As Integer
Dim motivo_epp As String
Dim msgtotal As String
motivo_percentual = "Deve ser: O imposto Debitado deve ser até 25% da Base de Calculo."
motivo_entradas = "Deve ser: Valores Contábeis = Base de Cálculo + Isentas não Tributadas + Outras."
motivo_epp = "Deve ser: Valores Contábeis >= Base de Cálculo + Isentas não Tributadas + Outras."
msgerro = ""
On Error GoTo erro_xx2

Dim cod_erro As Integer
CRITICA_DEM_saidas = True

         retorno_critica = critica_detalhe_demonstrativos(grid_saidas, linha_erro, cod_erro, "S", msgtotal)
         If Not retorno_critica Then
           Select Case cod_erro
            Case Is = 13
               msgerro = msgtotal
            Case Is = 15
               msgerro = "Erro de crítica no Demonstrativo das Operações (Saídas) - " & converte_linha_texto(linha_erro, "S") & Chr(13) & motivo_percentual
           Case Else
             If g_condicao = 3 Or g_condicao = 4 Then
                msgerro = "Erro de crítica no Demonstrativo das Operações (Saídas) - " & converte_linha_texto(linha_erro, "S") & Chr(13) & motivo_epp
             Else
                msgerro = "Erro de crítica no Demonstrativo das Operações (Saídas) - " & converte_linha_texto(linha_erro, "S") & Chr(13) & motivo_entradas
             End If
           End Select
            CRITICA_DEM_saidas = False
            Exit Function
         End If
         
   
         
erro_xx2:
If Err <> 0 Then
    MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
    Resume Next
End If
         
End Function

Function critica_detalhes() As Integer
critica_detalhes = True
msgerro = ""
On Error GoTo erro_xx3

            

         If SSTab1.TabEnabled(4) = True Then
            If Not verifica_preenchimento_dados_complementares(txt400, txt416, txt417, txt418, txt419, txt420, txt421, txt422, txt423, txt424, grid_DetDesp) Then
               msgerro = "Erro nos Dados do Exercício Anterior - pelo menos uma informação deve estar preenchida."
               critica_detalhes = False
               Exit Function
            End If
         End If
   
         If Not critica_livro_registro_inicial(txt416, txt417, txt418, txt419) Then
            msgerro = "Erro no Total do Estoque Inicial do Resumo."
        
            critica_detalhes = False
            Exit Function
         End If
         If Not critica_livro_registro_final(txt420, txt421, txt422, txt423) Then
            msgerro = "Erro no Total do Estoque Final do Resumo."
        
            critica_detalhes = False
            Exit Function
         End If
         If Not critica_detalhe_despesas(grid_DetDesp) Then
            msgerro = "Erro no Total do Detalhamento de Despesas."
        
            critica_detalhes = False
            Exit Function
         End If
        
erro_xx3:
If Err <> 0 Then
     MsgBox "Error Message :" & Str(Err) & " " & Error
    Resume Next
End If
        
End Function


Function critica_recolhimento() As Integer
Dim MeuValor As Double
Dim a As Integer
msgerro = ""

For a = 1 To 2
    MeuValor = acumula_grid_recolhimento(GridArray, a)
    If CCur(MeuValor) <> soma Then
      msgerro = "Erro no Total do Recolhimento da coluna " & IIf(a = 1, "Recolhidos", "A Recolher") & " "
      critica_recolhimento = False
      Exit Function
    Else
      critica_recolhimento = True
    End If
Next



End Function

Function deleta_itens_dem_entrada(item_inicial As Integer, item_final As Integer)
Dim ind As Integer
On Error GoTo erro_xx4

deleta_itens_dem_entrada = True
For ind = item_inicial To item_final

  If Not verifica_deleta_item(ind) Then
     deleta_itens_dem_entrada = False
     Exit Function
  End If


Next
erro_xx4:
If Err <> 0 Then
     MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
    Resume Next
End If

End Function

Private Function DeleteRow(bookm As Variant) As Boolean
    Dim Index As Long
    Dim i As Integer
    Dim j As Integer
    
    Index = IndexFromBookmark(bookm, False)
    If Index < 0 Or Index >= MaxRow Then
        DeleteRow = False
        Exit Function
    End If

    MaxRow = MaxRow - 1

'Shift the data in the array
    For i% = Index To MaxRow - 1
        For j% = 0 To MaxCol - 1
            GridArray(j%, i%) = GridArray(j%, i% + 1)
        Next j%
    Next i%
    
    ReDim Preserve GridArray(0 To MaxCol - 1, 0 To MaxRow - 1)
    DeleteRow = True
End Function

Private Function DeleteRow_DetDesp(bookm As Variant) As Boolean
    Dim Index As Long
    Dim i As Integer
    Dim j As Integer
    
    Index = IndexFromBookmark_DetDesp(bookm, False)
    If Index < 0 Or Index >= MaxRow_DetDesp Then
        DeleteRow_DetDesp = False
        Exit Function
    End If

    MaxRow_DetDesp = MaxRow_DetDesp - 1

'Shift the data in the array
    For i% = Index To MaxRow_DetDesp - 1
        For j% = 0 To MaxCol_DetDesp - 1
            grid_DetDesp(j%, i%) = grid_DetDesp(j%, i% + 1)
        Next j%
    Next i%
    
    ReDim Preserve grid_DetDesp(0 To MaxCol_DetDesp - 1, 0 To MaxRow_DetDesp - 1)
    DeleteRow_DetDesp = True
    
End Function

Private Function DeleteRow_fat(bookm As Variant) As Boolean
    Dim Index As Long
    Dim i As Integer
    Dim j As Integer
    
    Index = IndexFromBookmark_fat(bookm, False)
    If Index < 0 Or Index >= MaxRow_fat Then
        DeleteRow_fat = False
        Exit Function
    End If

    MaxRow_fat = MaxRow_fat - 1

'Shift the data in the array
    For i% = Index To MaxRow_fat - 1
        For j% = 0 To MaxCol_fat - 1
            grid_fatura(j%, i%) = grid_fatura(j%, i% + 1)
        Next j%
    Next i%
    
    ReDim Preserve grid_fatura(0 To MaxCol_fat - 1, 0 To MaxRow_fat - 1)
    DeleteRow_fat = True
End Function

Private Function DeleteRow_ICMS(bookm As Variant) As Boolean
    Dim Index As Long
    Dim i As Integer
    Dim j As Integer
    
    Index = IndexFromBookmark_ICMS(bookm, False)
    If Index < 0 Or Index >= MaxRow_ICMS Then
        DeleteRow_ICMS = False
        Exit Function
    End If

    MaxRow_ICMS = MaxRow_ICMS - 1

'Shift the data in the array
    For i% = Index To MaxRow_ICMS - 1
        For j% = 0 To MaxCol_ICMS - 1
            grid_ApurICMS(j%, i%) = grid_ApurICMS(j%, i% + 1)
        Next j%
    Next i%
    
    ReDim Preserve grid_ApurICMS(0 To MaxCol_ICMS - 1, 0 To MaxRow_ICMS - 1)
    DeleteRow_ICMS = True
    
End Function

Private Function DeleteRow_rec(bookm As Variant) As Boolean
    Dim Index As Long
    Dim i As Integer
    Dim j As Integer
    
    Index = IndexFromBookmark_rec(bookm, False)
    If Index < 0 Or Index >= MaxRow_rec Then
        DeleteRow_rec = False
        Exit Function
    End If

    MaxRow_rec = MaxRow_rec - 1

'Shift the data in the array
    For i% = Index To MaxRow_rec - 1
        For j% = 0 To MaxCol_rec - 1
            grid_recolim(j%, i%) = grid_recolim(j%, i% + 1)
        Next j%
    Next i%
    
    ReDim Preserve grid_recolim(0 To MaxCol_rec - 1, 0 To MaxRow_rec - 1)
    DeleteRow_rec = True
End Function

Private Function DeleteRow_s(bookm As Variant) As Boolean
    Dim Index As Long
    Dim i As Integer
    Dim j As Integer
    
    Index = IndexFromBookmark_s(bookm, False)
    If Index < 0 Or Index >= MaxRow_s Then
        DeleteRow_s = False
        Exit Function
    End If

    MaxRow_s = MaxRow_s - 1

'Shift the data in the array
    For i% = Index To MaxRow_s - 1
        For j% = 0 To MaxCol_s - 1
            grid_saidas(j%, i%) = grid_saidas(j%, i% + 1)
        Next j%
    Next i%
    
    ReDim Preserve grid_saidas(0 To MaxCol_s - 1, 0 To MaxRow_s - 1)
    DeleteRow_s = True
End Function

Sub desativa_label_pisca_pisca()

lbl_pisca_epp.Visible = False
lbl_epp_entr.Visible = False
LBL_APUR_PISCA.Visible = False
End Sub

Sub Finalizar_Declara()

Dim Erro1 As Variant, Erro2 As Variant, Erro3 As Variant, Erro4 As Variant, Erro5 As Variant

Erro1 = False
Erro2 = False
Erro3 = False
Erro4 = False
Erro5 = False

Entradas:
    
    If Not CRITICA_DEM_ENTRADAS Then
        Screen.MousePointer = 0
        Dec007.Image1(0).Picture = Dec007.ImageErro.Picture
        Dec007.Label3(0).Caption = msgerro
    Else
        Erro1 = True
        Dec007.Label3(0).Caption = "OK "
    End If

Saidas:
        If Not CRITICA_DEM_saidas Then
            Screen.MousePointer = 0
            Dec007.Image1(1).Picture = Dec007.ImageErro.Picture
            Dec007.Label3(1).Caption = msgerro
        Else
            Erro2 = True
            Dec007.Label3(1).Caption = "OK "
        End If

Apuracao_do_ICMS:
         If Not critica_apuracao_icms Then
            Screen.MousePointer = 0
            Dec007.Image1(2).Picture = Dec007.ImageErro.Picture
            Dec007.Label3(2).Caption = msgerro
         Else
            Erro3 = True
            Dec007.Label3(2).Caption = "OK "
         End If
      
Recolhimento:
           If Not critica_recolhimento Then
                Screen.MousePointer = 0
                Dec007.Image1(3).Picture = Dec007.ImageErro.Picture
                Dec007.Label3(3).Caption = msgerro
           Else
                Erro4 = True
                Dec007.Label3(3).Caption = "OK "
           End If

Detalhamento:

      If Not critica_detalhes Then
           Screen.MousePointer = 0
           Dec007.Image1(4).Picture = Dec007.ImageErro.Picture
           Dec007.Label3(4).Caption = msgerro
      Else
           Erro5 = True
           Dec007.Label3(4).Caption = "OK "
      End If


Dec007.Show 1
    
End Sub

Private Function GetNewBookmark() As Variant
On Error Resume Next
    ReDim Preserve GridArray(0 To MaxCol - 1, 0 To MaxRow)
    GetNewBookmark = MakeBookmark(MaxRow)
    MaxRow = MaxRow + 1
    
End Function

Private Function GetNewBookmark_DetDesp() As Variant
On Error Resume Next
    ReDim Preserve grid_DetDesp(0 To MaxCol_DetDesp - 1, 0 To MaxRow_DetDesp)
    GetNewBookmark_DetDesp = MakeBookmark(MaxRow_DetDesp)
    MaxRow_DetDesp = MaxRow_DetDesp + 1

End Function

Private Function GetNewBookmark_fat() As Variant
On Error Resume Next
    ReDim Preserve grid_fatura(0 To MaxCol_fat - 1, 0 To MaxRow_fat)
    GetNewBookmark_fat = MakeBookmark(MaxRow_fat)
    MaxRow_fat = MaxRow_fat + 1
    
End Function
Private Function GetNewBookmark_ICMS() As Variant
On Error Resume Next
    ReDim Preserve grid_ApurICMS(0 To MaxCol_ICMS - 1, 0 To MaxRow_ICMS)
    GetNewBookmark_ICMS = MakeBookmark(MaxRow_ICMS)
    MaxRow_ICMS = MaxRow_ICMS + 1

End Function
Private Function GetNewBookmark_rec() As Variant
On Error Resume Next
    ReDim Preserve grid_recolim(0 To MaxCol_rec - 1, 0 To MaxRow_rec)
    GetNewBookmark_rec = MakeBookmark(MaxRow_rec)
    MaxRow_rec = MaxRow_rec + 1
    
End Function

Private Function GetNewBookmark_s() As Variant
On Error Resume Next
    ReDim Preserve grid_saidas(0 To MaxCol_s - 1, 0 To MaxRow_s)
    GetNewBookmark_s = MakeBookmark(MaxRow_s)
    MaxRow_s = MaxRow_s + 1
    
End Function
Private Function GetRelativeBookmark(bookm As Variant, relpos As Integer) As Variant
    Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark(bookm, False) + relpos
    If Index < 0 Or Index >= MaxRow Then
        GetRelativeBookmark = Null
    Else
        GetRelativeBookmark = MakeBookmark(Index)
    End If
End Function
Private Function GetRelativeBookmark_DetDesp(bookm As Variant, relpos As Integer) As Variant
      Dim Index As Long
On Error Resume Next
          Index = IndexFromBookmark_DetDesp(bookm, False) + relpos
          If Index < 0 Or Index >= MaxRow_DetDesp Then
              GetRelativeBookmark_DetDesp = Null
          Else
              GetRelativeBookmark_DetDesp = MakeBookmark(Index)
          End If
End Function

Private Function GetRelativeBookmark_fat(bookm As Variant, relpos As Integer) As Variant
    Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark_fat(bookm, False) + relpos
    If Index < 0 Or Index >= MaxRow_fat Then
        GetRelativeBookmark_fat = Null
    Else
        GetRelativeBookmark_fat = MakeBookmark(Index)
    End If
End Function

Private Function GetRelativeBookmark_ICMS(bookm As Variant, relpos As Integer) As Variant
Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark_ICMS(bookm, False) + relpos
    If Index < 0 Or Index >= MaxRow_ICMS Then
        GetRelativeBookmark_ICMS = Null
    Else
        GetRelativeBookmark_ICMS = MakeBookmark(Index)
    End If
End Function

Private Function GetRelativeBookmark_rec(bookm As Variant, relpos As Integer) As Variant
    Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark_rec(bookm, False) + relpos
    If Index < 0 Or Index >= MaxRow_rec Then
        GetRelativeBookmark_rec = Null
    Else
        GetRelativeBookmark_rec = MakeBookmark(Index)
    End If
End Function

Private Function GetRelativeBookmark_s(bookm As Variant, relpos As Integer) As Variant
    Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark_s(bookm, False) + relpos
    If Index < 0 Or Index >= MaxRow_s Then
        GetRelativeBookmark_s = Null
    Else
        GetRelativeBookmark_s = MakeBookmark(Index)
    End If
End Function

Private Function GetUserData(bookm As Variant, colm As Integer) As Variant
     
        Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark(bookm, False)
    If Index < 0 Or Index >= MaxRow Or colm < 0 Or colm >= MaxCol Then
        GetUserData = Null
    Else
        GetUserData = GridArray(colm, Index)
    End If

End Function

Private Function GetUserData_DetDesp(bookm As Variant, colm As Integer) As Variant
 Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark_DetDesp(bookm, False)
    If Index < 0 Or Index >= MaxRow_DetDesp Or colm < 0 Or colm >= MaxCol_DetDesp Then
        GetUserData_DetDesp = Null
    Else
        GetUserData_DetDesp = grid_DetDesp(colm, Index)
    End If
    
End Function

Private Function GetUserData_fat(bookm As Variant, colm As Integer) As Variant
    Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark_fat(bookm, False)
    If Index < 0 Or Index >= MaxRow_fat Or colm < 0 Or colm >= MaxCol_fat Then
        GetUserData_fat = Null
    Else
        GetUserData_fat = grid_fatura(colm, Index)
    End If

     
  

End Function

Private Function GetUserData_ICMS(bookm As Variant, colm As Integer) As Variant
 Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark_ICMS(bookm, False)
    If Index < 0 Or Index >= MaxRow_ICMS Or colm < 0 Or colm >= MaxCol_ICMS Then
        GetUserData_ICMS = Null
    Else
       GetUserData_ICMS = grid_ApurICMS(colm, Index)
    End If
    
End Function

Private Function GetUserData_rec(bookm As Variant, colm As Integer) As Variant
    Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark_rec(bookm, False)
    If Index < 0 Or Index >= MaxRow_rec Or colm < 0 Or colm >= MaxCol_rec Then
        GetUserData_rec = Null
    Else
        GetUserData_rec = grid_recolim(colm, Index)
    End If

End Function
Private Function GetUserData_s(bookm As Variant, colm As Integer) As Variant
    Dim Index As Long
On Error Resume Next
    Index = IndexFromBookmark_s(bookm, False)
    If Index < 0 Or Index >= MaxRow_s Or colm < 0 Or colm >= MaxCol_s Then
        GetUserData_s = Null
    Else
        GetUserData_s = grid_saidas(colm, Index)
    End If
    
    
    

End Function
Private Function grava_item(item As Integer, valor As String) As Boolean
   On Error GoTo erro_grava_item

grava_item = True
  Query1 = "select * from tab_detalhe_declaracao where tdd_num_inscricao = '" & Trim(Str(msk_inscricao)) & "'"
  Query1 = Query1 & " and tdd_periodo_inicial = '" & g_periodo_inicial & "'"
  Query1 = Query1 & " and tdd_substitutiva = " & Val(g_tipo_decla)

  Query1 = Query1 & " and tdd_item = " & item

Set restauracontr = dbgim.OpenRecordset(Query1, dbOpenDynaset)
  If restauracontr.EOF And restauracontr.BOF Then
        query = "insert into tab_detalhe_declaracao"
        query = query & "(tdd_num_inscricao, tdd_periodo_inicial,  tdd_substitutiva, "
        query = query & "tdd_item, tdd_valor )"
        query = query & "values('" & Trim(Str(msk_inscricao)) & "','" & g_periodo_inicial & "','"
        query = query & Val(g_tipo_decla) & "',"
        query = query & item & "," & valor & ")"
   Else
        query = "update tab_detalhe_declaracao set "
        query = query & "tdd_valor = " & valor
        query = query & " where tdd_num_inscricao = '" & Trim(Str(msk_inscricao)) & "'"
        query = query & " and tdd_periodo_inicial = '" & g_periodo_inicial & "'"
        query = query & " and tdd_substitutiva = " & Val(g_tipo_decla)
        query = query & " and tdd_item = " & item
    End If
      restauracontr.Close
      
      If Not EjecutarQuery(query) Then
         Screen.MousePointer = 0
         grava_item = False
         Exit Function
      End If
   
 Exit Function

erro_grava_item:

     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next

End Function

Sub gravar_declaracao(tipo As Integer)

Dim sit As String
     On Error GoTo erro_grava_decla

Dim indficha As Integer


    sit = "I"
    
   query = "select * from tab_declaracao where tde_num_inscricao = '" & Trim(Str(msk_inscricao)) & "'"
   query = query & " and tde_periodo_inicial =  '" & g_periodo_inicial & "'"
   query = query & " and tde_substitutiva =  " & tipo

   Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
   
If restauracontr.EOF And restauracontr.BOF Then
    query = "insert into tab_declaracao"
    query = query & "(tde_num_inscricao, tde_periodo_inicial, tde_substitutiva, "
    query = query & "tde_Estado , tde_aliquota_m17, tde_condicao ) "
    
    query = query & "values('" & Trim(Str(msk_inscricao)) & "','" & g_periodo_inicial & "','"
    query = query & tipo & "','"
    query = query & sit & "'," & g_aliquota_Maior17 & "," & g_condicao & ")"
    
    If Not EjecutarQuery(query) Then
    End If
'Else
'   query = "update tab_declaracao set "
'   query = query & "tde_estado = '" & sit & "'"
'   query = query & " where tde_num_inscricao = '" & Trim(Str(msk_inscricao)) & "'"
'   query = query & " and tde_periodo_inicial = '" & g_periodo_inicial & "'"
'   query = query & " and tde_substitutiva = " & Val(g_tipo_decla)
'   query = query & " and tde_aliquota_m17 = " & g_aliquota_Maior17
'   query = query & " and tde_condicao = " & g_condicao

''     If Not EjecutarQuery(query) Then
 '   End If
End If


restauracontr.Close
  Exit Sub

erro_grava_decla:

     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next

End Sub

Private Function IndexFromBookmark(bookm As Variant, ReadPriorRows As Boolean) As Long
On Error Resume Next
    If IsNull(bookm) Then
        If ReadPriorRows Then
            IndexFromBookmark = MaxRow
        Else
            IndexFromBookmark = -1
        End If
    Else
        Dim Index As Long
        Index = Val(bookm)
        If Index < 0 Or Index >= MaxRow Then Index = -2000
            IndexFromBookmark = Index
    End If
End Function

Private Function IndexFromBookmark_DetDesp(bookm As Variant, ReadPriorRows As Boolean) As Long
On Error Resume Next
   If IsNull(bookm) Then
        If ReadPriorRows Then
            IndexFromBookmark_DetDesp = MaxRow_DetDesp
        Else
            IndexFromBookmark_DetDesp = -1
        End If
    Else
        Dim Index As Long
        Index = Val(bookm)
        If Index < 0 Or Index > MaxRow_DetDesp Then Index = -2000
            IndexFromBookmark_DetDesp = Index
    End If

End Function

Private Function IndexFromBookmark_fat(bookm As Variant, ReadPriorRows As Boolean) As Long
On Error Resume Next
    If IsNull(bookm) Then
        If ReadPriorRows Then
            IndexFromBookmark_fat = MaxRow_fat
        Else
            IndexFromBookmark_fat = -1
        End If
    Else
        Dim Index As Long
        Index = Val(bookm)
        If Index < 0 Or Index >= MaxRow_fat Then Index = -2000
            IndexFromBookmark_fat = Index
    End If
End Function
Private Function IndexFromBookmark_ICMS(bookm As Variant, ReadPriorRows As Boolean) As Long
On Error Resume Next
   If IsNull(bookm) Then
        If ReadPriorRows Then
            IndexFromBookmark_ICMS = MaxRow_ICMS
        Else
            IndexFromBookmark_ICMS = -1
        End If
    Else
        Dim Index As Long
        Index = Val(bookm)
        If Index < 0 Or Index >= MaxRow_ICMS Then Index = -2000
         IndexFromBookmark_ICMS = Index
    End If

End Function

Private Function IndexFromBookmark_rec(bookm As Variant, ReadPriorRows As Boolean) As Long
On Error Resume Next
    If IsNull(bookm) Then
        If ReadPriorRows Then
            IndexFromBookmark_rec = MaxRow_rec
        Else
            IndexFromBookmark_rec = -1
        End If
    Else
        Dim Index As Long
        Index = Val(bookm)
        If Index < 0 Or Index > MaxRow_rec Then Index = -2000
            IndexFromBookmark_rec = Index
    End If
End Function
Private Function IndexFromBookmark_s(bookm As Variant, ReadPriorRows As Boolean) As Long
On Error Resume Next
    If IsNull(bookm) Then
        If ReadPriorRows Then
            IndexFromBookmark_s = MaxRow_s
        Else
            IndexFromBookmark_s = -1
        End If
    Else
        Dim Index As Long
        Index = Val(bookm)
        If Index < 0 Or Index >= MaxRow_s Then Index = -2000
            IndexFromBookmark_s = Index
    End If
End Function
Sub inicia_Dbgrid()
GRID_DEM_ENTRADAS.Columns(0).Width = 2400
GRID_DEM_ENTRADAS.Columns(1).Width = 1250
GRID_DEM_ENTRADAS.Columns(2).Width = 1250
GRID_DEM_ENTRADAS.Columns(3).Width = 1250
GRID_DEM_ENTRADAS.Columns(4).Width = 1250
GRID_DEM_ENTRADAS.Columns(5).Width = 1250

grid_dem_saidas.Columns(0).Width = 2400
grid_dem_saidas.Columns(1).Width = 1250
grid_dem_saidas.Columns(2).Width = 1250
grid_dem_saidas.Columns(3).Width = 1250
grid_dem_saidas.Columns(4).Width = 1250
grid_dem_saidas.Columns(5).Width = 1250
'grid_dem_saidas.Width = 9175

End Sub

Sub limpa_capos_txt()

lbl_tipo_decl = ""
      txt_fatura_normal = ""
      msk_aliquota_frm4 = ""
      msk_fat_mensal_frm4 = ""
      msk_imposto_frm4 = ""
      
      
     
      txt201 = ""
      txt202 = ""
      txt203 = ""
      txt204 = ""
      txt205 = ""
      txt206 = ""
      txt207 = ""
      txt208 = ""
      txt209 = ""
      txt210 = ""
      txt211 = ""
      txt212 = ""
      txt213 = ""
      txt214 = ""
      txt215 = ""
      txt216 = ""
      txt217 = ""
      txt218 = ""
      txt236 = ""
      txt237 = ""
      txt238 = ""
      txt229 = ""
      txt242 = ""
      txt301.Text = ""
      txt300 = ""
      txt302 = ""
      txt303 = ""
'      txt314 = ""

      txt400 = ""
      txt415 = ""
      txt416 = ""

      txt417 = ""
      txt418 = ""
      txt419 = ""
      txt420 = ""
      txt421 = ""
      txt422 = ""
      txt423 = ""
      txt424 = ""

End Sub

Sub limpa_grids_entrada()

  ReDim GridArray(0 To 1, 0 To 12)
   monta_cabecalho_grid
   MaxCol = 2
   MaxRow = 13

   grid_Entra_MicroEPP.Refresh
  
    ReDim GridArray(0 To 5, 0 To 12)
    MaxCol = 6
    MaxRow = 13
    GRID_DEM_ENTRADAS.Refresh
     
End Sub

Sub limpa_tela_Det()

On Error GoTo erro_limpar

         ReDim GridArray(0, 0) As Variant  'demonstrativo das oper. entradas
         ReDim grid_saidas(0, 0) As Variant 'demonstrativo das opr saidas
         ReDim grid_ApurICMS(0, 0) As Variant 'Apuração do ICMS
         ReDim grid_recolim(0, 0) As Variant 'recolhimento do periodo
         ReDim grid_fatura(0, 0) As Variant 'faturamento frame 5
         ReDim grid_DetDesp(0, 0) As Variant 'faturamento frame 5
         
         cmb_tipo.Clear
         cmb_tipo.AddItem "Sim"
         cmb_tipo.AddItem "Não"
         
         erro_periodo = False
          
         If gformulario <> 1 Then
            g_num_inscricao = ""
            g_periodo_inicial = ""
            g_tipo_decla = 0
         End If
         
         msk_inscricao = g_num_inscricao
         txt_periodo_inicial = g_periodo_inicial
         txt_razao = ""
         txtCondContr = ""
         
        limpa_capos_txt
        Situacao.Caption = ""
         TestaCampoPreench
                  
         cargar_data
         cargar_datos_saidas
         cargar_grid_DetDesp
         cargar_grid_recolhimento
         cargar_grid_ApurICMS
         cargar_grid_faturamento

         grid_dem_saidas.MarqueeStyle = 5
         grid_recolhimento.MarqueeStyle = 5
'         grdApurICMS.MarqueeStyle = 5
         grdDetDesp.MarqueeStyle = 5
         grid_faturamento.MarqueeStyle = 5
         GRID_DEM_ENTRADAS.MarqueeStyle = 5
         
          AtualizaGrades
      
      SSTab1.Tab = 0
      SSTab1.Enabled = False
      
Exit Sub

erro_limpar:

     MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next

End Sub

Private Function MakeBookmark(Index As Long) As Variant
      MakeBookmark = Str$(Index)
End Function


Sub monta_cabecalho_ApurICMS()
On Error Resume Next
   grid_ApurICMS(0, 0) = "Débito do Imposto: Por Saídas ou Prestações"
   grid_ApurICMS(0, 1) = "                              Outros Débitos"
   grid_ApurICMS(0, 2) = "                              Estornos de Créditos"
   grid_ApurICMS(0, 3) = "                              Total do Débito"
   grid_ApurICMS(0, 4) = "Crédito do Imposto: Por entradas ou Aquisições"
   grid_ApurICMS(0, 5) = "                               Outros Créditos"
   grid_ApurICMS(0, 6) = "                               Estorno do Débito"
   grid_ApurICMS(0, 7) = "                               Saldo credor do período anterior"
   grid_ApurICMS(0, 8) = "                               Total do Crédito"
   grid_ApurICMS(0, 9) = "Apuração do Imposto: Saldo Devedor"
   grid_ApurICMS(0, 10) = "                                   Deduções"
   grid_ApurICMS(0, 11) = "                                   Imposto a Recolher"
   grid_ApurICMS(0, 12) = "                                   Saldo credor a transferir"

End Sub

Sub monta_cabecalho_DetDesp()
On Error Resume Next
   grid_DetDesp(0, 0) = "Pró-Labore"
   grid_DetDesp(0, 1) = "Comissões, Salários,Ordenados"
   grid_DetDesp(0, 2) = "Combustíveis e Lubrificantes"
   grid_DetDesp(0, 3) = "Encargos Sociais"
   grid_DetDesp(0, 4) = "Tributos Federais"
   grid_DetDesp(0, 5) = "Tributos Estaduais"
   grid_DetDesp(0, 6) = "Tributos Municipais"
   grid_DetDesp(0, 7) = "Água, Luz, Telefone"
   grid_DetDesp(0, 8) = "Aluguéis"
   grid_DetDesp(0, 9) = "Serviços Profissionais"
   grid_DetDesp(0, 10) = "Seguros"
   grid_DetDesp(0, 11) = "Fretes e carretos"
   grid_DetDesp(0, 12) = "Despesas Financeiras"
   grid_DetDesp(0, 13) = "Outras Despesas"
   grid_DetDesp(0, 14) = "TOTAL"

End Sub

Sub monta_cabecalho_FAt()
On Error Resume Next

   If txtCondContr = "Microempresa" Then
      Select Case Mid$(Trim(g_periodo_inicial), 5, 2)
       Case "03"
         grid_fatura(0, 0) = "Janeiro"
         grid_fatura(0, 1) = "Fevereiro"
         grid_fatura(0, 2) = "Março"
         grid_fatura(0, 3) = "Total"
       Case "06"
         grid_fatura(0, 0) = "Abril"
         grid_fatura(0, 1) = "Maio"
         grid_fatura(0, 2) = "Junho"
         grid_fatura(0, 3) = "Total"
      Case "09"
         grid_fatura(0, 0) = "Julho"
         grid_fatura(0, 1) = "Agosto"
         grid_fatura(0, 2) = "Setembro"
         grid_fatura(0, 3) = "Total"
      Case "12"
         grid_fatura(0, 0) = "Outubro"
         grid_fatura(0, 1) = "Novembro"
         grid_fatura(0, 2) = "Dezembro"
         grid_fatura(0, 3) = "Total"
    End Select
  
   End If
   
   
End Sub
Sub monta_cabecalho_grid()
On Error Resume Next

          GridArray(0, 0) = "Compras no Estado"
          GridArray(0, 1) = "Transferencias no Estado"
          GridArray(0, 2) = "Devoluções no Estado"
          GridArray(0, 3) = "Outras no Estado"
          GridArray(0, 4) = "Compras de outros Estados"
          GridArray(0, 5) = "Transf. de outros Estados"
          GridArray(0, 6) = "Devoluções de outros Estados"
          GridArray(0, 7) = "Outras de outros Estados"
          GridArray(0, 8) = "Compras no Exterior"
          GridArray(0, 9) = "Devoluções do Exterior"
          GridArray(0, 10) = "Outras entradas do Exterior"
          GridArray(0, 11) = "Totais de entradas"
          GridArray(0, 12) = ""

End Sub
Sub monta_cabecalho_recolhimento()
On Error Resume Next

Select Case g_condicao

Case Is = 1
       grid_recolim(0, 0) = "Normal"
       grid_recolim(0, 1) = "Substituição de Saídas"
       grid_recolim(0, 2) = "Substituição de Entradas"
       grid_recolim(0, 3) = "Antecipação Parcial"
       grid_recolim(0, 4) = "Diferencial de Alíquota"
       grid_recolim(0, 5) = "Importações"
       grid_recolim(0, 6) = "Total"

Case Is = 5, 6
       grid_recolim(0, 0) = "Pequeno Porte"
       grid_recolim(0, 1) = "Substituição de Saídas" '
       grid_recolim(0, 2) = "Substituição de Entradas"
       grid_recolim(0, 3) = "Antecipação Parcial"
       grid_recolim(0, 4) = "Diferencial de Alíquota"
       grid_recolim(0, 5) = "Importações"
       grid_recolim(0, 6) = "Total"

Case Is = 2
       grid_recolim(0, 0) = "Microempresa"
       grid_recolim(0, 1) = "Substituição de Saídas"
       grid_recolim(0, 2) = "Substituição de Entradas"
       grid_recolim(0, 3) = "Antecipação Parcial"
       grid_recolim(0, 4) = "Diferencial de Alíquota"
       grid_recolim(0, 5) = "Importações"
       grid_recolim(0, 6) = "Total"

Case Is = 3, 4, 7
       grid_recolim(0, 0) = "Normal"
       grid_recolim(0, 1) = "Pequeno Porte"
       grid_recolim(0, 2) = "Substituição de Saídas"
       grid_recolim(0, 3) = "Substituição de Entradas"
       grid_recolim(0, 4) = "Antecipação Parcial"
       grid_recolim(0, 5) = "Diferencial de Alíquota"
       grid_recolim(0, 6) = "Importações"
       grid_recolim(0, 7) = "Total"

End Select
  
End Sub

Sub monta_grid_saidas()
On Error Resume Next
          grid_saidas(0, 0) = "Vendas para o Estado"
          grid_saidas(0, 1) = "Transferencias para o Estado"
          grid_saidas(0, 2) = "Devoluções para o Estado"
          grid_saidas(0, 3) = "Outras para o Estado"
          grid_saidas(0, 4) = "Vendas para outros Estados"
          grid_saidas(0, 5) = "Transf. para outros Estados"
          grid_saidas(0, 6) = "Devoluções para outros Estados"
          grid_saidas(0, 7) = "Outras para outros Estados"
          grid_saidas(0, 8) = "Vendas para o Exterior"
          grid_saidas(0, 9) = "Devoluções para o Exterior"
          grid_saidas(0, 10) = "Outras para o Exterior"
          grid_saidas(0, 11) = "Totais de Saídas"
          grid_saidas(0, 12) = ""

End Sub

Sub posiciona_label_tipo_decl()

Select Case g_condicao
 Case Is = 1
              lbl_tipo_decl.Caption = "Normal"

 Case Is = 2
              lbl_tipo_decl.Caption = "Microempresa"
  Case Is = 3
              lbl_tipo_decl.Caption = "Com Alíquota de 1.5% S/Faturamento maior que 17%"
  Case Is = 5
              lbl_tipo_decl.Caption = "Com Alíquota de 1.5% S/Faturamento menor que 17%"

  Case Is = 4
              lbl_tipo_decl.Caption = "Com Alíquota de 2.0% S/Faturamento maior que 17%"
              
Case Is = 6
              lbl_tipo_decl.Caption = "Com Alíquota de 2.0% S/Faturamento menor que 17%"
Case Is = 7
   lbl_tipo_decl.Caption = "Com Faturamento maior que 240.000,00"

End Select
End Sub

Sub refresh_campos_grids()





End Sub

Private Function StoreUserData(bookm As Variant, colm As Integer, userval As Variant) As Boolean
    
    Dim Index As Long

    Index = IndexFromBookmark(bookm, False)
    If Index < 0 Or Index > MaxRow Or colm < 0 Or colm > MaxCol Then
        StoreUserData = False
    Else
        StoreUserData = True
        GridArray(colm, Index) = userval
    End If

    
End Function

Private Function StoreUserData_DetDesp(bookm As Variant, colm As Integer, userval As Variant) As Boolean

   Dim Index As Long
   
       Index = IndexFromBookmark_DetDesp(bookm, False)
       If Index < 0 Or Index > MaxRow_DetDesp Or colm < 0 Or colm > MaxCol_DetDesp Then
           StoreUserData_DetDesp = False
       Else
           StoreUserData_DetDesp = True
           grid_DetDesp(colm, Index) = userval
       End If

End Function

Private Function StoreUserData_fat(bookm As Variant, colm As Integer, userval As Variant) As Boolean
    Dim Index As Long

    Index = IndexFromBookmark_fat(bookm, False)
    If Index < 0 Or Index > MaxRow_fat Or colm < 0 Or colm > MaxCol_fat Then
        StoreUserData_fat = False
    Else
        StoreUserData_fat = True
        grid_fatura(colm, Index) = userval
    End If

    


End Function


Private Function StoreUserData_ICMS(bookm As Variant, colm As Integer, userval As Variant) As Boolean

   Dim Index As Long
   
       Index = IndexFromBookmark_ICMS(bookm, False)
       If Index < 0 Or Index > MaxRow_ICMS Or colm < 0 Or colm > MaxCol_ICMS Then
           StoreUserData_ICMS = False
       Else
           StoreUserData_ICMS = True
           grid_ApurICMS(colm, Index) = userval
       End If

End Function

Private Function StoreUserData_rec(bookm As Variant, colm As Integer, userval As Variant) As Boolean
    Dim Index As Long

    Index = IndexFromBookmark_rec(bookm, False)
    If Index < 0 Or Index > MaxRow_rec Or colm < 0 Or colm > MaxCol_rec Then
        StoreUserData_rec = False
    Else
        StoreUserData_rec = True
        grid_recolim(colm, Index) = userval
    End If


End Function

Private Function storeUserData_s(bookm As Variant, colm As Integer, userval As Variant) As Boolean
    Dim Index As Long

    Index = IndexFromBookmark_s(bookm, False)
    If Index < 0 Or Index > MaxRow_s Or colm < 0 Or colm > MaxCol_s Then
        storeUserData_s = False
    Else
        storeUserData_s = True
        grid_saidas(colm, Index) = userval
    End If
    
  

End Function
Sub TestaCampoPreench()
On Error Resume Next

   Dim recaliq As Recordset
   Dim campos As String
 
   AnoSis = Val(Year(Date))
   
   If Len(Trim(msk_inscricao)) <> 0 And Len(Trim(txt_periodo_inicial)) <> 0 And Len(Trim(cmb_tipo)) <> 0 Then
      cmd_imprimir.Enabled = True
      cmd_excluir.Enabled = True
      cmd_salvar.Enabled = True
      CmdFinalizar.Enabled = True
      Option1(0).Enabled = True
      Option1(1).Enabled = True
      fraIdentificação.Enabled = False
      campopreench = True
      SSTab1.Enabled = True
      Select Case g_condicao
         Case 1    'Normal
           fra_EntaNormal.Top = 650       'Entradas - Tela #1
           SSTab1.TabEnabled(1) = True         'Saídas - Tela #2
           Frame4.Visible = False
           frm_grid_fatu.Visible = False         'Faturamento - Tela #4 e 5
           Frame_tela3.Visible = True
            
            If Mid(Trim(txt_periodo_inicial), 1, 2) <> "03" And Len(g_num_cgc_cpf) <> 11 Then  'Março
               SSTab1.TabEnabled(4) = False     'GIA - Tela #8
               Navega_Ficha(0, 4) = 0
            
            ElseIf Mid(Trim(txt_periodo_inicial), 1, 2) = "03" And Len(g_num_cgc_cpf) = 11 Then  'Março
               SSTab1.TabEnabled(4) = False     'GIA - Tela #8
               Navega_Ficha(0, 4) = 0
            Else
               If g_ano_inicio_operacao <= AnoSis And Len(g_num_cgc_cpf) <> 11 Then
                    SSTab1.TabEnabled(4) = True      'GIA - Tela #8
                    Navega_Ficha(0, 4) = 1
               Else
                    SSTab1.TabEnabled(4) = False     'GIA - Tela #8
                    Navega_Ficha(0, 4) = 0
               End If
            End If
         Case 3, 4, 5, 6, 7 'EPP
            campos = "select * from tab_declaracao where tde_num_inscricao =   '" & g_num_inscricao & "'"
            campos = campos & " and [tde_periodo_inicial] = '" & g_periodo_inicial & "'"
            campos = campos & " and [tde_substitutiva] = " & g_tipo_decla
           
            Set recaliq = dbgim.OpenRecordset(campos, dbOpenDynaset)
            
            If Not recaliq.BOF And Not recaliq.EOF Then
               g_aliquota_Maior17 = recaliq.Fields("tde_aliquota_m17")
            End If
            recaliq.Close
            
            If g_condicao = 3 Or g_condicao = 4 Or g_condicao = 7 Then 'sim
               fra_EntaNormal.Top = 650      'Entradas - Tela #1A
               SSTab1.TabEnabled(1) = True         'tirei hoje 12/3/98
               Frame4.Visible = True
               frm_grid_fatu.Visible = False
               Frame_tela3.Visible = False
            Else 'nao
               fra_EntraMicroEPP.Top = 660    'Entradas - Tela #1A
               fra_EntraMicroEPP.Left = 75        'inclusao em 21/2/98 em funcao do 16 bits
               SSTab1.TabEnabled(1) = False         'Saídas - Tela #2
               Navega_Ficha(1, 1) = 1 'coloquei em 11/4/98 por ser saidas
               Frame_tela3.Visible = False
               Frame4.Visible = True
               frm_grid_fatu.Visible = False
            End If
  
            If Mid(Trim(txt_periodo_inicial), 1, 2) <> "03" And Len(g_num_cgc_cpf) = 11 Then    'Março
               SSTab1.TabEnabled(4) = False     'GIA - Tela #8
               Navega_Ficha(0, 4) = 0
            ElseIf Mid(Trim(txt_periodo_inicial), 1, 2) = "03" And Len(g_num_cgc_cpf) <> 11 Then 'Incluir o teste <= 15/05/98
                    SSTab1.TabEnabled(4) = True      'GIA - Tela #8
                    Navega_Ficha(0, 4) = 1
            Else
                    SSTab1.TabEnabled(4) = False     'GIA - Tela #8
                    Navega_Ficha(0, 4) = 0
            End If
         Case 2    'Microempresa
            fra_EntraMicroEPP.Top = 660   'Entradas - Tela #1A
            fra_EntraMicroEPP.Left = 75        'inclusao em 21/2/98 em funcao do 16 bits
            SSTab1.TabEnabled(1) = False         'Saídas - Tela #2
            fra_EntraMicroEPP.Refresh
            Frame_tela3.Visible = False
            Frame4.Visible = False              'Faturamento - Tela #4 e 5
            If Mid(Trim(txt_periodo_inicial), 1, 2) <> "03" Then    '1o.Semestre
               SSTab1.TabEnabled(4) = False     'GIA - Tela #8
               Navega_Ficha(0, 4) = 0
            Else
               If g_ano_inicio_operacao <= AnoSis And Len(g_num_cgc_cpf) <> 11 Then
                    SSTab1.TabEnabled(4) = True      'GIA - Tela #8
                    Navega_Ficha(0, 4) = 1
               Else
                    SSTab1.TabEnabled(4) = False     'GIA - Tela #8
                    Navega_Ficha(0, 4) = 0
               End If
            End If

      End Select
      SSTab1.Tab = 0
   Else
      cmd_imprimir.Enabled = False
      cmd_excluir.Enabled = False
      cmd_salvar.Enabled = False
      CmdFinalizar.Enabled = False
      Option1(0).Enabled = False
      Option1(1).Enabled = False
      fraIdentificação.Enabled = True
      SSTab1.Enabled = False
      campopreench = False
   End If
      
End Sub


Sub grava_grid_recolhimento()
Dim linha As Integer
Dim coluna As Integer
Dim item_a_gravar As Integer

On Error GoTo erro_grava_reco

If Not deleta_itens_dem_entrada(300, 325) Then Exit Sub
Select Case g_condicao
  Case Is = 1
    For coluna = 1 To 2
        For linha = 0 To 6
         If Trim(grid_recolim(coluna, linha)) <> "" And Trim(grid_recolim(coluna, linha)) <> "0,00" Then
            Select Case coluna
            Case 1
                If linha = 0 Then
                   item_a_gravar = linha + 304
                Else
                   item_a_gravar = linha + 306
                End If
            Case 2
                If linha = 0 Then
                   item_a_gravar = linha + 315
                Else
                   item_a_gravar = linha + 317
                   
                End If
            End Select
            valor_a_gravar = Str(grid_recolim(coluna, linha))
            If Not grava_item(item_a_gravar, valor_a_gravar) Then
                 Screen.MousePointer = 0
                 Exit Sub
            End If
        End If
        Next linha
     Next coluna
  Case Is = 5, 6
    For coluna = 1 To 2
        For linha = 0 To 6
          If Trim(grid_recolim(coluna, linha)) <> "" And Trim(grid_recolim(coluna, linha)) <> "0,00" Then
            Select Case coluna
            Case 1
                If linha = 0 Then
                   item_a_gravar = linha + 305
                Else
                   item_a_gravar = linha + 306
                End If
            Case 2
                If linha = 0 Then
                   item_a_gravar = linha + 316
                Else
                   item_a_gravar = linha + 317
                End If
            End Select
              valor_a_gravar = Str(grid_recolim(coluna, linha))
              If Not grava_item(item_a_gravar, valor_a_gravar) Then
                 Screen.MousePointer = 0
                 Exit Sub
              End If
          End If
           
        Next linha
     Next coluna
  Case Is = 2
    For coluna = 1 To 2
        For linha = 0 To 6
          Select Case coluna
          
          Case 1
                item_a_gravar = linha + 306
                If Trim(grid_recolim(coluna, linha)) <> "" And Trim(grid_recolim(coluna, linha)) <> "0,00" Then
                    valor_a_gravar = Str(grid_recolim(coluna, linha))
                    If Not grava_item(item_a_gravar, valor_a_gravar) Then
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
          Case 2
               item_a_gravar = linha + 317
                If Trim(grid_recolim(coluna, linha)) <> "" And Trim(grid_recolim(coluna, linha)) <> "0,00" Then
                    valor_a_gravar = Str(grid_recolim(coluna, linha))
                    If Not grava_item(item_a_gravar, valor_a_gravar) Then
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
          End Select
        Next linha
     Next coluna
     
  Case Is = 3, 4, 7
    For coluna = 1 To 2
        For linha = 0 To 7
            Select Case coluna
            Case 1
                If linha = 0 Or linha = 1 Then
                   item_a_gravar = linha + 304
                Else
                   item_a_gravar = linha + 305
                End If
            Case 2
                If linha = 0 Or linha = 1 Then
                   item_a_gravar = linha + 315
                Else
                   item_a_gravar = linha + 316
                End If
            End Select
            If Trim(grid_recolim(coluna, linha)) <> "" And Trim(grid_recolim(coluna, linha)) <> "0,00" Then
              valor_a_gravar = Str(grid_recolim(coluna, linha))
              If Not grava_item(item_a_gravar, valor_a_gravar) Then
                  Screen.MousePointer = 0
                 Exit Sub
              End If
          End If
        Next linha
     Next coluna
End Select


If g_condicao <> 2 Then
     If Trim(txt300.Text) <> "" Then
        valor_a_gravar = Str(txt300.Text)
        If Not grava_item(300, valor_a_gravar) Then
           Screen.MousePointer = 0
           Exit Sub
        End If
     End If
     
     If Trim(txt301.Text) <> "" Then
        valor_a_gravar = Str(txt301.Text)
        If Not grava_item(301, valor_a_gravar) Then
           Screen.MousePointer = 0
           Exit Sub
        End If
   End If
   
   If Trim(txt302.Text) <> "" Then
        valor_a_gravar = Str(txt302.Text)
        If Not grava_item(302, valor_a_gravar) Then
           Screen.MousePointer = 0
          Exit Sub
        End If
   End If
   
   If Trim(txt303.Text) <> "" Then
        valor_a_gravar = Str(txt303.Text)
        If Not grava_item(303, valor_a_gravar) Then
           Screen.MousePointer = 0
           Exit Sub
        End If
 End If
End If

  Exit Sub

erro_grava_reco:

     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next


End Sub


Sub grava_grid_Faturamento()
   Dim linha As Integer
   Dim coluna As Integer
   Dim item_a_gravar As Integer
   
   On Error GoTo erro_grava_fatura
   
    If Not deleta_itens_dem_entrada(217, 242) Then Exit Sub

   For coluna = 1 To 3
      For linha = 0 To 3
        If Trim(grid_fatura(coluna, linha)) <> "" And Trim(grid_fatura(coluna, linha)) <> "0,00" Then
           If linha < 3 Then
              Select Case Right(g_periodo_inicial, 2)
                Case Is = "03"
                    valor_a_gravar = linha + 1
                    item_a_gravar = linha + 217
                Case Is = "06"
                    valor_a_gravar = linha + 4
                    item_a_gravar = linha + 217
                Case Is = "09"
                    valor_a_gravar = linha + 7
                    item_a_gravar = linha + 217
                Case Is = "12"
                    valor_a_gravar = linha + 10
                    item_a_gravar = linha + 217
              End Select

              If Not grava_item(item_a_gravar, valor_a_gravar) Then
                 Screen.MousePointer = 0
                 Exit Sub
              End If
           End If
           If linha < 3 Then
               Select Case coluna
                 Case 1
                    item_a_gravar = linha + 223
                 Case 2
                    If linha < 3 Then
                       item_a_gravar = linha + 230
                    Else
                       Exit For
                    End If
                 Case 3
                    item_a_gravar = linha + 236
               End Select
           Else
               Select Case coluna
                 Case 1
                    item_a_gravar = 229
                 Case 2
                    Exit For
                 Case 3
                    item_a_gravar = 242
               End Select
           End If
              valor_a_gravar = Str(grid_fatura(coluna, linha))
              If Not grava_item(item_a_gravar, valor_a_gravar) Then
                 Screen.MousePointer = 0
                 Exit Sub
              End If
        
        End If
      Next linha
   Next coluna

  Exit Sub

erro_grava_fatura:

     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next


End Sub


Sub grava_grid_DetDesp()
Dim linha As Integer
Dim coluna As Integer
Dim item_a_gravar As Integer
On Error GoTo erro_grava_detdesp

If Not deleta_itens_dem_entrada(400, 424) Then Exit Sub

      For coluna = 1 To 1
          For linha = 0 To 14
             If Trim(grid_DetDesp(coluna, linha)) <> "" And Trim(grid_DetDesp(coluna, linha)) <> "0,00" Then
                item_a_gravar = linha + 401
                valor_a_gravar = Str(grid_DetDesp(coluna, linha))
                If Not grava_item(item_a_gravar, valor_a_gravar) Then
                   Screen.MousePointer = 0
                   Exit Sub
                End If
             End If
          Next linha
       Next coluna
       
        If Trim(txt400.Text) <> "" Then
          valor_a_gravar = Str(txt400.Text)
          If Not grava_item(400, valor_a_gravar) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
       End If
       
        If Trim(txt416.Text) <> "" Then
           valor_a_gravar = Str(txt416.Text)
           If Not grava_item(416, valor_a_gravar) Then
              Screen.MousePointer = 0
             Exit Sub
           End If
        End If
        
        If Trim(txt417.Text) <> "" Then
          valor_a_gravar = Str(txt417.Text)
          If Not grava_item(417, valor_a_gravar) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
        End If
        
        If Trim(txt418.Text) <> "" Then
          valor_a_gravar = Str(txt418.Text)
          If Not grava_item(418, valor_a_gravar) Then
             Screen.MousePointer = 0
            Exit Sub
          End If
        End If
        
        If Trim(txt419.Text) <> "" Then
          valor_a_gravar = Str(txt419.Text)
          If Not grava_item(419, valor_a_gravar) Then
             Screen.MousePointer = 0
            Exit Sub
          End If
        End If
        
        If Trim(txt420.Text) <> "" Then
          valor_a_gravar = Str(txt420.Text)
          If Not grava_item(420, valor_a_gravar) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
        End If
        
        If Trim(txt421.Text) <> "" Then
          valor_a_gravar = Str(txt421.Text)
          If Not grava_item(421, valor_a_gravar) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
        End If
        
        If Trim(txt422.Text) <> "" Then
          valor_a_gravar = Str(txt422.Text)
          If Not grava_item(422, valor_a_gravar) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
        End If
        
        If Trim(txt423.Text) <> "" Then
          valor_a_gravar = Str(txt423.Text)
          If Not grava_item(423, valor_a_gravar) Then
             Screen.MousePointer = 0
            Exit Sub
          End If
        End If

   If Trim(txt424.Text) <> "" Then
          valor_a_gravar = Str(txt424.Text)
          If Not grava_item(424, valor_a_gravar) Then
             Screen.MousePointer = 0
            Exit Sub
          End If
   End If
   
  Exit Sub

erro_grava_detdesp:

     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next

End Sub



Sub grava_grid_ApurICMS()
      Dim linha As Integer
      Dim coluna As Integer
      Dim item_a_gravar As Integer
  On Error GoTo erro_grava_apuricms
  
   If Not deleta_itens_dem_entrada(200, 221) Then Exit Sub

     For coluna = 1 To 1
        For linha = 0 To 17
        If Trim(grid_ApurICMS(coluna, linha)) <> "" And Trim(grid_ApurICMS(coluna, linha)) <> "0,00" Then
           item_a_gravar = linha + 200 + 1
           valor_a_gravar = Str(grid_ApurICMS(coluna, linha))
           If Not grava_item(item_a_gravar, valor_a_gravar) Then
              Screen.MousePointer = 0
             Exit Sub
           End If
        End If
        Next linha
     Next coluna
     
     If g_condicao = 1 Or g_condicao = 7 Then
       If Trim(txt_fatura_normal.Text) <> "" Then
          valor_a_gravar = Str(txt_fatura_normal.Text)
          If Not grava_item(200, valor_a_gravar) Then
              Screen.MousePointer = 0
              Exit Sub
          End If
       End If
    End If
      
    If g_condicao = 3 Or g_condicao = 4 Or g_condicao = 5 Or g_condicao = 6 Then
       If Trim(msk_fat_mensal_frm4.Text) <> "" Then
         valor_a_gravar = Str(msk_fat_mensal_frm4.Text)
         If Not grava_item(219, valor_a_gravar) Then
             Screen.MousePointer = 0
             Exit Sub
         End If
       End If
       
       If Trim(msk_aliquota_frm4.Text) <> "" Then
         valor_a_gravar = Str(msk_aliquota_frm4.Text)
         If Not grava_item(220, valor_a_gravar) Then
             Screen.MousePointer = 0
             Exit Sub
         End If
       End If
       
       If Trim(msk_imposto_frm4.Text) <> "" Then
         valor_a_gravar = Str(msk_imposto_frm4.Text)
         If Not grava_item(221, valor_a_gravar) Then
             Screen.MousePointer = 0
             Exit Sub
         End If
       End If
    End If
    
  Exit Sub

erro_grava_apuricms:

     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next

End Sub

Sub grava_detalhe_dem_saidas()
Dim linha As Integer
Dim coluna As Integer
Dim item_a_gravar As Integer
On Error GoTo erro_grava_saidas

If Not deleta_itens_dem_entrada(100, 159) Then Exit Sub
 
For coluna = 1 To 5
    For linha = 0 To 11
       If Trim(grid_saidas(coluna, linha)) <> "" And Trim(grid_saidas(coluna, linha)) <> "0,00" Then
           Select Case coluna
              Case Is = 1
                item_a_gravar = linha + 100
              Case Is = 2
                item_a_gravar = linha + 112
              Case Is = 3
                item_a_gravar = linha + 124
              Case Is = 4
                item_a_gravar = linha + 136
              Case Is = 5
                item_a_gravar = linha + 148
          End Select
          valor_a_gravar = Str(grid_saidas(coluna, linha))
          If Not grava_item(item_a_gravar, valor_a_gravar) Then
            Screen.MousePointer = 0
            Exit Sub
          End If
      
      End If
    Next linha
 Next coluna
  
 Exit Sub

erro_grava_saidas:

     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next

End Sub



Sub grava_detalhe_dem_entradas()
On Error GoTo erro_grava_entradas
Dim linha As Integer
Dim coluna As Integer
Dim item_a_gravar As Integer

Select Case g_condicao
   Case Is = 2, 5, 6
   If Not deleta_itens_dem_entrada(1, 12) Then Exit Sub
   
    For coluna = 1 To 1
        For linha = 0 To 11
         If Trim(GridArray(coluna, linha)) <> "" And Trim(GridArray(coluna, linha)) <> "0,00" Then
'          If CDbl(GridArray(coluna, linha)) <> 0 Then
              Select Case coluna
                  Case Is = 1
                      item_a_gravar = linha + 1
                  Case Is = 2
                      item_a_gravar = linha + 13
                  Case Is = 3
                      item_a_gravar = linha + 25
                  Case Is = 4
                      item_a_gravar = linha + 37
                  Case Is = 5
                      item_a_gravar = linha + 49
              End Select
              valor_a_gravar = Str(GridArray(coluna, linha))
             If Not grava_item(item_a_gravar, valor_a_gravar) Then
                Screen.MousePointer = 0
                Exit Sub
             End If
          End If
        Next linha
     Next coluna
Case Else
 If Not deleta_itens_dem_entrada(1, 60) Then Exit Sub
    For coluna = 1 To 5
       For linha = 0 To 11
'         For linha = 0 To 10
           If Trim(GridArray(coluna, linha)) <> "" And Trim(GridArray(coluna, linha)) <> "0,00" Then 'Or CDbl(Trim(Str(GridArray(coluna, linha)))) <> 0 Then
              Select Case coluna
                  Case Is = 1
                      item_a_gravar = linha + 1
                      
                  Case Is = 2
                      item_a_gravar = linha + 13
                      
                  Case Is = 3
                      item_a_gravar = linha + 25
                      
                  Case Is = 4
                      item_a_gravar = linha + 37
                                       
                  Case Is = 5
                      item_a_gravar = linha + 49
              End Select
              
              valor_a_gravar = Str(GridArray(coluna, linha))
             If Not grava_item(item_a_gravar, valor_a_gravar) Then
                Screen.MousePointer = 0
                Exit Sub
             End If
          
         End If
        Next linha
     Next coluna
     
End Select

Exit Sub

erro_grava_entradas:

     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next


End Sub

Sub AtualizaMsk()
On Error Resume Next 'GoTo erro_atualizamsk

   Dim Pegue As String
   Dim RecRazTipo As Recordset
  
   ' limpa_tela_Det
   APENAS_ENTREI = True
   msk_inscricao = g_num_inscricao
   txt_periodo_inicial = Mid$(g_periodo_inicial, 5, 2) & Mid$(g_periodo_inicial, 1, 4)
  
   If g_tipo_decla = 0 Then
      cmb_tipo.ListIndex = 0
   Else
      cmb_tipo.ListIndex = 1
   End If
   Pegue = "select * from tab_cad_contribuintes where tcc_nu_inscricao = '" & Trim(msk_inscricao) & "'"
   Set RecRazTipo = dbgim.OpenRecordset(Pegue, dbOpenDynaset)
   txt_razao = RecRazTipo.Fields("tcc_razao_social")
   g_num_cgc_cpf = RecRazTipo.Fields("tcc_cad_geral_contrib")
   g_ano_inicio_operacao = Val(Mid(RecRazTipo.Fields("tcc_da_inicio_op") & "", 5, 4))
   PER_INICIO = Val(Mid(RecRazTipo.Fields("tcc_da_inicio_op") & "", 5, 4) & Mid(RecRazTipo.Fields("tcc_da_inicio_op") & "", 3, 2)) & ""
    'g_condicao = RecRazTipo.Fields("tcc_tipo_contribuinte")
   RecRazTipo.Close
    
    query = "select * from tab_declaracao where tde_num_inscricao =   '" & g_num_inscricao & "'"
    query = query & " and [tde_periodo_inicial] = '" & g_periodo_inicial & "'"
    query = query & " and [tde_substitutiva] = " & g_tipo_decla
     
    Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
    If restauracontr.EOF Then Exit Sub
    g_aliquota_Maior17 = restauracontr.Fields("tde_aliquota_m17") & ""
    Select Case restauracontr.Fields("tde_condicao")
         Case 1
            txtCondContr = "Normal"
             lbl_tipo_decl.Caption = "Normal"
         Case 3, 4, 5, 6, 7
            txtCondContr = "Pequeno Porte"
            posiciona_label_tipo_decl
           
         Case 2
            txtCondContr = "Microempresa"
            lbl_tipo_decl.Caption = "Microempresa"
    End Select
    If g_condicao = 0 Then
      g_condicao = restauracontr.Fields("tde_condicao")
    End If
    
    EstadoDecla = restauracontr.Fields("tde_Estado")
    Situacao.Visible = True
      If EstadoDecla = "I" Then
         Situacao.Caption = "Incompleta"
         Option1(0).Value = True
      ElseIf EstadoDecla = "F" Or EstadoDecla = "G" Then
         Situacao.Caption = "Finalizada"
         Option1(1).Value = True
      End If
    restauracontr.Close
    've se a declaracao esta fechada e verifica o controle ,isto é, se nao foi alterado o banco
    verifica_controle_declaracao
    
erro_atualizamsk:
If Err <> 0 Then
     MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
    Resume Next
End If
End Sub


Sub AtualizaTela()

'On Error Resume Next

   ReDim GridArray(0, 0) As Variant  'demonstrativo das oper. entradas
   ReDim grid_saidas(0, 0) As Variant 'demonstrativo das opr saidas
   ReDim grid_ApurICMS(0, 0) As Variant 'Apuração do ICMS
   ReDim grid_recolim(0, 0) As Variant 'recolhimento do periodo
   ReDim grid_fatura(0, 0) As Variant 'faturamento frame 5
   ReDim grid_DetDesp(0, 0) As Variant 'faturamento frame 5
   
   APENAS_ENTREI = True
   cmb_tipo.Clear
   Situacao.Caption = ""
   cmb_tipo.AddItem "Não"
   cmb_tipo.AddItem "Sim"
   cmb_tipo.ListIndex = 0
   erro_periodo = False
   lbl_tipo_decl.Caption = ""
   limpa_capos_txt
   If gformulario = 0 Then
      g_num_inscricao = ""
      g_periodo_inicial = ""
      g_tipo_decla = 0
      g_condicao = 0
      g_ano_inicio_operacao = 0
      msk_inscricao = g_num_inscricao
      txt_periodo_inicial = g_periodo_inicial
      txt_razao = ""
      txtCondContr = ""
      txt_fatura_normal = ""
   ElseIf gformulario = 1 Then
        AtualizaMsk
   End If
   
   limpa_grids_entrada
   
   cargar_data
   cargar_datos_saidas
   cargar_grid_recolhimento
   cargar_grid_ApurICMS
   cargar_grid_faturamento
   cargar_grid_DetDesp
  
   AtualizaGrades
   
   TestaCampoPreench

   gformulario = 0
   
   SSTab1.Tab = 0
 '  monta_frames  ' inibi em 24/2/98
End Sub


Sub cargar_data()
On Error Resume Next
 Dim estado As String
 Dim i As Long
   
 query = "Select *  From tab_detalhe_declaracao "
 query = query & "Where tdd_num_inscricao  = '" & g_num_inscricao & "'"
 query = query & " and tdd_periodo_inicial = '" & g_periodo_inicial & "'"
 query = query & " and tdd_substitutiva = " & g_tipo_decla
 
 Select Case g_condicao
      Case Is = 2, 5, 6
            query = query & " and tdd_item between 01 and 12 "
            ReDim GridArray(0 To 1, 0 To 12)
            ReDim Preserve GridArray(0 To 1, 0 To 12)
            MaxCol = 2
            MaxRow = 13
      Case Else
            query = query & " and tdd_item between 01 and 60 "
            ReDim GridArray(0 To 5, 0 To 12)
            ReDim Preserve GridArray(0 To 5, 0 To 12)
            MaxCol = 6
            MaxRow = 13
 End Select
    
    
  
    Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
    If seleccion.BOF Or seleccion.EOF Then
        seleccion.Close
        monta_cabecalho_grid
        Select Case g_condicao
           Case Is = 2, 5, 6
              grid_Entra_MicroEPP.Refresh
           Case Else
            GRID_DEM_ENTRADAS.Refresh
        End Select
        Exit Sub
    End If
    
    seleccion.MoveFirst
    monta_cabecalho_grid
    i = 0
    While Not seleccion.EOF
      Select Case seleccion("tdd_item")
       Case Is = 1   '  coluna 1
         GridArray(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 2
         GridArray(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 3
         GridArray(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 4
         GridArray(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 5
         GridArray(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 6
         GridArray(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 7
         GridArray(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 8
         GridArray(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 9
         GridArray(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 10
         GridArray(1, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 11
         GridArray(1, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 12
          GridArray(1, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 13
          GridArray(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 14
         GridArray(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 15
         GridArray(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 16
         GridArray(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 17
         GridArray(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 18
         GridArray(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 19
         GridArray(2, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 20
         GridArray(2, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 21
         GridArray(2, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 22
         GridArray(2, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 23
         GridArray(2, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 24
         GridArray(2, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         'txt024.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     'coluna 3
       Case Is = 25
         GridArray(3, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 26
         GridArray(3, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 27
         GridArray(3, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 28
         GridArray(3, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 29
         GridArray(3, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 30
         GridArray(3, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 31
         GridArray(3, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 32
         GridArray(3, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 33
         GridArray(3, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 34
         GridArray(3, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 35
         GridArray(3, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 36
        GridArray(3, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
        txt205.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     
     'coluna 4
       Case Is = 37
         GridArray(4, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 38
         GridArray(4, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 39
         GridArray(4, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 40
         GridArray(4, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 41
         GridArray(4, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 42
         GridArray(4, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 43
         GridArray(4, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 44
         GridArray(4, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 45
         GridArray(4, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 46
         GridArray(4, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 47
         GridArray(4, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 48
        GridArray(4, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
        'txt048.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      'coluna 5
       Case Is = 49
         GridArray(5, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 50
         GridArray(5, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 51
         GridArray(5, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 52
         GridArray(5, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 53
         GridArray(5, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 54
         GridArray(5, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 55
         GridArray(5, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 56
         GridArray(5, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 57
         GridArray(5, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 58
         GridArray(5, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 59
         GridArray(5, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 60
         GridArray(5, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
        'txt060.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      End Select
        seleccion.MoveNext
        i = i + 1
    Wend
    Select Case g_condicao
       Case 2, 5, 6
            grid_Entra_MicroEPP.Refresh
            MaxCol = 2
            MaxRow = 13
       Case Else
            GRID_DEM_ENTRADAS.Refresh
            MaxRow = 13
            MaxCol = 6
     End Select
     seleccion.Close

End Sub

Sub cargar_datos_saidas()
On Error Resume Next
 Dim estado As String
   Dim i As Long
   
 query = "Select *  From tab_detalhe_declaracao "
 query = query & "Where tdd_num_inscricao  = '" & g_num_inscricao & "'"
 query = query & " and tdd_periodo_inicial = '" & g_periodo_inicial & "'"
 query = query & " and tdd_substitutiva = " & g_tipo_decla
 
 query = query & " and tdd_item between 100 and 159 "

   ReDim grid_saidas(0 To 5, 0 To 12)
   ReDim Preserve grid_saidas(0 To 5, 0 To 12)
  

    MaxCol_s = 6
    MaxRow_s = 13
    Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
    If seleccion.BOF Or seleccion.EOF Then
        seleccion.Close
        monta_grid_saidas
        grid_dem_saidas.Refresh
        Exit Sub
    End If
    seleccion.MoveFirst
    
    monta_grid_saidas
    i = 0
    While Not seleccion.EOF
      Select Case seleccion("tdd_item")
       Case Is = 100   '  coluna 1
         grid_saidas(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 101
         grid_saidas(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 102
         grid_saidas(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 103
         grid_saidas(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 104
         grid_saidas(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 105
         grid_saidas(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 106
         grid_saidas(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 107
         grid_saidas(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 108
         grid_saidas(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 109
         grid_saidas(1, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 110
         grid_saidas(1, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 111
         grid_saidas(1, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         'txt111.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         'coluna 2
       Case Is = 112
         grid_saidas(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 113
         grid_saidas(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 114
         grid_saidas(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 115
         grid_saidas(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 116
         grid_saidas(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 117
         grid_saidas(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 118
         grid_saidas(2, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 119
         grid_saidas(2, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 120
         grid_saidas(2, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 121
         grid_saidas(2, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 122
         grid_saidas(2, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 123
        grid_saidas(2, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
        'txt205.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")

     'coluna 3
       Case Is = 124
         grid_saidas(3, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 125
         grid_saidas(3, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 126
         grid_saidas(3, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 127
         grid_saidas(3, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 128
         grid_saidas(3, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 129
         grid_saidas(3, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 130
         grid_saidas(3, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 131
         grid_saidas(3, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 132
         grid_saidas(3, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 133
         grid_saidas(3, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 134
         grid_saidas(3, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 135
         grid_saidas(3, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         txt201.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")

     'coluna 4
       Case Is = 136
         grid_saidas(4, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 137
         grid_saidas(4, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 138
         grid_saidas(4, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 139
         grid_saidas(4, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 140
         grid_saidas(4, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 141
         grid_saidas(4, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 142
         grid_saidas(4, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 143
         grid_saidas(4, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 144
         grid_saidas(4, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 145
         grid_saidas(4, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 146
         grid_saidas(4, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 147
         grid_saidas(4, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         ' txt147.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")

     'coluna 5
       Case Is = 148
         grid_saidas(5, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 149
         grid_saidas(5, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 150
         grid_saidas(5, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 151
         grid_saidas(5, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 152
         grid_saidas(5, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 153
         grid_saidas(5, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 154
         grid_saidas(5, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 155
         grid_saidas(5, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 156
         grid_saidas(5, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 157
         grid_saidas(5, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 158
         grid_saidas(5, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 159
        grid_saidas(5, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
        ' txt159.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")

      End Select
        seleccion.MoveNext
        i = i + 1
    Wend
    
    grid_dem_saidas.Refresh
    seleccion.Close
End Sub


Sub monta_frames()
On Error Resume Next
desativa_label_pisca_pisca
Select Case SSTab1.Tab
  Case Is = 0
    fra_Dem_Saidas.Enabled = False
    fra_Dem_Saidas.Visible = False
    fra_Apur_ICMS.Visible = False
    fra_Recolhimento.Enabled = False
    fra_Recolhimento.Visible = False
    fra_Dados_Compl.Enabled = False
    fra_Dados_Compl.Visible = False
    frm_grid_fatu.Enabled = False
    frm_grid_fatu.Visible = False
    
    fra_EntaNormal.Left = 30
    fra_EntaNormal.Top = 650
    fra_EntaNormal.Enabled = True
    fra_EntaNormal.Visible = True
    fra_EntraMicroEPP.Enabled = False
    fra_EntraMicroEPP.Visible = False
    Select Case g_condicao
     Case 3, 4, 5, 6, 7
         If g_aliquota_Maior17 = 2 Then
            fra_EntraMicroEPP.Top = 660
            fra_EntraMicroEPP.Left = 75
            fra_EntaNormal.Enabled = False
            fra_EntaNormal.Visible = False
            fra_EntraMicroEPP.Enabled = True
            fra_EntraMicroEPP.Visible = True
        ElseIf g_aliquota_Maior17 = 1 Then
            fra_EntaNormal.Enabled = True
            fra_EntaNormal.Visible = True
            fra_EntraMicroEPP.Enabled = False
            fra_EntraMicroEPP.Visible = False
        End If
     Case 2
            fra_EntraMicroEPP.Top = 660
            fra_EntraMicroEPP.Left = 75
            fra_EntaNormal.Enabled = False
            fra_EntaNormal.Visible = False
            fra_EntraMicroEPP.Enabled = True
            fra_EntraMicroEPP.Visible = True
      End Select
Case Is = 1
    fra_EntaNormal.Enabled = False
    fra_EntaNormal.Visible = False
    fra_EntraMicroEPP.Enabled = False
    fra_EntraMicroEPP.Visible = False
    fra_Apur_ICMS.Visible = False
    fra_Recolhimento.Enabled = False
    fra_Recolhimento.Visible = False
    fra_Dados_Compl.Enabled = False
    fra_Dados_Compl.Visible = False
    frm_grid_fatu.Enabled = False
    frm_grid_fatu.Visible = False
    fra_Dem_Saidas.Left = 30
    fra_Dem_Saidas.Top = 650
    fra_Dem_Saidas.Enabled = True
    fra_Dem_Saidas.Visible = True
Case Is = 2
    fra_EntaNormal.Enabled = False
    fra_EntaNormal.Visible = False
    fra_EntraMicroEPP.Enabled = False
    fra_EntraMicroEPP.Visible = False
    fra_Dem_Saidas.Enabled = False
    fra_Dem_Saidas.Visible = False
    fra_Recolhimento.Enabled = False
    fra_Recolhimento.Visible = False
    Frame7.Visible = False
    fra_Dados_Compl.Enabled = False
    fra_Dados_Compl.Visible = False
    Select Case g_condicao
        Case Is = 1, 7              'Normal  e epp c/faturamento maior que 240.000
        '**** 'Faturamento - Tela #4 e 5  'ICMS - Tela #3
           frm_grid_fatu.Enabled = False
           frm_grid_fatu.Visible = False
           fra_Apur_ICMS.Left = 75
           fra_Apur_ICMS.Top = 585
           'Frame_tela3.Left = 20
           'Frame_tela3.Top = 20
           fra_Apur_ICMS.Visible = True
           fra_Apur_ICMS.Enabled = True
           
        Case Is = 3, 4, 5, 6    'EPP
            frm_grid_fatu.Enabled = False
            frm_grid_fatu.Visible = False
            If g_aliquota_Maior17 = 1 Then    'aliquota maior que 17%
              'Faturamento - Tela #4 e 5  'ICMS - Tela #3
               fra_Apur_ICMS.Left = 75 '210
               fra_Apur_ICMS.Enabled = True
               fra_Apur_ICMS.Visible = True
              ' Label_epp.Caption = "Faturamento do Mês (Receita Bruta - alíquota > 17%):"
            Else
              ' Label_epp.Caption = "Faturamento do Mês (Receita Bruta) :"
              fra_Apur_ICMS.Left = 75 '210
              fra_Apur_ICMS.Visible = True
              fra_Apur_ICMS.Enabled = True
             End If
        Case Is = 2                 'Microempresa'ICMS - Tela #3  'Faturamento - Tela #4 e 5
            fra_Apur_ICMS.Visible = False
            'fra_Apur_ICMS.Enabled = False
            frm_grid_fatu.Left = 1070
            frm_grid_fatu.Top = 1060
            frm_grid_fatu.Visible = True
            frm_grid_fatu.Enabled = True
    End Select
Case Is = 3
    fra_EntaNormal.Enabled = False
    fra_EntaNormal.Visible = False
    fra_EntraMicroEPP.Enabled = False
    fra_EntraMicroEPP.Visible = False
    fra_Dem_Saidas.Enabled = False
    fra_Dem_Saidas.Visible = False
    'fra_Apur_ICMS.Enabled = False
    fra_Apur_ICMS.Visible = False
    fra_Dados_Compl.Enabled = False
    fra_Dados_Compl.Visible = False
    frm_grid_fatu.Enabled = False
    frm_grid_fatu.Visible = False
    fra_Recolhimento.Left = 300
    fra_Recolhimento.Top = 800
'    txt314.Refresh
    fra_Recolhimento.Enabled = True
    fra_Recolhimento.Visible = True
    Frame6.Visible = True
    Frame7.Visible = True
    Frame7.Enabled = True
Case Is = 4
    fra_EntaNormal.Enabled = False
    fra_EntaNormal.Visible = False
    fra_EntraMicroEPP.Enabled = False
    fra_EntraMicroEPP.Visible = False
    fra_Dem_Saidas.Enabled = False
    fra_Dem_Saidas.Visible = False
    'fra_Apur_ICMS.Enabled = False
    fra_Apur_ICMS.Visible = False
    fra_Recolhimento.Enabled = False
    fra_Recolhimento.Visible = False
    frm_grid_fatu.Enabled = False
    frm_grid_fatu.Visible = False
    fra_Dados_Compl.Left = 75
    fra_Dados_Compl.Top = 670
    fra_Dados_Compl.Enabled = True
    fra_Dados_Compl.Visible = True

End Select

End Sub


Sub transporta_valores_Apur_ICMS()
  
  'If Val(txt205.Text) = 0 Then
  GRID_DEM_ENTRADAS.Row = 11
  txt205.Text = GRID_DEM_ENTRADAS.Columns(3).Value
  'End If
  
  'If Val(txt201.Text) = 0 Then
  grid_dem_saidas.Row = 11
  txt201.Text = grid_dem_saidas.Columns(3).Value
  'End If
  
  grid_ApurICMS(1, 0) = txt201.Text
  grid_ApurICMS(1, 4) = txt205.Text


End Sub

Function trata_EPP() As Integer
On Error GoTo erro_trata_epp
DoEvents
 vim_trata_epp = True
FRM_FATU.Show 1

'Me.Refresh
'DoEvents
trata_EPP = True
If G_opt_fat1 = True Then   '  empresa com menos que 160.000 aliq. 1.5%
    If g_condicao = 3 Then
       lbl_tipo_decl.Caption = "Com Alíquota de 1.5% S/Faturamento maior que 17%"
    ElseIf g_condicao = 5 Then
      lbl_tipo_decl.Caption = "Com Alíquota de 1.5% S/Faturamento menor que 17%"
    Else
        trata_EPP = False
    End If
ElseIf G_opt_fat2 = True Then
    If g_condicao = 4 Then
       lbl_tipo_decl.Caption = "Com Alíquota de 2.0% S/Faturamento maior que 17%"
    ElseIf g_condicao = 6 Then
       lbl_tipo_decl.Caption = "Com Alíquota de 2.0% S/Faturamento menor que 17%"
    Else
       trata_EPP = False
    End If
ElseIf G_opt_fat3 = True Then
    g_condicao = 7
   lbl_tipo_decl.Caption = "Com Faturamento maior que 240.000,00"
Else
  trata_EPP = False
End If

erro_trata_epp:
 If Err <> 0 Then
    MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next
End If
End Function
Sub verifica_controle_declaracao()
Dim snap_declaracao As Recordset
Dim controle As Double
Dim controle_decl As Double
Dim VARIAVEL As String
 On Error GoTo erro_controle
    
    query = "Select tde_controle ,tde_Estado From tab_declaracao "
    query = query & "Where tde_num_inscricao  = '" & g_num_inscricao & "'"
    query = query & " and tde_periodo_inicial = '" & g_periodo_inicial & "'"
    query = query & " and tde_substitutiva = " & g_tipo_decla
    query = query & " and tde_condicao = " & g_condicao
    
    Set snap_declaracao = dbgim.OpenRecordset(query, dbOpenSnapshot)
    
   If Not snap_declaracao.EOF Then
      If snap_declaracao!tde_Estado <> "I" Then
        controle = calcula_controle(g_num_inscricao, g_periodo_inicial, g_tipo_decla) * 100
        VARIAVEL = (snap_declaracao!tde_controle) & ""
        If Trim(VARIAVEL) = "" Then
           controle_decl = (snap_declaracao!tde_controle) & "0"
        Else
           controle_decl = (snap_declaracao!tde_controle)
        End If
        If Trim(Str(controle)) <> Trim(Str(controle_decl)) Then
            query = "update tab_declaracao set "
            query = query & "tde_estado = 'I'"
            query = query & " where tde_num_inscricao = '" & g_num_inscricao & "'"
            query = query & " and tde_periodo_inicial = '" & g_periodo_inicial & "'"
            query = query & " and tde_substitutiva = " & Val(g_tipo_decla)
            query = query & " and tde_condicao = " & g_condicao
           If Not EjecutarQuery(query) Then
           End If
            Option1(0).Value = True
            Option1(1).Value = False
        End If
      End If
   End If
 
 snap_declaracao.Close
 Exit Sub
erro_controle:
 MsgBox Err
 
End Sub

Function verifica_deleta_item(item As Integer) As Integer

verifica_deleta_item = True
  Query1 = "select * from tab_detalhe_declaracao where tdd_num_inscricao = '" & Trim(Str(msk_inscricao)) & "'"
  Query1 = Query1 & " and tdd_periodo_inicial = '" & g_periodo_inicial & "'"
  Query1 = Query1 & " and tdd_substitutiva = " & Val(g_tipo_decla)

  Query1 = Query1 & " and tdd_item = " & item

Set restauracontr = dbgim.OpenRecordset(Query1, dbOpenDynaset)
  If Not restauracontr.EOF And Not restauracontr.BOF Then
        query = "delete from  tab_detalhe_declaracao  "
        query = query & " where tdd_num_inscricao = '" & Trim(Str(msk_inscricao)) & "'"
        query = query & " and tdd_periodo_inicial = '" & g_periodo_inicial & "'"
        query = query & " and tdd_substitutiva = " & Val(g_tipo_decla)
        query = query & " and tdd_item = " & item
    If Not EjecutarQuery(query) Then
        Screen.MousePointer = 0
        verifica_deleta_item = False
        Exit Function
    End If
End If
restauracontr.Close
End Function

Private Sub Calcular_Click()
Dim MyAppID

On Error GoTo errocalc:
    
If rota = True Then
    MyAppID = Shell("C:\WINDOWS\CALC.EXE", 1) ' Run Calculator.
    rota = False
Else
    
    AppActivate "Calculadora"

End If

Exit Sub


errocalc:
If Err = 5 Then
    MyAppID = Shell("C:\WINDOWS\CALC.EXE", 1) ' Run Calculator.
End If

End Sub

Private Sub cmb_tipo_Change()

If APENAS_ENTREI = True Then
     APENAS_ENTREI = False
     Exit Sub
 End If
 

   If Trim(Mid(cmb_tipo, 1, 1)) <> "" Then
    reta = crit_tipo_decl(Mid(cmb_tipo, 1, 1))
    If Not reta Then
       MsgBox "Tipo de Declaração inválida.", 48, NOME_SISTEMA_DEFAULT
       cmb_tipo.SetFocus
       Exit Sub
    
    Else
      g_periodo_inicial = Mid(txt_periodo_inicial, 3, 4) & Mid(txt_periodo_inicial, 1, 2)
      g_aliquota_Maior17 = 0
      g_num_inscricao = Trim$(msk_inscricao)
      TestaCampoPreench
      DoEvents
     
     If campopreench = True Then
          SSTab1.Enabled = True
          SSTab1.SetFocus
      Else
         cmd_sair.SetFocus
         Exit Sub
      End If
    End If
   End If

End Sub

Private Sub cmb_tipo_Click()

If APENAS_ENTREI = True Then
     APENAS_ENTREI = False
     Exit Sub
 End If
 

   If Trim(Mid(cmb_tipo, 1, 1)) <> "" Then
    reta = crit_tipo_decl(Mid(cmb_tipo, 1, 1))
    If Not reta Then
       MsgBox "Tipo de Declaração inválida.", 48, NOME_SISTEMA_DEFAULT
       cmb_tipo.SetFocus
       Exit Sub
    
    Else
      g_periodo_inicial = Mid(txt_periodo_inicial, 3, 4) & Mid(txt_periodo_inicial, 1, 2)
      g_aliquota_Maior17 = 0   'nenhum
      g_num_inscricao = Trim$(msk_inscricao)
    
    End If
   End If



End Sub


Private Sub cmb_tipo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
   If Trim(Mid(cmb_tipo, 1, 1)) <> "" Then
    reta = crit_tipo_decl(Mid(cmb_tipo, 1, 1))
    If Not reta Then
       MsgBox "Tipo de Declaração inválida.", 48, NOME_SISTEMA_DEFAULT
       cmb_tipo.SetFocus
       Exit Sub
    Else
      g_periodo_inicial = Mid(txt_periodo_inicial, 3, 4) & Mid(txt_periodo_inicial, 1, 2)
      g_aliquota_Maior17 = 0
      g_num_inscricao = Trim$(msk_inscricao)
      TestaCampoPreench
      If campopreench = True Then
          SSTab1.Enabled = True
          SSTab1.SetFocus
      Else
         cmd_sair.SetFocus
         Exit Sub
      End If
    End If
   End If
End If

End Sub


Private Sub cmb_tipo_LostFocus()
Dim ret As Integer
Dim cTipoDecla As String

 If TypeOf Me.ActiveControl Is ComboBox Then
       If Me.ActiveControl.Tag = "tipo" Then
          GoTo continua
       End If
 End If


 If TypeOf Me.ActiveControl Is SSCommand Then
       If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "diretorio" Then
          Exit Sub
       End If
 End If
 
continua:
If erro_periodo = True Then
   Exit Sub
End If

 If TypeOf Me.ActiveControl Is MaskEdBox Then
     If Me.ActiveControl.Tag = "inscricao" Or Me.ActiveControl.Tag = "periodo" Or Me.ActiveControl.Tag = "periodo" Then
          Exit Sub
       End If
  End If
  
reta = crit_tipo_decl(Mid(cmb_tipo, 1, 1))
 If Not reta Then
     MsgBox "Tipo de Declaração inválida.", 48, NOME_SISTEMA_DEFAULT
     cmb_tipo.SetFocus
     Exit Sub
 End If

      g_periodo_inicial = Mid(txt_periodo_inicial, 3, 4) & Mid(txt_periodo_inicial, 1, 2)
      cTipoDecla = Mid(cmb_tipo.Text, 1, 1)
      g_aliquota_Maior17 = 0   'nenhum
      g_num_inscricao = Trim$(msk_inscricao)
     TestaCampoPreench
   If campopreench = True Then
      SSTab1.Enabled = True
      SSTab1.SetFocus
   Else
      cmd_sair.SetFocus
      Exit Sub
   End If

   If cTipoDecla = "N" Then   'Original
      query = "select * from tab_declaracao where tde_num_inscricao =   '" & g_num_inscricao & "'"
      query = query & " and [tde_periodo_inicial] = '" & g_periodo_inicial & "'"
      query = query & " and [tde_substitutiva] = 0"
      Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
      If restauracontr.EOF And restauracontr.BOF Then
   
         If g_condicao = 3 Or g_condicao = 4 Or g_condicao = 5 Or g_condicao = 6 Or g_condicao = 7 Then
            If Not trata_EPP Then
                cmd_iniciar_Click
                Exit Sub
            End If
         End If
         Select Case g_condicao
           Case 1
             txtCondContr = "Normal"
             lbl_tipo_decl.Caption = "Normal"
           Case 3, 4, 5, 6, 7
             txtCondContr = "Pequeno Porte"
           Case 2
             lbl_tipo_decl.Caption = "Microempresa"
             txtCondContr = "Microempresa"
         End Select
          g_tipo_decla = 0
          EstadoDecla = "I"
          Option1(0).Value = True
      Else
         g_aliquota_Maior17 = restauracontr.Fields("tde_aliquota_m17")
         EstadoDecla = restauracontr.Fields("tde_Estado")
         If EstadoDecla = "I" Then
            Option1(0).Value = True
         ElseIf EstadoDecla = "F" Or EstadoDecla = "G" Then
            Option1(1).Value = True
         End If
         Select Case restauracontr.Fields("tde_condicao")
               Case 1
                  txtCondContr = "Normal"
                   lbl_tipo_decl.Caption = "Normal"
               Case 3, 4, 5, 6, 7
                  txtCondContr = "Pequeno Porte"
               Case 2
                   lbl_tipo_decl.Caption = "Microempresa"
                  txtCondContr = "Microempresa"
'               Case 4
'                  txtCondContr = "EPP c/Alíquota >17%"
         End Select
        ' If g_condicao = 0 Then
           g_condicao = restauracontr.Fields("tde_condicao")
       '  End If
          query = "select * from tab_detalhe_declaracao where tdd_num_inscricao =   '" & g_num_inscricao & "'"
          query = query & " and [tdd_periodo_inicial] = '" & g_periodo_inicial & "'"
          query = query & " and [tdd_substitutiva] = " & g_tipo_decla
          Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
      End If
   Else                    'Substitutiva
      query = "select  Count(*) as qtdreg from tab_declaracao where tde_num_inscricao =   '" & g_num_inscricao & "'"
      query = query & " and [tde_periodo_inicial] = '" & g_periodo_inicial & "'"
      query = query & " and [tde_substitutiva] <> 0"
      Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
      If restauracontr!qtdReg = 0 Then
          If g_condicao = 3 Or g_condicao = 4 Or g_condicao = 5 Or g_condicao = 6 Or g_condicao = 7 Then
              If Not trata_EPP Then
                cmd_iniciar_Click
                Exit Sub
              End If

          End If
           g_tipo_decla = 1
           EstadoDecla = "I"
           Option1(0).Value = True
       Else
          g_condicao = restauracontr.Fields("tde_condicao")
          ret = MsgBox("Deseja criar uma nova substitutiva?", 36, NOME_SISTEMA_DEFAULT)
          If ret = 6 Then
             If g_condicao = 3 Or g_condicao = 4 Or g_condicao = 5 Or g_condicao = 6 Or g_condicao = 7 Then
               If Not trata_EPP Then
                  cmd_iniciar_Click
                  Exit Sub
                End If

             End If
             g_tipo_decla = restauracontr!qtdReg + 1
             EstadoDecla = "I"
             Option1(0).Value = True
          Else
             If restauracontr!qtdReg = 1 Then
                g_tipo_decla = 1
                query = "select * from tab_declaracao where tde_num_inscricao =   '" & g_num_inscricao & "'"
                query = query & " and [tde_periodo_inicial] = '" & g_periodo_inicial & "'"
                query = query & " and [tde_substitutiva] = <> 0"
                Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
                g_aliquota_Maior17 = restauracontr.Fields("tde_aliquota_m17")
                Select Case restauracontr.Fields("tde_condicao")
                      Case 1
                         txtCondContr = "Normal"
                        lbl_tipo_decl.Caption = "Normal"
                      Case 3, 4, 5, 6, 7
                         txtCondContr = "Pequeno Porte"
                      Case 2
                        lbl_tipo_decl.Caption = "Microempresa"
                         txtCondContr = "Microempresa"
                End Select
'                If g_condicao = 0 Then
                   g_condicao = restauracontr.Fields("tde_condicao")
'                End If
                EstadoDecla = restauracontr.Fields("tde_Estado")
                If EstadoDecla = "I" Then
                  Option1(0).Value = True
                ElseIf EstadoDecla = "F" Or EstadoDecla = "G" Then
                  Option1(1).Value = True
                End If
                query = "select * from tab_detalhe_declaracao where tdd_num_inscricao =   '" & g_num_inscricao & "'"
                query = query & " and [tdd_periodo_inicial] = '" & g_periodo_inicial & "'"
                query = query & " and [tdd_substitutiva] = " & g_tipo_decla
                Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
            Else
               MsgBox "Utilize o botão ''Diretório'' para selecionar a declaração.", 64, NOME_SISTEMA_DEFAULT
               campopreench = False   ' inclui em 24/2/1998
               fraIdentificação.Enabled = True  ' inclui em 24/2/1998
               gformulario = 0  ' inclui em 24/2/1998
               limpa_tela_Det    ' inclui em 24/2/1998
               SSTab1.Enabled = False  ' inclui em 24/2/1998
               msk_inscricao.SetFocus  ' inclui em 24/2/1998
               Exit Sub
            End If
         End If
      End If
End If
  
 monta_frames
 
      cargar_grid_recolhimento
      cargar_data
      cargar_datos_saidas
      If txtCondContr = "Normal" Or g_aliquota_Maior17 = 1 Then
         frm_grid_fatu.Visible = False
         frm_grid_fatu.Enabled = False
         frm_tela3.Visible = True
         'frm_tela3.Enabled = True
         Frame_tela3.Visible = True
         cargar_grid_ApurICMS
      ElseIf txtCondContr = "Pequeno Porte" Or g_aliquota_Maior17 = 1 Then
         frm_grid_fatu.Visible = False
         frm_grid_fatu.Enabled = False
         frm_tela3.Visible = True
         Frame_tela3.Visible = True
         cargar_grid_ApurICMS
      Else
        frm_tela3.Visible = False
         'frm_tela3.Enabled = False
'         frm_grid_fatu.Visible = True
         frm_grid_fatu.Enabled = True
         cargar_grid_faturamento
      End If

    If g_condicao = 1 Or g_condicao = 3 Or g_condicao = 4 Or g_condicao = 7 Then
       '   fra_EntraMicroEPP.Top = 7000
          fra_EntaNormal.Top = 650
          If g_condicao = 3 Or g_condicao = 4 Then
            'Frase = " Informe a seguir somente o detalhamento das operações com mercadorias cuja alíquota interna no Maranhão seja igual a 25% por cento e proceda a apuração normal."
             Frase = " Declare a seguir todas as operações com mercadorias " _
              & "cuja alíquota interna no Maranhão seja diferenciada (<17% e >17% =25%), " _
              & "registrando e apurando conforme instruções constantes " _
              & "na ajuda, quadro Critérios para Declaração de Contribuinte (Regime EPP)."
            ativa_label_pisca_apur
          End If
          SSTab1.Enabled = True
          SSTab1.Tab = 0
          fra_EntaNormal.Enabled = True
          GRID_DEM_ENTRADAS.Enabled = True
          GRID_DEM_ENTRADAS.SetFocus
          GRID_DEM_ENTRADAS.Refresh
      ElseIf g_condicao = 2 Or g_condicao = 5 Or g_condicao = 6 Then
          fra_EntaNormal.Top = 7000
          fra_EntraMicroEPP.Top = 660
          fra_EntraMicroEPP.Left = 75        'inclusao em 21/2/98 em funcao do 16 bits
          grid_Entra_MicroEPP.Enabled = True
          SSTab1.Enabled = True
          grid_Entra_MicroEPP.Refresh
          grid_Entra_MicroEPP.SetFocus
        End If

   TestaCampoPreench
   
End Sub


Private Sub cmd_buscar_contr_Click()
On Error Resume Next
  
If teve_alteracao_declaracao = True Then
  resposta = enviar_mensagem("Confirma a perda das alterações?")
  If resposta <> 1 Then
    Exit Sub
  End If
End If
gformulario = 0
AtualizaTela
SSTab1.Tab = 0
Me.Hide
quem_me_chamou = "DEC002"
dec005.Show 1
query = "select * from tab_cad_contribuintes where tcc_nu_inscricao = '" & Trim(Str(g_num_inscricao)) & "'"
Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
If Not restauracontr.EOF And Not restauracontr.BOF Then
     Me.Show
     msk_inscricao.Text = g_num_inscricao
     txt_razao = restauracontr.Fields("tcc_razao_social")
     g_num_cgc_cpf = restauracontr.Fields("tcc_cad_geral_contrib")
     g_ano_inicio_operacao = Val(Mid(restauracontr.Fields("tcc_da_inicio_op") & "", 5, 4))
     PER_INICIO = Val(Mid(restauracontr.Fields("tcc_da_inicio_op") & "", 5, 4) & Mid(restauracontr.Fields("tcc_da_inicio_op") & "", 3, 2))
     If g_condicao = 0 Then
              g_condicao = restauracontr.Fields("tcc_tipo_contribuinte")
              Select Case restauracontr.Fields("tcc_tipo_contribuinte")
                 Case 1
                      txtCondContr = "Normal"
                 Case 3, 4, 5, 6, 7
                      txtCondContr = "Pequeno Porte"
                 Case 2
                      txtCondContr = "Microempresa"
              End Select
         Else
             Select Case g_condicao
               Case 1
                  txtCondContr = "Normal"
               Case 3, 4, 5, 6, 7
                  txtCondContr = "Pequeno Porte "
               Case 2
                  txtCondContr = "Microempresa"
            End Select
        End If
        posiciona_label_tipo_decl
        SSTab1.Tab = 0
        carrega_navega   'coloquei 23/3/98
        monta_frames
        txt_periodo_inicial.SetFocus
 Else
    SSTab1.Tab = 0
    monta_frames
    Me.Show
    msk_inscricao.SetFocus
End If

 


   
  
End Sub

Private Sub cmd_buscar_contr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub


Private Sub cmd_diretorio_Click()
On Error Resume Next
  SSTab1.Tab = 0
  gformulario = 2
  Me.Hide
  dec003.Show 1
  Me.Show
  AtualizaTela
  SSTab1.Tab = 0
  monta_frames
  TAB_ANTERIOR = 0
  Select Case g_condicao
      Case 2, 5, 6
   
       fra_EntraMicroEPP.Top = 660
       fra_EntraMicroEPP.Left = 75
       fra_EntaNormal.Enabled = False
       fra_EntaNormal.Visible = False
       fra_EntraMicroEPP.Enabled = True
       fra_EntraMicroEPP.Visible = True
       grid_Entra_MicroEPP.SetFocus
      Case 1, 3, 4, 7
      ' fra_EntraMicroEPP.Left = -12000
      fra_EntaNormal.Enabled = True
      GRID_DEM_ENTRADAS.Enabled = True
      fra_EntaNormal.Visible = True
      GRID_DEM_ENTRADAS.Visible = True
      If g_condicao = 3 Or g_condicao = 4 Then
         If vim_de_onde = "DEC003" Then
            'Frase = " Informe a seguir somente o detalhamento das operações com mercadorias cuja alíquota interna no Maranhão seja igual a 25% por cento e proceda a apuração normal."
             Frase = " Declare a seguir todas as operações com mercadorias " _
              & "cuja alíquota interna no Maranhão seja diferenciada (<17% e > 17% = 25%)," _
              & "registrando e apurando conforme instruções constantes " _
              & "na ajuda, quadro Critérios para Declaração de Contribuinte (Regime EPP)."
            ativa_label_pisca_apur
         End If
      End If

      GRID_DEM_ENTRADAS.MarqueeStyle = 2
      GRID_DEM_ENTRADAS.SetFocus
    End Select
End Sub


Private Sub cmd_diretorio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub


Private Sub cmd_excluir_Click()

On Error Resume Next

Dim sel As Integer
Dim a As Integer
Dim Banco As Database
Dim result As Integer

Dim query2 As String
'sel = DBDados.SelBookmarks.Count - 1

'If DBDados.SelBookmarks.Count <> 0 Then
 resposta = enviar_mensagem("Excluir a declaração atual?")
 
 If resposta = 1 Then
        'For a = 0 To sel
            'Data1.Recordset.Bookmark = DBDados.SelBookmarks(a)
            'DBDados.AllowDelete = True
            query = "Delete * from tab_declaracao  where tde_num_inscricao = '" & Trim(msk_inscricao) _
            & "' And tde_periodo_inicial = '" & g_periodo_inicial _
            & "' And Tde_substitutiva = " & g_tipo_decla
            
            query2 = "Delete * from tab_detalhe_declaracao  where tdd_num_inscricao = '" & Trim(msk_inscricao) _
            & "' And tdd_periodo_inicial = '" & g_periodo_inicial _
            & "' And Tdd_substitutiva = " & g_tipo_decla
                                  
          '  Set Banco = OpenDatabase(App.Path + "\Dief.mdb", dbOpenDynaset)
             If Not EjecutarQuery(query) Then
                Screen.MousePointer = 0
              End If
           ' Banco.Execute query
             If Not EjecutarQuery(query2) Then
                Screen.MousePointer = 0
              End If
          '  Banco.Execute query2
           ' Banco.Close
        ' Next a
        teve_alteracao_declaracao = False
        g_num_inscricao = ""
        g_periodo_inicial = ""
        g_tipo_decla = 0
        g_condicao = 0
        g_ano_inicio_operacao = 0

         cmd_iniciar_Click

    End If
'End If
 
'limpa_tela
'Data1.Refresh
'DBDados.Refresh

End Sub


Private Sub cmd_excluir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub


Private Sub cmd_imprimir_Click()

On Error Resume Next

   Dim recordImp As Recordset
   Dim camposImp As String
   Dim datager As String
   Dim sitDecla As String
   
   Dim result As Integer
   camposImp = "select * from tab_declaracao where tde_num_inscricao =   '" & g_num_inscricao & "'"
   camposImp = camposImp & " and [tde_periodo_inicial] = '" & g_periodo_inicial & "'"
   camposImp = camposImp & " and [tde_substitutiva] = " & g_tipo_decla
   
   Me.MousePointer = 11
   Set recordImp = dbgim.OpenRecordset(camposImp, dbOpenDynaset)
   
   If recordImp.Fields("tde_da_geracao") <> "" Then
      datager = recordImp.Fields("tde_da_geracao")
   Else
      datager = "       "
   End If
   
   sitDecla = recordImp.Fields("tde_Estado")
   
   camposImp = "select * from tab_cad_contribuintes where tcc_nu_inscricao = '" & g_num_inscricao & "'"
   Set recordImp = dbgim.OpenRecordset(camposImp, dbOpenDynaset)
   
 '  result = MsgBox("Confirma impressão? ", vbYesNo, NOME_SISTEMA_DEFAULT)
    
  ' If result = 6 Then
      
      Call Impr_declara(g_num_inscricao, recordImp.Fields("tcc_razao_social"), g_periodo_inicial, sitDecla, datager, recordImp.Fields("tcc_num_crc_contador"), recordImp.Fields("tcc_uf_crc_contador"), g_tipo_decla)
      MsgBox "Fim de Impressão", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
      Me.MousePointer = 0
 '  End If
   
   recordImp.Close
End Sub


Private Sub cmd_imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub


Public Sub cmd_iniciar_Click()

On Error GoTo erro_iniciar

  If teve_alteracao_declaracao = True Then
     resposta = enviar_mensagem("Confirma a perda das alterações?")
     If resposta = 1 Then
       vim_trata_epp = False
        desativa_label_pisca_pisca
        teve_alteracao_declaracao = False
        gformulario = 0
        limpa_tela_Det
        SSTab1.Tab = 0
        monta_frames
        'msk_inscricao.SetFocus
     Else
       Exit Sub
     End If
  Else
       vim_trata_epp = False
        desativa_label_pisca_pisca
        teve_alteracao_declaracao = False
        gformulario = 0
'        AtualizaTela
        limpa_tela_Det
        SSTab1.Tab = 0
        monta_frames
        'msk_inscricao.SetFocus
 End If
 
 Exit Sub
 
erro_iniciar:

     MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     
    Resume Next

End Sub


Private Sub cmd_iniciar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub


Private Sub cmd_sair_Click()
cmd_iniciar_Click
If teve_alteracao_declaracao = True Then
   resposta = enviar_mensagem("Confirma a perda das alterações?")
   If resposta <> 1 Then
     Exit Sub
   End If
   
End If
  
If gformaAnte = 8 Then
   Unload Me
   gformaAnte = 1
   gformulario = 3
   dec003.Show
Else
   gformaAnte = 0
   Unload Me
End If
End Sub


Private Sub cmd_sair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub


Private Sub cmd_salvar_Click()
On Error GoTo erro_salvar

Dim cod_erro As Integer
Dim linha_erro As Integer
Dim linha As Integer
Dim retorno_critica As Double
Dim erro As Integer
Dim valor As String

Screen.MousePointer = 11
If cmb_tipo.ListIndex = 0 Then
    g_tipo_decla = 0
Else
     ' g_tipo_decla = 1
End If
     
gravar_declaracao g_tipo_decla

Entradas:
    
    If Not CRITICA_DEM_ENTRADAS Then
        MsgBox msgerro, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
        Navega_Ficha(1, 0) = 0
        Screen.MousePointer = 0
        Exit Sub
    Else
        gravar_declaracao g_tipo_decla
        Navega_Ficha(1, 0) = 1
        grava_detalhe_dem_entradas
    End If
    'Navega_Ficha(1, 0) = 1

Saidas:
    If Navega_Ficha(1, 1) = 1 Then
        If Not CRITICA_DEM_saidas Then
            MsgBox msgerro, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            Navega_Ficha(1, 1) = 1
             Screen.MousePointer = 0
            Exit Sub
        Else
            gravar_declaracao g_tipo_decla
            grava_detalhe_dem_saidas
            Navega_Ficha(1, 1) = 0
        End If
   End If

Apuracao_do_ICMS:
      If Navega_Ficha(1, 2) = 1 Then
         If Not critica_apuracao_icms Then
            MsgBox msgerro, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
            Navega_Ficha(1, 2) = 1
             Screen.MousePointer = 0
            Exit Sub
         Else
            gravar_declaracao g_tipo_decla
            Navega_Ficha(1, 2) = 1
            If g_condicao = 2 Then
                grava_grid_Faturamento
            Else
                grava_grid_ApurICMS
            End If
         End If
         Navega_Ficha(1, 2) = 0
      End If
      
Recolhimento:
      If Navega_Ficha(1, 3) = 1 Then
          
           If Not critica_recolhimento Then
               MsgBox msgerro, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                Navega_Ficha(1, 3) = 1
                Screen.MousePointer = 0
                Exit Sub
           Else
               grava_grid_recolhimento
           End If
           Navega_Ficha(1, 3) = 0
      End If


Detalhamento:
    If Navega_Ficha(1, 4) = 1 Then
       If Not critica_detalhes Then
           MsgBox msgerro, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           Navega_Ficha(1, 4) = 1
           Screen.MousePointer = 0
           Exit Sub
      Else
           gravar_declaracao g_tipo_decla
           grava_grid_DetDesp
           Navega_Ficha(1, 4) = 1
      End If
      Navega_Ficha(1, 4) = 0
    End If
      
    'verificar se existe registro detalhe gravado
    Dim snpadetalhe As Recordset
    Dim indficha As Integer
    teve_alteracao_declaracao = False
    For indficha = 0 To 4
          If Navega_Ficha(0, indficha) <> Navega_Ficha(1, indficha) Then
             Exit For
          End If
    Next indficha
     
    query = "select  Count(*) as qtdreg from tab_detalhe_declaracao where tdd_num_inscricao =   '" & g_num_inscricao & "'"
    query = query & " and [tdd_periodo_inicial] = '" & g_periodo_inicial & "'"
    query = query & " and [tdd_substitutiva] = " & g_tipo_decla
    Set snpadetalhe = dbgim.OpenRecordset(query, dbOpenDynaset)
    If snpadetalhe!qtdReg = 0 Then
       If indficha = 5 Then
          resposta = MsgBox("A Declaração não tem nenhuma ficha preenchida. Ela vai ser considerada finalizada ,tem certeza da gravação?", vbYesNo + vbDefaultButton1, NOME_SISTEMA_DEFAULT)
          If resposta = 6 Then
             If atualiza_declaracao(g_tipo_decla) Then
                MsgBox "Gravação efetuada com Sucesso.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
                Situacao.Caption = "Incompleta"
             End If
          End If
       End If
    Else
        If atualiza_declaracao(g_tipo_decla) Then
           MsgBox "Gravação efetuada com Sucesso.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           Situacao.Caption = "Incompleta"
         End If
    End If
        
    snpadetalhe.Close
    Screen.MousePointer = 0
    teve_alteracao_declaracao = False
    Exit Sub
  
erro_salvar:

     MsgBox "Houve Error na Gravação Informe o Código :" & Str(Err) & " - " & Error, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
     Resume Next


End Sub
Private Sub cmd_salvar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub


Private Sub CmdFinalizar_Click()

Dim ret As Integer
Finalizar_Declara
Exit Sub



End Sub



Private Sub Form_Activate()
On Error Resume Next
CentrarFormu Me
If msk_inscricao.Enabled = True Then
   msk_inscricao.SetFocus
End If

If gformaAnte = 8 Then
   cmd_diretorio.Enabled = False
   cmd_excluir.Enabled = False
Else
   cmd_diretorio.Enabled = True
   cmd_excluir.Enabled = True
End If

Navega_Ficha(0, 0) = 1
Navega_Ficha(0, 1) = 1
Navega_Ficha(0, 2) = 1
Navega_Ficha(0, 3) = 1
Navega_Ficha(0, 4) = 1
Navega_Ficha(1, 0) = 0
Navega_Ficha(1, 1) = 0
Navega_Ficha(1, 2) = 0
Navega_Ficha(1, 3) = 0
Navega_Ficha(1, 4) = 0


End Sub

Private Sub Form_Load()
    
Me.Top = 0
Me.Left = 0
TAB_ANTERIOR = 9
dec002.Enabled = True
inicia_Dbgrid
APENAS_ENTREI = False
MDIdec.Picture1.Visible = False
MDIdec.Enabled = False
teve_alteracao_declaracao = False
motivo_percentual = "Deve ser: O imposto Creditado/Debitado deve ser até 25% da Base de Calculo."
motivo_entradas = "Deve ser: Valores Contábeis = Base de Cálculo + Isentas não Tributadas + Outras."
apenas_troquei_estado = 0
TAB_ANTERIOR = 9
teve_alteracao_declaracao = False
Navega_Ficha(0, 0) = 1
Navega_Ficha(0, 1) = 1
Navega_Ficha(0, 2) = 1
Navega_Ficha(0, 3) = 1
Navega_Ficha(0, 4) = 1
Navega_Ficha(1, 0) = 0
Navega_Ficha(1, 1) = 0
Navega_Ficha(1, 2) = 0
Navega_Ficha(1, 3) = 0
Navega_Ficha(1, 4) = 0
erro_inscricao = False
monta_frames
vim_trata_epp = False

If vim_trata_epp = True Then Exit Sub
    
    AtualizaTela
    SSTab1.Tab = 0
    TAB_ANTERIOR = 0
    
    Select Case g_condicao
      
      Case 2, 5, 6
        fra_EntraMicroEPP.Top = 660
        fra_EntraMicroEPP.Left = 75
        fra_EntaNormal.Enabled = False
        fra_EntaNormal.Visible = False
        fra_EntraMicroEPP.Enabled = True
        fra_EntraMicroEPP.Visible = True
      
      Case 1, 3, 4, 7
      
        fra_EntaNormal.Enabled = True
        GRID_DEM_ENTRADAS.Enabled = True
        fra_EntaNormal.Visible = True
        GRID_DEM_ENTRADAS.Visible = True
        If g_condicao = 3 Or g_condicao = 4 Then
           If vim_de_onde = "DEC003" Then
             'Frase = " Informe a seguir somente o detalhamento das operações com mercadorias cuja alíquota interna no Maranhão seja igual a 25% por cento e proceda a apuração normal."
              Frase = " Declare a seguir todas as operações com mercadorias " _
              & "cuja alíquota interna no Maranhão seja diferenciada (<17% e >17% =25%), " _
              & "registrando e apurando conforme instruções constantes " _
              & "na ajuda, quadro Critérios para Declaração de Contribuinte (Regime EPP)."
              ativa_label_pisca_apur
           End If
        End If
        GRID_DEM_ENTRADAS.MarqueeStyle = 2
      
End Select
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

dec002.Enabled = False

End Sub


Private Sub Form_Unload(Cancel As Integer)
MDIdec.Picture1.Visible = True
MDIdec.Enabled = True
MDIdec.Picture1.SetFocus
End Sub

Private Sub fraIdentificação_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub




Private Sub grdDetDesp_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

Dim ret As Integer

If Not critica_valor_com_caracteres(grdDetDesp.Columns(ColIndex).Text) Then
  grdDetDesp.Columns(ColIndex).Text = "0,00"
End If
grdDetDesp.Columns(ColIndex).Text = Format(grdDetDesp.Columns(ColIndex).Text, "#,##0.#0")
End Sub

Private Sub grdDetDesp_Change()

Navega_Ficha(1, 4) = 1

End Sub


Private Sub grdDetDesp_GotFocus()
 
txt415 = acumula_grid_detalhes(grid_DetDesp)
txt415.Text = Format(txt415.Text, "#,##0.#0")

grid_DetDesp(1, 14) = txt415.Text
grdDetDesp.Col = 1
grdDetDesp.MarqueeStyle = 2

End Sub

Private Sub grdDetDesp_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = 39 Or KeyCode = 40 Then
       KeyCode = 13
  End If
      
   If KeyCode = 13 Then
       txt415 = acumula_grid_detalhes(grid_DetDesp)
       txt415.Text = Format(txt415.Text, "#,##0.#0")
       grid_DetDesp(1, 14) = txt415.Text
      If grdDetDesp.Row = 13 And grdDetDesp.Col <> 1 Then
         grdDetDesp.Col = grdDetDesp.Col
         grdDetDesp.Row = grdDetDesp.Row - 1
         grdDetDesp.Row = grdDetDesp.Row + 1
      ElseIf grdDetDesp.Row = 13 And grdDetDesp.Col = 1 Then
         grdDetDesp.Col = grdDetDesp.Col
         grdDetDesp.Row = 0
'         grdDetDesp.Row = grdDetDesp.Row + 1
      ElseIf grdDetDesp.Row > 13 And grdDetDesp.Col > 1 Then
         grdDetDesp.Col = 0
         grdDetDesp.Row = 0
      End If
   End If
   
   If KeyCode = 9 Or KeyCode = 13 Then
       If grdDetDesp.Row >= 13 Then
          If Frame9.Visible = True Then
            grdDetDesp.Col = 0
            grdDetDesp.Row = 0
            txt400.SetFocus
          Else
            grdDetDesp.Col = 0
            grdDetDesp.Row = grdDetDesp.Row + 1
          End If
          Exit Sub
       Else
           grdDetDesp.Col = 0
           grdDetDesp.Row = grdDetDesp.Row + 1
       End If
   End If
  If grdDetDesp.Col = 1 Then
     If KeyCode = 37 Or KeyCode = 39 Then
         KeyCode = 0
         grdDetDesp.Col = 1
     End If
  End If
End Sub


Private Sub grdDetDesp_KeyPress(KeyAscii As Integer)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub grdDetDesp_LostFocus()
   grdDetDesp.MarqueeStyle = 5

End Sub


Private Sub grdDetDesp_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   
   If SSTab1.Enabled Then
      If grdDetDesp.Col = 0 And LastCol > -1 Then
         grdDetDesp.Col = 1
   '      grdDetDesp.Row = grdDetDesp.Row - 1
      End If
   End If

End Sub


Private Sub grdDetDesp_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
   Dim i As Integer
   On Error Resume Next
   NewRowBookmark = GetNewBookmark_DetDesp()
   
   ' Loop over all the columns of the row, storing non-Null '
   ' values
       Dim newval As Variant
       For i% = 0 To RowBuf.ColumnCount - 1
           newval = RowBuf.Value(0, i%)
       
           If IsNull(newval) Then
   ' the RowBuf does not contain a value for this column.
   ' A default value should be set.  A convenient value
   ' is the default value for the column.
               newval = grdDetDesp.Columns(i%).DefaultValue
           End If
   ' Now store the new values.
           If Not StoreUserData_DetDesp(NewRowBookmark, i%, newval) Then
   ' storage of the data has failed.  Delete the added
   ' row using a VBA function DeleteRow which takes a
   ' bookmark as an argument.  Also, fail the update by
   ' clearing the RowCount.
               DeleteRow_DetDesp NewRowBookmark
               RowBuf.RowCount = 0 ' tell the grid the update failed
               Exit Sub            ' it failed, so exit the event
           End If
       Next i%

End Sub


Private Sub grdDetDesp_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
   On Error Resume Next
   Dim bookm As Variant
       Dim rowsFetched As Integer
       Dim relpos As Integer
       Dim i As Integer
       Dim j As Integer
       
       bookm = StartLocation
       If ReadPriorRows Then
   ' the grid is requesting data in rows prior to
   ' StartLocation
           relpos = -1
       Else
   ' the grid is requesting data in rows after
   ' StartLocation
           relpos = 1
       End If
    
       rowsFetched = 0
       For i% = 0 To RowBuf.RowCount - 1
           If (RowBuf.RowCount - 1) <> 0 Then
               ' Get the bookmark of the next available row
               bookm = GetRelativeBookmark_DetDesp(bookm, relpos)
           Else
               bookm = StartLocation
           End If
   ' If the next is BOF or EOF, then done
           If IsNull(bookm) Then Exit For
           For j% = 0 To RowBuf.ColumnCount - 1
               RowBuf.Value(i%, j%) = GetUserData_DetDesp(bookm, j%)
           Next j%
   ' Set the bookmark for the row
           RowBuf.Bookmark(i%) = bookm
   ' Increment the count of fetched rows
           rowsFetched = rowsFetched + 1
       Next i%
   
   ' tell the grid how many rows were fetched
       RowBuf.RowCount = rowsFetched


End Sub


Private Sub grdDetDesp_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
 Dim i As Integer
   On Error Resume Next
   ' Loop over all the columns of the row, storing non-Null
   ' values
       For i% = 0 To RowBuf.ColumnCount - 1
           If Not IsNull(RowBuf.Value(0, i%)) Then
               If Not StoreUserData_DetDesp(WriteLocation, i%, RowBuf.Value(0, i%)) Then
    ' storage of the data has failed.  Fail the update
                   RowBuf.RowCount = 0     ' tell the grid the update failed
                   Exit Sub                 ' it failed, so exit the event
               End If
           End If
       Next i%



    txt415 = acumula_grid_detalhes(grid_DetDesp)
       txt415.Text = Format(txt415.Text, "#,##0.#0")
   
         grid_DetDesp(1, 14) = txt415.Text
      
  
      
'txt314.Refresh
End Sub


Private Sub GRID_DEM_ENTRADAS_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  Dim ret As Integer

  If Not critica_valor_com_caracteres(GRID_DEM_ENTRADAS.Columns(ColIndex).Text) Then
    GRID_DEM_ENTRADAS.Columns(ColIndex).Text = "0,00"
  
  End If
  
  GRID_DEM_ENTRADAS.Columns(ColIndex).Text = Format(GRID_DEM_ENTRADAS.Columns(ColIndex).Text, "#,##0.#0")

End Sub

Private Sub GRID_DEM_ENTRADAS_Change()
    teve_alteracao_declaracao = True
End Sub


Private Sub GRID_DEM_ENTRADAS_GotFocus()

   GRID_DEM_ENTRADAS.Col = 1
   GRID_DEM_ENTRADAS.MarqueeStyle = 2

End Sub

Private Sub GRID_DEM_ENTRADAS_KeyDown(KeyCode As Integer, Shift As Integer)
     
If KeyCode = 40 And GRID_DEM_ENTRADAS.Col < 5 And GRID_DEM_ENTRADAS.Row >= 11 Then
       GRID_DEM_ENTRADAS.Col = GRID_DEM_ENTRADAS.Col + 1
       GRID_DEM_ENTRADAS.Row = 0
       KeyCode = 0
       Exit Sub
ElseIf KeyCode = 40 And GRID_DEM_ENTRADAS.Col = 5 And GRID_DEM_ENTRADAS.Row < 11 Then
      GRID_DEM_ENTRADAS.Row = GRID_DEM_ENTRADAS.Row + 1
      GRID_DEM_ENTRADAS.Row = GRID_DEM_ENTRADAS.Row - 1
      Exit Sub
ElseIf KeyCode = 40 And GRID_DEM_ENTRADAS.Col = 5 And GRID_DEM_ENTRADAS.Row >= 11 Then
       GRID_DEM_ENTRADAS.Col = 1
       GRID_DEM_ENTRADAS.Row = 0
       KeyCode = 0
       Exit Sub
End If
  
  If KeyCode = 39 Then
       KeyCode = 13
  End If
  
 If KeyCode = 13 Then
   desativa_label_pisca_pisca
   If GRID_DEM_ENTRADAS.Row = 11 And GRID_DEM_ENTRADAS.Col <> 5 Then
      GRID_DEM_ENTRADAS.Col = GRID_DEM_ENTRADAS.Col
      GRID_DEM_ENTRADAS.Row = GRID_DEM_ENTRADAS.Row - 1
      GRID_DEM_ENTRADAS.Row = GRID_DEM_ENTRADAS.Row + 1
   ElseIf GRID_DEM_ENTRADAS.Row = 11 And GRID_DEM_ENTRADAS.Col = 5 Then
      GRID_DEM_ENTRADAS.Col = GRID_DEM_ENTRADAS.Col
      GRID_DEM_ENTRADAS.Row = GRID_DEM_ENTRADAS.Row - 1
      GRID_DEM_ENTRADAS.Row = GRID_DEM_ENTRADAS.Row + 1
   ElseIf GRID_DEM_ENTRADAS.Row > 11 And GRID_DEM_ENTRADAS.Col > 5 Then
      GRID_DEM_ENTRADAS.Col = 1
      GRID_DEM_ENTRADAS.Row = 0
      KeyCode = 0
    End If
 End If

 If GRID_DEM_ENTRADAS.Col = 5 Then
    If KeyCode = 9 Or KeyCode = 13 Or KeyCode = 40 Then
       If GRID_DEM_ENTRADAS.Row >= 11 Then
          GRID_DEM_ENTRADAS.Col = 0
          GRID_DEM_ENTRADAS.Row = 0
          Exit Sub
       Else
          GRID_DEM_ENTRADAS.Col = 0
          GRID_DEM_ENTRADAS.Row = GRID_DEM_ENTRADAS.Row + 1
       End If
     End If
  End If
  
End Sub


Private Sub GRID_DEM_ENTRADAS_KeyPress(KeyAscii As Integer)
   
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
     KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
     KeyAscii = Asc(",")
   End If


End Sub


Private Sub GRID_DEM_ENTRADAS_LostFocus()

  GRID_DEM_ENTRADAS.MarqueeStyle = 5

End Sub


Private Sub GRID_DEM_ENTRADAS_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   
   If SSTab1.Enabled Then
      If GRID_DEM_ENTRADAS.Col = 0 And LastCol > -1 Then
         GRID_DEM_ENTRADAS.Col = LastCol
         GRID_DEM_ENTRADAS.Row = LastRow
      End If
  
   End If
   

End Sub


Private Sub GRID_DEM_ENTRADAS_SelChange(Cancel As Integer)

   If GRID_DEM_ENTRADAS.Col > 0 Then
     If GRID_DEM_ENTRADAS.Columns((GRID_DEM_ENTRADAS.Col)).Text <> GridArray((GRID_DEM_ENTRADAS.Col), GRID_DEM_ENTRADAS.Row) Then
       desativa_label_pisca_pisca
       GridArray((GRID_DEM_ENTRADAS.Col), GRID_DEM_ENTRADAS.Row) = GRID_DEM_ENTRADAS.Columns((GRID_DEM_ENTRADAS.Col)).Text
     End If
  End If

End Sub

Private Sub GRID_DEM_ENTRADAS_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim i As Integer
'' On Error Resume Next
NewRowBookmark = GetNewBookmark()

' Loop over all the columns of the row, storing non-Null '
' values
    Dim newval As Variant
    For i% = 0 To RowBuf.ColumnCount - 1
        newval = RowBuf.Value(0, i%)
    
        If IsNull(newval) Then
' the RowBuf does not contain a value for this column.
' A default value should be set.  A convenient value
' is the default value for the column.
            newval = GRID_DEM_ENTRADAS.Columns(i%).DefaultValue
        End If
' Now store the new values.
        If Not StoreUserData(NewRowBookmark, i%, newval) Then
' storage of the data has failed.  Delete the added
' row using a VBA function DeleteRow which takes a
' bookmark as an argument.  Also, fail the update by
' clearing the RowCount.
            DeleteRow NewRowBookmark
            RowBuf.RowCount = 0 ' tell the grid the update failed
            Exit Sub            ' it failed, so exit the event
        End If
    Next i%

End Sub


Private Sub GRID_DEM_ENTRADAS_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    On Error Resume Next
    Dim bookm As Variant
    Dim rowsFetched As Integer
    Dim relpos As Integer
    Dim i As Integer
    Dim j As Integer
    
    bookm = StartLocation
    If ReadPriorRows Then
        ' the grid is requesting data in rows prior to
        ' StartLocation
        relpos = -1
    Else
        ' the grid is requesting data in rows after
        ' StartLocation
        relpos = 1
    End If
 
    rowsFetched = 0
    For i% = 0 To RowBuf.RowCount - 1
        If (RowBuf.RowCount - 1) <> 0 Then
            ' Get the bookmark of the next available row
            bookm = GetRelativeBookmark(bookm, relpos)
        Else
            bookm = StartLocation
        End If
' If the next is BOF or EOF, then done
        If IsNull(bookm) Then Exit For
        For j% = 0 To RowBuf.ColumnCount - 1
            RowBuf.Value(i%, j%) = GetUserData(bookm, j%)
        Next j%
' Set the bookmark for the row
        RowBuf.Bookmark(i%) = bookm
' Increment the count of fetched rows
        rowsFetched = rowsFetched + 1
    Next i%

' tell the grid how many rows were fetched
    RowBuf.RowCount = rowsFetched


End Sub


Private Sub GRID_DEM_ENTRADAS_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim i As Integer
On Error Resume Next
' Loop over all the columns of the row, storing non-Null
' values
    For i% = 0 To RowBuf.ColumnCount - 1
        If Not IsNull(RowBuf.Value(0, i%)) Then
            If Not StoreUserData(WriteLocation, i%, RowBuf.Value(0, i%)) Then
                ' storage of the data has failed.  Fail the update
                RowBuf.RowCount = 0     ' tell the grid the update failed
                Exit Sub                ' it failed, so exit the event
            End If
        End If
    Next i%
    
End Sub


Private Sub grid_dem_saidas_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim ret As Integer

If Not critica_valor_com_caracteres(grid_dem_saidas.Columns(ColIndex).Text) Then
  grid_dem_saidas.Columns(ColIndex).Text = "0,00"
End If
  
grid_dem_saidas.Columns(ColIndex).Text = Format(grid_dem_saidas.Columns(ColIndex).Text, "#,##0.#0")
  
End Sub

Private Sub grid_dem_saidas_Change()

teve_alteracao_declaracao = True

End Sub


Private Sub grid_dem_saidas_GotFocus()
   
  
   grid_dem_saidas.Col = 1
   grid_dem_saidas.MarqueeStyle = 2

End Sub

Private Sub grid_dem_saidas_KeyDown(KeyCode As Integer, Shift As Integer)
   
If KeyCode = 40 And grid_dem_saidas.Col < 5 And grid_dem_saidas.Row >= 11 Then
       grid_dem_saidas.Col = grid_dem_saidas.Col + 1
       grid_dem_saidas.Row = 0
       KeyCode = 0
       Exit Sub
ElseIf KeyCode = 40 And grid_dem_saidas.Col = 5 And grid_dem_saidas.Row < 11 Then
      grid_dem_saidas.Row = grid_dem_saidas.Row + 1
      grid_dem_saidas.Row = grid_dem_saidas.Row - 1
      Exit Sub
ElseIf KeyCode = 40 And grid_dem_saidas.Col = 5 And grid_dem_saidas.Row >= 11 Then
       grid_dem_saidas.Col = 1
       grid_dem_saidas.Row = 0
       KeyCode = 0
       Exit Sub
End If
  
If KeyCode = 39 Then
     KeyCode = 13
End If

   
If KeyCode = 9 Or KeyCode = 13 Then
    desativa_label_pisca_pisca
     If grid_dem_saidas.Row = 11 And grid_dem_saidas.Col <> 5 Then
        grid_dem_saidas.Col = grid_dem_saidas.Col
        grid_dem_saidas.Row = grid_dem_saidas.Row - 1
        grid_dem_saidas.Row = grid_dem_saidas.Row + 1
     ElseIf grid_dem_saidas.Row = 11 And grid_dem_saidas.Col = 5 Then
        grid_dem_saidas.Col = grid_dem_saidas.Col
        grid_dem_saidas.Row = grid_dem_saidas.Row - 1
        grid_dem_saidas.Row = grid_dem_saidas.Row + 1
     ElseIf grid_dem_saidas.Row > 11 And grid_dem_saidas.Col > 5 Then
        grid_dem_saidas.Col = 1
        grid_dem_saidas.Row = 0
        KeyCode = 0
     End If
    
End If
 If grid_dem_saidas.Col = 5 Then
    If KeyCode = 9 Or KeyCode = 13 Then
      If grid_dem_saidas.Row >= 11 Then
        grid_dem_saidas.Row = 0
        grid_dem_saidas.Col = 0 ' 1
        KeyCode = 0
        Exit Sub
      Else
        grid_dem_saidas.Col = 0
        grid_dem_saidas.Row = grid_dem_saidas.Row + 1
      End If
   End If
End If

End Sub


Private Sub grid_dem_saidas_KeyPress(KeyAscii As Integer)

 If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
        KeyAscii = 0
 ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
        KeyAscii = Asc(",")
 End If

End Sub


Private Sub grid_dem_saidas_LostFocus()
  
  grid_dem_saidas.MarqueeStyle = 5

End Sub


Private Sub grid_dem_saidas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

     If SSTab1.Enabled Then
        If grid_dem_saidas.Col = 0 And LastCol > -1 Then
           grid_dem_saidas.Col = LastCol
           grid_dem_saidas.Row = LastRow
        End If
     End If

End Sub


Private Sub grid_dem_saidas_SelChange(Cancel As Integer)

   If grid_dem_saidas.Col > 0 Then
     If grid_dem_saidas.Columns((grid_dem_saidas.Col)).Text <> grid_saidas((grid_dem_saidas.Col), grid_dem_saidas.Row) Then
        desativa_label_pisca_pisca
        grid_saidas((grid_dem_saidas.Col), grid_dem_saidas.Row) = grid_dem_saidas.Columns((grid_dem_saidas.Col)).Text
     End If
  End If



End Sub

Private Sub grid_dem_saidas_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim i As Integer
On Error Resume Next
NewRowBookmark = GetNewBookmark_s()

' Loop over all the columns of the row, storing non-Null '
' values
    Dim newval As Variant
    For i% = 0 To RowBuf.ColumnCount - 1
        newval = RowBuf.Value(0, i%)
    
        If IsNull(newval) Then
' the RowBuf does not contain a value for this column.
' A default value should be set.  A convenient value
' is the default value for the column.
            newval = grid_dem_saidas.Columns(i%).DefaultValue
        End If
' Now store the new values.
        If Not storeUserData_s(NewRowBookmark, i%, newval) Then
' storage of the data has failed.  Delete the added
' row using a VBA function DeleteRow which takes a
' bookmark as an argument.  Also, fail the update by
' clearing the RowCount.
            DeleteRow_s NewRowBookmark
            RowBuf.RowCount = 0 ' tell the grid the update failed
            Exit Sub            ' it failed, so exit the event
        End If
    Next i%



End Sub


Private Sub grid_dem_saidas_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    On Error Resume Next
    Dim bookm As Variant
    Dim rowsFetched As Integer
    Dim relpos As Integer
    Dim i As Integer
    Dim j As Integer
    
    bookm = StartLocation
    If ReadPriorRows Then
' the grid is requesting data in rows prior to
' StartLocation
        relpos = -1
    Else
' the grid is requesting data in rows after
' StartLocation
        relpos = 1
    End If
 
    rowsFetched = 0
    For i% = 0 To RowBuf.RowCount - 1
        If (RowBuf.RowCount - 1) <> 0 Then
            ' Get the bookmark of the next available row
            bookm = GetRelativeBookmark_s(bookm, relpos)
        Else
            bookm = StartLocation
        End If
' If the next is BOF or EOF, then done
        If IsNull(bookm) Then Exit For
        For j% = 0 To RowBuf.ColumnCount - 1
            RowBuf.Value(i%, j%) = GetUserData_s(bookm, j%)
        Next j%
' Set the bookmark for the row
        RowBuf.Bookmark(i%) = bookm
' Increment the count of fetched rows
        rowsFetched = rowsFetched + 1
    Next i%

' tell the grid how many rows were fetched
    RowBuf.RowCount = rowsFetched


End Sub


Private Sub grid_dem_saidas_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim i As Integer
On Error Resume Next
' Loop over all the columns of the row, storing non-Null
' values
    For i% = 0 To RowBuf.ColumnCount - 1
        If Not IsNull(RowBuf.Value(0, i%)) Then
            If Not storeUserData_s(WriteLocation, i%, RowBuf.Value(0, i%)) Then
                ' storage of the data has failed.  Fail the update
                RowBuf.RowCount = 0     ' tell the grid the update failed
                Exit Sub                ' it failed, so exit the event
            End If
        End If
    Next i%

End Sub


Private Sub grid_Entra_MicroEPP_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

Dim ret As Integer

  If Not critica_valor_com_caracteres(grid_Entra_MicroEPP.Columns(ColIndex).Text) Then
    grid_Entra_MicroEPP.Columns(ColIndex).Text = "0,00"
  End If


   grid_Entra_MicroEPP.Columns(ColIndex).Text = Format(grid_Entra_MicroEPP.Columns(ColIndex).Text, "#,##0.#0")
  
  

End Sub

Private Sub grid_Entra_MicroEPP_GotFocus()

   grid_Entra_MicroEPP.Col = 1
   grid_Entra_MicroEPP.MarqueeStyle = 2

End Sub

Private Sub grid_Entra_MicroEPP_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 39 Or KeyCode = 40 Then
      KeyCode = 13
  End If
   
   If KeyCode = 13 Then
        
      If grid_Entra_MicroEPP.Row = 11 And grid_Entra_MicroEPP.Col <> 1 Then
         grid_Entra_MicroEPP.Col = grid_Entra_MicroEPP.Col
         grid_Entra_MicroEPP.Row = grid_Entra_MicroEPP.Row - 1
         grid_Entra_MicroEPP.Row = grid_Entra_MicroEPP.Row + 1
      ElseIf grid_Entra_MicroEPP.Row = 11 And grid_Entra_MicroEPP.Col = 1 Then
         grid_Entra_MicroEPP.Col = grid_Entra_MicroEPP.Col
         grid_Entra_MicroEPP.Row = grid_Entra_MicroEPP.Row - 1
         grid_Entra_MicroEPP.Row = grid_Entra_MicroEPP.Row + 1
      ElseIf grid_Entra_MicroEPP.Row > 11 And grid_Entra_MicroEPP.Col > 1 Then
         grid_Entra_MicroEPP.Col = 0
         grid_Entra_MicroEPP.Row = 0
      End If
   End If
      
      If grid_Entra_MicroEPP.Col = 1 Then
         If KeyCode = 9 Or KeyCode = 13 Then
            If grid_Entra_MicroEPP.Row >= 11 Then
               grid_Entra_MicroEPP.Col = 0
               grid_Entra_MicroEPP.Row = 0
               Exit Sub
            Else
               grid_Entra_MicroEPP.Col = 0
               grid_Entra_MicroEPP.Row = grid_Entra_MicroEPP.Row + 1
            End If
          End If
        
       End If

End Sub


Private Sub grid_Entra_MicroEPP_KeyPress(KeyAscii As Integer)
   
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub

Private Sub grid_Entra_MicroEPP_LostFocus()
  
  grid_Entra_MicroEPP.MarqueeStyle = 5


End Sub


Private Sub grid_Entra_MicroEPP_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   
   If SSTab1.Enabled Then
      If grid_Entra_MicroEPP.Col = 0 And LastCol > -1 Then
         grid_Entra_MicroEPP.Col = LastCol
         grid_Entra_MicroEPP.Row = LastRow
      End If
   End If

End Sub

Private Sub grid_Entra_MicroEPP_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim i As Integer
On Error Resume Next
NewRowBookmark = GetNewBookmark()

' Loop over all the columns of the row, storing non-Null '
' values
    Dim newval As Variant
    For i% = 0 To RowBuf.ColumnCount - 1
        newval = RowBuf.Value(0, i%)
    
        If IsNull(newval) Then
' the RowBuf does not contain a value for this column.
' A default value should be set.  A convenient value
' is the default value for the column.
            newval = grid_Entra_MicroEPP.Columns(i%).DefaultValue
        End If
' Now store the new values.
        If Not StoreUserData(NewRowBookmark, i%, newval) Then
' storage of the data has failed.  Delete the added
' row using a VBA function DeleteRow which takes a
' bookmark as an argument.  Also, fail the update by
' clearing the RowCount.
            DeleteRow NewRowBookmark
            RowBuf.RowCount = 0 ' tell the grid the update failed
            Exit Sub            ' it failed, so exit the event
        End If
    Next i%


End Sub


Private Sub grid_Entra_MicroEPP_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    On Error Resume Next
    Dim bookm As Variant
    Dim rowsFetched As Integer
    Dim relpos As Integer
    Dim i As Integer
    Dim j As Integer
    
    bookm = StartLocation
    If ReadPriorRows Then
' the grid is requesting data in rows prior to
' StartLocation
        relpos = -1
    Else
' the grid is requesting data in rows after
' StartLocation
        relpos = 1
    End If
 
    rowsFetched = 0
    For i% = 0 To RowBuf.RowCount - 1
        If (RowBuf.RowCount - 1) <> 0 Then
            ' Get the bookmark of the next available row
            bookm = GetRelativeBookmark(bookm, relpos)
        Else
            bookm = StartLocation
        End If
' If the next is BOF or EOF, then done
        If IsNull(bookm) Then Exit For
        For j% = 0 To RowBuf.ColumnCount - 1
            RowBuf.Value(i%, j%) = GetUserData(bookm, j%)
        Next j%
' Set the bookmark for the row
        RowBuf.Bookmark(i%) = bookm
' Increment the count of fetched rows
        rowsFetched = rowsFetched + 1
    Next i%

' tell the grid how many rows were fetched
    RowBuf.RowCount = rowsFetched



End Sub


Private Sub grid_Entra_MicroEPP_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim i As Integer
On Error Resume Next
' Loop over all the columns of the row, storing non-Null
' values
    For i% = 0 To RowBuf.ColumnCount - 1
        If Not IsNull(RowBuf.Value(0, i%)) Then
            If Not StoreUserData(WriteLocation, i%, RowBuf.Value(0, i%)) Then
 ' storage of the data has failed.  Fail the update
                RowBuf.RowCount = 0     ' tell the grid the update failed
                Exit Sub                 ' it failed, so exit the event
            End If
        End If
    Next i%

  

End Sub


Private Sub grid_faturamento_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

Dim ret As Integer

  If Not critica_valor_com_caracteres(grid_faturamento.Columns(ColIndex).Text) Then
    grid_faturamento.Columns(ColIndex).Text = "0,00"
  End If


   grid_faturamento.Columns(ColIndex).Text = Format(grid_faturamento.Columns(ColIndex).Text, "#,##0.#0")


End Sub

Private Sub grid_faturamento_Change()
Navega_Ficha(1, 2) = 1
End Sub


Private Sub grid_faturamento_GotFocus()
   
   

       txt229 = acumula_grid_faturamento(grid_fatura, 1)
       txt229.Text = Format(txt229.Text, "#,##0.#0")
      grid_fatura(1, 3) = txt229.Text
 
       txt242.Text = acumula_grid_faturamento(grid_fatura, 3)
       txt242.Text = Format(txt242.Text, "#,##0.#0")
      grid_fatura(3, 3) = txt242.Text
  


   
   grid_faturamento.Col = 1
   grid_faturamento.MarqueeStyle = 2


End Sub

Private Sub grid_faturamento_KeyDown(KeyCode As Integer, Shift As Integer)


    If KeyCode = 13 Then
   
        Select Case grid_faturamento.Col
           Case Is = 1
               txt229 = acumula_grid_faturamento(grid_fatura, grid_faturamento.Col)
               txt229.Text = Format(txt229.Text, "#,##0.#0")
              grid_fatura(1, 3) = txt229.Text
           Case Is = 2
                If grid_faturamento.Row = 0 Then
                   txt236.Text = calcula_apur_imposto_Faturamento_Versus_Aliquota(grid_fatura(1, grid_faturamento.Row), grid_faturamento.Columns(grid_faturamento.Col).Text)
                   txt236.Text = Format(txt236.Text, "#,##0.#0")
                   grid_fatura(3, grid_faturamento.Row) = txt236.Text
                ElseIf grid_faturamento.Row = 1 Then
                   txt237.Text = calcula_apur_imposto_Faturamento_Versus_Aliquota(grid_fatura(1, grid_faturamento.Row), grid_faturamento.Columns(grid_faturamento.Col).Text)
                   txt237.Text = Format(txt237.Text, "#,##0.#0")
                   grid_fatura(3, grid_faturamento.Row) = txt237.Text
                ElseIf grid_faturamento.Row = 2 Then
                   txt238.Text = calcula_apur_imposto_Faturamento_Versus_Aliquota(grid_fatura(1, grid_faturamento.Row), grid_faturamento.Columns(grid_faturamento.Col).Text)
                   txt238.Text = Format(txt238.Text, "#,##0.#0")
                   grid_fatura(3, grid_faturamento.Row) = txt238.Text
                End If
        '   Case Is = 3
               txt242.Text = acumula_grid_faturamento(grid_fatura, 3)
               txt242.Text = Format(txt242.Text, "#,##0.#0")
              grid_fatura(3, 3) = txt242.Text
          
        End Select


     ' If grid_faturamento.Row = 5 And grid_faturamento.Col = 1 Then
     '      grid_faturamento.Col = 3
     '      Exit Sub
     ' End If
      If grid_faturamento.Row = 2 And grid_faturamento.Col <> 2 Then
         grid_faturamento.Col = grid_faturamento.Col
         grid_faturamento.Row = grid_faturamento.Row - 1
         grid_faturamento.Row = grid_faturamento.Row + 1
      ElseIf grid_faturamento.Row = 2 And grid_faturamento.Col = 2 Then
         grid_faturamento.Col = grid_faturamento.Col
         grid_faturamento.Row = grid_faturamento.Row - 1
         grid_faturamento.Row = grid_faturamento.Row + 1
      ElseIf grid_faturamento.Row > 2 And grid_faturamento.Col > 2 Then
         grid_faturamento.Col = 0
         grid_faturamento.Row = 0
         KeyCode = 0
      Else
         'grid_faturamento.Col = grid_faturamento.Col - 1
         grid_faturamento.Row = grid_faturamento.Row + 1
         grid_faturamento.Row = grid_faturamento.Row - 1

      End If
   End If
      
      
      If grid_faturamento.Col = 2 Then
         If KeyCode = 9 Or KeyCode = 13 Then
            If grid_faturamento.Row >= 2 Then
               grid_faturamento.Col = 0
               grid_faturamento.Row = 0
               KeyCode = 0
               Exit Sub
            Else
               grid_faturamento.Col = 0
               grid_faturamento.Row = grid_faturamento.Row + 1
            End If
          End If
        
       End If




End Sub


Private Sub grid_faturamento_KeyPress(KeyAscii As Integer)

   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If


End Sub

Private Sub grid_faturamento_LostFocus()
  
  grid_faturamento.MarqueeStyle = 5

End Sub


Private Sub grid_faturamento_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   
   If SSTab1.Enabled Then
      If grid_faturamento.Col = 0 And LastCol > -1 Then
         grid_faturamento.Col = LastCol
         grid_faturamento.Row = LastRow
      End If
   End If

End Sub

Private Sub grid_faturamento_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim i As Integer
On Error Resume Next
NewRowBookmark = GetNewBookmark_fat()

' Loop over all the columns of the row, storing non-Null '
' values
    Dim newval As Variant
    For i% = 0 To RowBuf.ColumnCount - 1
        newval = RowBuf.Value(0, i%)
    
        If IsNull(newval) Then
' the RowBuf does not contain a value for this column.
' A default value should be set.  A convenient value
' is the default value for the column.
            newval = grid_faturamento.Columns(i%).DefaultValue
        End If
' Now store the new values.
        If Not StoreUserData_fat(NewRowBookmark, i%, newval) Then
' storage of the data has failed.  Delete the added
' row using a VBA function DeleteRow which takes a
' bookmark as an argument.  Also, fail the update by
' clearing the RowCount.
            DeleteRow_fat NewRowBookmark
            RowBuf.RowCount = 0 ' tell the grid the update failed
            Exit Sub            ' it failed, so exit the event
        End If
    Next i%


End Sub

Private Sub grid_faturamento_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    On Error Resume Next
    Dim bookm As Variant
    Dim rowsFetched As Integer
    Dim relpos As Integer
    Dim i As Integer
    Dim j As Integer
    
    bookm = StartLocation
    If ReadPriorRows Then
' the grid is requesting data in rows prior to
' StartLocation
        relpos = -1
    Else
' the grid is requesting data in rows after
' StartLocation
        relpos = 1
    End If
 
    rowsFetched = 0
    For i% = 0 To RowBuf.RowCount - 1
        If (RowBuf.RowCount - 1) <> 0 Then
            ' Get the bookmark of the next available row
            bookm = GetRelativeBookmark_fat(bookm, relpos)
        Else
            bookm = StartLocation
        End If
' If the next is BOF or EOF, then done
        If IsNull(bookm) Then Exit For
        For j% = 0 To RowBuf.ColumnCount - 1
            RowBuf.Value(i%, j%) = GetUserData_fat(bookm, j%)
        Next j%
' Set the bookmark for the row
        RowBuf.Bookmark(i%) = bookm
' Increment the count of fetched rows
        rowsFetched = rowsFetched + 1
    Next i%

' tell the grid how many rows were fetched
    RowBuf.RowCount = rowsFetched

End Sub


Private Sub grid_faturamento_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim i As Integer
On Error Resume Next
' Loop over all the columns of the row, storing non-Null
' values
    For i% = 0 To RowBuf.ColumnCount - 1
        If Not IsNull(RowBuf.Value(0, i%)) Then
            If Not StoreUserData_fat(WriteLocation, i%, RowBuf.Value(0, i%)) Then
 ' storage of the data has failed.  Fail the update
                RowBuf.RowCount = 0     ' tell the grid the update failed
                Exit Sub                 ' it failed, so exit the event
            End If
        End If
    Next i%


        Select Case grid_faturamento.Col
           Case Is = 1
               txt229 = acumula_grid_faturamento(grid_fatura, grid_faturamento.Col)
               txt229.Text = Format(txt229.Text, "#,##0.#0")
              grid_fatura(1, 3) = txt229.Text
           Case Is = 2
                If grid_faturamento.Row = 0 Then
                   txt236.Text = calcula_apur_imposto_Faturamento_Versus_Aliquota(grid_fatura(1, grid_faturamento.Row), grid_faturamento.Columns(grid_faturamento.Col).Text)
                   txt236.Text = Format(txt236.Text, "#,##0.#0")
                   grid_fatura(3, grid_faturamento.Row) = txt236.Text
                ElseIf grid_faturamento.Row = 1 Then
                   txt237.Text = calcula_apur_imposto_Faturamento_Versus_Aliquota(grid_fatura(1, grid_faturamento.Row), grid_faturamento.Columns(grid_faturamento.Col).Text)
                   txt237.Text = Format(txt237.Text, "#,##0.#0")
                   grid_fatura(3, grid_faturamento.Row) = txt237.Text
                ElseIf grid_faturamento.Row = 2 Then
                   txt238.Text = calcula_apur_imposto_Faturamento_Versus_Aliquota(grid_fatura(1, grid_faturamento.Row), grid_faturamento.Columns(grid_faturamento.Col).Text)
                   txt238.Text = Format(txt238.Text, "#,##0.#0")
                   grid_fatura(3, grid_faturamento.Row) = txt238.Text
                End If
        '   Case Is = 3
               txt242.Text = acumula_grid_faturamento(grid_fatura, 3)
               txt242.Text = Format(txt242.Text, "#,##0.#0")
              grid_fatura(3, 3) = txt242.Text
          
        End Select

  


End Sub


Private Sub grid_recolhimento_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

Dim ret As Integer

  If Not critica_valor_com_caracteres(grid_recolhimento.Columns(ColIndex).Text) Then
    grid_recolhimento.Columns(ColIndex).Text = "0,00"
  End If


   grid_recolhimento.Columns(ColIndex).Text = Format(grid_recolhimento.Columns(ColIndex).Text, "#,##0.#0")

End Sub

Private Sub grid_recolhimento_Change()
Navega_Ficha(1, 3) = 1
End Sub


Private Sub grid_recolhimento_GotFocus()
 
  grid_recolhimento.Col = 1
  grid_recolhimento.MarqueeStyle = 2
 
End Sub

Private Sub grid_recolhimento_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case g_condicao
   Case 1, 2, 5, 6
    
    If KeyCode = 40 And grid_recolhimento.Col < 2 And grid_recolhimento.Row >= 6 Then
       grid_recolhimento.Col = grid_recolhimento.Col + 1
       grid_recolhimento.Row = 0
       KeyCode = 0
       Exit Sub
    ElseIf KeyCode = 40 And grid_recolhimento.Col = 2 And grid_recolhimento.Row < 6 Then
          grid_recolhimento.Row = grid_recolhimento.Row + 1
          grid_recolhimento.Row = grid_recolhimento.Row - 1
          Exit Sub
    ElseIf KeyCode = 40 And grid_recolhimento.Col = 2 And grid_recolhimento.Row >= 6 Then
           grid_recolhimento.Col = 1
           grid_recolhimento.Row = 0
           KeyCode = 0
           Exit Sub
    End If
  
  If KeyCode = 39 Then
       KeyCode = 13
  End If
  
 If KeyCode = 13 Then
   desativa_label_pisca_pisca
   If grid_recolhimento.Row = 6 And grid_recolhimento.Col <> 2 Then
      grid_recolhimento.Col = grid_recolhimento.Col
      grid_recolhimento.Row = grid_recolhimento.Row - 1
      grid_recolhimento.Row = grid_recolhimento.Row + 1
   ElseIf grid_recolhimento.Row = 6 And grid_recolhimento.Col = 2 Then
      grid_recolhimento.Col = grid_recolhimento.Col
      grid_recolhimento.Row = grid_recolhimento.Row - 1
      grid_recolhimento.Row = grid_recolhimento.Row + 1
   ElseIf grid_recolhimento.Row > 6 And grid_recolhimento.Col > 2 Then
      grid_recolhimento.Col = 1
      grid_recolhimento.Row = 0
      KeyCode = 0
    End If
 End If

 If grid_recolhimento.Col = 2 Then
    If KeyCode = 9 Or KeyCode = 13 Or KeyCode = 40 Then
       If grid_recolhimento.Row >= 6 Then
          grid_recolhimento.Col = 0
          grid_recolhimento.Row = 0
          KeyCode = 0
          Exit Sub
       Else
          grid_recolhimento.Col = 0
          grid_recolhimento.Row = grid_recolhimento.Row + 1
       End If
     End If
  End If

Case 3, 4, 7
    If KeyCode = 40 And grid_recolhimento.Col < 2 And grid_recolhimento.Row >= 7 Then
       grid_recolhimento.Col = grid_recolhimento.Col + 1
       grid_recolhimento.Row = 0
       KeyCode = 0
       Exit Sub
    ElseIf KeyCode = 40 And grid_recolhimento.Col = 2 And grid_recolhimento.Row < 7 Then
          grid_recolhimento.Row = grid_recolhimento.Row + 1
          grid_recolhimento.Row = grid_recolhimento.Row - 1
          Exit Sub
    ElseIf KeyCode = 40 And grid_recolhimento.Col = 2 And grid_recolhimento.Row >= 7 Then
           grid_recolhimento.Col = 1
           grid_recolhimento.Row = 0
           KeyCode = 0
           Exit Sub
    End If
  
  If KeyCode = 39 Then
       KeyCode = 13
  End If
  
 If KeyCode = 13 Then
   desativa_label_pisca_pisca
   If grid_recolhimento.Row = 7 And grid_recolhimento.Col <> 2 Then
      grid_recolhimento.Col = grid_recolhimento.Col
      grid_recolhimento.Row = grid_recolhimento.Row - 1
      grid_recolhimento.Row = grid_recolhimento.Row + 1
   ElseIf grid_recolhimento.Row = 7 And grid_recolhimento.Col = 2 Then
      grid_recolhimento.Col = grid_recolhimento.Col
      grid_recolhimento.Row = grid_recolhimento.Row - 1
      grid_recolhimento.Row = grid_recolhimento.Row + 1
   ElseIf grid_recolhimento.Row > 7 And grid_recolhimento.Col > 2 Then
      grid_recolhimento.Col = 1
      grid_recolhimento.Row = 0
      KeyCode = 0
    End If
 End If

  If grid_recolhimento.Col = 2 Then
    If KeyCode = 9 Or KeyCode = 13 Or KeyCode = 40 Then
       If grid_recolhimento.Row >= 7 Then
          grid_recolhimento.Col = 0
          grid_recolhimento.Row = 0
          KeyCode = 0
          Exit Sub
       Else
          grid_recolhimento.Col = 0
          grid_recolhimento.Row = grid_recolhimento.Row + 1
       End If
     End If
  End If

End Select
End Sub


Private Sub grid_recolhimento_KeyPress(KeyAscii As Integer)
 
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If
   
End Sub


Private Sub grid_recolhimento_LostFocus()
      
   grid_recolhimento.MarqueeStyle = 5

End Sub


Private Sub grid_recolhimento_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim i As Integer
On Error Resume Next
NewRowBookmark = GetNewBookmark_rec()

' Loop over all the columns of the row, storing non-Null '
' values
    Dim newval As Variant
    For i% = 0 To RowBuf.ColumnCount - 1
        newval = RowBuf.Value(0, i%)
    
        If IsNull(newval) Then
' the RowBuf does not contain a value for this column.
' A default value should be set.  A convenient value
' is the default value for the column.
            newval = grid_recolhimento.Columns(i%).DefaultValue
        End If
' Now store the new values.
        If Not StoreUserData_rec(NewRowBookmark, i%, newval) Then
' storage of the data has failed.  Delete the added
' row using a VBA function DeleteRow which takes a
' bookmark as an argument.  Also, fail the update by
' clearing the RowCount.
            DeleteRow_rec NewRowBookmark
            RowBuf.RowCount = 0 ' tell the grid the update failed
            Exit Sub            ' it failed, so exit the event
        End If
    Next i%


End Sub


Private Sub grid_recolhimento_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    On Error Resume Next
    Dim bookm As Variant
    Dim rowsFetched As Integer
    Dim relpos As Integer
    Dim i As Integer
    Dim j As Integer
    
    bookm = StartLocation
    If ReadPriorRows Then
' the grid is requesting data in rows prior to
' StartLocation
        relpos = -1
    Else
' the grid is requesting data in rows after
' StartLocation
        relpos = 1
    End If
 
    rowsFetched = 0
    For i% = 0 To RowBuf.RowCount - 1
        If (RowBuf.RowCount - 1) <> 0 Then
            ' Get the bookmark of the next available row
            bookm = GetRelativeBookmark_rec(bookm, relpos)
        Else
            bookm = StartLocation
        End If
' If the next is BOF or EOF, then done
        If IsNull(bookm) Then Exit For
        For j% = 0 To RowBuf.ColumnCount - 1
            RowBuf.Value(i%, j%) = GetUserData_rec(bookm, j%)
        Next j%
' Set the bookmark for the row
        RowBuf.Bookmark(i%) = bookm
' Increment the count of fetched rows
        rowsFetched = rowsFetched + 1
    Next i%

' tell the grid how many rows were fetched
    RowBuf.RowCount = rowsFetched


End Sub


Private Sub grid_recolhimento_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim i As Integer
On Error Resume Next
' Loop over all the columns of the row, storing non-Null
' values
    For i% = 0 To RowBuf.ColumnCount - 1
        If Not IsNull(RowBuf.Value(0, i%)) Then
            If Not StoreUserData_rec(WriteLocation, i%, RowBuf.Value(0, i%)) Then
 ' storage of the data has failed.  Fail the update
                RowBuf.RowCount = 0     ' tell the grid the update failed
                Exit Sub                 ' it failed, so exit the event
            End If
        End If
    Next i%

'       txt314 = acumula_grid_recolhimento(grid_recolim)
 '      txt314.Text = Format(txt314.Text, "#,##0.#0")
       Select Case g_condicao
          Case Is = 3, 4, 7
'         grid_recolim(1, 9) = txt314.Text
        Case Is = 1, 2, 5, 6
'         grid_recolim(1, 8) = txt314.Text
      End Select
      
  
      
'txt314.Refresh

End Sub


Private Sub msk_aliquota_frm4_GotFocus()

  msk_aliquota_frm4.SelStart = 0
   msk_aliquota_frm4.SelLength = Len(msk_aliquota_frm4)

End Sub


Private Sub msk_aliquota_frm4_KeyDown(KeyCode As Integer, Shift As Integer)

   '   If KeyCode = 9 Or KeyCode = 13 Then
     '      msk_imposto_frm4.SetFocus
    '  End If

End Sub


Private Sub msk_aliquota_frm4_KeyPress(KeyAscii As Integer)

'Call PulaCampo(KeyAscii)

    '  If KeyAscii = 13 Then
    '      msk_imposto_frm4.SetFocus
    '  End If
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub SSTab1_LostFocus()
'desativa_label_pisca_pisca
End Sub

Private Sub Timer1_Timer()

If LBL_APUR_PISCA.Visible = True Then
    LBL_APUR_PISCA.Caption = Mid(Frase3, 1, Pos3)
    Pos3 = Pos3 + 1
    If Pos3 = Len(Frase3) Then Pos3 = 1
End If
End Sub







































Private Sub txt207_LostFocus()
If Not critica_valor_com_caracteres(txt207.Text) Then
    txt207.Text = "0,00"
End If
txt207.Text = Format(txt207.Text, "#,##0.#0")
grid_ApurICMS(1, 6) = txt207.Text
Call Calcula_206
End Sub

Private Sub txt209_LostFocus()
If Not critica_valor_com_caracteres(txt209.Text) Then
    txt209.Text = "0,00"
End If
txt209.Text = Format(txt209.Text, "#,##0.#0")
grid_ApurICMS(1, 8) = txt209.Text
Call Calcula_206

End Sub

Private Sub txt210_LostFocus()
If Not critica_valor_com_caracteres(txt210.Text) Then
    txt210.Text = "0,00"
End If
txt210.Text = Format(txt210.Text, "#,##0.#0")
grid_ApurICMS(1, 9) = txt210.Text
Call Calcula_206

End Sub

Private Sub txt211_LostFocus()
If Not critica_valor_com_caracteres(txt211.Text) Then
    txt211.Text = "0,00"
End If
txt211.Text = Format(txt211.Text, "#,##0.#0")
grid_ApurICMS(1, 10) = txt211.Text
Call Calcula_206

End Sub

Private Sub txt213_LostFocus()
If Not critica_valor_com_caracteres(txt213.Text) Then
    txt213.Text = "0,00"
End If
txt213.Text = Format(txt213.Text, "#,##0.#0")
Call Calcula_206
calcula_dados_ApuracaoImposto_chaim
grid_ApurICMS(1, 12) = txt213.Text

End Sub

Private Sub txt214_GotFocus()
txt214.SelStart = 0
txt214.SelLength = Len(txt214)

End Sub


Private Sub txt214_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
   KeyAscii = 0
ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
   KeyAscii = Asc(",")
End If
End Sub


Private Sub txt214_LostFocus()
If Not critica_valor_com_caracteres(txt214.Text) Then
    txt214.Text = "0,00"
End If
txt214.Text = Format(txt214.Text, "#,##0.#0")
calcula_dados_ApuracaoImposto_chaim
grid_ApurICMS(1, 13) = txt214.Text

End Sub

Private Sub txt215_GotFocus()
txt215.SelStart = 0
txt215.SelLength = Len(txt215)
End Sub


Private Sub txt215_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
   KeyAscii = 0
ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
   KeyAscii = Asc(",")
End If
End Sub


Private Sub txt215_LostFocus()

If Not critica_valor_com_caracteres(txt215.Text) Then
    txt215.Text = "0,00"
End If

txt215.Text = Format(txt215.Text, "#,##0.#0")
grid_ApurICMS(1, 14) = txt215.Text

End Sub

Private Sub txt216_GotFocus()
txt216.SelStart = 0
txt216.SelLength = Len(txt216)
End Sub


Private Sub txt216_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
   KeyAscii = 0
ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
   KeyAscii = Asc(",")
End If
End Sub


Private Sub txt216_LostFocus()
If Not critica_valor_com_caracteres(txt216.Text) Then
    txt216.Text = "0,00"
  End If

If txt216.Text <> "" And IsNumeric(txt216.Text) Then
   txt216.Text = Format(txt216.Text, "#,##0.#0")
   If CDbl(txt216.Text) > CDbl(txt215.Text) Then
      MsgBox "Deduções somente são permitidas até o limite do Saldo Devedor", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
      txt216.SetFocus
      Exit Sub
   End If
End If

txt216.Text = Format(txt216.Text, "#,##0.#0")
grid_ApurICMS(1, 15) = txt216.Text
calcula_dados_ApuracaoImposto_chaim
Navega_Ficha(1, 2) = 1

End Sub

Private Sub txt217_GotFocus()
txt217.SelStart = 0
txt217.SelLength = Len(txt217)
End Sub


Private Sub txt217_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
   KeyAscii = 0
ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
   KeyAscii = Asc(",")
End If
End Sub


Private Sub txt217_LostFocus()
If Not critica_valor_com_caracteres(txt217.Text) Then
    txt217.Text = "0,00"
End If
txt217.Text = Format(txt217.Text, "#,##0.#0")
grid_ApurICMS(1, 16) = txt217.Text
Navega_Ficha(1, 2) = 1

End Sub

Private Sub txt218_GotFocus()
txt218.SelStart = 0
txt218.SelLength = Len(txt218)
End Sub


Private Sub txt218_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
   KeyAscii = 0
ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
   KeyAscii = Asc(",")
End If
End Sub


Private Sub txt218_LostFocus()
If Not critica_valor_com_caracteres(txt218.Text) Then
    txt218.Text = "0,00"
  End If

txt218.Text = Format(txt218.Text, "#,##0.#0")
grid_ApurICMS(1, 17) = txt218.Text

End Sub

Private Sub txt301_GotFocus()

   txt301.SelStart = 0
   txt301.SelLength = Len(txt301)

End Sub


Private Sub txt301_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     txt302.SetFocus
End If
  
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
        KeyAscii = 0
 ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
        KeyAscii = Asc(",")
 End If
 
txt301 = tirapictu(txt301, ".")

If Len(Trim(txt301)) > 15 Then
   If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
  End If
  
End Sub


Private Sub txt301_LostFocus()

   If Not critica_valor_com_caracteres(txt301.Text) Then
    txt301.Text = "0,00"
  End If

txt301.Text = Format(txt301.Text, "#,##0.#0")
Navega_Ficha(1, 3) = 1
acumula_total_303
End Sub


Private Sub txt300_GotFocus()


   txt300.SelStart = 0
   txt300.SelLength = Len(txt300)

End Sub


Private Sub txt300_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     txt301.SetFocus
  End If
  
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
        KeyAscii = 0
 ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
        KeyAscii = Asc(",")
 End If
txt300 = tirapictu(txt300, ".")


 If Len(Trim(txt300)) > 15 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If
  
End Sub


Private Sub txt300_LostFocus()

   If Not critica_valor_com_caracteres(txt300.Text) Then
    txt300.Text = "0,00"
  End If


txt300.Text = Format(txt300.Text, "#,##0.#0")
acumula_total_303
Navega_Ficha(1, 3) = 1

End Sub


Private Sub txt302_GotFocus()

   txt302.SelStart = 0
   txt302.SelLength = Len(txt302)

End Sub


Private Sub txt302_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
      grid_recolhimento.Col = 1
      grid_recolhimento.Row = 0
      grid_recolhimento.MarqueeStyle = 2
      grid_recolhimento.SetFocus
   End If
End Sub


Private Sub txt302_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     grid_recolhimento.SetFocus
  End If
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
        KeyAscii = 0
 ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
        KeyAscii = Asc(",")
 End If

  
txt302 = tirapictu(txt302, ".")

If Len(Trim(txt302)) > 15 Then
   If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
  End If
End Sub


Private Sub txt302_LostFocus()

   If Not critica_valor_com_caracteres(txt302.Text) Then
    txt302.Text = "0,00"
  End If


txt302.Text = Format(txt302.Text, "#,##0.#0")
Navega_Ficha(1, 3) = 1
acumula_total_303
End Sub


Private Sub msk_fat_mensal_frm4_GotFocus()

  
   msk_fat_mensal_frm4.SelStart = 0
   msk_fat_mensal_frm4.SelLength = Len(msk_fat_mensal_frm4)

End Sub


Private Sub msk_fat_mensal_frm4_KeyPress(KeyAscii As Integer)

Call PulaCampo(KeyAscii)
  'If KeyAscii = 13 Then
  '       msk_aliquota_frm4.SetFocus
  'End If
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

msk_fat_mensal_frm4 = tirapictu(msk_fat_mensal_frm4, ".")
If Len(Trim(msk_fat_mensal_frm4)) > 15 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If


End Sub

Private Sub msk_fat_mensal_frm4_LostFocus()
      
   If Not critica_valor_com_caracteres(msk_fat_mensal_frm4.Text) Then
    msk_fat_mensal_frm4.Text = "0,00"
  End If
      
      
    msk_fat_mensal_frm4.Text = Format(msk_fat_mensal_frm4.Text, "#,##0.#0")
  Navega_Ficha(1, 2) = 1
  
 msk_imposto_frm4 = CALCULA_IMPOSTO_EPP(msk_fat_mensal_frm4, msk_aliquota_frm4)
    msk_imposto_frm4.Text = Format(msk_imposto_frm4.Text, "#,##0.#0")
   desativa_label_pisca_pisca
End Sub

Private Sub msk_imposto_frm4_GotFocus()
   
   msk_imposto_frm4.SelStart = 0
   msk_imposto_frm4.SelLength = Len(msk_imposto_frm4)

End Sub


Private Sub msk_imposto_frm4_KeyDown(KeyCode As Integer, Shift As Integer)
      
      If KeyCode = 9 Or KeyCode = 13 Then
         msk_fat_mensal_frm4.SetFocus
      End If

End Sub


Private Sub msk_imposto_frm4_KeyPress(KeyAscii As Integer)

'Call PulaCampo(KeyAscii)
If KeyAscii = 13 Then
  If g_condicao = 4 Then
    txt201.SetFocus
  Else
   msk_fat_mensal_frm4.SetFocus
  End If
 End If
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

'  If Len(Trim(msk_imposto_frm4)) > 15 Then
'      SendKeys "{BackSpace}"
'      Exit Sub
'  End If
  
End Sub

Private Sub msk_imposto_frm4_LostFocus()
      
   If Not critica_valor_com_caracteres(msk_imposto_frm4.Text) Then
    msk_imposto_frm4.Text = "0,00"
  End If
      
      
      
      msk_imposto_frm4.Text = Format(msk_imposto_frm4.Text, "#,##0.#0")
 Navega_Ficha(1, 2) = 1
  
End Sub


Private Sub msk_inscricao_GotFocus()

msk_inscricao.SelStart = 0
msk_inscricao.SelLength = Len(msk_inscricao)

End Sub


Private Sub msk_inscricao_KeyPress(KeyAscii As Integer)
Dim query As String
Dim restauracontr As Recordset
Dim MyInscricao As String

MyInscricao = msk_inscricao.Text
MyInscricao = tirapictu(MyInscricao, ".")
MyInscricao = tirapictu(MyInscricao, "-")

If Len(Trim(MyInscricao)) = 0 Then Exit Sub
 If KeyAscii = 9 Or KeyAscii = 13 Then 'Se teclar enter
    If Len(Trim(MyInscricao)) < 9 Then
       MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       msk_inscricao = ""
       msk_inscricao.SetFocus
       Exit Sub
    End If
    query = "select * from tab_cad_contribuintes where tcc_nu_inscricao = '" & Trim(Str(MyInscricao)) & "'"
    Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
    If restauracontr.EOF And restauracontr.BOF Then
       MsgBox "Número de Inscrição não Cadastrado.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       restauracontr.Close
       msk_inscricao.SetFocus
       Exit Sub
    End If

    txt_periodo_inicial.SetFocus
 End If
 
  If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep
  
 If Len(Trim(msk_inscricao)) > 9 Then
   '  SendKeys "{BackSpace}"
      Exit Sub
  End If
End Sub


Private Sub msk_inscricao_LostFocus()
Dim MyInscricao As String

If Len(Trim(msk_inscricao)) = 0 Then Exit Sub
If TypeOf Me.ActiveControl Is SSCommand Then
   If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "diretorio" Or Me.ActiveControl.Tag = "buscar" Then
      Exit Sub
   End If
End If

MyInscricao = msk_inscricao.Text
MyInscricao = tirapictu(MyInscricao, ".")
MyInscricao = tirapictu(MyInscricao, "-")
msk_inscricao.Text = MyInscricao

   If Len(Trim(MyInscricao)) < 9 Then
       MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       msk_inscricao = ""
       erro_inscricao = True
       msk_inscricao.SetFocus
       Exit Sub
  End If
    If GenDV(CStr(Mid(MyInscricao, 1, 8)), BASE9) <> Val(Mid(MyInscricao, 9, 1)) Then
       MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       msk_inscricao = ""
       erro_inscricao = True
       msk_inscricao.SetFocus
       Exit Sub
    End If
    query = "select * from tab_cad_contribuintes where tcc_nu_inscricao = '" & Trim(Str(MyInscricao)) & "'"
    Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
    If restauracontr.EOF And restauracontr.BOF Then
       MsgBox "Número de Inscrição não Cadastrado.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       restauracontr.Close
       erro_inscricao = True
       msk_inscricao.SetFocus
       Exit Sub
    End If
    txt_razao = restauracontr.Fields("tcc_razao_social")
    g_num_cgc_cpf = restauracontr.Fields("tcc_cad_geral_contrib")

    g_ano_inicio_operacao = Val(Mid(restauracontr.Fields("tcc_da_inicio_op") & "", 5, 4))
     PER_INICIO = Val(Mid(restauracontr.Fields("tcc_da_inicio_op") & "", 5, 4) & Mid(restauracontr.Fields("tcc_da_inicio_op") & "", 3, 2)) & ""
    If g_condicao = 0 Then
       g_condicao = restauracontr.Fields("tcc_tipo_contribuinte")
        Select Case restauracontr.Fields("tcc_tipo_contribuinte")
           Case 1
                txtCondContr = "Normal"
           Case 3, 4, 5, 6, 7
                txtCondContr = "Pequeno Porte"
           Case 2
                txtCondContr = "Microempresa"
        End Select
     Else
    Select Case g_condicao
         Case 1
            txtCondContr = "Normal"
         Case 3, 4, 5, 6, 7
            txtCondContr = "Pequeno Porte "
         Case 2
            txtCondContr = "Microempresa"
    End Select
     End If
    restauracontr.Close
    TestaCampoPreench
    erro_inscricao = False

End Sub


Private Sub txt303_GotFocus()

   txt303.SelStart = 0
   txt303.SelLength = Len(txt303)

End Sub


Private Sub txt303_KeyPress(KeyAscii As Integer)

 If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
        KeyAscii = 0
 ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
        KeyAscii = Asc(",")
 End If

  
  If Len(Trim(txt303)) > 15 Then
   '   SendKeys "{BackSpace}"
         KeyAscii = 0

      Exit Sub
  End If


End Sub


Private Sub txt303_LostFocus()


   If Not critica_valor_com_caracteres(txt303.Text) Then
    txt303.Text = "0,00"
  End If


   txt303.Text = Format(txt303.Text, "#,##0.#0")
  Navega_Ficha(1, 3) = 1

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

On Error Resume Next
DoEvents
  If TypeOf Me.ActiveControl Is SSCommand Then
       If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "salvar" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "diretorio" Or Me.ActiveControl.Tag = "excluir" Or Me.ActiveControl.Tag = "buscar" Then
          Exit Sub
       End If
    End If
  desativa_label_pisca_pisca
  'refresh_campos_grids
     Screen.MousePointer = 11
     Select Case TAB_ANTERIOR
       Case 0
           If Not CRITICA_DEM_ENTRADAS() Then
              Navega_Ficha(1, 0) = 1
              'TAB_ANTERIOR = 9
              'SSTab1.Tab = 0
            End If
        Case 1
           If Not CRITICA_DEM_saidas() Then
              Navega_Ficha(1, 1) = 1
              'TAB_ANTERIOR = 9
              'SSTab1.Tab = 1
           End If
        Case 2
           If Not critica_apuracao_icms() Then
              Navega_Ficha(1, 2) = 1
                'TAB_ANTERIOR = 9
                'SSTab1.Tab = 2
           End If
        Case 3
           If Not critica_recolhimento Then
               Navega_Ficha(1, 3) = 1
               Screen.MousePointer = 0
               'Exit Sub
           End If
        Case 4
           If Not critica_detalhes() Then
               Navega_Ficha(1, 4) = 1
              'TAB_ANTERIOR = 9
              'SSTab1.Tab = 4
            End If
     End Select
     Screen.MousePointer = 0
     
   monta_frames
 
  If SSTab1.Tab = 0 Then '"Entradas"
     TAB_ANTERIOR = 0
     teve_alteracao_declaracao = True
     Navega_Ficha(1, 0) = 1
     Select Case g_condicao
      Case 2, 5, 6
           
            fra_EntaNormal.Top = 7000
            fra_EntraMicroEPP.Top = 660
            fra_EntraMicroEPP.Left = 75    'inclusao em 21/2/98 em funcao do 16 bits
            If SSTab1.Enabled = True And grid_Entra_MicroEPP.Enabled = True Then
               grid_Entra_MicroEPP.Refresh
               fra_EntraMicroEPP.Enabled = True
               fra_EntraMicroEPP.Visible = True
               grid_Entra_MicroEPP.MarqueeStyle = 2
               grid_Entra_MicroEPP.SetFocus
            End If
   Case 1, 3, 4, 7
       ' fra_EntraMicroEPP.Top = 7000
        fra_EntaNormal.Top = 650
        
        If SSTab1.Enabled = True And GRID_DEM_ENTRADAS.Enabled = True Then

           fra_EntaNormal.Enabled = True
           fra_EntaNormal.Visible = True
           GRID_DEM_ENTRADAS.Refresh
           GRID_DEM_ENTRADAS.MarqueeStyle = 2
           GRID_DEM_ENTRADAS.SetFocus
        End If
    End Select
    If g_condicao = 3 Or g_condicao = 4 Then
       'Frase = " Informe a seguir somente o detalhamento das operações com mercadorias cuja alíquota interna no Maranhão seja igual a 25% por cento e proceda a apuração normal."
        Frase = " Declare a seguir todas as operações com mercadorias " _
              & "cuja alíquota interna no Maranhão seja diferenciada (<17% e >17% =25%), " _
              & "registrando e apurando conforme instruções constantes " _
              & "na ajuda, quadro Critérios para Declaração de Contribuinte (Regime EPP)."
       ativa_label_pisca_apur
    End If

 ElseIf SSTab1.Tab = 1 Then '"Saídas"
    TAB_ANTERIOR = 1
    teve_alteracao_declaracao = True
    Navega_Ficha(1, 1) = 1

    grid_dem_saidas.Refresh
    grid_dem_saidas.MarqueeStyle = 2
    grid_dem_saidas.SetFocus
    If g_condicao = 3 Or g_condicao = 4 Then
       'Frase = " Informe a seguir somente o detalhamento das operações com mercadorias cuja alíquota interna no Maranhão seja igual a 25% por cento e proceda a apuração normal."
        Frase = " Declare a seguir todas as operações com mercadorias " _
              & "cuja alíquota interna no Maranhão seja diferenciada (<17% e >17% =25%), " _
              & "registrando e apurando conforme instruções constantes " _
              & "na ajuda, quadro Critérios para Declaração de Contribuinte (Regime EPP)."
       ativa_label_pisca_apur
    End If
    
 ElseIf SSTab1.Tab = 2 Then '"Apuracao do ICMS"
    desativa_label_pisca_pisca
    teve_alteracao_declaracao = True
    Navega_Ficha(1, 2) = 1
    TAB_ANTERIOR = 2
    transporta_valores_Apur_ICMS
    acumula_total_204
    acumula_total_214
    calcula_dados_ApuracaoImposto_chaim
    Select Case g_condicao
        Case Is = 1, 7             'Normal
            Frame4.Visible = False
            frm_grid_fatu.Visible = False         'Faturamento - Tela #4 e 5
            'frm_tela3.Enabled = True
            Frame_tela3.Visible = True
            fra_Apur_ICMS.Visible = True
            frm_tela3.Visible = True            'ICMS - Tela #3
          '  grdApurICMS.Refresh
            txt_fatura_normal.SetFocus
            Exit Sub
         Case Is = 3, 4, 5, 6          'EPP
             If g_condicao = 3 Then
               msk_aliquota_frm4.Text = "1,5"
             ElseIf g_condicao = 4 Then
               msk_aliquota_frm4.Text = "2,0"
             ElseIf g_condicao = 5 Then
               msk_aliquota_frm4.Text = "1,5"
             ElseIf g_condicao = 6 Then
               msk_aliquota_frm4.Text = "2,0"
             End If
            If g_aliquota_Maior17 = 1 Then    'aliquota maior que 17%
               Frame4.Visible = True
               Frame_tela3.Visible = False
               frm_grid_fatu.Visible = False
               frm_tela3.Enabled = True           'ICMS - Tela #3
               frm_tela3.Visible = True           'ICMS - Tela #3
               msk_fat_mensal_frm4.SetFocus
            Else
               frm_tela3.Visible = False           'ICMS - Tela #3
               Frame_tela3.Visible = False
                frm_grid_fatu.Visible = False
               Frame4.Visible = True
               Frame4.Enabled = True
               msk_fat_mensal_frm4.SetFocus
             End If
        Case Is = 2                'Microempresa
            frm_tela3.Visible = False           'ICMS - Tela #3
            Frame_tela3.Visible = False
            Frame4.Visible = False              'Faturamento - Tela #4 e 5
            frm_grid_fatu.Enabled = True
            frm_grid_fatu.Visible = True
            grid_faturamento.Refresh
            grid_faturamento.MarqueeStyle = 2
            grid_faturamento.SetFocus

            Exit Sub
    End Select
    If g_condicao = 3 Or g_condicao = 4 Then
       Frase = " Informe se houver, no quadro de faturamento do mês somente o total das operações com mercadorias cuja alíquota interna no Maranhão seja inferior a 25%."
       ativa_label_pisca_apur
    End If

 ElseIf SSTab1.Tab = 3 Then     '"Recolhimentos no Periodo"
   TAB_ANTERIOR = 3
   teve_alteracao_declaracao = True
     Navega_Ficha(1, 3) = 1
   fra_Recolhimento.Enabled = True
   fra_Recolhimento.Visible = True
   'cargar_grid_recolhimento     'inclui em 21/3/98
   Select Case g_condicao
     Case Is = 1, 3, 4, 5, 6, 7
         Frame6.Visible = True
         Frame7.Left = 180      'inclui em 21/2/98
     Case Is = 2
         Frame6.Visible = False
         Frame7.Left = 2100     'inclui em 21/2/98
   End Select
   grid_recolhimento.Refresh
   Frame7.Enabled = True
   Frame7.Visible = True
   grid_recolhimento.MarqueeStyle = 2
   grid_recolhimento.SetFocus
   Select Case g_condicao
     Case Is = 1, 2, 5, 6
 
          
     Case Is = 3, 4, 7
 
   End Select
   
ElseIf SSTab1.Tab = 4 Then '"Dados do Exercicio anterior
   teve_alteracao_declaracao = True
  TAB_ANTERIOR = 4
   Navega_Ficha(1, 4) = 1
   cargar_grid_DetDesp
   grdDetDesp.SetFocus

End If

End Sub

Private Sub SSTab1_GotFocus()

desativa_label_pisca_pisca
 monta_frames

 If SSTab1.Tab = 0 Then '"Entradas"
    TAB_ANTERIOR = 0
    Navega_Ficha(1, 0) = 1
   If g_condicao = 3 Or g_condicao = 4 Then
       'Frase = " Informe a seguir somente o detalhamento das operações com mercadorias cuja alíquota interna no Maranhão seja igual a 25% por cento e proceda a apuração normal."
        Frase = " Declare a seguir todas as operações com mercadorias " _
              & "cuja alíquota interna no Maranhão seja diferenciada (<17% e >17% =25%), " _
              & "registrando e apurando conforme instruções constantes " _
              & "na ajuda, quadro Critérios para Declaração de Contribuinte (Regime EPP)."
       ativa_label_pisca_apur
    End If
    
    Select Case g_condicao
        Case Is = 1, 3, 4, 7
          If GRID_DEM_ENTRADAS.Enabled Then
             GRID_DEM_ENTRADAS.SetFocus
          End If
          Exit Sub
       Case Is = 2, 5, 6
          GRID_DEM_ENTRADAS.Col = 2
          GRID_DEM_ENTRADAS.Enabled = False
    End Select

 ElseIf SSTab1.Tab = 1 Then '"Saídas"
    If g_condicao = 3 Or g_condicao = 4 Then
      ''  Frase = " Informe a seguir somente o detalhamento das operações com mercadorias cuja alíquota interna no Maranhão seja igual a 25% por cento e proceda a apuração normal."
       Frase = " Declare a seguir todas as operações com mercadorias " _
              & "cuja alíquota interna no Maranhão seja diferenciada (<17% e >17% =25%), " _
              & "registrando e apurando conforme instruções constantes " _
              & "na ajuda, quadro Critérios para Declaração de Contribuinte (Regime EPP)."
       ativa_label_pisca_apur
    End If
 
ElseIf SSTab1.Tab = 2 Then '"Apuracao do ICMS"
    Select Case g_condicao
       Case Is = 1, 7    'Normal e epp c/faturamento maior 240.000,00
           ' desativa_label_pisca_pisca
            Frame4.Visible = False
            frm_tela3.Visible = True
            'frm_tela3.Enabled = True
        Case Is = 3, 4, 5, 6
            'desativa_label_pisca_pisca
            If g_aliquota_Maior17 = 1 Then    'aliquota maior que 17%
               Frame4.Visible = True
            Else
               Frame4.Visible = True
            End If
  
        Case Is = 2
            'desativa_label_pisca_pisca
            frm_grid_fatu.Visible = True
            frm_grid_fatu.Enabled = True
            Frame4.Visible = False
    End Select
   If g_condicao = 3 Or g_condicao = 4 Then
       Frase = " Informe se houver, no quadro de faturamento do mês somente o total das operações com mercadorias cuja alíquota interna no Maranhão seja inferior a 25%."
       ativa_label_pisca_apur
    End If

 ElseIf SSTab1.Tab = 3 Then '"Recolhimentos no Periodo"
   fra_Recolhimento.Enabled = True
   fra_Recolhimento.Visible = True
   Select Case g_condicao
     Case Is = 1, 3, 4, 5, 6, 7
         Frame6.Visible = True
        Frame7.Left = 180
     Case Is = 2
         Frame6.Visible = False
         Frame7.Left = 2100
   End Select
   fra_Recolhimento.Left = 300
   
             'inclui em 21/2/ 98
   
   grid_recolhimento.MarqueeStyle = 2
   grid_recolhimento.SetFocus
ElseIf SSTab1.Tab = 4 Then '"Dados do Exercicio anterior
   '   txt400.SetFocus
    grdDetDesp.SetFocus
End If

End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub

Private Sub txt_fatura_normal_GotFocus()

    txt_fatura_normal.SelStart = 0
    txt_fatura_normal.SelLength = Len(txt_fatura_normal)
    
End Sub


Private Sub txt_fatura_normal_KeyPress(KeyAscii As Integer)

Call PulaCampo(KeyAscii)

                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt_fatura_normal_LostFocus()

If Len(txt_fatura_normal) = 0 Then Exit Sub

If Not critica_valor_com_caracteres(txt_fatura_normal.Text) Then
    txt_fatura_normal.Text = "0,00"
  End If

 txt_fatura_normal.Text = Format(txt_fatura_normal.Text, "#,##0.#0")

Navega_Ficha(1, 2) = 1
End Sub

Private Sub txt_periodo_inicial_GotFocus()
    
    txt_periodo_inicial.SelStart = 0
    txt_periodo_inicial.SelLength = Len(txt_periodo_inicial)

End Sub


Private Sub txt_periodo_inicial_KeyPress(KeyAscii As Integer)
   Dim per_sis As Double
Dim per_prog As Double

If KeyAscii = 13 Or KeyAscii = 9 Then 'Se teclar enter

  If Len(Trim(msk_inscricao)) <> 0 And Len(Trim(txt_periodo_inicial)) = 0 Then
      MsgBox "Período inválido.", 48, NOME_SISTEMA_DEFAULT
      txt_periodo_inicial.Text = ""
      txt_periodo_inicial.SetFocus
      Exit Sub
   End If
   
   reta = crit_periodo(txt_periodo_inicial)
   If reta <> 0 Then
      MsgBox "Período inválido.", 48, NOME_SISTEMA_DEFAULT
      txt_periodo_inicial.Text = ""
      txt_periodo_inicial.SetFocus
      Exit Sub
   End If
   per_prog = Val(Mid$(txt_periodo_inicial, 3, 4) & Mid$(txt_periodo_inicial, 1, 2))
   per_sis = Val(Year(Date)) & Mid$(Date, 4, 2)
   If per_sis < per_prog Then
      MsgBox "Período não pode ser maior que a data atual.", 48, NOME_SISTEMA_DEFAULT
      txt_periodo_inicial.Text = ""
      txt_periodo_inicial.SetFocus
      Exit Sub
   End If
   If txtCondContr = "Microempresa" Then
      If Mid$(Trim(txt_periodo_inicial), 1, 2) <> "03" And Mid$(Trim(txt_periodo_inicial), 1, 2) <> "06" And Mid$(Trim(txt_periodo_inicial), 1, 2) <> "09" And Mid$(Trim(txt_periodo_inicial), 1, 2) <> "12" Then
         MsgBox "Período incompatível com a condição da declaração.", 48, NOME_SISTEMA_DEFAULT
         txt_periodo_inicial.Text = ""
         txt_periodo_inicial.SetFocus
         Exit Sub
      End If
   End If
   If Len(Trim(txt_periodo_inicial)) <> 0 And Val(Mid$(Trim(txt_periodo_inicial), 3, 6)) < 1991 Then
      MsgBox "Ano do Período não pode ser anterior a 1991.", 48, NOME_SISTEMA_DEFAULT
      txt_periodo_inicial.Text = ""
      txt_periodo_inicial.SetFocus
      Exit Sub
   End If
   cmb_tipo.SetFocus
   Exit Sub
End If
   
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep


   
End Sub


Private Sub txt_periodo_inicial_LostFocus()

Dim per_sis As Double
Dim per_prog As Double
If erro_inscricao = True Then Exit Sub

 If TypeOf Me.ActiveControl Is MaskEdBox Then
      If Me.ActiveControl.Tag = "inscricao" Then
         Exit Sub
      End If
 End If
   
   If TypeOf Me.ActiveControl Is SSCommand Then
      If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "iniciar" Or Me.ActiveControl.Tag = "diretorio" Or Me.ActiveControl.Tag = "buscar" Then
         Exit Sub
      End If
   End If
   
   If Len(Trim(msk_inscricao)) <> 0 And Len(Trim(txt_periodo_inicial)) = 0 Then
      MsgBox "Período inválido.", 48, NOME_SISTEMA_DEFAULT
      txt_periodo_inicial.SetFocus
      Exit Sub
   End If
   
   reta = crit_periodo(txt_periodo_inicial)
   If reta <> 0 Then
      MsgBox "Período inválido.", 48, NOME_SISTEMA_DEFAULT
      txt_periodo_inicial.SetFocus
      Exit Sub
   End If
   

   per_prog = Val(Mid$(txt_periodo_inicial, 3, 4) & Mid$(txt_periodo_inicial, 1, 2))
    per_sis = Val(Year(Date)) & Mid$(Date, 4, 2)
    
   If per_prog < PER_INICIO Then
      MsgBox "Período não pode ser menor que a data de inicio de operações.", 48, NOME_SISTEMA_DEFAULT
      txt_periodo_inicial.Text = ""
      txt_periodo_inicial.SetFocus
      Exit Sub
   End If
   
   
'   per_sis = Val(Mid$(Date, 7, 4) & Mid$(Date, 4, 2))
'   per_prog = Val(Mid$(txt_periodo_inicial, 3, 4) & Mid$(txt_periodo_inicial, 1, 2))
   
   If per_sis < per_prog Then
      MsgBox "Período não pode ser maior que a data atual.", 48, NOME_SISTEMA_DEFAULT
      txt_periodo_inicial.Text = ""
      txt_periodo_inicial.SetFocus
      Exit Sub
   End If
  
  If txtCondContr = "Microempresa" Then
   If Mid$(Trim(txt_periodo_inicial), 1, 2) <> "03" And Mid$(Trim(txt_periodo_inicial), 1, 2) <> "06" And Mid$(Trim(txt_periodo_inicial), 1, 2) <> "09" And Mid$(Trim(txt_periodo_inicial), 1, 2) <> "12" Then
         MsgBox "Período incompatível com a condição da declaração.", 48, NOME_SISTEMA_DEFAULT
         txt_periodo_inicial.SetFocus
         Exit Sub
      End If
   End If
   If Len(Trim(txt_periodo_inicial)) <> 0 And Val(Mid$(Trim(txt_periodo_inicial), 3, 6)) < 1991 Then
      MsgBox "Ano do Período não pode ser anterior a 1991.", 48, NOME_SISTEMA_DEFAULT
      txt_periodo_inicial.SetFocus
      Exit Sub
   End If

  ' TestaCampoPreench
   

End Sub













Private Sub txt201_GotFocus()

   txt201.SelStart = 0
   txt201.SelLength = Len(txt201)
   
End Sub


Private Sub txt201_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt201_LostFocus()

  grid_ApurICMS(1, 0) = txt201.Text
  Navega_Ficha(1, 2) = 1
  
  acumula_total_204
End Sub

Private Sub txt202_GotFocus()
   txt202.SelStart = 0
   txt202.SelLength = Len(txt202)

End Sub


Private Sub txt202_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt202_LostFocus()

If Not critica_valor_com_caracteres(txt202.Text) Then
    txt202.Text = "0,00"
  End If

 txt202.Text = Format(txt202.Text, "#,##0.#0")
  grid_ApurICMS(1, 1) = txt202.Text
  
  acumula_total_204
  calcula_dados_ApuracaoImposto_chaim
  
  Navega_Ficha(1, 2) = 1
End Sub

Private Sub txt203_GotFocus()
   txt203.SelStart = 0
   txt203.SelLength = Len(txt203)

End Sub


Private Sub txt203_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt203_LostFocus()

If Not critica_valor_com_caracteres(txt203.Text) Then
    txt203.Text = "0,00"
  End If

 txt203.Text = Format(txt203.Text, "#,##0.#0")
grid_ApurICMS(1, 2) = txt203.Text
acumula_total_204

calcula_dados_ApuracaoImposto_chaim

Navega_Ficha(1, 2) = 1
End Sub

Private Sub txt204_GotFocus()
   txt204.SelStart = 0
   txt204.SelLength = Len(txt204)

End Sub


Private Sub txt204_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt204_LostFocus()
If Not critica_valor_com_caracteres(txt204.Text) Then
    txt204.Text = "0,00"
  End If

 txt204.Text = Format(txt204.Text, "#,##0.#0")
grid_ApurICMS(1, 3) = txt204.Text
End Sub

Private Sub txt205_GotFocus()
   txt205.SelStart = 0
   txt205.SelLength = Len(txt205)

End Sub


Private Sub txt205_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt205_LostFocus()
If Not critica_valor_com_caracteres(txt205.Text) Then
    txt205.Text = "0,00"
End If
txt205.Text = Format(txt205.Text, "#,##0.#0")
grid_ApurICMS(1, 4) = txt205.Text
Navega_Ficha(1, 2) = 1
End Sub

Private Sub txt206_GotFocus()
   txt206.SelStart = 0
   txt206.SelLength = Len(txt206)

End Sub


Private Sub txt206_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt206_LostFocus()
If Not critica_valor_com_caracteres(txt206.Text) Then
    txt206.Text = "0,00"
End If
txt206.Text = Format(txt206.Text, "#,##0.#0")
grid_ApurICMS(1, 5) = txt206.Text
Navega_Ficha(1, 2) = 1
End Sub

Private Sub txt207_GotFocus()
   txt207.SelStart = 0
   txt207.SelLength = Len(txt207)

End Sub


Private Sub txt207_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt208_GotFocus()
   txt208.SelStart = 0
   txt208.SelLength = Len(txt208)

End Sub

Private Sub txt208_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt208_LostFocus()

If Not critica_valor_com_caracteres(txt208.Text) Then
    txt208.Text = "0,00"
  End If

txt208.Text = Format(txt208.Text, "#,##0.#0")
grid_ApurICMS(1, 7) = txt208.Text
Call Calcula_206
Navega_Ficha(1, 2) = 1
End Sub

Private Sub txt209_GotFocus()
   txt209.SelStart = 0
   txt209.SelLength = Len(txt209)

End Sub


Private Sub txt209_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt210_GotFocus()
  txt210.SelStart = 0
   txt210.SelLength = Len(txt210)

End Sub


Private Sub txt210_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt211_GotFocus()
txt211.SelStart = 0
txt211.SelLength = Len(txt211)
End Sub


Private Sub txt211_KeyPress(KeyAscii As Integer)

'If KeyAscii = 13 Then
'    msk_fat_mensal_frm4.SetFocus
'    Exit Sub
' End If
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt212_GotFocus()
  txt212.SelStart = 0
   txt212.SelLength = Len(txt212)

End Sub


Private Sub txt212_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt212_LostFocus()
If Not critica_valor_com_caracteres(txt212.Text) Then
    txt212.Text = "0,00"
End If

txt212.Text = Format(txt212.Text, "#,##0.#0")
grid_ApurICMS(1, 11) = txt212.Text
acumula_total_214

calcula_dados_ApuracaoImposto_chaim
Navega_Ficha(1, 2) = 1
End Sub

Private Sub txt213_GotFocus()
  txt213.SelStart = 0
  txt213.SelLength = Len(txt213)

End Sub


Private Sub txt213_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt400_GotFocus()
 
    txt400.SelStart = 0
    txt400.SelLength = Len(txt400)

End Sub

Private Sub txt400_KeyPress(KeyAscii As Integer)

Call PulaCampo(KeyAscii)

                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

End Sub


Private Sub txt400_LostFocus()

Navega_Ficha(1, 4) = 1

End Sub

Private Sub txt416_GotFocus()
 
    txt416.SelStart = 0
    txt416.SelLength = Len(txt416)

End Sub


Private Sub txt416_KeyPress(KeyAscii As Integer)

Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

txt416 = tirapictu(txt416, ".")
If Len(Trim(txt416)) > 15 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If


End Sub


Private Sub txt416_LostFocus()
   
   Dim valor As String
   
   If Not critica_valor_com_caracteres(txt416.Text) Then
    txt416.Text = "0,00"
  End If
 '  valor = Format$(txt416.Text, "#,###,###,###,##0.#0")
 '  valor = formata_p_impressao(valor, 20)
 '  txt416.Text = valor

 txt416.Text = Format(txt416.Text, "#,##0.#0")
 Navega_Ficha(1, 4) = 1
acumula_total_419
End Sub

Private Sub txt417_GotFocus()
 
    txt417.SelStart = 0
    txt417.SelLength = Len(txt417)

End Sub


Private Sub txt417_KeyPress(KeyAscii As Integer)

Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If


txt417 = tirapictu(txt417, ".")
If Len(Trim(txt417)) > 15 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If

End Sub


Private Sub txt417_LostFocus()

   If Not critica_valor_com_caracteres(txt417.Text) Then
    txt417.Text = "0,00"
  End If


 txt417.Text = Format(txt417.Text, "#,##0.#0")
Navega_Ficha(1, 4) = 1
acumula_total_419

End Sub

Private Sub txt418_GotFocus()
 
    txt418.SelStart = 0
    txt418.SelLength = Len(txt418)

End Sub


Private Sub txt418_KeyPress(KeyAscii As Integer)

Call PulaCampo(KeyAscii)

                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

txt418 = tirapictu(txt418, ".")
If Len(Trim(txt418)) > 15 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If


End Sub


Private Sub txt418_LostFocus()

   If Not critica_valor_com_caracteres(txt418.Text) Then
    txt418.Text = "0,00"
  End If

 txt418.Text = Format(txt418.Text, "#,##0.#0")
Navega_Ficha(1, 4) = 1
acumula_total_419

End Sub

Private Sub txt419_GotFocus()
 
    txt419.SelStart = 0
    txt419.SelLength = Len(txt419)

End Sub


Private Sub txt419_KeyPress(KeyAscii As Integer)

Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If
If Len(Trim(txt419)) > 15 Then
  '    SendKeys "{BackSpace}"
      Exit Sub
  End If

End Sub


Private Sub txt419_LostFocus()

   If Not critica_valor_com_caracteres(txt419.Text) Then
    txt419.Text = "0,00"
  End If

 txt419.Text = Format(txt419.Text, "#,##0.#0")
Navega_Ficha(1, 4) = 1

End Sub

Private Sub txt420_GotFocus()
 
    txt420.SelStart = 0
    txt420.SelLength = Len(txt420)

End Sub


Private Sub txt420_KeyPress(KeyAscii As Integer)

Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

txt420 = tirapictu(txt420, ".")
If Len(Trim(txt420)) > 15 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If


End Sub


Private Sub txt420_LostFocus()

   If Not critica_valor_com_caracteres(txt420.Text) Then
    txt420.Text = "0,00"
  End If

 txt420.Text = Format(txt420.Text, "#,##0.#0")
Navega_Ficha(1, 4) = 1
acumula_total_423
End Sub

Private Sub txt421_GotFocus()
 
    txt421.SelStart = 0
    txt421.SelLength = Len(txt421)

End Sub


Private Sub txt421_KeyPress(KeyAscii As Integer)

Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If


txt421 = tirapictu(txt421, ".")
If Len(Trim(txt421)) > 15 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If



End Sub


Private Sub txt421_LostFocus()

   If Not critica_valor_com_caracteres(txt421.Text) Then
    txt421.Text = "0,00"
  End If

 txt421.Text = Format(txt421.Text, "#,##0.#0")

Navega_Ficha(1, 4) = 1
acumula_total_423
    'txt422.SetFocus
End Sub

Private Sub txt422_GotFocus()
 
    txt422.SelStart = 0
    txt422.SelLength = Len(txt422)

End Sub


Private Sub txt422_KeyPress(KeyAscii As Integer)
   
 
   
Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If


txt422 = tirapictu(txt422, ".")
If Len(Trim(txt422)) > 15 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If


End Sub


Private Sub txt422_LostFocus()

   If Not critica_valor_com_caracteres(txt422.Text) Then
    txt422.Text = "0,00"
  End If

 txt422.Text = Format(txt422.Text, "#,##0.#0")
Navega_Ficha(1, 4) = 1
acumula_total_423
End Sub

Private Sub txt423_GotFocus()
 
    txt423.SelStart = 0
    txt423.SelLength = Len(txt423)

End Sub


Private Sub txt423_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = 13 Or KeyCode = 9 Then
     cmd_salvar.SetFocus
     Exit Sub
   End If

End Sub

Private Sub txt423_KeyPress(KeyAscii As Integer)


   If KeyAscii = 13 Or KeyAscii = 9 Then
     cmd_salvar.SetFocus
     Exit Sub
   End If
'Call PulaCampo(KeyAscii)
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If


If Len(Trim(txt423)) > 15 Then
   '   SendKeys "{BackSpace}"
      Exit Sub
  End If

End Sub


Private Sub txt423_LostFocus()

   If Not critica_valor_com_caracteres(txt423.Text) Then
    txt423.Text = "0,00"
  End If


 txt423.Text = Format(txt423.Text, "#,##0.#0")
Navega_Ficha(1, 4) = 1

End Sub

Private Sub txt424_GotFocus()

    txt424.SelStart = 0
    txt424.SelLength = Len(txt424)
    
End Sub


Private Sub txt424_KeyPress(KeyAscii As Integer)

Call PulaCampo(KeyAscii)
                    
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") And KeyAscii <> 8 Then
      KeyAscii = 0
   ElseIf KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
      KeyAscii = Asc(",")
   End If

txt424 = tirapictu(txt424, ".")
If Len(Trim(txt424)) > 15 Then
    If KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If


End Sub


Private Sub txt424_LostFocus()

   If Not critica_valor_com_caracteres(txt424.Text) Then
    txt424.Text = "0,00"
  End If


 txt424.Text = Format(txt424.Text, "#,##0.#0")
Navega_Ficha(1, 4) = 1

End Sub


