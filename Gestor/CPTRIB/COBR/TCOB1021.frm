VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCOB102 
   Caption         =   "VS"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1138
      Icone           =   "TCOB1021.frx":0000
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   3120
      Left            =   0
      TabIndex        =   32
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5503
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483637
      Ocultavel       =   0   'False
      Begin VB.TextBox txtDataEmissao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
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
         MaxLength       =   14
         TabIndex        =   1
         Tag             =   "Validade"
         Top             =   360
         Width           =   1605
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   22
         Left            =   3840
         TabIndex        =   52
         Top             =   420
         Width           =   1110
         _ExtentX        =   1958
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
         Caption         =   "Data Emissão:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   21
         Left            =   1200
         TabIndex        =   51
         Top             =   420
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   370
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         PictureMaskColor=   -2147483637
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Controle:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtControle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   0
         Tag             =   "Nome ou Razão Social:"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtAtividade 
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
         Left            =   2100
         TabIndex        =   8
         Top             =   2370
         Width           =   7725
      End
      Begin VB.TextBox txtDam 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2070
         MaxLength       =   14
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   -540
         Width           =   1305
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   0
         Left            =   465
         TabIndex        =   41
         Top             =   2070
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   370
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         PictureMaskColor=   -2147483637
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Nome de Fantasia:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   5
         Left            =   3840
         TabIndex        =   40
         Top             =   885
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         PictureMaskColor=   -2147483637
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CPF ou CNPJ:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   2
         Left            =   2835
         TabIndex        =   39
         Top             =   -285
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   318
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         PictureMaskColor=   -2147483637
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "NO. Pagamento:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   150
         Index           =   1
         Left            =   75
         TabIndex        =   38
         Top             =   2460
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   265
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         PictureMaskColor=   -2147483637
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Atividade Economica:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   4
         Left            =   1620
         TabIndex        =   37
         Top             =   870
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         PictureMaskColor=   -2147483637
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "IM.:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesq 
         Height          =   375
         Index           =   1
         Left            =   6720
         TabIndex        =   4
         Top             =   825
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VB.TextBox txtfantasia 
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
         Left            =   2100
         TabIndex        =   7
         Top             =   2010
         Width           =   7725
      End
      Begin VB.TextBox txtcgc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
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
         MaxLength       =   18
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "CPF ou  CGC:"
         Top             =   840
         Width           =   1605
      End
      Begin VB.TextBox txtIm 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
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
         MaxLength       =   18
         TabIndex        =   2
         Tag             =   "CPF ou  CGC:"
         Top             =   840
         Width           =   1605
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   7
         Left            =   960
         TabIndex        =   36
         Top             =   1290
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   318
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         PictureMaskColor=   -2147483637
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Doc.Origem:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtDocOrigem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
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
         MaxLength       =   14
         TabIndex        =   5
         Top             =   1245
         Width           =   1605
      End
      Begin Threed.SSPanel lbl 
         Height          =   150
         Index           =   9
         Left            =   15
         TabIndex        =   35
         Top             =   2835
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   265
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         PictureMaskColor=   -2147483637
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ativ. Econ. do Veiculo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtAtividadeEconomicaVEiculo 
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
         Left            =   2100
         TabIndex        =   9
         Top             =   2745
         Width           =   7725
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   11
         Left            =   3840
         TabIndex        =   34
         Top             =   1320
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   318
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         PictureMaskColor=   -2147483637
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Placa do Veiculo ou Doc Origem dos Anuncios"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtrazao 
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
         Left            =   2100
         TabIndex        =   6
         Tag             =   "Nome ou Razão Social:"
         Top             =   1650
         Width           =   7710
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   8
         Left            =   930
         TabIndex        =   33
         Top             =   1710
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   370
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         PictureMaskColor=   -2147483637
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Razão/Nome:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   2745
      Left            =   0
      TabIndex        =   43
      Top             =   3840
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4842
      Altura          =   1905
      Caption         =   " Documentos"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483637
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   23
         Left            =   240
         TabIndex        =   54
         Top             =   2040
         Width           =   1830
         _ExtentX        =   3228
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
         Caption         =   "Horário dos Eventos:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtHorarioEventos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   2130
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2040
         Width           =   7755
      End
      Begin VB.CheckBox chkProvisorio 
         Caption         =   "Provisório ?"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3840
         TabIndex        =   53
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtRestricao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2130
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1140
         Width           =   7755
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
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "Finalidade do Alvará"
         Top             =   750
         Width           =   7755
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   10
         Left            =   15
         TabIndex        =   48
         Top             =   780
         Width           =   2070
         _ExtentX        =   3651
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
         Caption         =   "Finalidade de Solicitação:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   12
         Left            =   1080
         TabIndex        =   47
         Top             =   1170
         Width           =   990
         _ExtentX        =   1746
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
         Caption         =   "Restrições:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   3
         Left            =   1170
         TabIndex        =   46
         Top             =   480
         Width           =   870
         _ExtentX        =   1535
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
         Caption         =   "Validade:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   6
         Left            =   6810
         TabIndex        =   45
         Top             =   390
         Width           =   630
         _ExtentX        =   1111
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
         Caption         =   "Tipo:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "TCOB1021.frx":031A
         Left            =   7470
         List            =   "TCOB1021.frx":0339
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "Tipo"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtValidade 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2130
         MaxLength       =   14
         TabIndex        =   10
         Tag             =   "Validade"
         Top             =   390
         Width           =   1605
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   13
         Left            =   840
         TabIndex        =   44
         Top             =   1590
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
         Caption         =   "Observações:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2130
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1560
         Width           =   7755
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL3 
      Height          =   2385
      Left            =   0
      TabIndex        =   49
      Top             =   6600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4207
      Altura          =   1905
      Caption         =   " ESPECIFICAÇÃO"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483637
      Ocultavel       =   0   'False
      Begin VB.Frame frmePublicidade 
         Caption         =   "PUBLICIDADE"
         Height          =   1980
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox txtTipoPublicidade 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   6195
         End
         Begin VB.TextBox txtCores 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   1605
            Width           =   6195
         End
         Begin VB.TextBox txtMaterialConfeccao 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   1260
            Width           =   6195
         End
         Begin VB.TextBox txtInscricaoTexto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   915
            Width           =   6195
         End
         Begin VB.TextBox txtDimensao 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
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
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   570
            Width           =   1635
         End
         Begin VB.TextBox txtLocal 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   570
            Width           =   3555
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   24
            Left            =   120
            TabIndex        =   63
            Top             =   915
            Width           =   1590
            _ExtentX        =   2805
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
            Caption         =   "Inscrições e texto"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   25
            Left            =   5520
            TabIndex        =   64
            Top             =   600
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
            Caption         =   "Dimensão"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   26
            Left            =   120
            TabIndex        =   65
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
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
            Caption         =   "Localização"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   27
            Left            =   120
            TabIndex        =   66
            Top             =   1260
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Material confecção"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   28
            Left            =   120
            TabIndex        =   67
            Top             =   1605
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Cores empregadas"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   29
            Left            =   120
            TabIndex        =   68
            Top             =   230
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Tipo Publicidade"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin VB.Frame frmFuncionamento 
         Caption         =   "FUNCIONAMENTO"
         Height          =   1815
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   8055
         Begin VB.TextBox txtSegSex 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            MaxLength       =   14
            TabIndex        =   16
            Top             =   360
            Width           =   1125
         End
         Begin VB.TextBox txtSabado 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            MaxLength       =   14
            TabIndex        =   18
            Top             =   900
            Width           =   1125
         End
         Begin VB.TextBox txtDomingo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            MaxLength       =   14
            TabIndex        =   20
            Top             =   1440
            Width           =   1125
         End
         Begin VB.TextBox txtSegSexAs 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   3600
            MaxLength       =   14
            TabIndex        =   17
            Top             =   360
            Width           =   1125
         End
         Begin VB.TextBox txtSabadoAs 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   3600
            MaxLength       =   14
            TabIndex        =   19
            Top             =   900
            Width           =   1125
         End
         Begin VB.TextBox txtDomingoAs 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   3600
            MaxLength       =   14
            TabIndex        =   21
            Top             =   1440
            Width           =   1125
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   20
            Left            =   3120
            TabIndex        =   56
            Top             =   1440
            Width           =   375
            _ExtentX        =   661
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
            Caption         =   "as"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   19
            Left            =   3120
            TabIndex        =   57
            Top             =   900
            Width           =   375
            _ExtentX        =   661
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
            Caption         =   "as"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   17
            Left            =   3120
            TabIndex        =   58
            Top             =   360
            Width           =   375
            _ExtentX        =   661
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
            Caption         =   "as"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   15
            Left            =   120
            TabIndex        =   59
            Top             =   1440
            Width           =   1590
            _ExtentX        =   2805
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
            Caption         =   "Domingo / Feriado"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   14
            Left            =   1080
            TabIndex        =   60
            Top             =   900
            Width           =   630
            _ExtentX        =   1111
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
            Caption         =   "Sábado"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   180
            Index           =   16
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   1590
            _ExtentX        =   2805
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
            Caption         =   "De Segunda à Sexta"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   3
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8280
         TabIndex        =   30
         Top             =   1680
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   8280
         TabIndex        =   28
         Top             =   600
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpa 
         Height          =   375
         Left            =   8280
         TabIndex        =   29
         Top             =   1140
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Threed.SSPanel lbl 
      Height          =   180
      Index           =   18
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   375
      _ExtentX        =   661
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
      Caption         =   "as"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "TCOB102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim Cobranca As New VSCobranca
Dim Imposto As New VSImposto
Dim InscMuni As String
Dim RazaoSocial As String
Dim NomeFantasia As String
Dim Atividade As String
Dim Restricoes As String
Dim Finalidade As String
Dim CPFCNPJ As String
Dim Endereco As String
Dim IM_Auxiliar As String
Dim Bairro As String
Dim Cidade As String
Dim Cep As String
Dim Uf As String
Dim Atividade_Veiculo As String
Public Carro As Boolean
Dim Validade As String
Dim CodSituacao As String
Private cae As String
Private observacao As String
Dim cf As String
Dim alteracao As Boolean

Private Sub chkProvisorio_Click()
    If chkProvisorio.Value = 1 Then
        txtValidade = Format(DateAdd("d", 30, Now), "DD/MM/YYYY")
    Else
        txtValidade = "31/12/" & Format(Now, "yyyy")
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim a As Integer
    Dim CodAlvara As String
    Dim codigoVerificador As String
    Dim Valores As String
    Dim Campos As String
    
    
    On Error Resume Next
    
    ''BCP
    'CodAlvara = BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto = '" & Imposto.NomeTributo(ttr_ALVARA) & " '")
    If alteracao = True Then
        CodAlvara = txtControle
    Else
        CodAlvara = Imposto.GeraNumNota(1, 5)
    End If
    'If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    Restricoes = "" & UCase(txtRestricao.Text)
    Finalidade = "" & UCase(txtMotivo)
    
    'Bdados.DeletaDados "TAB_ALVARA_IMPRESSO", "tai_tci_im ='" & txtIm & "' AND tai_periodo =" & Year(Date)
    'BCP - ATUALIZA AO INVES DE DELETAR DO BANCO
    If Len(Trim(txtDocOrigem)) = "" Or txtDocOrigem = 0 Then
        Bdados.Executa ("UPDATE TAB_ALVARA_IMPRESSO SET TAI_STATUS='CANCELADO' WHERE tai_tci_im ='" & txtIm & "' AND tai_periodo =" & Year(Date))
    End If
    codigoVerificador = Year(Date) & Right(txtIm, 2) & CodAlvara
    Valores = Bdados.PreparaValor(txtIm, Year(Date), Bdados.Converte(Date, TCDataHora), _
    AplicacoesVTFuncoes.Usuario, Bdados.Converte(txtValidade, TCDataHora), CodAlvara, txtSegSex, txtSegSexAs, _
    txtSabado, txtSabadoAs, txtDomingo, txtDomingoAs, codigoVerificador, txtDocOrigem, "EMITIDO", txtHorarioEventos, _
    txtLocal, txtDimensao, txtInscricaoTexto, txtMaterialConfeccao, txtCores, txtTipoPublicidade)
    
    Campos = "tai_tci_im,tai_periodo,tai_data_impressao,tai_tus_cod_usuario,tai_data_validade," _
    & "tai_cod_sequencial,tai_hf_segsex,tai_hf_segsex_as,tai_hf_sabado,tai_hf_sabado_as,tai_hf_domingo," _
    & "tai_hf_domingo_as,tai_cod_verificador,tai_doc_origem,tai_status,tai_hr_especifico," _
    & "tai_local,tai_dimensao,tai_texto,tai_material,tai_cores,tai_tipo_publicidade"
    
    
    'BCP - GLEYSON
    If alteracao = True Then
        
    Else
        'Bdados.GravaDados "TAB_ALVARA_IMPRESSO", Valores, Campos, "tai_tci_im='" & txtIm & "' AND tai_periodo =" & Year(Date)
        Bdados.InsereDados "TAB_ALVARA_IMPRESSO", Valores, Campos
    End If
    If Len(cf) = 0 Then
        cf = cae
    End If
    cae = Left(cf, 2) & "." & Mid(cf, 3, 2) & "-" & Mid(cf, 5, 1) & "-" & Right(cf, 2)
    
    Select Case cboTipo.ListIndex
        Case 0 'PADRAO
            With RPT
                If Not .DefinirArquivo(Bdados, App.Path + "\TAlvara.rpt") Then Exit Sub
                .Formulas "IM", InscMuni
                .Formulas "RAZAOSOCIAL", RazaoSocial
                .Formulas "NOMEFANTASIA", NomeFantasia
                If txtDocOrigem = "" Then
                   .Formulas "ATIVIDADE", Atividade
                Else
                    .Formulas "ATIVIDADE", txtAtividadeEconomicaVEiculo
                End If
                .Formulas "RESTRICOES", Restricoes
                .Formulas "OBJETIVO", Finalidade
                .Formulas "CPF/CNPJ", CPFCNPJ
                .Formulas "ENDERECO", Endereco
                .Formulas "BAIRRO", Bairro
                .Formulas "CidadeEmpresa", Cidade
                .Formulas "CEP", Cep
                .Formulas "ESTADO", Uf
                .Formulas "VALIDADE", txtValidade
                .Formulas "OBSERVACAO", txtObs
                .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
                .Formulas "CIDADE", Aplicacoes.municipio & " - " & UCase(Temp.PegaParametro(Bdados, "ESTADO CLIENTE"))
                .Formulas "DEPARTAMENTO", Temp.PegaParametro(Bdados, "SETOR")
                .Titulo = Imposto.NomeTributo(ttr_ALVARA)
                .Arvore = False
                .Visualizar
            End With
        Case 1 'VEICULO
            With RPT
                If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                    If Not .DefinirArquivo(Bdados, App.Path + "\TAlvaraMotoTaxiBarraMansa.rpt") Then Exit Sub
                Else
                    If Not .DefinirArquivo(Bdados, App.Path + "\TALVARA_VEICULO.rpt") Then Exit Sub
                End If
                .Formulas "IM", InscMuni
                .Formulas "RAZAOSOCIAL", RazaoSocial
                .Formulas "NOMEFANTASIA", NomeFantasia
                If txtDocOrigem = "" Then
                   .Formulas "ATIVIDADE", Atividade
                Else
                    .Formulas "ATIVIDADE", txtAtividadeEconomicaVEiculo
                End If
                .Formulas "RESTRICOES", Restricoes
                .Formulas "OBJETIVO", Finalidade
                .Formulas "CPF/CNPJ", CPFCNPJ
                .Formulas "ENDERECO", Endereco
                .Formulas "BAIRRO", Bairro
                .Formulas "CidadeEmpresa", Cidade
                .Formulas "CEP", Cep
                .Formulas "ESTADO", Uf
                .Formulas "VALIDADE", txtValidade
                .Formulas "OBSERVACAO", txtObs
                .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
                .Formulas "CIDADE", Aplicacoes.municipio & " - " & UCase(Temp.PegaParametro(Bdados, "ESTADO CLIENTE"))
                .Formulas "DEPARTAMENTO", Temp.PegaParametro(Bdados, "SETOR")
                .Titulo = Imposto.NomeTributo(ttr_ALVARA)
                .Arvore = False
                .Visualizar
            End With
        Case 2, 3, 6, 7, 8 'FUNCIONAMENTO/LOCALIZACAO/HORARIO ESPECIAL / EVENTOS / PUBLICIDADE
            Dim Arquivo As String
            
            If cboTipo.ListIndex = 2 Then
                Arquivo = "\TAlvaraFuncionamento.rpt"
            ElseIf cboTipo.ListIndex = 3 Then
                Arquivo = "\TAlvaraLocalizacao.rpt"
            ElseIf cboTipo.ListIndex = 6 Then
                Arquivo = "\TAlvaraHorarioEspecial.rpt"
            ElseIf cboTipo.ListIndex = 7 Then
                Arquivo = "\TAlvaraEventos.rpt"
            ElseIf cboTipo.ListIndex = 8 Then
                Arquivo = "\TAlvaraPublicidade.rpt"
            End If
            With RPT
                If Not .DefinirArquivo(Bdados, App.Path & Arquivo) Then Exit Sub
                ''BCP - GLEYSON
                Dim provisorio As String
                provisorio = "PROVISÓRIO"
                If chkProvisorio.Value = 0 Then  'FALSE
                    provisorio = ""
                End If
                
                .Formulas "TIPO_PUBLICIDADE", txtTipoPublicidade
                .Formulas "LOCAL", txtLocal
                .Formulas "CORES", txtCores
                .Formulas "DIMENSAO", txtDimensao
                .Formulas "TEXTO", txtInscricaoTexto
                .Formulas "MATERIAL", txtMaterialConfeccao
                
                
                .Formulas "PROVISORIO", provisorio
                .Formulas "ESPECIFICO", txtHorarioEventos
                
                .Formulas "CONTROLE", CodAlvara
                .Formulas "COD_ATIVIDADE", cae
                .Formulas "VT_Restricao", Restricoes
                .Formulas "CODIGO_VERIFICADOR", codigoVerificador
                
                .Formulas "SABADO", txtSabado
                .Formulas "SABADOAS", txtSabadoAs
                .Formulas "SEGSEX", txtSegSex
                .Formulas "SEGSEXAS", txtSegSexAs
                .Formulas "DOMINGO", txtDomingo
                .Formulas "DOMINGOAS", txtDomingoAs
                .Formulas "VEICULO", txtRestricao
                .Formulas "Atividade_Veiculo", txtAtividadeEconomicaVEiculo
                
                
                ''FIM
                .Formulas "OBJETIVO", Finalidade
                .Formulas "RAZAOSOCIAL", RazaoSocial
                .Formulas "NOMEFANTASIA", NomeFantasia
                .Formulas "CPF/CNPJ", CPFCNPJ
                .Formulas "IM", InscMuni
                .Formulas "ENDERECO", Trim$(Endereco) & " - " & Bairro & " - " & Cidade
                .Formulas "OBSERVACAO", txtObs
                .Formulas "EMISSAO", Format(txtDataEmissao, "DD/MM/YYYY")
                If txtDocOrigem = "" Then
                   .Formulas "ATIVIDADE", Atividade
                Else
                   .Formulas "ATIVIDADE", txtAtividadeEconomicaVEiculo
                End If
                .Formulas "VALIDADE", txtValidade
                .Formulas "CIDADE", Aplicacoes.municipio
                If cboTipo.ListIndex = 2 Then
                    .Formulas "VT_RESTRICOES", Restricoes
                End If
                .Titulo = Imposto.NomeTributo(ttr_ALVARA)
                .Arvore = False
                .Visualizar
            End With
        Case 4 'CRACHA DE TAXI
            With RPT
                If Not .DefinirArquivo(Bdados, App.Path + "\TCrachaMotoTaxi.rpt") Then Exit Sub
                .Selecao = "{Tab_Contribuinte.tci_im} = '" & txtIm & "' AND {TAB_ALVARA_IMPRESSO.TAI_PERIODO} =" & Year(Date) '& " AND {TAB_TRANSPORTADOR_VEICULO.TTV_PLACA} = '" & txtDocOrigem & "'"
                .Formulas "VTFinalidade", txtMotivo
                .Formulas "VT_ESTADO", Temp.PegaParametro(Bdados, "ESTADO")
                .Formulas "VT_PREFEITURA", Temp.PegaParametro(Bdados, "CLIENTE")
                .Formulas "VT_LOCAL_DATA", AplicacoesVTFuncoes.municipio & " - " & Temp.PegaParametro(Bdados, "ESTADO CLIENTE") & "  " & UCase(Left(FormatDateTime(Date, vbLongDate), 1)) & Right(FormatDateTime(Date, vbLongDate), Len(FormatDateTime(Date, vbLongDate)) - 1)
                .Formulas "VTFinalidade", txtMotivo
                .Titulo = Imposto.NomeTributo(ttr_ALVARA)
                .Arvore = False
                .Visualizar
            End With
        Case 5 'ALVARA DE LICENSA
            
            
            If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                With RPT
                    If Not .DefinirArquivo(Bdados, App.Path + "\ALVARABARRAMANSA.rpt") Then Exit Sub
                    .SubRelatorio = "TAlvaraSubAtividade.rpt"
                    .Selecao = "{TAB_ATIVIDADE_SECUNDARIA.TAS_TCI_IM} = '" & txtIm & "'"
                    .SubRelatorio = ""
                    .Formulas "NOMEFANTASIA", txtfantasia
                    .Formulas "ENDERECO ", Endereco
                    .Formulas "ATIVIDADE", Atividade
                    .Formulas "IM", txtIm
                    .Formulas "CADASTROUNICO", IM_Auxiliar
                    If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                        If CodSituacao = 1 Then
                            .Formulas "VALIDADE", Validade
                        Else
                            .Formulas "VALIDADE", txtValidade
                        End If
                    Else
                        .Formulas "VALIDADE", txtValidade
                    End If
                    .Formulas "CPF/CNPJ", txtcgc
                    .Formulas "RAZAOSOCIAL", txtrazao
                    .Formulas "DATADIA", Format(Day(Date), "00")
                    .Formulas "DATAMES", Format(Month(Date), "00")
                    .Formulas "DATAANO", Format(Year(Date), "00")
                    .Formulas "RESTRICOES", txtRestricao
                    .Visualizar
                End With
            ElseIf UCase(AplicacoesVTFuncoes.municipio) = "PETROLINA" Then
                With RPT
                    If Not .DefinirArquivo(Bdados, App.Path + "\TAlvarapETROLINA.rpt") Then Exit Sub
                                    .Formulas "IM", InscMuni
                    .Formulas "RAZAOSOCIAL", RazaoSocial
                    '.Formulas "NOMEFANTASIA", NomeFantasia
                    If txtDocOrigem = "" Then
                       .Formulas "ATIVIDADE", Atividade
                    Else
                        .Formulas "ATIVIDADE", txtAtividadeEconomicaVEiculo
                    End If
                    .Formulas "RESTRICOES", Restricoes
                    '.Formulas "OBJETIVO", Finalidade
                    .Formulas "CPF/CNPJ", CPFCNPJ
                    .Formulas "ENDERECO", Endereco
                    .Formulas "BAIRRO", Bairro
                    .Formulas "CidadeEmpresa", Cidade
                    .Formulas "CEP", Cep
                    .Formulas "VALIDADE", UCase(Left(FormatDateTime(txtValidade, vbLongDate), 1)) & Right(FormatDateTime(txtValidade, vbLongDate), Len(FormatDateTime(txtValidade, vbLongDate)) - 1)
                    .Formulas "DATA", UCase(Left(FormatDateTime(Date, vbLongDate), 1)) & Right(FormatDateTime(Date, vbLongDate), Len(FormatDateTime(Date, vbLongDate)) - 1)
                    .Formulas "ESTADO", Uf
                    '.Formulas "VALIDADE", txtValidade
                    '.Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
                    .Formulas "CIDADE", Aplicacoes.municipio
                    '.Formulas "DEPARTAMENTO", Temp.PegaParametro(Bdados, "SETOR")
                    .Visualizar
                 End With
            ElseIf UCase(AplicacoesVTFuncoes.municipio) = "COLINAS" Then
                With RPT
                    If Not .DefinirArquivo(Bdados, App.Path + "\TALVARA.rpt") Then Exit Sub
                    .Formulas "IM", InscMuni
                    .Formulas "RAZAOSOCIAL", RazaoSocial
                    .Formulas "NOMEFANTASIA", NomeFantasia
                    If txtDocOrigem = "" Then
                       .Formulas "ATIVIDADE", Atividade
                    Else
                        .Formulas "ATIVIDADE", txtAtividadeEconomicaVEiculo
                    End If
                    .Formulas "RESTRICOES", Restricoes
                    .Formulas "OBJETIVO", Finalidade
                    .Formulas "CPF/CNPJ", CPFCNPJ
                    .Formulas "ENDERECO", Endereco
                    .Formulas "BAIRRO", Bairro
                    .Formulas "CidadeEmpresa", Cidade
                    .Formulas "CEP", Cep
                    .Formulas "ESTADO", Uf
                    .Formulas "VALIDADE", txtValidade
    
                    .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
                    .Formulas "CIDADE", Aplicacoes.municipio & " - " & UCase(Temp.PegaParametro(Bdados, "ESTADO CLIENTE"))
                    .Formulas "DEPARTAMENTO", Temp.PegaParametro(Bdados, "SETOR")
                    .Titulo = Imposto.NomeTributo(ttr_ALVARA)
                    .Arvore = False
                    .Visualizar
                 End With
            End If
                
            
    End Select
    Screen.MousePointer = 0
    Set RPT = Nothing
    Call Util.Informa("ALVARÁ emitido.")
    If Me.Tag = "EXTERNO" Then cmdSair_Click
End Sub
Public Sub PreencheTela(Criterio As String, Optional Conexao As Object)
    Dim Rs As VSRecordset
    Dim Sql As String
        
    Sql = " Select * from Tab_Contribuinte " _
        & " where " & Criterio
    
    If Not Conexao Is Nothing Then Set Bdados = Conexao
    LimpaCampos Me
    chkProvisorio.Enabled = False
    chkProvisorio.Value = 0
    If Bdados.AbreTabela(Sql, Rs) Then
        chkProvisorio.Enabled = True
        TxtAtividade = Imposto.BuscaNomeCAE("" & Rs("tci_tae_cae"))
        txtIm = Rs!TCI_IM
        txtcgc = "" & Rs("tci_cgc_cpf")
        txtrazao = "" & Rs("tci_nome")
        txtfantasia = "" & Rs("tci_fantasia")
        Atividade = "" & UCase(TxtAtividade.Text)
        InscMuni = "" & Rs("tci_im")
        RazaoSocial = "" & Rs("tci_nome")
        NomeFantasia = "" & Rs("tci_fantasia")
        IM_Auxiliar = "" & Rs.Fields("tci_im_auxiliar")
        CPFCNPJ = "" & Rs("tci_cgc_cpf")
        Endereco = "" & Rs("tci_logradouro") & " " & Rs("tci_nome_logradouro") & "," & Rs("tci_numero") & " " & Rs("tci_complemento")
        Bairro = "" & Rs("tci_bairro")
        Cidade = "" & Rs("tci_cidade")
        Cep = "" & Rs("tci_cep")
        cae = "" & Rs("tci_tae_cae")
        cboTipo = "FUNCIONAMENTO"
        txtSegSex.Text = "08:00"
    txtSegSexAs.Text = "18:00"
    
    txtSabado.Text = "08:00"
    txtSabadoAs.Text = "12:00"
    
    txtDomingo.Text = ""
    txtDomingoAs.Text = ""
        If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
            Validade = "" & Pega_Situacao_Alvara("" & Rs.Fields("TCI_SIT_ALVARA"))
            If CodSituacao <> 1 Then
                txtValidade.Visible = False
            Else
                txtValidade.Visible = True
            End If
        End If
        Uf = "" & Rs("tci_uf")
        txtValidade = "31/12/" & Format(Now, "yyyy")
        
        Sql = "Select TTV_TAE_CAE, TTV_VEICULO,TTV_MARCA,TTV_COD_MODELO,TTV_ANO_FABRICACAO,TTV_PLACA,TTV_PLACA," & _
            "TTV_MUNICIPIO,TTV_UF,TTV_LICENCA,TTV_CHASSI FROM TAB_TRANSPORTADOR_VEICULO WHERE TTV_TCI_IM='" & txtIm & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            txtRestricao = "" & Rs!TTV_MARCA & "/" & Rs!TTV_COD_MODELO & ", PLACA " & Rs!ttv_placa & ", ANO " & _
            Rs!TTV_ANO_FABRICACAO & ", LICENCIAMENTO " & Rs!TTV_LICENCA & ", CHASSI " & Rs!TTV_CHASSI
            
        Else
            txtRestricao = ""
        End If
        If txtDocOrigem = "" Then
            txtAtividadeEconomicaVEiculo = "*****************************************************************************"
        End If
        Bdados.FechaTabela Rs
    Else
        Avisa "Contribuinte não cadastrado."
        txtcgc.SetFocus
    End If
    txtDataEmissao = Format(Now, "DD/MM/YYYY")
    Bdados.FechaTabela Rs
End Sub
Private Function Pega_Situacao_Alvara(Situacao As String)
    If Bdados.AbreTabela("select * from VIS_SITUACAO_ALVARA WHERE tge_codigo = '" & Situacao & "'") Then
        Pega_Situacao_Alvara = Bdados.Tabela("Tge_nome")
        CodSituacao = Bdados.Tabela("Tge_codigo")
    End If
End Function

Private Sub cboTipo_Click()
    frmePublicidade.Visible = False
    If cboTipo.ListIndex = 2 Then
        txtMotivo = "ALVARÁ DE LICENÇA E FUNCIONAMENTO"
    ElseIf cboTipo.ListIndex = 3 Then
        txtMotivo = "ALVARÁ DE LICENÇA DE FUNCIONAMENTO E LOCALIZAÇÃO"
    ElseIf cboTipo.ListIndex = 6 Then
        txtMotivo = "ALVARÁ DE HORÁRIO ESPECIAL"
    ElseIf cboTipo.ListIndex = 7 Then
        txtMotivo = "ALVARÁ DE EVENTOS"
    ElseIf cboTipo.ListIndex = 8 Then
        txtMotivo = "ALVARÁ DE PUBLICIDADE"
        frmePublicidade.Visible = True
        txtTipoPublicidade.SetFocus
    End If
    
    txtSegSex.Text = "08:00"
    txtSegSexAs.Text = "18:00"
    
    txtSabado.Text = "08:00"
    txtSabadoAs.Text = "12:00"
    
    txtDomingo.Text = ""
    txtDomingoAs.Text = ""
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdLimpa_Click()
    alteracao = False
End Sub

Private Sub cmdPesq_Click(Index As Integer)
    AplicacoesVTFuncoes.BuscaNoEconomico TcoJuridica, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    alteracao = False
    observacao = "ESTE ALVARÁ TERÁ VALIDADE PARA O FIM ACIMA MENCIONADO E DEVERÁ SER AFIXADO EM LOCAL VISÍVEL À FISCALIZAÇÃO"
    txtObs = observacao
    cabVisual.Exibir Bdados, Me.Name, App.Path
    'txtIm.SetFocus
    txtDataEmissao = Format(Now(), "DD/MM/YYYY")
End Sub
Private Sub Text1_Change()

End Sub

Private Sub txtControle_LostFocus()
    Dim Rs As VSRecordset
    Dim Sql As String
    If txtControle.Text <> "" Then
        Sql = " select * from TAB_ALVARA_IMPRESSO where  tai_cod_sequencial = " & txtControle
        If Bdados.AbreTabela(Sql, Rs) Then
            txtIm = Rs("tai_tci_im")
            Call txtIm_LostFocus
            txtIm = Rs("tai_tci_im")
            txtDataEmissao = Rs("tai_data_impressao")
            txtControle = Rs("tai_cod_sequencial")
            alteracao = True
            frmePublicidade.Visible = False
            If IsNull(Rs("tai_hf_segsex")) Then
                cboTipo = "PUBLICIDADE"
                frmFuncionamento.Visible = True
            Else
                cboTipo = "FUNCIONAMENTO"
            End If
            txtSegSex = IIf(IsNull(Rs("tai_hf_segsex")), "", Rs("tai_hf_segsex"))
            txtSegSexAs = IIf(IsNull(Rs("tai_hf_segsex_as")), "", Rs("tai_hf_segsex_as"))
                
            txtSabado = IIf(IsNull(Rs("tai_hf_sabado")), "", Rs("tai_hf_sabado"))
            txtSabadoAs = IIf(IsNull(Rs("tai_hf_sabado_as")), "", Rs("tai_hf_sabado_as"))
            
            txtDomingo = IIf(IsNull(Rs("tai_hf_domingo")), "", Rs("tai_hf_domingo"))
            txtHorarioEventos = IIf(IsNull(Rs("tai_hr_especifico")), "", Rs("tai_hr_especifico"))
            txtDomingoAs = IIf(IsNull(Rs("tai_hf_domingo_as")), "", Rs("tai_hf_domingo_as"))
            txtIm.Enabled = False
            txtcgc.Enabled = False
            
            
            txtLocal = IIf(IsNull(Rs("tai_local")), "", Rs("tai_local"))
            txtDimensao = IIf(IsNull(Rs("tai_dimensao")), "", Rs("tai_dimensao"))
            txtInscricaoTexto = IIf(IsNull(Rs("tai_texto")), "", Rs("tai_texto"))
            txtMaterialConfeccao = IIf(IsNull(Rs("tai_material")), "", Rs("tai_material"))
            txtCores = IIf(IsNull(Rs("tai_cores")), "", Rs("tai_cores"))
            txtTipoPublicidade = IIf(IsNull(Rs("tai_tipo_publicidade")), "", Rs("tai_tipo_publicidade"))
            
        End If
    Else
        alteracao = False
        txtIm.Enabled = True
        txtcgc.Enabled = True
    End If
End Sub

Private Sub txtDocOrigem_LostFocus()
    Dim Rs As VSRecordset
    Dim Sql As String
    If txtDocOrigem = "" Then Exit Sub
     
    'Checo se existe carros para o doc de origem...
    
     Sql = " Select * from Tab_Contribuinte ,TAB_TRANSPORTADOR_VEICULO" _
        & " where TTV_TCI_IM = TCI_IM AND  TTV_PLACA = " & Bdados.Converte(txtDocOrigem, tctexto) & " AND TTV_TCI_IM ='" & txtIm & "'"
         Carro = True
    If Not Bdados.AbreTabela(Sql) Then
            Carro = False
            Sql = "SELECT * FROM TAB_CONTRIBUINTE,TAB_ANUNCIO"
            Sql = Sql & " Where TAN_TCI_IM = TCI_IM"
            Sql = Sql & " AND TAN_DOC_ORIGEM = " & Bdados.Converte(txtDocOrigem, tctexto)
            Sql = Sql & " AND TAN_TCI_IM ='" & txtIm & "'"
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        LimpaCampos Me
        'TxtAtividade = Imposto.BuscaNomeCAE(Rs("tci_tae_cae"))
        txtRestricao = ""
        txtIm = Rs!TCI_IM
        txtcgc = "" & Rs("tci_cgc_cpf")
        txtrazao = "" & Rs("tci_nome")
        txtfantasia = "" & Rs("tci_fantasia")
        Atividade = "" & UCase(TxtAtividade.Text)
        InscMuni = "" & Rs("tci_im")
        RazaoSocial = "" & Rs("tci_nome")
        NomeFantasia = "" & Rs("tci_fantasia")
        CPFCNPJ = "" & Rs("tci_cgc_cpf")
        Endereco = "" & Rs("tci_logradouro") & " " & Rs("tci_nome_logradouro") & "," & Rs("tci_numero") & " " & Rs("tci_complemento")
        Bairro = "" & Rs("tci_bairro")
        Cidade = "" & Rs("tci_cidade")
        Cep = "" & Rs("tci_cep")
        Uf = "" & Rs("tci_uf")
        txtValidade = "31/12/" & Format(Now, "yyyy")
        cae = "" & Rs("tci_tae_cae")
        If Carro Then
            txtRestricao = "" & Rs!TTV_MARCA & "/" & Rs!TTV_COD_MODELO & ", PLACA " & Rs!ttv_placa & ", ANO " & _
            Rs!TTV_ANO_FABRICACAO & ", LICENCIAMENTO " & Rs!TTV_LICENCA & ", CHASSI " & Rs!TTV_CHASSI
            txtAtividadeEconomicaVEiculo = Imposto.BuscaNomeCAE(Rs("TTV_TAE_CAE"))
            lbl(9) = "Ativ. Econ. do Veiculo"
            txtDocOrigem = Rs!ttv_placa
        Else
            txtDocOrigem = Rs("TAN_DOC_ORIGEM")
            lbl(9) = "PUBLICIDADE"
            txtAtividadeEconomicaVEiculo = Pega_Nome_Taxa(Trim(Rs("TAN_MOVIMENTO")))
            TxtAtividade = "*****************************************************************************"
        End If
        If txtDocOrigem = "" Then
            txtAtividadeEconomicaVEiculo = "*****************************************************************************"
        End If
        Bdados.FechaTabela Rs
    Else
        Avisa "Contribuinte não cadastrado."
        LimpaCampos Me
        txtcgc.SetFocus
    End If
    Bdados.FechaTabela Rs
    txtObs = observacao
End Sub
Private Function Pega_Nome_Taxa(Cod As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "SELECT * FROM TAB_IMPOSTO WHERE TIP_COD_IMPOSTO = " & Bdados.Converte(Cod, tctexto)
    If Bdados.AbreTabela(Sql, Rs) Then
        Pega_Nome_Taxa = Rs.Fields("TIP_NOME_IMPOSTO")
    End If
End Function

Private Sub txtIm_LostFocus()
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Trim(txtIm) <> "" Then
            If IsNumeric(txtIm) Then txtIm = Imposto.FormataInscricao(txtIm, InscContrib)
             PreencheTela ("tci_im = '" & txtIm & "'")
        End If
    Else
        PreencheTela ("tci_im = '" & txtIm & "'")
    End If
    txtRestricao = Temp.PegaParametro(Bdados, "RESTRICAO PADRAO ALVARA")
    txtObs = observacao
    'txtObs = Temp.PegaParametro(Bdados, "OBSERVACAO PADRAO ALVARA")
End Sub

Private Sub txtValidade_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub Text1_LostFocus()
    
End Sub

Private Sub txtcgc_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    If Me.ActiveControl.Name = "cmdSair" Then Exit Sub
    If Trim(txtcgc) = "" Then Exit Sub
    If Len(txtcgc) = 11 Then
        txtcgc = Edita.FormataTexto(txtcgc, Cpf)
    ElseIf Len(txtcgc) = 14 And Mid(txtcgc, 4, 1) <> "." Then
        txtcgc = Edita.FormataTexto(txtcgc, Cgc)
    End If
    
    PreencheTela ("tci_cgc_cpf = '" & txtcgc & "'")
    txtRestricao = Temp.PegaParametro(Bdados, "RESTRICAO PADRAO ALVARA")
    txtObs = "ESTE ALVARÁ TERÁ VALIDADE PARA O FIM ACIMA MENCIONADO E DEVERÁ SER AFIXADO EM LOCAL VISÍVEL À FISCALIZAÇÃO"
    
    'txtObs = Temp.PegaParametro(Bdados, "OBSERVACAO PADRAO ALVARA")
End Sub

Private Sub txtDAM_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    If Me.ActiveControl.Name = "cmdSair" Then Exit Sub
    If Trim(txtDam) = "" Then Exit Sub
    txtRestricao = Temp.PegaParametro(Bdados, "RESTRICAO PADRAO ALVARA")
    'txtObs = Temp.PegaParametro(Bdados, "OBSERVACAO PADRAO ALVARA")
    txtObs = "ESTE ALVARÁ TERÁ VALIDADE PARA O FIM ACIMA MENCIONADO E DEVERÁ SER AFIXADO EM LOCAL VISÍVEL À FISCALIZAÇÃO"
    
    Sql = " Select * from tab_darm_recebido,Tab_Contribuinte " _
        & " where tdr_tgt_cod_pagamento = " & txtDam & " and tdr_im=tci_im and tdr_sit_pago <> 2"
    
    If Bdados.AbreTabela(Sql, Rs) Then
        txtcgc.TabStop = False
        TxtAtividade = Imposto.BuscaNomeCAE(Rs("tci_grupo_cae"))
        txtcgc = "" & Rs("tci_cgc_cpf")
        txtrazao = "" & Rs("tci_nome")
        txtfantasia = "" & Rs("tci_fantasia")
        Atividade = "" & UCase(TxtAtividade.Text)
        InscMuni = "" & Rs("tci_im")
        RazaoSocial = "" & Rs("tci_nome")
        NomeFantasia = "" & Rs("tci_fantasia")
        CPFCNPJ = "" & Rs("tci_cgc_cpf")
        Endereco = "" & Rs("tci_logradouro") & " " & Rs("tci_nome_logradouro") & "," & Rs("tci_numero") & " " & Rs("tci_complemento") & " " & Rs("tci_bairro")
        Bairro = "" & Rs("tci_bairro")
        Cidade = "" & Rs("tci_cidade")
        Cep = "" & Rs("tci_cep")
        Uf = "" & Rs("tci_uf")
        cae = "" & Rs("tci_tae_cae")
        txtValidade = "31/12/" & Format(Now, "yyyy")
        txtMotivo.SetFocus
    Else
        Call Util.Informa("DAM com falta de pagamento ou baixa no sistema.")
        txtDam.SetFocus
    End If
    Bdados.FechaTabela Rs
End Sub


Private Sub txtValidade_LostFocus()
    txtValidade = Edita.FormataTexto(txtValidade, data)
End Sub

