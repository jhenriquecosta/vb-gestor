VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCOB205 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleMode       =   0  'User
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   36
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCOB205.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   35
      Top             =   4470
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   13
         Top             =   105
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
         Left            =   4095
         TabIndex        =   14
         Top             =   105
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         Caption         =   "&Emitir DAM"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   5580
         TabIndex        =   15
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2160
      Left            =   60
      TabIndex        =   20
      Top             =   705
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   3810
      Altura          =   1905
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VB.TextBox txtCodReceita 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5130
         TabIndex        =   1
         Tag             =   "Codigo Tributo"
         Top             =   570
         Width           =   1485
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
         Left            =   5130
         TabIndex        =   3
         Tag             =   "Parcela"
         Top             =   1140
         Width           =   945
      End
      Begin VB.TextBox txtDtVenc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5130
         TabIndex        =   6
         Tag             =   "Data Vencimento"
         Top             =   1740
         Width           =   1335
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
         Left            =   3225
         TabIndex        =   5
         Tag             =   "Exercicio"
         Top             =   1740
         Width           =   1065
      End
      Begin VB.TextBox TXTIC 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Tag             =   "Insc. Cadastral"
         Top             =   1740
         Width           =   2235
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
         Left            =   150
         TabIndex        =   2
         Tag             =   "IM"
         Top             =   1140
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
         Left            =   150
         MaxLength       =   14
         TabIndex        =   0
         Tag             =   "NO. DAM"
         Top             =   570
         Width           =   2235
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   6
         Left            =   150
         TabIndex        =   28
         Top             =   1500
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
         Caption         =   "Insc. Cadastral"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   7
         Left            =   150
         TabIndex        =   27
         Top             =   345
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
         Caption         =   "Nº DAM"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   17
         Left            =   5130
         TabIndex        =   26
         Top             =   915
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
         Caption         =   "Parcela"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   1
         Left            =   5130
         TabIndex        =   25
         Top             =   1500
         Width           =   1440
         _ExtentX        =   2540
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
         Caption         =   "Data Vencimento"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   3255
         TabIndex        =   24
         Top             =   1500
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
         Caption         =   "Exercício"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   9
         Left            =   150
         TabIndex        =   23
         Top             =   915
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
         Caption         =   "Insc. Municipal"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   5130
         TabIndex        =   22
         Top             =   345
         Width           =   1365
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
         Caption         =   "Cód. Tributo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   960
      Width           =   375
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
      Left            =   660
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1230
      Width           =   5895
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4560
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   18
      Top             =   780
      Visible         =   0   'False
      Width           =   795
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   1138
      Icone           =   "TCOB205.frx":2123
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1500
      Left            =   75
      TabIndex        =   21
      Top             =   2925
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   2646
      Altura          =   1905
      Caption         =   " Detalhes"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   16
         Left            =   435
         TabIndex        =   34
         Top             =   742
         Width           =   1560
         _ExtentX        =   2752
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
         Caption         =   "Taxas Acessórias"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   15
         Left            =   3855
         TabIndex        =   33
         Top             =   1117
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
         Caption         =   "Total a Recolher"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   14
         Left            =   4665
         TabIndex        =   32
         Top             =   742
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
         Caption         =   "Juros"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   13
         Left            =   315
         TabIndex        =   31
         Top             =   1117
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
         Caption         =   "Imposto a Recolher"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   10
         Left            =   615
         TabIndex        =   30
         Top             =   367
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
         Caption         =   "Base de Cálculo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   18
         Left            =   4695
         TabIndex        =   29
         Top             =   367
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
         Caption         =   "Multa"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
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
         Left            =   5280
         TabIndex        =   10
         Tag             =   "Multa"
         Top             =   345
         Width           =   1275
      End
      Begin VB.TextBox txtSaldo 
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
         Left            =   2040
         TabIndex        =   7
         Tag             =   "Base Calculo"
         Top             =   345
         Width           =   1275
      End
      Begin VB.TextBox txtImposto 
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
         Left            =   2040
         TabIndex        =   9
         Tag             =   "Imposto"
         Top             =   1095
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
         Left            =   5280
         TabIndex        =   11
         Tag             =   "Juros"
         Top             =   720
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
         Left            =   5280
         TabIndex        =   12
         Tag             =   "Total"
         Top             =   1095
         Width           =   1275
      End
      Begin VB.TextBox txtTaxas 
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
         Left            =   2040
         TabIndex        =   8
         Tag             =   "Taxas"
         Top             =   720
         Width           =   1275
      End
   End
End
Attribute VB_Name = "TCOB205"
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
Dim CodPagamento As Double

Private Sub cmd_Click(Index As Integer)
    Dim a As Integer
    Dim Valores As String
    Dim Campos As String
    Dim ValorImposto As Double
    Dim RsCob As VSRecordset
    Dim rs As VSRecordset
    Dim Sql As String
    Dim SqlParc As String
    Dim Cobranca As New VSCobranca
    
    CodPagamento = Nvl(txtDam, 0)
    Data_Vencimento = txtDtVenc
    Cod_Tributo = txtCodReceita
    InscMuni = txtIM
    Screen.MousePointer = 11
    
    Select Case cmd(Index).Caption
        
        Case "&Emitir DAM"
            If Not Edita.CriticaCampos(Me) Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            Conta.GeraPagamento txtIM, txtic, txtCodReceita, CLng(IIf(IsNumeric(txtPeriodo), txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2))), txtDtVenc, CDbl(CDbl(txtImposto)), CDbl(Nvl(txtMulta, 0)), CDbl(Nvl(txtJuros, 0)), CDbl(txtDam), 0, CInt(Nvl(txtParcela, 0)), CDbl(txtTaxas), , IIf(txtParcela = 0, 1, 3)
            Conta.CriaContaContribuinte txtDam
            Informa "DAM gerado."
            LimpaCampos Me
        Case "Sai&r"
           Unload Me
    End Select
    Screen.MousePointer = 0
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub cmdLimpar_Click(Index As Integer)
    Edita.LimpaCampos Me
    txtDam.SetFocus
End Sub

Private Sub Form_Load()
            
    Dim Controle As Control
    Dim i As Byte
    
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub


Private Sub txtContribuinte_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtCodReceita_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtCodReceita) = "" Then Exit Sub
    Sql = "Select tip_nome_imposto from tab_imposto where tip_cod_imposto ='" & txtCodReceita & "'"
    If Not Bdados.AbreTabela(Sql, rs) Then
        Avisa "Código de receita inválido."
        txtCodReceita.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtDAM_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDAM_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim i As Byte
    If Trim(txtDam) = "" Then Exit Sub
    Sql = "Select tgt_im from tab_geracao_tributo where tgt_cod_pagamento =" & txtDam
    If Bdados.AbreTabela(Sql, rs) Then
        Informa "DAM já existente. Confirme o número."
        txtDam.SetFocus
        Bdados.FechaTabela rs
        Exit Sub
    Else
        Sql = "Select tgt_im from tab_geracao_tributo_parcela where tgt_cod_pagamento =" & txtDam
        If Bdados.AbreTabela(Sql, rs) Then
            Informa "DAM já existente. Confirme o número."
            txtDam.SetFocus
            Bdados.FechaTabela rs
            Exit Sub
        End If
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtDtVenc_LostFocus()
    txtDtVenc = Edita.FormataTexto(txtDtVenc, Data)
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    On Error GoTo trata
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIM = Imposto.FormataInscricao(txtIM, InscContrib)
    End If
trata:
    If Err.Number = 3265 Then
        Resume Next
    End If
End Sub

Private Sub txtImposto_LostFocus()
    txtImposto = Format(txtImposto, Const_Monetario)
End Sub

Private Sub txtJuros_LostFocus()
    txtJuros = Format(txtJuros, Const_Monetario)
End Sub

Private Sub txtMulta_LostFocus()
    txtMulta = Format(txtMulta, Const_Monetario)
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtSaldo_LostFocus()
    txtSaldo = Format(txtSaldo, Const_Monetario)
End Sub

Private Sub txtTotalNotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtTaxas_LostFocus()
    txtTaxas = Format(txtTaxas, Const_Monetario)
End Sub

Private Sub txtTotalImposto_LostFocus()
        txtTotalImposto = Format(txtTotalImposto, Const_Monetario)
End Sub
