VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCOB202 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCOB202"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   37
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCOB202.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   2025
      Left            =   90
      TabIndex        =   17
      Top             =   3540
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   3572
      Altura          =   1905
      Caption         =   " Documentos"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483637
      Ocultavel       =   0   'False
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
         TabIndex        =   19
         Tag             =   "Validade"
         Top             =   390
         Width           =   1605
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "TCOB202.frx":2123
         Left            =   7470
         List            =   "TCOB202.frx":2139
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Tag             =   "Tipo"
         Top             =   360
         Width           =   2415
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   6
         Left            =   6930
         TabIndex        =   20
         Top             =   390
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
         Caption         =   "Tipo:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   3
         Left            =   1290
         TabIndex        =   18
         Top             =   450
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
         Caption         =   "Validade:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   12
         Left            =   1080
         TabIndex        =   24
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
         Height          =   270
         Index           =   10
         Left            =   15
         TabIndex        =   22
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
         TabIndex        =   23
         Tag             =   "Finalidade do Alvará"
         Top             =   750
         Width           =   7755
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
         Height          =   795
         Left            =   2130
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   1140
         Width           =   7755
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2760
      Left            =   90
      TabIndex        =   35
      Top             =   735
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4868
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483637
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   11
         Left            =   3435
         TabIndex        =   36
         Top             =   495
         Width           =   3840
         _ExtentX        =   6773
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
         Caption         =   "Placa do Veiculo ou Doc Origem dos Anuncios"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
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
         TabIndex        =   16
         Top             =   2265
         Width           =   7725
      End
      Begin Threed.SSPanel lbl 
         Height          =   150
         Index           =   9
         Left            =   15
         TabIndex        =   15
         Top             =   2355
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   265
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
         Caption         =   "Ativ. Econ. do Veiculo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
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
         TabIndex        =   1
         Top             =   435
         Width           =   1305
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   7
         Left            =   960
         TabIndex        =   0
         Top             =   480
         Width           =   990
         _ExtentX        =   1746
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
         Caption         =   "Doc.Origem:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
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
         Left            =   8190
         MaxLength       =   18
         TabIndex        =   8
         Tag             =   "CPF ou  CGC:"
         Top             =   810
         Width           =   1605
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
         Left            =   4650
         MaxLength       =   18
         TabIndex        =   5
         Tag             =   "CPF ou  CGC:"
         Top             =   810
         Width           =   1605
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
         TabIndex        =   12
         Top             =   1530
         Width           =   7725
      End
      Begin VTOcx.cmdVISUAL cmdPesq 
         Height          =   375
         Index           =   1
         Left            =   6330
         TabIndex        =   6
         Top             =   765
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   4
         Left            =   7860
         TabIndex        =   7
         Top             =   840
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
         Caption         =   "IM.:"
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
         TabIndex        =   13
         Top             =   1980
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   265
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
         Caption         =   "Atividade Economica:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   2
         Left            =   675
         TabIndex        =   2
         Top             =   855
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
         Caption         =   "NO. Pagamento:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   5
         Left            =   3480
         TabIndex        =   4
         Top             =   855
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
         Caption         =   "CPF ou CNPJ:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   0
         Left            =   465
         TabIndex        =   11
         Top             =   1590
         Width           =   1710
         _ExtentX        =   3016
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
         Caption         =   "Nome de Fantasia:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   8
         Left            =   150
         TabIndex        =   9
         Top             =   1230
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
         Caption         =   "Nome ou Razão Social:"
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
         TabIndex        =   10
         Tag             =   "Nome ou Razão Social:"
         Top             =   1170
         Width           =   7710
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
         Left            =   2100
         MaxLength       =   14
         TabIndex        =   3
         Top             =   810
         Width           =   1305
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
         TabIndex        =   14
         Top             =   1890
         Width           =   7725
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   31
      Top             =   5640
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdLimpa 
         Height          =   375
         Left            =   6615
         TabIndex        =   34
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   7785
         TabIndex        =   33
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8955
         TabIndex        =   32
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   345
      Left            =   3780
      TabIndex        =   26
      Top             =   1440
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7440
      Top             =   1200
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1138
      Icone           =   "TCOB202.frx":2191
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   2
      Left            =   2340
      TabIndex        =   28
      Top             =   0
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
      Index           =   0
      Left            =   2340
      TabIndex        =   29
      Top             =   0
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdVISUAL1 
      Height          =   375
      Left            =   1170
      TabIndex        =   30
      Top             =   0
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Imprimir"
      Acao            =   4
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCOB202"
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
Dim Bairro As String
Dim Cidade As String
Dim Cep As String
Dim Uf As String
Dim Atividade_Veiculo As String
Public Carro As Boolean


Public Sub PreencheTela(Criterio As String, Optional Conexao As Object)
    Dim rs As VSRecordset
    Dim Sql As String
        
    Sql = " Select * from Tab_Contribuinte " _
        & " where " & Criterio
    
    If Not Conexao Is Nothing Then Set Bdados = Conexao
    LimpaCampos Me
    If Bdados.AbreTabela(Sql, rs) Then
        TxtAtividade = Imposto.BuscaNomeCAE("" & rs("tci_tae_cae"))
        txtIm = rs!TCI_IM
        txtcgc = "" & rs("tci_cgc_cpf")
        txtrazao = "" & rs("tci_nome")
        txtfantasia = "" & rs("tci_fantasia")
        Atividade = "" & UCase(TxtAtividade.Text)
        InscMuni = "" & rs("tci_im")
        RazaoSocial = "" & rs("tci_nome")
        NomeFantasia = "" & rs("tci_fantasia")
        CPFCNPJ = "" & rs("tci_cgc_cpf")
        Endereco = "" & rs("tci_logradouro") & " " & rs("tci_nome_logradouro") & "," & rs("tci_numero") & " " & rs("tci_complemento")
        Bairro = "" & rs("tci_bairro")
        Cidade = "" & rs("tci_cidade")
        Cep = "" & rs("tci_cep")
        Uf = "" & rs("tci_uf")
        txtValidade = "31/12/" & Format(Now, "yyyy")
        Sql = "Select TTV_TAE_CAE, TTV_VEICULO,TTV_MARCA,TTV_COD_MODELO,TTV_ANO_FABRICACAO,TTV_PLACA,TTV_PLACA," & _
            "TTV_MUNICIPIO,TTV_UF,TTV_LICENCA,TTV_CHASSI FROM TAB_TRANSPORTADOR_VEICULO WHERE TTV_TCI_IM='" & txtIm & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtRestricao = "" & rs!TTV_MARCA & "/" & rs!TTV_COD_MODELO & ", PLACA " & rs!ttv_placa & ", ANO " & _
            rs!TTV_ANO_FABRICACAO & ", LICENCIAMENTO " & rs!TTV_LICENCA & ", CHASSI " & rs!TTV_CHASSI
            
        Else
            txtRestricao = ""
        End If
        If txtDocOrigem = "" Then
            txtAtividadeEconomicaVEiculo = "*****************************************************************************"
        End If
        Bdados.FechaTabela rs
    Else
        Avisa "Contribuinte não cadastrado."
        txtcgc.SetFocus
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub cboTipo_Click()
    If cboTipo.ListIndex = 2 Then
        txtMotivo = "ALVARÁ DE LICENÇA E FUNCIONAMENTO"
    ElseIf cboTipo.ListIndex = 3 Then
        txtMotivo = "ALVARÁ DE LICENÇA DE FUNCIONAMENTO E LOCALIZAÇÃO"
    End If
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImprimir_Click()
    Dim a As Integer
    Dim CodAlvara As String
    Dim Valores As String
    Dim Campos As String
    
    On Error Resume Next

    CodAlvara = BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto = '" & Imposto.NomeTributo(ttr_ALVARA) & " '")
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    Restricoes = "" & UCase(txtRestricao.Text)
    Finalidade = "" & UCase(txtMotivo)
    
    Bdados.DeletaDados "TAB_ALVARA_IMPRESSO", "tai_tci_im='" & txtIm & "' AND tai_periodo =" & Year(Date)
    Valores = Bdados.PreparaValor(txtIm, Year(Date), Bdados.Converte(Date, TCDataHora), Aplicacoes.Usuario, Bdados.Converte(txtValidade, TCDataHora))
    Campos = "tai_tci_im,tai_periodo,tai_data_impressao,tai_tus_cod_usuario,tai_data_validade"
    Bdados.GravaDados "TAB_ALVARA_IMPRESSO", Valores, Campos, "tai_tci_im='" & txtIm & "' AND tai_periodo =" & Year(Date)
    Select Case cboTipo.ListIndex
        Case 0 'PADRAO
            With Rpt
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

                .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
                .Formulas "CIDADE", Aplicacoes.Municipio
                .Formulas "DEPARTAMENTO", Temp.PegaParametro(Bdados, "SETOR")
                .Titulo = Imposto.NomeTributo(ttr_ALVARA)
                .Arvore = False
                .Visualizar
            End With
        Case 1 'MOTO-TAXI
            With Rpt
                If Not .DefinirArquivo(Bdados, App.Path + "\TAlvaraMotoTaxi.rpt") Then Exit Sub
                .Selecao = "{Tab_Contribuinte.tci_im} = '" & txtIm & "'"
                .Formulas "VTFinalidade", txtMotivo
                .Formulas "VT_RESTRICAO", txtRestricao
                .Titulo = Imposto.NomeTributo(ttr_ALVARA)
                .Arvore = False
                .Visualizar
            End With
        Case 2, 3 'FUNCIONAMENTO/LOCALIZACAO
            Dim Arquivo As String
            If cboTipo.ListIndex = 2 Then
                Arquivo = "\TAlvaraFuncionamento.rpt"
            Else
                Arquivo = "\TAlvaraLocalizacao.rpt"
            End If
            With Rpt
                If Not .DefinirArquivo(Bdados, App.Path & Arquivo) Then Exit Sub
                .Formulas "VT_Restricao", Restricoes
                .Formulas "RAZAOSOCIAL", RazaoSocial
                .Formulas "NOMEFANTASIA", NomeFantasia
                .Formulas "CPF/CNPJ", CPFCNPJ
                .Formulas "IM", InscMuni
                .Formulas "ENDERECO", Trim$(Endereco) & " - " & Bairro
                If txtDocOrigem = "" Then
                   .Formulas "ATIVIDADE", Atividade
                Else
                    .Formulas "ATIVIDADE", txtAtividadeEconomicaVEiculo
                End If
                .Formulas "VALIDADE", txtValidade
                .Formulas "CIDADE", Aplicacoes.Municipio
                If cboTipo.ListIndex = 2 Then
                    .Formulas "VT_RESTRICOES", Restricoes
                End If
                .Titulo = Imposto.NomeTributo(ttr_ALVARA)
                .Arvore = False
                .Visualizar
            End With
        Case 4 'MOTO-TAXI
            With Rpt
                If Not .DefinirArquivo(Bdados, App.Path + "\TCrachaMotoTaxi.rpt") Then Exit Sub
                .Selecao = "{Tab_Contribuinte.tci_im} = '" & txtIm & "' AND {TAB_ALVARA_IMPRESSO.TAI_PERIODO} =" & Year(Date)
                .Formulas "VTFinalidade", txtMotivo
                .Formulas "VTFinalidade", txtMotivo
                .Titulo = Imposto.NomeTributo(ttr_ALVARA)
                .Arvore = False
                .Visualizar
            End With
        Case 5 'ALVARA DE LICENSA
            With Rpt
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
                .Formulas "CIDADE", Aplicacoes.Municipio
                '.Formulas "DEPARTAMENTO", Temp.PegaParametro(Bdados, "SETOR")
                '.Titulo = Imposto.NomeTributo(ttr_ALVARA)
                .Arvore = False
                .Imprimir
             End With
    End Select
    Screen.MousePointer = 0
    Set Rpt = Nothing
    Call Util.Informa("ALVARÁ emitido.")
    If Me.Tag = "EXTERNO" Then cmdSair_Click
End Sub

Private Sub cmdLimpa_Click()
    Edita.LimpaCampos Me
    txtDam.SetFocus
End Sub

Private Sub cmdPesq_Click(Index As Integer)
    AplicacoesVTFuncoes.BuscaNoEconomico TcoJuridica, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtDocOrigem_LostFocus()
    Dim rs As VSRecordset
    Dim Sql As String
    If txtDocOrigem = "" Then Exit Sub
     
    'Checo se existe carros para o doc de origem...
    
     Sql = " Select * from Tab_Contribuinte ,TAB_TRANSPORTADOR_VEICULO" _
        & " where TTV_TCI_IM = TCI_IM AND  TTV_PLACA = " & Bdados.Converte(txtDocOrigem, tctexto)
         Carro = True
    If Not Bdados.AbreTabela(Sql) Then
            Carro = False
            Sql = "SELECT * FROM TAB_CONTRIBUINTE,TAB_ANUNCIO"
            Sql = Sql & " Where TAN_TCI_IM = TCI_IM"
            Sql = Sql & " AND TAN_DOC_ORIGEM = " & Bdados.Converte(txtDocOrigem, tctexto)
    End If
    If Bdados.AbreTabela(Sql, rs) Then
        LimpaCampos Me
        'TxtAtividade = Imposto.BuscaNomeCAE(Rs("tci_tae_cae"))
        txtRestricao = ""
        txtIm = rs!TCI_IM
        txtcgc = "" & rs("tci_cgc_cpf")
        txtrazao = "" & rs("tci_nome")
        txtfantasia = "" & rs("tci_fantasia")
        Atividade = "" & UCase(TxtAtividade.Text)
        InscMuni = "" & rs("tci_im")
        RazaoSocial = "" & rs("tci_nome")
        NomeFantasia = "" & rs("tci_fantasia")
        CPFCNPJ = "" & rs("tci_cgc_cpf")
        Endereco = "" & rs("tci_logradouro") & " " & rs("tci_nome_logradouro") & "," & rs("tci_numero") & " " & rs("tci_complemento")
        Bairro = "" & rs("tci_bairro")
        Cidade = "" & rs("tci_cidade")
        Cep = "" & rs("tci_cep")
        Uf = "" & rs("tci_uf")
        txtValidade = "31/12/" & Format(Now, "yyyy")
        If Carro Then
            txtRestricao = "" & rs!TTV_MARCA & "/" & rs!TTV_COD_MODELO & ", PLACA " & rs!ttv_placa & ", ANO " & _
            rs!TTV_ANO_FABRICACAO & ", LICENCIAMENTO " & rs!TTV_LICENCA & ", CHASSI " & rs!TTV_CHASSI
            txtAtividadeEconomicaVEiculo = Imposto.BuscaNomeCAE(rs("TTV_TAE_CAE"))
            lbl(9) = "Ativ. Econ. do Veiculo"
            txtDocOrigem = rs!ttv_placa
        Else
            txtDocOrigem = rs("TAN_DOC_ORIGEM")
            lbl(9) = "PUBLICIDADE"
            txtAtividadeEconomicaVEiculo = Pega_Nome_Taxa(Trim(rs("TAN_MOVIMENTO")))
            TxtAtividade = "*****************************************************************************"
        End If
        If txtDocOrigem = "" Then
            txtAtividadeEconomicaVEiculo = "*****************************************************************************"
        End If
        Bdados.FechaTabela rs
    Else
        Avisa "Contribuinte não cadastrado."
        LimpaCampos Me
        txtcgc.SetFocus
    End If
    Bdados.FechaTabela rs
End Sub
Private Function Pega_Nome_Taxa(Cod As String) As String
    Dim Sql As String
    Dim rs As VSRecordset
    
    Sql = "SELECT * FROM TAB_IMPOSTO WHERE TIP_COD_IMPOSTO = " & Bdados.Converte(Cod, tctexto)
    If Bdados.AbreTabela(Sql, rs) Then
        Pega_Nome_Taxa = rs.Fields("TIP_NOME_IMPOSTO")
    End If
End Function
Private Sub txtIm_LostFocus()
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        If Trim(txtIm) <> "" Then
            If IsNumeric(txtIm) Then txtIm = Imposto.FormataInscricao(txtIm, InscContrib)
             PreencheTela ("tci_im = '" & txtIm & "'")
        End If
    Else
        PreencheTela ("tci_im = '" & txtIm & "'")
    End If
End Sub

Private Sub txtValidade_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub Text1_LostFocus()
    
End Sub

Private Sub txtcgc_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Me.ActiveControl.Name = "cmdSair" Then Exit Sub
    If Trim(txtcgc) = "" Then Exit Sub
    If Len(txtcgc) = 11 Then
        txtcgc = Edita.FormataTexto(txtcgc, Cpf)
    ElseIf Len(txtcgc) = 14 And Mid(txtcgc, 4, 1) <> "." Then
        txtcgc = Edita.FormataTexto(txtcgc, Cgc)
    End If
    
    PreencheTela ("tci_cgc_cpf = '" & txtcgc & "'")

End Sub

Private Sub txtDAM_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Me.ActiveControl.Name = "cmdSair" Then Exit Sub
    If Trim(txtDam) = "" Then Exit Sub
    Sql = " Select * from tab_darm_recebido,Tab_Contribuinte " _
        & " where tdr_tgt_cod_pagamento = " & txtDam & " and tdr_im=tci_im and tdr_sit_pago <> 2"
    
    If Bdados.AbreTabela(Sql, rs) Then
        txtcgc.TabStop = False
        TxtAtividade = Imposto.BuscaNomeCAE(rs("tci_grupo_cae"))
        txtcgc = "" & rs("tci_cgc_cpf")
        txtrazao = "" & rs("tci_nome")
        txtfantasia = "" & rs("tci_fantasia")
        Atividade = "" & UCase(TxtAtividade.Text)
        InscMuni = "" & rs("tci_im")
        RazaoSocial = "" & rs("tci_nome")
        NomeFantasia = "" & rs("tci_fantasia")
        CPFCNPJ = "" & rs("tci_cgc_cpf")
        Endereco = "" & rs("tci_logradouro") & " " & rs("tci_nome_logradouro") & "," & rs("tci_numero") & " " & rs("tci_complemento") & " " & rs("tci_bairro")
        Bairro = "" & rs("tci_bairro")
        Cidade = "" & rs("tci_cidade")
        Cep = "" & rs("tci_cep")
        Uf = "" & rs("tci_uf")
        txtValidade = "31/12/" & Format(Now, "yyyy")
        txtMotivo.SetFocus
    Else
        Call Util.Informa("DAM com falta de pagamento ou baixa no sistema.")
        txtDam.SetFocus
    End If
    Bdados.FechaTabela rs
End Sub


Private Sub txtValidade_LostFocus()
    txtValidade = Edita.FormataTexto(txtValidade, Data)
End Sub
