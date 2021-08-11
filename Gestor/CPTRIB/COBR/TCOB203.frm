VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TCOB203 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleMode       =   0  'User
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   48
      Top             =   6525
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   5610
         TabIndex        =   20
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Novo"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL Cmd 
         Height          =   375
         Index           =   1
         Left            =   4170
         TabIndex        =   19
         Top             =   105
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         Caption         =   "&Emitir DAM"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL Cmd 
         Height          =   375
         Index           =   2
         Left            =   6780
         TabIndex        =   21
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   43
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCOB203.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.cmdVISUAL CmdIc 
      Height          =   300
      Left            =   3465
      TabIndex        =   42
      Top             =   2250
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   529
      Caption         =   ""
      Acao            =   5
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
      Left            =   1530
      TabIndex        =   10
      Tag             =   "Exercicio"
      Top             =   3450
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
      Left            =   6540
      TabIndex        =   11
      Tag             =   "Data Vencimento"
      Top             =   3435
      Width           =   1335
   End
   Begin VB.TextBox txtImovel 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   2250
      Width           =   4080
   End
   Begin VB.TextBox txtIc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1530
      TabIndex        =   6
      Top             =   2250
      Width           =   1935
   End
   Begin VB.ComboBox cboImposto 
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
      ItemData        =   "TCOB203.frx":2123
      Left            =   720
      List            =   "TCOB203.frx":2125
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "Imposto"
      Top             =   660
      Width           =   7245
   End
   Begin VB.Timer tmr 
      Interval        =   10
      Left            =   8160
      Top             =   1710
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   1138
      Icone           =   "TCOB203.frx":2127
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   8250
      ScaleHeight     =   375
      ScaleWidth      =   765
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   795
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   270
      Width           =   5895
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1200
      TabIndex        =   23
      Top             =   -30
      Width           =   375
   End
   Begin Threed.SSPanel lbl 
      Height          =   270
      Index           =   5
      Left            =   60
      TabIndex        =   27
      Top             =   690
      Width           =   765
      _ExtentX        =   1349
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
      Caption         =   "Tributo:"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSFrame fra 
      Height          =   1215
      Index           =   3
      Left            =   60
      TabIndex        =   28
      Top             =   990
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   2143
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   16384
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Adquirente"
      ShadowStyle     =   1
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
         TabIndex        =   1
         Top             =   180
         Width           =   1485
      End
      Begin VB.TextBox txtCgc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox txtCedente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   4
         Tag             =   "Adquirente"
         Top             =   510
         Width           =   6345
      End
      Begin VB.TextBox txtEnderecoCedente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   5
         Tag             =   "Endereço Adquirente"
         Top             =   840
         Width           =   6345
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   4
         Left            =   5130
         TabIndex        =   29
         Top             =   225
         Width           =   900
         _ExtentX        =   1588
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
         Caption         =   "CNPJ/CPF:"
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
         TabIndex        =   30
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
         Caption         =   "Insc. Municipal:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   12
         Left            =   330
         TabIndex        =   31
         Top             =   540
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
         Caption         =   "Contribuinte:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   570
         TabIndex        =   32
         Top             =   870
         Width           =   1020
         _ExtentX        =   1799
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
         Caption         =   "Endereço:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesq 
         Height          =   315
         Left            =   2970
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
   End
   Begin Threed.SSPanel lbl 
      Height          =   180
      Index           =   6
      Left            =   60
      TabIndex        =   33
      Top             =   2280
      Width           =   1500
      _ExtentX        =   2646
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
      Caption         =   "Insc. Cadastral:"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSFrame fra 
      Height          =   855
      Index           =   2
      Left            =   60
      TabIndex        =   34
      Top             =   2550
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   1508
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   16384
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cedente"
      ShadowStyle     =   1
      Begin VB.TextBox txtEnderecoAdquirente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   9
         Top             =   480
         Width           =   6345
      End
      Begin VB.TextBox txtAdquirente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   8
         Top             =   150
         Width           =   6345
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   7
         Left            =   300
         TabIndex        =   35
         Top             =   210
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
         Caption         =   "Contribuinte:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   8
         Left            =   570
         TabIndex        =   36
         Top             =   540
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
         Caption         =   "Endereço:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSPanel lbl 
      Height          =   270
      Index           =   0
      Left            =   660
      TabIndex        =   37
      Top             =   3480
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
      Caption         =   "Exercício:"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel lbl 
      Height          =   270
      Index           =   1
      Left            =   4980
      TabIndex        =   38
      Top             =   3480
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
      Caption         =   "Data Vencimento:"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSFrame fra 
      Height          =   705
      Index           =   1
      Left            =   90
      TabIndex        =   39
      Top             =   5805
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   1244
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   16384
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Detalhes - DESCRIÇÃO DO IMÓVEL - LOTE - CASA - APTO - QUADRA - GLEBA - ÁREA"
      ShadowStyle     =   1
      Begin VB.TextBox txtObservacao 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   7695
      End
   End
   Begin Threed.SSFrame fraValor 
      Height          =   1350
      Left            =   75
      TabIndex        =   40
      Top             =   4455
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   2381
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   16384
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Base de Cálculo"
      ShadowStyle     =   1
      Begin VTOcx.txtVISUAL txtValorAvista 
         Height          =   330
         Left            =   360
         TabIndex        =   14
         Top             =   195
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   582
         Caption         =   "Valor à vista:"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtValorFinanciado 
         Height          =   330
         Left            =   60
         TabIndex        =   16
         Top             =   555
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   582
         Caption         =   "Valor financiado:"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtValorImovel 
         Height          =   330
         Left            =   150
         TabIndex        =   22
         Top             =   915
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         Caption         =   "Valor do Imóvel"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtAliquotaPropria 
         Height          =   330
         Left            =   2640
         TabIndex        =   15
         Top             =   195
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   582
         Caption         =   "Recursos Próprios (Alíquota%)"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtAliquotaFinanciada 
         Height          =   330
         Left            =   2775
         TabIndex        =   17
         Top             =   555
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   582
         Caption         =   "Parte Financiada (Alíquota%)"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtSubTotalaliquotaPropria 
         Height          =   330
         Left            =   6555
         TabIndex        =   44
         Top             =   195
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         Caption         =   ""
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtSubTotalAliquotaFinanciada 
         Height          =   330
         Left            =   6555
         TabIndex        =   45
         Top             =   555
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         Caption         =   ""
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtValorITBI 
         Height          =   330
         Left            =   5355
         TabIndex        =   46
         Top             =   915
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   582
         Caption         =   "Valor do ITBI"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   675
      Left            =   75
      TabIndex        =   41
      Top             =   3780
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   1191
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   16384
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Detalhes do Lançamento"
      ShadowStyle     =   1
      Begin VTOcx.txtVISUAL txtOcupa 
         Height          =   330
         Left            =   90
         TabIndex        =   12
         Top             =   240
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   582
         Caption         =   "Ocupação"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtDestino 
         Height          =   330
         Left            =   4140
         TabIndex        =   13
         Top             =   240
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         Caption         =   "Destinação"
         Text            =   ""
      End
   End
   Begin VTOcx.grdVISUAL GrdTaxas 
      Height          =   6135
      Left            =   7965
      TabIndex        =   47
      Top             =   660
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   10821
      Caption         =   "Taxas"
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
   End
End
Attribute VB_Name = "TCOB203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim CodImposto As String
Dim Sigla As String
Dim Exercicio As String
Dim Conta As New ContaCorrente
Dim Titular As String
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
Dim ValorMetro As Double
Dim TaxaParcela As Double
Dim Desconto As String
Dim Reducao As String
Dim CodPagamento As String
Dim CodTaxa As String
Dim Transf As New TransfImovel
Dim ImContribuinte As String
Dim String_Taxas As String
Dim Total_Taxas As Double
Function BuscaValorTaxa(NomeTaxa As String) As Double
    Dim rs As VSRecordset
    Dim Sql As String
    
    Sql = " select tpi_tip_cod_imposto, tpi_valor_taxa_fixa from Tab_Parametro_Imposto " & _
    " WHERE  tpi_tip_cod_imposto = (SELECT TIP_COD_IMPOSTO FROM TAB_IMPOSTO WHERE TIP_NOME_IMPOSTO ='" & NomeTaxa & "')"
    If Bdados.AbreTabela(Sql, rs) Then
        BuscaValorTaxa = IIf(IsNull(rs!tpi_valor_taxa_fixa), 0, rs!tpi_valor_taxa_fixa)
        CodTaxa = "" & rs!tpi_tip_cod_imposto
    Else
        BuscaValorTaxa = 0
    End If
    Bdados.FechaTabela rs
End Function
Public Sub GeraDam(Aliquota As Double)
    On Error GoTo trata
    Dim rs As VSRecordset
    Dim Sql As String
    Dim MaxCotas As Byte
    Dim Cobranca As New VSCobranca
    Dim CodBarra As New CodigoDeBarra
    Dim LinhaDigitavel As String
    Dim a As Byte
    
    Dim CgcPref As String
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TDAM_ITBI_Barra.rpt") Then Exit Sub
         .Formulas "InscMunicipal", InscMuni
         CgcPref = Temp.PegaParametro(Bdados, "CGC CLIENTE")
         .Formulas "CgcPrefeitura", "CNPJ " & Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
         .Formulas "nome", txtCgc & " - " & txtCedente
         .Formulas "documento ", CStr(CodPagamento)
         .Formulas "endereco", Edita.TiraPic(txtEnderecoCedente, "'")
         .Formulas "datavencimento", Data_Vencimento
         .Formulas "nomecedente", Edita.TiraPic(txtAdquirente, "'")
         .Formulas "enderecocedente", Edita.TiraPic(txtEnderecoAdquirente, "'")
         .Formulas "localizacao", txtIc & " - " & txtImovel
         .Formulas "exercicio", txtPeriodo
         .Formulas "ValorTributo", Format((CDbl(Tributo) * (100 + Nvl(Desconto, 0)) / 100), Const_Monetario)
         
         .Formulas "ValorMulta", Format(Multa, Const_Monetario)
         .Formulas "ValorJuros", Format(Juro, Const_Monetario)
         .Formulas "TaxaExpediente", Format(TaxaServico, Const_Monetario)
         .Formulas "ValorTotal", Format(CDbl(TotalImposto) + CDbl(TaxaServico), Const_Monetario)
         .Formulas "CodigoTributo", CodImposto
         .Formulas "NUM_NOTAS", "Valor Venal do Imóvel(R$): " & Format(txtValorImovel, Const_Monetario) & "     -       Aliquota(%): " & Format(Transf.AliquotaProprio, Const_Monetario)
         .Formulas "BASECALCULO", Format(txtValorImovel, Const_Monetario)
         
          Dim PosTraco As Byte
         .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
         .Formulas "EMISSAO", Imposto.BuscaDataGeracaoDam(CodPagamento)
         .Formulas "Imposto", Edita.TiraPic(UCase(cboImposto.Text), "'")
'         .Formulas "LinhaDigitavel", Cobranca.GeraCodBarra(CodPagamento, CodImposto, CDbl(TotalImposto) + CDbl(TaxaServico), PicBarra, Right(Exercicio, 4) & Left(Exercicio, 2), Data_Vencimento, 0, 1)
          If AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
            LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(InscMuni, Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITBI)), CDbl(TotalImposto) + CDbl(TaxaServico), Right(Exercicio, 4) & Left(Exercicio, 2), PicBarra, Data_Vencimento, 0, CStr(CodPagamento))
          Else
            LinhaDigitavel = CodBarra.CriaLinhaDigitavel(InscMuni, Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITBI)), CDbl(TotalImposto) + CDbl(TaxaServico), Right(Exercicio, 4) & Left(Exercicio, 2), Data_Vencimento, 0)
          End If
        .Formulas "LinhaDigitavel", LinhaDigitavel
        .Formulas "VT_LinhaBarra", CodBarra.LinhaBarraGerada
         .Formulas "OBSERVACAO", Edita.TiraPic(Trim(txtObservacao.Text), "'")
         
         '.Connect = Bdados.BDSistema.Connect
         .CopiasDetalhes = 3
         .Titulo = "Documento de Arrecadação Municipal - DAM"
         .Arvore = False
         .Visualizar
    End With
    Exit Sub
trata:
    If Err.Number = 20515 Or Err.Number = 3265 Then
        Rpt.Formulas "OBSERVACAO", ""
        Resume
    End If
    Avisa "O DAM não será gerado."
    Avisa Err.Description
    Exit Sub
End Sub

Sub GeraItbi(Optional BaseDeCalculo As Double)
        Dim Campos As String
        Dim Valores As String
        Dim Sql As String
        Dim rs As VSRecordset
        Dim Novo As Boolean
        Dim Aliquota As Double
        Dim Venc As String
        Dim Obrig As New Obrigacao
        ImContribuinte = txtIM
        If BaseDeCalculo = 0 Then
            Sql = "select tim_valor from tab_imovel where tim_ic='" & txtIc & "'"
            If Bdados.AbreTabela(Sql, rs) Then
                BaseDeCalculo = rs(0)
            End If
        End If
        'If cboTaxa.ListIndex = -1 Then
 '           TaxaServico = 0
        'Else
         '   TaxaServico = BuscaValorTaxa(cboTaxa.Text)
        'End If
        
        Sql = "Select tpi_aliquota,tpi_tip_cod_imposto from tab_parametro_imposto where tpi_tip_cod_imposto = '" & CodImposto & "'"
        If Bdados.AbreTabela(Sql, rs) Then Aliquota = rs(0)
        
        If Sigla = Imposto.NomeTributo(ttr_ITBI) Then
            'TotalImposto = (CDbl(Nvl(Trim(txtValorAvista), 0)) * (Transf.AliquotaProprio / 100))
            'TotalImposto = TotalImposto + (CDbl(Nvl(Trim(txtValorFinanciado), 0)) * (Transf.AliquotaFinanciado / 100))
            TotalImposto = txtValorITBI
        Else
            ' Busca Aliquota
            'sql = "Select tpi_aliquota,tpi_tip_cod_imposto from tab_parametro_imposto where tpi_tip_cod_imposto = '" & CodImposto & "'"
            'If Bdados.AbreTabela(sql, rs) Then
            '    Aliquota = rs(0)
'                TotalImposto = (CDbl(Nvl(Trim(txtValorImovel), 0)) * (Aliquota / 100))
            'End If
        End If
        
        Tributo = TotalImposto
        Valor_Imposto = TotalImposto
        CodPagamento = Conta.GeraCodPagamento(CodImposto)
        Venc = CStr(txtDtVenc)
        'Pega taxas
        Call Pega_taxas
        Valores = Bdados.PreparaValor(ImContribuinte, Bdados.Converte(txtIc, tctexto), CodImposto, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), Bdados.Converte(txtDtVenc, TCDataHora), Bdados.Converte(TotalImposto, TCDuplo), CodPagamento, Bdados.Converte(Imposto.BuscaDataGeracaoDam(CodPagamento), TCDataHora), Bdados.Converte(TaxaServico, TCDuplo))
      ' Conta.GeraPagamento   ImContribuinte, txtIc, CodImposto, Right(txtPeriodo, 4) & Left(txtPeriodo, 2), Venc, CDbl(TotalImposto), 0, 0, CodPagamento, 0, 0, TaxaServico, CodTaxa, EtcCreditoTributario
        Obrig.CriaObrigacao CodImposto, txtPeriodo, txtPeriodo, ImContribuinte, CDbl(TotalImposto), , , Venc, , , , , , , , , , etiContribuinte
        Transf.Gravar CodPagamento, CodImposto, Imposto.BuscaDataGeracaoDam(CodPagamento), txtIM, txtCgc, CDbl(Nvl(Trim(txtValorAvista), 0)), CDbl(Nvl(Trim(txtValorFinanciado), 0))
        Bdados.FechaTabela rs
        GeraDam Aliquota
End Sub

Private Sub cboIC_Click()
    CarregaImovel txtIc
End Sub
Sub CarregaImovel(InscIc As String)
    
End Sub





Private Sub cboImposto_Click()
    Dim Sql As String
    Dim rs As VSRecordset
    
    Sql = "select TIP_cod_IMPOSTO,tip_sigla_imposto from tab_imposto" & _
    " WHERE  TIP_nome_IMPOSTO ='" & cboImposto & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        CodImposto = rs(0)
        Sigla = rs!TIP_sigla_IMPOSTO
        'fraValor.Enabled = Sigla = Imposto.NomeTributo(ttr_ITBI)
    End If
    Bdados.FechaTabela rs
End Sub

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
    
    Documento = CodPagamento
    Data_Vencimento = txtDtVenc
    Cod_Tributo = NomeImposto
    CPFCNPJ = txtCgc
    InscMuni = txtIM
    RazaoSocial = txtCedente
    Juro = 0
    Multa = 0
    TotalImposto = 0
    Linhas = 0
    Screen.MousePointer = 11

    Select Case cmd(Index).Caption
        
        Case "&Emitir DAM"
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Call Pega_taxas
            TaxaServico = Total_Taxas
            'If cboTaxa.ListIndex = -1 Then
             '   TaxaServico = 0
            'Else
                'TaxaServico = BuscaValorTaxa(cboTaxa.Text)
            'End If
                    
            Sql = "Select tgt_valor_tributo,tgt_data_geracao from tab_geracao_tributo where tgt_im='" & txtIM & "' and tgt_tip_cod_imposto='" & CodImposto & "' and tgt_periodo=" & IIf(Len(txtPeriodo) = 4, txtPeriodo, Left(txtPeriodo, 2) & Right(txtPeriodo, 4))
            If Bdados.AbreTabela(Sql, rs) Then
                Avisa ("DAM já emitido no valor de R$" & Format(rs(0), Const_Monetario) & ", com em " & IIf(Not IsNull(rs(1)), rs(1), ""))
                Screen.MousePointer = 0
                Bdados.FechaTabela rs
                Exit Sub
            End If
            GeraItbi txtValorITBI
            Bdados.FechaTabela rs
            Util.Informa "DAM Emitido e Gravaçao efetuada com sucesso."
        Case "Sai&r"
           Unload Me
    End Select
    Screen.MousePointer = 0
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub CmdIc_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtIc
    If txtIM <> "" Then
        txtic_LostFocus
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtSubTotalaliquotaPropria = "0,00"
    txtAliquotaPropria = "0,00"
    txtValorAvista = "0,00"
    txtSubTotalAliquotaFinanciada = "0,00"
    txtAliquotaFinanciada = "0,00"
    txtValorFinanciado = "0,00"
    txtValorITBI = "0,00"
    cboImposto.SetFocus
End Sub



Private Sub cmdPesq_Click()
    'blnConsultaIM = True
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIM
    If txtIM <> "" Then
        txtIm_LostFocus
    End If
    'blnConsultaIM = False
End Sub

Private Sub Form_Load()
            
    Dim Controle As Control
    Dim i As Byte
    'Call Edita.AtualizaCombo(Bdados, cboTaxa, "select distinct(TIP_NOME_IMPOSTO) from tab_imposto" & _
    " WHERE  tip_cod_imposto in (SELECT tpi_tip_cod_imposto FROM Tab_Parametro_Imposto where tpi_tipo_tributo = 2 and tpi_valor_taxa_fixa > 0)")
'    cboTaxa.AddItem " "
    Grdtaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
    Call Edita.AtualizaCombo(Bdados, cboImposto, "select distinct(TIP_NOME_IMPOSTO) from tab_imposto" & _
    " WHERE  TIP_SIGLA_IMPOSTO LIKE '%ITBI%' OR TIP_NOME_IMPOSTO LIKE '%AFOR%'")
    Screen.MousePointer = 0
    'cboAliqFinancia.Preencher Bdados, "SELECT tge_codigo,tge_nome FROM VTSEG.DBO.TAB_GERAL where tge_tipo = 791 and tge_codigo <>0", 1
    'cboAliqFinancia.PreencherGeral Bdados, "ALIQUOTA ITBI FINANCIADO"
'    cboAliqProp.PreencherGeral Bdados, "ALIQUOTA ITBI PROPRIO"
    cabVisual.Exibir Bdados, Me.Name, App.Path
    txtPeriodo.Enabled = False
    txtPeriodo = Mid(Date, 4, 2) & "/" & Right(Date, 4)
    txtPeriodo_LostFocus
    txtSubTotalaliquotaPropria = "0,00"
    txtAliquotaPropria = "0,00"
    txtValorAvista = "0,00"
    txtSubTotalAliquotaFinanciada = "0,00"
    txtAliquotaFinanciada = "0,00"
    txtValorFinanciado = "0,00"
    txtValorITBI = "0,00"
    txtSubTotalaliquotaPropria.Enabled = False
    txtSubTotalAliquotaFinanciada.Enabled = False
    cboImposto.ListIndex = 0
    'txtValorAvista.Enabled = Bdados.AbreTabela("SELECT * FROM TAB_ACESSO_USUARIO WHERE TAU_TMO_COD_MODULO ='TCIU' and TAU_TFO_COD_FORMULARIO =103 AND TAU_TUS_COD_USUARIO='" & Aplicacoes.Usuario & "'")
End Sub

Private Sub Text1_Change()

End Sub




Private Sub GrdTaxas_ItemClick(ByVal Item As MSComctlLib.IListItem)
    Grdtaxas.Caption = "Taxas - " & Grdtaxas.SelectedItem & " - " & Grdtaxas.SelectedItem.SubItems(1)
End Sub

Private Sub txtAdquirente_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAliquotaFinanciada_Change()
    Calcula
End Sub

Private Sub txtAliquotaPropria_Change()
    Calcula
End Sub

Private Sub txtcgc_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtcgc_LostFocus()
    
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtCgc) = "" Then Exit Sub
    If Len(txtCgc) = 18 Then Exit Sub
    If Len(txtCgc) = 11 Then
        txtCgc = Edita.FormataTexto(txtCgc, Cpf)
    ElseIf Len(txtCgc) = 14 And Mid(txtCgc, 4, 1) <> "." Then
        txtCgc = Edita.FormataTexto(txtCgc, Cgc)
    End If
    Sql = "select tci_nome,tci_im,tci_logradouro,tci_nome_logradouro," & _
    " tci_numero,tci_complemento,tci_bairro,tci_cidade,tci_UF from Tab_Contribuinte" & _
    " where tci_cgc_cpf='" & txtCgc & "' and tci_tsc_cod_sit_cad =1"
    If Bdados.AbreTabela(Sql, rs) Then
        txtCedente = rs(0)
        txtIM = rs!TCI_IM
        Endereco = rs!tci_logradouro & " " & rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_COMPLEMENTO
        Bairro = rs!tci_BAIRRO
        Cod_Cidade = rs!tci_cidade
        Uf = rs!tci_UF
        Call txtIm_LostFocus
        txtCedente.Enabled = False
        txtEnderecoCedente.Enabled = False
    Else
        Avisa "Contribuinte não cadastrado. Utilize o módulo de Manutenção de Contribuintes."
        txtCgc.SetFocus
'        txtCedente.Enabled = True
'        txtCedente.SetFocus
'        txtEnderecoCedente.Enabled = True
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtContribuinte_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtFator_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtEnderecoAdquirente_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtic_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Valor As Double
    If txtIc = "" Then Exit Sub
    If Trim(txtIc) = "" Then
        If Not Confirma("O imóvel a ser transferido está situado na zona rural?") Then
            Avisa "A inscrição Cadastral deve ser informada."
            txtIc.SetFocus
            Exit Sub
        End If
        txtValorAvista = "":         txtValorImovel = ""
        txtEnderecoAdquirente = "": txtImovel = "": txtAdquirente = ""
    Else
        Sql = "select ttl_nome,tlg_nome,tba_nome,tim_numero,TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC,Tim_Zona,tim_valor,TIM_TCI_IM  ,TIM_SITUACAO_LOTE,tim_tci_im  from TAB_IMOVEL,VIS_BVT " & _
        "  where tim_ic='" & txtIc & "'  AND tim_tlg_cod_logradouro = " & _
        " tlg_cod_logradouro " 'AND VIS_BVT.tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio & " AND VIS_BVT.TBA_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio
        If Bdados.AbreTabela(Sql, rs) Then
            If "" & rs!TIM_SITUACAO_LOTE = 1 Then
                Informa "Imóvel desativado."
                txtIc.SetFocus
                Exit Sub
            End If
            txtImovel = rs(0) & " " & rs(1) & ", " & rs(2) & ", " & rs(3)
           ' ImContribuinte = Nvl("" & Rs!tim_tci_im, Const_ImAvulso)
            If Sigla = Imposto.NomeTributo(ttr_AFORO) Then
                Valor = Nvl("" & rs!TIM_VALOR_TERRENO, 0)
            Else
                Valor = CDbl(Nvl("" & rs!TIM_VALOR_EDIFIC, 0)) + CDbl(Nvl("" & rs!TIM_VALOR_TERRENO, 0))
            End If
            txtValorImovel = Format(Valor, Const_Monetario)
            txtValorAvista = Format(Valor, Const_Monetario)
            txtObservacao = BuscarDetalhes(txtIc)
            
            Sql = "SELECT tci_nome, tci_logradouro,tci_nome_logradouro,tci_bairro FROM TAB_CONTRIBUINTE WHERE TCI_IM = '" & rs!tim_tci_im & "'"
            If Bdados.AbreTabela(Sql, rs) Then
                txtAdquirente = "" & rs!tci_nome
                txtEnderecoAdquirente = rs!tci_logradouro & " " & rs!tci_nome_logradouro & " " & rs!tci_BAIRRO
            End If
'            Transf.BuscaDetalhesTransferencia txtIc
'            txtAliqFin = Transf.AliquotaFinanciado
'            txtAliqProp = Transf.AliquotaProprio
'            txtOcupa = Transf.OcupacaoLote
'            txtDestino = Transf.DestinoLote
        Else
            txtValorImovel = ""
            txtAdquirente = ""
            txtEnderecoAdquirente = ""
            Titular = Const_ImAvulso
        End If
        Bdados.FechaTabela rs
    End If
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    On Error GoTo trata
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIM = Imposto.FormataInscricao(txtIM, InscContrib)
    End If
    Dim Sql As String
    Dim rs As VSRecordset
    Dim RsIptu As VSRecordset
    If Trim(txtIM) = "" Then Exit Sub
    
    Sql = "select * from Tab_Contribuinte where tci_im='" & txtIM & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtCedente = "" & rs!tci_nome
        txtEnderecoCedente = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_BAIRRO & ", CEP: " & rs!tci_cep & ", " & rs!tci_cidade & "-" & rs!tci_UF
        txtCgc = "" & rs!TCI_CGC_CPF
        
        Endereco = "" & rs("tci_logradouro") & "  " & rs("tci_nome_logradouro") & "," & rs("tci_numero") & " " & rs("tci_complemento")
        txtEnderecoCedente = Endereco
        Bairro = "" & rs("tci_bairro") & ""
        Cod_Cidade = "" & rs("tci_cidade") & ""
        Cep = "" & rs("tci_cep") & ""
        Uf = "" & rs("tci_uf") & ""
        Titular = "" & rs!TCI_IM
        txtIc.SetFocus
    Else
        Util.Avisa "Inscrição não cadastrada."
        txtIM.SetFocus
    End If
    Bdados.FechaTabela rs
    Exit Sub
trata:
    If Err.Number = 3265 Then
        Resume Next
    End If
End Sub

Private Sub txtImovel_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPeriodo_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Hoje As String
    Dim SqlParc As String
    If Len(txtPeriodo) < 6 And Trim(txtPeriodo) <> "" Then
        Informa "Período inválido."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    If Len(txtPeriodo) = 6 Then
        txtPeriodo.MaxLength = 7
        txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
        txtPeriodo.MaxLength = 6
        Exercicio = txtPeriodo
    End If
    txtDtVenc = UltimoDiaDoMes(Format(Date, "dd/mm/yyyy"))
End Sub

Private Sub txtValorAvista_Change()
    Calcula
End Sub

Private Sub txtValorFinanciado_Change()
    Calcula
End Sub

Private Sub txtValorImovel_Change()
        If Sigla = Imposto.NomeTributo(ttr_ITBI) Then
            Transf.BuscaDetalhesTransferencia txtIc, CDbl(Nvl(txtValorImovel, 0))
'            txtAliqFin = Transf.AliquotaFinanciado
'            txtAliqProp = Transf.AliquotaProprio
            txtOcupa = Transf.OcupacaoLote
            txtDestino = Transf.DestinoLote
'            If Transf.AliquotaFinanciado = 0 And Transf.AliquotaProprio = 0 Then
'                Avisa "Cadastro do imóvel " & txtIc & " com problemas."
'                Exit Sub
'            End If

        End If
End Sub

Private Sub txtValorImovel_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub


Private Sub txtValorImovel_LostFocus()
    txtValorImovel = Edita.FormataTexto(txtValorImovel, Monetario, True)
End Sub

Private Function BuscarDetalhes(Ic As String) As String
    Dim Sql As String, rs As VSRecordset
    Dim Tipologia As String, Estrutura As String, Destinacao As String, AreaEdifTotal As String
    
    Sql = "SELECT TIM_QUADRA," & _
            " TIM_LOTE," & _
            " TDI_VALOR_ITEM" & _
        " FROM VIS_IMOVEL, VIS_DETALHE_IMOVEL" & _
        " WHERE TIM_IC = TDI_TIM_IC AND " & _
            " TIM_IC='" & Ic & "' AND" & _
            " TCO_DESCRICAO_COMPONENTE='ÁREA DO LOTE'"
    If Bdados.AbreTabela(Sql, rs) Then
        BuscarDetalhes = ("QD " & "" & rs!tim_quadra) & " - " & ("LOTE " & "" & rs!tim_Lote) & " - " & ("AREA LOTE " & "" & rs!TDI_VALOR_ITEM)
        Tipologia = Bdados.BuscaCodigo("SELECT TCO_DESCRICAO_COMPONENTE FROM VIS_DETALHE_IMOVEL WHERE TDI_TIM_IC='" & Ic & "' AND TDI_TIM_IC_UNIDADE=1 AND TDI_TGC_COD_GRUPO=9")
        Estrutura = Bdados.BuscaCodigo("SELECT TCO_DESCRICAO_COMPONENTE FROM VIS_DETALHE_IMOVEL WHERE TDI_TIM_IC='" & Ic & "' AND TDI_TIM_IC_UNIDADE=1 AND TDI_TGC_COD_GRUPO=10")
        Destinacao = Bdados.BuscaCodigo("SELECT TCO_DESCRICAO_COMPONENTE FROM VIS_DETALHE_IMOVEL WHERE TDI_TIM_IC='" & Ic & "' AND TDI_TIM_IC_UNIDADE=1 AND TDI_TGC_COD_GRUPO=11")
        AreaEdifTotal = Bdados.BuscaCodigo("SELECT TDI_VALOR_ITEM FROM VIS_DETALHE_IMOVEL WHERE TDI_TIM_IC='" & Ic & "' AND TDI_TIM_IC_UNIDADE=1 AND TDI_TGC_COD_GRUPO=113")
        If Tipologia <> "" Then
            BuscarDetalhes = BuscarDetalhes & Space(80) & _
                                Tipologia & " - " & _
                                Estrutura & " - " & _
                                Destinacao & " - " & _
                                "AREA EDIF TOTAL " & AreaEdifTotal
        End If
    End If
    Bdados.FechaTabela
End Function
Private Sub Calcula()
    txtSubTotalaliquotaPropria = CDbl(Nvl(txtAliquotaPropria, 0)) * CDbl(Nvl(txtValorAvista, 0)) / 100
    txtSubTotalAliquotaFinanciada = CDbl(Nvl(txtAliquotaFinanciada, 0)) * CDbl(Nvl(txtValorFinanciado, 0)) / 100
    
    txtValorImovel = CDbl(Nvl(Trim(txtValorFinanciado), 0)) + CDbl(Nvl(Trim(txtValorAvista), 0))
    
    txtValorITBI = CDbl(Nvl(txtSubTotalaliquotaPropria, 0)) + CDbl(Nvl(txtSubTotalAliquotaFinanciada, 0))
    
End Sub
Private Sub Pega_taxas()
    Dim i As Integer
    Dim Pos As Integer
    String_Taxas = ""
    Total_Taxas = 0
    For i = 1 To Grdtaxas.ListItems.Count
        If Grdtaxas.ListItems(i).Checked Then
            Pos = InStr(Grdtaxas.ListItems(i).SubItems(1), "-") - 1
            If String_Taxas = "" Then
                String_Taxas = String_Taxas & " [ " & Left(Grdtaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(Grdtaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            Else
                String_Taxas = String_Taxas & ", [ " & Left(Grdtaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(Grdtaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            End If
            Total_Taxas = Total_Taxas + CCur(Grdtaxas.ListItems(i).SubItems(2))
        End If
    Next
End Sub
