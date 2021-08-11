VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TPAR101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TPAR101.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.TextBox txtVence 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   2025
      TabIndex        =   0
      Tag             =   "Data Vencimento"
      Top             =   660
      Width           =   1260
   End
   Begin Threed.SSFrame fra 
      Height          =   705
      Index           =   3
      Left            =   90
      TabIndex        =   13
      Top             =   4935
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   1244
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
      Caption         =   "Resultados Parciais do Parcelamento"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtdebitoRestante 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   4290
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txtParcelaUm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   8655
         TabIndex        =   9
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txtPercEntrada 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6615
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1350
         Width           =   780
      End
      Begin VB.TextBox txtCotas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6165
         TabIndex        =   5
         Tag             =   "Cotas"
         Top             =   255
         Width           =   690
      End
      Begin VB.TextBox txtTotalParc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1455
         TabIndex        =   15
         Top             =   255
         Width           =   1245
      End
      Begin VB.TextBox txtValorPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5265
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1185
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   60
         TabIndex        =   16
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
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
         Caption         =   "Débito Parcelado"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   5
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   7125
         TabIndex        =   17
         Top             =   1080
         Width           =   1380
         _ExtentX        =   2434
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
         Caption         =   "Débito Pendente"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   7
         Left            =   3915
         TabIndex        =   18
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
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
         Caption         =   "Valor Parcelado"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   5
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   5
         Left            =   5625
         TabIndex        =   21
         Top             =   315
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
         Caption         =   "Cotas"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   6
         Left            =   5220
         TabIndex        =   22
         Top             =   1410
         Width           =   1320
         _ExtentX        =   2328
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
         Caption         =   "Perc. Entrada:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   9
         Left            =   6975
         TabIndex        =   23
         Top             =   315
         Width           =   1650
         _ExtentX        =   2910
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
         Caption         =   "Valor da Entrada(R$)"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   2850
         TabIndex        =   35
         Top             =   315
         Width           =   1380
         _ExtentX        =   2434
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
         Caption         =   "Débito Pendente"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   975
      Index           =   1
      Left            =   90
      TabIndex        =   19
      Top             =   5700
      Width           =   10890
      _ExtentX        =   19209
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
      Caption         =   "Detalhes"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtObservacao 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   1350
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   9435
      End
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   19
         Left            =   120
         TabIndex        =   20
         Top             =   240
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
         Caption         =   "Observações :"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VTOcx.grdVISUAL lstParcelas 
      Height          =   2430
      Left            =   90
      TabIndex        =   24
      Top             =   2415
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   4286
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      CheckBox        =   -1  'True
      Ordenavel       =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TPAR101.frx":2123
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   12
      Top             =   90
      Width           =   375
   End
   Begin VTOcx.cmdVISUAL cmdCancela 
      Height          =   375
      Left            =   8685
      TabIndex        =   7
      Top             =   6750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Limpar"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   9855
      TabIndex        =   8
      Top             =   6750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdParcela 
      Height          =   375
      Left            =   6495
      TabIndex        =   6
      Top             =   6750
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      Caption         =   "&Gerar Parcelamento"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin Threed.SSPanel lbl 
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   26
      Top             =   720
      Width           =   1860
      _ExtentX        =   3281
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
      Caption         =   "Vencimento Parcela 1"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   3
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSFrame fra 
      Height          =   1380
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   990
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   2434
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
      Caption         =   "Contribuinte"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.CheckBox chkAtualizacao 
         Caption         =   "Atualização Juros/Multa ?"
         Height          =   255
         Left            =   7080
         TabIndex        =   36
         Top             =   240
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.TextBox txtIm 
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
         Left            =   930
         TabIndex        =   2
         Top             =   240
         Width           =   1740
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   360
         TabIndex        =   28
         Top             =   600
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   556
         Caption         =   "Razão"
         Text            =   ""
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   15
         Left            =   165
         TabIndex        =   29
         Top             =   285
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
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
         Caption         =   "Inscricao"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   645
         Left            =   9465
         TabIndex        =   4
         Top             =   630
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1138
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   315
         Left            =   90
         TabIndex        =   30
         Top             =   960
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         Caption         =   "Endereço"
         Text            =   ""
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   2700
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   3090
         TabIndex        =   3
         Top             =   240
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   6570
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   225
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
   End
   Begin VTOcx.cboVISUAL cboImposto 
      Height          =   315
      Left            =   3345
      TabIndex        =   1
      Tag             =   "Tributo"
      Top             =   660
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   556
      Caption         =   "Tributo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
End
Attribute VB_Name = "TPAR101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Imposto As New VSImposto
Dim MaxCotas As Byte
Dim CodImp As String
Dim Cgc As String
Dim EnderecoContrib As String
Dim CodPagamento As Double
Dim TipoTransacao As TipoTransacao

Sub AtualizaLista()
    Dim Sql As String
    Dim ValorTotal As Double
    Dim Obrig As New Obrigacao
    Dim Conta As New ContaCorrente
    Dim Codigo As String
    Dim modo As TipoInscricaoObrigacao
    
    ValorTotal = 0
    txtTotalParc = ""
    txtdebitoRestante = ""
    txtParcelaUm = ""
    
        If txtIm <> "" Then
            Codigo = txtIm
            modo = etiContribuinte
            If chkAtualizacao.Value = 1 Then
                Conta.ExecutaAtualizacao txtIm, etiContribuinte, False, , ettParcelada, txtVence, 0, , , , CStr("" & cboImposto.Coluna(0).Valor)
            End If
        Else
            modo = etiImovel
            Codigo = txtImovel
            If chkAtualizacao.Value = 1 Then
                Conta.ExecutaAtualizacao txtImovel, etiImovel, False, , ettParcelada, txtVence, 0, , , , CStr("" & cboImposto.Coluna(0).Valor)
            End If
        End If
        If chkAtualizacao.Value = 1 Then
            
            If Obrig.CarregaListaObrigacaoAtualizada(lstParcelas, IIf(Trim(txtIm) = "", txtImovel, txtIm), CStr(cboImposto.Coluna(0).Valor), , etlParcelaveis, , modo) Then
                If lstParcelas.ListItems.Count > 0 Then ValorTotal = Format(lstParcelas.Colunas(11).Soma, Const_Monetario)
                If lstParcelas.ListItems.Count > 0 Then txtTotalParc = Format(0, Const_Monetario)
                If lstParcelas.ListItems.Count > 0 Then txtdebitoRestante = Format(ValorTotal, Const_Monetario)
            Else
                Avisa "Nenhum débito vencido ou lancado."
            End If
        Else
            If Obrig.CarregaListaObrigacaoAtualizada(lstParcelas, IIf(Trim(txtIm) = "", txtImovel, txtIm), CStr(cboImposto.Coluna(0).Valor), , etlNaoPagos, , modo) Then
                If lstParcelas.ListItems.Count > 0 Then ValorTotal = Format(lstParcelas.Colunas(6).Soma, Const_Monetario)
                If lstParcelas.ListItems.Count > 0 Then txtTotalParc = Format(0, Const_Monetario)
                If lstParcelas.ListItems.Count > 0 Then txtdebitoRestante = Format(ValorTotal, Const_Monetario)
            Else
                Avisa "Nenhum débito vencido ou lancado."
            End If
        
        End If
    
    Screen.MousePointer = 0
End Sub


Private Sub cmdBuscar_Click()
    
    If cboImposto.ListIndex < 0 Then
        Avisa "Informe o tributo."
        cboImposto.SetFocus
        Exit Sub
    End If
    AtualizaLista
    
    cmdParcela.Enabled = True
    txtCotas.SetFocus
End Sub

Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    lstParcelas.ListItems.Clear
    cmdParcela.Enabled = True
    txtIm.SetFocus
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImprime_Click()
    Dim Barra As Boolean
    Dim Cobranca As New VSCobranca
    
    Barra = False
    If CodPagamento = 0 Then
        Informa "Não há extrato para ser impresso."
        Exit Sub
    End If
    Screen.MousePointer = 11
    If Confirma("O extrato será usado para pagamento do débito?") Then
        If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TDAMExtratoBarra.rpt") Then Exit Sub
        Barra = True
    Else
        If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TDAMExtrato.rpt") Then Exit Sub
        
    End If
    With Rpt
        .Formulas "VT_EXTRATO ", CStr(CodPagamento)
        .Formulas "VT_PRAZO ", txtVence
        .Formulas "VT_OBS_GERAL ", txtObservacao
        .Selecao = "{TAB_PAGAMENTO_EXTRATO.TPE_COD_PAGAMENTO_EXTRATO} =" & CodPagamento
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        If Barra Then
'            .Formulas "LinhaDigitavel", Cobranca.GeraCodBarra(CStr(CodPagamento), 0, CDbl(txtValorPago), PicBarra, Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), txtVence)
        End If
        .Titulo = "Extrato de Lançamento"
        .Arvore = False
        .Visualizar
    End With
    Screen.MousePointer = 0
End Sub

Private Sub cmdParcela_Click()
    Dim Sql As String
    Dim CCorrente As New ContaCorrente
    Dim rs As VSRecordset
    Dim Conta As New ContaCorrente
    Dim i As Integer
    Dim CodImposto As String
    Dim Periodo As String
    Dim PeriodoInicial As Double, PeriodoFinal As Double
    Dim Parcelamento As Double
    Dim Valores As String
    Dim modo As TipoInscricaoObrigacao
    Dim Codigo As String
    Dim Motivo As String
    
    On Error GoTo Trata
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
    Valores = ""
    Periodo = Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2)
    If CDbl(Nvl(txtParcelaUm, 0)) = 0 Then
        Informa "Informe dados do parcelamento."
        Exit Sub
    End If
'    cmdBuscar_Click
    Screen.MousePointer = 11
    BuscaPeriodo lstParcelas, PeriodoInicial, PeriodoFinal
    
    CodImposto = cboImposto.Coluna(0).Valor
    If txtIm <> "" Then
        Codigo = txtIm
        modo = etiContribuinte
    Else
        Codigo = txtImovel
        modo = etiImovel
    End If

    Parcelamento = Conta.CriaParcelamento(Trim(Codigo), CStr(cboImposto.Coluna(0).Valor), _
         CInt(txtCotas), txtVence, txtTotalParc, lstParcelas, Nvl(txtParcelaUm, 0), modo, txtObservacao)
    If Parcelamento <> 0 Then
        
        cmdParcela.Enabled = False
        'separa tributos parcelados

        With Rpt
            
               If Not .DefinirArquivo(Bdados, App.Path + "\TermoParcela.rpt") Then Exit Sub
               .Formulas "NumParcelamento ", CStr(Parcelamento)
                .Formulas "Municipio ", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
                .Formulas "Imposto ", CStr(cboImposto.Coluna(2).Valor)
                .Formulas "Inscricao", IIf(Trim(txtIm) = "", txtImovel, txtIm)
                .Formulas "Contribuinte", txtRazao
                .Formulas "Endereco", txtEndereco
                If IsNumeric(txtTotalParc) Then
                    .Formulas "ValorExtenso", VBA.UCase(Extenso(CDbl(txtTotalParc.Text), "Reais", "Real"))
                End If
                .Formulas "VT_Periodo ", IIf(Len(CStr(PeriodoInicial)) = 4, CStr(PeriodoInicial), Right(CStr(PeriodoInicial), 2) & "/" & Left(CStr(PeriodoInicial), 4)) & " a " & IIf(Len(CStr(PeriodoFinal)) = 4, CStr(PeriodoFinal), Right(CStr(PeriodoFinal), 2) & "/" & Left(CStr(PeriodoFinal), 4))
                .Selecao = "{Tab_Parcelamento.TPA_NUM_PARCELAMENTO} = " & Parcelamento
                If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                    .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "GDA")
                Else
                    .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                End If
                .Titulo = "Termo de Parcelamento"
                .Arvore = False
                .Visualizar
        End With
        Set Rpt = Nothing
        Informa "Parcelamento gerado com sucesso."
        Bdados.FechaTabela rs
        Call cmdCancela_Click
    Else
        Avisa "Parcelamento não foi gerado. Verifique os parametros do sistema."
    End If
    Screen.MousePointer = 0
    Exit Sub
Trata:
    Avisa Err.Description
    Exit Sub
    Resume
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    AtualizaCabecalho lstParcelas
    cboImposto.Preencher Bdados, "Select  tip_cod_imposto,TIP_sigla_IMPOSTO  " & Bdados.Concatena & " ' # ' " & Bdados.Concatena & " tip_nome_imposto,tip_nome_imposto From TAB_IMPOSTO order by TIP_sigla_IMPOSTO asc", 1
    cboImposto.AddItem " "
    TipoTransacao = ettParcelada
End Sub

Private Sub lstParcELAS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstParcelas, ColumnHeader
End Sub

Private Sub txtCotas_Change()
    If Trim(txtCotas) <> "" And Trim(txtVence) <> "" Then
        If CDbl(txtCotas) = 0 Then Exit Sub
        If Temp.PegaParametro(Bdados, "DESCONTO DIFERENCIADO PARCELAMENTO") = "SIM" Then
            If CDbl(Nvl(Trim(txtCotas), 0)) > Temp.PegaParametro(Bdados, "MAX COTAS DESCONTO PARCELAMENTO") Then
                If Confirma("Limite máximo de parcelas ultrapassado para desconto. Deseja prosseguir?") Then
                    TipoTransacao = ettDividaAtiva
                    cmdBuscar_Click
                    Exit Sub
                End If
            Else
                If TipoTransacao = ettDividaAtiva Then
                    TipoTransacao = ettParcelada
                    cmdBuscar_Click
                    Exit Sub
                End If
            End If
        End If
            If chkAtualizacao.Value = 1 Then
                txtParcelaUm = CDbl(Nvl(txtTotalParc, 0)) / CDbl(Nvl(txtCotas, 0))
                txtParcelaUm = Format(CDbl(Nvl(txtParcelaUm, 0)) + (CDbl(Nvl(txtParcelaUm, 0)) * (DateDiff("m", Date, txtVence) / 100)), Const_Monetario)
            Else
                txtParcelaUm = CDbl(Nvl(txtTotalParc, 0)) / CDbl(Nvl(txtCotas, 0))
                txtParcelaUm = Format(CDbl(Nvl(txtParcelaUm, 0)), Const_Monetario)
            End If
    Else
        txtParcelaUm = ""
    End If
End Sub

Private Sub txtCotas_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub lstParcelas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        
        If Nvl(txtdebitoRestante, 0) = txtTotalParc Then
            'Queiroz (04/09/2002)
                'txtdebitoRestante = Format(CDbl(txtTotalParc) - CDbl(Item.SubItems(5)), Const_Monetario)
                    If chkAtualizacao.Value = 1 Then
                        txtdebitoRestante = Format(CDbl(Nvl(txtTotalParc, 0)) - CDbl(Nvl(Item.SubItems(10), 0)), Const_Monetario)
                        txtTotalParc = Format(CDbl(Nvl(txtTotalParc, 0)) + CDbl(Nvl(Item.SubItems(10), 0)), Const_Monetario)
                    Else
                        txtdebitoRestante = Format(CDbl(Nvl(txtTotalParc, 0)) - CDbl(Nvl(Item.SubItems(5), 0)), Const_Monetario)
                        txtTotalParc = Format(CDbl(Nvl(txtTotalParc, 0)) + CDbl(Nvl(Item.SubItems(5), 0)), Const_Monetario)
                    End If
        Else
            'Queiroz (04/09/2002)
                'txtdebitoRestante = Format(CDbl(txtdebitoRestante) - CDbl(Item.SubItems(5)), Const_Monetario)
                If chkAtualizacao.Value = 1 Then
                    txtdebitoRestante = Format(CDbl(Nvl(txtdebitoRestante, 0)) - CDbl(Nvl(Item.SubItems(10), 0)), Const_Monetario)
                    txtTotalParc = Format(CDbl(Nvl(txtTotalParc, 0)) + CDbl(Nvl(Item.SubItems(10), 0)), Const_Monetario)
                
                Else
                    txtdebitoRestante = Format(CDbl(Nvl(txtdebitoRestante, 0)) - CDbl(Nvl(Item.SubItems(5), 0)), Const_Monetario)
                    txtTotalParc = Format(CDbl(Nvl(txtTotalParc, 0)) + CDbl(Nvl(Item.SubItems(5), 0)), Const_Monetario)
                End If
        End If
        If CDbl(txtdebitoRestante) < 0 Then txtdebitoRestante = "0,00"
    Else
        'Queiroz (04/09/2002)
            'txtdebitoRestante = Format(CDbl(txtdebitoRestante) + CDbl(Item.SubItems(5)), Const_Monetario)
            If chkAtualizacao.Value = 1 Then
                txtdebitoRestante = Format(CDbl(Nvl(txtdebitoRestante, 0)) + CDbl(Nvl(Item.SubItems(10), 0)), Const_Monetario)
                txtTotalParc = Format(CDbl(Nvl(txtTotalParc, 0)) - CDbl(Nvl(Item.SubItems(10), 0)), Const_Monetario)
            Else
                txtdebitoRestante = Format(CDbl(Nvl(txtdebitoRestante, 0)) + CDbl(Nvl(Item.SubItems(5), 0)), Const_Monetario)
                txtTotalParc = Format(CDbl(Nvl(txtTotalParc, 0)) - CDbl(Nvl(Item.SubItems(5), 0)), Const_Monetario)
            End If
    End If
'    txtValorPago = Format(CDbl(txtTotalParc) - CDbl(txtdebitoRestante), Const_Monetario)
'    txtTotalParc = Format(CDbl(txtTotalParc) - CDbl(Item.SubItems(10)), Const_Monetario)
End Sub


Private Sub txtic_LostFocus()
    
End Sub

Private Sub txtIm_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco)
    
End Sub

Private Sub txtImovel_LostFocus()
 Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtIm.SetFocus
        End If
    End If
End Sub

Private Sub txtParcelaUm_LostFocus()
    txtParcelaUm = Format(Nvl(txtParcelaUm, 0), Const_Monetario)
End Sub


Private Sub txtPercEntrada_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtPercEntrada_LostFocus()
    txtCotas_Change
End Sub

Private Sub txtTotalParc_Change()
    If CDbl(Nvl(Trim(txtCotas), 0)) = 0 Then Exit Sub
    txtParcelaUm = CDbl(Nvl(txtTotalParc, 0)) / CDbl(Nvl(txtCotas, 0))
    txtParcelaUm = Format(CDbl(Nvl(txtParcelaUm, 0)) + (CDbl(Nvl(txtParcelaUm, 0)) * (DateDiff("m", Date, txtVence) / 100)), Const_Monetario)
End Sub

Private Sub txtValorPago_Change()
    txtCotas_Change
End Sub

Private Sub txtVence_LostFocus()
    If txtVence = "" Then Exit Sub
    txtVence = Edita.FormataTexto(txtVence, Data)
    txtCotas_Change
'    If CDate(txtVence) < Date Then
'        Util.Avisa "Data inválida."
'        txtVence.SetFocus
'    End If
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtInicio_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValidade_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub BuscaPeriodo(Lista As Object, PeridoInicial As Double, PeridoFinal As Double)
    Dim i As Integer
    Dim MudouInicial As Boolean
    
    PeridoInicial = 999999
    PeridoFinal = 0
    MudouInicial = False
    For i = 1 To lstParcelas.ListItems.Count
           If CDbl(Nvl(lstParcelas.ListItems(i).SubItems(3), 999999)) < PeridoInicial Then
                If lstParcelas.ListItems(i).Checked = True Then
                    PeridoInicial = lstParcelas.ListItems(i).SubItems(3)
                    MudouInicial = True
                End If
            End If
            If CDbl(Nvl(lstParcelas.ListItems(i).SubItems(3), 999999)) > PeridoFinal Then
                If lstParcelas.ListItems(i).Checked = True Then PeridoFinal = lstParcelas.ListItems(i).SubItems(3)
            End If
    Next
    If Not MudouInicial Then PeridoInicial = PeridoFinal
End Sub
