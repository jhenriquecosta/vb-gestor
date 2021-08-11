VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CUST101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLCO101"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   5
      Top             =   8370
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   847
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   8175
         TabIndex        =   4
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   7170
         TabIndex        =   3
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   1138
      Icone           =   "CUST101.frx":0000
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   6165
      Left            =   45
      TabIndex        =   7
      Top             =   2160
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10874
      CorBorda        =   32768
      Caption         =   "Lancamentos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1395
      Left            =   30
      TabIndex        =   8
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   705
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   2461
      Altura          =   1905
      Caption         =   " Dados do Proprietário"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   6540
         TabIndex        =   11
         Top             =   930
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         Caption         =   "&Adicionar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   510
         Left            =   240
         TabIndex        =   10
         Tag             =   "Tipo"
         ToolTipText     =   "783 - TIPO CONTA"
         Top             =   810
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   900
         Caption         =   "Tipo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   4210752
         Editavel        =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   480
         Left            =   4305
         TabIndex        =   2
         Tag             =   "                    "
         Top             =   795
         Visible         =   0   'False
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   847
         Caption         =   "Valor"
         Text            =   ""
         Formato         =   5
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   50
      End
      Begin VTOcx.cboVISUAL cboDescricao 
         Height          =   510
         Left            =   1875
         TabIndex        =   1
         Tag             =   "Tipo"
         ToolTipText     =   "783 - TIPO CONTA"
         Top             =   285
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   900
         Caption         =   "Descrição"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   4210752
         Editavel        =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtData 
         Height          =   480
         Left            =   255
         TabIndex        =   0
         Tag             =   "Conta/Subconta"
         Top             =   300
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   847
         Caption         =   "Data"
         Text            =   ""
         Formato         =   0
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   50
      End
   End
   Begin VTOcx.txtVISUAL txtCodigo 
      Height          =   480
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   847
      Caption         =   ""
      Text            =   ""
      Requerido       =   0   'False
      AlinhamentoRotulo=   1
      CorRotulo       =   16384
      CorTexto        =   4194304
   End
End
Attribute VB_Name = "CUST101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim CodLancamento As String
    Campos = "TLA_COD_LANCAMENTO,TLA_TIPO_LANCAMENTO,TLA_DATA,TLA_DESCRICAO,TLA_VALOR"
    CodLancamento =
    Valores = Bdados.PreparaValor()
    Bdados.InsereDados "TAB_LANCAMENTO", Valores, Campos
End Sub

Private Sub Form_Load()
    cboTipo.PreencherGeral Bdados, "TIPO LANCAMENTO"
End Sub
