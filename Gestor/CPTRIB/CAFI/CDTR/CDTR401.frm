VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDTR401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDTR401"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   1138
      Icone           =   "CDTR401.frx":0000
   End
   Begin VTOcx.grdVISUAL grdVeiculo 
      Height          =   3615
      Left            =   30
      TabIndex        =   13
      Top             =   2790
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6376
      CorBorda        =   32768
      Caption         =   "Veículos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.fraVISUAL fraVeiculo 
      Height          =   945
      Left            =   15
      TabIndex        =   14
      Top             =   1800
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   1667
      Altura          =   1905
      Caption         =   " Dados do Veículo"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   9120
         TabIndex        =   8
         Top             =   465
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtChassi 
         Height          =   300
         Left            =   4035
         TabIndex        =   6
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         Caption         =   "Chassi"
         Text            =   ""
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   20
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   495
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   503
         Caption         =   "Código"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   8
      End
      Begin VTOcx.txtVISUAL txtPlaca 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   495
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         Caption         =   "Placa"
         Text            =   ""
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   15
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   6045
         TabIndex        =   7
         Top             =   465
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1050
      Left            =   45
      TabIndex        =   15
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   705
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   1852
      Altura          =   1905
      Caption         =   " Dados do Proprietário"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   465
         TabIndex        =   3
         Top             =   750
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Tag             =   "Ins. Municipal"
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3165
         TabIndex        =   2
         Top             =   375
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   503
         Caption         =   "Nome/Razão Social"
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   375
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   16
      Top             =   6450
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   847
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   8415
         TabIndex        =   10
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
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   9420
         TabIndex        =   11
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
      Begin VTOcx.cmdVISUAL cmdRela 
         Height          =   345
         Left            =   7125
         TabIndex        =   9
         Top             =   105
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         Caption         =   "Relatório&s"
         Acao            =   8
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
         Icone           =   "CDTR401.frx":0C0A
      End
   End
End
Attribute VB_Name = "CDTR401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TransportePassageiro As New etransportePassageiro

Private Sub cmdBuscar_Click()
Dim Codigo As String, Placa As String, Chassi As String, Status As String
If txtIm = "" Then
    Avisa "Campo 'Ins. Municipal' deve ser informado."
    LimpaCampos Me
    Exit Sub
End If
If txtCodigo <> "" Then Codigo = txtCodigo
If txtPlaca <> "" Then Placa = txtPlaca
If txtChassi <> "" Then Chassi = txtChassi
If cboStatus <> "" Then Status = cboStatus.Coluna(1).VALOR

TransportePassageiro.PreencherGrdConsulta grdVeiculo, txtIm, Codigo, Placa, Chassi, Status
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdVeiculo.ListItems.Clear
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     cboStatus.Preencher Bdados, "select  TGE_NOME, TGE_CODIGO from vis_status_cad_fiscal"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
    
End Sub


