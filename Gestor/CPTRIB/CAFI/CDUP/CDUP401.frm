VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDUP401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDUP401"
   ClientHeight    =   6525
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   1138
      Icone           =   "CDUP401.frx":0000
   End
   Begin VTOcx.grdVISUAL grdUtilizacao 
      Height          =   3120
      Left            =   15
      TabIndex        =   11
      Top             =   2925
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   5503
      CorBorda        =   32768
      Caption         =   "Cadastros"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.fraVISUAL fraOcupacao 
      Height          =   975
      Left            =   15
      TabIndex        =   12
      Top             =   1890
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   1720
      Altura          =   1905
      Caption         =   " Dados da Utilização"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Top             =   405
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   660
         TabIndex        =   4
         Top             =   420
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   503
         Caption         =   "Código"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   8
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   6390
         TabIndex        =   6
         Top             =   405
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   582
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1125
      Left            =   30
      TabIndex        =   13
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   690
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1984
      Altura          =   1905
      Caption         =   " Dados do Proprietário"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   300
         Left            =   2760
         TabIndex        =   1
         Top             =   375
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3165
         TabIndex        =   2
         Top             =   375
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   503
         Caption         =   "Nome/Razão Social"
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Formato         =   8
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   465
         TabIndex        =   3
         Top             =   750
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   14
      Top             =   6045
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   847
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   8415
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   7
         Top             =   105
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         Caption         =   "Relatório&s"
         Acao            =   8
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
         Icone           =   "CDUP401.frx":0C0A
      End
   End
End
Attribute VB_Name = "CDUP401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UtilizacaoSolo  As New eUtilizacaoSolo

Private Sub cmdBuscar_Click()
Dim Codigo As String, Status As String

If txtCodigo <> "" Then Codigo = txtCodigo
If cboStatus <> "" Then Status = cboStatus.Coluna(1).VALOR

UtilizacaoSolo.PreencherGrdConsulta grdUtilizacao, txtIm, Codigo, Status
End Sub
Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdUtilizacao.ListItems.Clear
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
         cboStatus.PreencherGeral Bdados, "STATUS CADASTRO FISCAL"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
End Sub


Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub
