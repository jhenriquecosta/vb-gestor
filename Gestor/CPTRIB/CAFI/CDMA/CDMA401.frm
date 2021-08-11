VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDMA401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDMA401"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraOcupacao 
      Height          =   870
      Left            =   15
      TabIndex        =   11
      Top             =   1875
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1535
      Altura          =   1905
      Caption         =   " grdMaquinagrdMaquinagrdMaquina"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   315
         Left            =   9075
         TabIndex        =   7
         Top             =   420
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   660
         TabIndex        =   4
         Top             =   435
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         Caption         =   "Código"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   8
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   6015
         TabIndex        =   6
         Tag             =   "Segundo Horário"
         Top             =   420
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtSerie 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Tag             =   "Destinação "
         Top             =   435
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   503
         Caption         =   "Série (Nº )"
         Text            =   ""
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   50
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1125
      Left            =   15
      TabIndex        =   12
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   705
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1984
      Altura          =   1905
      Caption         =   " Dados do Proprietário"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
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
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3165
         TabIndex        =   2
         Top             =   375
         Width           =   7245
         _ExtentX        =   12779
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
         Tag             =   "Ins. Municipal"
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
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin VTOcx.grdVISUAL grdMaquina 
      Height          =   4050
      Left            =   15
      TabIndex        =   13
      Top             =   2805
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   7144
      CorBorda        =   32768
      Caption         =   "Equipamentos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1138
      Icone           =   "CDMA401.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   15
      Top             =   6915
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   847
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   8415
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   8
         Top             =   105
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         Caption         =   "Relatório&s"
         Acao            =   8
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
         Icone           =   "CDMA401.frx":031A
      End
   End
End
Attribute VB_Name = "CDMA401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MaquinaEquipamento  As New eMaquinaEquipamentoEletromec

Private Sub cmdBuscar_Click()
Dim Codigo As String, nSerie As String, Status As String

If txtCodigo <> "" Then Codigo = txtCodigo
If txtSerie <> "" Then nSerie = txtSerie
If cboStatus <> "" Then Status = cboStatus.Coluna(1).VALOR
MaquinaEquipamento.PreencherGrdConsulta grdMaquina, txtIm, Codigo, nSerie, Status
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdMaquina.ListItems.Clear
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

