VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDEF401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDEF401"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ControlBox      =   0   'False
   Icon            =   "CDEF401.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "CDEF401.frx":058A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1138
      Icone           =   "CDEF401.frx":26AD
   End
   Begin VTOcx.grdVISUAL grdAmbulante 
      Height          =   4020
      Left            =   45
      TabIndex        =   12
      Top             =   2760
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   7091
      CorBorda        =   32768
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.fraVISUAL fraSanitario 
      Height          =   930
      Left            =   60
      TabIndex        =   13
      Top             =   1770
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   1640
      Altura          =   1905
      Caption         =   " Dados do Ponto"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   315
         Left            =   6375
         TabIndex        =   5
         Top             =   465
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   3315
         TabIndex        =   4
         Top             =   465
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
         Left            =   615
         TabIndex        =   3
         Top             =   480
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
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1050
      Left            =   30
      TabIndex        =   9
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   705
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   1852
      Altura          =   1905
      Caption         =   " Dados do Contribuinte"
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
         Width           =   345
         _ExtentX        =   609
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
         Width           =   7170
         _ExtentX        =   12647
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
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   450
         TabIndex        =   14
         Top             =   735
         Width           =   9885
         _ExtentX        =   17436
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
      TabIndex        =   15
      Top             =   6795
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   847
      Begin VTOcx.cmdVISUAL cmdRela 
         Height          =   345
         Left            =   7125
         TabIndex        =   6
         Top             =   105
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         Caption         =   "Relatório&s"
         Acao            =   8
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
         Icone           =   "CDEF401.frx":29C7
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   9420
         TabIndex        =   8
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
         Left            =   8415
         TabIndex        =   7
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
End
Attribute VB_Name = "CDEF401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim AmbulanteEventual  As New eAmbulanteEventual

Private Sub cmdBuscar_Click()
    Dim Codigo As String, STATUS As String
    
    If txtCodigo <> "" Then Codigo = txtCodigo
    If cboStatus <> "" Then STATUS = cboStatus.Coluna(1).VALOR
    
    AmbulanteEventual.PreencherGrdConsulta grdAmbulante, txtIm, Codigo, STATUS

End Sub

Private Sub cmdLimpar_Click()
 LimpaCampos Me
 grdAmbulante.ListItems.Clear
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
