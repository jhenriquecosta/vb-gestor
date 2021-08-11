VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDAP101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDAP101"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CDAP101.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "CDAP101.frx":08CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   18
      Top             =   3705
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   7440
         TabIndex        =   14
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   9420
         TabIndex        =   16
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   8430
         TabIndex        =   15
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
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
      TabIndex        =   19
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1138
      Icone           =   "CDAP101.frx":29ED
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1080
      Left            =   30
      TabIndex        =   20
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   690
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   1905
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
         TabIndex        =   22
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
         TabIndex        =   1
         Top             =   375
         Width           =   7215
         _ExtentX        =   12726
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
         Tag             =   "Insc. Municipal"
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
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   465
         TabIndex        =   21
         Top             =   750
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin VTOcx.fraVISUAL fraVeiculo 
      Height          =   1800
      Left            =   15
      TabIndex        =   23
      Top             =   1815
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   3175
      Altura          =   1905
      Caption         =   " Dados do Veículo"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtComplemento 
         Height          =   480
         Left            =   4560
         TabIndex        =   13
         Top             =   1305
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   847
         Caption         =   "Complemento"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   100
      End
      Begin VTOcx.txtVISUAL txtInicioAtividadeCarro 
         Height          =   480
         Left            =   2925
         TabIndex        =   12
         Tag             =   "Início da Atividade"
         Top             =   1320
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   847
         Caption         =   "Início da Atividade"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtLicenca 
         Height          =   480
         Left            =   8580
         TabIndex        =   9
         Tag             =   "Licenciamento"
         Top             =   780
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   847
         Caption         =   "Licenciamento"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   20
      End
      Begin VTOcx.cboVISUAL cboUF 
         Height          =   510
         Left            =   2085
         TabIndex        =   11
         Tag             =   "UF"
         ToolTipText     =   "UF"
         Top             =   1290
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
         Caption         =   "UF"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtMunicipio 
         Height          =   480
         Left            =   135
         TabIndex        =   10
         Tag             =   "Município"
         Top             =   1320
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   847
         Caption         =   "Município"
         Text            =   ""
         Restricao       =   1
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   20
      End
      Begin VTOcx.txtVISUAL txtAnoFabric 
         Height          =   480
         Left            =   4725
         TabIndex        =   6
         Tag             =   "Ano"
         Top             =   780
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   847
         Caption         =   "Ano"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   4
         MinLen          =   4
      End
      Begin VTOcx.txtVISUAL txtPlaca 
         Height          =   480
         Left            =   5505
         TabIndex        =   7
         Tag             =   "Placa"
         Top             =   780
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   847
         Caption         =   "Placa"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   15
      End
      Begin VTOcx.txtVISUAL txtChassi 
         Height          =   480
         Left            =   6765
         TabIndex        =   8
         Tag             =   "Chassi"
         Top             =   780
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   847
         Caption         =   "Chassi"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   20
      End
      Begin VTOcx.txtVISUAL txtMarca 
         Height          =   480
         Left            =   1665
         TabIndex        =   4
         Tag             =   "Marca"
         Top             =   780
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   847
         Caption         =   "Marca"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   50
      End
      Begin VTOcx.txtVISUAL txtModelo 
         Height          =   480
         Left            =   3195
         TabIndex        =   5
         Tag             =   "Modelo"
         Top             =   780
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   847
         Caption         =   "Modelo"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   50
      End
      Begin VTOcx.txtVISUAL txtVeiculo 
         Height          =   480
         Left            =   135
         TabIndex        =   3
         Tag             =   "Veículo"
         Top             =   780
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   847
         Caption         =   "Veículo"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   50
      End
      Begin VTOcx.cboVISUAL cboAtividadeVeiculo 
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Tag             =   "Atividade Desempenhada"
         Top             =   420
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   556
         Caption         =   "Atividade Desempenhada"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   4210752
      End
   End
End
Attribute VB_Name = "CDAP101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AparelhoTransporte  As New eAparelhoTransporte
Dim Ativ As New Atividade

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

 
Private Sub LimpaCampo()
    txtVeiculo = ""
    txtMarca = ""
    txtModelo = ""
    txtAnoFabric = ""
    txtPlaca = ""
    txtMunicipio = ""
    cboUF.ListIndex = -1
    txtLicenca = ""
    txtChassi = ""
    cboAtividadeVeiculo.ListIndex = -1
    txtInicioAtividadeCarro = ""
    txtComplemento = ""
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub



Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
      
    With AparelhoTransporte
            .Im = txtIm
            .Veiculo = txtVeiculo
            .Marca = txtMarca
            .Modelo = txtModelo
            .AnoFabricacao = txtAnoFabric
            .Placa = txtPlaca
            .Municipio = txtMunicipio
            .UF = cboUF.Coluna(1).VALOR
            .Licenca = txtLicenca
            .Chassi = txtChassi
            .Atividade = cboAtividadeVeiculo.Coluna(1).VALOR
            .IniAtividadeCarro = txtInicioAtividadeCarro
            .Complemento = txtComplemento
        If .Salvar = True Then
            Avisa "Dados Salvos com Sucesso."
            LimpaCampo
        End If
    End With
    Screen.MousePointer = 0
End Sub
    

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     cboUF.PreencherGeral Bdados, "UF"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
     Ativ.PreencherCboAtiv cboAtividadeVeiculo
End Sub


Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub
Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub

