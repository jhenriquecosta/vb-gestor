VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDAP102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDAP102"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   ControlBox      =   0   'False
   Icon            =   "CDAP102.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "CDAP102.frx":08CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   19
      Top             =   7155
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   6585
         TabIndex        =   26
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   609
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   8550
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
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   9525
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
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   7575
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
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   1138
      Icone           =   "CDAP102.frx":29ED
   End
   Begin VTOcx.grdVISUAL grdVeiculo 
      Height          =   3390
      Left            =   30
      TabIndex        =   21
      Top             =   3750
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5980
      CorBorda        =   32768
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1125
      Left            =   45
      TabIndex        =   22
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   690
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
         TabIndex        =   24
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
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   503
         Caption         =   "Nome/Razão Social"
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   4210752
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
         CorRotulo       =   4210752
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   465
         TabIndex        =   23
         Top             =   750
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   4210752
         CorTexto        =   4194304
      End
   End
   Begin VTOcx.fraVISUAL fraVeiculo 
      Height          =   1800
      Left            =   30
      TabIndex        =   17
      Top             =   1875
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
      Begin VTOcx.txtVISUAL txtMunicipio 
         Height          =   480
         Left            =   135
         TabIndex        =   10
         Tag             =   "Município"
         Top             =   1305
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
      Begin VTOcx.cboVISUAL cboUFTransp 
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
      Begin VTOcx.txtVISUAL txtComplemento 
         Height          =   480
         Left            =   4560
         TabIndex        =   13
         Top             =   1320
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
   End
   Begin VTOcx.txtVISUAL txtcod 
      Height          =   285
      Left            =   255
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      Caption         =   "Ano"
      Text            =   ""
      Restricao       =   2
      AlinhamentoRotulo=   1
      CorRotulo       =   4210752
      CorTexto        =   4194304
      MaxLen          =   20
      MinLen          =   20
   End
End
Attribute VB_Name = "CDAP102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AparelhoTransporte  As New eAparelhoTransporte
Dim Ativ As New Atividade
    

Private Sub LimpaCampo()
txtVeiculo = ""
    txtMarca = ""
    txtModelo = ""
    txtAnoFabric = ""
    txtPlaca = ""
    txtMunicipio = ""
    cboUFTransp.ListIndex = -1
    txtLicenca = ""
    txtChassi = ""
    cboAtividadeVeiculo.ListIndex = -1
    txtInicioAtividadeCarro = ""
    txtComplemento = ""
End Sub

Private Sub cmdBuscar_Click()
   
    AparelhoTransporte.PreencherGrd grdVeiculo, txtIm
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

Private Sub cmdSalvar_Click()
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
      
    With AparelhoTransporte
            .icad = txtcod
            .Im = txtIm
            .Veiculo = txtVeiculo
            .Marca = txtMarca
            .Modelo = txtModelo
            .AnoFabricacao = txtAnoFabric
            .Placa = txtPlaca
            .Municipio = txtMunicipio
            .UF = CStr(cboUFTransp.Coluna(1).VALOR)
            .Licenca = txtLicenca
            .Chassi = txtChassi
            .Atividade = CStr(cboAtividadeVeiculo.Coluna(1).VALOR)
            .IniAtividadeCarro = txtInicioAtividadeCarro
            .Complemento = txtComplemento
        If .Salvar = True Then
            Avisa "Dados alterados com sucesso."
            LimpaCampo
            AparelhoTransporte.PreencherGrd grdVeiculo, txtIm
        End If
    End With
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     cboUFTransp.PreencherGeral Bdados, "UF"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
     Ativ.PreencherCboAtiv cboAtividadeVeiculo
End Sub


Private Sub grdVeiculo_DblClick()
    If grdVeiculo.ListItems.Count >= 1 Then
        fraVeiculo.Enabled = True
        txtcod = grdVeiculo.SelectedItem
        cboAtividadeVeiculo.SetarLinha grdVeiculo.SelectedItem.SubItems(14), 1
        txtVeiculo = grdVeiculo.SelectedItem.SubItems(2)
        txtMarca = grdVeiculo.SelectedItem.SubItems(3)
        txtModelo = grdVeiculo.SelectedItem.SubItems(4)
        txtAnoFabric = grdVeiculo.SelectedItem.SubItems(5)
        txtPlaca = grdVeiculo.SelectedItem.SubItems(6)
        txtChassi = grdVeiculo.SelectedItem.SubItems(7)
        txtLicenca = grdVeiculo.SelectedItem.SubItems(8)
        txtMunicipio = grdVeiculo.SelectedItem.SubItems(9)
        cboUFTransp.SetarLinha grdVeiculo.SelectedItem.SubItems(13), 1
        txtInicioAtividadeCarro = grdVeiculo.SelectedItem.SubItems(11)
        txtComplemento = grdVeiculo.SelectedItem.SubItems(12)
        txtIm = grdVeiculo.SelectedItem.SubItems(15)
        txtIm_LostFocus
  End If
End Sub

Private Sub txtIm_LostFocus()
      
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
   
End Sub
