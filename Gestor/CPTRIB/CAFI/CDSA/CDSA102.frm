VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDSA102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDSA102"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1138
      Icone           =   "CDSA102.frx":0000
   End
   Begin VTOcx.fraVISUAL fraSanitario 
      Height          =   2025
      Left            =   30
      TabIndex        =   12
      Top             =   1815
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   3572
      Altura          =   1905
      Caption         =   " Dados do Cadastro"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cboVISUAL cboUF 
         Height          =   510
         Left            =   7425
         TabIndex        =   8
         Tag             =   "UF"
         Top             =   885
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   900
         Caption         =   "UF"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtCep 
         Height          =   480
         Left            =   8370
         TabIndex        =   9
         Tag             =   "CEP"
         Top             =   900
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   847
         Caption         =   "CEP"
         Text            =   ""
         Formato         =   4
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   10
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboLogr 
         Height          =   510
         Left            =   2595
         TabIndex        =   3
         Tag             =   "Logradouro"
         Top             =   330
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   900
         Caption         =   "Logradouro (Nome)"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         Editavel        =   -1  'True
      End
      Begin VTOcx.cboVISUAL cboTipoLogr 
         Height          =   510
         Left            =   180
         TabIndex        =   2
         Tag             =   "Logradouro"
         Top             =   330
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   900
         Caption         =   "Logradouro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtNum 
         Height          =   480
         Left            =   6525
         TabIndex        =   4
         Top             =   345
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   847
         Caption         =   "Nº"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtCidade 
         Height          =   480
         Left            =   3765
         TabIndex        =   7
         Tag             =   "Cidade"
         Top             =   915
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   847
         Caption         =   "Cidade"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   80
      End
      Begin VTOcx.cboVISUAL cboBairro 
         Height          =   510
         Left            =   150
         TabIndex        =   6
         Tag             =   "Distrito ou Bairro"
         Top             =   900
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   900
         Caption         =   "Distrito ou Bairro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtComplemento 
         Height          =   480
         Left            =   7215
         TabIndex        =   5
         Top             =   345
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   847
         Caption         =   "Complemento"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   80
      End
      Begin VTOcx.txtVISUAL txtMotivo 
         Height          =   480
         Left            =   150
         TabIndex        =   10
         Tag             =   "Motivo"
         Top             =   1455
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   847
         Caption         =   "Motivo"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   100
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1050
      Left            =   45
      TabIndex        =   13
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   720
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
         TabIndex        =   15
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
         Left            =   3150
         TabIndex        =   1
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
   Begin VTOcx.grdVISUAL grdSanitario 
      Height          =   2970
      Left            =   30
      TabIndex        =   16
      Top             =   3930
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   5239
      CorBorda        =   32768
      Caption         =   "Cadastros"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   17
      Top             =   6930
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   6450
         TabIndex        =   22
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
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   7440
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
   Begin VTOcx.txtVISUAL txtcod 
      Height          =   285
      Left            =   255
      TabIndex        =   21
      Top             =   855
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
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
Attribute VB_Name = "CDSA102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sanitario  As New eSanitario
Dim Endereco As New eEndereco
Dim Ativ As New Atividade
Private Sub LimpaCampo()
      txtMotivo = ""
      cboTipoLogr.ListIndex = -1
      cboLogr.ListIndex = -1
      txtNum = ""
      cboBairro.ListIndex = -1
      txtComplemento = ""
      txtCidade = ""
      cboUF.ListIndex = -1
      txtCep = ""
End Sub





Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdSanitario.ListItems.Clear
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
      
    With Sanitario
        .Icod = txtcod
        .Im = txtIm
        .Motivo = txtMotivo
        .Logradouro = cboTipoLogr
        .NomeLogradouro = cboLogr
        .Numero = txtNum
        .Bairro = cboBairro
        .Complemento = txtComplemento
        .Cidade = txtCidade
        .UF = cboUF.Coluna(1).VALOR
        .Cep = txtCep
        If .Salvar = True Then
            Avisa "Dados Salvos com Sucesso."
            LimpaCampo
            Sanitario.PreencherGrd grdSanitario, txtIm
        End If
    End With
    Screen.MousePointer = 0
End Sub
    

Private Sub Form_Load()
     Dim Sql As String
     
    Sql = "select TLG_NOME, "
    Sql = Sql & " TLG_COD_LOGRADOURO"
    Sql = Sql & " from vis_logradouro  "
    Sql = Sql & " order by TLG_NOME "
    cboLogr.Preencher Bdados, Sql
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     With Endereco
        .PreencherCboTipoLogr cboTipoLogr
        .PreencherCboBairro cboBairro
     End With
     cboUF.PreencherGeral Bdados, "UF"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub grdSanitario_DblClick()
    If grdSanitario.ListItems.Count >= 1 Then
        fraSanitario.Enabled = True
        txtIm = grdSanitario.SelectedItem.SubItems(11)
        txtIm_LostFocus
        txtcod = grdSanitario.SelectedItem
        txtMotivo = grdSanitario.SelectedItem.SubItems(1)
        cboTipoLogr = grdSanitario.SelectedItem.SubItems(2)
        cboLogr = grdSanitario.SelectedItem.SubItems(3)
        txtNum = grdSanitario.SelectedItem.SubItems(4)
        txtComplemento = grdSanitario.SelectedItem.SubItems(5)
        cboBairro = grdSanitario.SelectedItem.SubItems(6)
        txtCep = grdSanitario.SelectedItem.SubItems(7)
        txtCidade = grdSanitario.SelectedItem.SubItems(8)
        cboUF.SetarLinha grdSanitario.SelectedItem.SubItems(10), 1
    End If
     
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
   
End Sub
Private Sub cmdBuscar_Click()
  Sanitario.PreencherGrd grdSanitario, txtIm
End Sub

