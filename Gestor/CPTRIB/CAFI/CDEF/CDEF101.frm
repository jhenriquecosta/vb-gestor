VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDEF101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDEF101"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   ControlBox      =   0   'False
   Icon            =   "CDEF101.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   14
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "CDEF101.frx":058A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   1138
      Icone           =   "CDEF101.frx":26AD
   End
   Begin VTOcx.fraVISUAL fraSanitario 
      Height          =   1980
      Left            =   30
      TabIndex        =   16
      Top             =   1770
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   3493
      Altura          =   1905
      Caption         =   " Dados do Ponto"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   510
         Left            =   180
         TabIndex        =   2
         Tag             =   "Tipo"
         ToolTipText     =   "784 - TIPO AMBULANTE"
         Top             =   300
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   900
         Caption         =   "Tipo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.cboVISUAL cboUF 
         Height          =   510
         Left            =   7425
         TabIndex        =   9
         Tag             =   "UF"
         Top             =   1410
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
         TabIndex        =   10
         Tag             =   "CEP"
         Top             =   1425
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
         TabIndex        =   4
         Tag             =   "Logradouro (nome)"
         Top             =   855
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   900
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         Editavel        =   -1  'True
      End
      Begin VTOcx.cboVISUAL cboTipoLogr 
         Height          =   510
         Left            =   165
         TabIndex        =   3
         Tag             =   "Logradouro"
         Top             =   855
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
         TabIndex        =   5
         Top             =   870
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
         Left            =   3750
         TabIndex        =   8
         Tag             =   "Cidade"
         Top             =   1440
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
         TabIndex        =   7
         Tag             =   "Distrito ou Bairro"
         Top             =   1425
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
         TabIndex        =   6
         Top             =   870
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   847
         Caption         =   "Complemento"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   100
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1035
      Left            =   45
      TabIndex        =   17
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   660
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   1826
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
         TabIndex        =   19
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
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   450
         TabIndex        =   18
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
      Height          =   450
      Left            =   0
      TabIndex        =   20
      Top             =   3810
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   8430
         TabIndex        =   12
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
         Left            =   9420
         TabIndex        =   13
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
         Left            =   7440
         TabIndex        =   11
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
End
Attribute VB_Name = "CDEF101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AmbulanteEventual  As New eAmbulanteEventual
Dim Endereco As New eEndereco
Dim Ativ As New atividade
Private Sub LimpaCampo()
      cboTipo.ListIndex = -1
      cboTipoLogr.ListIndex = -1
      cboLogr.ListIndex = -1
      txtNum = ""
      cboBairro.ListIndex = -1
      txtComplemento = ""
      txtCidade = ""
      cboUF.ListIndex = -1
      txtCep = ""
End Sub




Private Sub cboTipoLogr_LostFocus()
    If cboTipoLogr.ListIndex <> -1 Then
        Endereco.PreencherCboRua cboLogr, CInt(cboTipoLogr.Coluna(1).VALOR)
    Else
        Endereco.PreencherCboRua cboLogr
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
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
      
    With AmbulanteEventual
        .Im = txtIm
        .Tipo = cboTipo.Coluna(1).VALOR
        .Logradouro = cboTipoLogr.Coluna(0).VALOR
        .NomeLogradouro = cboLogr
        .Numero = txtNum
        .Bairro = cboBairro
        .Complemento = txtComplemento
        .Cidade = txtCidade
        .Uf = cboUF.Coluna(1).VALOR
        .Cep = txtCep
        If .Salvar = True Then
            Avisa "Dados Salvos com sucesso."
            LimpaCampo
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
        .PreencherCboRua cboLogr
     End With
     
     cboUF.PreencherGeral Bdados, "UF"
     cboTipo.PreencherGeral Bdados, "TIPO AMBULANTE"
     
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub



