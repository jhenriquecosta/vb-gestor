VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDOP101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDOP101"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1138
      Icone           =   "CDOP101.frx":0000
   End
   Begin VTOcx.fraVISUAL fraOcupacao 
      Height          =   2565
      Left            =   30
      TabIndex        =   20
      Top             =   1785
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   4524
      Altura          =   1905
      Caption         =   " Dados da Ocupação"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cboVISUAL cboUF 
         Height          =   510
         Left            =   7425
         TabIndex        =   14
         Tag             =   "UF"
         Top             =   1890
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
         TabIndex        =   15
         Tag             =   "CEP"
         Top             =   1905
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
      Begin VTOcx.cboVISUAL cboAtividade 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Tag             =   "AtividadeDesempenhada"
         Top             =   405
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   556
         Caption         =   "AtividadeDesempenhada"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   4210752
      End
      Begin VTOcx.txtVISUAL txtArea 
         Height          =   480
         Left            =   5835
         TabIndex        =   6
         Tag             =   "Área Ocupada (M2)"
         Top             =   795
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   847
         Caption         =   "Área Ocupada  (M2)"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   20
      End
      Begin VTOcx.cboVISUAL cboTipoOcupacao 
         Height          =   510
         Left            =   7590
         TabIndex        =   7
         Tag             =   "Tipo Ocupacao"
         Top             =   780
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   900
         Caption         =   "Tipo de  Ocupação"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtPerido 
         Height          =   480
         Left            =   4380
         TabIndex        =   5
         Tag             =   "Período"
         Top             =   795
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   847
         Caption         =   "Período  (meses)"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   20
      End
      Begin VTOcx.txtVISUAL txtInicio 
         Height          =   480
         Left            =   2970
         TabIndex        =   4
         Tag             =   "Inicio Ocupação"
         Top             =   795
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   847
         Caption         =   "Inicio Ocupação"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtObjeto 
         Height          =   480
         Left            =   150
         TabIndex        =   3
         Tag             =   "Objeto da Ocupação"
         Top             =   795
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   847
         Caption         =   "Objeto da Ocupação"
         Text            =   ""
         Restricao       =   1
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   50
      End
      Begin VTOcx.cboVISUAL cboLogr 
         Height          =   510
         Left            =   2595
         TabIndex        =   9
         Tag             =   "Logradouro (Nome)"
         Top             =   1335
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
         TabIndex        =   8
         Tag             =   "Logradouro"
         Top             =   1335
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
         TabIndex        =   10
         Top             =   1350
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
         Left            =   3735
         TabIndex        =   13
         Tag             =   "Cidade"
         Top             =   1920
         Width           =   3690
         _ExtentX        =   6509
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
         TabIndex        =   12
         Tag             =   "Distrito ou Bairro"
         Top             =   1905
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
         TabIndex        =   11
         Top             =   1350
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
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1050
      Left            =   45
      TabIndex        =   21
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   690
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   1852
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
         TabIndex        =   23
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
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   450
         TabIndex        =   22
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
      TabIndex        =   24
      Top             =   4395
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   7440
         TabIndex        =   16
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
         TabIndex        =   18
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
         TabIndex        =   17
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
End
Attribute VB_Name = "CDOP101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OcupacaoSoloPublico  As New eOcupacaoSoloPublico
Dim Endereco As New eEndereco
Dim Ativ As New Atividade
Private Sub LimpaCampo()
      txtObjeto = ""
      txtInicio = ""
      cboTipoOcupacao.ListIndex = -1
      txtPerido = ""
      txtArea = ""
      cboAtividade.ListIndex = -1
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
      
    With OcupacaoSoloPublico
        .Im = txtIm
        .ObjetoOcupacao = txtObjeto
        .TipoOcupacao = cboTipoOcupacao.Coluna(1).VALOR
        .DataOcupacao = txtInicio
        .Periodo = txtPerido
        .AreaOcupada = txtArea
        .AtividadeExercida = cboAtividade.Coluna(1).VALOR
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
     Ativ.PreencherCboAtiv cboAtividade
     cboUF.PreencherGeral Bdados, "UF"
     cboTipoOcupacao.Preencher Bdados, "select * from vis_Tipo_ocupacao_solo"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboTipoLogr_LostFocus()
    If cboTipoLogr.ListIndex <> -1 Then
        Endereco.PreencherCboRua cboLogr, CInt(cboTipoLogr.Coluna(1).VALOR)
    Else
        Endereco.PreencherCboRua cboLogr
    End If
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub


