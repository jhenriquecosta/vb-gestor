VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONT~1.OCX"
Begin VB.Form TFIS201 
   Caption         =   "TPRT101"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   1138
      Icone           =   "TFIS201.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   10
      Top             =   4380
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   7035
         TabIndex        =   7
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   8025
         TabIndex        =   8
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   6045
         TabIndex        =   6
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
      End
   End
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   3990
      Left            =   -30
      TabIndex        =   11
      Tag             =   "Documento gerencial"
      Top             =   330
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   7038
      _Version        =   131082
      TabCount        =   1
      TabOrientation  =   2
      Tabs            =   "TFIS201.frx":031A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3600
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   6350
         _Version        =   131082
         TabGuid         =   "TFIS201.frx":036F
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   6105
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   11460
            _ExtentX        =   20214
            _ExtentY        =   10769
            Altura          =   1905
            Caption         =   " Livro Fiscal (Modelos Diferentes)"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VTOcx.fraVISUAL fraVISUAL3 
               Height          =   1995
               Left            =   90
               TabIndex        =   18
               ToolTipText     =   "Pesquisa Contribuintes"
               Top             =   1515
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   3519
               Altura          =   1905
               Caption         =   " Formalização"
               CorTexto        =   16777215
               CorFaixa        =   32768
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.cboVISUAL cboFase 
                  Height          =   510
                  Left            =   60
                  TabIndex        =   5
                  Top             =   1350
                  Width           =   6645
                  _ExtentX        =   11721
                  _ExtentY        =   900
                  Caption         =   "Fase Processual"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.txtVISUAL txtInicio 
                  Height          =   495
                  Left            =   60
                  TabIndex        =   1
                  Top             =   330
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   873
                  Caption         =   "Data Início"
                  Text            =   ""
                  Formato         =   0
                  Restricao       =   2
                  AlinhamentoRotulo=   1
                  CorRotulo       =   0
                  AgruparValores  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtFim 
                  Height          =   495
                  Left            =   1890
                  TabIndex        =   2
                  Top             =   330
                  Width           =   1905
                  _ExtentX        =   3360
                  _ExtentY        =   873
                  Caption         =   "Data Término"
                  Text            =   ""
                  Formato         =   0
                  Restricao       =   2
                  AlinhamentoRotulo=   1
                  CorRotulo       =   0
                  AgruparValores  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtPeriodoFiscalizado 
                  Height          =   495
                  Left            =   3810
                  TabIndex        =   3
                  Top             =   330
                  Width           =   2835
                  _ExtentX        =   5001
                  _ExtentY        =   873
                  Caption         =   "Período Fiscalizado"
                  Text            =   ""
                  AlinhamentoRotulo=   1
                  CorRotulo       =   0
                  AgruparValores  =   0   'False
               End
               Begin VTOcx.cboVISUAL cboAutoridade 
                  Height          =   510
                  Left            =   60
                  TabIndex        =   4
                  Top             =   840
                  Width           =   8685
                  _ExtentX        =   15319
                  _ExtentY        =   900
                  Caption         =   "Autoridade Fiscal"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
            End
            Begin VTOcx.fraVISUAL fraProPrietario 
               Height          =   1125
               Left            =   90
               TabIndex        =   14
               ToolTipText     =   "Pesquisa Contribuintes"
               Top             =   330
               Width           =   8925
               _ExtentX        =   15743
               _ExtentY        =   1984
               Altura          =   1905
               Caption         =   " Qualificação do Contribuinte"
               CorTexto        =   16777215
               CorFaixa        =   32768
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.txtVISUAL txtEndereco 
                  Height          =   300
                  Left            =   465
                  TabIndex        =   17
                  Top             =   720
                  Width           =   8400
                  _ExtentX        =   14817
                  _ExtentY        =   529
                  Caption         =   "Endereço"
                  Text            =   ""
                  Enabled         =   0   'False
                  Requerido       =   0   'False
                  CorRotulo       =   0
                  CorTexto        =   4194304
                  RetirarMascara  =   0   'False
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
                  Restricao       =   2
                  CorRotulo       =   0
                  AgruparValores  =   0   'False
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtNomeContrib 
                  Height          =   285
                  Left            =   3120
                  TabIndex        =   16
                  Top             =   390
                  Width           =   5730
                  _ExtentX        =   10107
                  _ExtentY        =   503
                  Caption         =   "Nome"
                  Text            =   ""
                  Enabled         =   0   'False
                  CorRotulo       =   0
                  CorTexto        =   4194304
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.cmdVISUAL cmdBUsca 
                  Height          =   285
                  Left            =   2760
                  TabIndex        =   15
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
         End
      End
   End
End
Attribute VB_Name = "TFIS201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rpt As New VSRelatorio

Private Sub cmdBUsca_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtIm.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub
Private Sub ImprimirFicha()
Dim Selecao As String
    Dim i As Integer
    Dim rs As VSRecordset
    
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TProcessoFicha.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Formulas "vt_texto_cabecalho", "FICHA DE PROTOCOLO"
        .Formulas "VT_DATA", AplicacoesVTFuncoes.municipio & " -  " & Temp.PegaParametro(Bdados, "ESTADO CLIENTE") & "  " & FormatDateTime(Date, vbLongDate)
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Titulo = "Ficha de Processo"
        .SubRelatorio = "TSubProcesso.rpt"
        'TPR_ACAO,TPR_SUBACAO
        .Selecao = ""
        .SubRelatorio = ""
        .Selecao = Selecao
        .Imprimir
        
    End With
End Sub

Private Sub cmdSalvar_Click()
    Dim Fisc As New Fiscalizacao
    If Fisc.CriaFiscalizacao(txtIm, txtInicio, _
        txtFim, txtPeriodoFiscalizado, CStr(cboAutoridade.Coluna(0).Valor), CStr(cboFase.Coluna(0).Valor)) Then
        LimpaCampos Me
        Avisa "Dados gravados com sucesso."
        txtIm.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Fisc As New Fiscalizacao
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Fisc.Funcionario.PreencheComboFuncionario cboAutoridade
    Fisc.Rede.PreencheComboEtapas cboFase, etrProcesso
End Sub

Private Sub txtIm_LostFocus()
    Dim Ic As String
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtIm) = 10 Or Len(txtIm) = 11 Then
            Ic = Imposto.FormataInscricao(txtIm, InscContrib)
        Else
            Ic = txtIm
        End If
    Else
            Ic = txtIm
    End If
    txtIm = BuscaContribuinte(Ic, txtNomeContrib, txtEndereco)
End Sub
