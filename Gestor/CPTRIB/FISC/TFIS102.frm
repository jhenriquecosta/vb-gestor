VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TFIS102 
   BackColor       =   &H00FFF5EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   4680
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   10
      Top             =   7260
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   900
      CorFrente       =   0
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9000
         TabIndex        =   12
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   0
         CorFrente       =   0
      End
   End
   Begin ActiveTabs.SSActiveTabs tabGeral 
      Height          =   6420
      Left            =   60
      TabIndex        =   14
      Top             =   720
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   11324
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "TFIS102.frx":0000
      Images          =   "TFIS102.frx":00D7
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6000
         Index           =   0
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   10583
         _Version        =   131082
         TabGuid         =   "TFIS102.frx":072A
         Begin VTOcx.txtVISUAL txtDescProcesso 
            Height          =   315
            Left            =   195
            TabIndex        =   0
            Top             =   4350
            Width           =   7875
            _ExtentX        =   13891
            _ExtentY        =   556
            Caption         =   "Descrição"
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdSalvarProcesso 
            Height          =   405
            Left            =   8865
            TabIndex        =   4
            Top             =   5505
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdExcluirSistema 
            Height          =   405
            Left            =   9255
            TabIndex        =   5
            Top             =   5505
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.grdVISUAL grdProcessos 
            Height          =   4260
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   9525
            _ExtentX        =   16801
            _ExtentY        =   7514
            CorFundo        =   -2147483633
            Caption         =   "Processos"
            CorTitulo       =   4210688
         End
         Begin VTOcx.cmdVISUAL cmdLimparProcesso 
            Height          =   405
            Left            =   8475
            TabIndex        =   3
            Top             =   5505
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   6
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtPrazoProcesso 
            Height          =   300
            Left            =   45
            TabIndex        =   1
            Top             =   4710
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            Caption         =   "Prazo(Dias)"
            Text            =   ""
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            CorRotulo       =   0
            AgruparValores  =   0   'False
         End
         Begin VTOcx.cboVISUAL cboRespProcesso 
            Height          =   315
            Left            =   1950
            TabIndex        =   2
            Top             =   4710
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   556
            Caption         =   "Responsável"
            Text            =   ""
            AutoFocaliza    =   0   'False
            CorFundo        =   -2147483633
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6000
         Index           =   2
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   10583
         _Version        =   131082
         TabGuid         =   "TFIS102.frx":0752
         Begin VTOcx.cmdVISUAL cmdSalvarProcedimento 
            Height          =   405
            Left            =   8865
            TabIndex        =   25
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdExcluirProcedimento 
            Height          =   405
            Left            =   9255
            TabIndex        =   26
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.grdVISUAL grdProcedimentos 
            Height          =   4200
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   7408
            CorFundo        =   -2147483633
            Caption         =   "Procedimentos"
            CorTitulo       =   4210688
         End
         Begin VTOcx.cmdVISUAL cmdLimparProcedimento 
            Height          =   405
            Left            =   8475
            TabIndex        =   24
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   6
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtDescProcedimento 
            Height          =   315
            Left            =   195
            TabIndex        =   19
            Top             =   4290
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   556
            Caption         =   "Descrição"
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtPrazoProcedimento 
            Height          =   300
            Left            =   30
            TabIndex        =   20
            Top             =   4650
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            Caption         =   "Prazo(Dias)"
            Text            =   ""
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            CorRotulo       =   0
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtRelatorio 
            Height          =   495
            Left            =   1050
            TabIndex        =   22
            Top             =   4950
            Width           =   6585
            _ExtentX        =   11615
            _ExtentY        =   873
            Caption         =   "Caminho do Relatorio"
            Text            =   ""
            AlinhamentoRotulo=   1
            CorFundo        =   -2147483633
            CorRotulo       =   0
            AgruparValores  =   0   'False
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.cboVISUAL cboAutoridade 
            Height          =   315
            Left            =   2010
            TabIndex        =   21
            Top             =   4650
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   556
            Caption         =   "Responsável"
            Text            =   ""
            AutoFocaliza    =   0   'False
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdConsultaArquivo 
            Height          =   315
            Left            =   7740
            TabIndex        =   30
            Top             =   5190
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            Caption         =   ""
            Acao            =   5
            CorBorda        =   -2147483645
            CorFoco         =   -2147483628
         End
         Begin VTOcx.txtVISUAL txtOrdemProcedimento 
            Height          =   480
            Left            =   1020
            TabIndex        =   23
            Top             =   5430
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   847
            Caption         =   "Ordem"
            Text            =   ""
            Restricao       =   2
            AlinhamentoRotulo=   1
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            CorRotulo       =   0
            AgruparValores  =   0   'False
         End
         Begin VTOcx.cboVISUAL cboParametro 
            Height          =   510
            Left            =   1890
            TabIndex        =   31
            Top             =   5430
            Width           =   6285
            _ExtentX        =   11086
            _ExtentY        =   900
            Caption         =   "Relatos/Fundamentos Fiscais"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
            CorFundo        =   -2147483633
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6000
         Index           =   1
         Left            =   30
         TabIndex        =   27
         Top             =   30
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   10583
         _Version        =   131082
         TabGuid         =   "TFIS102.frx":077A
         Begin VTOcx.cmdVISUAL cmdSalvarFase 
            Height          =   405
            Left            =   8865
            TabIndex        =   11
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdExcluirFase 
            Height          =   405
            Left            =   9255
            TabIndex        =   13
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.grdVISUAL grdFases 
            Height          =   4260
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   7514
            CorFundo        =   -2147483633
            Caption         =   "Fases"
            CorTitulo       =   4210688
         End
         Begin VTOcx.cmdVISUAL cmdLimparFase 
            Height          =   405
            Left            =   8475
            TabIndex        =   9
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   6
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtDescFase 
            Height          =   315
            Left            =   195
            TabIndex        =   6
            Top             =   4350
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   556
            Caption         =   "Descrição"
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtPrazoFase 
            Height          =   300
            Left            =   30
            TabIndex        =   7
            Top             =   4710
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            Caption         =   "Prazo(Dias)"
            Text            =   ""
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            CorRotulo       =   0
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtOrdemFase 
            Height          =   300
            Left            =   2130
            TabIndex        =   8
            ToolTipText     =   "Padrão = 1"
            Top             =   4710
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   529
            Caption         =   "Ordem"
            Text            =   ""
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            CorRotulo       =   0
            AgruparValores  =   0   'False
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1138
      Icone           =   "TFIS102.frx":07A2
   End
End
Attribute VB_Name = "TFIS102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strCodSistema As String
Private strCodModulo As String
Private Fisc As New Fiscalizacao
Private Conta As New ContaCorrente

Private CodProcesso As String
Private CodFase As String
Private CodProcedimento As String

Public Function ChamaArquivos(Dialog As Object) As String
    On Error Resume Next
    With Dialog
        .CancelError = True
        .DialogTitle = "Arquivo de Relatório."
        .InitDir = "C:\Arquivos de programas\Sistema Gestor Municipal\"
        .Filter = "Arquivos do Crystal Report|*.RPT"
        .ShowOpen
        If .FileName <> "" Then
            ChamaArquivos = .FileName
        End If
    End With
End Function
Private Sub cmdConsultaArquivo_Click()
    txtRelatorio = ChamaArquivos(Dialogo)
End Sub

Private Sub cmdExcluirFase_Click()
    If Fisc.Rede.rCodEtapa <> 0 Then
        If Confirma("Confirma exclusão do registro?") Then
            If Fisc.Rede.ApagaEtapaRede(Fisc.Rede.rCodEtapa) Then
                Avisa "Registro excluído."
                PrepararFase
                CodFase = ""
            End If
        End If
    End If
End Sub

Private Sub cmdExcluirProcedimento_Click()
    If Fisc.Rede.rCodEtapa <> 0 Then
        If Confirma("Confirma exclusão do registro?") Then
            If Fisc.Rede.ApagaEtapaRede(Fisc.Rede.rCodEtapa) Then
                Avisa "Registro excluído."
                PrepararProcedimento
                CodProcedimento = ""
            End If
        End If
    End If
End Sub

Private Sub cmdExcluirSistema_Click()
    If Fisc.Rede.rCodEtapa <> 0 Then
        If Confirma("Confirma exclusão do registro?") Then
            If Fisc.Rede.ApagaEtapaRede(Fisc.Rede.rCodEtapa) Then
                Avisa "Registro excluído."
                PrepararProcesso
                CodProcesso = ""
            End If
        End If
    End If
End Sub

Private Sub cmdLimparFase_Click()
    txtDescFase = ""
    txtPrazoFase = ""
    CodFase = ""
    txtOrdemFase = 1
    Set Fisc = Nothing
    Set Fisc = New Fiscalizacao
    txtDescFase.SetFocus
    Fisc.Rede.LimpaDadosRede
End Sub

Private Sub cmdLimparProcedimento_Click()
    txtDescProcedimento = ""
    txtPrazoProcedimento = ""
    CodProcedimento = ""
    txtRelatorio = ""
    txtOrdemProcedimento = grdProcedimentos.AtualizarQtd + 1
    cboAutoridade.ListIndex = -1
    cboParametro.ListIndex = -1
    Set Fisc = Nothing
    Set Fisc = New Fiscalizacao
    txtDescProcedimento.SetFocus
    Fisc.Rede.LimpaDadosRede
End Sub


Private Sub cmdLimparProcesso_Click()
    txtDescProcesso = ""
    txtPrazoProcesso = ""
    CodProcesso = ""
    cboRespProcesso.ListIndex = -1
    Set Fisc = Nothing
    Set Fisc = New Fiscalizacao
    txtDescProcesso.SetFocus
    Fisc.Rede.LimpaDadosRede
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvarFase_Click()
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If CDbl(Nvl(CodFase, 0)) = 0 Then
        Fisc.Rede.rCodEtapa = Conta.GeraCodPagamento("36")
    Else
        Fisc.Rede.rCodEtapa = CodFase
    End If
    
    Fisc.Rede.rCaminhoRpt = ""
    Fisc.Rede.rCodEtapaPai = 0
    Fisc.Rede.rCodEtapaOrigem = grdProcessos.SelectedItem
    Fisc.Rede.rCodFuncionario = 0
    Fisc.Rede.rDescricao = txtDescFase
    Fisc.Rede.rOrdem = txtOrdemFase
    Fisc.Rede.rPrazo = Nvl(txtPrazoProcesso, 0)
    Fisc.Rede.rTipoEtapa = etrFase
    
    If Fisc.Rede.CriaEtapaRede(Fisc.Rede.rCodEtapa, Fisc.Rede.rDescricao, _
            Fisc.Rede.rCodEtapaPai, Fisc.Rede.rCodEtapaOrigem, _
            Fisc.Rede.rOrdem, Fisc.Rede.rCodFuncionario, Fisc.Rede.rCaminhoRpt, _
            Fisc.Rede.rTipoEtapa, Fisc.Rede.rPrazo) Then
        Avisa "Dados gravado com sucesso."
        PrepararFase
    End If
    CodFase = ""
End Sub

Private Sub cmdSalvarProcedimento_Click()
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If CDbl(Nvl(CodProcedimento, 0)) = 0 Then
        Fisc.Rede.rCodEtapa = Conta.GeraCodPagamento("36")
    Else
        Fisc.Rede.rCodEtapa = CodProcedimento
    End If
    
    Fisc.Rede.rCaminhoRpt = txtRelatorio
    Fisc.Rede.rCodEtapaPai = grdFases.SelectedItem
    Fisc.Rede.rCodEtapaOrigem = grdProcessos.SelectedItem
    Fisc.Rede.rCodFuncionario = CDbl(Nvl(CStr(cboAutoridade.Coluna(0).Valor), 0))
    Fisc.Rede.rDescricao = txtDescProcedimento
    Fisc.Rede.rOrdem = txtOrdemProcedimento
    Fisc.Rede.rPrazo = Nvl(txtPrazoProcedimento, 0)
    Fisc.Rede.rTipoEtapa = etrProcedimento
    If cboParametro.ListIndex >= 0 Then Fisc.Rede.rCodParametroFundamento = cboParametro.Coluna(0).Valor
    If Fisc.Rede.CriaEtapaRede(Fisc.Rede.rCodEtapa, Fisc.Rede.rDescricao, _
            Fisc.Rede.rCodEtapaPai, Fisc.Rede.rCodEtapaOrigem, _
            Fisc.Rede.rOrdem, Fisc.Rede.rCodFuncionario, Fisc.Rede.rCaminhoRpt, _
            Fisc.Rede.rTipoEtapa, Fisc.Rede.rPrazo, CDbl(Nvl(CStr(cboParametro.Coluna(0).Valor), 0))) Then
        Avisa "Dados gravado com sucesso."
        PrepararProcedimento
    End If
    CodProcedimento = ""
End Sub

Private Sub cmdSalvarProcesso_Click()
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If CDbl(Nvl(CodProcesso, 0)) = 0 Then
        Fisc.Rede.rCodEtapa = Conta.GeraCodPagamento("36")
    Else
        Fisc.Rede.rCodEtapa = CodProcesso
    End If
    
    Fisc.Rede.rCaminhoRpt = ""
    Fisc.Rede.rCodEtapaPai = 0
    Fisc.Rede.rCodEtapaOrigem = 0
    Fisc.Rede.rCodFuncionario = CDbl(Nvl(CStr(cboRespProcesso.Coluna(0).Valor), 0))
    Fisc.Rede.rDescricao = txtDescProcesso
    Fisc.Rede.rOrdem = 0
    Fisc.Rede.rPrazo = Nvl(txtPrazoProcesso, 0)
    Fisc.Rede.rTipoEtapa = etrProcesso
    
    If Fisc.Rede.CriaEtapaRede(Fisc.Rede.rCodEtapa, Fisc.Rede.rDescricao, _
            Fisc.Rede.rCodEtapaPai, Fisc.Rede.rCodEtapaOrigem, _
            Fisc.Rede.rOrdem, Fisc.Rede.rCodFuncionario, Fisc.Rede.rCaminhoRpt, _
            Fisc.Rede.rTipoEtapa, Fisc.Rede.rPrazo) Then
        Avisa "Dados gravado com sucesso."
        PrepararProcesso
    End If
    CodProcesso = ""
End Sub
Private Sub Form_Load()
    Dim Parametro As New Parametros
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    PrepararProcesso
    Fisc.Funcionario.PreencheComboFuncionario cboAutoridade
    Fisc.Funcionario.PreencheComboFuncionario cboRespProcesso
    Parametro.PreencheCombo cboParametro
End Sub

Private Sub PrepararProcesso()
    Fisc.Rede.PreencheGridRedeEtapas grdProcessos, etrProcesso
    grdFases.Preencher Bdados, ""
    grdProcedimentos.Preencher Bdados, ""
    txtDescProcesso = ""
    txtPrazoProcesso = ""
    cboRespProcesso.ListIndex = -1
End Sub

Private Sub PrepararFase()
    Fisc.Rede.PreencheGridRedeEtapas grdFases, etrFase, grdProcessos.SelectedItem
    grdProcedimentos.Preencher Bdados, ""
    txtDescFase = ""
    txtPrazoFase = ""
End Sub

Private Sub PrepararProcedimento()
    Fisc.Rede.PreencheGridRedeEtapas grdProcedimentos, etrProcedimento, grdFases.SelectedItem
    txtDescProcedimento = ""
    txtPrazoProcedimento = ""
    cboAutoridade.ListIndex = -1
    txtRelatorio = ""
    txtOrdemProcedimento = grdProcedimentos.AtualizarQtd + 1
    cboAutoridade.ListIndex = -1
End Sub

Private Sub grdFases_Click()
    If grdFases.ListItems.Count = 0 Then Exit Sub
    Fisc.Rede.CarregaDadosRede grdFases.SelectedItem
    CodFase = grdFases.SelectedItem
    txtDescFase = Fisc.Rede.rDescricao
    txtPrazoFase = Fisc.Rede.rPrazo
    txtOrdemFase = Fisc.Rede.rOrdem
    PrepararProcedimento
End Sub

Private Sub grdFases_DblClick()
    If grdFases.ListItems.Count = 0 Then Exit Sub
    tabGeral.Tabs(3).Selected = True
    txtDescProcedimento.SetFocus
End Sub

Private Sub grdProcedimentos_Click()
    If grdProcedimentos.ListItems.Count = 0 Then Exit Sub
    cmdLimparProcedimento_Click
    Fisc.Rede.CarregaDadosRede grdProcedimentos.SelectedItem
    
    CodProcedimento = grdProcedimentos.SelectedItem
    txtDescProcedimento = Fisc.Rede.rDescricao
    txtPrazoProcedimento = Fisc.Rede.rPrazo
    txtOrdemProcedimento = Fisc.Rede.rOrdem
    txtRelatorio = Fisc.Rede.rCaminhoRpt
    cboAutoridade.SetarLinha Fisc.Rede.rCodFuncionario, 0
    cboParametro.SetarLinha Fisc.Rede.rCodParametroFundamento, 0
End Sub

Private Sub grdProcessos_Click()
    If grdProcessos.ListItems.Count = 0 Then Exit Sub
    Fisc.Rede.CarregaDadosRede grdProcessos.SelectedItem
    CodProcesso = grdProcessos.SelectedItem
    txtDescProcesso = Fisc.Rede.rDescricao
    txtPrazoProcesso = Fisc.Rede.rPrazo
    cboRespProcesso.SetarLinha Fisc.Rede.rCodFuncionario, 0
    PrepararFase
End Sub

Private Sub grdProcessos_DblClick()
    If grdProcessos.ListItems.Count = 0 Then Exit Sub
    tabGeral.Tabs(2).Selected = True
    txtDescFase.SetFocus
End Sub


