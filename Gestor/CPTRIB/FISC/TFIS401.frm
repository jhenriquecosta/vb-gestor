VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TFIS401 
   Caption         =   "TPRT101"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   1138
      Icone           =   "TFIS401.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   6420
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   7035
         TabIndex        =   0
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
         TabIndex        =   1
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
      End
   End
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   6090
      Left            =   -30
      TabIndex        =   4
      Tag             =   "Documento gerencial"
      Top             =   330
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   10742
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      Tabs            =   "TFIS401.frx":031A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   5700
         Index           =   0
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   10054
         _Version        =   131082
         TabGuid         =   "TFIS401.frx":03E6
         Begin VTOcx.fraVISUAL fraProPrietario 
            Height          =   1095
            Left            =   30
            TabIndex        =   19
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   420
            Width           =   8925
            _ExtentX        =   15743
            _ExtentY        =   1931
            Altura          =   1905
            Caption         =   " Qualificação do Contribuinte"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Enabled         =   0   'False
            Begin VTOcx.txtVISUAL txtEndereco 
               Height          =   300
               Left            =   465
               TabIndex        =   22
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
               TabIndex        =   21
               Top             =   375
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   503
               Caption         =   "Ins. Municipal"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               CorRotulo       =   0
               AgruparValores  =   0   'False
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNomeContrib 
               Height          =   285
               Left            =   2790
               TabIndex        =   20
               Top             =   390
               Width           =   6060
               _ExtentX        =   10689
               _ExtentY        =   503
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               CorRotulo       =   0
               CorTexto        =   4194304
               RetirarMascara  =   0   'False
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   1455
            Left            =   30
            TabIndex        =   23
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   1590
            Width           =   8940
            _ExtentX        =   15769
            _ExtentY        =   2566
            Altura          =   1905
            Caption         =   " Levantamento Homologatório"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Enabled         =   0   'False
            Begin VTOcx.txtVISUAL txtInicio 
               Height          =   495
               Left            =   60
               TabIndex        =   28
               Top             =   330
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   873
               Caption         =   "Data Início"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtFim 
               Height          =   495
               Left            =   1890
               TabIndex        =   27
               Top             =   330
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   873
               Caption         =   "Data Término"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDevolucao 
               Height          =   495
               Left            =   3870
               TabIndex        =   26
               Top             =   330
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   873
               Caption         =   "Data Devolução"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtPeriodoFiscalizado 
               Height          =   495
               Left            =   5970
               TabIndex        =   25
               Top             =   330
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   873
               Caption         =   "Período Fiscalizado"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboAutoridade 
               Height          =   510
               Left            =   60
               TabIndex        =   24
               Top             =   840
               Width           =   8775
               _ExtentX        =   15478
               _ExtentY        =   900
               Caption         =   "Autoridade Fiscal"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Enabled         =   0   'False
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5700
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   10054
         _Version        =   131082
         TabGuid         =   "TFIS401.frx":040E
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   6225
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   10980
            Altura          =   1905
            Caption         =   " Nota Fiscal"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VTOcx.grdVISUAL grdAndamento 
               Height          =   3840
               Left            =   30
               TabIndex        =   29
               Top             =   990
               Width           =   9030
               _ExtentX        =   15928
               _ExtentY        =   6773
               CorBorda        =   32768
               Caption         =   "Andamento"
               CorTitulo       =   32768
               CorCaption      =   16777215
               CorDica         =   32768
               OcultarRodape   =   -1  'True
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5700
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   10054
         _Version        =   131082
         TabGuid         =   "TFIS401.frx":0436
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   6555
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   11460
            _ExtentX        =   20214
            _ExtentY        =   11562
            Altura          =   1905
            Caption         =   " Livro Fiscal (Modelos Diferentes)"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VTOcx.grdVISUAL grdProcedimentos 
               Height          =   2160
               Left            =   0
               TabIndex        =   30
               Top             =   3750
               Width           =   9030
               _ExtentX        =   15928
               _ExtentY        =   3810
               CorBorda        =   32768
               Caption         =   "Procedimentos"
               CorTitulo       =   32768
               CorCaption      =   16777215
               CorDica         =   32768
               OcultarRodape   =   -1  'True
            End
            Begin VTOcx.grdVISUAL grdFisc 
               Height          =   2130
               Left            =   0
               TabIndex        =   17
               Top             =   1860
               Width           =   9030
               _ExtentX        =   15928
               _ExtentY        =   3757
               CorBorda        =   32768
               Caption         =   "Fiscalizações"
               CorTitulo       =   32768
               CorCaption      =   16777215
               CorDica         =   32768
               OcultarRodape   =   -1  'True
            End
            Begin VTOcx.fraVISUAL fraVISUAL4 
               Height          =   1485
               Left            =   0
               TabIndex        =   9
               ToolTipText     =   "Pesquisa Contribuintes"
               Top             =   300
               Width           =   9045
               _ExtentX        =   15954
               _ExtentY        =   2619
               Altura          =   1905
               Caption         =   " Fiscalização"
               CorTexto        =   16777215
               CorFaixa        =   32768
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.cmdVISUAL cmdBuscar 
                  Height          =   330
                  Left            =   8040
                  TabIndex        =   16
                  Top             =   1080
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  Caption         =   "&Buscar"
                  Acao            =   5
                  CorBorda        =   32768
                  CorFrente       =   16384
               End
               Begin VTOcx.cmdVISUAL cmdBUsca 
                  Height          =   285
                  Left            =   3150
                  TabIndex        =   15
                  Top             =   720
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   503
                  Caption         =   ""
                  Acao            =   5
                  CorBorda        =   32768
                  CorFrente       =   16384
                  CorFoco         =   14737632
               End
               Begin VTOcx.txtVISUAL txtNome 
                  Height          =   285
                  Left            =   3510
                  TabIndex        =   14
                  Top             =   735
                  Width           =   5490
                  _ExtentX        =   9684
                  _ExtentY        =   503
                  Caption         =   ""
                  Text            =   ""
                  Enabled         =   0   'False
                  CorRotulo       =   0
                  CorTexto        =   4194304
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtInscricao 
                  Height          =   285
                  Left            =   300
                  TabIndex        =   13
                  Top             =   720
                  Width           =   2805
                  _ExtentX        =   4948
                  _ExtentY        =   503
                  Caption         =   "Contribuinte"
                  Text            =   ""
                  CorRotulo       =   0
                  AgruparValores  =   0   'False
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtDtfinal 
                  Height          =   285
                  Left            =   3300
                  TabIndex        =   12
                  Top             =   1080
                  Width           =   2685
                  _ExtentX        =   4736
                  _ExtentY        =   503
                  Caption         =   "Data Final"
                  Text            =   ""
                  Formato         =   0
                  Restricao       =   2
                  CorRotulo       =   0
                  AgruparValores  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtDtInicial 
                  Height          =   285
                  Left            =   390
                  TabIndex        =   11
                  Top             =   1080
                  Width           =   2685
                  _ExtentX        =   4736
                  _ExtentY        =   503
                  Caption         =   "Data Inicial"
                  Text            =   ""
                  Formato         =   0
                  Restricao       =   2
                  CorRotulo       =   0
                  AgruparValores  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtFiscalizacao 
                  Height          =   285
                  Left            =   75
                  TabIndex        =   10
                  Top             =   375
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   503
                  Caption         =   "Nº Fiscalização"
                  Text            =   ""
                  Restricao       =   2
                  CorRotulo       =   0
                  AgruparValores  =   0   'False
               End
            End
         End
      End
   End
   Begin VB.Menu mnuGeral 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuLinha01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInst 
         Caption         =   "1. Fase de Instauração"
         Begin VB.Menu mnuOrdem 
            Caption         =   "Imprimir Ordem de Serviço"
         End
         Begin VB.Menu mnuTIAF 
            Caption         =   "Imprimir TIAF"
         End
      End
      Begin VB.Menu mnuLinha 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApura 
         Caption         =   "2. Fase de Apuração"
         Begin VB.Menu mnuLevanta 
            Caption         =   "Levantamento de Dados"
         End
      End
      Begin VB.Menu mnuConclusao 
         Caption         =   "3. Fase de Conclusão"
         Begin VB.Menu mnuEncerrar 
            Caption         =   "Encerramento da Fiscalização"
         End
         Begin VB.Menu mnuTEAF 
            Caption         =   "Imprimir TEAF"
         End
      End
      Begin VB.Menu mnuLinha2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancela 
         Caption         =   "Cancelamento da Fiscalização"
      End
   End
End
Attribute VB_Name = "TFIS401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rpt As New VSRelatorio
Dim Fisc As New Fiscalizacao
Private Sub cmdBUsca_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscricao
End Sub

Private Sub cmdBuscar_Click()
    Fisc.PreencheGridFiscalizacao grdFisc, txtInscricao, txtFiscalizacao, txtDtInicial, txtDtfinal
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    TabDados.Tabs(1).Selected = True
    grdFisc.Preencher Bdados, ""
    txtFiscalizacao.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Fisc.Funcionario.PreencheComboFuncionario cboAutoridade
End Sub

Private Sub grdAndamento_DblClick()
    If grdAndamento.ListItems.Count = 0 Then Exit Sub
    TFIS202.Tag = grdAndamento.SelectedItem.SubItems(1)
    TFIS202.Caption = grdAndamento.SelectedItem.SubItems(2) & " | " & grdAndamento.SelectedItem.SubItems(3) & " | " & grdAndamento.SelectedItem
    TFIS202.Show
End Sub

Private Sub grdFisc_Click()
    If grdFisc.ListItems.Count > 0 Then Fisc.Rede.PreencheEtapasPossiveisRede grdProcedimentos, grdFisc.SelectedItem 'Nvl(grdFisc.SelectedItem.SubItems(11), 0)
    If grdFisc.ListItems.Count > 0 Then Fisc.Andamento.PreencheAndamentoFiscalizacao grdAndamento, grdFisc.SelectedItem
End Sub
Private Sub grdFisc_DblClick()
    Dim Sql As String
    Dim Rs As VSRecordset
    If grdFisc.ListItems.Count = 0 Then Exit Sub
    If Fisc.CarregaDadosFiscalizacao(grdFisc.SelectedItem) Then
        txtIm = BuscaContribuinte(Fisc.vIm, txtNomeContrib, txtEndereco)
        
        txtInscricao_LostFocus
        txtInicio = Fisc.vDataInicio
        txtFim = Fisc.vDataFim
        txtDevolucao = Fisc.vDataDevolucao
        txtPeriodoFiscalizado = Fisc.vPeriodoFiscalizado
        cboAutoridade.SetarLinha Fisc.vCodFuncionario, 0
        If grdAndamento.ListItems.Count > 0 Then TabDados.Tabs(2).Selected = True
    End If
End Sub

Private Sub grdFisc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuGeral
    End If
End Sub

Private Sub grdProcedimentos_DblClick()
    If grdProcedimentos.ListItems.Count > 0 Then
        TFIS202.Tag = grdFisc.SelectedItem
        TFIS202.Caption = grdProcedimentos.SelectedItem & " | " & grdProcedimentos.SelectedItem.SubItems(1)
        
        TFIS202.Show
        TabDados.Tabs(1).Selected = True
    End If
End Sub

Private Sub mnuCancela_Click()
    Dim Motivo As String
    
    If Trim(grdFisc.SelectedItem.SubItems(6)) <> "" Then
        Avisa "Fiscalização já encerrada em " & Trim(grdFisc.SelectedItem.SubItems(6)) & " não poderá ser cancelada. "
        Exit Sub
    End If
    If Confirma("Confirma o cancelamento da Fiscalização nº " & grdFisc.SelectedItem & "?") Then
        Motivo = Trim(Util.Entrada("Informe o Motivo do cancelamento", "Informação obrigatória"))
        If Trim(Motivo) = "" Then
            Erro "Cancelamento de Fiscalização não foi concluído por falta de motivo."
            Exit Sub
        End If
        If Fisc.CancelaFiscalizacao(grdFisc.SelectedItem) Then
            Avisa "Fiscaliazação cancelada com sucesso."
            cmdBuscar_Click
            Edita.LimpaCampos Me
            txtFiscalizacao.SetFocus
        End If
    End If
End Sub

Private Sub mnuEncerrar_Click()
    Dim Data As String
    If Trim(grdFisc.SelectedItem.SubItems(6)) <> "" Then
        Avisa "Fiscalização já encerrada em " & Trim(grdFisc.SelectedItem.SubItems(6)) & "."
        Exit Sub
    End If
    If Confirma("Confirma o encerramento da Fiscalização nº " & grdFisc.SelectedItem & "?") Then
        Data = Trim(Util.Entrada("Informe a Data do encerramento (dd/mm/aaaa)", "Informação obrigatória(dd/mm/aaaa)"))
        If Not IsDate(Trim(Data)) Then
            Avisa "Data inválida."
            Exit Sub
        End If
        If Trim(Data) = "" Then
            Erro "encerramento de Fiscalização não foi concluído por falta da data."
            Exit Sub
        End If
        If DateDiff("d", grdFisc.SelectedItem.SubItems(4), Data) < 0 Then
            Erro "Data de encerramento da ação fiscal não pode ser menor do que a data de início da ação fiscal."
            Exit Sub
        End If
        If Fisc.EncerraFiscalizacao(grdFisc.SelectedItem, Data) Then
            Avisa "Fiscaliazação encerrada com sucesso."
            If Confirma("Deseja imprimir o TEAF agora?") Then
                mnuTEAF_Click
            End If
            cmdBuscar_Click
            Edita.LimpaCampos Me
            txtFiscalizacao.SetFocus
        End If
    End If
End Sub

Private Sub mnuLevanta_Click()
    TFIS202.Tag = grdFisc.SelectedItem
    TFIS202.Caption = mnuLevanta.Caption
    TFIS202.Show 1
End Sub

Private Sub mnuTEAF_Click()
    Avisa "Em desenvolvimento."
End Sub

Private Sub mnuTIAF_Click()
    Dim Selecao As String
    Dim Rpt As New VSRelatorio
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path & "\TIAF.rpt") Then Exit Sub
        .Selecao = "{TAB_FISCALIZACAO.TFI_COD_FISCALIZACAO} = " & grdFisc.SelectedItem
        .Formulas "DOCUMENTO", UCase(Temp.PegaParametro(Bdados, "NOME TIAF"))
        .Formulas "VT_NUM_DOC", "FISCALIZAÇÃO Nº " & Left(grdFisc.SelectedItem, 2) & "." & Mid(grdFisc.SelectedItem, 3, 3) & "." & Mid(grdFisc.SelectedItem, 6, 3) & " / TIAF nº " & Left(grdFisc.SelectedItem.SubItems(1), 2) & "." & Mid(grdFisc.SelectedItem.SubItems(1), 3, 3) & "." & Mid(grdFisc.SelectedItem.SubItems(1), 6, 3)
        .Formulas "VT_PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
        .Formulas "VT_SECRETARIA", UCase(Temp.PegaParametro(Bdados, "SECRETARIA"))
        .Arvore = False
        .Visualizar
    End With
End Sub

Private Sub txtInscricao_LostFocus()
    Dim Ic As String
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtInscricao) = 10 Or Len(txtInscricao) = 11 Then
            Ic = Imposto.FormataInscricao(txtInscricao, InscContrib)
        Else
            Ic = txtInscricao
        End If
    Else
            Ic = txtInscricao
    End If
    txtInscricao = BuscaContribuinte(Ic, txtNome)
End Sub
