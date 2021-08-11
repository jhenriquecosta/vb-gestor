VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TFIS101 
   Caption         =   "TPRT101"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   1138
      Icone           =   "TFIS101.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   13
      Top             =   6990
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   7035
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   9
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
      Height          =   6600
      Left            =   -30
      TabIndex        =   14
      Tag             =   "Documento gerencial"
      Top             =   330
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   11642
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      Tabs            =   "TFIS101.frx":031A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   6210
         Left            =   -99969
         TabIndex        =   15
         Top             =   30
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   10954
         _Version        =   131082
         TabGuid         =   "TFIS101.frx":03D2
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   6225
            Left            =   0
            TabIndex        =   16
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
            Begin VB.TextBox txtFundamentacao 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   5325
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Top             =   870
               Width           =   9030
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6210
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   10954
         _Version        =   131082
         TabGuid         =   "TFIS101.frx":03FA
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   6555
            Left            =   0
            TabIndex        =   18
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
            Begin VTOcx.fraVISUAL fraVISUAL3 
               Height          =   1455
               Left            =   90
               TabIndex        =   24
               ToolTipText     =   "Pesquisa Contribuintes"
               Top             =   4635
               Width           =   8940
               _ExtentX        =   15769
               _ExtentY        =   2566
               Altura          =   1905
               Caption         =   " Levantamento Homologatório"
               CorTexto        =   16777215
               CorFaixa        =   32768
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.txtVISUAL txtInicio 
                  Height          =   495
                  Left            =   60
                  TabIndex        =   3
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
                  TabIndex        =   4
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
               Begin VTOcx.txtVISUAL txtDevolucao 
                  Height          =   495
                  Left            =   3870
                  TabIndex        =   5
                  Top             =   330
                  Width           =   1965
                  _ExtentX        =   3466
                  _ExtentY        =   873
                  Caption         =   "Data Devolução"
                  Text            =   ""
                  Formato         =   0
                  Restricao       =   2
                  AlinhamentoRotulo=   1
                  CorRotulo       =   0
                  AgruparValores  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtPeriodoFiscalizado 
                  Height          =   495
                  Left            =   5970
                  TabIndex        =   6
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
                  TabIndex        =   7
                  Top             =   840
                  Width           =   8775
                  _ExtentX        =   15478
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
               TabIndex        =   20
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
                  TabIndex        =   23
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
                  TabIndex        =   22
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
                  TabIndex        =   21
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
            Begin VTOcx.fraVISUAL fraVISUAL1 
               Height          =   3060
               Left            =   90
               TabIndex        =   19
               ToolTipText     =   "Pesquisa Contribuintes"
               Top             =   1515
               Width           =   8940
               _ExtentX        =   15769
               _ExtentY        =   5398
               Altura          =   1905
               Caption         =   " Formalização do Procedimento"
               CorTexto        =   16777215
               CorFaixa        =   32768
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.txtVISUAL txtInformacoes 
                  Height          =   1275
                  Left            =   60
                  TabIndex        =   2
                  Top             =   1710
                  Width           =   8745
                  _ExtentX        =   15425
                  _ExtentY        =   2249
                  Caption         =   "Informações Requeridas"
                  Text            =   ""
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtDocumentos 
                  Height          =   1275
                  Left            =   60
                  TabIndex        =   1
                  Top             =   345
                  Width           =   8745
                  _ExtentX        =   15425
                  _ExtentY        =   2249
                  Caption         =   "Documentos Solicitados"
                  Text            =   ""
                  AlinhamentoRotulo=   1
               End
            End
         End
      End
   End
End
Attribute VB_Name = "TFIS101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rpt As New VSRelatorio
Dim Fisc As New Fiscalizacao
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
    Dim Campos As String
    Dim Valores As String
    Dim CodFiscalizacao As String
    Dim Conta As New ContaCorrente
    Dim CodTIAF As String
    CodFiscalizacao = Conta.GeraCodPagamento("77")
    CodTIAF = Conta.GeraCodPagamento("13")
    Campos = "TFI_COD_FISCALIZACAO,TFI_TCI_IM,TFI_DOCUMENTOS_SOLICITADOS,TFI_INFORMACOES_SOLICITADAS," & _
        "TFI_DATA_FISCALIZACAO,TFI_DATA_INICIO,TFI_DATA_FIM,TFI_DATA_DEVOLUCAO,TFI_PERIODO_FISCALIZADO,TFI_TFU_COD_FUNCIONARIO," & _
        "TFI_FUNDAMENTACAO,TFI_DATA_TIAF,TFI_COD_TIAF,TFI_STATUS,TIF_USUARIO_TIAF"
    Valores = Bdados.PreparaValor(CodFiscalizacao, Bdados.Converte(txtIm, tctexto), _
            txtDocumentos, txtInformacoes, Bdados.Converte(Date, TCDataHora), _
            Bdados.Converte(txtInicio, TCDataHora), Bdados.Converte(txtFim, TCDataHora), _
            Bdados.Converte(txtDevolucao, TCDataHora), _
            txtPeriodoFiscalizado, cboAutoridade.Coluna(0).Valor, _
            txtFundamentacao, Bdados.Converte(Date, TCDataHora), CodTIAF, 1, AplicacoesVTFuncoes.Usuario)
    Bdados.GravaDados "TAB_PARAMETRO_TEXTO", Bdados.PreparaValor(txtFundamentacao), "TPT_TEXTO", "TPT_PARAMETRO = 'TIAF'"
    If Bdados.InsereDados("TAB_FISCALIZACAO", Valores, Campos) Then
        If Confirma("Deseja emitir o TIAF nº " & CodFiscalizacao & " agora?") Then
            Dim Selecao As String
            Dim Rpt As New VSRelatorio
                        
            If Trim(txtIm) <> "" Then
                Selecao = " {TAB_FISCALIZACAO.TFI_COD_TIAF} = " & CodTIAF
            End If
            
            With Rpt
                If Not .DefinirArquivo(Bdados, App.Path & "\TIAF.rpt") Then Exit Sub
                .Selecao = Selecao
                .Formulas "DOCUMENTO", UCase(Temp.PegaParametro(Bdados, "NOME TIAF"))
                .Formulas "VT_NUM_DOC", "FISCALIZAÇÃO Nº " & Left(CodFiscalizacao, 2) & "." & Mid(CodFiscalizacao, 3, 3) & "." & Mid(CodFiscalizacao, 6, 3) & " / TIAF nº " & Left(CodTIAF, 2) & "." & Mid(CodTIAF, 3, 3) & "." & Mid(CodTIAF, 6, 3)
                .Formulas "VT_PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
                .Formulas "VT_SECRETARIA", UCase(Temp.PegaParametro(Bdados, "SECRETARIA"))
        '        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                .Arvore = False
                .Visualizar
                Set Rpt = Nothing
            End With
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Dim Sql As String
    Dim rs As VSRecordset
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Fisc.Funcionario.PreencheComboFuncionario cboAutoridade
    Sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = 'TIAF'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtFundamentacao = "" & rs!TPT_TEXTO
    End If
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
