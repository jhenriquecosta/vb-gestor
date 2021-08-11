VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONT~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TARR1011 
   Caption         =   "TARR101"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   13150
      _Version        =   131082
      TabCount        =   2
      Tabs            =   "TARR1011.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   7065
         Left            =   -99969
         TabIndex        =   17
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   12462
         _Version        =   131082
         TabGuid         =   "TARR1011.frx":007C
         Begin Cabecalho.cabVISUAL cabVISUAL2 
            Height          =   645
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   1138
            Formulario      =   "RETORNO DE DOCUMENTOS BANCARIO"
            Descricao       =   "Selecione o diretorio para Recepção dos Pagamentos"
            Icone           =   "TARR1011.frx":00A4
         End
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   2940
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   5186
            Altura          =   1905
            Caption         =   " Consultar Por:"
            CorTexto        =   0
            CorFaixa        =   12632256
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VB.Frame Frame2 
               Caption         =   "Dados da Recpção"
               Height          =   1725
               Left            =   120
               TabIndex        =   25
               Top             =   1065
               Width           =   10395
               Begin VB.Frame Frame3 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   1155
                  Left            =   5580
                  TabIndex        =   26
                  Top             =   165
                  Width           =   3270
                  Begin VB.Label Label8 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "TOTAL DESCONTO:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   285
                     TabIndex        =   34
                     Top             =   330
                     Width           =   1500
                  End
                  Begin VB.Label LblTotalDesconto 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "0,00"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   195
                     Left            =   2700
                     TabIndex        =   33
                     Top             =   330
                     Width           =   360
                  End
                  Begin VB.Label LblTotalGeral 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "0,00"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   195
                     Left            =   2700
                     TabIndex        =   32
                     Top             =   840
                     Width           =   360
                  End
                  Begin VB.Label LblTotalJuros 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "0,00"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   195
                     Left            =   2700
                     TabIndex        =   31
                     Top             =   600
                     Width           =   360
                  End
                  Begin VB.Label LblTotalTitulo 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "0,00"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   195
                     Left            =   2700
                     TabIndex        =   30
                     Top             =   75
                     Width           =   360
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "TOTAL GERAL:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   615
                     TabIndex        =   29
                     Top             =   840
                     Width           =   1170
                  End
                  Begin VB.Label Label4 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "TOTAL JUROS:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   585
                     TabIndex        =   28
                     Top             =   585
                     Width           =   1185
                  End
                  Begin VB.Label Label3 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "TOTAL TITULO:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   540
                     TabIndex        =   27
                     Top             =   60
                     Width           =   1245
                  End
               End
               Begin MSComDlg.CommonDialog CommonDialog1 
                  Left            =   75
                  Top             =   270
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin MSComctlLib.ProgressBar Progresso 
                  Height          =   225
                  Left            =   90
                  TabIndex        =   35
                  Top             =   1410
                  Visible         =   0   'False
                  Width           =   10110
                  _ExtentX        =   17833
                  _ExtentY        =   397
                  _Version        =   393216
                  Appearance      =   0
                  Scrolling       =   1
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Lote:"
                  Height          =   195
                  Index           =   1
                  Left            =   1995
                  TabIndex        =   45
                  Top             =   195
                  Width           =   375
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Pagamentos:"
                  Height          =   195
                  Left            =   1020
                  TabIndex        =   44
                  Top             =   1155
                  Width           =   1350
               End
               Begin VB.Label LblDataRecepcao 
                  AutoSize        =   -1  'True
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   43
                  Top             =   900
                  Width           =   45
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Data Recepção:"
                  Height          =   195
                  Index           =   0
                  Left            =   1215
                  TabIndex        =   42
                  Top             =   900
                  Width           =   1155
               End
               Begin VB.Label LblAgencia 
                  AutoSize        =   -1  'True
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   41
                  Top             =   420
                  Width           =   45
               End
               Begin VB.Label Agencia 
                  AutoSize        =   -1  'True
                  Caption         =   "Agência:"
                  Height          =   195
                  Left            =   1740
                  TabIndex        =   40
                  Top             =   420
                  Width           =   630
               End
               Begin VB.Label LblConta 
                  AutoSize        =   -1  'True
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   39
                  Top             =   660
                  Width           =   45
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Conta:"
                  Height          =   195
                  Left            =   1860
                  TabIndex        =   38
                  Top             =   645
                  Width           =   495
               End
               Begin VB.Label LblTotalRegistro 
                  AutoSize        =   -1  'True
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   37
                  Top             =   1140
                  Width           =   45
               End
               Begin VB.Label lblLote 
                  AutoSize        =   -1  'True
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   36
                  Top             =   195
                  Width           =   60
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Informe o caminho dos arquivos de remessa.ret"
               Height          =   720
               Left            =   120
               TabIndex        =   21
               Top             =   360
               Width           =   10380
               Begin VB.TextBox txtCamminhoRemessa 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   90
                  TabIndex        =   22
                  Top             =   240
                  Width           =   7695
               End
               Begin VTOcx.cmdVISUAL cmdConsultaArquivo 
                  Height          =   405
                  Left            =   7860
                  TabIndex        =   23
                  Top             =   195
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   714
                  Caption         =   "Arquivo"
                  Acao            =   5
                  CorBorda        =   -2147483645
                  CorFoco         =   -2147483628
               End
               Begin VTOcx.cmdVISUAL cmdReceber 
                  Height          =   405
                  Left            =   9210
                  TabIndex        =   24
                  Top             =   195
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   714
                  Caption         =   "&Receber"
                  Acao            =   3
                  CorBorda        =   -2147483645
                  CorFoco         =   -2147483628
               End
            End
         End
         Begin VTOcx.grdVISUAL grdVISUAL1 
            Height          =   3060
            Left            =   120
            TabIndex        =   20
            Top             =   3840
            Width           =   10485
            _ExtentX        =   18494
            _ExtentY        =   5398
            Caption         =   "Dados"
            CheckBox        =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   7065
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   12462
         _Version        =   131082
         TabGuid         =   "TARR1011.frx":03BE
         Begin VB.CheckBox Check1 
            Caption         =   "Marcar Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   13
            Top             =   6600
            Width           =   1665
         End
         Begin Cabecalho.cabVISUAL cabVISUAL1 
            Height          =   645
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   1138
            Formulario      =   "REMESSA DE DOCUMENTOS BANCARIO"
            Descricao       =   "Selecione os pagamentos que deseja enviar ao Banco conveniado"
            Icone           =   "TARR1011.frx":03E6
         End
         Begin VTOcx.fraVISUAL txt 
            Height          =   1500
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   2646
            Altura          =   1905
            Caption         =   " Consultar Por:"
            CorTexto        =   0
            CorFaixa        =   12632256
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtObservacao 
               Height          =   285
               Left            =   300
               TabIndex        =   11
               Top             =   645
               Width           =   10065
               _ExtentX        =   17754
               _ExtentY        =   503
               Caption         =   "Observação"
               Text            =   ""
               TipoLetras      =   0
            End
            Begin VTOcx.txtVISUAL txtValor 
               Height          =   285
               Left            =   9360
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   315
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
            End
            Begin VTOcx.cmdVISUAL cmdAddDam 
               Height          =   390
               Left            =   9360
               TabIndex        =   9
               Top             =   945
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   688
               Caption         =   "Incluir"
               Acao            =   1
            End
            Begin VTOcx.txtVISUAL txtNome 
               Height          =   285
               Left            =   2595
               TabIndex        =   8
               Top             =   315
               Width           =   6660
               _ExtentX        =   11748
               _ExtentY        =   503
               Caption         =   "Dados"
               Text            =   ""
               Enabled         =   0   'False
               TipoLetras      =   0
            End
            Begin VTOcx.txtVISUAL txtDocumento 
               Height          =   285
               Left            =   300
               TabIndex        =   7
               Top             =   315
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   503
               Caption         =   "Doc DAM"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtDiretorio 
               Height          =   285
               Left            =   2595
               TabIndex        =   6
               Top             =   960
               Width           =   6660
               _ExtentX        =   11748
               _ExtentY        =   503
               Caption         =   "Diretorio Remessa"
               Text            =   ""
               TipoLetras      =   0
            End
            Begin VTOcx.cboVISUAL cboStatus 
               Height          =   315
               Left            =   6000
               TabIndex        =   5
               Top             =   700
               Visible         =   0   'False
               Width           =   4000
               _ExtentX        =   7064
               _ExtentY        =   556
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorRotulo       =   0
               Enabled         =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNumero 
               Height          =   285
               Left            =   300
               TabIndex        =   4
               Top             =   960
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   503
               Caption         =   "Número Remessa"
               Text            =   ""
            End
         End
         Begin VTOcx.grdVISUAL Grid 
            Height          =   4020
            Left            =   120
            TabIndex        =   12
            Top             =   2400
            Width           =   10485
            _ExtentX        =   18494
            _ExtentY        =   7091
            Caption         =   "Dados"
            CheckBox        =   -1  'True
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL1 
            Height          =   390
            Left            =   1800
            TabIndex        =   14
            Top             =   6600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   688
            Caption         =   "Gerar Arquivo Bradesco"
            Acao            =   3
         End
         Begin VTOcx.cmdVISUAL cmdLimpar 
            Height          =   390
            Left            =   4320
            TabIndex        =   15
            Top             =   6600
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   688
            Caption         =   "Limpar"
            Acao            =   6
         End
         Begin VTOcx.cmdVISUAL cmdSair 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   9840
            TabIndex        =   16
            Top             =   6600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            Caption         =   "Sai&r"
            Acao            =   7
         End
      End
   End
End
Attribute VB_Name = "TARR1011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim remHeaderBradesco           As New HeaderBradesco
Dim remBradesco                 As New RemessaBradesco
Dim trlBradesco                 As New TraillerBradesco
Dim rs                          As VSRecordset
Dim Sql                         As String
Dim NomeArquivo                 As String
Dim i                           As Integer
Dim Doc                         As Integer
Dim marcou                      As Boolean
Dim arquivo                     As String
Dim sequencialDetalheRemessa    As Integer
Dim obrigacao                   As String
Dim codigoStatus                As Integer
Dim obgris As String, sqlO As String

Dim bcpConta As String, bcpAgencia As String, bcpCarteira As String, bcpDigConta As String
Private Sub Check1_Click()
    Grid.MarcarTodos Check1.Value
End Sub
Private Sub cmdAddDam_Click()
    If codigoStatus <> 2 Then
        Avisa "O Status deste documento não permite inclusão na REMESSA"
        txtNome = ""
        txtDocumento = 0
        txtValor = 0
        Exit Sub
    End If
    Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_OBSERVACAO='" & txtObservacao & "' WHERE TOC_COD_OBRIGACAO=" & txtDocumento)
    obgris = obgris & txtDocumento & ","
    sqlO = "SELECT NUM_DOCUMENTO,COD_CLIENTE,NOME,VALOR_ATUAL FROM VIS_CONTA_CONTRIBUINTE WHERE NUM_DOCUMENTO IN(" & obgris
    sqlO = Left(sqlO, Len(sqlO) - 1)
    sqlO = sqlO & ")"
    If Grid.Preencher(Bdados, sqlO) Then
    Else
    End If
    
End Sub

Private Sub cmdConsultaArquivo_Click()
    With CommonDialog1
        .DialogTitle = "Selecione o arquivo retorno do bradesco"
        .Filter = "Arquivos do tipo retorno | *.ret"
        .ShowOpen
        If .FileName <> "" Then
            txtCamminhoRemessa = .FileName
        End If
    End With
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    Grid.ListItems.Clear
    'cboStatus.SetarLinha esrAberto, 1
End Sub

Private Sub montarHeder()
        arquivo = "CB" & Format(Now, "DDMM") & Format(txtNumero, "00") & ".REM"
        With remHeaderBradesco
            .IdentificacaoRegistro = "0"
            .IdentificaoArquivoRemessa = "1"
            .LiteralRemessa = "REMESSA"
            .CodigoServico = "01"
            .LiteralServico = preencherComCaractere("COBRANCA", 15, " ")
            .CodigoEmpresa = Format(4316876, "00000000000000000000")
            .NomeEmpresa = preencherComCaractere("PREFEITURA MUNICIPAL CODO", 30, " ")
            .NumeroCamaraCompensacao = "237"
            .NomeBanco = preencherComCaractere("BRADESCO", 15, " ")
            .DataGravacaoArquivo = Format(Now, "DDMMYY")
            .IdentificacaoSistema = "MX"
            .NumeroSequencialRemessa = Format(1, "0000000")
            .NumeroSequencialRegistroUmAUm = Format(1, "000000")
            arquivo = .gerarHeaderRemessa(txtDiretorio, arquivo) 'DIRETORIO + NOME ARQUIVO
        End With
End Sub
Private Sub montarDetalhe()
             sequencialDetalheRemessa = 2
            For Doc = 1 To Grid.ListItems.Count ' TOTAL DE TRIBUTOS
                If Grid.ListItems(Doc).Checked = True Then ' SE FOI MARCADO PARA REMESSA
                    obrigacao = Grid.ListItems(Doc)
                    If Bdados.AbreTabela("SELECT * FROM VIS_CONTA_CONTRIBUINTE WHERE NUM_DOCUMENTO=" & obrigacao) Then
                       
                        With remBradesco
                            .IdentificacaoRegistro = "1"
                            .AgenciaDebito = Format(0, "00000")
                            .DigitoAgenciaDebito = "0"
                            .RazaoContaCorrente = Format(0, "00000")
                            .ContaCorrente = Format(0, "0000000")
                            .DigitoContaCorrente = "0"
                            .IdentificacaoEmpresaCedenteNoBanco = "0" & bcpCarteira & bcpAgencia & bcpConta & bcpDigConta
                            .NumeroControleParticipante = "0000000000000000000000000"
                            .CodigoBancoCamaraCompensacao = "000"
                            .CampoMulta = 0
                            .PercentualMulta = Format(0, "0000")
                            .IdentificacaoTituloBanco = Format(Bdados.Tabela("NUM_DOCUMENTO"), "00000000000") ''
                            .DigitoAutoConferenciaNossoNumero = .gerarDigitoConferencia82(bcpCarteira)
                            .DescontoBonificacaoDia = Format(0, "0000000000")
                            .CondicaoParaEmissaoPapeladaCobranca = "2"
                            .EmitePapeletaDebitoAutomatico = "N"
                            .IndicadorRateioCredito = " "
                            .EnderecamentoAvisoDebAutoCC = "2"
                            .IndicacaoOcorrencia = "01"
                            .NumeroDocumento = Format(Bdados.Tabela("NUM_DOCUMENTO"), "0000000000")
                            .DataVencimentoTitulo = Format(Now, "DDMMYY")
                            Dim ValorTitulo As String
                            ValorTitulo = Format(Bdados.Tabela("VALOR_ATUAL"), "#,##0.00")
                            .ValorTitulo = Format(retiraSeparadores(ValorTitulo), "0000000000000")
                            .BancoEncarregadoCobranca = "000" '
                            .AgenciaDepositaria = "00000" '
                            .EspecieTitulo = "01"
                            .Identificacao = "N"
                            .DataEmissaoTitulo = Format(Bdados.Tabela("DATA_GERACAO"), "DDMMYY")
                            .PrimeiraInstrucao = "00"
                            .SegundaInstrucao = "00"
                            .ValorCobradoDiaAtraso = Format(0, "0000000000000")
                            .DataLimiteConcessaoDesconto = Format(0, "000000")
                            .ValorDesconto = Format(0, "0000000000000")
                            .ValorIOF = Format(0, "0000000000000")
                            .ValorAbatimento = Format(0, "0000000000000")
                            Dim cpfCnpj As String
                            cpfCnpj = retiraSeparadores(Bdados.Tabela("CPF_CNPJ"))
                            If Len(cpfCnpj) < 14 Then
                                .IdentificacaoTipoInscricaoSacado = "01" 'CPF
                            Else
                                .IdentificacaoTipoInscricaoSacado = "02" 'CNPJ
                            End If
                            .NumeroInscricaoSacado = Format(cpfCnpj, "00000000000000")
                            .NomeSacado = preencherComCaractere(Bdados.Tabela("NOME"), 40, " ")
                            Dim Endereco As String
                            Endereco = Bdados.Tabela("LOGRADOURO") & " " & Bdados.Tabela("NOME_LOGRADOURO") & " " & Bdados.Tabela("NUMERO_ENDERECO") & " " & Bdados.Tabela("BAIRRO") & " " & Bdados.Tabela("CIDADE") & " " & Bdados.Tabela("ESTADO")
                            .EnderecoCompleto = preencherComCaractere(Endereco, 40, " ")
                            .PrimeiraMensagem = preencherComCaractere(" ", 12, " ")
                            Dim ccep As String
                            ccep = Bdados.Tabela("CEP")
                            ccep = preencherComCaractere(ccep, 8, "0")
                            .cep = Left(ccep, 5)
                            .SufixoCEP = Right(ccep, 3)
                            .SacadorAvalistaSegundaMensagem = preencherComCaractere(" ", 60, " ")
                            .NumeroSequencialRegistro = Format(sequencialDetalheRemessa, "000000") ' SEQUENCIAL COMECANDO COM 2
                            sequencialDetalheRemessa = sequencialDetalheRemessa + 1
                            arquivo = .gerarDetalheRemessa(txtDiretorio, arquivo)
                            
                        End With
                    End If
                End If
            Next Doc
End Sub
Private Sub montarTrailer()
    With trlBradesco
        .IdentificacaoRegistro = "9"
        .NumeroSequencialRegistro = Format(sequencialDetalheRemessa, "000000") ' ULTIMO SEQUENCIAL DETALHE + 1
        arquivo = .gerarTrailerRemessa(txtDiretorio, arquivo)
    End With
End Sub
Private Sub transmitirRemessa()
    '746,1103 CONTA, OBRIGACAO
    For Doc = 1 To Grid.ListItems.Count ' TOTAL DE TRIBUTOS
        If Grid.ListItems(Doc).Checked = True Then ' SE FOI MARCADO PARA REMESSA
            obrigacao = Grid.ListItems(Doc)
            Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_STATUS_OBRIGACAO=12 WHERE TOC_COD_OBRIGACAO=" & obrigacao)
            Bdados.Executa ("UPDATE TAB_CONTA_CONTRIBUINTE SET tcc_status_conta=4 WHERE tcc_codigo_conta=" & obrigacao)
            Bdados.Executa ("INSERT INTO TAB_BCP_REMESSA (COD_REMESSA, COD_OBRIGACAO) VALUES('" & retiraSeparadores(arquivo) & "'," & obrigacao & ")")
        End If
    Next Doc
End Sub
Private Sub cmdVISUAL1_Click()
    marcou = False
    For i = 1 To Grid.ListItems.Count
        If Grid.ListItems(i).Checked Then
            marcou = True
            Exit For
        End If
    Next
    If marcou = False Then
        Avisa "Selecione um recebimento para geração do arquivo remessa"
        Exit Sub
    End If
    If Grid.ListItems.Count >= 1 Then
    If Me.Tag <> "" Then
        GoTo Vai
    End If
        If Confirma("Confirma geração do arquivos remessa?") Then
Vai:
            montarHeder
            montarDetalhe
            montarTrailer
            transmitirRemessa
            
            Grid.ListItems.Clear
            limparCampos
            txtDocumento.SetFocus
        End If
        
    End If
    Avisa "Arquivos gerados com sucesso."
    If Me.Tag = "" Then
        cmdBuscar_Especial_Click
    Else
        Unload Me
    End If

End Sub

Public Function preencherComCaractere(Texto As String, tamanhoDefinido As Integer, caractere As String) As String
    Dim tamanhoAtual As Integer, diferenca As Integer
    Dim novoTexto As String, caracs As String
    tamanhoAtual = Len(Texto)
    If tamanhoAtual = tamanhoDefinido Then
        novoTexto = Texto
    ElseIf tamanhoAtual < tamanhoDefinido Then
        diferenca = tamanhoDefinido - tamanhoAtual
        For i = 1 To diferenca
            caracs = caracs & caractere
        Next i
        novoTexto = Texto & caracs
    ElseIf tamanhoAtual > tamanhoDefinido Then
        diferenca = tamanhoAtual - tamanhoDefinido
        novoTexto = Left(Texto, tamanhoAtual - diferenca)
    End If
    Dim ass As Integer
    ass = Len(novoTexto)
    preencherComCaractere = novoTexto
End Function

Private Sub Form_Load()
     txtDiretorio = "C:\ORION TECNOLOGIAS\BRADESCO\Remessa\"
     txtNumero = 1
    bcpConta = Format(Temp.PegaParametro(Bdados, "CONVENIO CONTA"), "0000000")
     bcpAgencia = Format(Temp.PegaParametro(Bdados, "CONVENIO AGENCIA"), "00000")
    bcpCarteira = Format(Temp.PegaParametro(Bdados, "CONVENIO CARTEIRA"), "000")
    bcpDigConta = Right(Temp.PegaParametro(Bdados, "CONTA"), 1)
     'txtInicio = Format(Now, "DD/MM/YYYY")
     'txtFim = Format(Now, "DD/MM/YYYY")
End Sub
Private Function retiraSeparadores(Valor As String) As String
    Valor = Replace(Valor, ",", "")
    Valor = Replace(Valor, ".", "")
    Valor = Replace(Valor, "-", "")
    Valor = Replace(Valor, "/", "")
    retiraSeparadores = Valor
End Function


Private Sub txtDocumento_LostFocus()
    Dim Sql As String
    If Len(txtDocumento) > 0 Then
        If Bdados.AbreTabela("SELECT * FROM VIS_CONTA_CONTRIBUINTE WHERE NUM_DOCUMENTO=" & txtDocumento) Then
            codigoStatus = Bdados.Tabela("STATUS")
            txtNome = Bdados.Tabela("NOME")
            txtValor = Bdados.Tabela("VALOR_ATUAL")
        Else
            Avisa "DAM  não encontrado"
            limparCampos
        End If
    Else
        limparCampos
    End If
End Sub
Private Sub limparCampos()
            txtNome = ""
        txtDocumento = 0
        txtValor = 0
        txtObservacao = ""

End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
