VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TPRT201 
   Caption         =   "TPRT201"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2565
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   10860
      Begin VTOcx.cboVISUAL cboRestricao 
         Height          =   315
         Left            =   690
         TabIndex        =   7
         Tag             =   "Tributo"
         Top             =   -360
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   556
         Caption         =   "Restrição"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTIPO 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Tag             =   "C"
         Top             =   1605
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   556
         Caption         =   "Etapa     "
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtExercicioInicial 
         Height          =   300
         Left            =   9675
         TabIndex        =   1
         Tag             =   "C"
         Top             =   540
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtExercicioFinal 
         Height          =   300
         Left            =   9675
         TabIndex        =   3
         Tag             =   "C"
         Top             =   885
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Caption         =   "  "
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Tag             =   "C"
         Top             =   885
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtAtividade 
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Tag             =   "C"
         Top             =   1245
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   529
         Caption         =   "Atividade"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboFiscal 
         Height          =   315
         Left            =   6660
         TabIndex        =   0
         Tag             =   "C"
         Top             =   525
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   556
         Caption         =   "Fiscal/Ano"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtOutrosServicos 
         Height          =   300
         Left            =   4440
         TabIndex        =   5
         Tag             =   "C"
         Top             =   1605
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   529
         Caption         =   "Outros Serviços"
         Text            =   ""
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboFiscal2 
         Height          =   315
         Left            =   6660
         TabIndex        =   2
         Tag             =   "C"
         Top             =   885
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   556
         Caption         =   "Fiscal/Ano"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   9720
         TabIndex        =   72
         Top             =   2040
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "Iniciar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8760
         TabIndex        =   73
         Top             =   2040
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   1080
         TabIndex        =   74
         Top             =   2040
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdIncluirInscricao 
         Height          =   375
         Left            =   2355
         TabIndex        =   75
         Top             =   2040
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Alterar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   2950
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   200
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtNOME 
         Height          =   300
         Left            =   240
         TabIndex        =   77
         Tag             =   "C"
         Top             =   525
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   529
         Caption         =   "Nome     "
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   1080
         TabIndex        =   78
         Top             =   180
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         Caption         =   ""
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VB.Label lblStatus 
         Caption         =   "NOVA ORDEM DE SERVIÇO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4440
         TabIndex        =   79
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label lblServico 
         Caption         =   "SERVIÇO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   7560
         TabIndex        =   71
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Inscrição"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   8
         Top             =   1590
         Width           =   45
      End
   End
   Begin ActiveTabs.SSActiveTabs tabEtapa 
      Height          =   3855
      Left            =   0
      TabIndex        =   11
      Top             =   3240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6800
      _Version        =   131082
      TabCount        =   7
      Tabs            =   "TPRT201.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3465
         Left            =   -99969
         TabIndex        =   80
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT201.frx":01A7
         Begin VB.TextBox txtObservacao 
            Appearance      =   0  'Flat
            Height          =   2055
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   83
            Top             =   720
            Width           =   3975
         End
         Begin VTOcx.grdVISUAL grdHistorico 
            Height          =   3615
            Left            =   4200
            TabIndex        =   81
            Top             =   0
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   6376
            CorBorda        =   16711680
            Caption         =   "Histórico"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
         Begin VTOcx.cboVISUAL cboFiscalHistorico 
            Height          =   510
            Left            =   120
            TabIndex        =   82
            Top             =   120
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   900
            Caption         =   "Fiscal     "
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
            Alinhamento     =   1
         End
         Begin VTOcx.cmdVISUAL cmdIncluirHistorico 
            Height          =   375
            Left            =   120
            TabIndex        =   84
            Top             =   2880
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   661
            Caption         =   "Incluir Histórico"
            Acao            =   1
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   3465
         Left            =   -99969
         TabIndex        =   57
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT201.frx":01CF
         Begin VTOcx.cboVISUAL cboTipoRelatorio 
            Height          =   315
            Left            =   1000
            TabIndex        =   58
            Top             =   200
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   556
            Caption         =   "Relatório"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
         End
         Begin VTOcx.txtVISUAL txtDataInicio 
            Height          =   285
            Left            =   1000
            TabIndex        =   59
            Top             =   600
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   503
            Caption         =   "Inicio     "
            Text            =   ""
            Formato         =   0
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtDataFim 
            Height          =   285
            Left            =   3600
            TabIndex        =   60
            Top             =   600
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   503
            Caption         =   "Fim     "
            Text            =   ""
            Formato         =   0
            AgruparValores  =   0   'False
         End
         Begin VTOcx.cboVISUAL cboFiscalRelatorio 
            Height          =   315
            Left            =   1000
            TabIndex        =   61
            Top             =   1000
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   556
            Caption         =   "Fiscal     "
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdImprimirGerencial 
            Height          =   375
            Left            =   4700
            TabIndex        =   62
            Top             =   2000
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Imprimir"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cboVISUAL cboStatusRelatorio 
            Height          =   315
            Left            =   1000
            TabIndex        =   63
            Tag             =   "C"
            Top             =   1400
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   556
            Caption         =   "Etapa     "
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
         Height          =   3465
         Left            =   -99969
         TabIndex        =   49
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT201.frx":01F7
         Begin VB.TextBox txtTexto 
            Height          =   2655
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   600
            Width           =   10575
         End
         Begin VTOcx.cboVISUAL cboFinalizar 
            Height          =   315
            Left            =   2760
            TabIndex        =   68
            Tag             =   "C"
            Top             =   120
            Width           =   6330
            _ExtentX        =   11165
            _ExtentY        =   556
            Caption         =   "Etapa Finalizada"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdFinalizar 
            Height          =   375
            Left            =   9240
            TabIndex        =   69
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Encerrar"
            Acao            =   6
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtDataRelatorioFiscal 
            Height          =   315
            Left            =   0
            TabIndex        =   67
            Top             =   120
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   556
            Caption         =   "Data Rel. Fiscal"
            Text            =   ""
            Formato         =   0
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   3465
         Left            =   -99969
         TabIndex        =   44
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT201.frx":021F
         Begin VB.TextBox txtDocumentosIntimacao 
            Appearance      =   0  'Flat
            Height          =   2775
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   46
            Text            =   "TPRT201.frx":0247
            Top             =   120
            Width           =   10575
         End
         Begin VTOcx.cmdVISUAL cmdImprimeTermoIntimacao 
            Height          =   375
            Left            =   9360
            TabIndex        =   45
            Top             =   3000
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Imprimir"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   3465
         Left            =   -99969
         TabIndex        =   43
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT201.frx":025D
         Begin VB.Frame fmeNf 
            BorderStyle     =   0  'None
            Height          =   650
            Left            =   240
            TabIndex        =   65
            Top             =   2800
            Visible         =   0   'False
            Width           =   8415
            Begin VTOcx.txtVISUAL txtNumeroNota 
               Height          =   480
               Left            =   135
               TabIndex        =   24
               Tag             =   "A"
               Top             =   105
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   847
               Caption         =   "Nota Fiscal"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtEmissao 
               Height          =   480
               Left            =   1920
               TabIndex        =   25
               Tag             =   "A"
               Top             =   105
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   847
               Caption         =   "Emissão"
               Text            =   ""
               TipoLetras      =   0
               Formato         =   0
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtValorNota 
               Height          =   480
               Left            =   4200
               TabIndex        =   26
               Tag             =   "A"
               Top             =   120
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   847
               Caption         =   "R$ Nota"
               Text            =   ""
               TipoLetras      =   0
               Formato         =   5
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtBase 
               Height          =   480
               Left            =   6600
               TabIndex        =   27
               Tag             =   "A"
               Top             =   120
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   847
               Caption         =   "R$ Base C."
               Text            =   ""
               TipoLetras      =   0
               Formato         =   5
               AlinhamentoRotulo=   1
            End
         End
         Begin VTOcx.cmdVISUAL cmdImprimeAutoInfracao 
            Height          =   375
            Left            =   8760
            TabIndex        =   29
            Top             =   3015
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   661
            Caption         =   "Auto"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtPeriodo 
            Height          =   480
            Left            =   2160
            TabIndex        =   21
            Tag             =   "A"
            Top             =   2295
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   847
            Caption         =   "Periodo"
            Text            =   ""
            TipoLetras      =   0
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtISSDevido 
            Height          =   480
            Left            =   6840
            TabIndex        =   23
            Tag             =   "A"
            Top             =   2295
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   847
            Caption         =   "R$ Valor Devido"
            Text            =   ""
            TipoLetras      =   0
            Formato         =   5
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtVencimento 
            Height          =   480
            Left            =   4440
            TabIndex        =   22
            Tag             =   "A"
            Top             =   2295
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   847
            Caption         =   "Vencimento"
            Text            =   ""
            TipoLetras      =   0
            Formato         =   0
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL4 
            Height          =   375
            Left            =   8760
            TabIndex        =   28
            Top             =   2415
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   661
            Caption         =   "Incluir"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
            CorFundo        =   16777088
         End
         Begin VTOcx.grdVISUAL grdAuto 
            Height          =   2115
            Left            =   120
            TabIndex        =   47
            Top             =   120
            Width           =   10530
            _ExtentX        =   18574
            _ExtentY        =   3731
            Caption         =   "Documentos"
            OcultarRodape   =   -1  'True
         End
         Begin VTOcx.cboVISUAL cboDocumento 
            Height          =   510
            Left            =   240
            TabIndex        =   20
            Tag             =   "C"
            Top             =   2280
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   900
            Caption         =   "Documento"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
            Alinhamento     =   1
         End
         Begin VTOcx.cmdVISUAL cmdTermo 
            Height          =   375
            Left            =   9720
            TabIndex        =   66
            Top             =   3000
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   661
            Caption         =   "Termo"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   3465
         Left            =   -99969
         TabIndex        =   16
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT201.frx":0285
         Begin VB.CheckBox chk 
            Caption         =   "CONTRATO DE LOCAÇÃO (SE HOUVER)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   18
            Left            =   6720
            TabIndex        =   55
            Top             =   1920
            Width           =   3855
         End
         Begin VB.CheckBox chk 
            Caption         =   "CODIGO DO IPTU, N. DE FUNCIONÁRIOS E AREA CONTRUIDA DO ESTABELECIMENTO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   17
            Left            =   6720
            TabIndex        =   54
            Top             =   1200
            Width           =   3975
         End
         Begin VB.CheckBox chk 
            Caption         =   "CONTADOR RESPONSAVEL COM N. DE CRC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   6720
            TabIndex        =   53
            Top             =   960
            Width           =   3855
         End
         Begin VB.CheckBox chk 
            Caption         =   "CARTAO DO CNPJ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   6720
            TabIndex        =   52
            Top             =   600
            Width           =   3375
         End
         Begin VB.CheckBox chk 
            Caption         =   "ALVARA DE FUNCIONAMENTO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   6720
            TabIndex        =   51
            Top             =   240
            Width           =   3375
         End
         Begin VB.CheckBox chk 
            Caption         =   "FOTOCOPIA, CONTRATO SOCIAL, RG E CPF DO REPRESENTANTE LEGAL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   13
            Left            =   3480
            TabIndex        =   50
            Top             =   2400
            Width           =   6015
         End
         Begin VB.CheckBox chk 
            Caption         =   "COMPROVANTE DE PAG ISSQN-RF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   3480
            TabIndex        =   42
            Top             =   2040
            Width           =   5295
         End
         Begin VB.CheckBox chk 
            Caption         =   "REGISTRO OU ESCRITURA DO IMÓVEL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   3480
            TabIndex        =   41
            Top             =   1680
            Width           =   5415
         End
         Begin VB.CheckBox chk 
            Caption         =   "BALANÇO DO EXERCÍCIO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   3480
            TabIndex        =   40
            Top             =   1320
            Width           =   5295
         End
         Begin VB.CheckBox chk 
            Caption         =   "BLOCO E/OU NOTAS FISCAIS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   3480
            TabIndex        =   39
            Top             =   960
            Width           =   5295
         End
         Begin VB.CheckBox chk 
            Caption         =   "LIVRO FISCAL - ISSQN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   3480
            TabIndex        =   38
            Top             =   600
            Width           =   4935
         End
         Begin VB.CheckBox chk 
            Caption         =   "DEMONSTRAÇÃO DE RESULTADO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   3480
            TabIndex        =   37
            Top             =   240
            Width           =   5175
         End
         Begin VB.CheckBox chk 
            Caption         =   "IPTU:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   40
            TabIndex        =   36
            Top             =   2400
            Width           =   3255
         End
         Begin VB.CheckBox chk 
            Caption         =   "COMPROVANTE DE PAG. DO ISSQN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   40
            TabIndex        =   35
            Top             =   2040
            Width           =   3615
         End
         Begin VB.CheckBox chk 
            Caption         =   "DECLARAÇÃO DE IRPJ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   40
            TabIndex        =   34
            Top             =   1680
            Width           =   2655
         End
         Begin VB.CheckBox chk 
            Caption         =   "BALANCETE DE VERIFICACAO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   40
            TabIndex        =   33
            Top             =   1320
            Width           =   3255
         End
         Begin VB.CheckBox chk 
            Caption         =   "LIVRO RAZAO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   40
            TabIndex        =   32
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox chk 
            Caption         =   "LIVRO CAIXA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   40
            TabIndex        =   31
            Top             =   600
            Width           =   1695
         End
         Begin VB.CheckBox chk 
            Caption         =   "COMPARECER AO SETOR DE TRIBUTOS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   40
            TabIndex        =   30
            Top             =   240
            Width           =   3495
         End
         Begin VTOcx.cboVISUAL cboTipoAcaoFiscal 
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Tag             =   "Tributo"
            Top             =   2880
            Width           =   8970
            _ExtentX        =   15822
            _ExtentY        =   556
            Caption         =   "Ação Fiscal      "
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL1 
            Height          =   375
            Left            =   9240
            TabIndex        =   19
            Top             =   2880
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Imprimir"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3465
         Left            =   30
         TabIndex        =   12
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT201.frx":02AD
         Begin VTOcx.txtVISUAL txtDocumento 
            Height          =   540
            Left            =   120
            TabIndex        =   14
            Top             =   2355
            Width           =   10560
            _ExtentX        =   18627
            _ExtentY        =   953
            Caption         =   "Documento(s)"
            Text            =   ""
            Requerido       =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.grdVISUAL grdDOCS 
            Height          =   2595
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   4577
            Caption         =   "Documentos"
            OcultarRodape   =   -1  'True
         End
         Begin VTOcx.cmdVISUAL cmdADD 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   3000
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Adicionar"
            Acao            =   1
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdImprimirDiligencia 
            Height          =   375
            Left            =   1560
            TabIndex        =   17
            Top             =   3000
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Imprimir"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdAddPadrao 
            Height          =   375
            Left            =   9360
            TabIndex        =   56
            Top             =   3000
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Padrão"
            Acao            =   1
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1138
      Formulario      =   "Ordem de Serviço"
      Descricao       =   "Detalhes da Ordem de Serviço"
      Icone           =   "TPRT201.frx":02D5
   End
End
Attribute VB_Name = "TPRT201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim os As OrdemServico
Dim Tipo As Integer
Dim mOrdemServico As Long

Public Property Get OrdemServico() As String
    OrdemServico = mOrdemServico
End Property
 
Public Property Let OrdemServico(ByVal newValue As String)
    mOrdemServico = newValue
End Property
Private Sub cboDocumento_Click()
    If cboDocumento.Coluna(1).valor = 1 Then
        fmeNf.Visible = True
        txtPeriodo.Enabled = False
        txtVencimento.SetFocus
        txtPeriodo = 0
    Else
        fmeNf.Visible = False
        txtPeriodo.Enabled = True
        txtPeriodo.SetFocus
    End If
End Sub
Private Sub cboFinalizar_Click()
    Dim f As Integer
    f = cboFinalizar.ListIndex
    Dim Texto As String
    Dim Sql As String, umTab As String
    Dim rs As VSRecordset
    Dim Data As String, dataAbertura As String, dataEncerramento As String, dataTermo As String, dataNotificacao As String, dataAuto As String, ano As String, Processo As String
    Texto = ""
    umTab = "        "
    dataEncerramento = Format(Now, "dd/mm/yyyy")
        
    'podendo ser qualquer um dos status, fazer assim
    If Tipo = 2 Then 'diligencia
        Sql = "select data_autorizacao from tab_bcp_ordem_servico where codigo=" & mOrdemServico
        If Bdados.AbreTabela(Sql, rs) Then
            Do While rs.EOF = False
                Data = Format(rs("data_autorizacao"), "dd/mm/yyyy")
                rs.MoveNext
            Loop
        End If
        Texto = umTab & "Considerando que, o contribuinte acima qualificado, Mediante intimação escrita, atraves do Termo de Diligência Fiscal," _
        & " datado de " & Data & ", apresentou à autoridade administrativa todas as informações de que disponham com relação à bens, negócios ou atividades de terceiros;" _
        & vbCrLf & vbCrLf & umTab & "Atentando, ainda, para o fato de que a o fisco municipal procedeu a Diligencia Fiscal, analisando todos os documentos e informações apresentadas, conforme relato no Relatório Fiscal;" _
        & vbCrLf & vbCrLf & umTab & "Levando-se em conta, finalmente, a conclusão da DILIGÊNCIA FISCAL, Lavra-se o presente TERMO DE ENCERRAMENTO DE DILIGÊNCIA FISCAL." _
        & vbCrLf & vbCrLf & umTab & "Desta forma em " & dataEncerramento & ", lavra-se o presente TERMO DE ENCERRAMENTO DE DILIGÊNCIA FISCAL NÃO IMPEDE, que a Autoridade Competente, quando julgar necessário, abrir nova fiscalização."
        txtTexto = Texto
    Else
        Sql = "select data_abertura,ano,data_limite from vis_bcp_tiaf where SERVICO=" & mOrdemServico & " and os_status='TERMO DE INICIO DE AF'"
        If Bdados.AbreTabela(Sql, rs) Then
            Do While rs.EOF = False
                dataTermo = Format(rs("data_abertura"), "dd/mm/yyyy")
                rs.MoveNext
            Loop
        Else
            Mensagem ("Não consta registro de Inicio de Acão Fiscal para este PROCESSO, atualiza a OS e informe o data inicio do Termo")
            Exit Sub
        End If
        Sql = "select data_abertura,ano,data from vis_bcp_tiaf where SERVICO=" & mOrdemServico & " and cod_tipo"
        
        Select Case f
            Case 0
                Processo = "Termo de Diligência"
                Sql = Sql & "=2"
            Case 1
                Processo = "Termo de Notificação"
                Sql = Sql & "=6"
            Case 2
                Processo = "Termo de Auto de Infração"
                Sql = Sql & "=5"
            Case 3
                Processo = "Termo de Início de Ação Fiscal"
                Sql = Sql & " in (3,4)"
        End Select
        If Bdados.AbreTabela(Sql, rs) Then
            Do While rs.EOF = False
                dataAbertura = Format(rs("data_abertura"), "dd/mm/yyyy")
                Data = Format(rs("data"), "dd/mm/yyyy")
                ano = rs("ano")
                rs.MoveNext
            Loop
        Else
            Mensagem ("Não consta registro de data de PRAZO para a OS selecionada, atualiza a OS e define o prazo do processo")
            Exit Sub
        End If
        Sql = "select data_abertura,ano,data from vis_bcp_tiaf where SERVICO=" & mOrdemServico & " and os_status='NOTIFICACAO DO CONTRIBUINTE'"
        If Bdados.AbreTabela(Sql, rs) Then
            Do While rs.EOF = False
                dataNotificacao = Format(rs("data_abertura"), "dd/mm/yyyy")
                rs.MoveNext
            Loop
        Else
            Mensagem ("Não consta registro de Notificação para este PROCESSO, atualiza a OS e informe o data inicio da Notificação")
            Exit Sub
        End If
        
        'Sql = "select data from tab_bcp_os_encerramento where cod_os=" & mOrdemServico
        'If Bdados.AbreTabela(Sql, rs) Then
         '   Do While rs.EOF = False
          '      dataEncerramento = Format(rs("data"), "dd/mm/yyyy")
           '     rs.MoveNext
            'Loop
        'End If
        Texto = Texto & umTab & "Tendo em vista a Ordem de Serviço N. " & mOrdemServico & "/" & ano & " aberta em " & Data & ", e a lavratura do " _
        & "Termo de Início da Ação Fiscal, datado de " & dataTermo & ". Segue o presente Termo de Encerramento de Fiscalização. " _
        & vbCrLf & vbCrLf & umTab & "Considerando que, mediante Intimação escrita, o contribuinte, apresentou  à autoridade fiscal os documentos constantes " _
        & "no Termo de Recebimento de documentos, anexo, solicitados através de Notificação Fiscal datado de " & dataNotificacao & ". " _
        & vbCrLf & vbCrLf & umTab & "Considerando também que a Autoridade Fiscal, agindo de acordo com os Princípios da Legalidade, da Igualdade, " _
        & "da Capacidade Contribuitiva e da Moralidade, fiscalizou, analisou e processou todas as informações e documentos " _
        & "apresentados pelo Contribuinte, conforme Relatório Fiscal datado de " & Format(txtDataRelatorioFiscal, "dd/mm/yyyy") & ", constante na ordem de serviço acima citada."
        If Tipo = 5 Then 'auto de infraçao
            Sql = "select data_abertura,ano,data from vis_bcp_tiaf where SERVICO=" & mOrdemServico & " and os_status='AUTO INFRACAO'"
            If Bdados.AbreTabela(Sql, rs) Then
                Do While rs.EOF = False
                    dataAuto = Format(rs("data_abertura"), "dd/mm/yyyy")
                    rs.MoveNext
                Loop
            Else
                Mensagem ("Não consta registro de Auto de Infração para este PROCESSO, atualiza a OS e informe o data inicio do Auto de Infração")
                Exit Sub
            End If
            Texto = Texto & vbCrLf & vbCrLf & umTab & "Considerando finalmente, que, a Autoridade Fiscal, lavrou o Auto de Infração Nº " & mOrdemServico & ", datado de " & dataAuto & ", e o mesmo," _
            & " estando de acordo com as formalidades e procedimento legalmente exigidos conforme legislação vigente." _
            & vbCrLf & vbCrLf & umTab & "Desta forma em " & dataEncerramento & ", lavra-se o presente " & cboFinalizar.Text & " para o contribuinte acima citado. Destarte, o presente " _
            & "" & cboFinalizar.Text & " NÃO IMPEDE, que a Autoridade Competente, quando julgar necessário, possa observado os princípios constitucionais e tributários."
        Else
            Texto = Texto & vbCrLf & vbCrLf & umTab & "Desta forma em " & dataEncerramento & ", lavra-se o presente " & cboFinalizar.Text & " para o contribuinte acima citado. Destarte, o presente " _
            & "" & cboFinalizar.Text & " NÃO IMPEDE, que a Autoridade Competente, quando julgar necessário, possa observado os princípios constitucionais e tributários."
        End If
        txtTexto = Texto
    End If
End Sub
Private Sub cmdADD_Click()
    If os.IncluirDocumentosDiligencia(mOrdemServico, txtDocumento) Then
        exibeDiligencia
    End If
End Sub
Private Sub exibeDiligencia()
    If Not grdDOCS.preencher(Bdados, "SELECT COD_ITEM AS ITEM, DESCRICAO FROM TAB_BCP_OS_DILIGENCIA WHERE COD_ORDEM=" & mOrdemServico) Then
    End If
End Sub
Private Sub exibeAuto()
    If Not grdAuto.preencher(Bdados, "SELECT PERIODO,DATA_VENCIMENTO AS VENCIMENTO, ISS_DEVIDO,NUMERO_NOTA AS NOTA, DATA_EMISSAO AS EMISSAO,VALOR_NOTA AS TOTAL, BASE_CALCULO AS BASE_CALC, NOME_DOC AS TRIBUTO FROM VIS_BCP_AUTO WHERE SERVICO=" & mOrdemServico) Then
    End If
End Sub
Private Sub cmdAddPadrao_Click()
    Dim p As Long, i As Long, f As Long
    'i = InputBox("DILIGENCIA INICIAL", "ORDEM DE SERVIÇO")
    'f = InputBox("DILIGENCIA FINAL", "ORDEM DE SERVIÇO")
    'For p = i To f
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & mOrdemServico & ",1, 'FOTOCÓPIA - CONTRATO SOCIAL, IDENTIDADE E CPF DOS SÓCIOS')")
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & mOrdemServico & ",2, 'ALVARÁ DE FUNCIONAMENTO')")
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & mOrdemServico & ",3, 'CARTÃO DO CNPJ')")
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & mOrdemServico & ",4, 'CONTADOR RESPONSÁVEL (CRC, CPF, ENDEREÇO)')")
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & mOrdemServico & ",5, 'CONTRATO DE LOCAÇÃO (SE HOUVER)')")
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & mOrdemServico & ",6, 'INSCRIÇÃO DE IMOBILIÁRIA (IPTU), NÚMERO DE FUNCIONÁRIOS E ÁREA CONSTRUÍDA DO ESTABELECIMENTO')")
        exibeDiligencia
    'Next p
End Sub
Private Sub cmdFinalizar_Click()
    Dim ft As String
    ft = cboFinalizar.Text
    
    Dim os As OrdemServico
    Set os = New OrdemServico
        
    If Bdados.Executa("DELETE FROM TAB_BCP_OS_ENCERRAMENTO WHERE COD_OS=" & mOrdemServico) Then
    End If
    If Bdados.AbreTabela("SELECT * FROM VIS_BCP_OS_ENCERRAMENTO WHERE SERVICO=" & mOrdemServico) Then
        os.ImprimeEncerramento (mOrdemServico)
    Else
        If Not Util.Confirma("Confirma o encerramento deste processo?") Then
            Exit Sub
        End If
        If os.Encerrar(mOrdemServico, ft, txtTexto) Then
            os.ImprimeEncerramento (mOrdemServico)
        End If
    End If
    If Bdados.Executa("UPDATE TAB_BCP_ORDEM_SERVICO SET TIPO=99 WHERE CODIGO=" & mOrdemServico) Then
    End If
End Sub
Private Sub cmdImprimeAutoInfracao_Click()
    'os.ImprimeAutoInfracao mOrdemServico
    Dim Form As TPRT110
    Set Form = New TPRT110
    Form.carregar mOrdemServico
    Call txtIm_LostFocus
    Exit Sub
End Sub
Private Sub cmdImprimeTermoIntimacao_Click()
    Bdados.Executa ("UPDATE TAB_BCP_ORDEM_SERVICO SET INTIMACAO_DOCUMENTOS='" & txtDocumentosIntimacao & "' WHERE CODIGO=" & mOrdemServico)
    os.ImprimeTermoIntimacao mOrdemServico
End Sub
Private Sub cmdImprimir_Click()
    'If txtIm = "-" Then
        'mOrdemServico = 0
    'End If
    If mOrdemServico > 0 Then
    
        os.Imprime mOrdemServico, True
    Else
        i = InputBox("OS INICIAL", "ORDEM DE SERVIÇO")
        f = InputBox("OS FINAL", "ORDEM DE SERVIÇO")
        Dim xos As Long
        For xos = i To f
            os.Imprime completar(xos), True
        Next xos
    End If
End Sub
Private Sub cmdImprimirDiligencia_Click()
       ' Dim p As Long, i As Long, f As Long
        'i = InputBox("DILIGENCIA INICIAL", "ORDEM DE SERVIÇO")
        'f = InputBox("DILIGENCIA FINAL", "ORDEM DE SERVIÇO")
        
        'For p = i To f
          '   os.ImprimeDiligencia p
           ' os.ImprimeDiligencia p
        'Next p
        os.ImprimeDiligencia mOrdemServico
End Sub

Private Sub cmdIncluirHistorico_Click()
    If cboFiscalHistorico.Text = "" And txtObservacao.Text = "" Then
        Mensagem ("Informe o Fiscal e Observacao para gerar o histórico")
        Exit Sub
    End If
    If os.SalvarHistorico(OrdemServico, cboFiscal.Text, txtObservacao.Text) Then
        preencherHistorico
    End If
    txtObservacao.Text = ""
End Sub

Private Sub cmdIncluirInscricao_Click()
    If os.inserirInscricao(mOrdemServico, txtIm.Text) Then
    End If
    'If os.atualizaProcesso(mOrdemServico, CInt(cboTIPO.Coluna(1).Valor)) Then
    'End If
    Dim Form As TPRT109
    Set Form = New TPRT109
    Form.carregar mOrdemServico
    preencher
    Exit Sub
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub cmdSalvar_Click()
    If Len(cboFiscal) = 0 Then
        Mensagem "Informe o Primeiro Fiscal"
        Exit Sub
    End If
    If Len(cboFiscal2) = 0 Then
        Mensagem "Informe o Segundo Fiscal"
        Exit Sub
    End If
    
    If Len(cboTIPO) = 0 Then
        Mensagem "Informe o Tipo de OS"
        Exit Sub
    End If
    'If Bdados.AbreTabela("SELECT * FROM TAB_BCP_ORDEM_SERVICO WHERE IM_CONTRIBUINTE='" & txtIm & "' AND TIPO=" & cboTIPO.Coluna(1).valor & " AND PERIODO_INICIAL=" & CLng(txtExercicioInicial) & " AND PERIODO_FINAL=" & CLng(txtExercicioFinal)) Then
        'Mensagem "Já existe uma ordem de serviço aberta para este contribuinte"
        'Exit Sub
    'Else
        If Len(txtIm) > 0 Then
            mOrdemServico = os.Salvar(CInt(cboTIPO.Coluna(1).valor), Replace(txtIm, ".", ""), cboFiscal.Text, txtExercicioInicial, txtExercicioFinal, CDate(Format(Now, "DD/MM/YYYY")), Format(Now, "HH:mm"), txtOutrosServicos, cboFiscal2.Text)
            If mOrdemServico > 0 Then
                Mensagem "Ordem de serviço aberta com sucesso"
                preencher
                Call cmdImprimir_Click
            End If
        Else
            Dim nos As Integer, i As Integer
            Dim total As String
            total = InputBox("Numero de OS", "Geração de OS", 0)
            nos = CInt(IIf(Len(total) = 0, 0, total))
            If nos > 0 Then
                For i = 1 To nos
                    If os.Salvar(CInt(cboTIPO.Coluna(1).valor), "", cboFiscal.Text, txtExercicioInicial, txtExercicioFinal, CDate(Format(Now, "DD/MM/YYYY")), Format(Now, "HH:mm"), txtOutrosServicos, cboFiscal2.Text) Then
                    End If
                Next i
                Mensagem (nos & " Ordens de serviço abertas com sucesso!")
                Unload Me
            End If
        End If
    
        'Dim nos As Integer, i As Integer
        'nos = CInt(InputBox("Numero de OS", "Geração de OS"))
        'For i = 1 To nos
            
        'Next i
        
        
    'End If
End Sub

Private Sub cmdTermo_Click()
    Dim Form As TPRT111
    Set Form = New TPRT111
    Form.carregar mOrdemServico
        
End Sub

Private Sub cmdVISUAL1_Click()
    Dim docs As String
    Dim i As Integer
    docs = ""
    For i = 0 To chk.Count - 1
        If chk(i).Value = 1 Then
            docs = docs & "X"
        Else
            docs = docs & " "
        End If
    Next i
    'Bdados.Executa ("UPDATE TAB_BCP_ORDEM_SERVICO  DOCUMENTOS='" & docs & "' WHERE CODIGO=" & mOrdemServico)
    Bdados.Executa ("UPDATE TAB_BCP_ORDEM_SERVICO SET TIPO_ACAO_FISCAL='" & cboTipoAcaoFiscal.Text & "', DOCUMENTOS='" & docs & "' WHERE CODIGO=" & mOrdemServico)
    
    os.Imprime mOrdemServico, False
End Sub
Private Sub cmdVISUAL2_Click()
    Unload Me
End Sub
Private Sub cmdVISUAL3_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL4_Click()
    Dim campos As String
    Dim valores As String
    Dim trib As String, periodo As String
    Dim codtrib As Integer
    codtrib = cboDocumento.Coluna(1).valor
    If txtPeriodo = 0 Or txtPeriodo = "" Then
        Mensagem ("Informe o periodo e para ISSQN necessita-se dos dados fiscais abaixo")
        Exit Sub
    End If
    Select Case codtrib
        Case 1
            trib = "11130501"
        Case 2
            trib = "11130504"
        Case 3
            trib = "11210101"
        Case 4
            trib = "11120203"
    End Select
    If Len(txtPeriodo) > 4 Then
        periodo = Format(txtPeriodo, "000000")
    Else
        periodo = Format(txtPeriodo, "0000")
    End If
    periodo = "-" & periodo & "-"
    campos = "COD_OS, TIPO_DOC, PERIODO,DATA_VENCIMENTO,ISS_DEVIDO,NUMERO_NOTA,DATA_EMISSAO,VALOR_NOTA,BASE_CALCULO,JUROS,MULTA,ATUALIZACAO,ISS_ATUALIZADO,TRIBUTO"
    valores = Bdados.PreparaValor(mOrdemServico, codtrib, CStr(periodo), Format(txtVencimento.Text, "dd/mm/yy"), CCur(txtISSDevido), txtNumeroNota, Format(txtEmissao.Text, "dd/mm/yy"), txtValorNota, txtBase, 0, 0, 0, CCur(txtISSDevido), trib)
    If Bdados.InsereDados("Tab_bcp_auto", valores, campos) Then
        exibeAuto
    Else
        
    End If
    txtNumeroNota = ""
    txtEmissao = ""
    txtBase = ""
    txtValorNota = ""
End Sub

Private Sub Form_Load()
    Set os = New OrdemServico
    os.PreencheCombo cboTIPO
    os.PreencheCombo cboStatusRelatorio
    os.PreencheComboDocumento cboDocumento
    txtDataRelatorioFiscal = Format(Now, "DD/MM/YYYY")
    
    
    Dim ano As Long
    ano = Format(Now, "yyyy")
    txtExercicioFinal = ano
    txtExercicioInicial = ano
    Set os = New OrdemServico
    txtOutrosServicos = ""
    'SSTab.Tabs(2).Visible = False
    'SSTab.Tabs(3).Visible = False
    cboTipoAcaoFiscal.AddItem "TIAF - ACAO FISCAL"
    cboTipoAcaoFiscal.AddItem "FISCALIZAÇÃO"
    cboTipoAcaoFiscal.AddItem "AUTO DE INFRAÇÃO"
    cboTipoAcaoFiscal.AddItem "NOTIFICAÇÃO"
    cboTipoRelatorio.AddItem "RELATORIO ANALITICO"
    cboTipoRelatorio.AddItem "RESUMO POR FISCAL"
    cboTipoRelatorio.AddItem "RESUMO POR BAIRRO"
    cboTipoRelatorio.AddItem "RESUMO POR STATUS"
    cboTipoRelatorio.AddItem "PRAZO DOS PROCESSOS"
    cboFinalizar.AddItem "TERMO DE ENCERRAMENTO DE DILIGÊNCIA"
    cboFinalizar.AddItem "TERMO DE ENCERRAMENTO DE NOTIFICAÇÃO"
    cboFinalizar.AddItem "TERMO DE ENCERRAMENTO DE AUTO DE INFRAÇÃO"
    cboFinalizar.AddItem "TERMO DE ENCERRAMENTO DO INÍCIO DE AÇÃO FISCAL"
    
    Dim rs As VSRecordset
    If Bdados.AbreTabela("SELECT TUS_COD_USUARIO FROM TAB_USUARIO ORDER BY TUS_COD_USUARIO", rs) Then
        Do While Not rs.EOF
            cboFiscal.AddItem rs(0)
            cboFiscal2.AddItem rs(0)
            cboFiscalRelatorio.AddItem rs(0)
            cboFiscalHistorico.AddItem rs(0)
            rs.MoveNext
        Loop
    End If
    preencher
    
End Sub
Public Sub preencherHistorico()
    If Not grdHistorico.preencher(Bdados, "SELECT DATA,FISCAL,OBSERVACAO FROM TAB_BCP_OS_HISTORICO WHERE SERVICO='" & OrdemServico & "' ORDER BY DATA DESC") Then
            'Mensagem "Não existem ordem de serviço para o contribuinte"
    End If
End Sub
Public Sub preencher()
    lblServico = "SERVIÇO: "
    'cboFiscal = Rs("FISCAL")
    'cboFiscal2 = Rs("FISCAL2")
    'txtIm = Rs("INSCRICAO")
    'cboTIPO = Rs("ETAPA")
    cboFiscal.Enabled = True
    cboFiscal2.Enabled = True
    cboTIPO.Enabled = True
    txtOutrosServicos.Enabled = True
    cmdSalvar.Enabled = True
    Dim x As Integer
    Dim rs As VSRecordset
    If mOrdemServico > 0 Then
        cmdImprimir.Enabled = True
        cmdIncluirInscricao.Enabled = True
        If Bdados.AbreTabela("SELECT * FROM VIS_BCP_ORDEM_SERVICO WHERE SERVICO=" & mOrdemServico, rs) Then
            'txtDAM = IIf(IsNull(Rs(0)), "", Rs(0))
            lblServico = "SERVIÇO: " & rs("SERVICO")
            lblStatus = "STATUS: " & rs("SITUACAO")
            cboFiscal = rs("FISCAL")
            cboFiscal2 = rs("FISCAL2")
            cboTipoAcaoFiscal = IIf(IsNull(rs("TIPO_ACAO_FISCAL")), "", rs("TIPO_ACAO_FISCAL"))
            If Not (IsNull(rs("INSCRICAO"))) Then
                txtIm = rs("INSCRICAO")
            End If
            cboTIPO = rs("ETAPA")
            Tipo = rs("TIPO")
    
            If txtIm <> "" Then
                Call txtIm_LostFocus
            End If
            cboFiscal.Enabled = False
            cboFiscal2.Enabled = False
            
            txtIm.Enabled = False
            cmdPesquisaInscricao.Enabled = False
            
            cboTIPO.Enabled = False
            txtOutrosServicos.Enabled = False
            cmdSalvar.Enabled = False
            
            exibeAuto
            exibeDiligencia
            preencherHistorico
            habilitaBotoes (True)
            If Not (IsNull(rs("DOCUMENTOS"))) Then
                For x = 1 To Len(rs("DOCUMENTOS"))
                Dim mark As String 'marcado com X ou espaço em brando cada DOCUMENTO na tela
                    mark = Mid(rs("DOCUMENTOS"), x, 1)
                    chk(x - 1).Value = IIf(mark = "X", 1, 0)
                Next x
            End If
            
        End If
    Else
        cmdImprimir.Enabled = False
        cmdIncluirInscricao.Enabled = False
        txtIm.Enabled = True
        cmdPesquisaInscricao.Enabled = True
        habilitaBotoes (False)
        
    End If
    
End Sub
Private Sub habilitaBotoes(e As Boolean)
    cmdADD.Enabled = e
    cmdAddPadrao.Enabled = e
    cmdImprimirDiligencia.Enabled = e
    Dim x As Integer
    x = 1
    Do While x <= 7
        tabEtapa.Tabs(x).Enabled = e
        x = x + 1
    Loop
End Sub
Private Sub grdOrdens_Click()
    
    lblOrdem = "ORDEM DE SERVIÇO N. " & mOrdemServico
    
    'txtDocumentosIntimacao = grdOrdens.SelectedItem.SubItems(26)
End Sub


Private Sub txtEmissao_LostFocus()
    If cboDocumento.Coluna(1).valor = 1 Then
        If Len(txtEmissao) Then
            txtPeriodo = Mid(txtEmissao, 3, 4) & Right(txtEmissao, 4)
        Else
            txtPeriodo = 0
        End If
    End If
End Sub

Private Sub txtIm_LostFocus()
     Dim rs As VSRecordset
    Dim Sql As String
    'lblOrdem = ""
    If Len(txtIm) > 0 Or txtIm <> "-" Then
        Sql = " Select * from Tab_Contribuinte " _
            & " where tci_im='" & txtIm & "'"
        'If Not Conexao Is Nothing Then Set Bdados = Conexao
        If Bdados.AbreTabela(Sql, rs) Then
            txtAtividade = Imposto.BuscaNomeCAE("" & rs("tci_tae_cae"))
            txtNOME = "" & rs("tci_nome")
            txtEndereco = "" & rs("tci_logradouro") & " " & rs("tci_nome_logradouro") & "," & rs("tci_numero") & " " & rs("tci_complemento") & " " & rs("tci_bairro")
            
        Else
            txtAtividade = ""
            txtNOME = ""
            txtEndereco = ""
        End If
    Else
            txtAtividade = ""
            txtNOME = ""
            txtEndereco = ""
    End If
    
End Sub
Private Sub cmdImprimirGerencial_Click()
    'On Error GoTo trata
    Dim CodRelatorio As Integer
    Set Rpt = Nothing
   ' Screen.MousePointer = vbHourglass
    CodRelatorio = cboTipoRelatorio.ListIndex + 1
    Set Rpt = New VSRelatorio
        If DefinirArquivo(CodRelatorio) Then
            If DefinirFormulas(CodRelatorio) Then
                If DefinirSelecao(CodRelatorio) Then
                    'Rpt.Arvore = False
                    Rpt.visualizar
                End If
            End If
        End If

'trata:
 '   Screen.MousePointer = vbNormal
End Sub

Private Function DefinirArquivo(CodRelatorio As Integer) As Boolean
    DefinirArquivo = Rpt.DefinirArquivo(Bdados, App.Path + "\TBCP100" & CodRelatorio & ".rpt")
End Function
Private Function DefinirFormulas(CodRelatorio As Integer) As Boolean
    DefinirFormulas = True
    Rpt.LimparFormulas
    Rpt.Formulas "CLIENTE", Temp.PegaParametro(Bdados, "CLIENTE")
    Rpt.Formulas "TITULO", cboTipoRelatorio
End Function
Private Function DefinirSelecao(CodRelatorio As Integer) As Boolean
    Dim Filtro As String, descricao As String
    DefinirSelecao = True
    Filtro = ""
    descricao = ""
    descricao = "Filtrado por:"
    Dim rel As Integer
    rel = cboTipoRelatorio.ListIndex + 1
    Select Case rel
    Case 1, 2, 3, 4
        If Len(cboFiscalRelatorio) > 0 Then
             Filtro = Filtro & "{VIS_BCP_ORDEM_SERVICO.FISCAL} ='" & cboFiscalRelatorio.Text & "' AND "
             descricao = descricao & " Fiscal:" & cboFiscalRelatorio.Text
         End If
         If Len(cboStatusRelatorio) > 0 Then
             Filtro = Filtro & "{VIS_BCP_ORDEM_SERVICO.COD_TIPO} =" & cboStatusRelatorio.Coluna(1).valor & " AND "
             descricao = descricao & " Status:" & cboStatusRelatorio.Text
         End If
         If Len(txtDataInicio) > 0 And Len(txtDataFim) > 0 Then
            Filtro = Filtro & "{VIS_BCP_ORDEM_SERVICO.DATA} >=" & retornarData(txtDataInicio) & " AND {VIS_BCP_ORDEM_SERVICO.DATA} <=" & retornarData(txtDataFim) & " AND "
            descricao = descricao & " Data Emissão: " & txtDataInicio.Text & " até " & txtDataFim.Text
         End If
         If Filtro <> "" Then
             Filtro = Left(Filtro, Len(Filtro) - 4)
         End If
    Case 5
        Dim rt As VSRecordset
        Dim dl As Date, datual As Date
        Dim os As Long
        If Bdados.AbreTabela("select codigo, data_limite,os_status from vis_bcp_tiaf", rt) Then
        End If
        datual = Format(Now, "DD/MM/YYYY")
        Do While rt.EOF = False
            If CDate(rt(1)) < datual Then
                If Bdados.Executa("UPDATE TAB_BCP_TIAF SET TIAF_STATUS='FORA DO PRAZO' WHERE COD_OS=" & rt(0) & " and os_status='" & rt(2) & "'") Then
                End If
            End If
            rt.MoveNext
        Loop
    End Select
    If DefinirSelecao = True Then
        If Filtro <> "" Then
            Rpt.Selecao = Filtro
            Rpt.Formulas "FILTRO", descricao
        End If
    End If
End Function
Private Function retornarData(Data As String) As String
    Dim nd As String
    Data = Replace(Data, "/", "")
    nd = Right(Data, 4) & "-" & Mid(Data, 3, 2) & "-" & Left(Data, 2)
    retornarData = "'" & nd & "'"
End Function
Private Sub formatar()
    If Len(txtIm) > 0 Then
        Dim Im As String
        Im = txtIm.Text
        Im = Left(Im, 8) & "-" & Right(Im, 2)
        'txtIM = im
    Else
        txtIm = ""
    End If
End Sub
Private Sub txtValorNota_LostFocus()
    txtBase = txtValorNota
End Sub
Private Function completar(os As Long) As Long
    Dim osC As String
    osC = "10000000"
    osC = Left(osC, Len(osC) - Len(CStr(os)))
    osC = osC & os
    completar = osC
End Function

