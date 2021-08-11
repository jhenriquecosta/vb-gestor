VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TPRT101 
   Caption         =   "TPRT101"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10890
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
      Height          =   2085
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   10860
      Begin VB.TextBox txtIm 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1080
         TabIndex        =   0
         Top             =   200
         Width           =   2055
      End
      Begin VTOcx.cboVISUAL cboRestricao 
         Height          =   315
         Left            =   690
         TabIndex        =   12
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
         TabIndex        =   5
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
      Begin VTOcx.txtVISUAL txtNOME 
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Tag             =   "C"
         Top             =   525
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   529
         Caption         =   "Nome     "
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtExercicioInicial 
         Height          =   300
         Left            =   9675
         TabIndex        =   2
         Tag             =   "C"
         Top             =   540
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Caption         =   ""
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtExercicioFinal 
         Height          =   300
         Left            =   9675
         TabIndex        =   4
         Tag             =   "C"
         Top             =   885
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Caption         =   "  "
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Tag             =   "C"
         Top             =   885
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtAtividade 
         Height          =   300
         Left            =   240
         TabIndex        =   16
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
         Left            =   6300
         TabIndex        =   1
         Tag             =   "C"
         Top             =   525
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         Caption         =   "Fiscal/Periodo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtOutrosServicos 
         Height          =   300
         Left            =   4440
         TabIndex        =   6
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
         Left            =   6300
         TabIndex        =   3
         Tag             =   "C"
         Top             =   885
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         Caption         =   "Fiscal/Periodo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
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
         TabIndex        =   76
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   14
         Top             =   1590
         Width           =   45
      End
   End
   Begin ActiveTabs.SSActiveTabs tabEtapa 
      Height          =   3855
      Left            =   0
      TabIndex        =   17
      Top             =   3120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6800
      _Version        =   131082
      TabCount        =   7
      TagVariant      =   ""
      Tabs            =   "TPRT101.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   3465
         Left            =   30
         TabIndex        =   69
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":01AE
         Begin VTOcx.cboVISUAL cboTipoRelatorio 
            Height          =   315
            Left            =   1000
            TabIndex        =   70
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
            TabIndex        =   71
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
            TabIndex        =   72
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
            TabIndex        =   73
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
            TabIndex        =   74
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
            TabIndex        =   75
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
         Left            =   30
         TabIndex        =   61
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":01D6
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   3465
         Left            =   30
         TabIndex        =   56
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":01FE
         Begin VB.TextBox txtDocumentosIntimacao 
            Appearance      =   0  'Flat
            Height          =   2775
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   58
            Text            =   "TPRT101.frx":0226
            Top             =   120
            Width           =   10575
         End
         Begin VTOcx.cmdVISUAL cmdImprimeTermoIntimacao 
            Height          =   375
            Left            =   9360
            TabIndex        =   57
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
         Left            =   30
         TabIndex        =   54
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":023C
         Begin VB.Frame fmeNf 
            BorderStyle     =   0  'None
            Height          =   650
            Left            =   240
            TabIndex        =   77
            Top             =   2800
            Visible         =   0   'False
            Width           =   8415
            Begin VTOcx.txtVISUAL txtNumeroNota 
               Height          =   480
               Left            =   135
               TabIndex        =   35
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
               TabIndex        =   36
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
               TabIndex        =   37
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
               TabIndex        =   38
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
            TabIndex        =   40
            Top             =   2895
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
            TabIndex        =   32
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   39
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
            TabIndex        =   59
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
            TabIndex        =   31
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
            TabIndex        =   78
            Top             =   2880
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
         Left            =   30
         TabIndex        =   25
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":0264
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
            Left            =   6600
            TabIndex        =   67
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
            Left            =   6600
            TabIndex        =   66
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
            Left            =   6600
            TabIndex        =   65
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
            Left            =   6600
            TabIndex        =   64
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
            Left            =   6600
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   53
            Top             =   2040
            Width           =   5295
         End
         Begin VB.CheckBox chk 
            Caption         =   "DCTF"
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
            TabIndex        =   52
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
            TabIndex        =   51
            Top             =   1320
            Width           =   5295
         End
         Begin VB.CheckBox chk 
            Caption         =   "NOTAS FISCAIS"
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            Left            =   240
            TabIndex        =   47
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
            Left            =   240
            TabIndex        =   46
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
            Left            =   240
            TabIndex        =   45
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
            Left            =   240
            TabIndex        =   44
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
            Left            =   240
            TabIndex        =   43
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
            Left            =   240
            TabIndex        =   42
            Top             =   600
            Width           =   1695
         End
         Begin VB.CheckBox chk 
            Caption         =   "LIVRO DIARIO"
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
            Left            =   240
            TabIndex        =   41
            Top             =   240
            Width           =   1695
         End
         Begin VTOcx.cboVISUAL cboTipoAcaoFiscal 
            Height          =   315
            Left            =   240
            TabIndex        =   28
            Tag             =   "Tributo"
            Top             =   2880
            Width           =   7410
            _ExtentX        =   13070
            _ExtentY        =   556
            Caption         =   "Ação Fiscal      "
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL1 
            Height          =   375
            Left            =   7920
            TabIndex        =   29
            Top             =   2880
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Imprimir"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL3 
            Height          =   375
            Left            =   9360
            TabIndex        =   30
            Top             =   2880
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Sai&r"
            Acao            =   7
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3465
         Left            =   30
         TabIndex        =   18
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":028C
         Begin VTOcx.grdVISUAL grdDOCS 
            Height          =   2355
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   4154
            Caption         =   "Documentos"
            OcultarRodape   =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtDocumento 
            Height          =   540
            Left            =   0
            TabIndex        =   20
            Top             =   2355
            Width           =   10680
            _ExtentX        =   18838
            _ExtentY        =   953
            Caption         =   "Documento(s)"
            Text            =   ""
            Requerido       =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.cmdVISUAL cmdADD 
            Height          =   375
            Left            =   6480
            TabIndex        =   21
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
            Left            =   7920
            TabIndex        =   26
            Top             =   3000
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Imprimir"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL2 
            Height          =   375
            Left            =   9360
            TabIndex        =   27
            Top             =   3000
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Sai&r"
            Acao            =   7
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdAddPadrao 
            Height          =   375
            Left            =   120
            TabIndex        =   68
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
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3465
         Left            =   30
         TabIndex        =   22
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   6112
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":02B4
         Begin VB.TextBox txtImAberto 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   480
            TabIndex        =   24
            Text            =   "Inscrição"
            Top             =   3000
            Width           =   2055
         End
         Begin VTOcx.grdVISUAL grdOrdens 
            Height          =   2865
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   5054
            OcultarRodape   =   -1  'True
         End
         Begin VTOcx.cmdVISUAL cmdSalvar 
            Height          =   375
            Left            =   6360
            TabIndex        =   7
            Top             =   3000
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Iniciar"
            Acao            =   3
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdSair 
            Height          =   375
            Left            =   9300
            TabIndex        =   10
            Top             =   3000
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Sai&r"
            Acao            =   7
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdImprimir 
            Height          =   375
            Left            =   7800
            TabIndex        =   8
            Top             =   3000
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Imprimir"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdIncluirInscricao 
            Height          =   375
            Left            =   2640
            TabIndex        =   9
            Top             =   3000
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   661
            Caption         =   "Atualizar OS"
            Acao            =   8
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   1138
      Icone           =   "TPRT101.frx":02DC
   End
   Begin VB.Label lblOrdem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   0
      TabIndex        =   55
      Top             =   2760
      Width           =   10815
   End
End
Attribute VB_Name = "TPRT101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim os As ordemServico
Dim codOs As Long, Tipo As Integer

Private Sub cboDocumento_Click()
    If cboDocumento.Coluna(1).Valor = 1 Then
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

Private Sub cmdADD_Click()
    If os.IncluirDocumentosDiligencia(codOs, txtDocumento) Then
        exibeDiligencia
    End If
End Sub
Private Sub exibeDiligencia()
    If Not grdDOCS.Preencher(Bdados, "SELECT COD_ITEM AS ITEM, DESCRICAO FROM TAB_BCP_OS_DILIGENCIA WHERE COD_ORDEM=" & codOs) Then
    End If
End Sub
Private Sub exibeAuto()
    If Not grdAuto.Preencher(Bdados, "SELECT PERIODO,DATA_VENCIMENTO AS VENCIMENTO, ISS_DEVIDO,NUMERO_NOTA AS NOTA, DATA_EMISSAO AS EMISSAO,VALOR_NOTA AS TOTAL, BASE_CALCULO AS BASE_CALC, NOME_DOC AS TRIBUTO FROM VIS_BCP_AUTO WHERE CODIGO=" & codOs) Then
    End If
End Sub

Private Sub cmdAddPadrao_Click()
    Dim p As Long, i As Long, f As Long
    'i = InputBox("DILIGENCIA INICIAL", "ORDEM DE SERVIÇO")
    'f = InputBox("DILIGENCIA FINAL", "ORDEM DE SERVIÇO")
    'For p = i To f
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & codOs & ",1, 'FOTOCÓPIA - CONTRATO SOCIAL, IDENTIDADE E CPF DOS SÓCIOS')")
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & codOs & ",2, 'ALVARÁ DE FUNCIONAMENTO')")
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & codOs & ",3, 'CARTÃO DO CNPJ')")
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & codOs & ",4, 'CONTADOR RESPONSÁVEL (CRC, CPF, ENDEREÇO)')")
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & codOs & ",5, 'CONTRATO DE LOCAÇÃO (SE HOUVER)')")
        Bdados.Executa ("INSERT INTO TAB_BCP_OS_DILIGENCIA (COD_ORDEM,COD_ITEM,DESCRICAO) VALUES(" & codOs & ",6, 'INSCRIÇÃO DE IMOBILIÁRIA (IPTU), NÚMERO DE FUNCIONÁRIOS E ÁREA CONSTRUÍDA DO ESTABELECIMENTO')")
        exibeDiligencia
    'Next p
End Sub

Private Sub cmdImprimeAutoInfracao_Click()
    'os.ImprimeAutoInfracao codOs
    Dim Form As TPRT110
    Set Form = New TPRT110
    Form.carregar codOs
    Call txtIm_LostFocus
    habilitarTab (Tipo)
    Exit Sub
End Sub

Private Sub cmdImprimeTermoIntimacao_Click()
    Bdados.Executa ("UPDATE TAB_BCP_ORDEM_SERVICO SET INTIMACAO_DOCUMENTOS='" & txtDocumentosIntimacao & "' WHERE CODIGO=" & codOs)
    os.ImprimeTermoIntimacao codOs
End Sub

Private Sub cmdImprimir_Click()
    If txtIm = "-" Then
        codOs = 0
    End If
    If codOs > 0 Then
        os.Imprime codOs, True
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
        os.ImprimeDiligencia codOs
End Sub
Private Sub cmdIncluirInscricao_Click()
    If os.inserirInscricao(codOs, txtImAberto.Text) Then
    End If
    'If os.atualizaProcesso(codOs, CInt(cboTIPO.Coluna(1).Valor)) Then
    'End If
    Dim Form As TPRT109
    Set Form = New TPRT109
    txtIm = txtImAberto
    Form.carregar codOs
    Call txtIm_LostFocus
    Exit Sub
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
    'If Bdados.AbreTabela("SELECT * FROM TAB_BCP_ORDEM_SERVICO WHERE IM_CONTRIBUINTE='" & txtIm & "' AND TIPO=" & cboTIPO.Coluna(1).Valor & " AND PERIODO_INICIAL=" & CLng(txtExercicioInicial) & " AND PERIODO_FINAL=" & CLng(txtExercicioFinal)) Then
     '   Mensagem "Já existe uma ordem de serviço aberta para este contribuinte"
      '  Exit Sub
    'Else
        If Len(txtIm) > 0 Then
            If os.Salvar(CInt(cboTIPO.Coluna(1).Valor), Replace(txtIm, ".", ""), cboFiscal.Text, txtExercicioInicial, txtExercicioFinal, CDate(Format(Now, "DD/MM/YYYY")), Format(Now, "HH:mm"), txtOutrosServicos, cboFiscal2.Text) Then
            Else
            End If
        Else
            Dim nos As Integer, i As Integer
            nos = CInt(InputBox("Numero de OS", "Geração de OS"))
            For i = 1 To nos
                If os.Salvar(CInt(cboTIPO.Coluna(1).Valor), "-", cboFiscal.Text, txtExercicioInicial, txtExercicioFinal, CDate(Format(Now, "DD/MM/YYYY")), Format(Now, "HH:mm"), txtOutrosServicos, cboFiscal2.Text) Then
                Else
                End If
            Next i
        End If
    
        'Dim nos As Integer, i As Integer
        'nos = CInt(InputBox("Numero de OS", "Geração de OS"))
        'For i = 1 To nos
            
        'Next i
        Mensagem "Ordem de serviço aberta com sucesso"
        exibir
    'End If
End Sub

Private Sub cmdTermo_Click()
    Dim Form As TPRT111
    Set Form = New TPRT111
    Form.carregar codOs
        
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
    Bdados.Executa ("UPDATE TAB_BCP_ORDEM_SERVICO SET TIPO_ACAO_FISCAL='" & cboTipoAcaoFiscal.Text & "', DOCUMENTOS='" & docs & "' WHERE CODIGO=" & codOs)
    os.Imprime codOs, False
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
    codtrib = cboDocumento.Coluna(1).Valor
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
    valores = Bdados.PreparaValor(codOs, codtrib, CStr(periodo), Format(txtVencimento.Text, "dd/mm/yy"), CCur(txtISSDevido), txtNumeroNota, Format(txtEmissao.Text, "dd/mm/yy"), txtValorNota, txtBase, 0, 0, 0, CCur(txtISSDevido), trib)
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
    Set os = New ordemServico
    habilitarTab (1)
    os.PreencheCombo cboTIPO
    os.PreencheCombo cboStatusRelatorio
    os.PreencheComboDocumento cboDocumento
    
    Dim ano As Long
    ano = Format(Now, "yyyy")
    txtExercicioFinal = ano
    txtExercicioInicial = ano
    Set os = New ordemServico
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
    
    
    Dim rs As VSRecordset
    If Bdados.AbreTabela("SELECT TUS_COD_USUARIO FROM TAB_USUARIO ORDER BY TUS_COD_USUARIO", rs) Then
        Do While Not rs.EOF
            cboFiscal.AddItem rs(0)
            cboFiscal2.AddItem rs(0)
            cboFiscalRelatorio.AddItem rs(0)
            rs.MoveNext
        Loop
    End If
    'cboTipoNf.AddItem "PRESTADA"
    'cboTipoNf.AddItem "TOMADA"
    'cboStatusNf.AddItem "EMITDA"
    'cboStatusNf.AddItem "CANCELADA"
End Sub
Private Sub grdOrdens_dblClick()
    On Error GoTo err
    'Call txtIm_LostFocus
    'exibeDiligencia
    'exibeAuto
    'formatar
    'habilitarTab (Tipo)
    
    Tipo = grdOrdens.SelectedItem.SubItems(19)
    codOs = grdOrdens.SelectedItem
    cboTIPO = grdOrdens.SelectedItem.SubItems(2)
    cboFiscal = grdOrdens.SelectedItem.SubItems(8)
    cboFiscal2 = grdOrdens.SelectedItem.SubItems(36)
    
    txtExercicioInicial = grdOrdens.SelectedItem.SubItems(3)
    txtExercicioFinal = grdOrdens.SelectedItem.SubItems(4)
    txtIm = grdOrdens.SelectedItem.SubItems(5)
    txtImAberto = grdOrdens.SelectedItem.SubItems(5)
    lblOrdem = "ORDEM DE SERVIÇO N. " & codOs
    'formatar
    Call txtIm_LostFocus
    habilitarTab (Tipo)
    exibeAuto
    exibeDiligencia
    grdOrdens_Click
err:
End Sub
Private Sub grdOrdens_Click()
    
    lblOrdem = "ORDEM DE SERVIÇO N. " & codOs
    
    'txtDocumentosIntimacao = grdOrdens.SelectedItem.SubItems(26)
End Sub

Private Sub txtEmissao_LostFocus()
    If cboDocumento.Coluna(1).Valor = 1 Then
        If Len(txtEmissao) Then
            txtPeriodo = Mid(txtEmissao, 3, 4) & Right(txtEmissao, 4)
        Else
            txtPeriodo = 0
        End If
    End If
End Sub

Private Sub txtIm_LostFocus()
     Dim rs As VSRecordset
    Dim Sql As String, Im As String
    Im = txtIm
    lblOrdem = ""
    If Len(Trim(Im)) > 0 And Im <> "-" Then
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
    exibir
End Sub
Private Sub exibir()
    Dim Im As String
    Im = txtIm
    If Len(Im) > 0 And Im <> "-" Then
        If Not grdOrdens.Preencher(Bdados, "SELECT * FROM VIS_BCP_ORDEM_SERVICO_REL WHERE INSCRICAO='" & txtIm & "' ORDER BY CODIGO DESC") Then
            Mensagem "Não existem ordem de serviço para o contribuinte"
        Else
        End If
    Else
        If Not grdOrdens.Preencher(Bdados, "SELECT * FROM VIS_BCP_ORDEM_SERVICO_REL ORDER BY CODIGO DESC") Then
            Mensagem "Não existem ordem de serviço para o contribuinte"
        Else
        End If
    End If
    formatar
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
             Filtro = Filtro & "{VIS_BCP_ORDEM_SERVICO.COD_TIPO} =" & cboStatusRelatorio.Coluna(1).Valor & " AND "
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
Private Sub habilitarTab(Etapa As Integer)
    Dim x As Integer
    x = 2
    Do While x <= 5
        tabEtapa.Tabs(x).Enabled = False
        x = x + 1
    Loop
    'tabEtapa.Tabs(1).Enabled = True
    x = Etapa
    Select Case Etapa
        Case 2
            x = 2
        Case 4
            x = 3
        Case 5
            x = 4
    End Select
    tabEtapa.Tabs(x).Enabled = True
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

