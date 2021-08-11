VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{D8A7CA9C-BFF7-11D5-9D50-00D0590D0C80}#1.0#0"; "CTREEOPT.OCX"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU402a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU402a"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   7290
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10500
         TabIndex        =   54
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin ActiveTabs.SSActiveTabs ssTab 
      Height          =   6240
      Left            =   0
      TabIndex        =   31
      Top             =   870
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   11007
      _Version        =   131082
      TabCount        =   3
      CaptionOrientation=   1
      PictureBackgroundStyle=   1
      HotTracking     =   1
      BeginProperty FontHotTracking {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "TCIU402a.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   5850
         Left            =   -99969
         TabIndex        =   52
         Top             =   360
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   10319
         _Version        =   131082
         TabGuid         =   "TCIU402a.frx":00CE
         Begin VTOcx.fraVISUAL fraVISUAL6 
            Height          =   5115
            Left            =   6540
            TabIndex        =   56
            Top             =   90
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   9022
            Altura          =   1905
            Caption         =   " Dados Adicionais da Edificação"
            CorTexto        =   0
            CorFaixa        =   12632256
            CorFundo        =   16777215
            Ocultavel       =   0   'False
            Begin VTOcx.fraVISUAL fraUnidade 
               Height          =   915
               Left            =   690
               TabIndex        =   61
               Top             =   2940
               Visible         =   0   'False
               Width           =   4005
               _ExtentX        =   7064
               _ExtentY        =   1614
               Altura          =   1905
               Caption         =   " "
               CorTexto        =   0
               CorFaixa        =   16777215
               CorFundo        =   16777215
               Ocultavel       =   0   'False
            End
            Begin VTOcx.cboVISUAL cboUnidades 
               Height          =   510
               Left            =   2340
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   1740
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   900
               Caption         =   "Unidades Edificadas"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtAnoconstrucao 
               Height          =   495
               Left            =   60
               TabIndex        =   58
               Tag             =   "111"
               Top             =   870
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   873
               Caption         =   "Ano Construção"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtAreaUnidade 
               Height          =   495
               Left            =   60
               TabIndex        =   57
               Tag             =   "112"
               Top             =   300
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   873
               Caption         =   "Área Edificada Unidade"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
         End
         Begin cTreeOpt.XTreeOpt treCadBP 
            Height          =   5115
            Left            =   0
            TabIndex        =   53
            Tag             =   "2"
            Top             =   90
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   9022
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            Indentation     =   400.251983642578
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5850
         Left            =   -99969
         TabIndex        =   51
         Top             =   360
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   10319
         _Version        =   131082
         TabGuid         =   "TCIU402a.frx":00F6
         Begin cTreeOpt.XTreeOpt treCadBT 
            Height          =   5055
            Left            =   0
            TabIndex        =   38
            Tag             =   "1"
            Top             =   90
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   8916
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            Indentation     =   400.251983642578
         End
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   5085
            Left            =   6540
            TabIndex        =   55
            Top             =   90
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   8969
            Altura          =   1905
            Caption         =   " Dimensões do Terreno"
            CorTexto        =   0
            CorFaixa        =   12632256
            CorFundo        =   16777215
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtDataCadastro 
               Height          =   495
               Left            =   2880
               TabIndex        =   62
               Top             =   2040
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   873
               Caption         =   "Data Cadastramento"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtAno 
               Height          =   495
               Left            =   2880
               TabIndex        =   47
               Top             =   2640
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   873
               Caption         =   "Ano Aquisição"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtAreaEdifTotal 
               Height          =   495
               Left            =   2880
               TabIndex        =   50
               Tag             =   "113"
               Top             =   4380
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   873
               Caption         =   "Área Edificada Total"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtTestadaCampo 
               Height          =   495
               Left            =   150
               TabIndex        =   46
               Tag             =   "107"
               Top             =   4350
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Nº de Testadas"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtTestada4 
               Height          =   495
               Left            =   120
               TabIndex        =   42
               Tag             =   "105"
               Top             =   2040
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Testada 4"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtTestada3 
               Height          =   495
               Left            =   120
               TabIndex        =   41
               Tag             =   "103"
               Top             =   1470
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Testada 3"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtTestada2 
               Height          =   495
               Left            =   120
               TabIndex        =   40
               Tag             =   "101"
               Top             =   900
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Testada 2"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtProfundidade 
               Height          =   495
               Left            =   2880
               TabIndex        =   48
               Tag             =   "115"
               Top             =   3240
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   873
               Caption         =   "Profundidade"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtAreaLote 
               Height          =   495
               Left            =   2880
               TabIndex        =   49
               Tag             =   "108"
               Top             =   3810
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   873
               Caption         =   "Área do Lote"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtTrechoLogr4 
               Height          =   495
               Left            =   120
               TabIndex        =   45
               Tag             =   "106"
               Top             =   3810
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Trecho/Seção Logr. 4"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtTrechoLogr3 
               Height          =   495
               Left            =   120
               TabIndex        =   44
               Tag             =   "104"
               Top             =   3210
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Trecho/Seção Logr. 3"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtTrechoLogr2 
               Height          =   495
               Left            =   120
               TabIndex        =   43
               Tag             =   "102"
               Top             =   2640
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Trecho/Seção Logr. 2"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtTestadaPrin 
               Height          =   495
               Left            =   120
               TabIndex        =   39
               Tag             =   "100"
               Top             =   300
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Testada Principal"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   16777215
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5850
         Left            =   30
         TabIndex        =   32
         Top             =   360
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   10319
         _Version        =   131082
         TabGuid         =   "TCIU402a.frx":011E
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1815
            Left            =   0
            TabIndex        =   33
            Top             =   90
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   3201
            Altura          =   1905
            Caption         =   " Localização do Imóvel"
            CorTexto        =   0
            CorFaixa        =   12632256
            CorFundo        =   16777215
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboTipo 
               Height          =   315
               Left            =   9480
               TabIndex        =   3
               Top             =   330
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               Caption         =   "Tipo"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtCodReduzido 
               Height          =   315
               Left            =   3630
               TabIndex        =   1
               Top             =   330
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   556
               Caption         =   "Cód. Reduzido"
               Text            =   ""
               Restricao       =   2
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtInscImob 
               Height          =   315
               Left            =   120
               TabIndex        =   0
               Top             =   330
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   556
               Caption         =   "Insc. Imobiliária"
               Text            =   ""
               Restricao       =   2
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtIncAnterior 
               Height          =   315
               Left            =   6360
               TabIndex        =   2
               Top             =   330
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               Caption         =   "Insc. Anterior"
               Text            =   ""
               Restricao       =   2
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtCodLogr 
               Height          =   315
               Left            =   420
               TabIndex        =   4
               Top             =   690
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               Caption         =   "Código Logr."
               Text            =   ""
               Restricao       =   2
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtLogradouroImovel 
               Height          =   315
               Left            =   2610
               TabIndex        =   34
               Top             =   690
               Width           =   6525
               _ExtentX        =   11509
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtNumero 
               Height          =   315
               Left            =   9810
               TabIndex        =   5
               Top             =   690
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               Caption         =   "Número"
               Text            =   ""
               CorFundo        =   16777215
            End
            Begin VTOcx.cboVISUAL cboLoteamento 
               Height          =   315
               Left            =   6270
               TabIndex        =   7
               Top             =   1050
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   556
               Caption         =   "Loteamento"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.cboVISUAL cboBairro 
               Height          =   315
               Left            =   990
               TabIndex        =   6
               Top             =   1050
               Width           =   5265
               _ExtentX        =   9287
               _ExtentY        =   556
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtQuadra 
               Height          =   315
               Left            =   900
               TabIndex        =   8
               Top             =   1410
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               Caption         =   "Quadra"
               Text            =   ""
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtLote 
               Height          =   315
               Left            =   2340
               TabIndex        =   9
               Top             =   1410
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
               Caption         =   "Lote"
               Text            =   ""
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtSecao 
               Height          =   315
               Left            =   3600
               TabIndex        =   10
               Top             =   1410
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               Caption         =   "Seção"
               Text            =   ""
               CorFundo        =   16777215
            End
            Begin VTOcx.cboVISUAL cboPredio 
               Height          =   315
               Left            =   4950
               TabIndex        =   11
               Top             =   1410
               Width           =   6435
               _ExtentX        =   11351
               _ExtentY        =   556
               Caption         =   "Prédio/Condomínio"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   16777215
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   735
            Left            =   0
            TabIndex        =   35
            Top             =   1920
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   1296
            Altura          =   1905
            Caption         =   " Dados Adicionais da Localização"
            CorTexto        =   0
            CorFaixa        =   12632256
            CorFundo        =   16777215
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtComplemento 
               Height          =   315
               Left            =   5550
               TabIndex        =   15
               Top             =   330
               Width           =   5835
               _ExtentX        =   10292
               _ExtentY        =   556
               Caption         =   "Complemento"
               Text            =   ""
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtLoja 
               Height          =   315
               Left            =   360
               TabIndex        =   12
               Top             =   330
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   556
               Caption         =   "No. Loja/Sala"
               Text            =   ""
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtApto 
               Height          =   315
               Left            =   2460
               TabIndex        =   13
               Top             =   330
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               Caption         =   "No. Apto."
               Text            =   ""
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtBloco 
               Height          =   315
               Left            =   4170
               TabIndex        =   14
               Top             =   330
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "Bloco"
               Text            =   ""
               CorFundo        =   16777215
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   2325
            Left            =   30
            TabIndex        =   36
            Top             =   2670
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   4101
            Altura          =   1905
            Caption         =   " Dados do Proprietário"
            CorTexto        =   0
            CorFaixa        =   12632256
            CorFundo        =   16777215
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtRg 
               Height          =   315
               Left            =   -30
               TabIndex        =   25
               Top             =   1920
               Width           =   5355
               _ExtentX        =   9446
               _ExtentY        =   556
               Caption         =   "RG/Órgão Expedidor"
               Text            =   ""
               TipoLetras      =   0
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtCpfProp 
               Height          =   315
               Left            =   7080
               TabIndex        =   26
               Top             =   1920
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   556
               Caption         =   "CPF/CNPJ"
               Text            =   ""
               Restricao       =   2
               CorFundo        =   16777215
            End
            Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
               Height          =   315
               Left            =   3510
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   330
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.txtVISUAL txtUF 
               Height          =   315
               Left            =   7800
               TabIndex        =   29
               Top             =   1410
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               Caption         =   "UF"
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.cboVISUAL cboMunicipio 
               Height          =   315
               Left            =   900
               TabIndex        =   23
               Top             =   1410
               Width           =   6765
               _ExtentX        =   11933
               _ExtentY        =   556
               Caption         =   "Município"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtCEP 
               Height          =   315
               Left            =   9330
               TabIndex        =   24
               Top             =   1380
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               Caption         =   "CEP"
               Text            =   ""
               Formato         =   4
               Restricao       =   2
               CorFundo        =   16777215
            End
            Begin VTOcx.cboVISUAL cboLogrProp 
               Height          =   315
               Left            =   720
               TabIndex        =   18
               Top             =   690
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.cboVISUAL cboBairroProp 
               Height          =   315
               Left            =   1170
               TabIndex        =   21
               Top             =   1050
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   556
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   16777215
               Editavel        =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtNumeroProp 
               Height          =   315
               Left            =   9780
               TabIndex        =   20
               Top             =   690
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               Caption         =   "Número"
               Text            =   ""
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtContribuinte 
               Height          =   315
               Left            =   3870
               TabIndex        =   17
               Top             =   330
               Width           =   7455
               _ExtentX        =   13150
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtInscMunicipal 
               Height          =   315
               Left            =   60
               TabIndex        =   16
               Top             =   330
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   556
               Caption         =   "Inscrição/Cadastro"
               Text            =   ""
               CorFundo        =   16777215
            End
            Begin VTOcx.cboVISUAL cboNomeLogrProp 
               Height          =   315
               Left            =   3870
               TabIndex        =   19
               Top             =   690
               Width           =   5745
               _ExtentX        =   10134
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   16777215
               Editavel        =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtComplementoProp 
               Height          =   315
               Left            =   6360
               TabIndex        =   22
               Top             =   1050
               Width           =   4965
               _ExtentX        =   8758
               _ExtentY        =   556
               Caption         =   "Complemento"
               Text            =   ""
               CorFundo        =   16777215
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   735
            Left            =   30
            TabIndex        =   37
            Top             =   5070
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   1296
            Altura          =   1905
            Caption         =   " Ocupante do Imóvel"
            CorTexto        =   0
            CorFaixa        =   12632256
            CorFundo        =   16777215
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtCPFOcupante 
               Height          =   315
               Left            =   8130
               TabIndex        =   28
               Top             =   330
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   556
               Caption         =   "CPF/CNPJ"
               Text            =   ""
               Restricao       =   2
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtOcupante 
               Height          =   315
               Left            =   1170
               TabIndex        =   27
               Top             =   330
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   556
               Caption         =   "Nome"
               Text            =   ""
               CorFundo        =   16777215
            End
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   1138
      Icone           =   "TCIU402a.frx":0146
   End
End
Attribute VB_Name = "TCIU402a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Edificacoes() As Edificacao
Private Tree As New TreeViewBci
Function ImovelJaCadastrado(Inscricao As String)
    Dim rs As VSRecordset
    Dim Sql As String
    Dim aux As String
       
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        Sql = "Select * from tab_imovel_historico where tim_ic_auxiliar ='" & Inscricao & "'"
    Else
        Sql = "Select * from tab_imovel_historico where tIM_ic ='" & Inscricao & "'"
    End If
    If Bdados.AbreTabela(Sql, rs) Then
        ImovelJaCadastrado = True
    End If
End Function

Private Sub cboMunicipio_Click()
    txtUF = cboMunicipio.Coluna(2).Valor
End Sub

Private Sub cboTipo_Click()
    If cboTipo.Coluna(1).Valor = 2 Then
        ssTab.Tabs(3).Enabled = False
    Else
        ssTab.Tabs(3).Enabled = True
    End If
End Sub

Private Sub cboUnidades_Click()
    If CInt(Nvl(cboUnidades.Text, 0)) <> 0 Then
        txtAnoconstrucao = Edificacoes(CInt(Nvl(cboUnidades.Text, 0))).Componente(txtAnoconstrucao.Tag)
        txtAreaUnidade = Edificacoes(CInt(cboUnidades.Text)).Componente(txtAreaUnidade.Tag)
        Tree.SetaTreeViewEdificacao treCadBP, Edificacoes, CInt(cboUnidades.Text)
        fraUnidade.Visible = True
    End If
    DoEvents
End Sub

Private Sub CmdConsultaContribuinte_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscMunicipal
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.Tag <> "" Then
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            txtCodReduzido = Left(Trim(Me.Tag), Len(Trim(Me.Tag)) - 5)
            txtCodReduzido_LostFocus
        Else
            txtInscImob = Left(Trim(Me.Tag), Len(Trim(Me.Tag)) - 5)
            Call txtInscImob_LostFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static i As Byte
    If KeyCode = 113 Then
        Select Case ssTab.SelectedTab.Index
            Case 1
                ssTab.Tabs(2).Selected = True
                treCadBT.SetFocus
            Case 2
                ssTab.Tabs(3).Selected = True
                treCadBP.SetFocus
            Case 3
                ssTab.Tabs(1).Selected = True
                txtInscImob.SetFocus
        End Select
        DoEvents
    End If
End Sub

Private Sub Form_Load()
    Tree.CarregaListaComponentes treCadBT
    Tree.CarregaListaComponentes treCadBP
    cboLoteamento.Preencher Bdados, "SELECT TLO_COD_LOTEAMENTO,TLO_DESCRICAO FROM TAB_LOTEAMENTO ORDER BY 2", 1
    cboPredio.Preencher Bdados, "SELECT TED_COD_EDIFICIO,TED_DESCRICAO FROM TAB_EDIFICIO ORDER BY 2", 1
    cboBairro.Preencher Bdados, "SELECT TBA_COD_BAIRRO,TBA_NOME FROM TAB_BAIRRO ORDER BY 2", 1
    cboBairroProp.Preencher Bdados, "SELECT TBA_COD_BAIRRO,TBA_NOME FROM TAB_BAIRRO ORDER BY 2", 1
    cboMunicipio.Preencher Bdados, "SELECT TMU_COD_MUNICIPIO,TMU_NOME,TUF_UF FROM TAB_MUNICIPIO,TAB_UF " & _
        "WHERE TMU_TUF_COD_UF = TUF_COD_UF ORDER  BY 2", 1
    cboLogrProp.Preencher Bdados, "SELECT TTL_COD_TIP_LOGR,TTL_NOME FROM TAB_TIPO_LOGR ORDER BY 2", 1
    cboNomeLogrProp.Preencher Bdados, "SELECT DISTINCT tlg_nome FROM TAB_LOGRADOURO ORDER BY 1"
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Tag
    cboTipo.PreencherGeral Bdados, "TIPO LOTE"
    ReDim Edificacoes(1 To 1) As Edificacao
    
End Sub

Private Sub txtCodLogr_LostFocus()
    On Error GoTo TrataErro
    Dim Query As String
    Dim rs As VSRecordset
    If Trim(txtCodLogr) = "" Then Exit Sub
    Query = "SELECT TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome, " & _
        " TAB_BAIRRO.TBA_NOME FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
        " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
         " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & txtCodLogr & "'"
    If Bdados.AbreTabela(Query, rs) Then
        txtLogradouroImovel = rs(0) & " " & rs(1)
    Else
        Avisa "Código de logradouro inválido."
        txtCodLogr.SetFocus
    End If
    Bdados.FechaTabela rs
    Exit Sub
TrataErro:
    If Err.Number = 3265 Then
        Resume Next
    Else
        Util.Erro Err.Description
    End If
End Sub


Private Sub txtCodReduzido_LostFocus()
    txtInscImob_LostFocus
End Sub

Private Sub txtInscImob_LostFocus()
    Dim rs As VSRecordset
    Dim Sql As String
    Dim i As Integer
    If Trim(txtInscImob) = "" And (txtCodReduzido) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If Trim(txtCodReduzido) <> "" Then
        txtInscImob.Enabled = True
        txtCodReduzido.Enabled = False
    Else
        txtInscImob.Enabled = False
        txtCodReduzido.Enabled = True
    End If
    
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        If txtCodReduzido = "" Then
            Sql = "Select * from tab_imovel_historico where tim_ic_auxiliar ='" & txtInscImob & "'  AND TIM_COD_MUDANCA=" & CInt(Right(Trim(Me.Tag), 5))
        Else
            Sql = "Select * from tab_imovel_historico where tim_ic ='" & txtCodReduzido & "' AND TIM_COD_MUDANCA=" & CInt(Right(Trim(Me.Tag), 5))
        End If
    Else
        Sql = "Select * from tab_imovel_historico where tIM_ic ='" & txtInscImob & "'  AND TIM_COD_MUDANCA=" & CInt(Right(Trim(Me.Tag), 5))
    End If
    txtCodReduzido = Trim(txtCodReduzido)
    If Bdados.AbreTabela(Sql, rs) Then
        txtIncAnterior = "" & IIf(rs!tim_ic_anterior = 0, "", rs!tim_ic_anterior)
        txtCodLogr = "" & rs!tim_tlg_cod_logradouro
        txtCodLogr_LostFocus
        txtNumero = "" & rs!tim_numero
        txtAno = "" & rs!tim_ano_aquis
        cboLoteamento.SetarLinha "" & rs!tim_loteamento, 0
        cboPredio.SetarLinha "" & rs!tim_ted_cod_edificio, 0
        cboBairro.SetarLinha "" & rs!tim_TBA_COD_BAIRRO, 0
        txtQuadra = "" & rs!tim_QUADRA
        txtLote = "" & rs!tim_lote
        txtOcupante = "" & rs!tim_ocupante
        txtCPFOcupante = "" & rs!tim_cgc_cpf_ocupante
        
        txtIncAnterior = "" & rs!tim_ic_anterior
        txtSecao = "" & rs!tim_secao
        
        txtComplemento = "" & rs!tim_complemento
        txtBloco = "" & rs!TIM_BLOCO
        txtApto = "" & rs!TIM_APTO
        txtLoja = "" & rs!TIM_SALA_LOJA
        If txtCodReduzido <> "" Then
            txtInscImob = Trim("" & rs!tim_ic_auxiliar)
        End If
        txtDataCadastro = "" & rs!tim_DATA_CADASTRO
        'VOU PEGAR O CONTRIBUINTE
        txtInscMunicipal = "" & rs!tim_tci_im
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            txtCodReduzido = rs!TIM_IC
        Else
            txtCodReduzido = rs!tim_ic_auxiliar
        End If
        txtInscMunicipal_LostFocus
        cboTipo.SetarLinha Nvl("" & rs!tim_tipo_imovel, 0), 1
        cboTipo_Click
        'VOU PEGAR OS DETALHES
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            Sql = "Select tim_ic from tab_imovel where tim_ic ='" & txtCodReduzido & "' or tim_ic_condominio ='" & txtCodReduzido & "'"
        Else
            Sql = "Select tim_ic from tab_imovel where tim_ic ='" & txtInscImob & "' or tim_ic_condominio ='" & txtInscImob & "'"
        End If
        If Bdados.AbreTabela(Sql) Then
            Dim Inscricoes As String
            Inscricoes = ""
            Bdados.Tabela.MoveFirst
            Do
                Inscricoes = Inscricoes & "'" & Bdados.Tabela(0) & "',"
                Bdados.Tabela.MoveNext
            Loop While Not Bdados.Tabela.EOF
            Inscricoes = Left(Inscricoes, Len(Inscricoes) - 1)
        End If
        Sql = "Select * from TAB_DETALHE_IMOVEL_HISTORICO where TDI_TIM_IC in (" & Inscricoes & ") AND TDI_TIM_COD_MUDANCA=" & CInt(Right(Trim(Me.Tag), 5)) & " order by tdi_tim_ic_unidade,tdi_tgc_cod_grupo asc"
        If Bdados.AbreTabela(Sql, rs) Then
            cboUnidades.Clear
            Do
                If CInt(Nvl(rs!tdi_tim_ic_unidade, 0)) = 0 Then 'LOTE
                    If rs!tdi_tgc_cod_grupo >= 100 Then
                        Dim Controle As Control
                        On Error Resume Next
                        For Each Controle In Controls
                            If IsNumeric(Controle.Tag) Then
                                If CInt(Controle.Tag) = rs!tdi_tgc_cod_grupo Then
                                    Controle.Text = rs!TDI_VALOR_ITEM
                                    Exit For
                                End If
                            End If
                        Next
                        On Error GoTo 0
                    Else
                        For i = 1 To treCadBT.NodesCollection.Count
                            If IsNumeric(Left(treCadBT.NodesCollection(i).Key, 3)) Then
                                If rs!tdi_tgc_cod_grupo = CInt(Mid(treCadBT.NodesCollection(i).Key, 4, 3)) _
                                   And rs!tdi_tco_cod_componente = CInt(Left(treCadBT.NodesCollection(i).Key, 3)) Then
                                    treCadBT.Value(i) = 1
                                End If
                            End If
                        Next
                    End If
                ElseIf CInt(Nvl(rs!tdi_tim_ic_unidade, 0)) > 0 Then 'EDIFICACOES
                    cboUnidades.ListIndex = cboUnidades.ListCount - 1
                    If rs!tdi_tim_ic_unidade <> CInt(Nvl(cboUnidades, 0)) Then
                        ReDim Preserve Edificacoes(1 To rs!tdi_tim_ic_unidade) As Edificacao
                        cboUnidades.AddItem Format(rs!tdi_tim_ic_unidade, "0000")
                    End If
                    Edificacoes(rs!tdi_tim_ic_unidade).Componente(rs!tdi_tgc_cod_grupo) = rs!TDI_VALOR_ITEM
                    If rs!tdi_tgc_cod_grupo >= 100 Then Edificacoes(rs!tdi_tim_ic_unidade).Subjetivo(rs!tdi_tgc_cod_grupo) = True
                End If
                rs.MoveNext
            Loop While Not rs.EOF
            cboUnidades_Click
        End If
    Else
        Avisa "Imóvel não encontrado."
        txtInscImob = ""
        txtCodReduzido = ""
        txtInscImob.Enabled = True
        txtCodReduzido.Enabled = True
        Exit Sub
    End If
    On Error Resume Next
    txtIncAnterior.SetFocus
    On Error GoTo 0
    Bdados.FechaTabela rs
    Screen.MousePointer = 0
    treCadBP.ExpandAll
    treCadBT.ExpandAll
End Sub

Private Sub txtInscMunicipal_LostFocus()
    Dim rs As VSRecordset
    Dim cadastro As New VSImposto
    Dim Sql As String
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtInscMunicipal) <> "" Then
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            If PosPic(txtInscMunicipal, "-") = 0 Then txtInscMunicipal = cadastro.FormataInscricao(txtInscMunicipal, InscContrib)
        End If
        Sql = "Select  tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, " & _
        " tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF,TCI_COD_LOGRADOURO,tci_rg from Tab_Contribuinte where tci_im = '" & txtInscMunicipal & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtContribuinte = "" & rs(0)  'Rs!tci_Nome
            cboLogrProp = CStr("" & rs(1))
            cboNomeLogrProp = "" & rs(2) '!tci_nome_logradouro
            txtNumeroProp = "" & rs(3)  '!tci_numero
            txtComplementoProp = "" & rs(4)  '!tci_complemento
            cboBairroProp = "" & rs(5)  '!tci_bairro
            txtCep = "" & rs!TCI_CEP
            cboMunicipio = rs(7)
            txtUF = rs(8) '!tci_UF
            txtCpfProp = "" & rs!TCI_CGC_CPF
            txtRg = "" & rs!tci_rg
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtInscMunicipal.Enabled = True
            txtInscMunicipal.SetFocus
        End If
    End If
    Bdados.FechaTabela rs
End Sub
