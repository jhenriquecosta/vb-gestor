VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{D8A7CA9C-BFF7-11D5-9D50-00D0590D0C80}#1.0#0"; "CTREEOPT.OCX"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU203a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU203"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   33
      Top             =   7200
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   9285
         TabIndex        =   84
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10470
         TabIndex        =   86
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   8040
         TabIndex        =   82
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin ActiveTabs.SSActiveTabs ssTab 
      Height          =   6465
      Left            =   0
      TabIndex        =   34
      Top             =   690
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   11404
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
      Tabs            =   "TCIU203a.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   6075
         Left            =   30
         TabIndex        =   57
         Top             =   360
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   10716
         _Version        =   131082
         TabGuid         =   "TCIU203a.frx":00DD
         Begin VTOcx.fraVISUAL fraCondominio 
            Height          =   3000
            Left            =   6525
            TabIndex        =   60
            Top             =   60
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   5292
            Altura          =   1905
            Caption         =   " Dados Adicionais da Edificação"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtAreaIrregular 
               Height          =   495
               Left            =   285
               TabIndex        =   63
               Tag             =   "122"
               Top             =   885
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   873
               Caption         =   "Área Edificada Irregular"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VB.Frame Frame1 
               Caption         =   "Fração Ideal"
               Height          =   540
               Left            =   300
               TabIndex        =   97
               Top             =   1395
               Width           =   2070
               Begin VTOcx.txtVISUAL txtFracaoIdeal 
                  Height          =   300
                  Left            =   90
                  TabIndex        =   98
                  Tag             =   "123"
                  Top             =   180
                  Width           =   1905
                  _ExtentX        =   3360
                  _ExtentY        =   529
                  Caption         =   ""
                  Text            =   ""
               End
            End
            Begin VTOcx.txtVISUAL txtPavimentos 
               Height          =   495
               Left            =   285
               TabIndex        =   61
               Tag             =   "110"
               Top             =   315
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   873
               Caption         =   "Pavimentos"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.fraVISUAL fraUnidade 
               Height          =   735
               Left            =   2205
               TabIndex        =   87
               Top             =   2130
               Visible         =   0   'False
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   1296
               Altura          =   1905
               Caption         =   " "
               CorTexto        =   0
               CorFaixa        =   -2147483633
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.cmdVISUAL cmdExcluiEdific 
                  Height          =   375
                  Left            =   195
                  TabIndex        =   90
                  Top             =   360
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   661
                  Caption         =   "&Excluir"
                  Acao            =   2
                  CorBorda        =   8421504
                  CorFrente       =   16384
               End
               Begin VTOcx.cmdVISUAL cmdNovaEdific 
                  Height          =   375
                  Left            =   1305
                  TabIndex        =   89
                  Top             =   360
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   661
                  Caption         =   "&Nova"
                  Acao            =   1
                  CorBorda        =   8421504
                  CorFrente       =   16384
               End
               Begin VTOcx.txtVISUAL txtAlteracao 
                  Height          =   285
                  Left            =   30
                  TabIndex        =   88
                  Tag             =   "111"
                  Top             =   0
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   503
                  Caption         =   "Alterando Unidade"
                  Text            =   ""
                  Enabled         =   0   'False
                  Restricao       =   2
                  AlinhamentoTexto=   1
               End
            End
            Begin VTOcx.cboVISUAL cboUnidades 
               Height          =   315
               Left            =   180
               TabIndex        =   85
               TabStop         =   0   'False
               Top             =   2055
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   556
               Caption         =   "Unidades "
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cmdVISUAL cmdAdicionar 
               Height          =   375
               Left            =   2520
               TabIndex        =   65
               Top             =   1575
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   661
               Caption         =   "&Adicionar Edificação"
               Acao            =   1
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.txtVISUAL txtAnoconstrucao 
               Height          =   495
               Left            =   2490
               TabIndex        =   64
               Tag             =   "111"
               Top             =   885
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   873
               Caption         =   "Ano Construção"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtAreaUnidade 
               Height          =   495
               Left            =   2475
               TabIndex        =   62
               Tag             =   "112"
               Top             =   315
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   873
               Caption         =   "Área Edificada Regular"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
         End
         Begin cTreeOpt.XTreeOpt treCadBP 
            Height          =   5970
            Left            =   60
            TabIndex        =   58
            Tag             =   "2"
            Top             =   75
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   10530
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
            IconSet         =   0
         End
         Begin VTOcx.fraVISUAL fraProprBC 
            Height          =   2190
            Left            =   30
            TabIndex        =   94
            Top             =   3855
            Visible         =   0   'False
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   3863
            Altura          =   1905
            Caption         =   " Dados do Proprietário da Unidade"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtRgCon 
               Height          =   315
               Left            =   4260
               TabIndex        =   81
               Top             =   1800
               Width           =   7065
               _ExtentX        =   12462
               _ExtentY        =   556
               Caption         =   "RG/Órgão Expedidor"
               Text            =   ""
               TipoLetras      =   0
            End
            Begin VTOcx.txtVISUAL txtCpfPropCon 
               Height          =   315
               Left            =   8085
               TabIndex        =   80
               Top             =   1455
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   556
               Caption         =   "CPF/CNPJ"
               Text            =   ""
               Formato         =   1
               Restricao       =   2
            End
            Begin VTOcx.txtVISUAL txtInscMunicipalCon 
               Height          =   315
               Left            =   60
               TabIndex        =   70
               Top             =   330
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   556
               Caption         =   "Inscrição/Cadastro"
               Text            =   ""
               Restricao       =   2
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.cmdVISUAL cmdNovoContribCon 
               Height          =   345
               Left            =   3870
               TabIndex        =   96
               Top             =   300
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   609
               Caption         =   ""
               Acao            =   6
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cmdVISUAL CmdConsultaContribuinteCon 
               Height          =   315
               Left            =   3510
               TabIndex        =   95
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
            Begin VTOcx.txtVISUAL txtUFCon 
               Height          =   315
               Left            =   4935
               TabIndex        =   78
               Top             =   1440
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   556
               Caption         =   "UF"
               Text            =   ""
            End
            Begin VTOcx.cboVISUAL cboMunicipioCon 
               Height          =   315
               Left            =   915
               TabIndex        =   77
               Top             =   1440
               Width           =   4020
               _ExtentX        =   7091
               _ExtentY        =   556
               Caption         =   "Município"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCEPCon 
               Height          =   315
               Left            =   5685
               TabIndex        =   79
               Top             =   1440
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   556
               Caption         =   "CEP"
               Text            =   ""
               Formato         =   4
               Restricao       =   2
            End
            Begin VTOcx.cboVISUAL cboLogrPropCon 
               Height          =   315
               Left            =   720
               TabIndex        =   72
               Top             =   720
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboBairroPropCon 
               Height          =   315
               Left            =   1170
               TabIndex        =   75
               Top             =   1080
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   556
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Editavel        =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtNumeroPropCon 
               Height          =   315
               Left            =   9780
               TabIndex        =   74
               Top             =   720
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               Caption         =   "Número"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtContribuinteCon 
               Height          =   315
               Left            =   4260
               TabIndex        =   71
               Top             =   330
               Width           =   7065
               _ExtentX        =   12462
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               CorFundo        =   16777215
            End
            Begin VTOcx.cboVISUAL cboNomeLogrPropCon 
               Height          =   315
               Left            =   3870
               TabIndex        =   73
               Top             =   720
               Width           =   5745
               _ExtentX        =   10134
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   16777215
               Editavel        =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtComplementoPropCon 
               Height          =   315
               Left            =   6360
               TabIndex        =   76
               Top             =   1080
               Width           =   4965
               _ExtentX        =   8758
               _ExtentY        =   556
               Caption         =   "Complemento"
               Text            =   ""
            End
         End
         Begin VTOcx.fraVISUAL fraCompBC 
            Height          =   735
            Left            =   30
            TabIndex        =   93
            Top             =   3120
            Visible         =   0   'False
            Width           =   11430
            _ExtentX        =   20161
            _ExtentY        =   1296
            Altura          =   1905
            Caption         =   " Dados Adicionais da Localização da Unidade"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtComplementoCon 
               Height          =   315
               Left            =   5550
               TabIndex        =   69
               Top             =   330
               Width           =   5595
               _ExtentX        =   9869
               _ExtentY        =   556
               Caption         =   "Complemento"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtLojaCon 
               Height          =   315
               Left            =   360
               TabIndex        =   66
               Top             =   330
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   556
               Caption         =   "No. Loja/Sala"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtAptoCon 
               Height          =   315
               Left            =   2460
               TabIndex        =   67
               Top             =   330
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               Caption         =   "No. Apto."
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtBlocoCon 
               Height          =   315
               Left            =   4170
               TabIndex        =   68
               Top             =   330
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "Bloco"
               Text            =   ""
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   6075
         Left            =   30
         TabIndex        =   56
         Top             =   360
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   10716
         _Version        =   131082
         TabGuid         =   "TCIU203a.frx":0105
         Begin cTreeOpt.XTreeOpt treCadBT 
            Height          =   5940
            Left            =   0
            TabIndex        =   41
            Tag             =   "1"
            Top             =   90
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   10478
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
            IconSet         =   0
         End
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   5925
            Left            =   6540
            TabIndex        =   59
            Top             =   90
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   10451
            Altura          =   1905
            Caption         =   " Dimensões do Terreno"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtTotalPontos 
               Height          =   495
               Left            =   2880
               TabIndex        =   55
               Tag             =   "114"
               Top             =   5040
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   873
               Caption         =   "Total Pontos"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtDataCadastro 
               Height          =   495
               Left            =   2850
               TabIndex        =   50
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
            End
            Begin VTOcx.txtVISUAL txtAno 
               Height          =   495
               Left            =   2880
               TabIndex        =   51
               Top             =   2640
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   873
               Caption         =   "Ano Aquisição"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtAreaEdifTotal 
               Height          =   495
               Left            =   2880
               TabIndex        =   54
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
            End
            Begin VTOcx.txtVISUAL txtTestadaCampo 
               Height          =   495
               Left            =   150
               TabIndex        =   49
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
            End
            Begin VTOcx.txtVISUAL txtTestada4 
               Height          =   495
               Left            =   120
               TabIndex        =   45
               Tag             =   "105"
               Top             =   2040
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Testada 4(Fundo)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtTestada3 
               Height          =   495
               Left            =   120
               TabIndex        =   44
               Tag             =   "103"
               Top             =   1470
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Testada 3(Esquerda)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtTestada2 
               Height          =   495
               Left            =   120
               TabIndex        =   43
               Tag             =   "101"
               Top             =   900
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Testada 2(Direita)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtProfundidade 
               Height          =   495
               Left            =   2880
               TabIndex        =   52
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
            End
            Begin VTOcx.txtVISUAL txtAreaLote 
               Height          =   495
               Left            =   2880
               TabIndex        =   53
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
            End
            Begin VTOcx.txtVISUAL txtTrechoLogr4 
               Height          =   495
               Left            =   120
               TabIndex        =   48
               Tag             =   "106"
               Top             =   3810
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Trecho/Seção Logr. 4"
               Text            =   ""
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtTrechoLogr3 
               Height          =   495
               Left            =   120
               TabIndex        =   47
               Tag             =   "104"
               Top             =   3210
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Trecho/Seção Logr. 3"
               Text            =   ""
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtTrechoLogr2 
               Height          =   495
               Left            =   120
               TabIndex        =   46
               Tag             =   "102"
               Top             =   2640
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   873
               Caption         =   "Trecho/Seção Logr. 2"
               Text            =   ""
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtTestadaPrin 
               Height          =   495
               Left            =   120
               TabIndex        =   42
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
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6075
         Left            =   30
         TabIndex        =   35
         Top             =   360
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   10716
         _Version        =   131082
         TabGuid         =   "TCIU203a.frx":012D
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1800
            Left            =   0
            TabIndex        =   36
            Top             =   90
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   3175
            Altura          =   1905
            Caption         =   " Localização do Imóvel"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtCodBairro 
               Height          =   315
               Left            =   495
               TabIndex        =   8
               Top             =   1035
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   556
               Caption         =   "Cod.Bairro"
               Text            =   ""
               Restricao       =   2
            End
            Begin VTOcx.cboVISUAL CboLogradouroImovel 
               Height          =   315
               Left            =   4845
               TabIndex        =   5
               Top             =   690
               Width           =   4710
               _ExtentX        =   8308
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.cboVISUAL CboTipoLogradouroImovel 
               Height          =   315
               Left            =   555
               TabIndex        =   4
               Top             =   675
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   556
               Caption         =   "Tipo Logr."
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cmdVISUAL cmdVISUAL1 
               Height          =   315
               Left            =   3450
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   330
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.cboVISUAL cboTipo 
               Height          =   315
               Left            =   9570
               TabIndex        =   3
               ToolTipText     =   "TIPO LOTE"
               Top             =   330
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               Caption         =   "Tipo"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCodReduzido 
               Height          =   315
               Left            =   3825
               TabIndex        =   1
               Top             =   330
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               Caption         =   "Cód. BCI"
               Text            =   ""
               Restricao       =   2
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtInscImob 
               Height          =   315
               Left            =   30
               TabIndex        =   0
               Top             =   360
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   556
               Caption         =   "Insc. Imobiliária"
               Text            =   ""
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtIncAnterior 
               Height          =   315
               Left            =   6390
               TabIndex        =   2
               Top             =   330
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   556
               Caption         =   "Insc. Anterior"
               Text            =   ""
               Restricao       =   2
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCodLogr 
               Height          =   315
               Left            =   2925
               TabIndex        =   6
               Top             =   675
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               Caption         =   "Cod.Logr"
               Text            =   ""
               Restricao       =   2
            End
            Begin VTOcx.txtVISUAL txtLogradouroImovel 
               Height          =   315
               Left            =   2610
               TabIndex        =   37
               Top             =   1905
               Width           =   6525
               _ExtentX        =   11509
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNumero 
               Height          =   315
               Left            =   9810
               TabIndex        =   7
               Top             =   690
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               Caption         =   "Número"
               Text            =   ""
            End
            Begin VTOcx.cboVISUAL cboLoteamento 
               Height          =   315
               Left            =   6270
               TabIndex        =   10
               Top             =   1050
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   556
               Caption         =   "Loteamento"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboBairro 
               Height          =   315
               Left            =   2460
               TabIndex        =   9
               Top             =   1035
               Width           =   3765
               _ExtentX        =   6641
               _ExtentY        =   556
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtQuadra 
               Height          =   315
               Left            =   810
               TabIndex        =   11
               Top             =   1410
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               Caption         =   "Quadra"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtLote 
               Height          =   315
               Left            =   2310
               TabIndex        =   12
               Top             =   1410
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
               Caption         =   "Lote"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtSecao 
               Height          =   315
               Left            =   3600
               TabIndex        =   13
               Top             =   1410
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               Caption         =   "Seção"
               Text            =   ""
            End
            Begin VTOcx.cboVISUAL cboPredio 
               Height          =   315
               Left            =   4950
               TabIndex        =   14
               Top             =   1410
               Width           =   6435
               _ExtentX        =   11351
               _ExtentY        =   556
               Caption         =   "Prédio/Condomínio"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   735
            Left            =   0
            TabIndex        =   38
            Top             =   1920
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   1296
            Altura          =   1905
            Caption         =   " Dados Adicionais da Localização"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtComplemento 
               Height          =   315
               Left            =   5550
               TabIndex        =   18
               Top             =   330
               Width           =   5835
               _ExtentX        =   10292
               _ExtentY        =   556
               Caption         =   "Complemento"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtLoja 
               Height          =   315
               Left            =   360
               TabIndex        =   15
               Top             =   330
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   556
               Caption         =   "No. Loja/Sala"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtApto 
               Height          =   315
               Left            =   2460
               TabIndex        =   16
               Top             =   330
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               Caption         =   "No. Apto."
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtBloco 
               Height          =   315
               Left            =   4170
               TabIndex        =   17
               Top             =   330
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "Bloco"
               Text            =   ""
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   2235
            Left            =   30
            TabIndex        =   39
            Top             =   2670
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   3942
            Altura          =   1905
            Caption         =   " Dados do Proprietário"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VB.CheckBox chkPac 
               Caption         =   "Proj. Moradia Popular?"
               Height          =   195
               Left            =   9000
               TabIndex        =   99
               Top             =   1800
               Width           =   2295
            End
            Begin VTOcx.txtVISUAL txtCpfProp 
               Height          =   315
               Left            =   5730
               TabIndex        =   29
               Top             =   1770
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   556
               Caption         =   "CPF/CNPJ"
               Text            =   ""
               Formato         =   1
               Restricao       =   2
            End
            Begin VTOcx.txtVISUAL txtRg 
               Height          =   315
               Left            =   390
               TabIndex        =   28
               Top             =   1770
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   556
               Caption         =   "RG/Órgão Exp."
               Text            =   ""
               TipoLetras      =   0
            End
            Begin VTOcx.cmdVISUAL cmdNovoContrib 
               Height          =   345
               Left            =   4050
               TabIndex        =   91
               Top             =   300
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   609
               Caption         =   ""
               Acao            =   6
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
               Height          =   315
               Left            =   3690
               TabIndex        =   83
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
               TabIndex        =   32
               Top             =   1410
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               Caption         =   "UF"
               Text            =   ""
               Enabled         =   0   'False
            End
            Begin VTOcx.cboVISUAL cboMunicipio 
               Height          =   315
               Left            =   900
               TabIndex        =   26
               Top             =   1410
               Width           =   6885
               _ExtentX        =   12144
               _ExtentY        =   556
               Caption         =   "Município"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Editavel        =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtCEP 
               Height          =   315
               Left            =   9330
               TabIndex        =   27
               Top             =   1380
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               Caption         =   "CEP"
               Text            =   ""
               Formato         =   4
               Restricao       =   2
            End
            Begin VTOcx.cboVISUAL cboLogrProp 
               Height          =   315
               Left            =   720
               TabIndex        =   21
               Top             =   690
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboBairroProp 
               Height          =   315
               Left            =   1170
               TabIndex        =   24
               Top             =   1050
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   556
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Editavel        =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtNumeroProp 
               Height          =   315
               Left            =   9780
               TabIndex        =   23
               Top             =   690
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               Caption         =   "Número"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtContribuinte 
               Height          =   315
               Left            =   4410
               TabIndex        =   20
               Top             =   330
               Width           =   6945
               _ExtentX        =   12250
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtInscMunicipal 
               Height          =   315
               Left            =   60
               TabIndex        =   19
               Top             =   330
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   556
               Caption         =   "Inscrição/Cadastro"
               Text            =   ""
               Restricao       =   2
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboNomeLogrProp 
               Height          =   315
               Left            =   3945
               TabIndex        =   22
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
               TabIndex        =   25
               Top             =   1050
               Width           =   4965
               _ExtentX        =   8758
               _ExtentY        =   556
               Caption         =   "Complemento"
               Text            =   ""
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   735
            Left            =   30
            TabIndex        =   40
            Top             =   4950
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   1296
            Altura          =   1905
            Caption         =   " Ocupante do Imóvel"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtCPFOcupante 
               Height          =   315
               Left            =   8130
               TabIndex        =   31
               Top             =   330
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   556
               Caption         =   "CPF/CNPJ"
               Text            =   ""
               Formato         =   1
               Restricao       =   2
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtOcupante 
               Height          =   315
               Left            =   1170
               TabIndex        =   30
               Top             =   330
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   556
               Caption         =   "Nome"
               Text            =   ""
            End
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   1138
      Icone           =   "TCIU203a.frx":0155
   End
End
Attribute VB_Name = "TCIU203a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Edificacoes() As Edificacao
Private Tree As New TreeViewBci
Dim BuscaGeralContribuinte As Boolean
Private Limpa_Condominio As Boolean

Rem Raimundo processo que me retorna o código reduzido para cada unidade
Private Type Codigo_Unidade
    Unidade As String
    Codigo  As String
End Type
Dim Unidade(1 To 100) As Codigo_Unidade
Dim Contador As Integer


Public Sub HabilitaCaixa(Status As Boolean)
    txtInscMunicipal.Enabled = Not Status
    txtContribuinte.Enabled = Status
    cboLogrProp.Enabled = Status
    cboNomeLogrProp.Enabled = Status
    txtNumeroProp.Enabled = Status
    txtComplementoProp.Enabled = Status
    cboBairroProp.Enabled = Status
    txtCEP.Enabled = Status
    cboMunicipio.Enabled = Status
    txtUF.Enabled = Status
    txtInscMunicipal = ""
    txtContribuinte = ""
    cboLogrProp = ""
    cboNomeLogrProp = ""
    txtNumeroProp = ""
    txtComplementoProp = ""
    cboBairroProp = ""
    txtCEP = ""
    cboMunicipio = ""
    txtUF = ""
End Sub

Function ImovelJaCadastrado(Inscricao As String)
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim aux As String
       
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        Sql = "Select * from tab_imovel where tim_ic_auxiliar ='" & Inscricao & "'"
    Else
        Sql = "Select * from tab_imovel where tIM_ic ='" & Inscricao & "'"
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        ImovelJaCadastrado = True
    End If
End Function

Private Sub cboBairro_Click()
    txtCodBairro = cboBairro.Coluna(0).Valor
End Sub

Private Sub CboLogradouroImovel_Click()
    txtCodLogr = CboLogradouroImovel.Coluna(0).Valor
End Sub

Private Sub cboMunicipio_Click()
    txtUF = cboMunicipio.Coluna(2).Valor
End Sub

Private Sub cboPredio_Click()
   Dim Sql As String
   Dim Rs  As VSRecordset
   Sql = "Select * from tab_edificio where ted_cod_edificio = '" & cboPredio.Coluna(0).Valor & "'"
   If Bdados.AbreTabela(Sql, Rs) Then
         CboLogradouroImovel.SetarLinha Trim(Rs.Fields("TED_TLG_COD_LOGRADOURO"))
         CboTipoLogradouroImovel.SetarLinha Rs.Fields("TED_TTL_COD_TIPO_LOGRA")
         txtNumero = "" & Rs.Fields("TED_NUMERO")
         cboBairro.SetarLinha Rs.Fields("TED_TBA_COD_BAIRRO")
         cboLoteamento.SetarLinha Rs.Fields("ted_tlo_cod_loteamento")
   End If
End Sub

Private Sub cboTipo_Click()
    If cboTipo.Coluna(1).Valor = 2 Then
        ssTab.Tabs(3).Enabled = False
    Else
        ssTab.Tabs(3).Enabled = True
    End If
End Sub

Private Sub cboUnidades_Click()
    Dim Unidade As Integer
    
    If Limpa_Condominio = True Then Exit Sub
    
    If CInt(Nvl(cboUnidades.Text, 0)) <> 0 Then
        Unidade = CInt(Nvl(cboUnidades.Text, 0))
        txtAnoconstrucao = Edificacoes(CInt(Nvl(cboUnidades.Text, 0))).Componente(txtAnoconstrucao.Tag)
        txtAreaUnidade = Edificacoes(CInt(cboUnidades.Text)).Componente(txtAreaUnidade.Tag)
        txtPavimentos = Nvl(Edificacoes(CInt(cboUnidades.Text)).Componente(txtPavimentos.Tag), 1)
        txtAreaIrregular = Nvl(Edificacoes(CInt(cboUnidades.Text)).Componente(txtAreaIrregular.Tag), 0)
        txtFracaoIdeal = Edificacoes(CInt(cboUnidades.Text)).Componente(txtFracaoIdeal.Tag)
        Tree.SetaTreeViewEdificacao treCadBP, Edificacoes, CInt(cboUnidades.Text)
        
        txtAptoCon = Edificacoes(Unidade).Bc.Endereco.Apto
        txtBlocoCon = Edificacoes(Unidade).Bc.Endereco.Bloco
        txtComplementoCon = Edificacoes(Unidade).Bc.Endereco.Complemento
        
        txtLojaCon = Edificacoes(Unidade).Bc.Endereco.SalaLoja
        txtInscMunicipalCon = Edificacoes(Unidade).Bc.Proprietario.Inscricao
        cboBairroPropCon.Text = Edificacoes(Unidade).Bc.Proprietario.Bairro
        txtCEPCon = Edificacoes(Unidade).Bc.Proprietario.CEP
        cboMunicipioCon.Text = Edificacoes(Unidade).Bc.Proprietario.Cidade
        txtComplementoPropCon = Edificacoes(Unidade).Bc.Proprietario.Complemento
        cboLogrPropCon.Formato = formDocumento
        txtCpfPropCon.Formato = formDocumento
        txtCpfPropCon = Edificacoes(Unidade).Bc.Proprietario.Cpf

        cboLogrPropCon.Text = Edificacoes(Unidade).Bc.Proprietario.Logradouro
        txtContribuinteCon = Edificacoes(Unidade).Bc.Proprietario.Nome
        cboNomeLogrPropCon.Text = Edificacoes(Unidade).Bc.Proprietario.NomeLogradouro
        txtNumeroPropCon = Edificacoes(Unidade).Bc.Proprietario.Numero
        txtRgCon = Edificacoes(Unidade).Bc.Proprietario.Rg
        txtCpfPropCon = Edificacoes(Unidade).Bc.Proprietario.Cpf
        txtUFCon = Edificacoes(Unidade).Bc.Proprietario.UF

        fraUnidade.Visible = True
        txtAlteracao = cboUnidades
    End If
    DoEvents
End Sub

Private Function ValidaUnidadeEdificada() As Boolean
    On Error Resume Next
    Dim i As Integer
    Dim Area As String
    
Rem    Area = 0
Rem    For i = 1 To cboUnidades.ListCount
Rem        Area = CCur(Area) + CCur(Edificacoes(i).Componente(txtAreaUnidade.Tag))
Rem    Next
Rem
Rem    If CCur(Area) > CCur(Nvl(txtAreaEdifTotal, 0)) Then
Rem        Util.Avisa "Total da Área Edificada da(s) unidade(s) não pode ser maior que Área Edificada Total."
Rem        ValidaUnidadeEdificada = False
Rem        ssTab.Tabs(2).Selected = True
Rem        txtAreaEdifTotal.SetFocus
Rem        Screen.MousePointer = 0
Rem    Else
Rem        ValidaUnidadeEdificada = True
Rem    End If
ValidaUnidadeEdificada = True
End Function
Private Sub cmdAdicionar_Click()
    Dim Unidade As Integer
    
    
   ' If (Trim(txtAnoconstrucao)) = "" Or (Trim(txtAreaUnidade)) = "" Then
    '    Avisa "Informe todos os dados da edificacão."
        'txtAreaUnidade.SetFocus
     '   Exit Sub
    'End If
    If fraUnidade.Visible = False Then
        ReDim Preserve Edificacoes(1 To cboUnidades.ListCount + 1) As Edificacao
        cboUnidades.AddItem Format(cboUnidades.ListCount + 1, "000")
        Unidade = UBound(Edificacoes)
    Else
        Unidade = CInt(Nvl(txtAlteracao, 1))
    End If
    Edificacoes(Unidade).Componente(txtAnoconstrucao.Tag) = txtAnoconstrucao
    Edificacoes(Unidade).Componente(txtAreaUnidade.Tag) = txtAreaUnidade
    Edificacoes(Unidade).Componente(txtPavimentos.Tag) = txtPavimentos
    Edificacoes(Unidade).Componente(txtAreaIrregular.Tag) = txtAreaIrregular
    
    Edificacoes(Unidade).Componente(txtFracaoIdeal.Tag) = txtFracaoIdeal
    
    
    Edificacoes(Unidade).Subjetivo(txtPavimentos.Tag) = True
    Edificacoes(Unidade).Subjetivo(txtAnoconstrucao.Tag) = True
    Edificacoes(Unidade).Subjetivo(txtAreaUnidade.Tag) = True
    Edificacoes(Unidade).Subjetivo(txtAreaIrregular.Tag) = True
    
     If cboPredio.ListIndex > -1 Then
       'Dados da Unidade
        Edificacoes(Unidade).Bc.Endereco.Apto = txtAptoCon
        Edificacoes(Unidade).Bc.Endereco.Bloco = txtBlocoCon
        Edificacoes(Unidade).Bc.Endereco.Complemento = txtComplementoCon
        Edificacoes(Unidade).Bc.Endereco.SalaLoja = txtLojaCon
       'Dados do Contribuinte
        Edificacoes(Unidade).Bc.Proprietario.Inscricao = txtInscMunicipalCon
        Edificacoes(Unidade).Bc.Proprietario.Bairro = cboBairroPropCon.Text
        Edificacoes(Unidade).Bc.Proprietario.CEP = txtCEPCon
        Edificacoes(Unidade).Bc.Proprietario.Cidade = cboMunicipioCon.Text
        Edificacoes(Unidade).Bc.Proprietario.Complemento = txtComplementoPropCon
        Edificacoes(Unidade).Bc.Proprietario.Cpf = txtCpfPropCon
        Edificacoes(Unidade).Bc.Proprietario.Logradouro = cboLogrPropCon.Text
        Edificacoes(Unidade).Bc.Proprietario.Nome = txtContribuinteCon
        Edificacoes(Unidade).Bc.Proprietario.NomeLogradouro = cboNomeLogrPropCon.Text
        Edificacoes(Unidade).Bc.Proprietario.Numero = txtNumeroPropCon
        Edificacoes(Unidade).Bc.Proprietario.Rg = txtRgCon
        Edificacoes(Unidade).Bc.Proprietario.UF = txtUFCon
    End If
    Tree.AdicionaEdificacao treCadBP, Edificacoes, Unidade
    cmdNovaEdific_Click
End Sub

Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    ssTab.Tabs(1).Selected = True
    txtInscImob.Enabled = True
    txtCodReduzido.Enabled = True
    Tree.CarregaListaComponentes treCadBT
    Tree.CarregaListaComponentes treCadBP
    Screen.MousePointer = 0
    fraUnidade.Visible = False
    txtInscImob.SetFocus
End Sub

Private Sub CmdConsultaContribuinte_Click()
    BuscaGeralContribuinte = True
    
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscMunicipal
    BuscaGeralContribuinte = False
End Sub

Private Sub CmdConsultaContribuinteCon_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscMunicipalCon
End Sub

Private Sub cmdExcluiEdific_Click()
    Dim i As Integer
    If Confirma("Confirma a exclusão da unidade selecionada?") Then
        Edificacoes(cboUnidades).Deletado = True
        cboUnidades.Clear
        For i = 1 To UBound(Edificacoes)
            If Not Edificacoes(i).Deletado Then cboUnidades.AddItem Format(i, "000")
        Next
        cmdNovaEdific_Click
    End If
End Sub

Private Sub cmdNovaEdific_Click()
    txtAreaUnidade = ""
    txtAnoconstrucao = ""
    txtAreaIrregular = ""
    txtContribuinteCon = ""
    cboLogrPropCon.ListIndex = -1
    cboNomeLogrPropCon = ""
    txtFracaoIdeal = ""
    txtNumeroPropCon = ""
    txtComplementoPropCon = ""
    cboBairroPropCon.ListIndex = -1
    txtPavimentos = ""
    cboUnidades.ListIndex = -1
    cboLogrPropCon.ListIndex = -1
    cboLogrPropCon.Text = ""
    cboBairroPropCon.ListIndex = -1
    cboBairroPropCon.Text = ""
    txtCEPCon = ""
    cboMunicipioCon = ""
    txtUFCon = ""
    txtCpfPropCon = ""
    txtRgCon = ""
    txtInscMunicipalCon = ""
    txtAptoCon = ""
    txtBlocoCon = ""
    txtComplementoCon = ""
    txtLojaCon = ""
    Limpa_Condominio = True
    cboUnidades.ListIndex = cboUnidades.ListCount - 1
    fraUnidade.Visible = False
    Limpa_Condominio = False
    'txtPavimentos.SetFocus
    DoEvents
End Sub

Private Sub cmdNovoContrib_Click()
    Call HabilitaCaixa(True)
    txtOcupante = ""
    txtContribuinte.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Lote As New BCI
    Dim InscricaoMunicipal As String
    Dim InscricaoReduzida As String
    Dim InscricaoCadastral As String
    Screen.MousePointer = 11
    Dim Boletim As TipoBoletim
    Dim Conta As New ContaCorrente
    Dim Insc As String
    Dim Motivo As String
    Dim UnidadeCondominio As Integer
    Dim Contador As Integer
    
    Dim Unidade As Integer
    
    If Not ValidaUnidadeEdificada Then Exit Sub
    'BCP TEMP TESTE
     
     cmdAdicionar_Click
    
    Do
        Motivo = Edita.TiraPic(Edita.TiraPic(Trim(Util.Entrada("Informe o motivo da alteracão", "Justificativa.")), Chr(13)), Chr(13))
    Loop While Trim(Motivo) = ""
    InscricaoMunicipal = txtInscMunicipal
    Lote.CarregaDadosContribuinte InscricaoMunicipal, txtContribuinte, txtCpfProp, "", cboLogrProp, Trim(cboNomeLogrProp), _
             txtNumeroProp, txtComplementoProp, "", Trim(cboBairroProp), txtCEP, cboMunicipio, txtUF, txtRg
    If Not Lote.InsereContribuinte() Then Exit Sub
    InscricaoCadastral = txtInscImob
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        InscricaoReduzida = txtCodReduzido
        Lote.CarregaDadosImovel InscricaoCadastral, txtIncAnterior, "0", "0", "", "", CStr(CboLogradouroImovel.Coluna(0).Valor), CStr(cboBairro.Coluna(0).Valor), _
             txtNumero, txtComplemento, txtLote, txtQuadra, CStr(cboLoteamento.Coluna(0).Valor), CInt(cboTipo.Coluna(1).Valor), txtOcupante, _
             txtCPFOcupante, , , , , , , , , , _
             txtBloco, InscricaoReduzida, txtSecao, Trim(CStr(cboPredio.Coluna(0).Valor)), , txtApto, txtLoja, CInt(Nvl(txtAno, 0)), CStr(CboTipoLogradouroImovel.Coluna(0).Valor), , chkPac.Value
    Else
        Lote.CarregaDadosImovel InscricaoCadastral, txtIncAnterior, "0", "0", "", "", CStr(CboLogradouroImovel.Coluna(0).Valor), CStr(cboBairro.Coluna(0).Valor), _
             txtNumero, txtComplemento, txtLote, txtQuadra, CStr(cboLoteamento.Coluna(0).Valor), CInt(cboTipo.Coluna(1).Valor), txtOcupante, _
             txtCPFOcupante, , , , , , , , , , txtBloco, , txtSecao, _
             CStr(cboPredio.Coluna(0).Valor), Format(Date, "dd/mm/yyyy"), txtApto, txtLoja, CInt(Nvl(txtAno, 0)), CStr(CboTipoLogradouroImovel.Coluna(0).Valor), , chkPac.Value
    End If
    If Not Lote.InsereTerritorio() Then Exit Sub
    
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        'BCP
        Bdados.DeletaDados "TAB_DETALHE_IMOVEL", "TDI_TIM_IC='" & InscricaoReduzida & "'"
        Lote.GravaBoletimTerritorial treCadBT, InscricaoReduzida, 0
        If cboPredio.ListIndex = -1 Then 'QUANDO NAO FOR CONDOMINIO
            If cboUnidades.ListCount > 0 Then Lote.GravaBoletimPredial Edificacoes, InscricaoReduzida
        End If
        If cboUnidades.ListCount > 0 Then Lote.GravaBoletimPredial Edificacoes, InscricaoReduzida
        Lote.GravaComponentes InscricaoReduzida, Me, 100, 110, True, 0, 0
        Lote.GravaComponente InscricaoReduzida, 0, Nvl(txtProfundidade, 0), txtProfundidade.Tag, 0
        Lote.GravaComponente InscricaoReduzida, 0, Nvl(txtAreaEdifTotal, 0), txtAreaEdifTotal.Tag, 0
        Lote.GravaComponente InscricaoReduzida, 0, Nvl(txtTotalPontos, 0), txtTotalPontos.Tag, 0
        Insc = InscricaoReduzida
    Else
        'BCP
        Bdados.DeletaDados "TAB_DETALHE_IMOVEL", "TDI_TIM_IC='" & InscricaoCadastral & "'"
        Lote.GravaBoletimTerritorial treCadBT, InscricaoCadastral, 0
        If cboPredio.ListIndex = -1 Then 'QUANDO NAO FOR CONDOMINIO
            If cboUnidades.ListCount > 0 Then Lote.GravaBoletimPredial Edificacoes, InscricaoCadastral
        End If
        If cboUnidades.ListCount > 0 Then Lote.GravaBoletimPredial Edificacoes, InscricaoCadastral
        Lote.GravaComponentes InscricaoCadastral, Me, 100, 110, True, 0, 0
        Lote.GravaComponente InscricaoCadastral, 0, Nvl(txtProfundidade, 0), txtProfundidade.Tag, 0
        Lote.GravaComponente InscricaoCadastral, 0, Nvl(txtAreaEdifTotal, 0), txtAreaEdifTotal.Tag, 0
        Lote.GravaComponente InscricaoCadastral, 0, Nvl(txtTotalPontos, 0), txtTotalPontos.Tag, 0
        Insc = InscricaoCadastral
    End If
    Bdados.DeletaDados "TAB_IMOVEL", "TIM_IC_CONDOMINIO ='" & Insc & "'"
    'Gravo os dados do condominio...
   If cboPredio.ListIndex <> -1 Then 'QUANDO FOR CONDOMINIO
        Dim UnidadeCond() As Edificacao
        UnidadeCondominio = Right(txtInscImob, 3)
        For Contador = 0 To cboUnidades.ListCount - 1
            UnidadeCondominio = UnidadeCondominio + 1
            Unidade = cboUnidades.List(Contador)
            Lote.CarregaDadosContribuinte Edificacoes(Unidade).Bc.Proprietario.Inscricao, Edificacoes(Unidade).Bc.Proprietario.Nome, Edificacoes(Unidade).Bc.Proprietario.Cpf, "", Edificacoes(Unidade).Bc.Proprietario.Logradouro, Trim(Edificacoes(Unidade).Bc.Proprietario.NomeLogradouro), _
                Edificacoes(Unidade).Bc.Proprietario.Numero, Edificacoes(Unidade).Bc.Proprietario.Complemento, "", Trim(Edificacoes(Unidade).Bc.Proprietario.Bairro), Edificacoes(Unidade).Bc.Proprietario.CEP, Edificacoes(Unidade).Bc.Proprietario.Cidade, Edificacoes(Unidade).Bc.Proprietario.UF, Edificacoes(Unidade).Bc.Proprietario.Rg
            If Not Lote.InsereContribuinte() Then
                Avisa "Erro ao gravar contribuinte da unidade."
                Screen.MousePointer = 0
                Exit Sub
            End If
            InscricaoCadastral = Left(txtInscImob, Len(txtInscImob) - 3) & Format(UnidadeCondominio, "000")
            If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                Rem PONTO PARA TER CUIDADO...
                If Edificacoes(Unidade).Inscricao = "" Then
                    InscricaoReduzida = Conta.GeraCodPagamento("CADASTRO IMOBILIARIO")
                Else
                    InscricaoReduzida = Edificacoes(Unidade).Inscricao
                End If
                Lote.CarregaDadosImovel InscricaoCadastral, "", Format(CStr(UnidadeCondominio), "000"), "0", Insc, "", txtCodLogr, CStr(cboBairro.Coluna(0).Valor), _
                txtNumero, Edificacoes(Unidade).Bc.Endereco.Complemento, txtLote, txtQuadra, CStr(cboLoteamento.Coluna(0).Valor), CInt(cboTipo.Coluna(1).Valor), Edificacoes(Unidade).Bc.Proprietario.Nome, _
                Edificacoes(Unidade).Bc.Proprietario.Cpf, , , , , , , , , , _
                Edificacoes(Unidade).Bc.Endereco.Bloco, InscricaoReduzida, txtSecao, Trim(CStr(cboPredio.Coluna(0).Valor)), Format(Date, "dd/mm/yyyy"), Edificacoes(Unidade).Bc.Endereco.Apto, Edificacoes(Unidade).Bc.Endereco.SalaLoja, CInt(Nvl(txtAno, 0)), , txtFracaoIdeal, chkPac.Value
            Else
                Lote.CarregaDadosImovel InscricaoCadastral, "", Format(CStr(UnidadeCondominio), "000"), "0", Insc, "", txtCodLogr, CStr(cboBairro.Coluna(0).Valor), _
                txtNumero, Edificacoes(Unidade).Bc.Endereco.Complemento, txtLote, txtQuadra, CStr(cboLoteamento.Coluna(0).Valor), CInt(cboTipo.Coluna(1).Valor), Edificacoes(Unidade).Bc.Proprietario.Nome, _
                Edificacoes(Unidade).Bc.Proprietario.Cpf, , , , , , , , , , Edificacoes(Unidade).Bc.Endereco.Bloco, , txtSecao, _
                CStr(cboPredio.Coluna(0).Valor), Format(Date, "dd/mm/yyyy"), Edificacoes(Unidade).Bc.Endereco.Apto, Edificacoes(Unidade).Bc.Endereco.SalaLoja, CInt(Nvl(txtAno, 0)), , txtFracaoIdeal, chkPac.Value
            End If
            
            If Not Lote.InsereTerritorio() Then
                Avisa "Erro ao gravar unidade."
                Screen.MousePointer = 0
                Exit Sub
            Else
                ReDim UnidadeCond(Unidade To Unidade) As Edificacao
                UnidadeCond(Unidade) = Edificacoes(Unidade)
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
                    Lote.GravaBoletimPredial UnidadeCond, InscricaoReduzida
                Else
                    Lote.GravaBoletimPredial UnidadeCond, InscricaoCadastral
                End If
            End If
        Next
    End If
    Lote.GravaHistorico Insc, Motivo
    Avisa "Dados gravados com sucesso. Registro No " & Insc
    chkPac = 0
    cmdCancela_Click
End Sub

Private Sub cmdVISUAL1_Click()
    LimpaCampos Me
    'AplicacoesVTFuncoes.BuscaInscricao InscImovel , txtInscImob,
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtCodReduzido.Enabled = True
        AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtCodReduzido
        If Trim(txtCodReduzido) <> "" Then
            txtCodReduzido_LostFocus
        Else
            txtCodReduzido.Enabled = False
        End If
    Else
        txtInscImob.Enabled = True
        AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtInscImob
        If Trim(txtInscImob) <> "" Then
            txtInscImob_LostFocus
        Else
            txtInscImob.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Activate()
    If Me.Tag <> "" Then
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            txtCodReduzido = Mid(Trim(Me.Tag), 2)
            txtCodReduzido_LostFocus
        Else
            txtInscImob = Mid(Trim(Me.Tag), 2)
            Call txtInscImob_LostFocus
        End If
        cmdSalvar.Enabled = False
        cmdCancela.Enabled = False
        cmdAdicionar.Enabled = False
        cmdExcluiEdific.Enabled = False
        cmdNovaEdific.Enabled = False
        
    End If
    
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 117 Then
        ssTab.Tabs(1).Selected = True
        txtInscImob.SetFocus
    ElseIf KeyCode = 118 Then
        ssTab.Tabs(2).Selected = True
        treCadBT.SetFocus
    ElseIf KeyCode = 119 Then
        ssTab.Tabs(3).Selected = True
        txtPavimentos.SetFocus
    End If
    ssTab_Click
End Sub

Private Sub Form_Load()
    Tree.CarregaListaComponentes treCadBT
    Tree.CarregaListaComponentes treCadBP
    cboLoteamento.Preencher Bdados, "SELECT TLO_COD_LOTEAMENTO,TLO_DESCRICAO FROM TAB_LOTEAMENTO ORDER BY 2", 1
    cboPredio.Preencher Bdados, "SELECT TED_COD_EDIFICIO,TED_DESCRICAO FROM TAB_EDIFICIO ORDER BY 2", 1
    cboBairro.Preencher Bdados, "SELECT TBA_COD_BAIRRO,TBA_NOME FROM TAB_BAIRRO ORDER BY 2", 1
    cboBairroProp.Preencher Bdados, "SELECT TBA_COD_BAIRRO,TBA_NOME FROM TAB_BAIRRO ORDER BY 2", 1
    cboMunicipio.Preencher Bdados, "SELECT TMU_COD_MUNICIPIO,TMU_NOME,tuf_uf FROM TAB_MUNICIPIO,TAB_UF " & _
        "WHERE TMU_TUF_COD_UF = TUF_COD_UF ORDER  BY 2", 1
    cboLogrProp.Preencher Bdados, "SELECT TTL_COD_TIP_LOGR,TTL_NOME FROM TAB_TIPO_LOGR ORDER BY 2", 1
    cboLogrPropCon.Preencher Bdados, "SELECT TTL_COD_TIP_LOGR,TTL_NOME FROM TAB_TIPO_LOGR ORDER BY 2", 1
    cboNomeLogrProp.Preencher Bdados, "SELECT DISTINCT tlg_nome FROM TAB_LOGRADOURO ORDER BY 1"
    
   ' cabVISUAL1.Exibir Bdados, Me.Name, App.Path
   ' rodVISUAL1.Exibir Bdados, Me.Tag
    
    cboTipo.PreencherGeral Bdados, "TIPO LOTE"
    CboLogradouroImovel.Preencher Bdados, "SELECT  TLG_COD_LOGRADOURO,   TLG_NOME FROM TAB_LOGRADOURO ORDER BY 1", 1
    CboTipoLogradouroImovel.Preencher Bdados, "SELECT TTL_COD_TIP_LOGR,TTL_NOME FROM TAB_TIPO_LOGR ORDER BY 2", 1
    'BCP
    ReDim Edificacoes(1 To 2) As Edificacao
    
    'OLD
    'ReDim Edificacoes(1 To 1) As Edificacao
    
    BuscaGeralContribuinte = False
    If UCase(AplicacoesVTFuncoes.municipio) = "PETROLINA" Then
        txtInscMunicipal.Caption = "Inscricão/Cadastro"
    Else
        txtInscMunicipal.Caption = "Cadastro Único"
        txtInscMunicipal.Width = txtInscMunicipal.Width - 400
        txtInscMunicipal.Left = cboLogrProp.Left - 300
    End If
    fraUnidade.Visible = False
End Sub

Private Sub ssTab_Click()
     If ssTab.Tabs(3).Selected Then
        If cboPredio.ListIndex <> -1 Or cboPredio.Text <> "" Then
            fraCompBC.Visible = True
            fraProprBC.Visible = True
            treCadBP.Height = 3000
            fraCondominio.Height = treCadBP.Height
        Else
            treCadBP.Height = 5955
            fraCondominio.Height = treCadBP.Height
            fraCompBC.Visible = False
            fraProprBC.Visible = False
        End If
    End If
End Sub

Private Sub treCadBP_CheckClick(ItemNode As ComctlLib.INode, Value As cTreeOpt.OptionTreeCheckTypes)
    Tree.MarcaUnico treCadBP, ItemNode.Key, CInt(Value)
End Sub

Private Sub treCadBT_CheckClick(ItemNode As ComctlLib.INode, Value As cTreeOpt.OptionTreeCheckTypes)
    Tree.MarcaUnico treCadBT, ItemNode.Key, CInt(Value)
End Sub

Private Sub txtAreaEdifTotal_LostFocus()
    On Error Resume Next
    If txtAreaEdifTotal <> "" Then
        ssTab.Tabs(3).Selected = True
    End If
    CalculaFracao
End Sub

Private Sub txtAreaLote_LostFocus()
    CalculaProfundidade
End Sub

Private Sub txtAreaUnidade_LostFocus()
    CalculaFracao
End Sub

Private Sub txtCodBairro_LostFocus()
    If txtCodBairro = "" Then Exit Sub
    cboBairro.SetarLinha txtCodBairro
    If cboBairro.Text = "" Then
        Avisa "Bairro não cadastrado."
        txtCodBairro.SetFocus
    End If
End Sub

Private Sub txtCodLogr_LostFocus()
    If txtCodLogr = "" Then Exit Sub
    CboLogradouroImovel.SetarLinha txtCodLogr
    If CboLogradouroImovel.Text = "" Then
        Avisa "Código de logradouro não dadastrado."
        txtCodLogr.SetFocus
    End If
'    On Error GoTo TrataErro
'    Dim Query As String
'    Dim rs As VSRecordset
'    If Trim(txtCodLogr) = "" Then Exit Sub
'    Query = "SELECT TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome, " & _
'        " TAB_BAIRRO.TBA_NOME FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
'        " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
'         " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & txtCodLogr & "' and tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
'    If Bdados.AbreTabela(Query, rs) Then
'        txtLogradouroImovel = rs(0) & " " & rs(1)
'        CboLogradouroImovel.SetarLinha txtCodLogr
'    Else
'        Avisa "Código de logradouro inválido."
'        txtCodLogr.SetFocus
'    End If
'    Bdados.FechaTabela rs
'    Exit Sub
'TrataErro:
'    If Err.Number = 3265 Then
'        Resume Next
'    Else
'        Util.Erro Err.Description
'    End If
End Sub


Private Sub txtCodReduzido_LostFocus()
    txtInscImob_LostFocus
End Sub

Private Sub txtCpfProp_Change()
    txtCpfProp.Formato = formNenhum
End Sub

Private Sub txtCpfProp_LostFocus()
    If Trim(txtCpfProp) = "" Then Exit Sub
    If Len(Edita.TiraTudo(Trim(txtCpfProp))) = 11 Then
        txtCpfProp.Formato = formCPF
    ElseIf Len(Edita.TiraTudo(Trim(txtCpfProp))) = 14 Then
        txtCpfProp.Formato = formCGC
    End If
End Sub

Private Sub txtFracaoIdeal_LostFocus()
    txtFracaoIdeal = Format(txtFracaoIdeal, "###,###,###,##0.000")
End Sub

Private Sub txtInscImob_LostFocus()
    On Error Resume Next
    Dim Rs As VSRecordset
    Dim rsEdificacao As VSRecordset
    Dim CodImovel As String
    Dim Sql As String
    Dim i As Integer
    ReDim Edificacoes(1 To 1) As Edificacao
    
    If BuscaGeralContribuinte Then Exit Sub
    If (Trim(txtInscImob) = "" And (txtCodReduzido) = "") Or (txtCodReduzido.Enabled = False Or txtInscImob.Enabled = False) Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    txtCpfProp.Formato = formDocumento
    If Trim(txtCodReduzido) <> "" Then
        txtInscImob.Enabled = True
        txtCodReduzido.Enabled = False
    Else
        txtInscImob.Enabled = False
        txtCodReduzido.Enabled = True
    End If
    
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        If txtCodReduzido = "" Then
            Sql = "Select * from tab_imovel where TIM_IC_AUXILIAR ='" & txtInscImob & "'"
        Else
            Sql = "Select * from tab_imovel where TIM_IC ='" & txtCodReduzido & "'"
        End If
    Else
        If Trim(txtCodReduzido) = "" And Trim(txtInscImob) <> "" Then
            Sql = "Select * from tab_imovel where TIM_IC ='" & txtInscImob & "'"
        ElseIf Trim(txtCodReduzido) <> "" Then
            Sql = "Select * from tab_imovel where TIM_IC_AUXILIAR ='" & txtCodReduzido & "'"
        End If
    End If
    txtCodReduzido = Trim(txtCodReduzido)
    txtCPFOcupante.Formato = formDocumento
    If Bdados.AbreTabela(Sql, Rs) Then
        txtIncAnterior = "" & IIf(Rs!tim_ic_anterior = 0, "", Rs!tim_ic_anterior)
        'CodImovel = "" & rs.Fields("tci_im")
        CboLogradouroImovel.SetarLinha "" & Rs!tim_tlg_cod_logradouro
        CboTipoLogradouroImovel.SetarLinha "" & Rs.Fields("TIM_TTL_COD_TIP_LOGR")
        txtNumero = "" & Rs!tim_numero
        cboPredio.SetarLinha Nvl("" & Rs!tim_ted_cod_edificio, 0)
        txtAno = "" & Rs!tim_ano_aquis
        cboLoteamento.SetarLinha "" & Rs!tim_loteamento, 0
        cboPredio.SetarLinha "" & Rs!tim_ted_cod_edificio, 0
        cboBairro.SetarLinha "" & Rs!tim_TBA_COD_BAIRRO, 0
        txtQuadra = "" & Rs!tim_QUADRA
        txtLote = "" & Rs!tim_lote
        txtOcupante = "" & Rs!tim_ocupante
        txtCPFOcupante.Formato = formDocumento
        txtCPFOcupante = "" & Rs!tim_cgc_cpf_ocupante
        cboTipo.SetarLinha Nvl("" & Rs!tim_tipo_imovel, 0), 1
        cboTipo_Click
        txtIncAnterior = "" & Rs!tim_ic_anterior
        txtSecao = "" & Rs!tim_secao
        
        txtComplemento = "" & Rs!tim_complemento
        txtBloco = "" & Rs!TIM_BLOCO
        txtApto = "" & Rs!TIM_APTO
        txtLoja = "" & Rs!TIM_SALA_LOJA
        chkPac = IIf(Rs!tim_Pac = True, 1, 0)

        If txtCodReduzido <> "" Then
            txtInscImob = Trim("" & Rs!tim_ic_auxiliar)
        End If
        txtDataCadastro = Format("" & Rs!tim_DATA_CADASTRO, "dd/MM/yyyy")
        'VOU PEGAR O CONTRIBUINTE
        txtInscMunicipal = "" & Rs!tim_tci_im
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            txtCodReduzido = "" & Rs!TIM_IC
            txtInscImob = "" & Rs!tim_ic_auxiliar
        Else
            txtCodReduzido = "" & Rs!tim_ic_auxiliar
            txtInscImob = "" & Rs!TIM_IC
        End If
        
        
        txtInscMunicipal_LostFocus
        'VOU PEGAR OS DETALHES
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            Sql = "Select tim_ic from tab_imovel where tim_ic ='" & txtCodReduzido & "' or tim_ic_condominio ='" & txtCodReduzido & "'"
        Else
            If Trim(txtInscImob) = "" Then
                Avisa "Lote sem inscricao imobiliária. Defina uma para prosseguir."
                Exit Sub
            End If
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
        Sql = "Select TAB_DETALHE_IMOVEL.*,TAB_GRUPO_COMPONENTE_AVANCADO.TGC_CATEGORIA from TAB_DETALHE_IMOVEL,TAB_GRUPO_COMPONENTE_AVANCADO where TDI_TGC_COD_GRUPO = TGC_COD_GRUPO AND TDI_TIM_IC in (" & Trim(Inscricoes) & ") order by tdi_tim_ic_unidade,tdi_tgc_cod_grupo asc"
        If Bdados.AbreTabela(Sql, Rs) Then
            cboUnidades.Clear
            Do
                If CInt(Nvl(Rs!TGC_CATEGORIA, 0)) = 1 Or CInt(Nvl(Rs!TGC_CATEGORIA, 0)) = 3 Then  'LOTE
                    If Rs!tdi_tgc_cod_grupo >= 100 Then
                        Dim Controle As Control
                        On Error Resume Next
                        For Each Controle In Controls
                            If IsNumeric(Controle.Tag) Then
                                If CInt(Controle.Tag) = Rs!tdi_tgc_cod_grupo Then
                                    Controle.Text = "" & Rs!TDI_VALOR_ITEM
                                    Exit For
                                End If
                            End If
                        Next
                        On Error GoTo 0
                        
                    Else
                        For i = 1 To treCadBT.NodesCollection.Count
                            If IsNumeric(Left(treCadBT.NodesCollection(i).Key, 3)) Then
                                If Rs!tdi_tgc_cod_grupo = CInt(Mid(treCadBT.NodesCollection(i).Key, 4, 3)) _
                                   And Rs!tdi_tco_cod_componente = CInt(Left(treCadBT.NodesCollection(i).Key, 3)) Then
                                    treCadBT.Value(i) = 1
                                End If
                            End If
                        Next
                    End If
                    If Rs!tdi_tgc_cod_grupo >= 100 And Rs!tdi_tim_ic_unidade > 0 Then
                        cboUnidades.ListIndex = cboUnidades.ListCount - 1
                        'Edificacoes(rs!tdi_tim_ic_unidade).Componente(rs!tdi_tgc_cod_grupo) = rs!TDI_VALOR_ITEM
                        'Edificacoes(rs!tdi_tim_ic_unidade).Subjetivo(rs!tdi_tgc_cod_grupo) = True
                        If Rs!tdi_tim_ic_unidade <> CInt(Nvl(cboUnidades, 0)) Then
                            ReDim Preserve Edificacoes(1 To Rs!tdi_tim_ic_unidade) As Edificacao
                            cboUnidades.AddItem Format(Rs!tdi_tim_ic_unidade, "000")
                            Edificacoes(Rs!tdi_tim_ic_unidade).Inscricao = Trim(Rs!tdi_tim_ic)
                            Pega_Dados_Unidade txtInscImob, Rs!tdi_tim_ic_unidade
                        End If
                    End If
                ElseIf CInt(Nvl(Rs!TGC_CATEGORIA, 0)) >= 2 And Rs!tdi_tim_ic_unidade > 0 Then 'EDIFICACOES
                    cboUnidades.ListIndex = cboUnidades.ListCount - 1
                    If Rs!tdi_tim_ic_unidade <> CInt(Nvl(cboUnidades, 0)) Then
                        ReDim Preserve Edificacoes(1 To Rs!tdi_tim_ic_unidade) As Edificacao
                        cboUnidades.AddItem Format(Rs!tdi_tim_ic_unidade, "000")
                        Edificacoes(Rs!tdi_tim_ic_unidade).Inscricao = Trim(Rs!tdi_tim_ic)
                        Pega_Dados_Unidade txtInscImob, Rs!tdi_tim_ic_unidade
                    End If
                    Edificacoes(Rs!tdi_tim_ic_unidade).Componente(Rs!tdi_tgc_cod_grupo) = Rs!TDI_VALOR_ITEM
                    If Rs!tdi_tgc_cod_grupo >= 100 Then Edificacoes(Rs!tdi_tim_ic_unidade).Subjetivo(Rs!tdi_tgc_cod_grupo) = True
                    
                End If
                Rs.MoveNext
            Loop While Not Rs.EOF
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
    cboBairro_Click
    CboLogradouroImovel_Click
'    CalculaFracao
    On Error Resume Next
    txtIncAnterior.SetFocus
    On Error GoTo 0
    Bdados.FechaTabela Rs
    Screen.MousePointer = 0
    treCadBP.ExpandAll
    treCadBT.ExpandAll
End Sub

Private Sub CalculaFracao()
    On Error Resume Next
Rem    txtFracaoIdeal = Nvl(txtAreaUnidade, 0) / Nvl(txtAreaEdifTotal, 0)
End Sub
Private Sub Pega_Dados_Unidade(Imovel As String, Unidade As Integer)
    Dim rsC As VSRecordset
    Dim RsI As VSRecordset
    Dim Sql As String
    Dim rsImovel As VSRecordset
    If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        If Not AplicacoesVTFuncoes.municipio = "VERDEJANTE" Then
            Sql = "Select * from tab_imovel where tim_ic_auxiliar = '" & Left(Trim(Imovel), Len(Imovel) - 3) & Format(Right(Trim(Imovel), 3) + Unidade, "000") & "'"
        Else
            Sql = "Select * from tab_imovel where tim_ic = '" & Imovel & "'"
        End If
    Else
        Sql = "Select * from tab_imovel where tim_ic_auxiliar = '" & Imovel & "'"
    End If
    If Bdados.AbreTabela(Sql, rsImovel) Then
        'PEGO OS DADOS DO CONTRIBUINTE...
        Sql = "Select * from tab_contribuinte where tci_im = '" & rsImovel.Fields("tim_tci_im") & "'"
        If Bdados.AbreTabela(Sql, rsC) Then
            
            Edificacoes(Unidade).Bc.Endereco.Apto = "" & rsImovel.Fields("TIM_APTO")
            Edificacoes(Unidade).Bc.Endereco.Bloco = "" & rsImovel.Fields("TIM_BLOCO")
            Edificacoes(Unidade).Bc.Endereco.Complemento = "" & rsImovel.Fields("TIM_COMPLEMENTO")
            Edificacoes(Unidade).Bc.Endereco.SalaLoja = "" & rsImovel.Fields("TIM_SALA_LOJA")
        
            Edificacoes(Unidade).Bc.Proprietario.Inscricao = "" & rsImovel.Fields("tim_tci_im")
            Edificacoes(Unidade).Bc.Proprietario.Bairro = "" & rsC.Fields("TCI_BAIRRO")
            Edificacoes(Unidade).Bc.Proprietario.CEP = "" & rsC.Fields("tci_cep")
            Edificacoes(Unidade).Bc.Proprietario.Cidade = "" & rsC.Fields("tci_cidade")
            Edificacoes(Unidade).Bc.Proprietario.Complemento = "" & rsC.Fields("TCI_COMPLEMENTO")
            Edificacoes(Unidade).Bc.Proprietario.Cpf = "" & rsC.Fields("tci_cgc_cpf")
            Edificacoes(Unidade).Bc.Proprietario.Logradouro = "" & rsC.Fields("TCI_LOGRADOURO")
            Edificacoes(Unidade).Bc.Proprietario.Nome = "" & rsC.Fields("tci_nome")
            Edificacoes(Unidade).Bc.Proprietario.NomeLogradouro = "" & rsC.Fields("tci_nome_logradouro")
            Edificacoes(Unidade).Bc.Proprietario.Numero = "" & rsC.Fields("tci_numero")
            Edificacoes(Unidade).Bc.Proprietario.Rg = "" & rsC.Fields("tci_rg")
            Edificacoes(Unidade).Bc.Proprietario.UF = "" & rsC.Fields("tci_uf")
        End If
    End If
End Sub

Private Sub txtInscMunicipal_LostFocus()
    Dim Rs As VSRecordset
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
        If Bdados.AbreTabela(Sql, Rs) Then
            txtContribuinte = "" & Rs(0)  'Rs!tci_Nome
            cboLogrProp = CStr("" & Rs(1))
            cboNomeLogrProp = "" & Rs(2) '!tci_nome_logradouro
            txtNumeroProp = "" & Rs(3)  '!tci_numero
            txtComplementoProp = "" & Rs(4)  '!tci_complemento
            cboBairroProp = "" & Rs(5)  '!tci_bairro
            txtCEP = "" & Rs!TCI_CEP
            cboMunicipio = "" & Rs(7)
            txtUF = "" & Rs!tci_UF
            txtCpfProp.Formato = formDocumento
            
            txtCpfProp = "" & Rs!TCI_CGC_CPF
            txtRg = "" & Rs!tci_rg
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtInscMunicipal.Enabled = True
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtInscMunicipalCon_LostFocus()
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim cadastro As New VSImposto
    
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtInscMunicipalCon) <> "" Then
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            txtInscMunicipalCon = cadastro.FormataInscricao(txtInscMunicipalCon, InscContrib)
        End If
        Sql = "Select  tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, " & _
        " tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF,TCI_COD_LOGRADOURO,tci_rg from Tab_Contribuinte where tci_im = '" & txtInscMunicipalCon & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            txtContribuinteCon = "" & Rs(0)  'Rs!tci_Nome
            cboLogrPropCon = CStr("" & Rs(1))
            cboNomeLogrPropCon = "" & Rs(2) '!tci_nome_logradouro
            txtNumeroPropCon = "" & Rs(3)  '!tci_numero
            txtComplementoPropCon = "" & Rs(4)  '!tci_complemento
            cboBairroPropCon = "" & Rs(5)  '!tci_bairro
            txtCEPCon = "" & Rs!TCI_CEP
            cboMunicipioCon = "" & Rs(7)
            txtUFCon = "" & Rs(8)
            txtCpfPropCon.Formato = formDocumento
            txtCpfPropCon = "" & Rs!TCI_CGC_CPF
            txtRgCon = "" & Rs!tci_rg
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtInscMunicipalCon.Enabled = True
            txtInscMunicipalCon.SetFocus
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtProfundidade_LostFocus()
    If Not (Trim(txtAreaLote) <> "" Or UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA") Then
        If Trim(txtTestadaPrin) <> "" And Trim(txtProfundidade) <> "" Then
            txtAreaLote = CDbl(txtTestadaPrin) * CDbl(txtProfundidade)
        End If
    End If
End Sub
Private Sub CalculaProfundidade()
    If Nvl(txtTestadaPrin, 0) = 0 Then Exit Sub
    If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
        txtProfundidade = (Nvl(txtAreaLote, 0) * Nvl(txtTestadaPrin, 0) / 30) ^ 0.5
    Else
        Exit Sub
        txtProfundidade = (Nvl(txtAreaLote, 0) / Nvl(txtTestadaPrin, 0))
    End If
End Sub
Private Sub txtTestadaPrin_LostFocus()
    If Not (Trim(txtAreaLote) <> "" Or UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA") Then
        If Trim(txtTestadaPrin) <> "" And Trim(txtProfundidade) <> "" Then
            txtAreaLote = CDbl(txtTestadaPrin) * CDbl(txtProfundidade)
        End If
    End If
    CalculaProfundidade
End Sub

Private Sub txtVISUAL1_Change()
    
End Sub

Private Sub txtVISUAL1_LostFocus()

End Sub
