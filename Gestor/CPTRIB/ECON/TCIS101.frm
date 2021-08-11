VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIS101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIS"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   ControlBox      =   0   'False
   Icon            =   "TCIS101.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs tabCadastro 
      Height          =   7400
      Left            =   60
      TabIndex        =   51
      Top             =   690
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   13044
      _Version        =   131082
      TabCount        =   8
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      Tabs            =   "TCIS101.frx":08CA
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel8 
         Height          =   7005
         Left            =   30
         TabIndex        =   145
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   12356
         _Version        =   131082
         TabGuid         =   "TCIS101.frx":0AA2
         Begin VB.TextBox txtObservacaoCompleta 
            Height          =   5415
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   146
            Text            =   "TCIS101.frx":0ACA
            Top             =   720
            Width           =   9735
         End
         Begin VTOcx.txtVISUAL txtAreaEstabelecimento 
            Height          =   480
            Left            =   120
            TabIndex        =   147
            Top             =   120
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   847
            Caption         =   "Area do Estabelecimento"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoRotulo=   1
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   7005
         Left            =   30
         TabIndex        =   120
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   12356
         _Version        =   131082
         TabGuid         =   "TCIS101.frx":0AD5
         Begin VTOcx.grdVISUAL grdAtividade 
            Height          =   4620
            Left            =   105
            TabIndex        =   121
            Top             =   1545
            Width           =   9645
            _ExtentX        =   17013
            _ExtentY        =   8149
            CorBorda        =   16711680
            Caption         =   "Atividades"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
         Begin VTOcx.cboVISUAL cboGrupoAtividade 
            Height          =   315
            Left            =   195
            TabIndex        =   122
            Tag             =   "Grupo"
            Top             =   405
            Width           =   9525
            _ExtentX        =   16801
            _ExtentY        =   556
            Caption         =   "Grupo de Atividade"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL CboAtividade 
            Height          =   315
            Left            =   1050
            TabIndex        =   123
            Tag             =   "Grupo"
            Top             =   750
            Width           =   8670
            _ExtentX        =   15293
            _ExtentY        =   556
            Caption         =   "Atividade"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdAdAtividade 
            Height          =   375
            Left            =   7215
            TabIndex        =   124
            Top             =   1110
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
            Caption         =   "&Adicionar "
            Acao            =   3
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdExcluir 
            Height          =   375
            Left            =   8490
            TabIndex        =   125
            Top             =   1110
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
            Caption         =   "&Excluir"
            Acao            =   2
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   7005
         Left            =   30
         TabIndex        =   52
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   12356
         _Version        =   131082
         TabGuid         =   "TCIS101.frx":0AFD
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   1350
            Left            =   60
            TabIndex        =   53
            Top             =   75
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   2381
            Altura          =   1905
            Caption         =   " Contribuinte"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtIncriAuxiliar 
               Height          =   480
               Left            =   5385
               TabIndex        =   5
               Top             =   795
               Width           =   1300
               _ExtentX        =   2302
               _ExtentY        =   847
               Caption         =   "Inscri.Anterior"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtRG 
               Height          =   480
               Left            =   90
               TabIndex        =   0
               Top             =   315
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   847
               Caption         =   "RG"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtFantasia 
               Height          =   480
               Left            =   90
               TabIndex        =   4
               Top             =   795
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   847
               Caption         =   "Nome Fantasia"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtFoneFax 
               Height          =   480
               Left            =   6700
               TabIndex        =   6
               Top             =   795
               Width           =   1360
               _ExtentX        =   2408
               _ExtentY        =   847
               Caption         =   "Fone/Fax"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCateg 
               Height          =   480
               Left            =   7335
               TabIndex        =   9
               Top             =   1380
               Visible         =   0   'False
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   847
               Caption         =   "Categoria"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtProtocolo 
               Height          =   480
               Left            =   90
               TabIndex        =   1
               Top             =   315
               Visible         =   0   'False
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   847
               Caption         =   "Nº Protocolo"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCnh 
               Height          =   480
               Left            =   6090
               TabIndex        =   8
               Top             =   1410
               Visible         =   0   'False
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   847
               Caption         =   "CNH"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtAutoriza 
               Height          =   480
               Left            =   6375
               TabIndex        =   10
               Top             =   -525
               Visible         =   0   'False
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   847
               Caption         =   "Nº Autorização"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtRuc 
               Height          =   480
               Left            =   8200
               TabIndex        =   7
               Top             =   795
               Width           =   1400
               _ExtentX        =   2461
               _ExtentY        =   847
               Caption         =   "Ins. Estadual"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtRazao 
               Height          =   480
               Left            =   3465
               TabIndex        =   3
               Tag             =   "Nome ou Razão Social"
               Top             =   315
               Width           =   6150
               _ExtentX        =   10848
               _ExtentY        =   847
               Caption         =   "Nome ou Razão Social"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCgc 
               Height          =   480
               Left            =   1620
               TabIndex        =   2
               Top             =   315
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   847
               Caption         =   "CPF ou CNPJ"
               Text            =   ""
               Formato         =   10
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1425
            Left            =   60
            TabIndex        =   54
            Top             =   1455
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   2514
            Altura          =   1905
            Caption         =   " Localização"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL cmdVISUAL1 
               Height          =   285
               Left            =   4140
               TabIndex        =   142
               TabStop         =   0   'False
               Top             =   525
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtIc 
               Height          =   480
               Left            =   2745
               TabIndex        =   13
               Top             =   330
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   847
               Caption         =   "Insc. Cadastral"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtNum 
               Height          =   480
               Left            =   9015
               TabIndex        =   16
               Top             =   330
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   847
               Caption         =   "Nº"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCep 
               Height          =   480
               Left            =   8475
               TabIndex        =   21
               Tag             =   "CEP"
               Top             =   855
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   847
               Caption         =   "CEP"
               Text            =   ""
               Formato         =   4
               Restricao       =   2
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboImovel 
               Height          =   510
               Left            =   1260
               TabIndex        =   12
               Tag             =   "Imóvel"
               Top             =   330
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   900
               Caption         =   "Imóvel"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboLogr 
               Height          =   510
               Left            =   5805
               TabIndex        =   15
               Top             =   330
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   900
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Editavel        =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtCidade 
               Height          =   480
               Left            =   5460
               TabIndex        =   19
               Tag             =   "Cidade"
               Top             =   855
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   847
               Caption         =   "Cidade"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboTipoLogr 
               Height          =   510
               Left            =   4500
               TabIndex        =   14
               Top             =   330
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   900
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboEstabelece 
               Height          =   510
               Left            =   60
               TabIndex        =   11
               Tag             =   "Estabelecido"
               Top             =   330
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   900
               Caption         =   "Estabelecido"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboUF 
               Height          =   510
               Left            =   7665
               TabIndex        =   20
               Tag             =   "UF"
               Top             =   840
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   900
               Caption         =   "UF"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboBairro 
               Height          =   510
               Left            =   2715
               TabIndex        =   18
               Top             =   840
               Width           =   2745
               _ExtentX        =   4842
               _ExtentY        =   900
               Caption         =   "Distrito ou Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Editavel        =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtComplemento 
               Height          =   480
               Left            =   60
               TabIndex        =   17
               Top             =   855
               Width           =   2670
               _ExtentX        =   4710
               _ExtentY        =   847
               Caption         =   "Complemento"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   3550
            Left            =   60
            TabIndex        =   55
            Top             =   2910
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   6271
            Altura          =   1905
            Caption         =   " Atividade"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtCodAtividade 
               Height          =   480
               Left            =   50
               TabIndex        =   29
               Top             =   2200
               Width           =   980
               _ExtentX        =   1720
               _ExtentY        =   847
               Caption         =   "Codigo"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboSitAlvara 
               Height          =   315
               Left            =   105
               TabIndex        =   32
               Top             =   3180
               Width           =   3800
               _ExtentX        =   6694
               _ExtentY        =   556
               Caption         =   "Situação do Alvará"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtInicioPrestacaoServico 
               Height          =   300
               Left            =   6240
               TabIndex        =   34
               Top             =   3195
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   529
               Caption         =   "Início Prest. Serv."
               Text            =   ""
               Formato         =   0
               Restricao       =   2
            End
            Begin VTOcx.fraVISUAL FraVariavel 
               Height          =   300
               Left            =   105
               TabIndex        =   126
               Top             =   345
               Width           =   4620
               _ExtentX        =   8149
               _ExtentY        =   529
               Status          =   1
               Altura          =   1905
               Caption         =   " Variáveis de Cadastro"
               CorTexto        =   16777215
               CorFaixa        =   16711680
               CorFundo        =   -2147483633
               Begin VTOcx.cboVISUAL CbovFuncionarioSUS 
                  Height          =   315
                  Left            =   465
                  TabIndex        =   130
                  Tag             =   "Logradouro"
                  Top             =   1425
                  Width           =   2460
                  _ExtentX        =   4339
                  _ExtentY        =   556
                  Caption         =   "Funcionário - SUS"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
               End
               Begin VTOcx.txtVISUAL txtvDataFormatura 
                  Height          =   285
                  Left            =   645
                  TabIndex        =   129
                  Top             =   1095
                  Width           =   3405
                  _ExtentX        =   6006
                  _ExtentY        =   503
                  Caption         =   "Data Formatura"
                  Text            =   ""
                  Formato         =   0
                  Restricao       =   2
               End
               Begin VTOcx.txtVISUAL txtvQtdItems 
                  Height          =   285
                  Left            =   285
                  TabIndex        =   128
                  Top             =   765
                  Width           =   3765
                  _ExtentX        =   6641
                  _ExtentY        =   503
                  Caption         =   "Quantidade de Item"
                  Text            =   ""
                  Restricao       =   2
               End
               Begin VTOcx.txtVISUAL txtvAnuncio 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   127
                  Top             =   435
                  Width           =   2730
                  _ExtentX        =   4815
                  _ExtentY        =   503
                  Caption         =   "Anúncio"
                  Text            =   ""
                  Formato         =   5
                  Restricao       =   2
               End
            End
            Begin VTOcx.txtVISUAL txtDataReabertura 
               Height          =   285
               Left            =   6960
               TabIndex        =   41
               Tag             =   "Início da Atividade"
               Top             =   3675
               Visible         =   0   'False
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   503
               Caption         =   "Dt.Reabertura"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
            End
            Begin VTOcx.txtVISUAL txtDataEncerramento 
               Height          =   285
               Left            =   4350
               TabIndex        =   40
               Tag             =   "Início da Atividade"
               Top             =   3615
               Visible         =   0   'False
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               Caption         =   "Dt.Encerramento"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
            End
            Begin VTOcx.cboVISUAL cboMatrizFilial 
               Height          =   510
               Left            =   75
               TabIndex        =   27
               Top             =   1365
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   900
               Caption         =   "Matriz/Filial"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL CboTipoCadastro 
               Height          =   315
               Left            =   255
               TabIndex        =   31
               Top             =   2805
               Width           =   3700
               _ExtentX        =   6535
               _ExtentY        =   556
               Caption         =   "Tipo de Cadastro"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboPorte 
               Height          =   510
               Left            =   1700
               TabIndex        =   28
               Tag             =   "Porte da Empresa"
               Top             =   1350
               Width           =   2500
               _ExtentX        =   4419
               _ExtentY        =   900
               Caption         =   "Porte da Empresa"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboObrigIss 
               Height          =   510
               Left            =   5910
               TabIndex        =   25
               Tag             =   "Obrigação do ISSQN"
               Top             =   690
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   900
               Caption         =   "Recolhimento do ISSQN"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboAtivPoder 
               Height          =   510
               Left            =   4215
               TabIndex        =   24
               Tag             =   "Atividade Exercida Poder"
               Top             =   690
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   900
               Caption         =   "Atv.Exercida Poder"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboNatJur 
               Height          =   510
               Left            =   60
               TabIndex        =   22
               Tag             =   "Natureza Jurídica"
               Top             =   690
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   900
               Caption         =   "Natureza Jurídica"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboClassAtiv 
               Height          =   510
               Left            =   1920
               TabIndex        =   23
               Tag             =   "Classificação de Atividade"
               Top             =   690
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   900
               Caption         =   "Classificação de Atividade"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtEmpregados 
               Height          =   480
               Left            =   75
               TabIndex        =   35
               Tag             =   "Empregados"
               Top             =   1365
               Visible         =   0   'False
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   847
               Caption         =   "Empregados"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.fraVISUAL fraVISUAL4 
               Height          =   1005
               Left            =   4215
               TabIndex        =   56
               Top             =   1245
               Width           =   5400
               _ExtentX        =   9525
               _ExtentY        =   1773
               Altura          =   1905
               Caption         =   " Somente para Autônomos"
               CorTexto        =   16777215
               CorFaixa        =   16711680
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.txtVISUAL txtCRC 
                  Height          =   285
                  Left            =   3240
                  TabIndex        =   144
                  Top             =   675
                  Visible         =   0   'False
                  Width           =   2070
                  _ExtentX        =   3651
                  _ExtentY        =   503
                  Caption         =   "Nº CRC"
                  Text            =   ""
               End
               Begin VTOcx.cboVISUAL cboNivel 
                  Height          =   315
                  Left            =   1845
                  TabIndex        =   143
                  Top             =   315
                  Width           =   3495
                  _ExtentX        =   6165
                  _ExtentY        =   556
                  Caption         =   "Nível de Instrução"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
               End
               Begin VTOcx.txtVISUAL txtConselho 
                  Height          =   285
                  Left            =   75
                  TabIndex        =   36
                  Top             =   330
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   503
                  Caption         =   "Conselho"
                  Text            =   ""
               End
               Begin VTOcx.txtVISUAL txtRegistro 
                  Height          =   285
                  Left            =   135
                  TabIndex        =   37
                  Top             =   675
                  Width           =   2070
                  _ExtentX        =   3651
                  _ExtentY        =   503
                  Caption         =   "Nº Registro"
                  Text            =   ""
               End
            End
            Begin VTOcx.cboVISUAL cboAtivSecund 
               Height          =   510
               Left            =   75
               TabIndex        =   44
               Top             =   3630
               Width           =   4650
               _ExtentX        =   8202
               _ExtentY        =   900
               Caption         =   "Atividade Secundária"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboAtivSecund2 
               Height          =   510
               Left            =   4755
               TabIndex        =   45
               Top             =   3735
               Width           =   4950
               _ExtentX        =   8731
               _ExtentY        =   900
               Caption         =   "Atividade Secundária"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboIsento 
               Height          =   510
               Left            =   8745
               TabIndex        =   26
               Tag             =   "Isento"
               Top             =   675
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   900
               Caption         =   "Isento"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtDtInicio 
               Height          =   285
               Left            =   6225
               TabIndex        =   33
               Tag             =   "Início da Atividade"
               Top             =   2820
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   503
               Caption         =   "Início da Atividade"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
            End
            Begin VTOcx.txtVISUAL txtFator 
               Height          =   480
               Left            =   8200
               TabIndex        =   39
               Top             =   2220
               Visible         =   0   'False
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   847
               Caption         =   ""
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboPonto 
               Height          =   315
               Left            =   4785
               TabIndex        =   43
               Top             =   3645
               Width           =   4845
               _ExtentX        =   8546
               _ExtentY        =   556
               Caption         =   "Ponto Recepção"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboRamo 
               Height          =   315
               Left            =   225
               TabIndex        =   42
               Top             =   3585
               Width           =   4470
               _ExtentX        =   7885
               _ExtentY        =   556
               Caption         =   "Ramo Atividade"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cmdVISUAL cmdAdAtiv 
               Height          =   300
               Left            =   9300
               TabIndex        =   38
               Tag             =   "TCIS101"
               Top             =   2415
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   529
               Caption         =   ""
               Acao            =   5
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cboVISUAL cboAtivServ 
               Height          =   510
               Left            =   1050
               TabIndex        =   30
               Tag             =   "Atividade Principal"
               Top             =   2220
               Width           =   8200
               _ExtentX        =   14473
               _ExtentY        =   900
               Caption         =   "Atividade Principal"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
         Height          =   7005
         Left            =   30
         TabIndex        =   57
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   12356
         _Version        =   131082
         TabGuid         =   "TCIS101.frx":0B25
         Begin VTOcx.fraVISUAL fraAnu 
            CausesValidation=   0   'False
            Height          =   1920
            Left            =   135
            TabIndex        =   58
            Top             =   345
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   3387
            Altura          =   1905
            Caption         =   " Dados do Veículo de Divulgacão"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboItem 
               Height          =   510
               Left            =   60
               TabIndex        =   66
               Top             =   825
               Width           =   9570
               _ExtentX        =   16880
               _ExtentY        =   900
               Caption         =   "Sub Publicidade"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtFatorMutiplicador 
               Height          =   480
               Left            =   3945
               TabIndex        =   65
               Top             =   1365
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   847
               Caption         =   "Mutiplicador"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtValorApagar 
               Height          =   480
               Left            =   8340
               TabIndex        =   64
               Top             =   1365
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   847
               Caption         =   "Valor"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtValor 
               Height          =   480
               Left            =   7095
               TabIndex        =   63
               Top             =   1365
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   847
               Caption         =   "Valor em UFM"
               Text            =   ""
               Enabled         =   0   'False
               TipoLetras      =   0
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtDimensao 
               Height          =   480
               Left            =   75
               TabIndex        =   62
               Top             =   1365
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   847
               Caption         =   "Descrição"
               Text            =   ""
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtArea 
               Height          =   480
               Left            =   2430
               TabIndex        =   61
               Top             =   1365
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   847
               Caption         =   "Área Total"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.cboVISUAL cboMovimento 
               Height          =   510
               Left            =   60
               TabIndex        =   60
               Top             =   315
               Width           =   9570
               _ExtentX        =   16880
               _ExtentY        =   900
               Caption         =   "Publicidade"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtDataInstalacao 
               Height          =   480
               Left            =   5325
               TabIndex        =   59
               Top             =   1365
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   847
               Caption         =   "Data Instalação"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
         End
         Begin VTOcx.cmdVISUAL cmdAdAnuncio 
            Height          =   390
            Left            =   165
            TabIndex        =   67
            Top             =   3885
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   688
            Caption         =   "&Adicionar Anúncio"
            Acao            =   1
            Enabled         =   0   'False
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.grdVISUAL grdAnuncio 
            Height          =   2145
            Left            =   150
            TabIndex        =   68
            Top             =   4305
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   3784
            CorBorda        =   16711680
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
            OcultarRodape   =   -1  'True
         End
         Begin Threed.SSCheck chkCad 
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   69
            Top             =   105
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   344
            _Version        =   196610
            Caption         =   "Cadastrar"
         End
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   1545
            Left            =   150
            TabIndex        =   70
            Top             =   2310
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   2725
            Altura          =   1905
            Caption         =   " Localização do Imóvel"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboBairroAnuncio 
               Height          =   315
               Left            =   990
               TabIndex        =   75
               Top             =   1050
               Width           =   8565
               _ExtentX        =   15108
               _ExtentY        =   556
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNumeroAnuncio 
               Height          =   315
               Left            =   7950
               TabIndex        =   74
               Top             =   690
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               Caption         =   "Número"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtInscImob 
               Height          =   315
               Left            =   120
               TabIndex        =   73
               Top             =   330
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   556
               Caption         =   "Insc. Imobiliária"
               Text            =   ""
               Restricao       =   2
            End
            Begin VTOcx.cboVISUAL cboTipoLOgraAnuncio 
               Height          =   315
               Left            =   540
               TabIndex        =   72
               Top             =   690
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   556
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Editavel        =   -1  'True
            End
            Begin VTOcx.cboVISUAL cboLogradouroAnuncio 
               Height          =   315
               Left            =   3750
               TabIndex        =   71
               Top             =   705
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   7005
         Left            =   30
         TabIndex        =   76
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   12356
         _Version        =   131082
         TabGuid         =   "TCIS101.frx":0B4D
         Begin VTOcx.cmdVISUAL cmdAdVeiculo 
            Height          =   375
            Left            =   90
            TabIndex        =   77
            Top             =   2235
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   661
            Caption         =   "&Adicionar Veículo"
            Acao            =   1
            Enabled         =   0   'False
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.grdVISUAL grdVeiculo 
            Height          =   3840
            Left            =   60
            TabIndex        =   78
            Top             =   2655
            Width           =   9750
            _ExtentX        =   17198
            _ExtentY        =   6773
            CorBorda        =   16711680
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
            OcultarRodape   =   -1  'True
            CheckBox        =   -1  'True
         End
         Begin Threed.SSCheck chkCad 
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   79
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   344
            _Version        =   196610
            Caption         =   "Cadastrar"
         End
         Begin VTOcx.cboVISUAL cboAtividadeVeiculo 
            Height          =   315
            Left            =   90
            TabIndex        =   131
            Top             =   360
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   556
            Caption         =   "Atividade Desempenhada"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.fraVISUAL fraTrans 
            Height          =   1395
            Left            =   60
            TabIndex        =   80
            Top             =   750
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   2461
            Altura          =   1905
            Caption         =   " Representante Legal"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtLicenca 
               Height          =   480
               Left            =   2685
               TabIndex        =   139
               Top             =   810
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   847
               Caption         =   "Licenciamento"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboUFTransp 
               Height          =   510
               Left            =   4020
               TabIndex        =   140
               Top             =   795
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   900
               Caption         =   "UF"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtMunicipio 
               Height          =   480
               Left            =   120
               TabIndex        =   138
               Top             =   810
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   847
               Caption         =   "Município"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtAnoFabric 
               Height          =   480
               Left            =   5910
               TabIndex        =   135
               Top             =   300
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   847
               Caption         =   "Ano"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtPlaca 
               Height          =   480
               Left            =   6735
               TabIndex        =   136
               Top             =   300
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   847
               Caption         =   "Placa"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtChassi 
               Height          =   480
               Left            =   7920
               TabIndex        =   137
               Top             =   300
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   847
               Caption         =   "Chassi"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtMarca 
               Height          =   480
               Left            =   1800
               TabIndex        =   133
               Top             =   300
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   847
               Caption         =   "Marca"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtModelo 
               Height          =   480
               Left            =   3660
               TabIndex        =   134
               Top             =   300
               Width           =   2220
               _ExtentX        =   3916
               _ExtentY        =   847
               Caption         =   "Modelo"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtVeiculo 
               Height          =   480
               Left            =   120
               TabIndex        =   132
               Top             =   300
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   847
               Caption         =   "Veículo"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtInicioAtividadeCarro 
               Height          =   480
               Left            =   4875
               TabIndex        =   141
               Top             =   810
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   847
               Caption         =   "Inicio Atividade "
               Text            =   ""
               Formato         =   0
               AlinhamentoRotulo=   1
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   7005
         Left            =   30
         TabIndex        =   81
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   12356
         _Version        =   131082
         TabGuid         =   "TCIS101.frx":0B75
         Begin VTOcx.fraVISUAL fraRepresen1 
            CausesValidation=   0   'False
            Height          =   1365
            Left            =   60
            TabIndex        =   82
            Top             =   525
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   2408
            Altura          =   1905
            Caption         =   " Representante Legal"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtImRepresentante 
               Height          =   480
               Left            =   120
               TabIndex        =   85
               Top             =   300
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   847
               Caption         =   "Im"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   8
               AlinhamentoRotulo=   1
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNomeRepresentante 
               Height          =   480
               Left            =   2145
               TabIndex        =   84
               Top             =   795
               Width           =   7485
               _ExtentX        =   13203
               _ExtentY        =   847
               Caption         =   "Nome"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCpfRepresentante 
               Height          =   480
               Left            =   120
               TabIndex        =   83
               Top             =   795
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   847
               Caption         =   "CPF"
               Text            =   ""
               Formato         =   1
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
         End
         Begin VTOcx.fraVISUAL fraRepresen2 
            CausesValidation=   0   'False
            Height          =   1410
            Left            =   60
            TabIndex        =   86
            Top             =   1920
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   2487
            Altura          =   1905
            Caption         =   " Endereço do Representante"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboTipoLogrRepresentante 
               Height          =   510
               Left            =   75
               TabIndex        =   88
               Top             =   300
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   900
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboLogrRepresentante 
               Height          =   510
               Left            =   1335
               TabIndex        =   89
               Top             =   300
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   900
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Editavel        =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtNumRepresentante 
               Height          =   480
               Left            =   4950
               TabIndex        =   90
               Top             =   315
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   847
               Caption         =   "Nº"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtComplementoRepresentante 
               Height          =   480
               Left            =   5625
               TabIndex        =   91
               Top             =   315
               Width           =   4020
               _ExtentX        =   7091
               _ExtentY        =   847
               Caption         =   "Complemento"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboBairroRepresentante 
               Height          =   510
               Left            =   90
               TabIndex        =   92
               Top             =   810
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   900
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtTelefoneRepresentante 
               Height          =   480
               Left            =   4905
               TabIndex        =   93
               Top             =   825
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   847
               Caption         =   "Telefone"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCidadeRepresentante 
               Height          =   480
               Left            =   6255
               TabIndex        =   94
               Top             =   825
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   847
               Caption         =   "Cidade"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboUfRepresentante 
               Height          =   510
               Left            =   8820
               TabIndex        =   87
               Top             =   810
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   900
               Caption         =   "UF"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
         End
         Begin Threed.SSCheck chkCad 
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   95
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   344
            _Version        =   196610
            Caption         =   "Cadastrar"
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   7005
         Left            =   30
         TabIndex        =   96
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   12356
         _Version        =   131082
         TabGuid         =   "TCIS101.frx":0B9D
         Begin VTOcx.fraVISUAL fraContador 
            CausesValidation=   0   'False
            Height          =   1590
            Left            =   105
            TabIndex        =   97
            Top             =   495
            Width           =   7470
            _ExtentX        =   13176
            _ExtentY        =   2805
            Altura          =   1905
            Caption         =   " Dados do Contador"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboContador 
               Height          =   510
               Left            =   75
               TabIndex        =   98
               Top             =   435
               Width           =   7350
               _ExtentX        =   12965
               _ExtentY        =   900
               Caption         =   "Contador"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtCgcEscritorio 
               Height          =   480
               Left            =   5025
               TabIndex        =   101
               Top             =   990
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   847
               Caption         =   "CNPJ Escritório"
               Text            =   ""
               Formato         =   2
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCpfContador 
               Height          =   480
               Left            =   75
               TabIndex        =   99
               Top             =   990
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   847
               Caption         =   "CPF"
               Text            =   ""
               Formato         =   1
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCrcContador 
               Height          =   480
               Left            =   2550
               TabIndex        =   100
               Top             =   990
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   847
               Caption         =   "CRC"
               Text            =   ""
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
            End
         End
         Begin Threed.SSCheck chkCad 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   102
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   344
            _Version        =   196610
            Caption         =   "Cadastrar"
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   7005
         Left            =   30
         TabIndex        =   103
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   12356
         _Version        =   131082
         TabGuid         =   "TCIS101.frx":0BC5
         Begin Threed.SSCheck chkCad 
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   104
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   344
            _Version        =   196610
            Caption         =   "Cadastrar"
         End
         Begin VTOcx.fraVISUAL fraSocio1 
            Height          =   900
            Left            =   60
            TabIndex        =   105
            Top             =   390
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   1588
            Altura          =   1905
            Caption         =   " Dados do Sócio"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Enabled         =   0   'False
            Begin VTOcx.txtVISUAL txtCargoSocio 
               Height          =   480
               Left            =   7470
               TabIndex        =   108
               Top             =   315
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   847
               Caption         =   "Cargo"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtNomeSocio 
               Height          =   480
               Left            =   2100
               TabIndex        =   107
               Top             =   315
               Width           =   5355
               _ExtentX        =   9446
               _ExtentY        =   847
               Caption         =   "Nome"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCpfSocio 
               Height          =   480
               Left            =   75
               TabIndex        =   106
               Top             =   315
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   847
               Caption         =   "CPF"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   1
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
         End
         Begin VTOcx.fraVISUAL fraSocio2 
            Height          =   1410
            Left            =   60
            TabIndex        =   109
            Top             =   1320
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   2487
            Altura          =   1905
            Caption         =   " Endereço do Sócio"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Enabled         =   0   'False
            Begin VTOcx.cboVISUAL cboUFSocio 
               Height          =   510
               Left            =   8820
               TabIndex        =   117
               Top             =   810
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   900
               Caption         =   "UF"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Enabled         =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCidadeSocio 
               Height          =   480
               Left            =   6255
               TabIndex        =   116
               Top             =   825
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   847
               Caption         =   "Cidade"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtTelSocio 
               Height          =   480
               Left            =   4905
               TabIndex        =   115
               Top             =   825
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   847
               Caption         =   "Telefone"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboBairroSocio 
               Height          =   510
               Left            =   90
               TabIndex        =   114
               Top             =   810
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   900
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Enabled         =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCompSocio 
               Height          =   480
               Left            =   5625
               TabIndex        =   113
               Top             =   315
               Width           =   4020
               _ExtentX        =   7091
               _ExtentY        =   847
               Caption         =   "Complemento"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtNumSocio 
               Height          =   480
               Left            =   4950
               TabIndex        =   112
               Top             =   315
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   847
               Caption         =   "Nº"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboLogrSocio 
               Height          =   510
               Left            =   1635
               TabIndex        =   111
               Top             =   300
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   900
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Editavel        =   -1  'True
               Enabled         =   0   'False
            End
            Begin VTOcx.cboVISUAL cboTipoLogrSocio 
               Height          =   510
               Left            =   75
               TabIndex        =   110
               Top             =   300
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   900
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Enabled         =   0   'False
            End
         End
         Begin VTOcx.cmdVISUAL cmdAdEdif 
            Height          =   375
            Left            =   60
            TabIndex        =   118
            Top             =   2775
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   661
            Caption         =   "&Adicionar Sócio"
            Acao            =   1
            Enabled         =   0   'False
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.grdVISUAL grdSocio 
            Height          =   3195
            Left            =   60
            TabIndex        =   119
            Top             =   3180
            Width           =   9750
            _ExtentX        =   17198
            _ExtentY        =   5636
            CorBorda        =   16711680
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
            OcultarRodape   =   -1  'True
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   50
      Top             =   8130
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   714
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   330
         Left            =   5610
         TabIndex        =   46
         Top             =   60
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   6705
         TabIndex        =   47
         Top             =   60
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   7800
         TabIndex        =   48
         Top             =   60
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   8895
         TabIndex        =   49
         Top             =   60
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   148
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   1138
      Icone           =   "TCIS101.frx":0BED
   End
End
Attribute VB_Name = "TCIS101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cadastro As VSImposto
Dim Transportador As eTransportador
Dim Contribuinte As eContribuinte
Dim Endereco As eEndereco
Dim Contador As eContador
Dim atividade As atividade
Dim Imovel As eImovel
Dim Socio As eSocio
Dim Representante As eRepresentante
Dim GraveiContrib As Boolean
Dim VaiGravarSocio As Boolean
Dim InscricaoMunicipal As String
Dim InscricaoAuxiliar As String
Dim Anuncio As eAnuncio

Private Function VerificaCpfCgc() As Boolean
    Dim strCPF As String
    Dim strCGC As String
    Dim blnValido As Boolean
    
    blnValido = True
    
    If Trim(txtCgc) = "" Then Exit Function
    If Len(Edita.TiraTudo(txtCgc)) = 11 Then
        strCPF = Edita.TiraTudo(txtCgc)
        If Util.ValidaCpf(strCPF) = False Then
            blnValido = False
        End If
        Select Case strCPF
            Case String(11, "1")
                blnValido = False
            Case String(11, "2")
                blnValido = False
            Case String(11, "3")
                blnValido = False
            Case String(11, "4")
                blnValido = False
            Case String(11, "5")
                blnValido = False
            Case String(11, "6")
                blnValido = False
            Case String(11, "7")
                blnValido = False
            Case String(11, "8")
                blnValido = False
            Case String(11, "9")
                blnValido = False
            Case String(11, "0")
                blnValido = False
        End Select
    ElseIf Len(Edita.TiraTudo(txtCgc)) = 14 Then
        strCGC = Edita.TiraTudo(txtCgc)
'        If Util.ValidaCgc(strCGC) = False Then
'            blnValido = False
'            Exit Function
'        End If
    Else
        blnValido = False
    End If
    
    VerificaCpfCgc = blnValido
    
    If blnValido = False Then
        Util.Avisa "Cpf ou Cnpj Inválido"
    End If
    
End Function

Private Sub cboAtivServ_LostFocus()
    Dim RetFator As String
    Dim RetNivel As String
    If Trim(cboAtivServ) = "" Then Exit Sub
    txtCodAtividade = atividade.TrazerCodigo(cboAtivServ.Text)
    
    If atividade.BuscaFator(cboAtivServ, RetFator, RetNivel) Then
        'txtFator.Visible = True
        cboNivel.SetarLinha RetNivel, 1
        txtFator.Caption = RetFator
        txtFator.Tag = "Fator"
        'txtFator.SetFocus
    Else
        txtFator.Visible = False
        txtFator.Tag = ""
    End If
End Sub

Private Sub cboClassAtiv_Click()
    atividade.PreencherCboAtiv cboAtivServ, CStr(cboClassAtiv.Coluna(1).Valor)
End Sub

Private Sub cboContador_Click()
    If cboContador.ListIndex = -1 Then Exit Sub
    If Contribuinte.Buscar(CStr(cboContador.Coluna(1).Valor), , False) = True Then
                    txtCpfContador.Formato = formDocumento
'        txtCgcEscritorio = "": txtCrcContador = "": txtCpfContador = ""
        txtCrcContador = Contribuinte.Registro
        If Len(Trim(Contribuinte.CgcCpf)) = 14 And Not IsNumeric(Contribuinte.CgcCpf) Then
            txtCpfContador = Contribuinte.CgcCpf
            txtCgcEscritorio.SetFocus
        Else
            txtCgcEscritorio = Contribuinte.CgcCpf
            txtCpfContador = Contribuinte.CgcCpf
            txtCpfContador.SetFocus
        End If
    End If
End Sub
Private Sub verificaFantasia()
    If Len(txtFantasia) = 0 Then
        txtFantasia = txtRazao
    End If
End Sub
    
Private Sub cboEstabelece_LostFocus()
    If cboEstabelece = "SIM" Then
        txtIc.Tag = "Insc. Cadastral"
        cboImovel.Tag = "Imovel"
    Else
        txtIc.Tag = ""
        cboImovel.Tag = ""
    End If
End Sub

Private Sub cboGrupoAtividade_Click()
    atividade.PreencherCboAtiv CboAtividade, CStr(cboGrupoAtividade.Coluna(1).Valor)
End Sub

Private Sub cboItem_Click()
 On Error Resume Next
    txtValor = cboItem.Coluna(2).Valor
    Calcula
End Sub

Private Sub cboMovimento_Click()
  txtValor = "0,00"
  txtValorApagar = "0,00"
  If Bdados.AbreTabela("SELECT * From TAB_PARAMETRO_DETALHE where  TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(cboMovimento.Coluna(0).Valor, tctexto)) Then
        cboItem.Enabled = True
        cboItem.Preencher Bdados, "SELECT TPD_TIP_COD_IMPOSTO,tpd_descricao ,tpd_valor_ufm,tpd_item  From TAB_PARAMETRO_DETALHE where  TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(cboMovimento.Coluna(0).Valor, tctexto), 1
        txtDimensao.Text = ""
        txtDimensao.Enabled = Not txtDimensao.Enabled
        txtArea.Text = ""
        txtArea.Enabled = Not txtArea.Enabled
        txtFatorMutiplicador.Text = ""
        txtFatorMutiplicador.Enabled = Not txtFatorMutiplicador.Enabled
    Else
        txtDimensao.Enabled = True
        txtArea.Enabled = True
        txtFatorMutiplicador.Enabled = True
        cboItem.Clear
        cboItem.Enabled = False
    End If
End Sub

Private Sub cboTipoLogrRepresentante_Click()
    If cboTipoLogrRepresentante.ListIndex = -1 Then Exit Sub
    Endereco.PreencherCboRua cboLogrRepresentante, cboTipoLogrRepresentante
End Sub

Private Sub cboTipoLogrSocio_Click()
    If cboTipoLogrSocio.ListIndex = -1 Then Exit Sub
    Endereco.PreencherCboRua cboLogrSocio, cboTipoLogrSocio
End Sub

Private Sub chkCad_Click(Index As Integer, Value As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            fraSocio1.Enabled = chkCad(Index).Value
            fraSocio2.Enabled = chkCad(Index).Value
            cmdAdEdif.Enabled = chkCad(Index).Value
            txtCpfSocio.SetFocus
        Case 1
            fraContador.Enabled = chkCad(Index).Value
            cboContador.SetFocus
        Case 2
            fraRepresen1.Enabled = chkCad(Index).Value
            fraRepresen2.Enabled = chkCad(Index).Value
            txtCpfRepresentante.SetFocus
        Case 3
            fraTrans.Enabled = chkCad(Index).Value
            cmdAdVeiculo.Enabled = chkCad(Index).Value
            cboAtividadeVeiculo.Enabled = chkCad(Index).Value
            cboAtividadeVeiculo.SetFocus
        Case 4
            fraAnu.Enabled = chkCad(Index).Value
            cmdAdAnuncio.Enabled = chkCad(Index).Value
            cboMovimento.SetFocus
    End Select
End Sub

Private Sub cmdAdAnuncio_Click()
   Dim RetIm                           As String
    Dim i                                   As Byte
    Dim Sql                               As String
    Dim rs                                As VSRecordset
    Dim Index                            As Integer
    If cboMovimento.ListIndex = -1 Then Exit Sub
    
   Index = grdAnuncio.ListItems.Count + 1
   grdAnuncio.ListItems.Add Index, , Index
   grdAnuncio.ListItems.Item(Index).SubItems(1) = cboMovimento.Coluna(0).Valor & " - " & cboMovimento.Text
   grdAnuncio.ListItems.Item(Index).SubItems(2) = txtDimensao
   grdAnuncio.ListItems.Item(Index).SubItems(3) = txtArea
   grdAnuncio.ListItems.Item(Index).SubItems(4) = txtDataInstalacao
   grdAnuncio.ListItems.Item(Index).SubItems(5) = txtValor
   grdAnuncio.ListItems.Item(Index).SubItems(6) = txtValorApagar
   If cboItem.Enabled Then
        grdAnuncio.ListItems.Item(Index).SubItems(7) = cboItem.Coluna(3).Valor & " - " & cboItem.Coluna(1).Valor
   End If
   grdAnuncio.ListItems.Item(Index).SubItems(8) = TiraPic(InscricaoMunicipal, "-") & Format(Index, "00")
   txtDimensao = ""
   txtArea = ""
   txtDataInstalacao = ""
   cboMovimento.ListIndex = -1
   txtValor = "0,00"
   If cboItem.Enabled Then
     txtDataInstalacao.SetFocus
   Else
    txtDimensao.SetFocus
   End If
   End Sub



Private Function Pega_Posicao_Grupo_Atividade()
    
End Function

Private Sub cmdAdAtividade_Click()
    Set Contribuinte = New eContribuinte
    Call Contribuinte.Adiciona_Atividade_Secundaria(grdAtividade, cboGrupoAtividade, CboAtividade)
    
End Sub

Private Sub cmdExcluir_Click()
    Contribuinte.Apaga_Atividade_Secundaria grdAtividade
End Sub

Private Sub cmdSalvar_Click()
    On Error Resume Next
    Dim Item As Object
    Dim Conta As New ContaCorrente
    
    If Not GraveiContrib Then
        'If Not Edita.CriticaCampos(Me) Then Exit Sub
        
        cboTipoLogr.Tag = ""
        cboLogr.Tag = ""
        CboAtividade.Tag = ""
        cboGrupoAtividade.Tag = ""
        txtDataReabertura.Tag = ""
        txtDataEncerramento.Tag = ""
        txtDataReabertura.Tag = ""
        txtInscImob.Tag = ""
        txtIc.Tag = ""
        cboImovel.Tag = ""
        cboPorte.Tag = ""
        cboIsento.Tag = ""
        
        If VerificaCpfCgc = False Then
            'Util.Avisa "Informe o Documento do contribuinte - CPF/CNPJ"
            
            'tabCadastro.Tabs(1).Selected = True
            'txtCgc.SetFocus
            'Exit Sub
        End If
        If Len(Trim(txtCodAtividade)) = 0 Or txtCodAtividade = 0 Then
            If cboAtivServ.Coluna(1).Valor = cboAtivSecund.Coluna(1).Valor Then
                Util.Avisa "Atividade principal deve ser diferente da atividade segundária."
                cboAtivSecund.SetFocus
                Exit Sub
            End If
        End If
        Screen.MousePointer = 11
        'Buscando as Inscricoesno
'        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
'            InscricaoMunicipal = Conta.GeraCodPagamento("CADASTRO ECONOMICO")
'            InscricaoAuxiliar = Cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
'        Else
            InscricaoMunicipal = Cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
            InscricaoAuxiliar = ""
'
'        End If
        'Vou cadastrar o contribuinte
        With Contribuinte
            .Rg = txtRG
            .Im = InscricaoMunicipal
            .ImAuxiliar = txtIncriAuxiliar
            'BCP
            .ImAnterior = txtIncriAuxiliar
            .AreaEstabelecimento = txtAreaEstabelecimento
            .Obs = txtObservacaoCompleta
           ' .Crc = txtCrc
            'FIM BCP
            .CgcCpf = txtCgc
            .Nome = txtRazao
            .TipoCadastro = CStr(Nvl("" & cboTipoCadastro.Coluna(1).Valor, 0))
            .Fantasia = txtFantasia
            .InicioPrestacaoServico = txtInicioPrestacaoServico
            .SituacaoAlvara = cboSitAlvara.Coluna(1).Valor
            .Cep = txtCep
            .Matriz_Filial = CStr(Nvl("" & cboMatrizFilial.Coluna(1).Valor, 0))
            .Data_Encerramento = txtDataEncerramento
            .VariavelAnuncio = txtvAnuncio
            .VariavelDataFormatura = txtvDataFormatura
            .VariavelFuncionarioSUS = CbovFuncionarioSUS.Coluna(1).Valor
            .VariavelQuantidadeItem = txtvQtdItems
            .Data_Reabertura = txtDataReabertura
            .CodGrupo = Nvl("" & CStr(Nvl("" & cboClassAtiv.Coluna(1).Valor, 0)), 0)
            .CodSitCadastral = 1
            .CodNatureza = Nvl("" & CStr(Nvl("" & cboNatJur.Coluna(1).Valor, 0)), 0)
            .CodAtivPoder = Nvl("" & cboAtivPoder.Coluna(1).Valor, 0)
            .Estabelecido = Nvl("" & CStr(Nvl("" & cboEstabelece.Coluna(1).Valor, 0)), 0)
            .DataCadastro = Bdados.Converte(Date, TCDataHora)
            .Nome_Tela = Me.Caption
            If txtCodAtividade = 0 Or Len(Trim(txtCodAtividade)) = 0 Then
                .CodAtividade = Nvl("" & CStr(Nvl("" & cboAtivServ.Coluna(1).Valor, 0)), 0)
            Else
                .CodAtividade = txtCodAtividade
            End If
            .CodUsuario = AplicacoesVTFuncoes.Usuario
            .Logradouro = cboTipoLogr
            .NomeLogradouro = cboLogr
            .Numero = txtNum
            .Complemento = txtComplemento
            .Bairro = cboBairro
            .Cidade = txtCidade
            .Uf = cboUf
            .GrupoAtividade = Nvl("" & CStr(Nvl("" & cboAtivServ.Coluna(2).Valor, 0)), 0)
            .InicioAtividade = Bdados.Converte(txtDtInicio, TCDataHora)
            .TipoContribuinte = CStr(Nvl("" & cboNatJur.Coluna(1).Valor, 0))
            .TipoRecolhimentoIss = CStr(Nvl("" & cboObrigIss.Coluna(1).Valor, 0))
            .Ruc = txtRuc
            If Trim(txtFator) = "" Then
                txtFator = "1"
            End If
            .FatorAlvara = txtFator
            .Conselho = txtConselho
            .Registro = IIf(Trim(txtRegistro) = "", 0, txtRegistro)
            If cboImovel.ListIndex <> -1 Then
                .ImovelProprio = CStr(Nvl("" & cboImovel.Coluna(1).Valor, 0))
            End If
            .NumEmpregado = Nvl(txtEmpregados, 0)
            .PorteEmpresa = CStr(Nvl("" & cboPorte.Coluna(1).Valor, 0))
            .CodAtividadeSec = CStr(Nvl("" & IIf(Trim(cboAtivSecund) = "", 0, cboAtivSecund.Coluna(1).Valor), 0))
            .CodAtividadeTerc = CStr(Nvl("" & IIf(Trim(cboAtivSecund2) = "", 0, cboAtivSecund2.Coluna(1).Valor), 0))
            .NivelEscolar = cboNivel.Coluna(1).Valor
            .Protocolo = txtProtocolo
            If Trim(txtIc) <> "" Then .Ic = txtIc
            If Nvl(CStr(Nvl("" & cboIsento.Coluna(1).Valor, 0)), 0) Then .Isento = CStr(Nvl("" & cboIsento.Coluna(1).Valor, 0))
            .CodRamo = Nvl(CStr(Nvl("" & cboRamo.Coluna(1).Valor, 0)), 0)
            .FoneFax = CDbl(IIf(txtFoneFax = "", 0, txtFoneFax))
            .CNH = txtCnh
            .Categoria = txtCateg
            .Autorizacao = Nvl(txtAutoriza, 0)
            .PontoRecepcao = cboPonto
            'Salva dados do contribuinte
            .Salvar
        End With
        'GRAVANDO SOCIOS
        If chkCad(0).Value Then
            If grdSocio.ListItems.Count <> "0" Then
                For Each Item In grdSocio.ListItems
                    Item.Selected = True
                    With Socio
                        .Im = InscricaoMunicipal
                        .Cpf = Item
                        .Nome = Item.SubItems(1)
                        .Cargo = Item.SubItems(2)
                        .TipoLogr = Item.SubItems(3)
                        .Logr = Item.SubItems(4)
                        .Numero = Item.SubItems(5)
                        .Complemento = Item.SubItems(6)
                        .Bairro = Item.SubItems(7)
                        .Telefone = Item.SubItems(8)
                        .Cidade = Item.SubItems(9)
                        .Uf = Item.SubItems(10)
                        .Salvar
                    End With
                Next
            End If
        End If
        'GRAVANDO CONTADOR
        If chkCad(1).Value Then
            With Contador
                .Im = InscricaoMunicipal
                .Crc = txtCrcContador
                .Contador = cboContador
                .Cpf = txtCpfContador
                .CGCEscritorio = txtCgcEscritorio
                .Salvar
            End With
        End If
        'GRAVANDO REPRESENTANTE
        If chkCad(2).Value Then
            If Trim(txtCpfRepresentante) <> "" And Trim(txtNomeRepresentante) <> "" Then
                With Representante
                    .Im = InscricaoMunicipal
                    .Cpf = txtCpfRepresentante
                    .Nome = txtNomeRepresentante
                    .TipoLogr = cboTipoLogr
                    .Logr = cboLogr
                    If Trim(txtNumRepresentante) <> "" Then .Numero = Bdados.Converte(txtNumRepresentante, tctexto)
                    .Complemento = txtComplementoRepresentante
                    .Bairro = cboBairroRepresentante
                    .Cidade = txtCidadeRepresentante
                    .Telefone = txtTelefoneRepresentante
                    .Uf = cboUf
                    If Trim(txtImRepresentante) <> "" Then .ImRepresentante = Bdados.Converte(txtImRepresentante, tctexto)
                    .Salvar
                End With
            End If
        End If
        'GRAVANDO TRANSPORTADOR
        If chkCad(3).Value Then
             If grdVeiculo.ListItems.Count <> "0" Then 'Se definiu algum carro
                'If cboAtividadeVeiculo.ListIndex = -1 Then
                '    Avisa "Informe a atividade desempenhada."
                '    cboAtividadeVeiculo.SetFocus
                'End If
                'Contribuinte.AtividadeTransporte = cboAtividadeVeiculo.Coluna(1).VALOR
                For Each Item In grdVeiculo.ListItems
                 '   grdVeiculo.ListItems(Item).Selected = True
                    With Transportador
                        .Im = InscricaoMunicipal
                        .Veiculo = Item
                        .Marca = Item.SubItems(1)
                        .CodModelo = Item.SubItems(2)
                        If IsNumeric(Item.SubItems(3)) Then .AnoFabricacao = Item.SubItems(3)
                        .Placa = Item.SubItems(4)
                        .Chassi = Item.SubItems(5)
                        .municipio = Item.SubItems(6)
                        .Uf = Item.SubItems(7)
                        .Licensa = CInt(Item.SubItems(8))
                        Dim Pos As Integer
                        Pos = InStr(Item.SubItems(9), "-")
                        .atividade = Left(Item.SubItems(9), Pos - 1)
                        .IniAtividadeCarro = Item.SubItems(10)
                        .Salvar
                    End With
                Next
            End If
        End If
       'GRAVANDO ANUNCIOS
        Anuncio.Excluir InscricaoMunicipal 'EXCLUI ANUNCIOS
        If chkCad(4).Value Then
            If grdAnuncio.ListItems.Count <> "0" Then
                For Each Item In grdAnuncio.ListItems
                    With Anuncio
                        .Im = Bdados.Converte(InscricaoMunicipal, tctexto)
                        .icad = Bdados.Converte(Item.Text, tctexto)
                        .Movimento = Trim(Bdados.Converte(Trim(Left(Item.SubItems(1), 9)), tctexto))
                        .Dimensao = Item.SubItems(2)
                        .Area = Item.SubItems(3)
                        .DataInstalacao = Item.SubItems(4)
                        .Valor_UFM = Item.SubItems(5)
                        .Valor_Apagar = Item.SubItems(6)
                        'Dim Pos As Integer
                        Dim t As Integer
                        Pos = Val(InStr(Item.SubItems(7), " - ")) - 1
                        If Pos > -1 Then
                            .Item = Left(Item.SubItems(7), Pos)
                        Else
                            .Item = 0
                        End If
                        .Doc_Origem = Item.SubItems(8)
                        .Salvar
                    End With
                Next
            End If
        End If
        Call Contribuinte.Grava_Atividade_Secundaria(grdAtividade, InscricaoMunicipal)
        Call Util.Informa("Registro gravado com sucesso. Inscricão Municipal Gerada Nº: " & InscricaoMunicipal & ".")
        cmdLimpar_Click
        Screen.MousePointer = 0
    End If
    Exit Sub
trata:
        Erro Err.Number & " - " & Err.Description
        Screen.MousePointer = 0
        Exit Sub
        Resume
End Sub

Private Sub cmdAdAtiv_Click()
    Dim CodGrupo As Double
    CodGrupo = cboClassAtiv.ListIndex
    TATV401.Tag = cboClassAtiv.Text
    TATV401.Show 1
    cboClassAtiv_Click
    cboClassAtiv.ListIndex = CodGrupo
    If TCIS101.Tag <> "" Then
        cboAtivServ.ListIndex = ListIndexDe(cboAtivServ, TCIS101.Tag)
    End If
    Unload TATV101
End Sub

Private Sub cmdAdEdif_Click()
    Dim ItmX As Object
    Dim i As Byte
    If Trim(txtCpfSocio) = "" Then Exit Sub
    Set ItmX = grdSocio.ListItems.Add(, , txtCpfSocio)
    With ItmX
        .SubItems(1) = txtNomeSocio
        .SubItems(2) = txtCargoSocio
        .SubItems(3) = cboTipoLogrSocio
        .SubItems(4) = cboLogrSocio
        .SubItems(5) = txtNumSocio
        .SubItems(6) = txtCompSocio
        .SubItems(7) = cboBairroSocio
        .SubItems(8) = txtTelSocio
        .SubItems(9) = txtCidadeSocio
        .SubItems(10) = cboUFSocio
    End With
    txtCpfSocio = ""
    txtNomeSocio = ""
    txtCargoSocio = ""
    cboTipoLogrSocio = ""
    cboLogrSocio = ""
    txtNumSocio = ""
    txtCompSocio = ""
    cboBairroSocio = ""
    txtTelSocio = ""
    txtCidadeSocio = ""
    cboUFSocio.ListIndex = -1
    txtCpfSocio.SetFocus
End Sub

Private Sub MontaCabGrid()
'grid veiculo
    grdVeiculo.ColumnHeaders.Add , , "Veículo"
    grdVeiculo.ColumnHeaders.Add , , "Marca"
    grdVeiculo.ColumnHeaders.Add , , "Modelo"
    grdVeiculo.ColumnHeaders.Add , , "Ano Fabricação"
    grdVeiculo.ColumnHeaders.Add , , "Placa"
    grdVeiculo.ColumnHeaders.Add , , "Chassi"
    grdVeiculo.ColumnHeaders.Add , , "Cidade"
    grdVeiculo.ColumnHeaders.Add , , "UF"
    grdVeiculo.ColumnHeaders.Add , , "Licença"
    grdVeiculo.ColumnHeaders.Add , , "Atividade"
    grdVeiculo.ColumnHeaders.Add , , "Ini.Atividade"
'grid socio
    grdSocio.ColumnHeaders.Add , , "CPF"
    grdSocio.ColumnHeaders.Add , , "Nome"
    grdSocio.ColumnHeaders.Add , , "Cargo"
    grdSocio.ColumnHeaders.Add , , "Endereço"
    grdSocio.ColumnHeaders.Add , , "Endereço"
    grdSocio.ColumnHeaders.Add , , "Número"
    grdSocio.ColumnHeaders.Add , , "Compl."
    grdSocio.ColumnHeaders.Add , , "Bairro"
    grdSocio.ColumnHeaders.Add , , "Telefone"
    grdSocio.ColumnHeaders.Add , , "Cidade"
    grdSocio.ColumnHeaders.Add , , "UF"
    
 'Atividade
    grdAtividade.ColumnHeaders.Add , , "Classificação", 4000
    grdAtividade.ColumnHeaders.Add , , "Atividade", 4000
End Sub

Private Sub cmdAdVeiculo_Click()
    Dim RetIm       As String
    Dim i               As Byte
    Dim Index        As Integer
    
    Dim Sql           As String
    Dim rs            As VSRecordset
    
    If Trim(txtPlaca) = "" Then Exit Sub
    If Transportador.VerificaChassi(txtChassi, RetIm) Then
        Util.Informa "Chassi cadastrado para contribuinte IM = '" & RetIm & "'."
        Exit Sub
    End If
    
    Index = grdVeiculo.ListItems.Count + 1
    grdVeiculo.ListItems.Add Index, , txtVeiculo
    grdVeiculo.ListItems.Item(Index).SubItems(1) = txtMarca
    grdVeiculo.ListItems.Item(Index).SubItems(2) = txtModelo
    grdVeiculo.ListItems.Item(Index).SubItems(3) = txtAnoFabric
    grdVeiculo.ListItems.Item(Index).SubItems(4) = txtPlaca
    grdVeiculo.ListItems.Item(Index).SubItems(5) = txtChassi
    grdVeiculo.ListItems.Item(Index).SubItems(6) = txtMunicipio
    grdVeiculo.ListItems.Item(Index).SubItems(7) = cboUFTransp
    grdVeiculo.ListItems.Item(Index).SubItems(8) = txtLicenca
    grdVeiculo.ListItems.Item(Index).SubItems(9) = cboAtividadeVeiculo.Coluna(1).Valor & " - " & cboAtividadeVeiculo.Text
    grdVeiculo.ListItems.Item(Index).SubItems(10) = txtInicioAtividadeCarro
'    Set ItmX = grdVeiculo.ListItems.Add(, , txtVeiculo)
'    With ItmX
'        .SubItems(1) = txtMarca
'        .SubItems(2) = txtModelo
'        .SubItems(3) = txtAnoFabric
'        .SubItems(4) = txtPlaca
'        .SubItems(5) = txtChassi
'        .SubItems(6) = txtMunicipio
'        .SubItems(7) = cboUFTransp
'        .SubItems(8) = txtLicenca
'    End With
    txtVeiculo = ""
    txtMarca = ""
    txtModelo = ""
    txtInicioAtividadeCarro = ""
    txtAnoFabric = ""
    txtPlaca = ""
    txtChassi = ""
    txtMunicipio = ""
    txtLicenca = ""
    txtCidadeSocio = ""
    cboUFTransp.ListIndex = -1
    cboAtividadeVeiculo.ListIndex = -1
    cboAtividadeVeiculo.SetFocus
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{tab}"
End Sub

Private Sub cmdImprimir_Click()
        Screen.MousePointer = 11
        If Trim(InscricaoMunicipal) <> "" Then Imposto.ImprimeFC InscricaoMunicipal, Rpt
        Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
    'tabCadastro.TabEnable'd(0) = True
    tabCadastro.Tabs(1).Selected = True
    txtRG.SetFocus
    InscricaoMunicipal = ""
    Edita.LimpaCampos Me
    GraveiContrib = False
    VaiGravarSocio = False
    grdSocio.ListItems.Clear
    grdVeiculo.ListItems.Clear
    grdAnuncio.ListItems.Clear
    chkCad(0).Value = 0
    chkCad(1).Value = 0
    chkCad(2).Value = 0
    chkCad(3).Value = 0
    txtCidade = Aplicacoes.municipio
    cboUf.Text = "MA"
    txtCep = CepCliente
End Sub

Private Sub cmdSair_Click()
        Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtIc
End Sub

Private Sub Form_Activate()
    cboMovimento.Preencher Bdados, "Select tip_cod_imposto as [Código Receita] , tip_nome_imposto as Tributo FROM Tab_Imposto where tip_sigla_Imposto = 'PUBL'", 1
    If AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        cboPonto.Visible = False
    End If
    tabCadastro.Tabs(7).Enabled = False
    cboTipoCadastro.PreencherGeral Bdados, "TIPO CADASTRO ECONOMICO"
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
        txtImRepresentante.Formato = formNenhum
    Else
        txtImRepresentante.Formato = formDoisDigitos
    End If
    If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
        FraVariavel.Visible = True
    Else
        FraVariavel.Visible = False
    End If
    cboSitAlvara.PreencherGeral Bdados, "SITUACAO ALVARA"
    CbovFuncionarioSUS.PreencherGeral Bdados, "FUNCIONARIO SUS"
    cboNivel.Enabled = False
    txtCep.Formato = formNenhum
End Sub

Private Sub Form_Load()
    Set Contribuinte = New eContribuinte
    Set Transportador = New eTransportador
    Set Cadastro = New VSImposto
    Set Endereco = New eEndereco
    Set Contador = New eContador
    Set atividade = New atividade
    Set Socio = New eSocio
    Set Imovel = New eImovel
    Set Representante = New eRepresentante
    Set Anuncio = New eAnuncio
    cboUFTransp.Tag = ""
    txtCRC = 0
    txtAreaEstabelecimento = 0
    txtObservacaoCompleta = ""
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision

    Screen.MousePointer = 0
    With Endereco
        .PreencherCboTipoLogr cboTipoLogr
        .PreencherCboTipoLogr cboTipoLogr
         
        .PreencherCboTipoLogr cboTipoLogrSocio
        .PreencherCboTipoLogr cboTipoLogrRepresentante
        .PreencherPonto cboPonto
        .PreencherCboBairro cboBairro
        .PreencherCboBairro cboBairroRepresentante
        .PreencherCboBairro cboBairroSocio
    End With
    
    cboLogr.Preencher Bdados, "select tlg_cod_logradouro,TLG_NOME from tab_logradouro", 1
    cboLogrRepresentante.Preencher Bdados, "select tlg_cod_logradouro,TLG_NOME from tab_logradouro", 1
    Contador.PreencherCboContador cboContador
    cboEstabelece.PreencherGeral Bdados, "ESTABELECIDO"
    cboImovel.PreencherGeral Bdados, "TIPO IMÓVEL"
    cboObrigIss.PreencherGeral Bdados, "TIPO RECOLHIMENTO ISS"
    cboIsento.PreencherGeral Bdados, "SIM OU NÃO"
    cboNivel.PreencherGeral Bdados, "NIVEL INSTRUÇÃO"
    cboPorte.PreencherGeral Bdados, "PORTE EMPRESA"
    cboMatrizFilial.PreencherGeral Bdados, "TIPO EMPRESA"
    With atividade
        .PreencheCombo cboClassAtiv, iaGrupoAtividade
        .PreencheCombo cboGrupoAtividade, iaGrupoAtividade
        .PreencheCombo cboRamo, iaRamo
        .PreencherCboPoder cboAtivPoder
        .PreencherCboAtiv cboAtivServ
        .PreencherCboAtiv cboAtivSecund
        .PreencherCboAtiv cboAtividadeVeiculo
        .PreencherCboAtiv cboAtivSecund2
        .PreencherCboNturJur cboNatJur
    End With
    cboUf.PreencherGeral Bdados, "UF"
    cboUFSocio.PreencherGeral Bdados, "UF"
    cboUfRepresentante.PreencherGeral Bdados, "UF"
    cboUFTransp.PreencherGeral Bdados, "UF"
    
    GraveiContrib = False
    
    txtCidade = Aplicacoes.municipio
    cboUf.Text = Temp.PegaParametro(Bdados, "ESTADO UF")
    txtCep = CepCliente
    MontaCabGrid
    Anuncio.MontarGrid grdAnuncio
    
End Sub



Private Sub grdAnuncio_DblClick()
   Dim Contador As Integer
    
    With grdAnuncio
        If grdAnuncio.SelectedItem Is Nothing Then Exit Sub
        cboMovimento.SetarLinha Trim(Left(.SelectedItem.SubItems(1), 9)), 0
        cboMovimento_Click
        Contador = InStr(grdAnuncio.SelectedItem.SubItems(7), " - ")
        If .SelectedItem.SubItems(7) <> "" Then cboItem.ListIndex = ListIndexDe(cboItem, Trim(Right(.SelectedItem.SubItems(7), Len(.SelectedItem.SubItems(7)) - Contador)))
        txtDimensao = .SelectedItem.SubItems(2)
        txtArea = .SelectedItem.SubItems(3)
        txtDataInstalacao = .SelectedItem.SubItems(4)
        If txtValor = "" Then
            txtValor = .SelectedItem.SubItems(5)
        End If
        txtValorApagar = .SelectedItem.SubItems(6)
        .ListItems.Remove .SelectedItem.Index
        txtArea_LostFocus
        For Contador = 1 To grdAnuncio.ListItems.Count
            grdAnuncio.ListItems(Contador) = Contador
            grdAnuncio.ListItems(Contador).SubItems(8) = Left(grdAnuncio.ListItems(Contador).SubItems(8), Len(grdAnuncio.ListItems(Contador).SubItems(8)) - 2) & Format(Contador, "00")
        Next
    End With
End Sub


Private Sub grdAtividade_Click()
    Contribuinte.Seta_Atividade_Secundaria grdAtividade, cboGrupoAtividade, CboAtividade
End Sub



Private Sub grdSocio_DblClick()
    If grdSocio.SelectedItem Is Nothing Then Exit Sub
    txtCpfSocio = grdSocio.SelectedItem
    txtNomeSocio = grdSocio.SelectedItem.SubItems(1)
    txtCargoSocio = grdSocio.SelectedItem.SubItems(2)
    cboTipoLogrSocio = grdSocio.SelectedItem.SubItems(3)
    cboLogrSocio = grdSocio.SelectedItem.SubItems(4)
    txtNumSocio = grdSocio.SelectedItem.SubItems(5)
    txtCompSocio = grdSocio.SelectedItem.SubItems(6)
    cboBairroSocio = grdSocio.SelectedItem.SubItems(7)
    txtTelSocio = grdSocio.SelectedItem.SubItems(8)
    txtCidadeSocio = grdSocio.SelectedItem.SubItems(9)
    cboUFSocio = grdSocio.SelectedItem.SubItems(10)
    grdSocio.ListItems.Remove (grdSocio.SelectedItem.Index)
End Sub

Private Sub grdVeiculo_DblClick()
    If grdVeiculo.SelectedItem Is Nothing Then Exit Sub
    txtVeiculo = grdVeiculo.SelectedItem
    txtMarca = grdVeiculo.SelectedItem.SubItems(1)
    txtModelo = grdVeiculo.SelectedItem.SubItems(2)
    txtAnoFabric = grdVeiculo.SelectedItem.SubItems(3)
    txtPlaca = grdVeiculo.SelectedItem.SubItems(4)
    txtMunicipio = grdVeiculo.SelectedItem.SubItems(6)
    cboUFTransp = grdVeiculo.SelectedItem.SubItems(7)
    txtLicenca = grdVeiculo.SelectedItem.SubItems(8)
    txtChassi = grdVeiculo.SelectedItem.SubItems(5)
    Dim Pos As Integer
    Pos = InStr(grdVeiculo.SelectedItem.SubItems(9), "-")
    cboAtividadeVeiculo.ListIndex = ListIndexDe(cboAtividadeVeiculo, CStr(Trim(Right(grdVeiculo.SelectedItem.SubItems(9), Len(grdVeiculo.SelectedItem.SubItems(9)) - Pos - 1))))
    txtInicioAtividadeCarro = grdVeiculo.SelectedItem.SubItems(10)
    grdVeiculo.ListItems.Remove (grdVeiculo.SelectedItem.Index)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Transportador = Nothing
    Set Contribuinte = Nothing
    Set Endereco = Nothing
    Set Contador = Nothing
    Set Imovel = Nothing
    Set Representante = Nothing
End Sub

Private Sub txtArea_LostFocus()
 Calcula
End Sub

Private Sub txtcgc_LostFocus()
    On Error GoTo TrataErro
    If Trim(txtCgc) = "" Then Exit Sub
    
    If txtCgc = "99999999999" Or txtCgc = "999.999.999-99" Or txtCgc = "00000000000" Or txtCgc = "000.000.000-00" Then
        Util.Avisa "Valor do CPF inválido."
        txtCgc.SetFocus
    End If
    
    'TEMPORARIAMENTE(13/07/2004)
    If Len(Edita.TiraTudo(txtCgc)) = 11 Then
        txtCgc.Formato = formCPF
    ElseIf Len(Edita.TiraTudo(txtCgc)) = 14 And IsNumeric(Edita.TiraTudo(txtCgc)) Then
        txtCgc.Formato = formCGC
    Else
        Util.Informa "Cpf ou Cnpj inválido."
        tabCadastro.Tabs(1).Selected = True
        txtCgc.SetFocus
        Exit Sub
    End If
    
    If VerificaCpfCgc = False Then
        txtCgc.SetFocus
        Exit Sub
    End If
    
    If txtCgc = "" Then Exit Sub
    If Cadastro.VerificaEmpresaAntiga(txtCgc, txtRazao) = 1 Then
        If Not Util.Confirma("Já existe uma empresa cadastrada com o mesmo CNPJ/CPF. Confirma cadastro.") Then
            txtCgc.SetFocus
            Exit Sub
        End If
    End If
    txtCgc.Formato = formNenhum
    
    Exit Sub
TrataErro:
    Util.Erro Err.Description
End Sub

Private Sub txtCidade_LostFocus()
    If Trim(txtCidade) = "" Then
        txtCidade = Aplicacoes.municipio
        cboUf.Text = "MA"
        txtCep = CepCliente
    End If
End Sub

Private Sub txtCodAtividade_LostFocus()
    If atividade.Buscar(txtCodAtividade, True, 0) Then
        cboAtivServ.Text = atividade.Nome
    End If
End Sub

Private Sub txtCpfSocio_LostFocus()
    If Trim(txtCpfSocio) = "" Then Exit Sub
    If Socio.Buscar(, txtCpfSocio) Then
        With Socio
            txtCpfSocio = .Cpf
            txtNomeSocio = .Nome
            txtCargoSocio = .Cargo
            cboTipoLogrSocio = .TipoLogr
            cboLogrSocio = .Logr
            txtNumSocio = .Numero
            txtCompSocio = .Complemento
            cboBairroSocio = Bairro
            txtTelSocio = .Telefone
            txtCidadeSocio = .Cidade
            cboUFSocio = .Uf
        End With
    Else
        If Contribuinte.Buscar(, txtCpfSocio, False) Then
            With Contribuinte
                txtCpfSocio = .CgcCpf
                txtNomeSocio = .Nome
                txtCargoSocio = ""
                cboTipoLogrSocio = .Logradouro
                cboLogrSocio = .NomeLogradouro
                txtNumSocio = .Numero
                txtCompSocio = .Complemento
                cboBairroSocio = Bairro
                txtTelSocio = .FoneFax
                txtCidadeSocio = .Cidade
                cboUFSocio = .Uf
            End With
        End If
    End If
End Sub

Private Sub txtFatorMutiplicador_Change()
    If txtFatorMutiplicador = "" Or txtFatorMutiplicador = "0" Then
        txtFatorMutiplicador = 1
    End If
    Calcula
End Sub

Private Sub txtIC_LostFocus()
    If Trim(txtIc) = "" Then Exit Sub
    If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) <> 1 Then
        txtIc = Cadastro.FormataInscricao(txtIc, InscImovel)
    End If
    If Imovel.BuscarImovel(txtIc, cboTipoLogr, cboLogr, txtNum, txtComplemento, cboBairro, txtCep, txtCidade, cboUf) = False Then
        Util.Informa ("Imóvel não cadastrado.")
        cboTipoLogr.ListIndex = -1
        txtNum = ""
    End If
End Sub

Private Sub txtImRepresentante_LostFocus()
    If Trim(txtImRepresentante) = "" Then Exit Sub

    With Contribuinte
        If .Buscar(txtImRepresentante, , False) Then
            txtCpfRepresentante.Formato = formDocumento
            txtCpfRepresentante = .CgcCpf
            txtNomeRepresentante = .Nome
            cboTipoLogrRepresentante.SetarLinha .Logradouro
            cboLogrRepresentante = .NomeLogradouro
            txtNumRepresentante = .Numero
            txtComplementoRepresentante = .Complemento
            cboBairroRepresentante.SetarLinha .Bairro
            txtTelefoneRepresentante = .FoneFax
            txtCidadeRepresentante = .Cidade
            cboUfRepresentante.SetarLinha .Uf
        Else
            Util.Avisa "Contribuinte não encontrado."
        End If
    End With
End Sub

Private Sub txtchassi_lostfocus()
    Dim RetIm As String
    If Trim(txtChassi) <> "" Then
        If Transportador.VerificaChassi(txtChassi, RetIm) Then
            Informa "Chassi já cadastrado para contribuinte IM = " & RetIm & "."
            txtPlaca = ""
        End If
    End If
End Sub

Private Sub txtRazao_LostFocus()
    If Trim(txtRazao) = "" Then Exit Sub
    If Cadastro.VerificaEmpresaAntiga(txtCgc, txtRazao) = 2 Then
        If Not Util.Confirma("Já existe uma empresa cadastrada com a mesma razao social. Confirma cadastro.") Then
            txtRazao.SetFocus
            Exit Sub
        End If
    End If
    txtFantasia = txtRazao
End Sub
Private Sub Calcula()
 Dim Sql As String
    Dim rs As VSRecordset
    If cboMovimento.Text = "" Then Exit Sub
    If txtFatorMutiplicador = "" Then
        txtFatorMutiplicador = 0
    End If
    If txtArea <> "" Or txtValor <> "" Then
        'Pego os dados dos anucios..
        Sql = "Select * from Tab_Parametro_taxas where TPT_TIP_COD_IMPOSTO =  " & Bdados.Converte(cboMovimento.Coluna(0).Valor, tctexto)
        If Bdados.AbreTabela(Sql, rs) Then
            'Checo se a taxa está tabelada ou não
            If rs.Fields("TPT_TIPO") = 1 Then ' Tem Tabela
                'Faço o laço para ferificar onde o cara se enquadra.
                Do Until rs.EOF
                    If Val(txtArea) >= Val(rs.Fields("TPT_LIMITE_INFERIOR")) And Val(txtArea) <= Val(rs.Fields("TPT_LIMITE_SUPERIOR")) Then
                        If txtFatorMutiplicador <> "" Then
                            txtValor = rs.Fields("TPD_VALOR_UFM") * txtFatorMutiplicador
                            txtValorApagar = Calcula_UFM(txtValor, Converete_Real)
                        Else
                            txtValor = rs.Fields("TPD_VALOR_UFM")
                            txtValorApagar = Calcula_UFM(txtValor, Converete_Real)
                        End If
                    End If
                    rs.MoveNext
                Loop
            Else
                'Não é tabelado então eu pego o valor que foi estimado...
                If txtFatorMutiplicador <> "" Then
                    txtValor = rs.Fields("TPT_VALOR_UFM") * txtFatorMutiplicador
                    txtValorApagar = Calcula_UFM(txtValor, Converete_Real)
                Else
                    txtValor = rs.Fields("TPT_VALOR_UFM")
                    txtValorApagar = Calcula_UFM(txtValor, Converete_Real)
                End If
            End If
         Else
            Sql = "Select * from Tab_Parametro_Detalhe where TPD_TIP_COD_IMPOSTO =  " & Bdados.Converte(cboMovimento.Coluna(0).Valor, tctexto) & " and tpd_item = " & Bdados.Converte(cboItem.Coluna(3).Valor, tctexto)
            If Bdados.AbreTabela(Sql, rs) Then
                If txtFatorMutiplicador <> "" Then
                    txtValor = rs.Fields("TPD_VALOR_UFM") * txtFatorMutiplicador
                    txtValorApagar = Calcula_UFM(txtValor, Converete_Real)
                Else
                    txtValor = rs.Fields("TPD_VALOR_UFM")
                    txtValorApagar = Calcula_UFM(txtValor, Converete_Real)
                End If
            End If
        End If
    End If
    txtValor.Enabled = False
    txtValorApagar.Enabled = False
    txtFatorMutiplicador.Enabled = True
End Sub






