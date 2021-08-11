VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIP101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   81
      Top             =   7320
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   6795
         TabIndex        =   62
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Novo"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   7965
         TabIndex        =   63
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9135
         TabIndex        =   64
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin ActiveTabs.SSActiveTabs tabCad 
      Height          =   6540
      Left            =   75
      TabIndex        =   67
      Top             =   705
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   11536
      _Version        =   131082
      TabCount        =   2
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
      Tabs            =   "TCIP101.frx":0000
      Images          =   "TCIP101.frx":0092
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   6120
         Left            =   30
         TabIndex        =   69
         Top             =   30
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   10795
         _Version        =   131082
         TabGuid         =   "TCIP101.frx":0BDA
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   1395
            Left            =   75
            TabIndex        =   75
            Top             =   30
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   2461
            Altura          =   1905
            Caption         =   " Informações Gerais"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin Threed.SSPanel lblEndereco 
               Height          =   900
               Left            =   7155
               TabIndex        =   76
               Top             =   405
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   1588
               _Version        =   196610
               Font3D          =   3
               CaptionStyle    =   1
               ForeColor       =   32768
               Windowless      =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Ocupação"
               BorderWidth     =   1
               BevelOuter      =   0
               RoundedCorners  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboPedol 
               Height          =   510
               Left            =   4755
               TabIndex        =   32
               Tag             =   "42"
               Top             =   825
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   900
               Caption         =   "Pedologia"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboSit 
               Height          =   510
               Left            =   2415
               TabIndex        =   31
               Tag             =   "43"
               Top             =   825
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   900
               Caption         =   "Situação"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboPatrimonio 
               Height          =   510
               Left            =   90
               TabIndex        =   30
               Tag             =   "53"
               Top             =   825
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   900
               Caption         =   "Patrimônio"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboTopogr 
               Height          =   510
               Left            =   4755
               TabIndex        =   29
               Tag             =   "41"
               Top             =   300
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   900
               Caption         =   "Topografia"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboUtilizacao 
               Height          =   510
               Left            =   2415
               TabIndex        =   28
               Tag             =   "52"
               Top             =   300
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   900
               Caption         =   "Utilização"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboOcupacao 
               Height          =   510
               Left            =   90
               TabIndex        =   27
               Tag             =   "51"
               Top             =   300
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   900
               Caption         =   "Ocupação"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   1005
            Left            =   75
            TabIndex        =   77
            Top             =   1440
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   1773
            Altura          =   1905
            Caption         =   " Dimensões do Imóvel"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtUnids 
               Height          =   285
               Left            =   7545
               TabIndex        =   40
               Tag             =   "106"
               Top             =   660
               Width           =   2160
               _ExtentX        =   3810
               _ExtentY        =   503
               Caption         =   "Unidades no Lote"
               Text            =   ""
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtPavimento 
               Height          =   285
               Left            =   5160
               TabIndex        =   39
               Tag             =   "115"
               Top             =   660
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   503
               Caption         =   "Nº de Pavimentos"
               Text            =   ""
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtAConstr 
               Height          =   285
               Left            =   2520
               TabIndex        =   38
               Tag             =   "104"
               Top             =   660
               Width           =   2550
               _ExtentX        =   4498
               _ExtentY        =   503
               Caption         =   "Área Constr. na Unid"
               Text            =   ""
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtProf 
               Height          =   285
               Left            =   465
               TabIndex        =   37
               Tag             =   "102"
               Top             =   660
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   503
               Caption         =   "Profundidade"
               Text            =   ""
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtTTotal 
               Height          =   285
               Left            =   7860
               TabIndex        =   36
               Tag             =   "101"
               Top             =   330
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   503
               Caption         =   "Testada Total"
               Text            =   ""
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtATotal 
               Height          =   285
               Left            =   5115
               TabIndex        =   35
               Tag             =   "105"
               Top             =   330
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   503
               Caption         =   "Área Total Constr."
               Text            =   ""
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtATerreno 
               Height          =   285
               Left            =   3225
               TabIndex        =   34
               Tag             =   "103"
               Top             =   330
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   503
               Caption         =   "Área do Lote"
               Text            =   ""
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtPrin 
               Height          =   285
               Left            =   135
               TabIndex        =   33
               Tag             =   "100"
               Top             =   330
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   503
               Caption         =   "Testada Principal"
               Text            =   ""
               Restricao       =   3
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL6 
            Height          =   675
            Left            =   75
            TabIndex        =   78
            Top             =   2460
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   1191
            Altura          =   1905
            Caption         =   " Área em m²"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtArea 
               Height          =   285
               Left            =   690
               TabIndex        =   41
               Tag             =   "108"
               Top             =   330
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   503
               Caption         =   "Construída"
               Text            =   ""
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtAreaNao 
               Height          =   285
               Left            =   3000
               TabIndex        =   42
               Tag             =   "109"
               Top             =   330
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   503
               Caption         =   "Não Construída"
               Text            =   ""
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtTotalArea 
               Height          =   285
               Left            =   6270
               TabIndex        =   43
               Tag             =   "110"
               Top             =   330
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               Caption         =   "Total"
               Text            =   ""
               Restricao       =   3
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL7 
            Height          =   1890
            Left            =   75
            TabIndex        =   79
            Top             =   3150
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   3334
            Altura          =   1905
            Caption         =   " Informações Sobre a Edificação"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboInstElet 
               Height          =   510
               Left            =   4995
               TabIndex        =   54
               Tag             =   "50"
               Top             =   1320
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Instalação Elétrica"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboParede 
               Height          =   510
               Left            =   2535
               TabIndex        =   53
               Tag             =   "48"
               Top             =   1320
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Paredes"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboPosicao 
               Height          =   510
               Left            =   90
               TabIndex        =   52
               Tag             =   "56"
               Top             =   1320
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Posicionamento"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboConservacao 
               Height          =   510
               Left            =   7440
               TabIndex        =   51
               Tag             =   "57"
               Top             =   810
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Estado de Conservação"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboInstSanit 
               Height          =   510
               Left            =   4995
               TabIndex        =   50
               Tag             =   "49"
               Top             =   810
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Instalação Sanitária"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboCobert 
               Height          =   510
               Left            =   2535
               TabIndex        =   49
               Tag             =   "47"
               Top             =   810
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Cobertura"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboPiso 
               Height          =   510
               Left            =   7440
               TabIndex        =   47
               Tag             =   "54"
               Top             =   300
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Piso"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboForro 
               Height          =   510
               Left            =   4995
               TabIndex        =   46
               Tag             =   "55"
               Top             =   300
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Forro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboEstrut 
               Height          =   510
               Left            =   2535
               TabIndex        =   45
               Tag             =   "46"
               Top             =   300
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Estrutura"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboTipConstr 
               Height          =   510
               Left            =   90
               TabIndex        =   44
               Tag             =   "45"
               Top             =   300
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Tipo de Construção"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboFachada 
               Height          =   510
               Left            =   90
               TabIndex        =   48
               Tag             =   "58"
               Top             =   810
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   900
               Caption         =   "Alinhamento"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL8 
            Height          =   1020
            Left            =   75
            TabIndex        =   80
            Top             =   5055
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   1799
            Altura          =   1905
            Caption         =   " Serviços Públicos a Disposição"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtServCalc 
               Height          =   285
               Left            =   5295
               TabIndex        =   60
               Tag             =   "113"
               Top             =   675
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   503
               Caption         =   "Testada Servida"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   3
            End
            Begin VTOcx.cboVISUAL cboLixo 
               Height          =   315
               Left            =   7695
               TabIndex        =   61
               Tag             =   "40"
               Top             =   315
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   556
               Caption         =   "Coleta de Lixo"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboCalcam 
               Height          =   315
               Left            =   5100
               TabIndex        =   59
               Tag             =   "39"
               Top             =   315
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   556
               Caption         =   "Cons. Calçamento"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboLimp 
               Height          =   315
               Left            =   2715
               TabIndex        =   57
               Tag             =   "38"
               Top             =   315
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   556
               Caption         =   "Limpeza Pública"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboIlumPub 
               Height          =   315
               Left            =   90
               TabIndex        =   55
               Tag             =   "37"
               Top             =   315
               Width           =   2460
               _ExtentX        =   4339
               _ExtentY        =   556
               Caption         =   "Iluminição Pública"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtServLimp 
               Height          =   285
               Left            =   2700
               TabIndex        =   58
               Tag             =   "112"
               Top             =   675
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   503
               Caption         =   "Testada Servida"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtServIlum 
               Height          =   285
               Left            =   240
               TabIndex        =   56
               Tag             =   "111"
               Top             =   675
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   503
               Caption         =   "Testada Servida"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   3
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6120
         Left            =   30
         TabIndex        =   68
         Top             =   30
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   10795
         _Version        =   131082
         TabGuid         =   "TCIP101.frx":0C02
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1425
            Left            =   60
            TabIndex        =   70
            Top             =   30
            Width           =   9960
            _ExtentX        =   17568
            _ExtentY        =   2514
            Altura          =   1905
            Caption         =   " Imóvel"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtCepIm 
               Height          =   480
               Left            =   8520
               TabIndex        =   10
               Top             =   855
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   847
               Caption         =   "CEP"
               Text            =   ""
               Formato         =   4
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtLoteamento 
               Height          =   480
               Left            =   5625
               TabIndex        =   7
               Top             =   855
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   847
               Caption         =   "Loteamento"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtQuadra 
               Height          =   480
               Left            =   3825
               TabIndex        =   5
               Top             =   855
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   847
               Caption         =   "Quadra"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtLote 
               Height          =   480
               Left            =   6750
               TabIndex        =   8
               Top             =   855
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   847
               Caption         =   "Lote"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtSecao 
               Height          =   480
               Left            =   7650
               TabIndex        =   9
               Top             =   855
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   847
               Caption         =   "Seção"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboBairro 
               Height          =   510
               Left            =   105
               TabIndex        =   4
               Tag             =   "Bairro"
               Top             =   825
               Width           =   3675
               _ExtentX        =   6482
               _ExtentY        =   900
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtComplemento 
               Height          =   480
               Left            =   7095
               TabIndex        =   3
               Top             =   315
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   847
               Caption         =   "Complemento"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtNumero 
               Height          =   480
               Left            =   4725
               TabIndex        =   6
               Tag             =   "Nº"
               Top             =   855
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   847
               Caption         =   "Nº"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboLogr 
               Height          =   315
               Left            =   3795
               TabIndex        =   2
               Tag             =   "Logradouro"
               Top             =   480
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboTipoLogr 
               Height          =   510
               Left            =   1965
               TabIndex        =   1
               Tag             =   "Logradouro"
               Top             =   285
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   900
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtInscAnterior 
               Height          =   480
               Left            =   120
               TabIndex        =   0
               Top             =   315
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   847
               Caption         =   "Insc. Anterior"
               Text            =   ""
               Formato         =   7
               AlinhamentoRotulo=   1
               AgruparValores  =   0   'False
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   1860
            Left            =   60
            TabIndex        =   71
            Top             =   1485
            Width           =   9960
            _ExtentX        =   17568
            _ExtentY        =   3281
            Altura          =   1905
            Caption         =   " Contribuinte"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin Threed.SSCheck chkEndereco 
               Height          =   225
               Left            =   8640
               TabIndex        =   11
               Top             =   45
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   397
               _Version        =   196610
               ForeColor       =   16777215
               BackColor       =   32768
               Caption         =   "Mesmo Lote"
            End
            Begin VTOcx.cmdVISUAL cmdBuscarContrib 
               Height          =   330
               Left            =   1560
               TabIndex        =   13
               Top             =   450
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   582
               Caption         =   ""
               Acao            =   5
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cmdVISUAL cmdNovoContr 
               Height          =   330
               Left            =   1965
               TabIndex        =   14
               Top             =   450
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   582
               Caption         =   ""
               Acao            =   6
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.txtVISUAL txtIm 
               Height          =   480
               Left            =   105
               TabIndex        =   12
               Top             =   285
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   847
               Caption         =   "Insc. Municipal"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               Mascara         =   "00000000-00"
            End
            Begin VTOcx.txtVISUAL txtNomeContrib 
               Height          =   480
               Left            =   2385
               TabIndex        =   15
               Tag             =   "Nome"
               Top             =   285
               Width           =   7545
               _ExtentX        =   13309
               _ExtentY        =   847
               Caption         =   "Nome"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboTipoLogrContrib 
               Height          =   315
               Left            =   105
               TabIndex        =   16
               Tag             =   "Logradouro"
               Top             =   960
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNomeLogrContrib 
               Height          =   285
               Left            =   1725
               TabIndex        =   17
               Tag             =   "Logradouro"
               Top             =   975
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   503
               Caption         =   ""
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtNumeroContrib 
               Height          =   480
               Left            =   4470
               TabIndex        =   18
               Tag             =   "Nº"
               Top             =   780
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   847
               Caption         =   "Nº"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCompContrib 
               Height          =   480
               Left            =   5115
               TabIndex        =   19
               Top             =   780
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   847
               Caption         =   "Complemento"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtBairroContrib 
               Height          =   480
               Left            =   7425
               TabIndex        =   20
               Tag             =   "Bairro"
               Top             =   780
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   847
               Caption         =   "Bairro"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCep 
               Height          =   480
               Left            =   105
               TabIndex        =   21
               Top             =   1290
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   847
               Caption         =   "CEP"
               Text            =   ""
               Formato         =   4
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtMunic 
               Height          =   480
               Left            =   1725
               TabIndex        =   22
               Tag             =   "Município"
               Top             =   1290
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   847
               Caption         =   "Município"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboUF 
               Height          =   315
               Left            =   5130
               TabIndex        =   23
               Tag             =   "UF"
               Top             =   1470
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VB.Label Label2 
               Caption         =   "Logradouro"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   120
               TabIndex        =   73
               Top             =   765
               Width           =   1635
            End
            Begin VB.Label Label4 
               Caption         =   "UF"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   5130
               TabIndex        =   72
               Top             =   1275
               Width           =   435
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   780
            Left            =   60
            TabIndex        =   74
            Top             =   3360
            Width           =   9960
            _ExtentX        =   17568
            _ExtentY        =   1376
            Altura          =   1905
            Caption         =   " Detalhes"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboAforado 
               Height          =   315
               Left            =   5445
               TabIndex        =   25
               Tag             =   "Aforado"
               Top             =   360
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   556
               Caption         =   "Aforado"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtAnoAq 
               Height          =   285
               Left            =   1905
               TabIndex        =   24
               Tag             =   "Ano de Aquisição"
               Top             =   360
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   503
               Caption         =   "Ano de Aquisição"
               Text            =   ""
               Restricao       =   2
               MaxLen          =   4
               MinLen          =   4
            End
         End
         Begin VTOcx.grdVISUAL grdContribuinte 
            Height          =   2010
            Left            =   60
            TabIndex        =   26
            Top             =   4170
            Width           =   9960
            _ExtentX        =   17568
            _ExtentY        =   3545
            CorBorda        =   16711680
            Caption         =   "Contribuintes"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
      End
   End
   Begin VB.TextBox txtFatorFixo 
      Height          =   285
      Left            =   8640
      TabIndex        =   65
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   4560
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1138
      Icone           =   "TCIP101.frx":0C2A
   End
End
Attribute VB_Name = "TCIP101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As VSImposto
Dim Endereco As cEndereco
Dim Contribuinte As cContribuinte
Dim Imovel As cImovel
Dim NovoContrib As Boolean
Dim Sql As String
Sub SetaComboTSU()
    txtUnids = "1"
    txtCepIm = Trim(Temp.PegaParametro(Bdados, "CEP"))
End Sub

Sub CalculaAreaTerreno()

    Dim Constr As Single
    Dim Terreno As Single
    On Error Resume Next
    
    Constr = IIf(Trim(txtPrin) = "", 0, CSng(txtPrin))
    Terreno = IIf(Trim(txtProf) = "", 0, CSng(txtProf))
    txtATerreno = CStr(Constr * Terreno)
End Sub

Function TotalProva(Valor As String) As String
    Static Total As Double
    If Trim(Valor) = "" Then Valor = "0"
    Total = CDbl(Valor) + Total
    TotalProva = Total
End Function

Private Sub VerificaTestada(CboRef As Object, TxtRef As Object, Codigo As Byte)
    If CboRef.Text = "SIM" Then
        TxtRef.Tag = Codigo
        TxtRef.TabStop = True
        TxtRef.Enabled = True
        TxtRef.Text = txtPrin
        TxtRef.SetFocus
    Else
        TxtRef.Tag = ""
        TxtRef.Text = ""
        TxtRef.TabStop = False
        TxtRef.Enabled = False
    End If
End Sub

Public Sub HabilitaCaixa(Status As Boolean)
    txtIm.Enabled = Not Status
    txtNomeContrib.Enabled = Status
    cboTipoLogrContrib.Enabled = Status
    txtNomeLogrContrib.Enabled = Status
    txtNumeroContrib.Enabled = Status
    txtCompContrib.Enabled = Status
    txtBairroContrib.Enabled = Status
    txtCep.Enabled = Status
    txtMunic.Enabled = Status
    cboUF.Enabled = Status
    If Status Then
        txtIm = ""
        txtNomeContrib = ""
        cboTipoLogrContrib.ListIndex = -1
        txtNomeLogrContrib = ""
        txtNumeroContrib = ""
        txtCompContrib = ""
        txtBairroContrib = ""
        txtCep = ""
        txtMunic = ""
        cboUF.ListIndex = -1
    End If
End Sub
Sub CalculaArea()
    Dim Constr As Single
    Dim Terreno As Single
    On Error Resume Next
    
    Constr = IIf(Trim(txtArea) = "", 0, CSng(txtArea))
    Terreno = IIf(Trim(txtAreaNao) = "", 0, CSng(txtAreaNao))
    txtTotalArea = CStr(Constr + Terreno)
End Sub

Private Sub cboBairro_Click()
    lblEndereco = cboTipoLogr & " " & cboLogr & " " & cboBairro
End Sub

Private Sub cboCalcam_LostFocus()
    If cboCalcam.ListIndex <> -1 Then Call VerificaTestada(cboCalcam, txtServCalc, 113)
End Sub

Private Sub cboIlumPub_LostFocus()
    If cboIlumPub.ListIndex <> -1 Then Call VerificaTestada(cboIlumPub, txtServIlum, CByte(111))
End Sub

Private Sub cboLimp_LostFocus()
    If cboLimp.ListIndex <> -1 Then Call VerificaTestada(cboLimp, txtServLimp, 112)
End Sub

Private Sub cboLogr_Click()
    Call cadastro.BuscaLogradouro(Bairro, cboLogr, cboBairro)
    lblEndereco = cboTipoLogr & " " & cboLogr & " " & cboBairro
End Sub

Private Sub cboTipoLogr_Click()
    Call cadastro.BuscaLogradouro(Rua, cboTipoLogr, cboLogr)
    lblEndereco = cboTipoLogr & " " & cboLogr & " " & cboBairro
End Sub

Private Sub chkEndereco_Click(Value As Integer)
    On Error Resume Next
    If Value Then
        cboTipoLogrContrib = cboTipoLogr.Text
        txtNomeLogrContrib = cboLogr
        txtNumeroContrib = txtNumero
        txtCompContrib = txtComplemento
        txtBairroContrib = cboBairro
        txtCep = txtCepIm
        txtNomeContrib.SetFocus
    Else
        cboTipoLogrContrib.ListIndex = -1
        txtNomeLogrContrib = ""
        txtNumeroContrib = ""
        txtCompContrib = ""
        txtBairroContrib = ""
        cboTipoLogrContrib.SetFocus
    End If
End Sub

Private Sub cmdBuscarContrib_Click()
    If Contribuinte.PreencherGrd(grdContribuinte, txtNomeContrib) = False Then
        Util.Avisa "Nenhum contribuinte encontrado."
    End If
End Sub

Private Sub cmdNovoContr_Click()
    NovoContrib = True
    txtIm = ""
    Call HabilitaCaixa(True)
    txtNomeContrib.SetFocus
End Sub

Private Sub cmdSalvar_Click()
    On Error Resume Next
    Dim Valores As String
    Dim Campos As String
    Dim DataReab As Date
    Dim RsAux As VSRecordset
    Dim Rs As VSRecordset
    Dim InscricaoMunicipal As String
    Dim InscricaoCadastral As String
    Dim CodLogr As String
    Dim CodBairr As Long
    Dim DtVenc As String
    Dim SitCadastral As String
    'Verifica se tem area construida, se nao tiver, nao se deve preencher os componentes, caso contrario deve.
    If txtArea = 0 Then
        cboTipConstr.Tag = ""
        cboFachada.Tag = ""
        cboPosicao.Tag = ""
        cboEstrut.Tag = ""
        cboCobert.Tag = ""
        cboParede.Tag = ""
        cboForro.Tag = ""
        cboInstSanit.Tag = ""
        cboInstElet.Tag = ""
        cboPiso.Tag = ""
        cboConservacao.Tag = ""
    Else
        cboTipConstr.Tag = "45"
        cboFachada.Tag = "58"
        cboPosicao.Tag = "56"
        cboEstrut.Tag = "46"
        cboCobert.Tag = "47"
        cboParede.Tag = "48"
        cboForro.Tag = "55"
        cboInstSanit.Tag = "49"
        cboInstElet.Tag = "50"
        cboPiso.Tag = "54"
        cboConservacao.Tag = "57"
    End If
        If Edita.CriticaCampos(Me) Then
            txtFatorFixo.Tag = "1000"
            Screen.MousePointer = 11
            CodLogr = CStr(cadastro.PegaCodLogr(cboBairro.Text, cboTipoLogr.Text, cboLogr.Text))
            CodBairr = Endereco.BuscaBairro(cboBairro)
            'Buscando as Inscricoes
            InscricaoCadastral = cadastro.GeraInscCadastral(Right(Date, 1), 31, 1)
            If NovoContrib Then
                InscricaoMunicipal = cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
            Else
                InscricaoMunicipal = txtIm
            End If
            
            If Not cadastro.ContribuinteHabilitado(InscricaoMunicipal, SitCadastral) Then
                Call Util.Avisa("O Contribuinte está " & SitCadastral & " e não pode adquirir novos imóveis.")
                Screen.MousePointer = 0
                txtFatorFixo.Tag = ""
                Exit Sub
            End If
            'Vou gravar o Contribuinte
            If NovoContrib Then
                If Contribuinte.GravarContribuinte(InscricaoMunicipal, Trim(txtNomeContrib), _
                cboTipoLogrContrib, Trim(txtNomeLogrContrib), Trim(txtNumeroContrib), _
                Trim(txtCompContrib), Trim(txtBairroContrib), Trim(txtCep), _
                Trim(txtMunic), cboUF) Then
                    NovoContrib = True
                End If
            End If
            If Imovel.Buscar(InscricaoCadastral) = False Then
                With Imovel
                    .Im = InscricaoMunicipal
                    .CodLogradouro = CodLogr
                    .Numero = txtNumero
                    .Complemento = txtComplemento
                    .Loteamento = txtLoteamento
                    .Secao = Trim(txtSecao)
                    .Quadra = Trim(txtQuadra)
                    .Lote = Trim(txtLote)
                    .CEP = Trim(txtCepIm)
                    .AnoAquisicao = Trim(txtAnoAq)
                    .Aforado = Left(cboAforado, 1)
                    .IcAnterior = Util.Nvl(txtInscAnterior, 0)
                    .CodBairro = CodBairr
                    .Gravar (InscricaoCadastral)
                End With
            End If
            txtIm = InscricaoMunicipal
            'gravo os componentes
            Call cadastro.GravaComponente(InscricaoCadastral, "0", txtArea, txtAreaNao, Me)
            Call Util.Informa("Registro gravado com sucesso.")
            txtFatorFixo.Tag = ""
            Call Util.Informa("Inscricão Cadastral Gerada Nº: " & InscricaoCadastral)
            If NovoContrib Then Call Util.Informa("Inscricão Municipal Gerada Nº: " & InscricaoMunicipal)
            NovoContrib = True
            cmdNovo_Click
            DoEvents
            cboTipoLogr.Enabled = True
            txtInscAnterior.SetFocus
            Screen.MousePointer = 0
            Call HabilitaCaixa(True)
        End If
End Sub

Private Sub cmdNovo_Click()
    NovoContrib = True
    Call Edita.LimpaCampos(Me)
    cboTipoLogr.Enabled = True
    tabCad.Tabs(1).Selected = True
    txtInscAnterior.SetFocus
    grdContribuinte.ListItems.Clear
    SetaComboTSU
End Sub

Private Sub cmdSair_Click()
    NovoContrib = True
    Unload Me
End Sub

Private Sub Form_Load()
        
    Dim Controle As Control
    Dim i As Byte
    Dim Rs As VSRecordset
    Dim Sql As String
    '*********Setando classes
    Set cadastro = New VSImposto
    Set Endereco = New cEndereco
    Set Contribuinte = New cContribuinte
    Set Imovel = New cImovel
    '**********Preenchendo as combos
    Endereco.PreencherComboLogr cboLogr
    Endereco.PreencherComboTipoLogr cboTipoLogr
    Endereco.PreencherComboBairro cboBairro
    Endereco.PreencherComboTipoLogr cboTipoLogrContrib
    cboUF.PreencherGeral Bdados, "UF"
    cboAforado.PreencherGeral Bdados, "SIM OU NÃO"
    On Error Resume Next
    For Each Controle In Controls
        If IsNumeric(Controle.Tag) Then
            If Val(Controle.Tag) < 100 Then
                Imovel.preenchercomponente Controle, Controle.Tag
            End If
        End If
    Next
    On Error GoTo 0

    Screen.MousePointer = 0
    '*********setando o cabecalho e rodape
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    NovoContrib = True

    SetaComboTSU
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Endereco = Nothing
    Set Contribuinte = Nothing
    Set Imovel = Nothing
End Sub

Private Sub grdContribuinte_DblClick()
    txtIm = grdContribuinte.SelectedItem
    Call txtIm_LostFocus
End Sub


Private Sub txtArea_Change()
    CalculaArea
End Sub

Private Sub txtAreaNao_Change()
    CalculaArea
End Sub

Private Sub txtAterreno_Change()
    Dim AreaTerreno As Double
    Dim AreaConst As Double
    
    AreaTerreno = Util.Nvl(txtATerreno, 0)
    AreaConst = Util.Nvl(txtArea, 0)
    txtAreaNao = AreaTerreno - AreaConst
End Sub

Private Sub txtAtotal_Change()
    txtArea = txtATotal
    Dim AreaTerreno As Double
    Dim AreaConst As Double
    
    AreaTerreno = Util.Nvl(txtATerreno, 0)
    AreaConst = Util.Nvl(txtArea, 0)
    txtAreaNao = AreaTerreno - AreaConst
End Sub

Private Sub txtIm_LostFocus()
    Dim Rs As VSRecordset
    Dim NomeContrib As String, TipoLogrContr As String, LogrContr As String, NumeroContr As String, CompContri As String, _
           BairroContr As String, CepContr As String, MunicContr As String, UFContr As String
    
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIm) <> "" Then
        If Contribuinte.BuscarContribuinte(txtIm, NomeContrib, TipoLogrContr, LogrContr, NumeroContr, CompContri, BairroContr, CepContr, MunicContr, UFContr) Then
            txtNomeContrib = NomeContrib
            cboTipoLogrContrib.ListIndex = cadastro.BuscaCodLogr("" & TipoLogrContr)
            txtNomeLogrContrib = LogrContr
            txtNumeroContrib = NumeroContr
            txtCompContrib = CompContri
            txtBairroContrib = BairroContr
            txtCep = CepContr
            txtMunic = MunicContr
            cboUF = UFContr
            NovoContrib = False
            HabilitaCaixa False
            txtIm.Enabled = False
            txtAnoAq.SetFocus
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtIm.Enabled = True
            txtIm.SetFocus
            NovoContrib = True
            HabilitaCaixa True
            txtIm = ""
            Exit Sub
        End If
    End If
End Sub

Private Sub txtMunic_LostFocus()
    If Trim(txtMunic) = "" Then txtMunic = Aplicacoes.municipio
End Sub

Private Sub txtprin_Change()
    txtTTotal = txtPrin
    CalculaAreaTerreno
End Sub

Private Sub txtProf_Change()
    CalculaAreaTerreno
End Sub

