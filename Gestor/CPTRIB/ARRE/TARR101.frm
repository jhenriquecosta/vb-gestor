VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TARR101 
   Caption         =   "Gestão de Arrecadação - 20171019"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   7455
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   13150
      _Version        =   131082
      TabCount        =   5
      Tabs            =   "TARR101.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   7065
         Left            =   -99969
         TabIndex        =   95
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   12462
         _Version        =   131082
         TabGuid         =   "TARR101.frx":0114
         Begin VTOcx.fraVISUAL fraVISUAL6 
            Height          =   1095
            Left            =   120
            TabIndex        =   105
            Top             =   4920
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   1931
            Altura          =   1905
            Caption         =   " OBSERVAÇÕES"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483644
            Begin VTOcx.cmdVISUAL cmdTransferir 
               Height          =   495
               Left            =   8400
               TabIndex        =   108
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   873
               Caption         =   "Gravar"
               Acao            =   3
            End
            Begin VTOcx.cmdVISUAL cmdVISUAL8 
               Height          =   495
               Left            =   9480
               TabIndex        =   107
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   873
               Caption         =   "Sair"
               Acao            =   7
            End
            Begin VTOcx.txtVISUAL txtObsTransferencia 
               Height          =   765
               Left            =   120
               TabIndex        =   106
               Top             =   280
               Width           =   8265
               _ExtentX        =   14579
               _ExtentY        =   1349
               Caption         =   "Observação"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   975
            Left            =   120
            TabIndex        =   101
            Top             =   3840
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   1720
            Altura          =   1905
            Caption         =   " INSCRIÇÃO DESTINO"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483644
            Begin VTOcx.txtVISUAL txtDestinoNome 
               Height          =   495
               Left            =   2640
               TabIndex        =   104
               Top             =   360
               Width           =   7845
               _ExtentX        =   13838
               _ExtentY        =   873
               Caption         =   "Nome - Endereço"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cmdVISUAL cmdVISUAL7 
               Height          =   315
               Left            =   2200
               TabIndex        =   103
               TabStop         =   0   'False
               Top             =   550
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtDestinoInscricao 
               Height          =   500
               Left            =   120
               TabIndex        =   102
               Top             =   360
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   873
               Caption         =   "Destino"
               Text            =   ""
               Enabled         =   0   'False
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   975
            Left            =   120
            TabIndex        =   97
            Top             =   720
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   1720
            Altura          =   1905
            Caption         =   " INSCRIÇÃO ORIGEM (CONTRIBUINTE)"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483644
            Begin VB.Frame Frame5 
               Height          =   495
               Left            =   8280
               TabIndex        =   110
               Top             =   360
               Width           =   2175
               Begin VB.OptionButton optImposto 
                  Caption         =   "OUTROS"
                  Height          =   195
                  Index           =   1
                  Left            =   960
                  TabIndex        =   112
                  Top             =   200
                  Width           =   1095
               End
               Begin VB.OptionButton optImposto 
                  Caption         =   "IPTU"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   111
                  Top             =   200
                  Value           =   -1  'True
                  Width           =   735
               End
            End
            Begin VTOcx.txtVISUAL txtOrigemNome 
               Height          =   495
               Left            =   2640
               TabIndex        =   100
               Top             =   360
               Width           =   5565
               _ExtentX        =   9816
               _ExtentY        =   873
               Caption         =   "Contribuinte"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cmdVISUAL cmdOrigemInscricao 
               Height          =   315
               Left            =   2200
               TabIndex        =   99
               TabStop         =   0   'False
               Top             =   550
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtOrigemInscricao 
               Height          =   500
               Left            =   120
               TabIndex        =   98
               Top             =   360
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   873
               Caption         =   "Origem"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               EnterEqvTab     =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
         End
         Begin Cabecalho.cabVISUAL cabVISUAL5 
            Height          =   645
            Left            =   0
            TabIndex        =   96
            Top             =   0
            Width           =   10755
            _ExtentX        =   18971
            _ExtentY        =   1138
            Formulario      =   "TRANSFERÊNCIA DE DÉBITOS"
            Descricao       =   "Transfere os débitos entre contribuintes"
            Icone           =   "TARR101.frx":013C
         End
         Begin VTOcx.grdVISUAL lstDebitos 
            Height          =   2220
            Left            =   120
            TabIndex        =   109
            Top             =   1800
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   3916
            CorBorda        =   16711680
            Caption         =   "Débitos"
            CorTitulo       =   16711680
            CorCaption      =   -2147483634
            CheckBox        =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   7065
         Left            =   -99969
         TabIndex        =   80
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   12462
         _Version        =   131082
         TabGuid         =   "TARR101.frx":0456
         Begin VB.Frame Frame4 
            Caption         =   "Informe o DAM para Alteração"
            Height          =   1455
            Left            =   120
            TabIndex        =   84
            Top             =   720
            Width           =   10455
            Begin VTOcx.cmdVISUAL cmdBuscarDam 
               Height          =   375
               Left            =   3000
               TabIndex        =   64
               Top             =   1005
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   661
               Caption         =   "&Buscar"
               Acao            =   5
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.txtVISUAL txtDAM 
               Height          =   540
               Left            =   1560
               TabIndex        =   63
               Top             =   840
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   953
               Caption         =   "Número DAM"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cmdVISUAL cmdAlterarDam 
               Height          =   375
               Left            =   4440
               TabIndex        =   65
               Top             =   1005
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   661
               Caption         =   "&Alterar"
               Acao            =   3
               CorBorda        =   8421504
               CorFrente       =   16384
               CorFundo        =   16777088
            End
            Begin VTOcx.cmdVISUAL cmdVISUAL5 
               CausesValidation=   0   'False
               Height          =   375
               Left            =   9360
               TabIndex        =   86
               Top             =   960
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               Caption         =   "Sai&r"
               Acao            =   7
               CorBorda        =   16711680
               CorFundo        =   16777152
            End
            Begin VTOcx.txtVISUAL txtNossoNumero 
               Height          =   540
               Left            =   120
               TabIndex        =   62
               Top             =   840
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   953
               Caption         =   "Banco Número"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtIm 
               Height          =   555
               Left            =   120
               TabIndex        =   90
               Top             =   240
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   979
               Caption         =   "Inscricão"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
               Height          =   315
               Left            =   3000
               TabIndex        =   91
               TabStop         =   0   'False
               Top             =   480
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
            End
            Begin VB.Label lblValorReal 
               Caption         =   "R$ Real:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   6000
               TabIndex        =   92
               Top             =   1080
               Width           =   3255
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   3060
            Left            =   120
            TabIndex        =   82
            Top             =   3840
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   5398
            Altura          =   1905
            Caption         =   " Dados da Obrigração"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL cmdAlteraTributo 
               Height          =   300
               Left            =   7920
               TabIndex        =   94
               TabStop         =   0   'False
               Top             =   2520
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   529
               Caption         =   ""
               Acao            =   1
               CorFundo        =   12648447
            End
            Begin VTOcx.cboVISUAL cboImposto 
               Height          =   315
               Left            =   120
               TabIndex        =   93
               Tag             =   "Tributo"
               Top             =   2520
               Width           =   7770
               _ExtentX        =   13705
               _ExtentY        =   556
               Caption         =   "Tributo"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCredito 
               Height          =   495
               Left            =   8520
               TabIndex        =   89
               Tag             =   "Data Vencimento"
               Top             =   2280
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   873
               Caption         =   "Data Crédito"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VB.CheckBox chkComprovante 
               Alignment       =   1  'Right Justify
               Caption         =   "Comprovante ?"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   8520
               TabIndex        =   88
               Top             =   1920
               Width           =   1695
            End
            Begin VTOcx.txtVISUAL txtObservacao 
               Height          =   765
               Left            =   120
               TabIndex        =   79
               Top             =   1680
               Width           =   8265
               _ExtentX        =   14579
               _ExtentY        =   1349
               Caption         =   "Observação"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtMultaMargem 
               Height          =   540
               Left            =   3840
               TabIndex        =   74
               Top             =   1080
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   953
               Caption         =   " %"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   2
               CorRotulo       =   192
               ValorPadrao     =   "0"
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtJurosMargem 
               Height          =   540
               Left            =   2040
               TabIndex        =   72
               Top             =   1080
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   953
               Caption         =   " %"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   2
               CorRotulo       =   192
               ValorPadrao     =   "0"
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cboVISUAL cboStatus 
               Height          =   510
               Left            =   7320
               TabIndex        =   70
               Tag             =   "Status"
               Top             =   360
               Width           =   2880
               _ExtentX        =   5080
               _ExtentY        =   900
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtInscricao 
               Height          =   510
               Left            =   90
               TabIndex        =   66
               Top             =   420
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   900
               Caption         =   "Contribuinte"
               Text            =   ""
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtPeriodo 
               Height          =   510
               Left            =   2070
               TabIndex        =   67
               Tag             =   "Período"
               Top             =   420
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   900
               Caption         =   "Periodo"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtMulta 
               Height          =   540
               Left            =   4440
               TabIndex        =   75
               Top             =   1080
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   953
               Caption         =   "Multa"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtValorOriginal 
               Height          =   540
               Left            =   120
               TabIndex        =   71
               Tag             =   "Valor Obrigação"
               Top             =   1080
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   953
               Caption         =   "Valor Original"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtJuros 
               Height          =   540
               Left            =   2640
               TabIndex        =   73
               Top             =   1080
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   953
               Caption         =   "Juros"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtTributo 
               Height          =   510
               Left            =   3840
               TabIndex        =   68
               Top             =   420
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   900
               Caption         =   "Tributo"
               Text            =   ""
               Enabled         =   0   'False
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtVence 
               Height          =   510
               Left            =   5880
               TabIndex        =   69
               Tag             =   "Data Vencimento"
               Top             =   420
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   900
               Caption         =   "Vencimento"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtCorrecao 
               Height          =   540
               Left            =   5880
               TabIndex        =   76
               Top             =   1080
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   953
               Caption         =   "Correcão"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtDesconto 
               Height          =   540
               Left            =   7320
               TabIndex        =   77
               Top             =   1080
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   953
               Caption         =   "Descto (%)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               CorRotulo       =   192
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtParcelamento 
               Height          =   465
               Left            =   240
               TabIndex        =   81
               Top             =   1920
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   820
               Caption         =   "Cod. Parcelamento"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtDescontoReal 
               Height          =   540
               Left            =   8400
               TabIndex        =   78
               Top             =   1080
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   953
               Caption         =   "Desconto R$"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               Requerido       =   0   'False
               AlinhamentoRotulo=   1
               AutoTAB         =   -1  'True
            End
         End
         Begin Cabecalho.cabVISUAL cabVISUAL4 
            Height          =   645
            Left            =   120
            TabIndex        =   83
            Top             =   0
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   1138
            Formulario      =   "ALTERAÇÃO DE DAM"
            Descricao       =   "Alterar Valores do DAM"
            Icone           =   "TARR101.frx":047E
         End
         Begin VTOcx.grdVISUAL lstObrig 
            Height          =   1770
            Left            =   120
            TabIndex        =   85
            Top             =   2280
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   3122
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   192
            OcultarRodape   =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   7065
         Left            =   -99969
         TabIndex        =   49
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   12462
         _Version        =   131082
         TabGuid         =   "TARR101.frx":0798
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   855
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   1508
            Altura          =   1905
            Caption         =   " FILTRO"
            CorTexto        =   0
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Begin VB.TextBox txtArquivo 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4440
               TabIndex        =   54
               Top             =   360
               Width           =   1260
            End
            Begin VTOcx.cboVISUAL cboDados 
               Height          =   315
               Left            =   2640
               TabIndex        =   53
               Top             =   360
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               Caption         =   "Opção"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtArquivo1 
               Height          =   315
               Left            =   5760
               TabIndex        =   55
               Top             =   360
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
            End
            Begin VTOcx.cboVISUAL cboOrdem 
               Height          =   315
               Left            =   8040
               TabIndex        =   61
               Top             =   360
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               Caption         =   "ORDEM"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cmdVISUAL cmdImprimir 
               Height          =   375
               Left            =   10080
               TabIndex        =   58
               Top             =   360
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               Caption         =   ""
               Acao            =   4
               CorFundo        =   16777088
            End
            Begin VTOcx.cmdVISUAL cmdBuscar 
               Height          =   375
               Left            =   7560
               TabIndex        =   57
               Top             =   360
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               Caption         =   ""
               Acao            =   5
               CorFundo        =   16777088
            End
            Begin VTOcx.txtVISUAL txtAno 
               Height          =   315
               Left            =   6840
               TabIndex        =   56
               Top             =   360
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   556
               Caption         =   "/"
               Text            =   ""
            End
            Begin VTOcx.cboVISUAL cboTIPO 
               Height          =   315
               Left            =   120
               TabIndex        =   52
               Top             =   360
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               Caption         =   "TIPO"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
         End
         Begin Cabecalho.cabVISUAL cabVISUAL3 
            Height          =   645
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   1138
            Formulario      =   "GERENCIAMENTO DE PAGAMENTOS"
            Descricao       =   "Gerenciamento das REMESSAS e RETORNOS gerados"
            Icone           =   "TARR101.frx":07C0
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL4 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   9360
            TabIndex        =   59
            Top             =   6600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "Sai&r"
            Acao            =   7
            CorFundo        =   16777088
         End
         Begin VTOcx.grdVISUAL grdGerencial 
            Height          =   4860
            Left            =   120
            TabIndex        =   60
            Top             =   1920
            Width           =   10485
            _ExtentX        =   18494
            _ExtentY        =   8573
            Caption         =   "Dados"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CheckBox        =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   7065
         Left            =   -99969
         TabIndex        =   23
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   12462
         _Version        =   131082
         TabGuid         =   "TARR101.frx":0ADA
         Begin VTOcx.grdVISUAL grdRETORNO 
            Height          =   4380
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   10485
            _ExtentX        =   18494
            _ExtentY        =   7726
            Caption         =   "Dados"
            CheckBox        =   -1  'True
         End
         Begin Cabecalho.cabVISUAL cabVISUAL2 
            Height          =   645
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   1138
            Formulario      =   "RETORNO DE DOCUMENTOS BANCARIO"
            Descricao       =   "Selecione o diretorio para Recepção dos Pagamentos"
            Icone           =   "TARR101.frx":0B02
         End
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1065
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   1879
            Altura          =   1905
            Caption         =   " Consultar Por:"
            CorTexto        =   0
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VB.Frame Frame2 
               Caption         =   "Dados da Recpção"
               Height          =   1725
               Left            =   120
               TabIndex        =   28
               Top             =   1065
               Width           =   10395
               Begin VB.Frame Frame3 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   1155
                  Left            =   5580
                  TabIndex        =   29
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
                     TabIndex        =   37
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
                     TabIndex        =   36
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
                     TabIndex        =   35
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
                     TabIndex        =   34
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
                     TabIndex        =   33
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
                     TabIndex        =   32
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
                     TabIndex        =   31
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
                     TabIndex        =   30
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
                  TabIndex        =   38
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
                  TabIndex        =   48
                  Top             =   195
                  Width           =   375
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Pagamentos:"
                  Height          =   195
                  Left            =   1020
                  TabIndex        =   47
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
                  TabIndex        =   46
                  Top             =   900
                  Width           =   45
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Data Recepção:"
                  Height          =   195
                  Index           =   0
                  Left            =   1215
                  TabIndex        =   45
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
                  TabIndex        =   44
                  Top             =   420
                  Width           =   45
               End
               Begin VB.Label Agencia 
                  AutoSize        =   -1  'True
                  Caption         =   "Agência:"
                  Height          =   195
                  Left            =   1740
                  TabIndex        =   43
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
                  TabIndex        =   42
                  Top             =   660
                  Width           =   45
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Conta:"
                  Height          =   195
                  Left            =   1860
                  TabIndex        =   41
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
                  TabIndex        =   40
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
                  TabIndex        =   39
                  Top             =   195
                  Width           =   60
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Informe o caminho dos arquivos de retorno.ret"
               Height          =   720
               Left            =   120
               TabIndex        =   27
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
                  TabIndex        =   11
                  Top             =   240
                  Width           =   7695
               End
               Begin VTOcx.cmdVISUAL cmdConsultaArquivo 
                  Height          =   405
                  Left            =   7860
                  TabIndex        =   12
                  Top             =   195
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   714
                  Caption         =   "Arquivo"
                  Acao            =   5
                  CorBorda        =   -2147483645
                  CorFundo        =   16777152
                  CorFoco         =   -2147483628
               End
               Begin VTOcx.cmdVISUAL cmdReceber 
                  Height          =   405
                  Left            =   9210
                  TabIndex        =   13
                  Top             =   195
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   714
                  Caption         =   "&Receber"
                  Acao            =   3
                  CorBorda        =   -2147483645
                  CorFundo        =   16777152
                  CorFoco         =   -2147483628
               End
            End
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL2 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   9360
            TabIndex        =   14
            Top             =   6600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "Sai&r"
            Acao            =   7
            CorFundo        =   16777152
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   7065
         Left            =   30
         TabIndex        =   16
         Top             =   360
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   12462
         _Version        =   131082
         TabGuid         =   "TARR101.frx":0E1C
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
            TabIndex        =   7
            Top             =   6600
            Value           =   1  'Checked
            Width           =   1665
         End
         Begin Cabecalho.cabVISUAL cabVISUAL1 
            Height          =   645
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   1138
            Formulario      =   "REMESSA DE DOCUMENTOS BANCARIO"
            Descricao       =   "Selecione os pagamentos que deseja enviar ao Banco conveniado"
            Icone           =   "TARR101.frx":0E44
         End
         Begin VTOcx.fraVISUAL txt 
            Height          =   1860
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   3281
            Altura          =   1905
            Caption         =   " Consultar Por:"
            CorTexto        =   0
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL cmdVISUAL6 
               Height          =   375
               Left            =   9960
               TabIndex        =   87
               Top             =   960
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               Caption         =   ""
               Acao            =   9
               CorFundo        =   16777088
            End
            Begin VTOcx.cmdVISUAL cmdListarDam 
               Height          =   390
               Left            =   2520
               TabIndex        =   1
               Top             =   360
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   688
               Caption         =   "DAMs Emitidos"
               Acao            =   1
               CorFundo        =   16777152
            End
            Begin VTOcx.txtVISUAL txtData 
               Height          =   285
               Left            =   360
               TabIndex        =   0
               Top             =   480
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   503
               Caption         =   "Data"
               Text            =   ""
               Formato         =   0
            End
            Begin MSComDlg.CommonDialog CommonDialog2 
               Left            =   360
               Top             =   1560
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VTOcx.cmdVISUAL cmdVISUAL3 
               Height          =   375
               Left            =   9360
               TabIndex        =   5
               Top             =   0
               Visible         =   0   'False
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
               Caption         =   "Arquivo"
               Acao            =   5
               CorBorda        =   16711680
               CorFrente       =   0
               CorFundo        =   16777088
               CorFoco         =   -2147483628
            End
            Begin VTOcx.txtVISUAL txtValor 
               Height          =   285
               Left            =   8880
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1035
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
               TabIndex        =   6
               Top             =   1905
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   688
               Caption         =   "Alterar"
               Acao            =   1
               CorBorda        =   16711680
               CorFrente       =   0
               CorFundo        =   16777088
            End
            Begin VTOcx.txtVISUAL txtNome 
               Height          =   285
               Left            =   2595
               TabIndex        =   20
               Top             =   1035
               Width           =   6180
               _ExtentX        =   10901
               _ExtentY        =   503
               Caption         =   "Dados"
               Text            =   ""
               Enabled         =   0   'False
               TipoLetras      =   0
            End
            Begin VTOcx.txtVISUAL txtDocumento 
               Height          =   285
               Left            =   300
               TabIndex        =   2
               Top             =   1035
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   503
               Caption         =   "Doc DAM"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtDiretorio 
               Height          =   285
               Left            =   2595
               TabIndex        =   4
               Top             =   1440
               Width           =   7740
               _ExtentX        =   13653
               _ExtentY        =   503
               Caption         =   "Diretorio Remessa"
               Text            =   ""
               TipoLetras      =   0
            End
            Begin VTOcx.txtVISUAL txtNumero 
               Height          =   285
               Left            =   300
               TabIndex        =   3
               Top             =   1440
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   503
               Caption         =   "Número Remessa"
               Text            =   ""
            End
            Begin VTOcx.cboVISUAL cboStatus1 
               Height          =   315
               Left            =   6000
               TabIndex        =   19
               Top             =   480
               Visible         =   0   'False
               Width           =   4005
               _ExtentX        =   7064
               _ExtentY        =   556
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorRotulo       =   0
               Enabled         =   0   'False
            End
            Begin VB.Line Line1 
               X1              =   240
               X2              =   10320
               Y1              =   840
               Y2              =   840
            End
         End
         Begin VTOcx.grdVISUAL Grid 
            Height          =   3660
            Left            =   120
            TabIndex        =   22
            Top             =   2760
            Width           =   10485
            _ExtentX        =   18494
            _ExtentY        =   6456
            Caption         =   "Dados"
            CheckBox        =   -1  'True
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL1 
            Height          =   390
            Left            =   1920
            TabIndex        =   8
            Top             =   6600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   688
            Caption         =   "Gerar Arquivo"
            Acao            =   3
            CorBorda        =   16711680
            CorFundo        =   16777152
         End
         Begin VTOcx.cmdVISUAL cmdLimpar 
            Height          =   390
            Left            =   3720
            TabIndex        =   9
            Top             =   6600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   688
            Caption         =   "Limpar"
            Acao            =   6
            CorBorda        =   16711680
            CorFundo        =   16777152
         End
         Begin VTOcx.cmdVISUAL cmdSair 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   9360
            TabIndex        =   10
            Top             =   6600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Sai&r"
            Acao            =   7
            CorBorda        =   16711680
            CorFundo        =   16777152
         End
      End
   End
End
Attribute VB_Name = "TARR101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim remHeaderBradesco           As New HeaderBradesco
Dim remBradesco                 As New RemessaBradesco
Dim trlBradesco                 As New TraillerBradesco
Dim Rs                          As VSRecordset
Dim Sql                         As String
Dim NomeArquivo                 As String
Dim i                           As Integer
Dim Doc                         As Integer
Dim marcou                      As Boolean
Dim arquivo                     As String
Dim sequencialDetalheRemessa    As Integer
Dim obrigacao                   As String
Dim codigoStatus                As Integer
Dim ValorDesconto As Currency
Dim obgris As String, sqlO As String
Dim CodObrigacao As String
Dim sqlGeral As String
Dim bcpRast As Boolean, bcpBancoArrecadacao As String, bcpBanco As String, bcpConta As String, bcpAgencia As String, bcpCarteira As String, bcpDigConta As String, bcpConvenioArrecadacao As String, bcpClienteRazao As String
Dim log As String

Dim isIptu As Boolean


Private Sub cboDados_Click()
    txtArquivo = cboDados
End Sub

Private Sub cboTIPO_Click()
    cboDados.Clear
    txtArquivo = ""
    Dim Tipo As Integer
    Tipo = cboTIPO.ListIndex
    If Tipo = 0 Then 'Remessa
        AtualizaCombo Bdados, cboDados, "SELECT  DISTINCT LEFT(COD_REMESSA, 6) AS REMESSA FROM TAB_BCP_REMESSA WHERE (SUBSTRING(COD_REMESSA, 3, 2)) = MONTH(GETDATE()) AND (SUBSTRING(COD_REMESSA, 7, 4)) = YEAR(GETDATE())"
    ElseIf Tipo = 1 Then 'Retorno
        'BASA
        If Temp.PegaParametro(Bdados, "BANCO ARRECADACAO") = 3 Then
            AtualizaCombo Bdados, cboDados, "SELECT  DISTINCT LEFT(COD_RETORNO, 4) AS RETORNO FROM TAB_BCP_RETORNO WHERE (SUBSTRING(COD_RETORNO, 3, 2)) = MONTH(GETDATE()) AND (SUBSTRING(COD_RETORNO, 5, 4)) = YEAR(GETDATE())"
        Else
        End If
    ElseIf Tipo = 2 Then 'Previsao
        AtualizaCombo Bdados, cboDados, "SELECT  DISTINCT DATA_CREDITO AS RETORNO FROM TAB_BCP_RETORNO WHERE (SUBSTRING(COD_RETORNO, 3, 2)) = MONTH(GETDATE()) AND (SUBSTRING(COD_RETORNO, 5, 4)) = YEAR(GETDATE())"
    End If
    
    
End Sub

Private Sub Check1_Click()
    Grid.MarcarTodos Check1.Value
End Sub
Private Sub cmdAddDam_Click()
    Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_OBSERVACAO='" & txtObservacao & "' WHERE TOC_COD_OBRIGACAO=" & txtDocumento)
End Sub

Private Sub cmdAlterarDam_Click()
Dim Obrig As New obrigacao
    Dim resultado As Boolean
    Dim Motivo As String
    
    Motivo = ""
    
    
    If Not Util.Confirma("Confirma a Alteração da Obrigação?") Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    If Obrig.AlteraObrigacaoTOBR201(CodObrigacao, txtVence, txtValorOriginal, txtMulta, txtJuros, 0, txtCorrecao, CInt(cboStatus.Coluna(1).Valor), txtDesconto, txtPeriodo, Motivo, txtParcelamento) Then
        'ALTERO NA TAB_COTAS_PARCELAMENTO DE ACORDO COM O NOVO PROCESSO DO PARCELAMENTO...
         If lstObrig.SelectedItem.SubItems(16) <> "" And lstObrig.SelectedItem.SubItems(16) <> "0" Then
            'checo se o parcelamento existe
            'É UM PARCELAMENTO
            Dim RsCotsa As VSRecordset
            
            If Bdados.AbreTabela("select TCP_STATUS_OBRIGACAO_PARCELA from TAB_COTAS_PARCELAMENTO where TCP_NUM_COTA = '" & lstObrig.SelectedItem & "'", RsCotsa) Then
                If Not IsNull(RsCotsa.Fields("TCP_STATUS_OBRIGACAO_PARCELA")) Then
                    Bdados.GravaDados "TAB_COTAS_PARCELAMENTO", CStr(cboStatus.Coluna(1).Valor), "TCP_STATUS_OBRIGACAO_PARCELA", "TCP_NUM_COTA = '" & lstObrig.SelectedItem & "'"
                End If
            End If
         End If
         If txtObservacao <> "" Then
            Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_OBSERVACAO='" & txtObservacao & "' WHERE TOC_COD_OBRIGACAO=" & CodObrigacao)
         End If
         
         Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_LOG_ALTERACAO='" & log & ":=" & Format(Now, "DD-MM-YY") & ";" & AplicacoesVTFuncoes.usuario & ";" & cboStatus & ";" & Format(txtValorOriginal, "#,##0.00") & ";D:" & Format(txtDescontoReal, "#,##0.00") & "' WHERE TOC_COD_OBRIGACAO=" & CodObrigacao)
        Avisa "Registro gravado."
        If chkComprovante = 1 Then
            Bdados.Executa ("INSERT INTO TAB_BCP_RETORNO (COD_RETORNO, COD_OBRIGACAO,DATA_OCRRENCIA_BANCO,DATA_CREDITO,STATUS,TIPO) VALUES('M" & Year(Now()) & "','" & CodObrigacao & "','" & txtCredito & "','" & txtCredito & "','5','MANUAL')")
        End If
        
    Else
        Avisa "Problemas ao gravar registro."
    End If
    'BCP
    Dim c As String, v As String, d As String
        
    If Len(txtDescontoReal) > 0 Then
        d = txtDescontoReal
    Else
        d = 0
    End If
    c = "tcc_desconto_concedido"
    v = Bdados.PreparaValor(Bdados.Converte(CCur(txtDescontoReal), TCDuplo))
    v = Bdados.GravaDados("Tab_Conta_Contribuinte", v, c, "tcc_codigo_conta =" & CodObrigacao)
    
    'FIM BCP
    cmdBuscarDam_Click
    LimparTodosCampos
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdAlteraTributo_Click()
    Dim Imposto, Inscricao As String
    Inscricao = txtInscricao
    Imposto = CStr(cboImposto.Coluna(0).Valor)
    If Inscricao = "" Or Imposto = "" Then
        Avisa "Informe a inscrição e imposto"
        Exit Sub
    End If
    Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_TIP_COD_IMPOSTO='" & Imposto & "', TOC_INSCRICAO = '" & Inscricao & "' WHERE TOC_COD_OBRIGACAO=" & CodObrigacao)
    Bdados.Executa ("UPDATE TAB_CONTA_CONTRIBUINTE SET TCC_TIP_COD_IMPOSTO='" & Imposto & "',TCC_INSCRICAO = '" & Inscricao & "', TCC_IM = '" & Inscricao & "' WHERE TCC_CODIGO_CONTA=" & CodObrigacao)
    Avisa "Tipo Imposto Alterado"
End Sub

Private Sub cmdBuscar_Click()
    Dim Tipo As Integer
    Tipo = cboTIPO.ListIndex
    
    If Tipo = 0 Then 'Remessa
        bcpRast = False
        sqlGeral = "SELECT * FROM VIS_BCP_REMESSA WHERE COD_REMESSA = '" & txtArquivo & txtAno & "' ORDER BY COD_OBRIGACAO"
    ElseIf Tipo = 1 Then 'Retorno
        bcpRast = False
        sqlGeral = "SELECT * FROM VIS_BCP_RETORNO WHERE COD_RETORNO = '" & txtArquivo & txtAno & "' ORDER BY COD_RETORNO,COD_OBRIGACAO"
    ElseIf Tipo = 2 Then 'Previsao
        bcpRast = False
        sqlGeral = "SELECT * FROM VIS_BCP_RETORNO WHERE DATA_CREDITO = '" & txtArquivo & "' ORDER BY DATA_CREDITO,COD_RETORNO,COD_OBRIGACAO"
    ElseIf Tipo = 3 Then 'Rastreamento Retorno
        bcpRast = False
        sqlGeral = "SELECT * FROM VIS_BCP_RASTREAMENTO WHERE COD_RETORNO = '" & txtArquivo & txtAno & "' ORDER BY COD_RETORNO,COD_OBRIGACAO"
    ElseIf Tipo = 4 Then 'Rastreamento Documento
        sqlGeral = "SELECT * FROM VIS_BCP_RASTREAMENTO WHERE COD_OBRIGACAO in ("
        obgris = obgris & "'" & txtArquivo & "',"
        sqlGeral = sqlGeral & obgris
        sqlGeral = Left(sqlGeral, Len(sqlGeral) - 1) & ") ORDER BY COD_OBRIGACAO"
        bcpRast = True
    ElseIf Tipo = 6 Then 'Conciliação
        sqlGeral = "SELECT * FROM VIS_BCP_RASTREAMENTO where previsao >= '" & txtArquivo & "' and previsao <='" & txtArquivo1 & "'"
        sqlGeral = sqlGeral & " ORDER BY COD_OBRIGACAO"
        bcpRast = True
    
    End If
    If grdGerencial.Preencher(Bdados, sqlGeral) Then
    End If
End Sub

Private Sub cmdBuscarDam_Click()
    Dim Obrig As obrigacao
    Set Obrig = New obrigacao
    If txtDAM = "" Then
        'Exit Sub
    End If
    If Not Obrig.MostraObrigacaoGerada(lstObrig, "", txtIm, , , "", _
            "", , , , "", , IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False), txtDAM) Then
            Avisa "Nenhum registro encontrado."
            txtIm.SetFocus
    End If
    'If Not Obrig.MostraObrigacaoGerada(lstObrig, "", "", , , "", _
     '       "", , , , "", , IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False), txtDAM) Then
      '      Avisa "Nenhum registro encontrado."
            
    'End If
    
End Sub

Private Sub cmdTransferir_Click()
    If txtDestinoInscricao = "" Then
        Mensagem ("Informe a inscrição de destino")
        Exit Sub
    End If
    
    Dim obrigacao As String, data As String, usuario As String, Sql As String
    Dim TipoInscricao As Integer
    data = Format(Now, "dd/mm/yyyy")
    usuario = UCase(AplicacoesVTFuncoes.usuario)
    For Doc = 1 To lstDebitos.ListItems.Count ' TOTAL DE TRIBUTOS
        If lstDebitos.ListItems(Doc).Checked = True Then ' SE FOI MARCADO PARA REMESSA
            obrigacao = lstDebitos.ListItems(Doc).SubItems(1)
            Sql = "INSERT INTO TAB_TRANSF_DEBITOS (data, usuario,origem,obrigacao,destino,obs)"
            Sql = Sql & " VALUES('" & data & "','" & usuario & "','" & txtOrigemInscricao & "','" & obrigacao & "',"
            Sql = Sql & "'" & txtDestinoInscricao & "','" & txtObsTransferencia & "')"
            Bdados.Executa (Sql)
            
            If isIptu = True Then
                TipoInscricao = 1
            Else
                Bdados.Executa ("UPDATE TAB_CONTRIBUINTE SET TCI_OBSERVACAO= TCI_NOME, TCI_NOME= 'CADASTRO MOBILIARIO INVÁLIDO', TCI_FANTASIA =  'CADASTRO MOBILIARIO INVÁLIDO' WHERE TCI_IM='" & txtOrigemInscricao & "'")
                TipoInscricao = 2
            End If
            
            Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_TIPO_INSCRICAO = " & TipoInscricao & ", TOC_INSCRICAO='" & txtDestinoInscricao & "' WHERE TOC_COD_OBRIGACAO=" & obrigacao)
            Bdados.Executa ("UPDATE TAB_CONTA_CONTRIBUINTE SET TCC_TIPO_INSCRICAO = " & TipoInscricao & ", tcc_inscricao='" & txtDestinoInscricao & "' WHERE tcc_codigo_conta=" & obrigacao)
            
        End If
    Next Doc
    lstDebitos.ListItems.Clear
    txtObsTransferencia = ""
    txtOrigemInscricao = ""
    txtOrigemNome = ""
    txtDestinoInscricao = ""
    txtDestinoNome = ""
    txtOrigemInscricao.SetFocus
    Mensagem ("Transferência realizada com sucesso!")
End Sub

Private Sub cmdVISUAL7_Click()
    If isIptu = True Then
        AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtDestinoInscricao
    Else
        AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtDestinoInscricao
    End If
End Sub

Private Sub cmdVISUAL8_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim Obrig As obrigacao
    Set Obrig = New obrigacao
    Obrig.PreencheComboTributo cboImposto, False
End Sub

Private Sub cmdOrigemInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtOrigemInscricao
End Sub

Private Sub optImposto_Click(Index As Integer)
    isIptu = optImposto(0).Value
End Sub


Private Sub txtOrigemInscricao_LostFocus()
    Dim Rs As VSRecordset
    Dim sqlIptus As String, igualOuDif As String
    Dim cadastro As New VSImposto
    If Trim(txtOrigemInscricao) <> "" Then
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            txtOrigemInscricao = cadastro.FormataInscricao(txtOrigemInscricao, InscContrib)
        End If
        Sql = "Select  tci_Nome from Tab_Contribuinte where tci_im = '" & txtOrigemInscricao & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            txtOrigemNome = "" & Rs(0)  'Rs!tci_Nome
            txtObsTransferencia = "TRANSFERENCIA DE DEBITOS ENTRE CONTRIBUINTES, ORIGEM: " & txtOrigemInscricao & " : " & txtOrigemNome

            If isIptu = True Then
                igualOuDif = "="
            Else
                igualOuDif = "<>"
            End If
            sqlIptus = "SELECT SIGLA ,OBRIGACAO, PERIODO, VALOR, STATUS,OBSERVACAO FROM VIS_OBRIGACAO WHERE SIGLA " & igualOuDif & " 'IPTU' AND INSCRICAO = '" & txtOrigemInscricao & "'"
            If lstDebitos.Preencher(Bdados, sqlIptus) Then
            End If
            Dim i As Integer
            For i = 1 To lstDebitos.ListItems.Count
                    lstDebitos.ListItems(i).Checked = True
            Next i
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub txtDestinoInscricao_LostFocus()
    Dim RsI As VSRecordset
    Dim SqlI As String
    Dim cadastro As New VSImposto
    If isIptu = True Then
        If Trim(txtDestinoInscricao) <> "" Then
            SqlI = "Select  tipologradouro, logradouro, tim_numero, tba_nome from vis_imovel where tim_ic = '" & txtDestinoInscricao & "'"
            If Bdados.AbreTabela(SqlI, RsI) Then
                txtDestinoNome = "" & RsI(0) & " " & RsI(1) & " " & RsI(2) & ", " & RsI(3) 'Rs!tci_Nome
            Else
                Call Util.Informa("Contribuinte não cadastrado.")
            End If
        End If
    Else
        If Trim(txtDestinoInscricao) <> "" Then
            SqlI = "Select  tci_cgc_cpf, tci_nome, tci_fantasia from tab_contribuinte where tci_im = '" & txtDestinoInscricao & "'"
            If Bdados.AbreTabela(SqlI, RsI) Then
                txtDestinoNome = "" & RsI(0) & " - " & RsI(1) & " (" & RsI(2) & ")"
            Else
                Call Util.Informa("Contribuinte não cadastrado.")
            End If
        End If
    End If
    txtObsTransferencia = txtObsTransferencia & " PARA " & txtDestinoInscricao & " : " & txtDestinoNome
    Bdados.FechaTabela RsI
End Sub

Private Sub txtNossoNumero_LostFocus()
    If txtNossoNumero <> "" Then
        If Bdados.AbreTabela("SELECT TOC_COD_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_NOSSO_NUMERO=" & txtNossoNumero, Rs) Then
            txtDAM = IIf(IsNull(Rs(0)), "", Rs(0))
        End If
    End If
End Sub



Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub
Private Sub LimparTodosCampos()
    txtDesconto = 0
    CodObrigacao = 0
    txtInscricao = ""
    txtTributo = ""
    txtPeriodo = 0
    txtPeriodo = 0
    txtVence = ""
    txtValorOriginal = Format(0, Const_Monetario)
    txtMulta = Format(0, Const_Monetario)
    txtJuros = Format(0, Const_Monetario)
    txtCorrecao = Format(0, Const_Monetario)
    'txtTaxas = Format(Nvl(lstObrig.SelectedItem.SubItems(10), 0), Const_Monetario)
    txtParcelamento = 0
    'cboStatus.SetarLinha Format(0), 1
    txtDesconto = Format(0, Const_Monetario)
    txtJurosMargem = 0
    txtMultaMargem = 0
    txtDescontoReal = (CCur(0))
    
    'Dim Rs As VSRecordset
    txtObservacao = ""
    log = "-"
End Sub
Private Sub lstObrig_DblClick()
    
    Dim Obrig As New obrigacao
    
    
    If lstObrig.ListItems.Count = 0 Then Exit Sub
    txtDesconto = 0
    CodObrigacao = lstObrig.SelectedItem
    txtInscricao = lstObrig.SelectedItem.SubItems(1)
    txtTributo = lstObrig.SelectedItem.SubItems(3)
    txtPeriodo = lstObrig.SelectedItem.SubItems(4)
    txtPeriodo = IIf(Len(Trim(txtPeriodo)) = 4, txtPeriodo, Right(txtPeriodo, 2) & Left(txtPeriodo, 4))
    txtVence = lstObrig.SelectedItem.SubItems(5)
    txtValorOriginal = Format(lstObrig.SelectedItem.SubItems(6), Const_Monetario)
    txtMulta = Format(lstObrig.SelectedItem.SubItems(8), Const_Monetario)
    txtJuros = Format(lstObrig.SelectedItem.SubItems(7), Const_Monetario)
    txtCorrecao = Format(IIf(Len(lstObrig.SelectedItem.SubItems(17)) = 0, 0, lstObrig.SelectedItem.SubItems(17)), Const_Monetario)
    'txtTaxas = Format(Nvl(lstObrig.SelectedItem.SubItems(10), 0), Const_Monetario)
    txtParcelamento = lstObrig.SelectedItem.SubItems(16)
    cboStatus.SetarLinha Format(Nvl(lstObrig.SelectedItem.SubItems(15), -1)), 1
    txtDesconto = Format(Nvl(lstObrig.SelectedItem.SubItems(18), 0), Const_Monetario)
    txtJurosMargem = 0
    txtMultaMargem = 0
    txtDescontoReal = (CCur(txtValorOriginal) + CCur(txtMulta) + CCur(txtJuros)) * (txtDesconto / 100)
    
    'Dim Rs As VSRecordset
    Dim valorReal As Double
    valorReal = CCur(txtValorOriginal) + CCur(txtJuros) + CCur(txtMulta) + CCur(txtCorrecao) - CCur(txtDescontoReal)
    lblValorReal = "R$ Real: " & Format(valorReal, Const_Monetario)
    If Bdados.AbreTabela("SELECT TOC_OBSERVACAO,TOC_LOG_ALTERACAO,TOC_TIP_COD_IMPOSTO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO=" & CodObrigacao, Rs) Then
        txtObservacao = IIf(IsNull(Rs(0)), "", Rs(0))
        log = IIf(IsNull(Rs(1)), "", Rs(1))
    Else
        txtObservacao = ""
        log = "-"
    End If
   
End Sub
Private Sub cmdConsultaArquivo_Click()
    With CommonDialog1
        .DialogTitle = "Selecione o arquivo retorno"
        '.Filter = "Arquivos do tipo retorno | *.ret"
        .Filter = "Arquivos do tipo retorno | *.*"
        
        .ShowOpen
        If .FileName <> "" Then
            txtCamminhoRemessa = .FileName
            NomeArquivo = .FileTitle
        End If
    End With
End Sub

Private Sub cmdImprimir_Click()
     Screen.MousePointer = 11
     Dim cabecalho As String
     cabecalho = txtArquivo
    With RPT
        If cboTIPO.ListIndex = 0 Then 'REMESSA
            If Not .DefinirArquivo(Bdados, App.Path + "\TREMESSA.rpt") Then Exit Sub
            If Len(txtArquivo) > 0 Then
                .SELECAO = "{VIS_BCP_REMESSA.COD_REMESSA} = '" & txtArquivo & txtAno & "'"
            Else
                .SELECAO = "{VIS_BCP_REMESSA.ANO_REMESSA} = '" & txtAno & "'"
            End If
        ElseIf cboTIPO.ListIndex = 1 Then 'RETORNO
            If cboOrdem.ListIndex = 0 Then 'DOCUMENTO
                If Not .DefinirArquivo(Bdados, App.Path + "\TRETORNO.rpt") Then Exit Sub
                .SELECAO = "{VIS_BCP_RETORNO.COD_RETORNO} = '" & txtArquivo & txtAno & "'"
            Else 'TRIBUTO
                If Not .DefinirArquivo(Bdados, App.Path + "\TRETORNOTRIBUTO.rpt") Then Exit Sub
                .SELECAO = "{VIS_BCP_RETORNO.COD_RETORNO} = '" & txtArquivo & txtAno & "' "
            End If
        ElseIf cboTIPO.ListIndex = 2 Then 'PREVISAO
            If Not .DefinirArquivo(Bdados, App.Path + "\TPREVISAOCREDITO.rpt") Then Exit Sub
            .SELECAO = "{VIS_BCP_RETORNO.DATA_CREDITO} = '" & txtArquivo & "' AND {VIS_BCP_RETORNO.STATUS} <>2 AND {VIS_BCP_RETORNO.STATUS} <>10" '2=REGISTRADO NO BANCO, 6 E 17 COMPENSADOS
        ElseIf cboTIPO.ListIndex = 3 Then 'RASTREAMENTO RETORNO
            If Not .DefinirArquivo(Bdados, App.Path + "\TRASTREAMENTOPAGAMENTO.rpt") Then Exit Sub
            If Len(txtArquivo) > 0 Then
                .SELECAO = "{VIS_BCP_RASTREAMENTO.COD_RETORNO} = '" & txtArquivo & txtAno & "'"
            End If
        ElseIf cboTIPO.ListIndex = 4 Then 'RASTREAMENTO DOCUMENTO
            If Not .DefinirArquivo(Bdados, App.Path + "\TRASTREAMENTOPAGAMENTO.rpt") Then Exit Sub
            If grdGerencial.ListItems.Count > 0 And bcpRast = True Then
                Dim xx As Integer
                obgris = ""
                sqlGeral = ""
                txtArquivo = ""
                'For xx = 1 To grdGerencial.ListItems.Count
                    'obgris = "{VIS_BCP_RASTREAMENTO.COD_OBRIGACAO} = '" & grdGerencial(xx).SubItems(1) & "' or "
                    'sqlGeral = sqlGeral & obgris
                'Next xx
                'sqlGeral = Left(sqlGeral, Len(sqlGeral) - 3)
                
                .SELECAO = "{VIS_BCP_RASTREAMENTO.COD_OBRIGACAO} = '" & txtArquivo & "'"
            Else
                .SELECAO = "{VIS_BCP_RASTREAMENTO.ANO_RETORNO} =  '" & txtAno & "'"
            End If
        ElseIf cboTIPO.ListIndex = 5 Then 'PREVISAO TRIBUTO
            If Not .DefinirArquivo(Bdados, App.Path + "\TPREVISAOCREDITO.rpt") Then Exit Sub
            Dim Filtro As String
            If Len(txtArquivo) = 0 Then
                Mensagem "Informe a Data Inicial"
                txtArquivo.SetFocus
                Exit Sub
            End If
            If Len(txtArquivo1) = 0 Then
                Mensagem "Informe a Data Final"
                txtArquivo1.SetFocus
                Exit Sub
            End If
            Dim di As String, df As String, m As String, a As String 'd=dia, m=mes, a=ano - Alias parcial de data credito na tab_bcp_retorno
            di = Format(Left(txtArquivo, 2), "00")
            df = Format(Left(txtArquivo1, 2), "00")
            
            m = Format(Mid(txtArquivo1, 4, 2), "00")
            a = Format(Right(txtArquivo1, 4), "0000")
            
            
            Filtro = Filtro & "{VIS_BCP_RETORNO.DIA_CREDITO} >='" & di _
            & "' AND {VIS_BCP_RETORNO.DIA_CREDITO} <='" & df _
            & "' AND {VIS_BCP_RETORNO.MES_CREDITO} ='" & m _
            & "' AND {VIS_BCP_RETORNO.ANO_CREDITO} ='" & a & "'"
            'Filtro = Filtro & " AND {VIS_BCP_RETORNO.STATUS} <>2 OR {VIS_BCP_RETORNO.STATUS} <>10"  '2=REGISTRADO NO BANCO, 6 E 17 COMPENSADOS
            .SELECAO = Filtro
            cabecalho = txtArquivo & " a " & txtArquivo1
        ElseIf cboTIPO.ListIndex = 6 Then 'CONCILIAÇÃO
            If Not .DefinirArquivo(Bdados, App.Path + "\TRASTREAMENTOPAGAMENTO.rpt") Then Exit Sub
            If Len(txtArquivo) > 0 Then
                .SELECAO = "{VIS_BCP_RASTREAMENTO.PREVISAO} >='" & txtArquivo & "' AND {VIS_BCP_RASTREAMENTO.PREVISAO} <='" & txtArquivo1 & "'"
            End If
        End If
        
        .Formulas "ANO", txtAno
        .Formulas "ARQUIVO", cabecalho
        obgris = ""
        sqlGeral = ""
        'txtArquivo = ""
        .Visualizar
    End With
    Set RPT = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    Grid.ListItems.Clear
    'cboStatus.SetarLinha esrAberto, 1
End Sub
Private Sub Form_Load()
    cmdVISUAL1.Enabled = False
    Dim Obrig As New obrigacao
    Dim User As String
    'cboStatus.Enabled = False
    'chkComprovante.Enabled = False
    txtCredito.Visible = True
    Dim Rs As VSRecordset
    Dim sim As Integer
    cboStatus.Enabled = sim
    chkComprovante.Enabled = sim
    txtCredito.Visible = sim
    If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1264 Then 'pinheiro
        User = UCase(AplicacoesVTFuncoes.usuario)
        If Bdados.AbreTabela("SELECT TUS_ARRECADACAO from tab_usuario WHERE TUS_COD_USUARIO='" & User & "'", Rs) Then
            sim = Rs(0)
            cboStatus.Enabled = sim
            chkComprovante.Enabled = sim
            txtCredito.Visible = sim
        End If
        'If User <> "MARCELOSILVA" And User <> "HENRIQUE" And User <> "MAURICIO" Then
         '   cboStatus.Enabled = False
          '  chkComprovante.Enabled = False
           ' txtCredito.Visible = False
        'End If
    
    End If
    optImposto(0).Value = True
    optImposto(1).Value = False
    isIptu = True
    'DELETE FROM Tab_Conta_Contribuinte;
    'DELETE FROM TAB_OBRIGACAO_CONTRIBUINTE;
    'DELETE FROM TAB_BCP_REMESSA;
    'DELETE FROM TAB_BCP_RETORNO;
    
    
    'select * FROM Tab_Conta_Contribuinte;
    'select * FROM TAB_OBRIGACAO_CONTRIBUINTE;
    'select * FROM TAB_BCP_REMESSA;
    'select * FROM TAB_BCP_RETORNO;
    
    
    'cabVisual.Exibir Bdados, Me.Name, App.Path
    'rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    'Obrig.PreencheComboTributo CboImposto, False
    cboStatus.PreencherGeral Bdados, "STATUS OBRIGACAO"
    txtDescontoReal = 0
    bcpBanco = Temp.PegaParametro(Bdados, "BANCO ARRECADACAO")
     If bcpBanco = "1" Then 'BB codo pinheiro
        txtDiretorio = "C:\BancoBrasil\BBTransf\Remessa\"
        txtCamminhoRemessa = "C:\BancoBrasil\BBTransf\Retorno\"
     ElseIf bcpBanco = "3" Then 'BASA
        txtDiretorio = "C:\STCPCLT_ATP\O0055ATPCLIENT\SAIDA\"
        txtCamminhoRemessa = "C:\STCPCLT_ATP\O0055ATPCLIENT\ENTRADA\"
     Else 'BRADESCO
        txtDiretorio = "C:\ORION TECNOLOGIAS\BRADESCO\REMESSA\"
        txtCamminhoRemessa = "C:\ORION TECNOLOGIAS\BRADESCO\RETORNO\"
     End If
     Dim Remessa As String
     'Dim Rs As VSRecordset
     Remessa = Format(Now, "DDMM")
     Remessa = Remessa & "01" & Format(Now, "YYYY")
     If Bdados.AbreTabela("SELECT COD_REMESSA FROM TAB_BCP_REMESSA WHERE COD_REMESSA=" & Remessa, Rs) Then
        If Rs.RecordCount = 0 Then
            txtNumero = 1
        Else
            txtNumero = 2
        End If
     Else
        txtNumero = 1
     End If
    bcpBancoArrecadacao = Format(Temp.PegaParametro(Bdados, "BANCO ARRECADACAO"), "000")
    bcpConta = Format(Temp.PegaParametro(Bdados, "CONVENIO CONTA"), "0000000")
    bcpAgencia = Format(Temp.PegaParametro(Bdados, "CONVENIO AGENCIA"), "00000")
    bcpCarteira = Format(Temp.PegaParametro(Bdados, "CONVENIO CARTEIRA"), "000")
    bcpConvenioArrecadacao = Format(Temp.PegaParametro(Bdados, "CONVENIO ARRECADACAO"), "00000000000000000000") '' CODIGO EMPRESA
    bcpDigConta = Right(Temp.PegaParametro(Bdados, "CONTA"), 1)
    txtData = Format(Now, "DD/MM/YYYY")
    bcpClienteRazao = Temp.PegaParametro(Bdados, "RAZAO")
     'txtInicio = Format(Now, "DD/MM/YYYY")
     'txtFim = Format(Now, "DD/MM/YYYY")
     cboTIPO.AddItem "REMESSA"
     cboTIPO.AddItem "RETORNO"
     cboTIPO.AddItem "PREVISAO DIARIA"
     cboTIPO.AddItem "RASTREAR RETORNO"
     cboTIPO.AddItem "RASTREAR DOCUMTO"
     cboTIPO.AddItem "PREVISAO TRIBUTO"
     cboTIPO.AddItem "CONCILIAÇÃO"
     
     cboOrdem.AddItem "DOCUMENTO"
     cboOrdem.AddItem "TRIBUTO"
     txtAno = Format(Now, "YYYY")
     obgris = ""
End Sub
Private Sub montarHeder()
        Dim cliente, NomeBanco As String
        arquivo = "CB" & Format(Now, "DDMM") & Format(txtNumero, "00") & ".REM"
        NomeBanco = "BRADESCO"
        If bcpBanco = 3 Then 'basa
            cliente = Format(Temp.PegaParametro(Bdados, "CONVENIO ARRECADACAO"), "000000000") '' CODIGO EMPRESA
            arquivo = cliente & Format(Now, "MM") & Format(Now, "DD") & ".REM." & Format(txtNumero, "000")
            NomeBanco = "Banco Amazonia"
            
        End If
        With remHeaderBradesco
            .IdentificacaoRegistro = "0"
            .IdentificaoArquivoRemessa = "1"
            .LiteralRemessa = "REMESSA"
            .CodigoServico = "01"
            .LiteralServico = preencherComCaractere("COBRANCA", 15, " ")
            .CodigoEmpresa = Format(bcpConvenioArrecadacao, "00000000000000000000")
            .NomeEmpresa = preencherComCaractere(bcpClienteRazao, 30, " ")
            .NumeroCamaraCompensacao = bcpBancoArrecadacao
            .NomeBanco = preencherComCaractere(NomeBanco, 15, " ")
            .DataGravacaoArquivo = Format(Now, "DDMMYY")
            .IdentificacaoSistema = "MX"
            Dim sequencialRemessa As Long
            sequencialRemessa = Imposto.GeraNumCorrelativo(1, 101)
            .NumeroSequencialRemessa = Format(sequencialRemessa, "0000000")
            .NumeroSequencialRegistroUmAUm = Format(1, "000000")
            arquivo = .gerarHeaderRemessa(txtDiretorio, arquivo) 'DIRETORIO + NOME ARQUIVO
            
        End With
End Sub
Private Sub montarDetalhe()
             sequencialDetalheRemessa = 2
             Dim ValorTitulo As String
             Dim cpfCnpj As String
             Dim Endereco As String
             Dim ccep As String
             If bcpBanco = 3 Then 'basa
                For Doc = 1 To Grid.ListItems.Count ' TOTAL DE TRIBUTOS
                    If Grid.ListItems(Doc).Checked = True Then ' SE FOI MARCADO PARA REMESSA
                        obrigacao = Grid.ListItems(Doc)
                        If Bdados.AbreTabela("SELECT * FROM VIS_CONTA_CONTRIBUINTE WHERE NUM_DOCUMENTO=" & obrigacao) Then
                           
                            With remBradesco
                                .IdentificacaoRegistro = "1"
                                .Filler02 = "0000000000000000000"
                                .IdentificacaoEmpresaCedenteNoBanco = "0" & bcpCarteira & bcpAgencia & Format(bcpConta, "00000000")
                                .NumeroControleParticipante = "0000000000000000000000000"
                                .Filler05 = "00000000"
                                
                                .IdentificacaoTituloBanco = Format(Bdados.Tabela("TOC_NOSSO_NUMERO"), "000000000000") '"000000000000"
                                
                                .DescontoBonificacaoDia = "0000000000"
                                .CondicaoParaEmissaoPapeladaCobranca = "2"
                                .DebitoAutomatico = "N"
                                .Filler10 = preencherComCaractere(" ", 14, " ")
                                .IndicacaoOcorrencia = "01"
                                .NumeroDocumento = Format(Bdados.Tabela("NUM_DOCUMENTO"), "0000000000")
                                .DataVencimentoTitulo = Format(Bdados.Tabela("DATA_VENCIMENTO"), "DDMMYY")
                                ValorTitulo = Format(Bdados.Tabela("VALOR_ATUAL"), "#,##0.00")
                                .ValorTitulo = Format(retiraSeparadores(ValorTitulo), "0000000000000")
                                .BancoEncarregadoCobranca = bcpBancoArrecadacao
                                .AgenciaDepositaria = "00000"
                                .EspecieTitulo = "01"
                                .Identificacao = "A"
                                .DataEmissaoTitulo = Format(Bdados.Tabela("DATA_GERACAO_BOLETO"), "DDMMYY")
                                .Instrucao1 = "00"
                                .Instrucao2 = "00"
                                .ValorCobradoDiaAtraso = "0000000000000"
                                .DataLimiteConcessaoDesconto = Format(Bdados.Tabela("DATA_VENCIMENTO"), "DDMMYY")
                                .ValorDesconto = "0000000000000"
                                .ValorIOF = "0000000000000"
                                .ValorAbatimento = "0000000000000"
                                
                                cpfCnpj = "00000000000000"
                                If IsNull(Bdados.Tabela("CPF_CNPJ")) Then
                                    .IdentificacaoTipoInscricaoSacado = "99" 'OUTROS
                                ElseIf Len(Bdados.Tabela("CPF_CNPJ")) = 0 Then
                                    .IdentificacaoTipoInscricaoSacado = "99" 'OUTROS
                                Else
                                    If Not IsNull(Bdados.Tabela("CPF_CNPJ")) Then
                                        cpfCnpj = retiraSeparadores(Bdados.Tabela("CPF_CNPJ"))
                                    End If
                                    If Len(cpfCnpj) <= 11 Then
                                      .IdentificacaoTipoInscricaoSacado = "01" 'CPF
                                    Else
                                        .IdentificacaoTipoInscricaoSacado = "02" 'CNPJ
                                    End If
                                End If
                                .NumeroInscricaoSacado = Format(cpfCnpj, "00000000000000")
                                .NomeSacado = preencherComCaractere(Bdados.Tabela("NOME"), 40, " ")
                                Endereco = Bdados.Tabela("LOGRADOURO") & " " & Bdados.Tabela("NOME_LOGRADOURO") & " " & Bdados.Tabela("NUMERO_ENDERECO") & " "
                                'Endereco = "PREFEITURA MUNICIPAL"
                                
                                .EnderecoCompleto = preencherComCaractere(Endereco, 40, " ")
                                .Bairro = preencherComCaractere(Bdados.Tabela("BAIRRO"), 12, " ")
                                '.Bairro = preencherComCaractere("BAIRRO", 12, " ")
                                
                                ccep = "00000000" 'retiraSeparadores(IIf(IsNull(Bdados.Tabela("CEP")), "00000000", Bdados.Tabela("CEP")))
                                ccep = retiraSeparadores(IIf(IsNull(Bdados.Tabela("CEP")), "00000000", Bdados.Tabela("CEP")))
                                
                                ccep = preencherComCaractere(ccep, 8, "0")
                                .Cep = Left(ccep, 5)
                                .SufixoCEP = Right(ccep, 3)
                                .Municipio = preencherComCaractere(Bdados.Tabela("CIDADE"), 15, " ")
                                '.Municipio = preencherComCaractere("CIDADE", 15, " ")
                                
                                .UF = preencherComCaractere(Bdados.Tabela("ESTADO"), 2, " ")
                                '.UF = preencherComCaractere("UF", 2, " ")
                                
                                
                                .Filler36 = preencherComCaractere(" ", 43, " ")
                                .NumeroSequencialRegistro = Format(sequencialDetalheRemessa, "000000") ' SEQUENCIAL COMECANDO COM 2
                                sequencialDetalheRemessa = sequencialDetalheRemessa + 1
                                arquivo = .gerarDetalheRemessaBasa(txtDiretorio, arquivo)
                                
                                '.AgenciaDebito = Format(0, "00000")
                                '.DigitoAgenciaDebito = "0"
                                '.RazaoContaCorrente = Format(0, "00000")
                                '.ContaCorrente = Format(0, "0000000")
                                '.DigitoContaCorrente = "0"
                                '.IdentificacaoEmpresaCedenteNoBanco = "0" & bcpCarteira & bcpAgencia & bcpConta & bcpDigConta
                                '.NumeroControleParticipante = "0000000000000000000000000"
                                '.CodigoBancoCamaraCompensacao = "000"
                                '.CampoMulta = 0
                                '.PercentualMulta = Format(0, "0000")
                                '.IdentificacaoTituloBanco = Format(Bdados.Tabela("NUM_DOCUMENTO"), "00000000000") ''
                                '.DigitoAutoConferenciaNossoNumero = .gerarDigitoConferencia82(bcpCarteira)
                                '.DescontoBonificacaoDia = Format(0, "0000000000")
                                '.CondicaoParaEmissaoPapeladaCobranca = "2"
                                '.EmitePapeletaDebitoAutomatico = "N"
                                '.IndicadorRateioCredito = " "
                                '.EnderecamentoAvisoDebAutoCC = "2"
                                '.IndicacaoOcorrencia = "01"
                                '.NumeroDocumento = Format(Bdados.Tabela("NUM_DOCUMENTO"), "0000000000")
                                '.DataVencimentoTitulo = Format(Bdados.Tabela("DATA_VENCIMENTO"), "DDMMYY")
                                
                                'ValorTitulo = Format(Bdados.Tabela("VALOR_ATUAL"), "#,##0.00")
                                '.ValorTitulo = Format(retiraSeparadores(ValorTitulo), "0000000000000")
                                '.BancoEncarregadoCobranca = "000" '
                                '.AgenciaDepositaria = "00000" '
                                '.EspecieTitulo = "01"
                                '.Identificacao = "N"
                                '.DataEmissaoTitulo = Format(Bdados.Tabela("DATA_GERACAO_BOLETO"), "DDMMYY")
                                '.PrimeiraInstrucao = "00"
                                '.SegundaInstrucao = "00"
                                '.ValorCobradoDiaAtraso = Format(0, "0000000000000")
                                '.DataLimiteConcessaoDesconto = Format(0, "000000")
                                '.ValorDesconto = Format(0, "0000000000000")
                                '.ValorIOF = Format(0, "0000000000000")
                                '.ValorAbatimento = Format(0, "0000000000000")
                                
                                'cpfCnpj = "00000000000000"
                                'If IsNull(Bdados.Tabela("CPF_CNPJ")) Then
                                 '   .IdentificacaoTipoInscricaoSacado = "99" 'OUTROS
                                'ElseIf Len(Bdados.Tabela("CPF_CNPJ")) = 0 Then
                                 '   .IdentificacaoTipoInscricaoSacado = "99" 'OUTROS
                                'Else
                                    
                                 '   cpfCnpj = retiraSeparadores(Bdados.Tabela("CPF_CNPJ"))
                                  '  If Len(cpfCnpj) < 14 Then
                                   '     .IdentificacaoTipoInscricaoSacado = "01" 'CPF
                                    'Else
                                     '   .IdentificacaoTipoInscricaoSacado = "02" 'CNPJ
                                    'End If
                                'End If
                                '.NumeroInscricaoSacado = Format(cpfCnpj, "00000000000000")
                                '.NomeSacado = preencherComCaractere(Bdados.Tabela("NOME"), 40, " ")
                                
                                'Endereco = Bdados.Tabela("LOGRADOURO") & " " & Bdados.Tabela("NOME_LOGRADOURO") & " " & Bdados.Tabela("NUMERO_ENDERECO") & " " & Bdados.Tabela("BAIRRO") & " " & Bdados.Tabela("CIDADE") & " " & Bdados.Tabela("ESTADO")
                                '.EnderecoCompleto = preencherComCaractere(Endereco, 40, " ")
                                '.PrimeiraMensagem = preencherComCaractere(" ", 12, " ")
                                
                                'ccep = retiraSeparadores(IIf(IsNull(Bdados.Tabela("CEP")), "00000000", Bdados.Tabela("CEP")))
                                'ccep = preencherComCaractere(ccep, 8, "0")
                                '.cep = Left(ccep, 5)
                                '.SufixoCEP = Right(ccep, 3)
                                '.SacadorAvalistaSegundaMensagem = preencherComCaractere(" ", 60, " ")
                                '.NumeroSequencialRegistro = Format(sequencialDetalheRemessa, "000000") ' SEQUENCIAL COMECANDO COM 2
                                
                            End With
                        End If
                    End If
                Next Doc
             
             Else 'demais bancos
              
             
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
                                .DataVencimentoTitulo = Format(Bdados.Tabela("DATA_VENCIMENTO"), "DDMMYY")
                                
                                ValorTitulo = Format(Bdados.Tabela("VALOR_ATUAL"), "#,##0.00")
                                .ValorTitulo = Format(retiraSeparadores(ValorTitulo), "0000000000000")
                                .BancoEncarregadoCobranca = "000" '
                                .AgenciaDepositaria = "00000" '
                                .EspecieTitulo = "01"
                                .Identificacao = "N"
                                .DataEmissaoTitulo = Format(Bdados.Tabela("DATA_GERACAO_BOLETO"), "DDMMYY")
                                .PrimeiraInstrucao = "00"
                                .SegundaInstrucao = "00"
                                .ValorCobradoDiaAtraso = Format(0, "0000000000000")
                                .DataLimiteConcessaoDesconto = Format(0, "000000")
                                .ValorDesconto = Format(0, "0000000000000")
                                .ValorIOF = Format(0, "0000000000000")
                                .ValorAbatimento = Format(0, "0000000000000")
                                
                                cpfCnpj = "00000000000000"
                                If IsNull(Bdados.Tabela("CPF_CNPJ")) Then
                                    .IdentificacaoTipoInscricaoSacado = "99" 'OUTROS
                                ElseIf Len(Bdados.Tabela("CPF_CNPJ")) = 0 Then
                                    .IdentificacaoTipoInscricaoSacado = "99" 'OUTROS
                                Else
                                    
                                    cpfCnpj = retiraSeparadores(Bdados.Tabela("CPF_CNPJ"))
                                    If Len(cpfCnpj) < 14 Then
                                        .IdentificacaoTipoInscricaoSacado = "01" 'CPF
                                    Else
                                        .IdentificacaoTipoInscricaoSacado = "02" 'CNPJ
                                    End If
                                End If
                                .NumeroInscricaoSacado = Format(cpfCnpj, "00000000000000")
                                .NomeSacado = preencherComCaractere(Bdados.Tabela("NOME"), 40, " ")
                                'Dim Endereco As String
                                Endereco = Bdados.Tabela("LOGRADOURO") & " " & Bdados.Tabela("NOME_LOGRADOURO") & " " & Bdados.Tabela("NUMERO_ENDERECO") & " " & Bdados.Tabela("BAIRRO") & " " & Bdados.Tabela("CIDADE") & " " & Bdados.Tabela("ESTADO")
                                .EnderecoCompleto = preencherComCaractere(Endereco, 40, " ")
                                .PrimeiraMensagem = preencherComCaractere(" ", 12, " ")
                                'Dim ccep As String
                                ccep = retiraSeparadores(IIf(IsNull(Bdados.Tabela("CEP")), "00000000", Bdados.Tabela("CEP")))
                                ccep = preencherComCaractere(ccep, 8, "0")
                                .Cep = Left(ccep, 5)
                                .SufixoCEP = Right(ccep, 3)
                                .SacadorAvalistaSegundaMensagem = preencherComCaractere(" ", 60, " ")
                                .NumeroSequencialRegistro = Format(sequencialDetalheRemessa, "000000") ' SEQUENCIAL COMECANDO COM 2
                                sequencialDetalheRemessa = sequencialDetalheRemessa + 1
                                arquivo = .gerarDetalheRemessa(txtDiretorio, arquivo)
                                
                            End With
                        End If
                    End If
                Next Doc
            End If
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
    'arquivo = Replace(arquivo, "CB", "")
    'arquivo = Replace(arquivo, ".REM", "")
    'arquivo = arquivo & Format(Now, "yyyy")
    Dim Remessa As String
    Remessa = Format(Now, "DD") & Format(Now, "MM") & "01" & Format(Now, "YYYY")
    For Doc = 1 To Grid.ListItems.Count ' TOTAL DE TRIBUTOS
        If Grid.ListItems(Doc).Checked = True Then ' SE FOI MARCADO PARA REMESSA
            obrigacao = Grid.ListItems(Doc)
            'Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_STATUS_OBRIGACAO=12,TOC_REMESSA=" & arquivo & " WHERE TOC_COD_OBRIGACAO=" & obrigacao)
            'NAO MAIS MUDO PARA TRANSMITIDO
            Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_REMESSA=" & Remessa & " WHERE TOC_COD_OBRIGACAO=" & obrigacao)
            Bdados.Executa ("UPDATE TAB_CONTA_CONTRIBUINTE SET tcc_status_conta=4 WHERE tcc_codigo_conta=" & obrigacao)
            Bdados.Executa ("INSERT INTO TAB_BCP_REMESSA (COD_REMESSA, COD_OBRIGACAO) VALUES('" & Remessa & "'," & obrigacao & ")")
        End If
    Next Doc
End Sub

Private Sub cmdListarDam_Click()
   'teste remover depois
   'Dim nnumero As String
   'nnumero = Imposto.GeraNumCorrelativo(1, 199)  'BASA NOSSO NUMERO
    Dim erro As String
    Dim errados As Integer
    cmdVISUAL1.Enabled = False
    sqlO = "SELECT NUM_DOCUMENTO AS DOCUMENTO,COD_CLIENTE,CPF_CNPJ,NOME,STATUS,VALOR_ATUAL AS VALOR, DATA_EMISSAO AS EMISSAO, DATA_VENCIMENTO AS VENCTO, TOC_NOSSO_NUMERO AS NUMERO FROM VIS_CONTA_CONTRIBUINTE WHERE STATUS=2 AND (TOC_REMESSA IS NULL OR  TOC_REMESSA=''  OR  TOC_REMESSA ='1') ORDER BY TOC_NOSSO_NUMERO"
    If Grid.Preencher(Bdados, sqlO) Then
    Else
    End If
    Dim i As Integer
    For i = 1 To Grid.ListItems.Count
            Grid.ListItems(i).Checked = True
    Next i
    erro = ""
    errados = 0
    For i = 1 To Grid.ListItems.Count ' TOTAL DE TRIBUTOS
    If Grid.ListItems(i).Checked = True Then ' SE FOI MARCADO PARA REMESSA
        obrigacao = Grid.ListItems(i)
        If Bdados.AbreTabela("SELECT * FROM VIS_CONTA_CONTRIBUINTE WHERE NUM_DOCUMENTO=" & obrigacao) Then
            
            If IsNull(Bdados.Tabela("ESTADO")) Or IsNull(Bdados.Tabela("CEP")) Or IsNull(Bdados.Tabela("LOGRADOURO")) Or IsNull(Bdados.Tabela("NOME_LOGRADOURO")) Or IsNull(Bdados.Tabela("BAIRRO")) Or IsNull(Bdados.Tabela("CIDADE")) Then
                errados = errados + 1
                erro = erro & obrigacao & vbCrLf
                'Mensagem ("DAM com problemas de endereço:" & obrigacao)
                'Exit Sub
            End If
        End If
    End If
    Next i
    If Len(Trim(erro)) > 0 Then
        Mensagem (errados & " DAMs com problemas de endereço: " & vbCrLf & erro)
        Exit Sub
    End If
    cmdVISUAL1.Enabled = True
    
End Sub
Private Sub cmdVISUAL6_Click()
    If obrigacao = "" Then
        Util.Informa ("Selecione a obrigação")
        Exit Sub
    End If
    If Util.Confirma("Deseja remover o DAM " & obrigacao & " da lista de remessa") = True Then
        Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_REMESSA='MOVED'" & " WHERE TOC_COD_OBRIGACAO=" & obrigacao)
    End If
End Sub
Private Function retornarData(data As String) As String
    Dim nd As String
    data = Replace(data, "/", "")
    nd = Right(data, 4) & "-" & Mid(data, 3, 2) & "-" & Left(data, 2)
    retornarData = nd
End Function

Private Sub cmdReceber_Click()
    'BCP RETORNO
    Dim Linha As String
    Dim diretorio As String
    Dim Rs As VSRecordset
    diretorio = txtCamminhoRemessa
    Dim bb As Boolean
    bb = False
    Dim docRet As Long, status As Integer, sob As Integer, scc As Integer, dataOcorrencia As String, dataCredito As String, tres As String
    tres = Left(NomeArquivo, 3)
    obgris = ""
    Dim nomeB As String
    If Temp.PegaParametro(Bdados, "BANCO ARRECADACAO") = "1" Then 'BB
        If tres = "CBR" Then
            NomeArquivo = Mid(NomeArquivo, 8, 8)
            nomeB = Format(NomeArquivo, "00000000")
            dataOcorrencia = Left(NomeArquivo, 2) & "/" & Mid(NomeArquivo, 3, 2) & "/" & Right(NomeArquivo, 4)
            NomeArquivo = Left(NomeArquivo, 4) & "00" & Right(NomeArquivo, 4)
            bb = True
        Else
            NomeArquivo = Replace(NomeArquivo, "CB", "")
            nomeB = NomeArquivo
            NomeArquivo = Replace(NomeArquivo, "RET", "")
            NomeArquivo = NomeArquivo & Format(Now, "YYYY")
            bb = False
        End If
    Else
        NomeArquivo = Replace(NomeArquivo, "CB", "")
        NomeArquivo = Replace(NomeArquivo, "RET", "")
        NomeArquivo = NomeArquivo & Format(Now, "YYYY")
        nomeB = NomeArquivo
    End If
    
    Open diretorio For Input As #1
    Do While Not EOF(1)
        Line Input #1, Linha
        If Left(Linha, 1) = 0 Or Left(Linha, 1) = 9 Then
            If Mid(Linha, 89, 6) = "BRASIL" Then
                nomeB = Mid(Linha, 95, 4) & "00" & txtAno
            ElseIf Mid(Linha, 3, 7) = "RETORNO" Then 'BASA
                nomeB = Mid(Linha, 95, 4) & txtAno 'Data DDMM gravacao arquivo - 095 - 98 header retorno
            End If
            
        Else
            If Temp.PegaParametro(Bdados, "BANCO ARRECADACAO") = "3" Then 'BASA
                docRet = Mid(Linha, 117, 10) ' 117 a 126
                status = Mid(Linha, 109, 2) ' 109 a 110
                dataOcorrencia = Mid(Linha, 111, 6) '111 a 116
                dataCredito = Mid(Linha, 296, 6) '296 a 301
                dataOcorrencia = gerarDataRetorno(dataOcorrencia)
                dataCredito = gerarDataRetorno(dataCredito)
                sob = 0
                If status = 2 Then
                    sob = 14
                    scc = 5
                    dataCredito = ""
                ElseIf status = 6 Or status = 9 Or status = 10 Or status = 15 Or status = 17 Then
                    sob = 3
                End If
            
            ElseIf Temp.PegaParametro(Bdados, "BANCO ARRECADACAO") = "1" Then 'BB codo/pinheiro
                If bb = True Then
                    docRet = Mid(Linha, 71, 10) ' 71 a 80
                    status = Mid(Linha, 109, 2) ' 109 a 110
                    dataOcorrencia = Mid(Linha, 111, 6) '111 a 116
                    dataCredito = Mid(Linha, 176, 6) '176 a 181
                    dataOcorrencia = gerarDataRetorno(dataOcorrencia)
                    dataCredito = gerarDataRetorno(dataCredito)
                    sob = 3
                Else
                    docRet = Mid(Linha, 71, 11) ' 71 a 81
                    status = Mid(Linha, 109, 2) ' 109 a 110
                    dataOcorrencia = Mid(Linha, 111, 6) '111 a 116
                    dataCredito = Mid(Linha, 296, 6) '296 a 301
                    dataOcorrencia = gerarDataRetorno(dataOcorrencia)
                    dataCredito = gerarDataRetorno(dataCredito)
                
                    If status = 2 Or status = 10 Then
                        sob = 14
                        scc = 5
                        dataCredito = ""
                    ElseIf status = 6 Or status = 17 Then
                        sob = 3
                    End If
                End If
            Else 'GRAJAU BRADESCO
                docRet = Mid(Linha, 71, 11) ' 71 a 81
                status = Mid(Linha, 109, 2) ' 109 a 110
                dataOcorrencia = Mid(Linha, 111, 6) '111 a 116
                dataCredito = Mid(Linha, 296, 6) '296 a 301
                dataOcorrencia = gerarDataRetorno(dataOcorrencia)
                dataCredito = gerarDataRetorno(dataCredito)
                
                If status = 2 Then
                    sob = 14
                    scc = 5
                    dataCredito = ""
                ElseIf status = 6 Or status = 17 Or status = 10 Then
                    sob = 3
                End If
            End If
            If sob = 3 Then
                'so atualizado se for para pago
                Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_STATUS_OBRIGACAO=" & sob & " WHERE TOC_COD_OBRIGACAO=" & docRet)
                Bdados.Executa ("UPDATE TAB_CONTA_CONTRIBUINTE SET tcc_status_conta=" & scc & "WHERE tcc_codigo_conta=" & docRet)
            End If
            If Bdados.AbreTabela("SELECT COD_OBRIGACAO FROM TAB_BCP_RETORNO WHERE STATUS=" & status & " AND COD_OBRIGACAO=" & docRet, Rs) = False Then 'SE AINDA NAO EXISTE RETORNO
                Bdados.Executa ("INSERT INTO TAB_BCP_RETORNO (COD_RETORNO, COD_OBRIGACAO,STATUS,DATA_OCRRENCIA_BANCO,DATA_CREDITO) VALUES('" & retiraSeparadores(nomeB) & "','" & docRet & "'," & status & ",'" & dataOcorrencia & "','" & dataCredito & "')")
            End If
            obgris = obgris & docRet & ","
    
        End If
    Loop
    Close #1
    
    sqlO = "SELECT NUM_DOCUMENTO,COD_CLIENTE,CPF_CNPJ,NOME,VALOR_ATUAL FROM VIS_CONTA_CONTRIBUINTE WHERE NUM_DOCUMENTO IN(" & obgris
    sqlO = Left(sqlO, Len(sqlO) - 1)
    sqlO = sqlO & ")"
    If grdRETORNO.Preencher(Bdados, sqlO) Then
    Else
    End If
End Sub
Private Function gerarDataRetorno(dataArquivo As String) As String
    gerarDataRetorno = Left(dataArquivo, 2) & "/" & Mid(dataArquivo, 3, 2) & "/20" & Right(dataArquivo, 2)
End Function
'FIM BCP
Private Sub cmdVISUAL1_Click()
    'On Error GoTo Err
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
        'GoTo Vai
    End If
        If Confirma("Confirma geração do arquivos remessa?") Then
'Vai:
            
            For i = 1 To Grid.ListItems.Count
                If Grid.ListItems(i).Checked Then
                    'If Grid.ListItems(i).SubItems(2) = "" Then
                      '  Mensagem "O Contribuinte do DAM de numero " & Grid.ListItems(i) & " nao possui CPF/CNPJ, favor atualize seu cadastro"
                       ' Exit Sub
                    'End If
                End If
                If Grid.ListItems(i).Checked Then
                   If Grid.ListItems(i).SubItems(4) <> 2 Then
                        Dim S As String
                        S = Grid.ListItems(i).SubItems(4)
                        If S = 12 Then
                            S = "TRANSMITIDO"
                        ElseIf S = 14 Then
                            S = "REGISTRO BANCARIO"
                        ElseIf S = 8 Then
                            S = "CANCELADO"
                        ElseIf S = 3 Then
                            S = "PAGO"
                        End If
                        
                        Mensagem "O DAM de numero " & Grid.ListItems(i) & " nao pode ser transmitido por estar: " & S
                        Exit Sub
                    Else
                        If bcpBanco = 3 Then 'basa
                            Dim obr As String
                            obr = Grid.ListItems(i)
                            
                            'If Bdados.AbreTabela("SELECT TOC_NOSSO_NUMERO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO=" & obr) Then
                             '   If IsNull(Bdados.Tabela("TOC_NOSSO_NUMERO")) Then
                              '      Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_NOSSO_NUMERO=" & GeraBasaNossoNumero & " WHERE TOC_COD_OBRIGACAO=" & obr)
                               ' End If
                            'End If
                            
                            
                        End If
                    End If
                End If
            Next i
            montarHeder
            montarDetalhe
            montarTrailer
            transmitirRemessa
            
            Grid.ListItems.Clear
            limparCampos
            txtNumero = txtNumero + 1
            txtNumero.SetFocus
        End If
        
    End If
    Avisa "Arquivos gerados com sucesso."
    If Me.Tag = "" Then
       ' cmdBuscar_Especial_Click
    Else
        Unload Me
    End If
'Err:
        
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
Private Sub cmdVISUAL2_Click()
    Unload Me
End Sub
Private Sub cmdVISUAL3_Click()
    With CommonDialog2
        .DialogTitle = "Selecione o um diretorio do bradesco"
        .ShowOpen
        If .FileName <> "" Then
            txtDiretorio = .FileName
        End If
    End With
End Sub

Private Sub cmdVISUAL4_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL5_Click()
    Unload Me
End Sub


Private Function retiraSeparadores(Valor As String) As String
    Valor = Replace(Valor, ",", "")
    Valor = Replace(Valor, ".", "")
    Valor = Replace(Valor, "-", "")
    Valor = Replace(Valor, "/", "")
    Valor = Replace(Valor, "_", "")
    Valor = Replace(Valor, "-", "")
    retiraSeparadores = Valor
End Function

Private Sub Grid_Click()
    obrigacao = Grid.ListItems(Grid.SelectedItem.Index)
    exibe
End Sub


Private Sub txtDesconto_LostFocus()
    If txtDesconto <> "" Then
    Dim ValorTotal As Currency
        ValorTotal = CCur(txtValorOriginal) + CCur(txtJuros) + CCur(txtMulta)
        ValorDesconto = ValorTotal * (CCur(txtDesconto / 100))
        txtDescontoReal = ValorDesconto
    Else
        txtDescontoReal = 0
    End If
End Sub

Private Sub txtDescontoReal_LostFocus()
    Dim m As Double
    m = CCur(txtDescontoReal) * 100
    m = CCur(m) / CCur(txtValorOriginal)
    txtDesconto = CCur(m)
End Sub

Private Sub txtDocumento_LostFocus()
    Dim Sql As String
    obrigacao = txtDocumento
    exibe
End Sub
Private Sub exibe()
    If Len(obrigacao) > 0 Then
        If Bdados.AbreTabela("SELECT * FROM VIS_CONTA_CONTRIBUINTE WHERE NUM_DOCUMENTO=" & obrigacao) Then
            codigoStatus = Bdados.Tabela("STATUS")
            txtNome = Bdados.Tabela("NOME")
            txtValor = Bdados.Tabela("VALOR_ATUAL")
            txtObservacao = IIf(IsNull(Bdados.Tabela("OBSERVACAO")), "", Bdados.Tabela("OBSERVACAO"))
            txtDocumento = Bdados.Tabela("NUM_DOCUMENTO")
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

Private Sub txtJurosMargem_LostFocus()
    If txtJurosMargem <> "" Then
        txtJuros = CCur(txtValorOriginal) * (CCur(txtJurosMargem / 100))
    Else
        txtJuros = 0
    End If
End Sub

Private Sub txtMultaMargem_LostFocus()
    If txtMultaMargem <> "" Then
        txtMulta = CCur(txtValorOriginal) * (CCur(txtMultaMargem / 100))
    Else
        txtMulta = 0
    End If
End Sub
