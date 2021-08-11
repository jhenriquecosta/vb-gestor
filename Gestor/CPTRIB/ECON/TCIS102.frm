VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.1#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCIS102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIS102"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   ControlBox      =   0   'False
   Icon            =   "TCIS102.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   60
      ScaleHeight     =   570
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCIS102.frx":08CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1138
      Icone           =   "TCIS102.frx":29ED
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   26
      Top             =   7650
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   847
      Begin VTOcx.cmdVISUAL cmdPesq 
         Height          =   375
         Index           =   1
         Left            =   4170
         TabIndex        =   120
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8790
         TabIndex        =   25
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   7635
         TabIndex        =   24
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6480
         TabIndex        =   23
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   5325
         TabIndex        =   22
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin ActiveTabs.SSActiveTabs tabCadastro 
      Height          =   6870
      Left            =   -15
      TabIndex        =   27
      Top             =   645
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   12118
      _Version        =   131082
      TabCount        =   7
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
      Tabs            =   "TCIS102.frx":342F
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
         Height          =   6480
         Left            =   30
         TabIndex        =   33
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   11430
         _Version        =   131082
         TabGuid         =   "TCIS102.frx":35CC
         Begin ActiveTabs.SSActiveTabs TabAnuncio 
            Height          =   6255
            Left            =   60
            TabIndex        =   34
            Top             =   -90
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   11033
            _Version        =   131082
            TabCount        =   2
            TabOrientation  =   2
            Tabs            =   "TCIS102.frx":35F4
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel8 
               Height          =   5865
               Left            =   30
               TabIndex        =   35
               Top             =   30
               Width           =   9795
               _ExtentX        =   17277
               _ExtentY        =   10345
               _Version        =   131082
               TabGuid         =   "TCIS102.frx":368A
               Begin VTOcx.cmdVISUAL cmdAdAnuncio 
                  Height          =   375
                  Left            =   -1980
                  TabIndex        =   36
                  Top             =   1545
                  Visible         =   0   'False
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   661
                  Caption         =   "&Adicionar Anúncio"
                  Acao            =   1
                  Enabled         =   0   'False
                  CorBorda        =   8421504
                  CorFrente       =   16384
               End
               Begin VTOcx.fraVISUAL fraAnu 
                  CausesValidation=   0   'False
                  Height          =   1920
                  Left            =   60
                  TabIndex        =   37
                  Top             =   150
                  Width           =   9675
                  _ExtentX        =   17066
                  _ExtentY        =   3387
                  Altura          =   1905
                  Caption         =   " Dados do Veículo de Divulgacão"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Begin VTOcx.cboVISUAL cboItem 
                     Height          =   510
                     Left            =   75
                     TabIndex        =   45
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
                     Left            =   4275
                     TabIndex        =   44
                     Top             =   1365
                     Width           =   1380
                     _ExtentX        =   2434
                     _ExtentY        =   847
                     Caption         =   "Mutiplicador"
                     Text            =   ""
                     Restricao       =   3
                     AlinhamentoRotulo=   1
                     AlinhamentoTexto=   1
                  End
                  Begin VTOcx.txtVISUAL txtValorApagar 
                     Height          =   480
                     Left            =   8340
                     TabIndex        =   43
                     Top             =   1365
                     Width           =   1245
                     _ExtentX        =   2196
                     _ExtentY        =   847
                     Caption         =   "Valor"
                     Text            =   ""
                     Enabled         =   0   'False
                     Formato         =   5
                     Restricao       =   3
                     AlinhamentoRotulo=   1
                     AlinhamentoTexto=   1
                  End
                  Begin VTOcx.txtVISUAL txtValor 
                     Height          =   480
                     Left            =   7080
                     TabIndex        =   42
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
                     Left            =   90
                     TabIndex        =   41
                     Top             =   1365
                     Width           =   2670
                     _ExtentX        =   4710
                     _ExtentY        =   847
                     Caption         =   "Descrição"
                     Text            =   ""
                     AlinhamentoRotulo=   1
                     RetirarMascara  =   0   'False
                  End
                  Begin VTOcx.txtVISUAL txtArea 
                     Height          =   480
                     Left            =   2760
                     TabIndex        =   40
                     Top             =   1365
                     Width           =   1515
                     _ExtentX        =   2672
                     _ExtentY        =   847
                     Caption         =   "Área Total"
                     Text            =   ""
                     Restricao       =   3
                     AlinhamentoRotulo=   1
                     AlinhamentoTexto=   1
                  End
                  Begin VTOcx.cboVISUAL cboMovimento 
                     Height          =   510
                     Left            =   75
                     TabIndex        =   39
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
                     Left            =   5655
                     TabIndex        =   38
                     Top             =   1365
                     Width           =   1425
                     _ExtentX        =   2514
                     _ExtentY        =   847
                     Caption         =   "Data Instalação"
                     Text            =   ""
                     Formato         =   0
                     Restricao       =   2
                     AlinhamentoRotulo=   1
                     AlinhamentoTexto=   1
                  End
               End
               Begin VTOcx.fraVISUAL fraVISUAL5 
                  Height          =   1545
                  Left            =   60
                  TabIndex        =   46
                  Top             =   2115
                  Width           =   9675
                  _ExtentX        =   17066
                  _ExtentY        =   2725
                  Altura          =   1905
                  Caption         =   " Localização do Imóvel"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Begin VTOcx.cboVISUAL cboBairroAnuncio 
                     Height          =   315
                     Left            =   990
                     TabIndex        =   51
                     Top             =   1050
                     Width           =   8565
                     _ExtentX        =   15108
                     _ExtentY        =   556
                     Caption         =   "Bairro"
                     Text            =   ""
                     AutoFocaliza    =   0   'False
                     Editavel        =   -1  'True
                  End
                  Begin VTOcx.txtVISUAL txtInscImob 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   50
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
                     TabIndex        =   49
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
                     TabIndex        =   48
                     Top             =   705
                     Width           =   5805
                     _ExtentX        =   10239
                     _ExtentY        =   556
                     Caption         =   ""
                     Text            =   ""
                     AutoFocaliza    =   0   'False
                     Editavel        =   -1  'True
                  End
                  Begin VTOcx.cmdVISUAL cmdVISUAL1 
                     Height          =   315
                     Left            =   3780
                     TabIndex        =   47
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
               End
               Begin VTOcx.fraVISUAL fraVISUAL6 
                  Height          =   1545
                  Left            =   60
                  TabIndex        =   52
                  Top             =   3720
                  Width           =   9675
                  _ExtentX        =   17066
                  _ExtentY        =   2725
                  Altura          =   1905
                  Caption         =   " Obs"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Begin VB.TextBox txtObs 
                     Appearance      =   0  'Flat
                     Height          =   1065
                     Left            =   90
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   53
                     Top             =   360
                     Width           =   9495
                  End
               End
               Begin VTOcx.cmdVISUAL CmdLimparAnuncio 
                  Height          =   375
                  Left            =   8580
                  TabIndex        =   54
                  Top             =   5340
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   661
                  Caption         =   "&Limpar"
                  Acao            =   6
                  CorBorda        =   8421504
                  CorFrente       =   16384
               End
               Begin VTOcx.cmdVISUAL cmdExcluir 
                  Height          =   375
                  Left            =   7425
                  TabIndex        =   55
                  Top             =   5340
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   661
                  Caption         =   "&Excluir"
                  Acao            =   2
                  CorBorda        =   8421504
                  CorFrente       =   16384
               End
               Begin VTOcx.cmdVISUAL cmdSalvarAnuncio 
                  Height          =   375
                  Left            =   6270
                  TabIndex        =   56
                  Top             =   5340
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   661
                  Caption         =   "&Salvar"
                  Acao            =   3
                  CorBorda        =   8421504
                  CorFrente       =   16384
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
               Height          =   5865
               Left            =   30
               TabIndex        =   57
               Top             =   30
               Width           =   9795
               _ExtentX        =   17277
               _ExtentY        =   10345
               _Version        =   131082
               TabGuid         =   "TCIS102.frx":36B2
               Begin VTOcx.grdVISUAL grdAnuncio 
                  Height          =   5985
                  Left            =   45
                  TabIndex        =   58
                  Top             =   135
                  Width           =   9750
                  _ExtentX        =   17198
                  _ExtentY        =   10557
                  CorBorda        =   32768
                  CorTitulo       =   32768
                  CorCaption      =   16777215
                  CorDica         =   32768
                  OcultarRodape   =   -1  'True
               End
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6480
         Left            =   30
         TabIndex        =   28
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   11430
         _Version        =   131082
         TabGuid         =   "TCIS102.frx":36DA
         Begin VB.Frame FraDados 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   3825
            Left            =   45
            TabIndex        =   122
            Top             =   2895
            Width           =   9765
            Begin VTOcx.fraVISUAL fraVISUAL3 
               Height          =   3435
               Left            =   15
               TabIndex        =   123
               Top             =   30
               Width           =   9735
               _ExtentX        =   17171
               _ExtentY        =   6059
               Altura          =   1905
               Caption         =   " Atividade"
               CorTexto        =   16777215
               CorFaixa        =   32768
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.fraVISUAL FraVariavel 
                  Height          =   300
                  Left            =   60
                  TabIndex        =   150
                  Top             =   315
                  Width           =   4620
                  _ExtentX        =   8149
                  _ExtentY        =   529
                  Status          =   1
                  Altura          =   1905
                  Caption         =   " Variáveis de Cadastro"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Begin VTOcx.txtVISUAL txtvAnuncio 
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   154
                     Top             =   435
                     Width           =   2730
                     _ExtentX        =   4815
                     _ExtentY        =   503
                     Caption         =   "Anúncio"
                     Text            =   ""
                     Formato         =   5
                     Restricao       =   2
                  End
                  Begin VTOcx.txtVISUAL txtvQtdItems 
                     Height          =   285
                     Left            =   285
                     TabIndex        =   153
                     Top             =   765
                     Width           =   3765
                     _ExtentX        =   6641
                     _ExtentY        =   503
                     Caption         =   "Quantidade de Item"
                     Text            =   ""
                     Restricao       =   2
                  End
                  Begin VTOcx.txtVISUAL txtvDataFormatura 
                     Height          =   285
                     Left            =   645
                     TabIndex        =   152
                     Top             =   1095
                     Width           =   3405
                     _ExtentX        =   6006
                     _ExtentY        =   503
                     Caption         =   "Data Formatura"
                     Text            =   ""
                     Formato         =   0
                     Restricao       =   2
                  End
                  Begin VTOcx.cboVISUAL CbovFuncionarioSUS 
                     Height          =   315
                     Left            =   465
                     TabIndex        =   151
                     Top             =   1425
                     Width           =   2460
                     _ExtentX        =   4339
                     _ExtentY        =   556
                     Caption         =   "Funcionário - SUS"
                     Text            =   ""
                     AutoFocaliza    =   0   'False
                  End
               End
               Begin VTOcx.cboVISUAL cboSitCad 
                  Height          =   315
                  Left            =   105
                  TabIndex        =   149
                  Tag             =   "Situação Cadastral"
                  Top             =   2895
                  Width           =   4770
                  _ExtentX        =   8414
                  _ExtentY        =   556
                  Caption         =   "Situação Cadastral"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
               End
               Begin VTOcx.cmdVISUAL cmdAdAtiv 
                  Height          =   300
                  Left            =   6165
                  TabIndex        =   148
                  Top             =   1995
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   529
                  Caption         =   ""
                  Acao            =   5
                  CorBorda        =   8421504
                  CorFrente       =   16384
               End
               Begin VTOcx.cboVISUAL cboPonto 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   146
                  Top             =   3510
                  Width           =   4845
                  _ExtentX        =   8546
                  _ExtentY        =   556
                  Caption         =   "Ponto Recepção"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
               End
               Begin VTOcx.txtVISUAL txtFator 
                  Height          =   480
                  Left            =   6600
                  TabIndex        =   145
                  Top             =   1815
                  Visible         =   0   'False
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   847
                  Caption         =   ""
                  Text            =   ""
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtDtInicio 
                  Height          =   480
                  Left            =   7965
                  TabIndex        =   144
                  Tag             =   "Início da Atividade"
                  Top             =   1815
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   847
                  Caption         =   "Início Atividade"
                  Text            =   ""
                  Formato         =   0
                  Restricao       =   2
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.cboVISUAL cboIsento 
                  Height          =   510
                  Left            =   3810
                  TabIndex        =   143
                  Tag             =   "Isento"
                  Top             =   1245
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   900
                  Caption         =   "Isento"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboAtivSecund2 
                  Height          =   510
                  Left            =   4770
                  TabIndex        =   142
                  Top             =   3735
                  Width           =   4950
                  _ExtentX        =   8731
                  _ExtentY        =   900
                  Caption         =   "Atividade Secundária"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboAtivSecund 
                  Height          =   510
                  Left            =   75
                  TabIndex        =   141
                  Top             =   3435
                  Width           =   4650
                  _ExtentX        =   8202
                  _ExtentY        =   900
                  Caption         =   "Atividade Secundária"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.fraVISUAL fraVISUAL4 
                  Height          =   885
                  Left            =   4755
                  TabIndex        =   137
                  Top             =   900
                  Width           =   4890
                  _ExtentX        =   8625
                  _ExtentY        =   1561
                  Altura          =   1905
                  Caption         =   " Somente para Autônomos"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Begin VTOcx.txtVISUAL txtRegistro 
                     Height          =   480
                     Left            =   1260
                     TabIndex        =   140
                     Top             =   300
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   847
                     Caption         =   "Nº Registro"
                     Text            =   ""
                     AlinhamentoRotulo=   1
                  End
                  Begin VTOcx.txtVISUAL txtConselho 
                     Height          =   480
                     Left            =   120
                     TabIndex        =   139
                     Top             =   300
                     Width           =   1110
                     _ExtentX        =   1958
                     _ExtentY        =   847
                     Caption         =   "Conselho"
                     Text            =   ""
                     AlinhamentoRotulo=   1
                  End
                  Begin VTOcx.cboVISUAL cboNivel 
                     Height          =   510
                     Left            =   2685
                     TabIndex        =   138
                     Top             =   285
                     Width           =   1965
                     _ExtentX        =   3466
                     _ExtentY        =   900
                     Caption         =   "Nível de Instrução"
                     Text            =   ""
                     AutoFocaliza    =   0   'False
                     Alinhamento     =   1
                  End
               End
               Begin VTOcx.txtVISUAL txtEmpregados 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   136
                  Tag             =   "Empregados"
                  Top             =   1260
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   847
                  Caption         =   "Empregados"
                  Text            =   ""
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.cboVISUAL cboClassAtiv 
                  Height          =   510
                  Left            =   2250
                  TabIndex        =   135
                  Tag             =   "Classificação de Atividade"
                  Top             =   735
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   900
                  Caption         =   "Classificação de Atividade"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboNatJur 
                  Height          =   510
                  Left            =   60
                  TabIndex        =   134
                  Tag             =   "Natureza Jurídica"
                  Top             =   735
                  Width           =   2070
                  _ExtentX        =   3651
                  _ExtentY        =   900
                  Caption         =   "Natureza Jurídica"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboAtivPoder 
                  Height          =   510
                  Left            =   4740
                  TabIndex        =   133
                  Tag             =   "Atividade Exercida Poder"
                  Top             =   360
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   900
                  Caption         =   "Atv.Exercida Poder"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboObrigIss 
                  Height          =   510
                  Left            =   6465
                  TabIndex        =   132
                  Tag             =   "Obrigação do ISSQN"
                  Top             =   360
                  Width           =   3225
                  _ExtentX        =   5689
                  _ExtentY        =   900
                  Caption         =   "Tipo de recolhimento do ISSQN"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboPorte 
                  Height          =   510
                  Left            =   1320
                  TabIndex        =   131
                  Tag             =   "Porte da Empresa"
                  Top             =   1245
                  Width           =   2490
                  _ExtentX        =   4392
                  _ExtentY        =   900
                  Caption         =   "Porte da Empresa"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboAtivServ 
                  Height          =   510
                  Left            =   75
                  TabIndex        =   130
                  Tag             =   "Atividade Principal"
                  Top             =   1800
                  Width           =   6075
                  _ExtentX        =   10716
                  _ExtentY        =   900
                  Caption         =   "Atividade Principal"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboTipoCadastro 
                  Height          =   315
                  Left            =   5025
                  TabIndex        =   129
                  Top             =   2385
                  Width           =   4605
                  _ExtentX        =   8123
                  _ExtentY        =   556
                  Caption         =   "Tipo de Cadastro"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
               End
               Begin VTOcx.txtVISUAL txtDataReabertura 
                  Height          =   285
                  Left            =   6960
                  TabIndex        =   128
                  Tag             =   "Início da Atividade"
                  Top             =   3660
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
                  Left            =   3735
                  TabIndex        =   127
                  Tag             =   "Início da Atividade"
                  Top             =   3570
                  Visible         =   0   'False
                  Width           =   3120
                  _ExtentX        =   5503
                  _ExtentY        =   503
                  Caption         =   "Dt.Encerramento"
                  Text            =   ""
                  Formato         =   0
                  Restricao       =   2
               End
               Begin VTOcx.cboVISUAL cboMatrizFilial 
                  Height          =   315
                  Left            =   765
                  TabIndex        =   126
                  Top             =   2445
                  Width           =   4110
                  _ExtentX        =   7250
                  _ExtentY        =   556
                  Caption         =   "Matriz/Filial"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
               End
               Begin VTOcx.cboVISUAL cboSitAlvara 
                  Height          =   315
                  Left            =   5340
                  TabIndex        =   125
                  Top             =   3060
                  Width           =   4290
                  _ExtentX        =   7567
                  _ExtentY        =   556
                  Caption         =   "Situação do Alvará"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
               End
               Begin VTOcx.txtVISUAL txtInicioPrestacaoServico 
                  Height          =   285
                  Left            =   5025
                  TabIndex        =   124
                  Top             =   2745
                  Width           =   4575
                  _ExtentX        =   8070
                  _ExtentY        =   503
                  Caption         =   "Início da Prestação do Serviço"
                  Text            =   ""
                  Formato         =   0
                  Restricao       =   2
               End
               Begin VTOcx.cboVISUAL cboRamo 
                  Height          =   315
                  Left            =   375
                  TabIndex        =   147
                  Top             =   1980
                  Width           =   4335
                  _ExtentX        =   7646
                  _ExtentY        =   556
                  Caption         =   "Ramo Atividade"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
               End
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   1380
            Left            =   60
            TabIndex        =   29
            Top             =   30
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   2434
            Altura          =   1905
            Caption         =   " Contribuinte"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
               Height          =   315
               Left            =   9330
               TabIndex        =   155
               TabStop         =   0   'False
               Top             =   480
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtIncriAuxiliar 
               Height          =   480
               Left            =   5385
               TabIndex        =   5
               Top             =   825
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   847
               Caption         =   "Inscri.Auxiliar"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtFantasia 
               Height          =   480
               Left            =   90
               TabIndex        =   4
               Top             =   825
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   847
               Caption         =   "Nome Fantasia"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtAutoriza 
               Height          =   480
               Left            =   9705
               TabIndex        =   10
               Top             =   825
               Visible         =   0   'False
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   847
               Caption         =   "Nº Autorização"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtIm 
               Height          =   480
               Left            =   3195
               TabIndex        =   2
               Top             =   300
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   847
               Caption         =   "Ins. Municipal"
               Text            =   ""
               Formato         =   8
               Restricao       =   2
               AlinhamentoRotulo=   1
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtFoneFax 
               Height          =   480
               Left            =   7020
               TabIndex        =   7
               Top             =   825
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   847
               Caption         =   "Fone/Fax"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCateg 
               Height          =   480
               Left            =   7380
               TabIndex        =   9
               Top             =   1455
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
               TabIndex        =   0
               Top             =   300
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   847
               Caption         =   "Nº Protocolo"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCnh 
               Height          =   480
               Left            =   6135
               TabIndex        =   8
               Top             =   1410
               Visible         =   0   'False
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   847
               Caption         =   "CNH"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtRuc 
               Height          =   480
               Left            =   8370
               TabIndex        =   6
               Top             =   825
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   847
               Caption         =   "Ins. Estadual"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtRazao 
               Height          =   480
               Left            =   4440
               TabIndex        =   3
               Tag             =   "Nome ou Razão Social"
               Top             =   300
               Width           =   4920
               _ExtentX        =   8678
               _ExtentY        =   847
               Caption         =   "Nome ou Razão Social"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCgc 
               Height          =   480
               Left            =   1380
               TabIndex        =   1
               Top             =   300
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   847
               Caption         =   "CPF ou CNPJ"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1425
            Left            =   60
            TabIndex        =   30
            Top             =   1455
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   2514
            Altura          =   1905
            Caption         =   " Localização"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL cmdVISUAL3 
               Height          =   285
               Left            =   4155
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   525
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               Caption         =   ""
               Acao            =   5
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
            Begin VTOcx.cboVISUAL cboLogr 
               Height          =   510
               Left            =   5730
               TabIndex        =   15
               Top             =   315
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   900
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Editavel        =   -1  'True
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
            Begin VTOcx.cboVISUAL cboTipoLogr 
               Height          =   510
               Left            =   4500
               TabIndex        =   14
               Top             =   315
               Width           =   1260
               _ExtentX        =   2223
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
               Left            =   2730
               TabIndex        =   18
               Top             =   840
               Width           =   2745
               _ExtentX        =   4842
               _ExtentY        =   900
               Caption         =   "Distrito ou Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
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
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   6480
         Left            =   30
         TabIndex        =   59
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   11430
         _Version        =   131082
         TabGuid         =   "TCIS102.frx":3702
         Begin VTOcx.fraVISUAL fraTrans 
            Height          =   1395
            Left            =   60
            TabIndex        =   60
            Top             =   885
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   2461
            Altura          =   1905
            Caption         =   " Representante Legal"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Enabled         =   0   'False
            Begin VTOcx.txtVISUAL txtLicenca 
               Height          =   480
               Left            =   2685
               TabIndex        =   70
               Top             =   810
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   847
               Caption         =   "Licenciamento"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboUFTransp 
               Height          =   510
               Left            =   4020
               TabIndex        =   69
               Top             =   795
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   900
               Caption         =   "UF"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Enabled         =   0   'False
            End
            Begin VTOcx.txtVISUAL txtMunicipio 
               Height          =   480
               Left            =   120
               TabIndex        =   68
               Top             =   810
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   847
               Caption         =   "Município"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtAnoFabric 
               Height          =   480
               Left            =   5910
               TabIndex        =   67
               Top             =   300
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   847
               Caption         =   "Ano"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtPlaca 
               Height          =   480
               Left            =   6735
               TabIndex        =   66
               Top             =   300
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   847
               Caption         =   "Placa"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtChassi 
               Height          =   480
               Left            =   7920
               TabIndex        =   65
               Top             =   300
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   847
               Caption         =   "Chassi"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtMarca 
               Height          =   480
               Left            =   1800
               TabIndex        =   64
               Top             =   300
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   847
               Caption         =   "Marca"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtModelo 
               Height          =   480
               Left            =   3660
               TabIndex        =   63
               Top             =   300
               Width           =   2220
               _ExtentX        =   3916
               _ExtentY        =   847
               Caption         =   "Modelo"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtVeiculo 
               Height          =   480
               Left            =   120
               TabIndex        =   62
               Top             =   300
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   847
               Caption         =   "Veículo"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtInicioAtividadeCarro 
               Height          =   480
               Left            =   4830
               TabIndex        =   61
               Top             =   810
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   847
               Caption         =   "Inicio Atividade "
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   0
               AlinhamentoRotulo=   1
            End
         End
         Begin VTOcx.cmdVISUAL cmdAdVeiculo 
            Height          =   375
            Left            =   60
            TabIndex        =   71
            Top             =   2325
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
            Height          =   3510
            Left            =   60
            TabIndex        =   72
            Top             =   2745
            Width           =   9750
            _ExtentX        =   17198
            _ExtentY        =   6191
            CorBorda        =   32768
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
         Begin Threed.SSCheck chkCad 
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   73
            Top             =   30
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   344
            _Version        =   196610
            Caption         =   "Cadastrar"
         End
         Begin VTOcx.cboVISUAL cboAtividadeVeiculo 
            Height          =   315
            Left            =   120
            TabIndex        =   74
            Top             =   480
            Width           =   9600
            _ExtentX        =   16933
            _ExtentY        =   556
            Caption         =   "Atividade Desempenhada"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   6480
         Left            =   30
         TabIndex        =   75
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   11430
         _Version        =   131082
         TabGuid         =   "TCIS102.frx":372A
         Begin VTOcx.fraVISUAL fraRepresen1 
            Height          =   1365
            Left            =   60
            TabIndex        =   76
            Top             =   525
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   2408
            Altura          =   1905
            Caption         =   " Representante Legal"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Enabled         =   0   'False
            Begin VTOcx.txtVISUAL txtImRepresentante 
               Height          =   480
               Left            =   120
               TabIndex        =   79
               Top             =   300
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   847
               Caption         =   "Im"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNomeRepresentante 
               Height          =   480
               Left            =   2145
               TabIndex        =   78
               Top             =   795
               Width           =   7485
               _ExtentX        =   13203
               _ExtentY        =   847
               Caption         =   "Nome"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCpfRepresentante 
               Height          =   480
               Left            =   120
               TabIndex        =   77
               Top             =   795
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
         Begin VTOcx.fraVISUAL fraRepresen2 
            Height          =   1410
            Left            =   60
            TabIndex        =   80
            Top             =   1920
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   2487
            Altura          =   1905
            Caption         =   " Endereço do Representante"
            CorTexto        =   16777215
            CorFaixa        =   32768
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
               TabIndex        =   87
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
               TabIndex        =   86
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
               TabIndex        =   85
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
               TabIndex        =   84
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
               TabIndex        =   83
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
               TabIndex        =   82
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
               TabIndex        =   81
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
            TabIndex        =   89
            Top             =   30
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   344
            _Version        =   196610
            Caption         =   "Cadastrar"
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   6480
         Left            =   30
         TabIndex        =   90
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   11430
         _Version        =   131082
         TabGuid         =   "TCIS102.frx":3752
         Begin VTOcx.fraVISUAL fraContador 
            Height          =   1590
            Left            =   1215
            TabIndex        =   91
            Top             =   495
            Width           =   7470
            _ExtentX        =   13176
            _ExtentY        =   2805
            Altura          =   1905
            Caption         =   " Dados do Contador"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Enabled         =   0   'False
            Begin VTOcx.cboVISUAL cboContador 
               Height          =   510
               Left            =   75
               TabIndex        =   95
               Top             =   435
               Width           =   7350
               _ExtentX        =   12965
               _ExtentY        =   900
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Enabled         =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCgcEscritorio 
               Height          =   480
               Left            =   5025
               TabIndex        =   94
               Top             =   990
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   847
               Caption         =   "CNPJ Escritório"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   2
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCpfContador 
               Height          =   480
               Left            =   75
               TabIndex        =   93
               Top             =   990
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   847
               Caption         =   "CPF"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   1
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCrcContador 
               Height          =   480
               Left            =   2550
               TabIndex        =   92
               Top             =   990
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   847
               Caption         =   "CRC"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
            End
         End
         Begin Threed.SSCheck chkCad 
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   96
            Top             =   30
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   344
            _Version        =   196610
            Caption         =   "Cadastrar"
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   6480
         Left            =   30
         TabIndex        =   97
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   11430
         _Version        =   131082
         TabGuid         =   "TCIS102.frx":377A
         Begin VTOcx.fraVISUAL fraSocio1 
            Height          =   900
            Left            =   60
            TabIndex        =   98
            Top             =   390
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   1588
            Altura          =   1905
            Caption         =   " Dados do Sócio"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtCargoSocio 
               Height          =   480
               Left            =   7470
               TabIndex        =   101
               Top             =   315
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   847
               Caption         =   "Cargo"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtNomeSocio 
               Height          =   480
               Left            =   2100
               TabIndex        =   100
               Top             =   315
               Width           =   5355
               _ExtentX        =   9446
               _ExtentY        =   847
               Caption         =   "Nome"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCpfSocio 
               Height          =   480
               Left            =   75
               TabIndex        =   99
               Top             =   315
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
         Begin VTOcx.fraVISUAL fraSocio2 
            Height          =   1410
            Left            =   60
            TabIndex        =   102
            Top             =   1320
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   2487
            Altura          =   1905
            Caption         =   " Endereço do Sócio"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboUFSocio 
               Height          =   510
               Left            =   8820
               TabIndex        =   110
               Top             =   810
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   900
               Caption         =   "UF"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtCidadeSocio 
               Height          =   480
               Left            =   6255
               TabIndex        =   109
               Top             =   825
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   847
               Caption         =   "Cidade"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtTelSocio 
               Height          =   480
               Left            =   4905
               TabIndex        =   108
               Top             =   825
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   847
               Caption         =   "Telefone"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboBairroSocio 
               Height          =   510
               Left            =   90
               TabIndex        =   107
               Top             =   810
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   900
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtCompSocio 
               Height          =   480
               Left            =   5625
               TabIndex        =   106
               Top             =   315
               Width           =   4020
               _ExtentX        =   7091
               _ExtentY        =   847
               Caption         =   "Complemento"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtNumSocio 
               Height          =   480
               Left            =   4950
               TabIndex        =   105
               Top             =   315
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   847
               Caption         =   "Nº"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cboVISUAL cboLogrSocio 
               Height          =   510
               Left            =   1635
               TabIndex        =   104
               Top             =   300
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   900
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
               Editavel        =   -1  'True
            End
            Begin VTOcx.cboVISUAL cboTipoLogrSocio 
               Height          =   510
               Left            =   75
               TabIndex        =   103
               Top             =   300
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   900
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
         End
         Begin VTOcx.cmdVISUAL cmdAdEdif 
            Height          =   375
            Left            =   60
            TabIndex        =   111
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
            Height          =   3015
            Left            =   60
            TabIndex        =   112
            Top             =   3180
            Width           =   9750
            _ExtentX        =   17198
            _ExtentY        =   5318
            CorBorda        =   32768
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
         Begin Threed.SSCheck chkCad 
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   113
            Top             =   30
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   344
            _Version        =   196610
            Caption         =   "Cadastrar"
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel9 
         Height          =   6480
         Left            =   30
         TabIndex        =   114
         Top             =   30
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   11430
         _Version        =   131082
         TabGuid         =   "TCIS102.frx":37A2
         Begin VTOcx.grdVISUAL grdAtividade 
            Height          =   4620
            Left            =   105
            TabIndex        =   115
            Top             =   1545
            Width           =   9645
            _ExtentX        =   17013
            _ExtentY        =   8149
            CorBorda        =   32768
            Caption         =   "Atividades"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
         Begin VTOcx.cboVISUAL cboGrupoAtividade 
            Height          =   315
            Left            =   195
            TabIndex        =   116
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
            TabIndex        =   117
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
            TabIndex        =   118
            Top             =   1110
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
            Caption         =   "&Adicionar "
            Acao            =   3
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL2 
            Height          =   375
            Left            =   8490
            TabIndex        =   119
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
   End
End
Attribute VB_Name = "TCIS102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cadastro As VSImposto
Dim Conta As New ContaCorrente
Dim Transportador As eTransportador
Dim Contribuinte As eContribuinte
Dim Endereco As eEndereco
Dim Contador As eContador
Dim atividade As atividade
Dim Imovel As eImovel
Dim Socio As eSocio
Dim Representante As eRepresentante
Dim GraveiContrib  As Boolean
Dim InscricaoMunicipal As String
Dim CodMudanca As String
Dim Anuncio As eAnuncio
Dim DataBaixa As String
Dim icad As String
Sub PreencheTela(Im As String, Cpf As String)
    On Error GoTo trata
    If Trim(Im) = "" And Trim(Cpf) = "" Then Exit Sub
    Screen.MousePointer = 11
    With Contribuinte
        If .BuscarHistorico(Im, Cpf, True, Data, Motivo) Then
            txtIm = .Im
            If Trim(.Ic) <> Null Or .Ic <> "" Then
                txtIc = .Ic
                txtIC_LostFocus
            End If
            InscricaoMunicipal = txtIm
            txtCgc = .CgcCpf
            txtIncriAuxiliar = .ImAuxiliar
            txtcgc_LostFocus
            txtRazao = .Nome
            txtInicioPrestacaoServico = .InicioPrestacaoServico
            cboSitAlvara.SetarLinha .SituacaoAlvara, 1
            cboTipoCadastro.SetarLinha .TipoCadastro, 1
            txtvAnuncio = .VariavelAnuncio
            
            txtDataEncerramento = .Data_Encerramento
            txtDataReabertura = .Data_Reabertura
            txtvDataFormatura = .VariavelDataFormatura
            CbovFuncionarioSUS.SetarLinha .VariavelFuncionarioSUS, 1
            txtvQtdItems = .VariavelQuantidadeItem
            cboMatrizFilial.SetarLinha "" & .Matriz_Filial, 1
            txtFantasia = .Fantasia
            cboTipoLogr.SetarLinha .Logradouro
            cboLogr.SetarLinha .NomeLogradouro
            cboLogr.Text = .NomeLogradouro
            txtNum = .Numero
            txtComplemento = .Complemento
            cboBairro.SetarLinha .Bairro
            cboBairro.Text = .Bairro
            
            txtCep = .Cep
            txtFator = Format(Nvl(.FatorAlvara, 0), Const_Monetario)
            DoEvents
            txtDtInicio = .InicioAtividade
            txtCidade = .Cidade
            cboUF.SetarLinha .Uf
            cboUF.Text = .Uf
            cboRamo.SetarLinha .CodRamo, 1
            If Trim(.FoneFax) <> Null Or .FoneFax <> "" Then
                txtFoneFax = .FoneFax
            End If
            txtCateg = .Categoria
            txtAutoriza = .Autorizacao
            cboPonto.SetarLinha .PontoRecepcao
            cboNatJur.SetarLinha .CodNatureza, 1
            cboClassAtiv.SetarLinha .CodGrupo, 1
            cboAtivPoder.SetarLinha .CodAtivPoder, 1
            cboObrigIss.SetarLinha .TipoRecolhimentoIss, 1
            txtRuc = .Ruc
            cboImovel.SetarLinha .ImovelProprio, 1
            txtConselho = .Conselho
            If Trim(.Registro) <> "" And .Registro <> "0" Then txtRegistro = .Registro
            cboNivel.SetarLinha .NivelEscolar, 1
            txtEmpregados = .NumEmpregado
            cboPorte.SetarLinha .PorteEmpresa, 1
            txtProtocolo = .Protocolo
            cboSitCad.SetarLinha .CodSitCadastral, 1
            cboEstabelece.SetarLinha .Estabelecido, 1
            cboEstabelece_LostFocus
            txtIm.Enabled = False
            
            cboIsento.SetarLinha .Isento, 1
            cboAtivServ.SetarLinha .CodAtividade, 1
            cboAtivServ_LostFocus
            If Not IsNull(.CodAtividadeSec) Then
                cboAtivSecund.SetarLinha .CodAtividadeSec, 1
            End If
            If Not IsNull(.CodAtividadeTerc) Then
                cboAtivSecund2.SetarLinha .CodAtividadeTerc, 1
            End If
        Else
            Avisa "Contribuinte não encontrado."
            If txtIm = "" Then txtCgc.SetFocus Else txtIm.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    End With
    If Socio.PreencherGrd(grdSocio, InscricaoMunicipal) Then chkCad(0).Value = ssCBChecked: chkCad_Click 0, 1 'CARREGA TOSOS OS SOCIOS
    With Contador   'CARREGA DADOS DO CONTADOR
        If .Buscar(InscricaoMunicipal) Then
            chkCad(1).Value = ssCBChecked: chkCad_Click 1, True
            cboContador = .Contador
            txtCrcContador = .Crc
            If Len(.Cpf) > 14 Then
                txtCgcEscritorio = .Cpf
            ElseIf Len(.Cpf) = 11 Then
                txtCpfContador = .Cpf
            End If
            If CStr(.CGCEscritorio) = CStr(0) Then .CGCEscritorio = ""
            If Len(.CGCEscritorio) <= 14 And Trim(.CGCEscritorio) <> "" Then
                txtCpfContador = .CGCEscritorio
            Else
                txtCgcEscritorio = .CGCEscritorio
            End If
            txtCrcContador = .Crc
        End If
    End With
    With Representante 'CARREGA DADOS DO REPRESENTANTE
        If .Buscar(InscricaoMunicipal) Then
            chkCad(2).Value = ssCBChecked: chkCad_Click 2, 1
            txtCpfRepresentante = .Cpf
            txtNomeRepresentante = .Nome
            cboTipoLogrRepresentante.SetarLinha .TipoLogr
            cboLogrRepresentante = .Logr
            txtNumRepresentante = .Numero
            txtComplementoRepresentante = .Complemento
            cboBairroRepresentante.SetarLinha .Bairro
            txtTelefoneRepresentante = .Telefone
            txtCidadeRepresentante = .Cidade
            cboUfRepresentante = .Uf
            txtImRepresentante = .ImRepresentante
        End If
    End With
    If Transportador.PreencherGrd(grdVeiculo, InscricaoMunicipal) Then chkCad(3).Value = ssCBChecked: chkCad_Click 3, 1        'carrega todos os veiculos
    Contribuinte.PreencherAtividadeSecundarias grdAtividade, txtIm
    'CARREGA OS ANÚNCIOS
    If Anuncio.PreencherGrd(grdAnuncio, InscricaoMunicipal) Then
        chkCad(4).Value = ssCBChecked
        chkCad_Click 4, 1
        Dim iLACO As Integer
    End If
    
    Screen.MousePointer = 0
    DoEvents
    Exit Sub
trata:
    Erro Err.Number & " - " & Err.Description
    Screen.MousePointer = 0
    Exit Sub
    Resume
End Sub

Private Sub cboAtivServ_LostFocus()
    On Error Resume Next
    Dim RetFator As String
    Dim RetNivel     As String
    If Trim(cboAtivServ) = "" Then Exit Sub
    If atividade.BuscaFator(cboAtivServ, RetFator, RetNivel) Then
        txtFator.Visible = True
        cboNivel.SetarLinha RetNivel, 1
        txtFator.Caption = RetFator
        txtFator.Tag = "Fator"
        txtFator.SetFocus
        'txtFator = ""
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
        txtCgcEscritorio = "": txtCrcContador = "": txtCpfContador = ""
        txtCrcContador = Contribuinte.Registro
        If Len(Trim(Contribuinte.CgcCpf)) = 14 And Not IsNumeric(Contribuinte.CgcCpf) Then
            txtCpfContador = Contribuinte.CgcCpf
        Else
            txtCgcEscritorio = Contribuinte.CgcCpf
        End If
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
    Dim SQL As String
    txtValor = "0,00"
    txtValorApagar = "0,00"
    SQL = "SELECT * "
    SQL = SQL & " From TAB_PARAMETRO_DETALHE "
    SQL = SQL & " where  TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(cboMovimento.Coluna(0).Valor, tctexto)
    If Bdados.AbreTabela(SQL) Then
        cboItem.Enabled = True
        cboItem.Preencher Bdados, "SELECT TPD_TIP_COD_IMPOSTO,TPD_DESCRICAO,tpd_valor_ufm,tpd_item  From TAB_PARAMETRO_DETALHE where   TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(cboMovimento.Coluna(0).Valor, tctexto), 1
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

Private Sub cboSitCad_Click()
    If Me.Tag <> "" Then Exit Sub
    If cboSitCad.Coluna(1).Valor <> 1 Then
        Do
            DataBaixa = Util.Entrada("Informe a data do processo.", "", Date)
        Loop While DataBaixa <> ""
    Else
        DataBaixa = ""
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
            txtVeiculo.SetFocus
        Case 4
            fraAnu.Enabled = chkCad(Index).Value
            cmdAdAnuncio.Enabled = chkCad(Index).Value
            If chkCad(Index).Value = ssCBUnchecked Then
                cboMovimento.ListIndex = -1
                txtValor = "0,00"
                txtValorApagar = "0,00"
            End If
            txtValor.Enabled = False
            txtValorApagar.Enabled = False

    End Select
End Sub

Private Sub cmdAdAnuncio_Click()
    Dim RetIm                           As String
    Dim i                                   As Byte
    Dim SQL                               As String
    Dim Rs                                As VSRecordset
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
   grdAnuncio.ListItems.Item(Index).SubItems(8) = TiraPic(txtIm, "-") & Format(Index, "00")
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

Private Sub cmdAdAtividade_Click()
    Contribuinte.Adiciona_Atividade_Secundaria grdAtividade, cboGrupoAtividade, CboAtividade
End Sub

Private Sub cmdExcluir_Click()
    If grdAnuncio.ListItems.Count >= 1 Then
        If Confirma("Deseja excluir?") Then
            If Anuncio.Excluir(grdAnuncio.SelectedItem) Then
                Avisa "Anuncio excluido com sucesso."
                Anuncio.PreencherGrd grdAnuncio, txtIm
                TabAnuncio.Tabs(1).Selected = True
            End If
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    Screen.MousePointer = 11
    If Trim(txtIm) <> "" Then Imposto.ImprimeFC txtIm, Rpt
    Screen.MousePointer = 0
End Sub

Private Sub CmdLimparAnuncio_Click()
    LimpaAnuncio
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSalvar_Click()
    On Error GoTo trata
    Dim Item As Object
    Dim Motivo As String
    If Not GraveiContrib Then
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
        If Not Edita.CriticaCampos(Me) Then Exit Sub
        If InscricaoMunicipal = "" Then
            Avisa "Contribuinte não encontrado."
            Exit Sub
        End If
            If cboAtivSecund <> "" Then
                If cboAtivServ.Coluna(1).Valor = cboAtivSecund.Coluna(1).Valor Then
                Util.Avisa "Atividade principal deve ser diferente da atividade segundária."
                cboAtivSecund.SetFocus
                Exit Sub
            End If
        End If
        Screen.MousePointer = 11
        '*****************************************Grava Historico do imovel************************************
        If 1 = 2 Then
                        CodMudanca = Conta.GeraCodPagamento(0)
                        '- TAB_CONTRIBUINTE_HISTORICO
                        Contribuinte.GravarHistorico CStr(CodMudanca), InscricaoMunicipal
                        '- TAB_SOCIO
                        Socio.GravarHistorico CStr(CodMudanca), InscricaoMunicipal
                        '- TAB_CONTADOR
                        Contador.GravarHistorico CStr(CodMudanca), InscricaoMunicipal
                        '- TAB_TRANSPORTADOR_VEICULO
                        Transportador.GravarHistorico CStr(CodMudanca), InscricaoMunicipal
        End If
            '*****************************************FIM DO HISTORICO*********************************************
                'Vou cadastrar o contribuinte
                Do
                    Motivo = Entrada("Informe o motivo da alteração.", "CIAP")
                Loop Until Motivo <> ""
        With Contribuinte
            .Im = InscricaoMunicipal
            .CgcCpf = txtCgc
            .ImAuxiliar = txtIncriAuxiliar
            .Nome = txtRazao
            .Data = Bdados.Converte(Date, TCDataHora)
            .Motivo = Motivo
            .SituacaoAlvara = CStr(cboSitAlvara.Coluna(1).Valor)
            .InicioPrestacaoServico = txtInicioPrestacaoServico
            .Fantasia = txtFantasia
            .Cep = txtCep
            .Matriz_Filial = CStr(cboMatrizFilial.Coluna(1).Valor)
            If txtDataReabertura <> "" Then
                .Data_Encerramento = txtDataEncerramento
            End If
            If txtDataReabertura <> "" Then
                .Data_Reabertura = txtDataReabertura
            End If
            .VariavelAnuncio = txtvAnuncio
            .VariavelDataFormatura = txtvDataFormatura
            .VariavelFuncionarioSUS = CbovFuncionarioSUS.Coluna(1).Valor
            .VariavelQuantidadeItem = txtvQtdItems
            .CodGrupo = CStr(cboClassAtiv.Coluna(1).Valor)
            .TipoCadastro = CStr(cboTipoCadastro.Coluna(1).Valor)
            .CodSitCadastral = cboSitCad.Coluna(1).Valor
            .CodNatureza = CStr(cboNatJur.Coluna(1).Valor)
            .CodAtivPoder = cboAtivPoder.Coluna(1).Valor
            .Estabelecido = CStr(cboEstabelece.Coluna(1).Valor)
            .DataCadastro = Bdados.Converte(Date, TCDataHora)
            .CodAtividade = CStr(Nvl(CStr(cboAtivServ.Coluna(1).Valor), 0))
            If Trim(cboAtivSecund) <> "" Then .CodAtividadeSec = CStr(cboAtivSecund.Coluna(1).Valor)
            .CodUsuario = AplicacoesVTFuncoes.Usuario
            .Logradouro = cboTipoLogr
            .NomeLogradouro = cboLogr
            .Numero = txtNum
            .Complemento = txtComplemento
            .Bairro = cboBairro
            .Cidade = txtCidade
            .Uf = cboUF
            If Trim(cboAtivSecund) <> "" Then .GrupoAtividade = CStr(cboAtivServ.Coluna(2).Valor)
            .InicioAtividade = Bdados.Converte(txtDtInicio, TCDataHora)
            .TipoContribuinte = CStr(cboNatJur.Coluna(1).Valor)
            .TipoRecolhimentoIss = CStr(cboObrigIss.Coluna(1).Valor)
            .Ruc = txtRuc
            If Trim(txtFator) = "" Then
                txtFator = "1"
            End If
            .FatorAlvara = txtFator
            .Conselho = txtConselho
            .Registro = IIf(Trim(txtRegistro) = "", 0, txtRegistro)
            If CStr(cboImovel.Coluna(1).Valor) <> "" Then
                .ImovelProprio = CStr(cboImovel.Coluna(1).Valor)
            End If
            .NumEmpregado = Nvl(txtEmpregados, 0)
            .PorteEmpresa = Nvl(CStr(cboPorte.Coluna(1).Valor), 0)
            .CodAtividadeSec = CStr(IIf(Trim(cboAtivSecund) = "", 0, cboAtivSecund.Coluna(1).Valor))
            .CodAtividadeTerc = CStr(IIf(Trim(cboAtivSecund2) = "", 0, cboAtivSecund2.Coluna(1).Valor))
            .NivelEscolar = cboNivel.Coluna(1).Valor
            .Protocolo = txtProtocolo
            If Trim(txtIc) <> "" Then .Ic = txtIc
            .Isento = Nvl(CStr(cboIsento.Coluna(1).Valor), 0)
            .CodRamo = Nvl(CStr(cboRamo.Coluna(1).Valor), 0)
            .FoneFax = CDbl(IIf(txtFoneFax = "", 0, txtFoneFax))
            .CNH = txtCnh
            .Categoria = txtCateg
            .Autorizacao = Nvl(txtAutoriza, 0)
            .PontoRecepcao = cboPonto
            'Salva dados do contribuinte
            .Salvar
        End With
        'GRAVANDO SOCIOS
        Socio.Excluir InscricaoMunicipal ' exclui todos registros d socios
        If chkCad(0).Value Then
            If grdSocio.ListItems.Count > 0 Then
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
        Contador.Excluir InscricaoMunicipal
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
        Representante.Excluir InscricaoMunicipal
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
                    .Uf = cboUF
                    If Trim(txtImRepresentante) <> "" Then .ImRepresentante = Bdados.Converte(txtImRepresentante, tctexto)
                    .Salvar
                End With
            End If
        End If
                'GRAVANDO TRANSPORTADOR
        Transportador.Excluir InscricaoMunicipal ' exclui todos registros d veiculos
        If chkCad(3).Value Then
             If grdVeiculo.ListItems.Count <> "0" Then 'Se definiu algum carro
                'Contribuinte.AtividadeTransporte = cboAtividadeVeiculo.Coluna(1).VALOR
                For Each Item In grdVeiculo.ListItems
                 '   grdVeiculo.ListItems(Item).Selected = True
                    With Transportador
                        .Im = InscricaoMunicipal
                        .Veiculo = Item
                        .Marca = Item.SubItems(1)
                        .CodModelo = Item.SubItems(2)
                        .AnoFabricacao = Item.SubItems(3)
                        .Placa = Item.SubItems(4)
                        .Chassi = Item.SubItems(5)
                        .Municipio = Item.SubItems(6)
                        .Uf = Item.SubItems(7)
                        .Licensa = Item.SubItems(8)
                        'Pego a posição do  Traço...
                        Dim Pos As Integer
                        Pos = InStr(Item.SubItems(9), "-")
                        .atividade = Left(Item.SubItems(9), Pos - 1)
                        If IsDate(Item.SubItems(10)) Then .IniAtividadeCarro = Item.SubItems(10)
                        .Salvar
                    End With
                Next
            End If
        End If
        Contribuinte.Grava_Atividade_Secundaria grdAtividade, InscricaoMunicipal
        Call Util.Informa("Registro gravado com sucesso. Inscricão Municipal Nº: " & InscricaoMunicipal & ".")
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
    grdSocio.ColumnHeaders.Add , , "Logr."
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
    Dim RetIm As String
    
    Dim i As Byte
    Dim SQL As String
    Dim Index As Integer
    
    Dim Rs As VSRecordset
    If Trim(txtPlaca) = "" Then Exit Sub
    'If Transportador.VerificaChassi(txtChassi, RetIm) Then
   '     Util.Informa "Placa cadastrada para contribuinte IM = '" & RetIm & "'."
  '      Exit Sub
 '   End If
    
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
    
    
    'Set ItmX = grdVeiculo.ListItems.Add(, , txtVeiculo)
    'With ItmX
    '    .SubItems(1) = txtMarca
    '    .SubItems(2) = txtModelo
    '    .SubItems(3) = txtAnoFabric
    '    .SubItems(4) = txtPlaca
    '    .SubItems(5) = txtChassi
    '    .SubItems(6) = txtMunicipio
    '    .SubItems(7) = cboUFTransp
    '    .SubItems(8) = txtLicenca
    'End With
    
    txtVeiculo = ""
    txtMarca = ""
    txtModelo = ""
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

Private Sub cmdLimpar_Click()
        Edita.LimpaCampos Me
        txtIm.Enabled = True
        txtCgc.Enabled = True
        InscricaoMunicipal = ""
        tabCadastro.Tabs(1).Selected = True
        Screen.MousePointer = 0
        grdSocio.ListItems.Clear
        grdVeiculo.ListItems.Clear
        grdAnuncio.ListItems.Clear
        Anuncio.MontarGrid grdAnuncio
        chkCad_Click 0, 0
        chkCad_Click 1, 0
        chkCad_Click 2, 0
        chkCad_Click 3, 0
        Screen.MousePointer = 0
        txtProtocolo.SetFocus
End Sub

Private Sub cmdPesq_Click(Index As Integer)
    Select Case Index
        Case 1
            cmdLimpar_Click
            AplicacoesVTFuncoes.BuscaNoEconomico TcoJuridica, txtIm
    End Select
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvarAnuncio_Click()
    Dim Anuncio As New eAnuncio
     
    If icad = "" Then
        icad = Conta.GeraCodPagamento("ICAD")
    End If
    With Anuncio
        .Im = txtIm
        .icad = icad
        .Movimento = CStr(cboMovimento.Coluna(0).Valor)
        .Dimensao = txtDimensao
        .Area = txtArea
        .DataInstalacao = txtDataInstalacao
        .Valor_UFM = txtValor
        .Valor_Apagar = txtValorApagar
        .Multiplicador = txtFatorMutiplicador
        .SubPublicidade = cboItem.Text
        .Doc_Origem = txtIm & icad
        If txtInscImob.Text <> "" Then
            .ICLocal = txtInscImob
        Else
            .TipoLogra = CStr(cboTipoLOgraAnuncio.Text)
            .Logradouro = CStr(cboLogradouroAnuncio.Text)
            .Bairro = CStr(cboBairroAnuncio.Text)
        End If
        .Obs = txtObs
        If .Salvar = True Then
            .PreencherGrd grdAnuncio, txtIm
            TabAnuncio.Tabs(1).Selected = True
        End If
        icad = ""
    End With
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao 1, txtInscImob
End Sub

Private Sub cmdVISUAL2_Click()
    Contribuinte.Apaga_Atividade_Secundaria grdAtividade
End Sub

Private Sub cmdVISUAL3_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtIc
End Sub

Private Sub Form_Activate()
'    Dim i As Byte
    If AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIm.Formato = formNenhum
    End If
    
    If AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        cboPonto.Visible = False
    End If
    txtCep.Formato = formNenhum
    cboNivel.Enabled = False
     cboSitAlvara.PreencherGeral Bdados, "SITUACAO ALVARA"
    cboTipoCadastro.PreencherGeral Bdados, "TIPO CADASTRO ECONOMICO"
cboMatrizFilial.PreencherGeral Bdados, "TIPO EMPRESA"
'    Set Contribuinte = New eContribuinte
'    Set Transportador = New eTransportador
'    Set Endereco = New eEndereco
'    Set Contador = New eContador
'    Set Atividade = New Atividade
'    Set Socio = New eSocio
'    Set Imovel = New eImovel
 '   Set Representante = New eRepresentante
'    Set Cadastro = New VSImposto
    
'    cabVisual.Exibir Bdados, Me.Name, App.Path
'    rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
    
    Screen.MousePointer = 0
    
    With Endereco
        .PreencherCboTipoLogr cboTipoLOgraAnuncio
        .PreencherCboBairro cboBairroAnuncio
        
    End With
'    Contador.PreencherCboContador cboContador
'    Contribuinte.PreencherCboSitCad cboSitCad
'    cboEstabelece.PreencherGeral Bdados, "ESTABELECIDO"
'    cboImovel.PreencherGeral Bdados, "TIPO IMÓVEL"
'    cboObrigIss.PreencherGeral Bdados, "TIPO RECOLHIMENTO ISS"
'    cboIsento.PreencherGeral Bdados, "SIM OU NÃO"
'    cboNivel.PreencherGeral Bdados, "NIVEL INSTRUÇÃO"
'    cboPorte.PreencherGeral Bdados, "PORTE EMPRESA"
    With atividade
'        .PreencheCombo cboClassAtiv, iaGrupoAtividade
'        .PreencheCombo cboRamo, iaRamo
'        .PreencherCboPoder cboAtivPoder
        .PreencherCboAtiv cboAtivServ
        .PreencherCboAtiv cboAtivSecund
        .PreencherCboAtiv cboAtivSecund2
'        .PreencherCboNturJur cboNatJur
    End With
'    cboUF.PreencherGeral Bdados, "UF"
'    cboUFSocio.PreencherGeral Bdados, "UF"
'    cboUfRepresentante.PreencherGeral Bdados, "UF"
'    cboUFTransp.PreencherGeral Bdados, "UF"
    
    GraveiContrib = False
    
    If Me.Tag <> "" Then
        txtIm = Me.Tag
        txtIM_LostFocus
        fraSocio1.Enabled = False
        fraSocio2.Enabled = False
        fraContador.Enabled = False
        fraRepresen1.Enabled = False
        fraRepresen2.Enabled = False
        fraTrans.Enabled = False
        cmdSalvar.Enabled = False
        cmdLimpar.Enabled = False
        FraDados.Enabled = False
    End If
    CbovFuncionarioSUS.PreencherGeral Bdados, "FUNCIONARIO SUS"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        tabCadastro.Tab = IIf(tabCadastro.Tab = 3, 0, tabCadastro.Tab)
        If tabCadastro.Tab <> 0 Then chkCad(tabCadastro.Tab - 1).SetFocus
    End If
End Sub

Private Sub Form_Load()
    Set Contribuinte = New eContribuinte
    Set Transportador = New eTransportador
    Set Endereco = New eEndereco
    Set Contador = New eContador
    Set atividade = New atividade
    Set Socio = New eSocio
    Set Imovel = New eImovel
    Set Representante = New eRepresentante
    Set Cadastro = New VSImposto
    Set Anuncio = New eAnuncio
    'tabCadastro.Tabs(6).Enabled = False
    Screen.MousePointer = 0
    txtValor.Enabled = False
    txtDataInstalacao.Formato = formData
  
  cabVISUAL1.Exibir Bdados, Me.Name, App.Path
  rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
    'Bdados.CriaViews "Vis_Contrib_Servico", Bdados.BDSistema, TCIS101.ConsultaISS
    With Endereco
        .PreencherCboTipoLogr cboTipoLogr
        .PreencherCboTipoLogr cboTipoLogrSocio
        .PreencherCboTipoLogr cboTipoLogrRepresentante
        .PreencherPonto cboPonto
        .PreencherCboBairro cboBairro
        .PreencherCboBairro cboBairroRepresentante
        .PreencherCboBairro cboBairroSocio
    End With
    cboLogr.Preencher Bdados, "select tlg_cod_logradouro,TLG_NOME from tab_logradouro", 1
    cboLogradouroAnuncio.Preencher Bdados, "select tlg_cod_logradouro,TLG_NOME from tab_logradouro", 1
    Contador.PreencherCboContador cboContador
    Contribuinte.PreencherCboSitCad cboSitCad
    cboEstabelece.PreencherGeral Bdados, "ESTABELECIDO"
    cboImovel.PreencherGeral Bdados, "TIPO IMÓVEL"
    cboObrigIss.PreencherGeral Bdados, "TIPO RECOLHIMENTO ISS"
    cboIsento.PreencherGeral Bdados, "SIM OU NÃO"
    cboNivel.PreencherGeral Bdados, "NIVEL INSTRUÇÃO"
    cboPorte.PreencherGeral Bdados, "PORTE EMPRESA"
    With atividade
        .PreencheCombo cboClassAtiv, iaGrupoAtividade
        .PreencheCombo cboGrupoAtividade, iaGrupoAtividade
        .PreencheCombo cboRamo, iaRamo
        .PreencherCboPoder cboAtivPoder
        .PreencherCboAtiv cboAtivServ
        .PreencherCboAtiv cboAtivSecund
        .PreencherCboAtiv cboAtivSecund2
        .PreencherCboNturJur cboNatJur
        .PreencherCboAtiv cboAtividadeVeiculo
    End With
    cboUF.PreencherGeral Bdados, "UF"
    cboUFSocio.PreencherGeral Bdados, "UF"
    cboUfRepresentante.PreencherGeral Bdados, "UF"
    cboUFTransp.PreencherGeral Bdados, "UF"
    
    GraveiContrib = False
    MontaCabGrid
    Anuncio.MontarGrid grdAnuncio
    cboMovimento.Preencher Bdados, "Select tip_cod_imposto , tip_nome_imposto as Tributo FROM Tab_Imposto where tip_sigla_Imposto = '" & Imposto.NomeTributo(ttr_PUBLICIDADE) & "'", 1
    
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
    Else
        txtIm.Formato = formDoisDigitos
    End If
    If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
        cboRamo.Visible = False
        cboPonto.Visible = False
    End If
End Sub

Private Sub PegaDadosAnuncios()
    If Anuncio.Buscar(grdAnuncio.SelectedItem) Then
        LimpaAnuncio
        With Anuncio
            icad = .icad
            cboMovimento.SetarLinha .Movimento
            cboMovimento_Click
            If .SubPublicidade <> "" Then
                cboItem = .SubPublicidade
            End If
                txtDimensao = .Dimensao
                txtArea = .Area
                Rem txtArea_LostFocus
                txtFatorMutiplicador = .Multiplicador
                txtDataInstalacao = .DataInstalacao
                txtValor = .Valor_UFM
                txtValorApagar = .Valor_Apagar
                If .ICLocal <> "" Then
                    txtInscImob = .ICLocal
                    Call txtInscImob_LostFocus
                Else
                    cboTipoLOgraAnuncio = .TipoLogra
                    cboLogradouroAnuncio = .Logradouro
                    cboBairroAnuncio = .Bairro
                End If
                txtObs = .Obs
            
            TabAnuncio.Tabs(2).Selected = True
        End With
    End If
End Sub
Private Sub LimpaAnuncio()
    On Error Resume Next
    cboMovimento.ListIndex = -1
    cboItem.ListIndex = -1
    txtDimensao = ""
    txtArea = ""
    txtFatorMutiplicador = ""
    txtDataInstalacao = ""
    txtValor = ""
    txtValorApagar = ""
    txtInscImob = ""
    cboTipoLOgraAnuncio.ListIndex = -1
    cboLogradouroAnuncio.ListIndex = -1
    cboBairroAnuncio.ListIndex = -1
    txtObs = ""
    icad = ""
    cboMovimento.SetFocus
End Sub

Private Sub grdAnuncio_DblClick()
    If grdAnuncio.ListItems.Count >= 1 Then
        PegaDadosAnuncios
    End If

End Sub

Private Sub grdAtividade_DblClick()
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
    txtChassi = grdVeiculo.SelectedItem.SubItems(10)
    grdVeiculo.ListItems.Remove (grdVeiculo.SelectedItem.Index)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Transportador = Nothing
    Set Contribuinte = Nothing
    Set Endereco = Nothing
    Set Contador = Nothing
    Set Imovel = Nothing
    Set Representante = Nothing
    Set Conta = Nothing
End Sub

Private Sub txtArea_LostFocus()
    Calcula
End Sub
Private Sub Calcula()
 Dim SQL As String
    Dim Rs As VSRecordset
    If cboMovimento.Text = "" Then Exit Sub
    If txtFatorMutiplicador = "" Then
        txtFatorMutiplicador = 0
    End If
    If txtArea <> "" Or txtValor <> "" Then
        'Pego os dados dos anucios..
        SQL = "Select * from Tab_Parametro_taxas where TPT_TIP_COD_IMPOSTO =  " & Bdados.Converte(cboMovimento.Coluna(0).Valor, tctexto)
        If Bdados.AbreTabela(SQL, Rs) Then
            'Checo se a taxa está tabelada ou não
            If Rs.Fields("TPT_TIPO") = 1 Then ' Tem Tabela
                'Faço o laço para ferificar onde o cara se enquadra.
                Do Until Rs.EOF
                    If Val(Nvl(txtArea, 0)) >= Val(Rs.Fields("TPT_LIMITE_INFERIOR")) And Val(txtArea) <= Val(Rs.Fields("TPT_LIMITE_SUPERIOR")) Then
                        If txtFatorMutiplicador <> "" Then
                            txtValor = Rs.Fields("TPT_VALOR_UFM") * CDbl(Nvl(txtFatorMutiplicador, 0))
                            txtValorApagar = Calcula_UFM(Nvl(txtValor, 0), Converete_Real)
                        Else
                            txtValor = Rs.Fields("TPT_VALOR_UFM")
                            txtValorApagar = Calcula_UFM(Nvl(txtValor, 0), Converete_Real)
                        End If
                    End If
                    Rs.MoveNext
                Loop
            Else
                'Não é tabelado então eu pego o valor que foi estimado...
                If txtFatorMutiplicador <> "" Then
                    txtValor = Rs.Fields("TPT_VALOR_UFM") * CDbl(Nvl(txtFatorMutiplicador, 0))
                    txtValorApagar = Calcula_UFM(Nvl(txtValor, 0), Converete_Real)
                Else
                    txtValor = Rs.Fields("TPT_VALOR_UFM")
                    txtValorApagar = Calcula_UFM(Nvl(txtValor, 0), Converete_Real)
                End If
            End If
         Else
            SQL = "Select * from Tab_Parametro_Detalhe where TPD_TIP_COD_IMPOSTO =  " & Bdados.Converte(cboMovimento.Coluna(0).Valor, tctexto) & " and tpd_item = " & Bdados.Converte(cboItem.Coluna(3).Valor, tctexto)
            If Bdados.AbreTabela(SQL, Rs) Then
                If Not IsNumeric(txtFatorMutiplicador) Then Exit Sub
                If Trim(txtFatorMutiplicador) <> "" Then
                    txtValor = Rs.Fields("TPD_VALOR_UFM") * CDbl(Nvl(txtFatorMutiplicador, 0)) * CDbl(Nvl(txtArea, 0))
                    txtValorApagar = Calcula_UFM(Nvl(txtValor, 0), Converete_Real)
                Else
                    txtValor = Rs.Fields("TPD_VALOR_UFM")
                    txtValorApagar = Calcula_UFM(Nvl(txtValor, 0), Converete_Real)
                End If
            End If
        End If
    End If
    txtFatorMutiplicador.Enabled = True
End Sub

Private Sub txtcgc_LostFocus()
'    If Trim(txtCgc) = "" Then Exit Sub
'     If txtCgc = "99999999999" Or txtCgc = "999.999.999-99" Or txtCgc = "00000000000" Or txtCgc = "000.000.000-00" Then
'        Util.Avisa "Valor do CPF inválido."
'        txtCgc.SetFocus
'    End If
'    txtCgc.Formato = formNenhum
'    If Len(txtCgc) = 11 Then
'        txtCgc.Formato = formCPF
'    ElseIf Len(txtCgc) = 14 And Not IsNumeric(txtCgc) Then
'         txtCgc.Formato = formCPF
'    ElseIf Len(txtCgc) = 14 And IsNumeric(txtCgc) Then
'         txtCgc.Formato = formCGC
'    ElseIf Len(txtCgc) = 18 And Not IsNumeric(txtCgc) Then
'         txtCgc.Formato = formCGC
'    Else
'        Util.Informa "Cpf ou Cnpj inválido."
'        txtCgc.SetFocus
'        txtCgc.Formato = formNenhum
'        Exit Sub
'    End If
'    If txtCgc = "" Or txtIm <> "" Then txtCgc.Formato = formNenhum: Exit Sub
'    Call PreencheTela(txtIm, txtCgc)
'    txtCgc.Formato = formNenhum
End Sub

Private Sub txtFator_Change()
'    Calcula
End Sub

Private Sub txtFatorMutiplicador_LostFocus()
Calcula
End Sub

Private Sub txtIC_LostFocus()
    If Trim(txtIc) = "0" Then
        txtIc = ""
    End If
    If Trim(txtIc) = "" Then Exit Sub
    If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) <> 1 Then
        txtIc = Cadastro.FormataInscricao(txtIc, InscImovel)
    End If
    If Imovel.BuscarImovel(txtIc, cboTipoLogr, cboLogr, txtNum, txtComplemento, cboBairro, txtCep, txtCidade, cboUF) = False Then
        Util.Informa ("Imóvel não cadastrado.")
        cboTipoLogr.ListIndex = -1
        txtNum = ""
    End If
End Sub


Private Sub txtIM_LostFocus()
    If Trim(txtIm) = "" Then Exit Sub
    If txtIm.Enabled = False Then Exit Sub
    Call PreencheTela(txtIm, txtCgc)
End Sub

Private Sub txtImRepresentante_LostFocus()
    If Trim(txtImRepresentante) = "" Then Exit Sub
    With Contribuinte
        If .Buscar(txtImRepresentante, , False) Then
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
Private Sub txtCpfSocio_LostFocus()
    Dim CpfBusca As String
    If Trim(txtCpfSocio) = "" Then Exit Sub
    CpfBusca = txtCpfSocio
    If Socio.Buscar(, CpfBusca) Then
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
        If Contribuinte.Buscar(, CpfBusca, False) Then
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

Private Sub txtInscImob_LostFocus()
    Dim SQL As String
    Dim Rs As VSRecordset
    
    If Trim(txtInscImob) = "" Then Exit Sub
    If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) <> 1 Then
        txtInscImob = Cadastro.FormataInscricao(txtInscImob, InscImovel)
        
    End If
    SQL = "Select * from Vis_Imovel Where tim_Ic = '" & txtInscImob & "' And TBA_TMU_COD_MUNICIPIO = " & Aplicacoes.Codigo_Municipio & " And tlg_tmu_cod_municipio = " & Aplicacoes.Codigo_Municipio
    If Bdados.AbreTabela(SQL, Rs) Then
        txtInscImob = Rs!tim_ic
        cboBairroAnuncio = Rs!tba_nome
        cboLogradouroAnuncio = Rs!tlg_nome
        cboTipoLOgraAnuncio = Rs!ttl_nome
    Else
        Util.Informa ("Imóvel não cadastrado.")
        cboTipoLOgraAnuncio.ListIndex = -1
    End If

End Sub

