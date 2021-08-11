VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TDEC110 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DECL101"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TDEC110.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs TabDec 
      Height          =   4815
      Left            =   30
      TabIndex        =   25
      Top             =   3120
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   8493
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      Tabs            =   "TDEC110.frx":0342
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4425
         Index           =   0
         Left            =   30
         TabIndex        =   26
         Top             =   30
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   7805
         _Version        =   131082
         TabGuid         =   "TDEC110.frx":03EA
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   4605
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   8123
            Altura          =   1905
            Caption         =   " Apurac�o das Sa�das - Apura��o de imposto sobre as notas emitidas"
            CorFaixa        =   16711680
            Ocultavel       =   0   'False
            Begin VTOcx.fraVISUAL fraNormal 
               Height          =   2175
               Index           =   3
               Left            =   0
               TabIndex        =   43
               Top             =   300
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   3836
               Altura          =   1905
               Caption         =   " Nota Fiscal Emitida"
               CorTexto        =   65535
               CorFaixa        =   16711680
               Ocultavel       =   0   'False
               Begin VTOcx.txtVISUAL txtTomadorRazao 
                  Height          =   555
                  Left            =   3840
                  TabIndex        =   47
                  Top             =   960
                  Width           =   7140
                  _ExtentX        =   12594
                  _ExtentY        =   979
                  Caption         =   "TOMADOR"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
                  CorFundo        =   14737632
               End
               Begin VTOcx.cmdVISUAL cmdBuscaTomador 
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   1155
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   556
                  Caption         =   ""
                  Acao            =   5
                  CorBorda        =   16711680
                  CorFrente       =   16384
                  CorFundo        =   16777088
               End
               Begin VTOcx.txtVISUAL txtTomadorCpfCnpj 
                  Height          =   525
                  Left            =   1560
                  TabIndex        =   7
                  Top             =   960
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   926
                  Caption         =   "CPF/CNPJ Tomador"
                  Text            =   ""
                  Restricao       =   2
                  AlinhamentoRotulo=   1
                  CorFundo        =   14737632
                  AgruparValores  =   0   'False
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtTMNumNota 
                  Height          =   525
                  Left            =   1560
                  TabIndex        =   4
                  Top             =   360
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   926
                  Caption         =   "N� Nota Fiscal"
                  Text            =   ""
                  Restricao       =   2
                  AlinhamentoRotulo=   1
                  AlinhamentoTexto=   1
                  CorFundo        =   14737632
               End
               Begin VTOcx.txtVISUAL txtTMInscricao 
                  Height          =   525
                  Left            =   5520
                  TabIndex        =   6
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   926
                  Caption         =   "CPF/CNPJ Prestador"
                  Text            =   ""
                  Restricao       =   2
                  AlinhamentoRotulo=   1
                  CorFundo        =   14737632
                  AgruparValores  =   0   'False
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtAidf 
                  Height          =   525
                  Left            =   4200
                  TabIndex        =   44
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   926
                  Caption         =   "No. AIDF"
                  Text            =   ""
                  AlinhamentoRotulo=   1
                  AlinhamentoTexto=   1
                  CorFundo        =   14737632
               End
               Begin VB.CheckBox chkCancelada 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Nota Cancelada"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  TabIndex        =   45
                  Top             =   480
                  Width           =   1365
               End
               Begin VTOcx.txtVISUAL txtTMSaldoDevedor 
                  Height          =   525
                  Left            =   8160
                  TabIndex        =   15
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   926
                  Caption         =   "Saldo Devedor"
                  Text            =   ""
                  Enabled         =   0   'False
                  Formato         =   5
                  Restricao       =   3
                  AlinhamentoRotulo=   1
                  AlinhamentoTexto=   1
                  CorFundo        =   14737632
                  ValorPadrao     =   "0"
               End
               Begin VTOcx.txtVISUAL txtTMImpostoRetido 
                  Height          =   525
                  Left            =   6840
                  TabIndex        =   14
                  Top             =   1560
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   926
                  Caption         =   "ISS Retido"
                  Text            =   ""
                  Formato         =   5
                  Restricao       =   3
                  AlinhamentoRotulo=   1
                  AlinhamentoTexto=   1
                  CorFundo        =   14737632
                  ValorPadrao     =   "0"
               End
               Begin VTOcx.txtVISUAL txtTMAliq 
                  Height          =   525
                  Left            =   2880
                  TabIndex        =   11
                  Top             =   1560
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   926
                  Caption         =   "Aliq(%)"
                  Text            =   ""
                  Formato         =   5
                  Restricao       =   3
                  AlinhamentoRotulo=   1
                  AlinhamentoTexto=   1
                  CorFundo        =   14737632
                  ValorPadrao     =   "0"
               End
               Begin VTOcx.txtVISUAL txtTMIcms 
                  Height          =   525
                  Left            =   1560
                  TabIndex        =   10
                  Top             =   1560
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   926
                  Caption         =   "Vlr ICMS"
                  Text            =   ""
                  Formato         =   5
                  Restricao       =   3
                  AlinhamentoRotulo=   1
                  AlinhamentoTexto=   1
                  CorFundo        =   14737632
                  ValorPadrao     =   "0"
               End
               Begin VTOcx.cmdVISUAL cmdAdicionarNotaTM 
                  Height          =   375
                  Left            =   9480
                  TabIndex        =   16
                  ToolTipText     =   "Adicionar"
                  Top             =   1680
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   661
                  Caption         =   "&Incluir"
                  Acao            =   1
                  CorBorda        =   16711680
                  CorFrente       =   0
                  CorFundo        =   16777088
               End
               Begin VTOcx.txtVISUAL txtTMSaldo 
                  Height          =   525
                  Left            =   4200
                  TabIndex        =   12
                  Top             =   1560
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   926
                  Caption         =   "Base Calculo"
                  Text            =   ""
                  Enabled         =   0   'False
                  Formato         =   5
                  Restricao       =   3
                  AlinhamentoRotulo=   1
                  AlinhamentoTexto=   1
                  CorFundo        =   14737632
                  ValorPadrao     =   "0"
               End
               Begin VTOcx.txtVISUAL txtTMImpostoDevido 
                  Height          =   525
                  Left            =   5520
                  TabIndex        =   13
                  Top             =   1560
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   926
                  Caption         =   "ISS Devido"
                  Text            =   ""
                  Enabled         =   0   'False
                  Formato         =   5
                  Restricao       =   3
                  AlinhamentoRotulo=   1
                  AlinhamentoTexto=   1
                  CorFundo        =   14737632
                  ValorPadrao     =   "0"
               End
               Begin VTOcx.txtVISUAL txtTMValor 
                  Height          =   525
                  Left            =   120
                  TabIndex        =   9
                  Top             =   1560
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   926
                  Caption         =   "Vlr da Nota"
                  Text            =   ""
                  Formato         =   5
                  Restricao       =   3
                  AlinhamentoRotulo=   1
                  AlinhamentoTexto=   1
                  CorFundo        =   14737632
                  ValorPadrao     =   "0"
               End
               Begin VTOcx.txtVISUAL txtTMEmissao 
                  Height          =   525
                  Left            =   2880
                  TabIndex        =   5
                  Top             =   360
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   926
                  Caption         =   "Data Emiss�o"
                  Text            =   ""
                  Formato         =   0
                  Restricao       =   2
                  AlinhamentoRotulo=   1
                  AlinhamentoTexto=   1
                  CorFundo        =   14737632
               End
            End
            Begin VTOcx.grdVISUAL grdSaida 
               Height          =   1980
               Left            =   0
               TabIndex        =   31
               Top             =   2520
               Width           =   11100
               _ExtentX        =   19579
               _ExtentY        =   3493
               CorBorda        =   16711680
               Caption         =   "Notas emitidas"
               CorTitulo       =   16711680
               CorCaption      =   16777215
               CorDica         =   16711680
               OcultarRodape   =   -1  'True
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4425
         Index           =   1
         Left            =   30
         TabIndex        =   27
         Top             =   30
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   7805
         _Version        =   131082
         TabGuid         =   "TDEC110.frx":0412
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   4635
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   8176
            Altura          =   1905
            Caption         =   " Apurac�o das Entradas - Apura��o de imposto sobre as notas recebidas"
            CorFaixa        =   5346129
            Ocultavel       =   0   'False
            Begin VTOcx.fraFUTURO fraFUTURO1 
               Height          =   5325
               Index           =   0
               Left            =   0
               TabIndex        =   33
               Top             =   0
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   9393
               Caption         =   "Apura��o do Imposto"
               Descricao       =   "Apura��o de imposto sobre as notas emitidas"
               corFaixa        =   16711680
               corFundo        =   14737632
               corTexto        =   16384
               Icone           =   "TDEC110.frx":043A
               Ocultavel       =   0   'False
               Altura          =   2000
               Begin VTOcx.fraVISUAL fraNormal 
                  Height          =   735
                  Index           =   2
                  Left            =   60
                  TabIndex        =   42
                  Top             =   2370
                  Width           =   11175
                  _ExtentX        =   19711
                  _ExtentY        =   1296
                  Altura          =   1905
                  Caption         =   " Resumo de Recolhimento"
                  CorTexto        =   0
                  CorFaixa        =   16711680
                  Ocultavel       =   0   'False
                  Begin VTOcx.txtVISUAL txtItemDecl 
                     Height          =   315
                     Index           =   14
                     Left            =   360
                     TabIndex        =   40
                     Top             =   330
                     Width           =   3645
                     _ExtentX        =   6429
                     _ExtentY        =   556
                     Caption         =   "Total Retido por Tomadores"
                     Text            =   ""
                     Enabled         =   0   'False
                     Formato         =   5
                     Restricao       =   3
                     AlinhamentoTexto=   1
                     CorFundo        =   14737632
                  End
                  Begin VTOcx.txtVISUAL txtItemDecl 
                     Height          =   315
                     Index           =   13
                     Left            =   8100
                     TabIndex        =   41
                     Top             =   330
                     Width           =   2655
                     _ExtentX        =   4683
                     _ExtentY        =   556
                     Caption         =   "Total a Recolher"
                     Text            =   ""
                     Enabled         =   0   'False
                     Formato         =   5
                     Restricao       =   3
                     AlinhamentoTexto=   1
                     CorFundo        =   14737632
                  End
               End
               Begin VTOcx.fraVISUAL fraNormal 
                  Height          =   1635
                  Index           =   5
                  Left            =   60
                  TabIndex        =   35
                  Top             =   720
                  Width           =   11175
                  _ExtentX        =   19711
                  _ExtentY        =   2884
                  Altura          =   1905
                  Caption         =   " Resumo das Notas de Sa�da"
                  CorTexto        =   0
                  CorFaixa        =   16711680
                  Ocultavel       =   0   'False
                  Begin VTOcx.txtVISUAL txtItemDecl 
                     Height          =   315
                     Index           =   6
                     Left            =   7290
                     TabIndex        =   39
                     Top             =   1110
                     Width           =   3525
                     _ExtentX        =   6218
                     _ExtentY        =   556
                     Caption         =   "Imposto devido em notas"
                     Text            =   ""
                     Enabled         =   0   'False
                     Formato         =   5
                     Restricao       =   3
                     AlinhamentoTexto=   1
                     CorFundo        =   14737632
                  End
                  Begin VTOcx.txtVISUAL txtItemDecl 
                     Height          =   315
                     Index           =   5
                     Left            =   1350
                     TabIndex        =   38
                     Top             =   1110
                     Width           =   2655
                     _ExtentX        =   4683
                     _ExtentY        =   556
                     Caption         =   "Base de Calculo"
                     Text            =   ""
                     Enabled         =   0   'False
                     Formato         =   5
                     Restricao       =   3
                     AlinhamentoTexto=   1
                     CorFundo        =   14737632
                  End
                  Begin VTOcx.txtVISUAL txtItemDecl 
                     Height          =   315
                     Index           =   3
                     Left            =   990
                     TabIndex        =   36
                     Top             =   720
                     Width           =   3015
                     _ExtentX        =   5318
                     _ExtentY        =   556
                     Caption         =   "Valor total em notas"
                     Text            =   ""
                     Formato         =   5
                     Restricao       =   3
                     AlinhamentoTexto=   1
                     CorFundo        =   14737632
                  End
                  Begin VTOcx.txtVISUAL txtItemDecl 
                     Height          =   315
                     Index           =   4
                     Left            =   7740
                     TabIndex        =   37
                     Top             =   720
                     Width           =   3075
                     _ExtentX        =   5424
                     _ExtentY        =   556
                     Caption         =   "Total sujeito a ICMS"
                     Text            =   ""
                     Formato         =   5
                     Restricao       =   3
                     AlinhamentoTexto=   1
                     CorFundo        =   14737632
                  End
                  Begin VTOcx.txtVISUAL txtItemDecl 
                     Height          =   315
                     Index           =   2
                     Left            =   8640
                     TabIndex        =   18
                     Top             =   330
                     Width           =   2175
                     _ExtentX        =   3836
                     _ExtentY        =   556
                     Caption         =   "Nota Final"
                     Text            =   ""
                     Restricao       =   2
                     AlinhamentoTexto=   1
                     CorFundo        =   14737632
                  End
                  Begin VTOcx.txtVISUAL txtItemDecl 
                     Height          =   315
                     Index           =   1
                     Left            =   1770
                     TabIndex        =   17
                     Top             =   330
                     Width           =   2235
                     _ExtentX        =   3942
                     _ExtentY        =   556
                     Caption         =   "Nota Inicial"
                     Text            =   ""
                     Restricao       =   2
                     AlinhamentoTexto=   1
                     CorFundo        =   14737632
                  End
               End
               Begin VTOcx.grdVISUAL Grdtaxas 
                  Height          =   435
                  Left            =   60
                  TabIndex        =   34
                  Top             =   5565
                  Width           =   11190
                  _ExtentX        =   19738
                  _ExtentY        =   767
                  Caption         =   "Taxas"
                  OcultarRodape   =   -1  'True
                  CheckBox        =   -1  'True
                  Ordenavel       =   0   'False
               End
            End
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   24
      Top             =   8235
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdFinaliza 
         Height          =   375
         Left            =   6720
         TabIndex        =   20
         Top             =   90
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   661
         Caption         =   "Finalizar Declarac�o"
         Acao            =   1
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmLimpar 
         Height          =   375
         Left            =   8820
         TabIndex        =   21
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4770
         TabIndex        =   19
         Top             =   90
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Salvar Declarac�o"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10110
         TabIndex        =   22
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1080
      Left            =   30
      TabIndex        =   28
      Top             =   660
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   1905
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   16384
         CorFundo        =   16777088
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   285
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   503
         Caption         =   "Per�odo (mm/aaaa)"
         Text            =   ""
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   3840
         TabIndex        =   23
         Top             =   720
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   556
         Caption         =   "Insc.Municipal"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   510
         Left            =   8760
         TabIndex        =   3
         Top             =   530
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   900
         Caption         =   "Declarac�o"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
   End
   Begin VTOcx.grdVISUAL grdDec 
      Height          =   1485
      Left            =   30
      TabIndex        =   29
      Top             =   1800
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   2619
      CorBorda        =   16711680
      Caption         =   "Declarac�es"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   16711680
      OcultarRodape   =   -1  'True
   End
   Begin Cabecalho.cabVISUAL cabVisual1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   1138
      Icone           =   "TDEC110.frx":0D14
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "mnuGeral"
      Visible         =   0   'False
      Begin VB.Menu mnuDeletar 
         Caption         =   "Deletar"
      End
   End
End
Attribute VB_Name = "TDEC110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declaracao As VsTFuncoes.cDeclaracao
Private GerarIM As Boolean
Private AliqISSQN As Double, ISSQNFixo As Double
Private AliqISSST As Double, ISSSTFixo As Double

Private TotalImpostoST As Double
Private TotalBaseST As Double
Private TotalImpostoDevidoSaida As Double
Private TotalImpostoRetidoSaida As Double
Private TotalBaseSaida As Double
Private TotalICMSSujeito As Double
Private DeduzValores As Boolean
Private ContribuinteEndereco As String
Private ContribuinteAtividade As String
Dim Notas() As New NotaFiscal
Dim Modalidade As Integer
Dim ClassGrid As New grdEditavel
Dim String_Taxas As String
Dim Total_Taxas As Double

Private Sub AtualizaTextoSaidas(Indice As Integer, DeduzValoresTotais As Boolean)
    If grdSaida.ListItems.Count = 0 Then Exit Sub
    txtTMInscricao = grdSaida.ListItems(Indice).Text
    txtTMNumNota = grdSaida.ListItems(Indice).SubItems(1)
    txtTMEmissao = grdSaida.ListItems(Indice).SubItems(2)
    txtTMValor = grdSaida.ListItems(Indice).SubItems(3)
    txtTMIcms = grdSaida.ListItems(Indice).SubItems(4)
    txtTMSaldo = grdSaida.ListItems(Indice).SubItems(5)
    txtTMImpostoDevido = grdSaida.ListItems(Indice).SubItems(6)
    txtTMImpostoRetido = grdSaida.ListItems(Indice).SubItems(7)
    txtTMSaldoDevedor = grdSaida.ListItems(Indice).SubItems(8)
    chkCancelada.Value = grdSaida.ListItems(Indice).SubItems(9)
    txtTMAliq = grdSaida.ListItems(Indice).SubItems(10)
    txtAidf = grdSaida.ListItems(Indice).SubItems(11)
    txtTomadorCpfCnpj = grdSaida.ListItems(Indice).SubItems(11)
    
    If DeduzValoresTotais Then
        If Trim(txtTMImpostoRetido) <> Trim(txtTMImpostoDevido) Then
            TotalBaseSaida = TotalBaseSaida  '- CDbl(Nvl(txtTMSaldo, 0))
        End If
        TotalImpostoRetidoSaida = TotalImpostoRetidoSaida - CDbl(Nvl(txtTMImpostoRetido, 0))
        TotalImpostoDevidoSaida = TotalImpostoDevidoSaida - CDbl(Nvl(txtTMImpostoDevido, 0))
        grdSaida.ListItems.Remove Indice
    End If
    AtualizaApuracao
End Sub

Private Sub PreencheDeclaracao()
    On Error Resume Next
    Dim NumDec As String
    Dim i As Integer
    If UCase(Me.ActiveControl.Name) = "CMLIMPAR" Or UCase(Me.ActiveControl.Name) = "CMDSAIR" Then Exit Sub
    IniciaTotalizadores
    grdSaida.ListItems.Clear
    
    DeduzValores = False
    If Declaracao.Buscar(txtIM, txtPeriodo, CInt(cboTipo.Coluna(1).Valor), Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN))) Then
        cboTipo.SetarLinha Declaracao.Tipo, 1
        Declaracao.PreencheCamposDeclaracao grdSaida, grdSaida, txtItemDecl(1), txtItemDecl(2), txtItemDecl(3), txtItemDecl(4)
        TotalBaseSaida = txtItemDecl(3)
        TotalICMSSujeito = txtItemDecl(4)
        For i = 1 To grdSaida.ListItems.Count
            AtualizaTextoSaidas i, DeduzValores
            cmdAdicionarNotaTM_Click
        Next
        DoEvents
        AtualizaApuracao
        txtItemDecl_LostFocus 3
        TabDec.Tabs(3).Selected = True
        txtItemDecl(1).SetFocus
    
    End If
    DeduzValores = True
End Sub

Private Sub CarregaItensGrid(Lista As Object, CFOP As eTipoNotaOperacao)
    Dim i As Integer
    Dim j As Byte
    Dim Nota As NotaFiscal
    
    On Error GoTo TRATA
    
    For i = 1 To Lista.ListItems.Count
        Set Nota = New NotaFiscal
        If Trim(Lista.ListItems(i).ListSubItems(1).Text) <> "" And Trim(Lista.ListItems(i).ListSubItems(1).Text) <> "0,00" Then
            Nota.BaseCalculo = Lista.ListItems(i).ListSubItems(5).Text
            Nota.Data = Lista.ListItems(i).SubItems(2)
            Nota.ImpostoDevido = Lista.ListItems(i).SubItems(6)
            Nota.ImpostoRetido = Lista.ListItems(i).SubItems(7)
            Nota.Numero = Lista.ListItems(i).SubItems(1)
            Nota.Status = IIf(CFOP = etoSaida, Lista.ListItems(i).SubItems(9), 0)
            Nota.TipoOperacao = CFOP
            Nota.ValorMaterialICMS = Lista.ListItems(i).SubItems(4)
            Nota.ValorTotal = Lista.ListItems(i).SubItems(3)
            Nota.Destinatario = Lista.ListItems(i)
            Nota.Aliquota = Nvl(Lista.ListItems(i).SubItems(10), 0)
            Nota.AIDF = Nvl(Lista.ListItems(i).SubItems(11), 0)
            Nota.Tomador = Nvl(Lista.ListItems(i).SubItems(12), 0)
            
        End If
        Declaracao.Notas.Adicionar Nota
        DoEvents
    Next
    Exit Sub
TRATA:
    If Err.Number = 35600 Then
        Exit Sub
    End If
End Sub

Private Sub FormataRegistro(ByRef Inscricao As Object)
    Select Case Len(Inscricao.Text)
        Case 10
            Inscricao.Text = Imposto.FormataInscricao(Inscricao.Text, InscContrib)
        Case 11
            Inscricao.Text = Edita.FormataTexto(Inscricao, Cpf)
        Case 14
            Inscricao.Text = Edita.FormataTexto(Inscricao, Cgc)
    End Select
End Sub

Private Sub AtualizaApuracao()
    On Error Resume Next
    'RESUMO NOTAS SAIDA
    txtItemDecl(3) = TotalBaseSaida + TotalICMSSujeito
    txtItemDecl(4) = TotalICMSSujeito
    txtItemDecl(5) = TotalBaseSaida ' - TotalICMSSujeito
    txtItemDecl(6) = TotalImpostoDevidoSaida
    'NOTAS DE ENTRADA
    txtItemDecl(10) = TotalImpostoST
    txtItemDecl(11) = 0
    txtItemDecl(12) = TotalImpostoST
    'NOTAS EMITIDAS(SAIDAS)
    txtItemDecl(7) = TotalImpostoRetidoSaida
    txtItemDecl(8) = 0
    txtItemDecl(9) = TotalImpostoRetidoSaida
    'TOTAL RECOLHIMENTO
    txtItemDecl(13) = TotalImpostoDevidoSaida + TotalImpostoST - TotalImpostoRetidoSaida
'    txtItemDecl(100) = (100 * CDbl(txtItemDecl(13))) / (TotalBaseSaida + TotalBaseST)
End Sub

Private Sub IniciaTotalizadores()
    TotalImpostoST = 0
    TotalBaseSaida = 0
    TotalBaseST = 0
    TotalImpostoDevidoSaida = 0
    TotalImpostoRetidoSaida = 0
    TotalICMSSujeito = 0
End Sub

Private Sub cboTipo_Click()
    Declaracao.CarregaGrid grdDec, txtIM, txtPeriodo, CInt(cboTipo.Coluna(1).Valor), , Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN))
End Sub

Private Sub cboTipo_LostFocus()
  If cboTipo.Coluna(1).Valor = 2 Then
        Avisa "Utilize o formul�rio de ENTREGA DE DECLARAC�O NEGATIVA."
        cboTipo.ListIndex = -1
        cboTipo.SetFocus
    Else
        PreencheDeclaracao
    End If
End Sub

Private Sub chkCancelada_Click()
    txtTMIcms.Enabled = Not CBool(chkCancelada.Value)
    txtTMImpostoRetido.Enabled = Not CBool(chkCancelada.Value)
    txtTMInscricao.Enabled = Not CBool(chkCancelada.Value)
    txtTMValor.Enabled = Not CBool(chkCancelada.Value)
    
    txtTMIcms = 0
    txtTMImpostoDevido = 0
    txtTMImpostoRetido = 0
    txtTMInscricao = ""
    txtTMSaldo = 0
    txtTMValor = 0
    txtTMSaldoDevedor = 0
End Sub

Private Sub cmdAdicionarNotaTM_Click()
    
    Dim Linha As Object
    On Error Resume Next
    If txtIM = "" Or txtRazao = "" Then
       Util.Mensagem ("Informe corretamente o contribuinte")
       txtIM.SetFocus
       Exit Sub
    End If
    
    If txtPeriodo = "" Then
       Util.Informa ("Informe o Periodo Corretamente!")
       txtPeriodo.SetFocus
       Exit Sub
    End If
       
    If (Trim$(txtTMInscricao) <> "" Xor chkCancelada.Value = 1) And Trim$(txtTMNumNota) <> "" Then
        If DeduzValores Then
            Set Linha = grdSaida.ListItems.Add(, , IIf(chkCancelada.Value = 0, txtTMInscricao, txtIM))
            Linha.SubItems(1) = txtTMNumNota
            Linha.SubItems(2) = txtTMEmissao
            Linha.SubItems(3) = txtTMValor
            Linha.SubItems(4) = txtTMIcms
            Linha.SubItems(5) = txtTMSaldo
            Linha.SubItems(6) = txtTMImpostoDevido
            Linha.SubItems(7) = txtTMImpostoRetido
            Linha.SubItems(8) = txtTMSaldoDevedor
            Linha.SubItems(10) = txtTMAliq
            Linha.SubItems(11) = txtAidf
            Linha.SubItems(9) = chkCancelada.Value
            Linha.SubItems(12) = txtTomadorCpfCnpj
        End If
        If Trim(txtTMImpostoRetido) <> Trim(txtTMImpostoDevido) Then
            TotalBaseSaida = TotalBaseSaida + CDbl(Nvl(txtTMSaldo, 0))
        End If
        TotalImpostoRetidoSaida = TotalImpostoRetidoSaida + CDbl(Nvl(txtTMImpostoRetido, 0))
        TotalImpostoDevidoSaida = TotalImpostoDevidoSaida + CDbl(Nvl(txtTMImpostoDevido, 0))
        TotalICMSSujeito = TotalICMSSujeito + CDbl(Nvl(txtTMIcms, 0))
'        txtItemDecl(100) = (100 * CDbl(txtItemDecl(13))) / (TotalBaseSaida + TotalBaseST)
        AtualizaApuracao
        chkCancelada.Value = 0
        'txtTMInscricao = ""
        txtTMNumNota = 0
        txtTMEmissao = ""
        txtTMValor = 0
        txtTMIcms = 0
        txtTMSaldo = 0
        txtTMImpostoDevido = 0
        txtTMImpostoRetido = 0
        txtTMSaldoDevedor = 0
        
        chkCancelada.SetFocus
        txtTMNumNota.SetFocus
    Else
        Avisa "Informe os dados corretamente."
    End If
End Sub

Private Sub cmdBuscaTomador_Click()
AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtTomadorCpfCnpj, txtTomadorRazao
End Sub

Private Sub CmdConsultaContribuinte_Click()
    'blnConsultaIM = True
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIM
    'blnConsultaIM = False
End Sub

Private Sub cmdFinaliza_Click()
    Dim NumDec As String
    If Confirma("Ao finalizar a declarac�o, ela s� poder� ser modificada atrav�s de uma retificadora. Deseja prosseguir?") Then
        If Trim(txtItemDecl(1)) = "" Or Trim(txtItemDecl(2)) = "" Then
            Avisa "Informe nota inicial e nota final."
            TabDec.Tabs(2).Selected = True
            txtItemDecl(1).SetFocus
            Exit Sub
        End If
        If cboTipo.Coluna(1).ListIndex < 0 Then
            Avisa "Informe tipo da declaracao."
            cboTipo.SetFocus
            Exit Sub
        End If
        If CDbl(Nvl(Trim(txtItemDecl(13)), 0)) = 0 Then
            If Not Confirma("Valor Devido igual a zero. Prosseguir?") Then
                TabDec.Tabs(3).Selected = True
                txtItemDecl(3).SetFocus
                Exit Sub
            End If
        End If
        
        
        If Not Edita.CriticaCampos(Me) Then Exit Sub
        
        If txtIM = "" Then Exit Sub
    
        If Trim(txtItemDecl(1)) = "" Or Trim(txtItemDecl(2)) = "" Then
            Avisa "Informe nota inicial e nota final."
            TabDec.Tabs(3).Selected = True
            txtItemDecl(1).SetFocus
            Exit Sub
        End If
        If cboTipo.Coluna(1).ListIndex < 0 Then
            Avisa "Informe tipo da declaracao."
            cboTipo.SetFocus
            Exit Sub
        End If
        If txtPeriodo = "" Then
            Util.Avisa "Informe o per�odo."
            txtPeriodo.SetFocus
            Exit Sub
        End If
        If CDbl(Nvl(Trim(txtItemDecl(13)), 0)) = 0 Then
            If Not Confirma("Valor Devido igual a zero. Prosseguir?") Then
                TabDec.Tabs(3).Selected = True
                txtItemDecl(3).SetFocus
                Exit Sub
            End If
        End If
    
    Declaracao.Im = txtIM
    Declaracao.Periodo = txtPeriodo
    Declaracao.CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN))
    Declaracao.Data = Format(Date, "dd/mm/yyyy")
    Declaracao.Origem = orgSistema
    Declaracao.Recepcao = Date
    Declaracao.Versao = Nvl(Temp.PegaParametro(Bdados, "VERSAO DEC"), 2)
    Declaracao.Tipo = cboTipo.Coluna(1).Valor
    Declaracao.BaseGeral = TotalBaseSaida + TotalBaseST
    Declaracao.Status = decFinalizada
    CarregarItens
    If Declaracao.Gravar() Then
        Avisa "Declara��o gravada com sucesso."
        Declaracao.Salvar_Sem_Finalizar , , , , CDbl(Nvl(txtItemDecl(13), 0))
        txtIM.SetFocus
    End If
    End If
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If txtIM = "" Then Exit Sub
    
    If Trim(txtItemDecl(1)) = "" Or Trim(txtItemDecl(2)) = "" Then
        Avisa "Informe nota inicial e nota final."
        'TabDec.Tabs(3).Selected = True
        'txtItemDecl(1).SetFocus
        Exit Sub
    End If
    If cboTipo.Coluna(1).ListIndex < 0 Then
        Avisa "Informe tipo da declaracao."
        cboTipo.SetFocus
        Exit Sub
    End If
    If txtPeriodo = "" Then
        Util.Avisa "Informe o per�odo."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    If CDbl(Nvl(Trim(txtItemDecl(13)), 0)) = 0 Then
        If Not Confirma("Valor Devido igual a zero. Prosseguir?") Then
            TabDec.Tabs(2).Selected = True
            txtItemDecl(2).SetFocus
            Exit Sub
        End If
    End If
    
    Declaracao.Im = txtIM
    Declaracao.Periodo = txtPeriodo
    Declaracao.CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN))
    Declaracao.Data = Format(Date, "dd/mm/yyyy")
    Declaracao.Origem = orgSistema
    Declaracao.Recepcao = Date
    Declaracao.Versao = Nvl(Temp.PegaParametro(Bdados, "VERSAO DEC"), 2)
    Declaracao.Tipo = cboTipo.Coluna(1).Valor
    Declaracao.Status = decAberta
    Declaracao.BaseGeral = TotalBaseSaida + TotalBaseST
    CarregarItens
    If Declaracao.Gravar() Then
        Avisa "Declara��o gravada com sucesso."
        'Call Pega_taxas
        Declaracao.Salvar_Sem_Finalizar True, , , , CDbl(Nvl(txtItemDecl(13), 0))
        txtIM.SetFocus
    End If
End Sub

Private Sub cmLimpar_Click()
    Edita.LimpaCampos Me
'    grdDec.ListItems.Clear
    grdSaida.ListItems.Clear
    TabDec.Tabs(1).Selected = True
    IniciaTotalizadores
    txtIM.SetFocus
End Sub

Private Sub AtualizaDEC(CodModalidade As Integer)
    Dim Sql As String
    
    Sql = "SELECT TCD_TMD_COD_DECLARACAO,TCD_COD_CAMPO as COD,TCD_CAMPO as ITEM,TCD_VALOR_CAMPO AS VALOR FROM TAB_CONTEUDO_DECLARACAO WHERE TCD_TMD_COD_DECLARACAO = " & CodModalidade
End Sub
Private Sub Form_Load()
    Dim Sql As String
    IniciaTotalizadores
    
    cabVisual1.Exibir Bdados, Me.Name, App.Path
    'rodVISUAL1.Exibir Bdados, Me.Tag
    AtualizaDEC 0
    Set Imposto = New VsTFuncoes.VSImposto
    DeduzValores = True
    PrepararGrid grdSaida, 50
    Set Declaracao = New cDeclaracao
    cboTipo.PreencherGeral Bdados, "TIPO DECLARACAO"
'    TabDec.Tabs(4).Enabled = False
    Sql = "SELECT TCD_COD_CAMPO as Item ,TCD_CAMPO as Descricao, ' ' as Valor FROM " & _
        "TAB_CONTEUDO_DECLARACAO WHERE TCD_TMD_COD_DECLARACAO = " & 1
    Grdtaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
    Set Atividade = CreateObject("VsTEcon.atividade")
    
    
    txtPeriodo = Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Declaracao = Nothing
    Set Atividade = Nothing
    IniciaTotalizadores
End Sub

Private Sub grdDec_DblClick()
    On Error Resume Next
    If grdDec.ListItems.Count >= 1 Then
        cmLimpar_Click
        txtIM = grdDec.SelectedItem
        txtPeriodo = Right(grdDec.SelectedItem.SubItems(1), 2) & "/" & Left(grdDec.SelectedItem.SubItems(1), 4)
        cboTipo.SetarLinha grdDec.SelectedItem.SubItems(2), 1
        cboTipo_LostFocus
    End If
End Sub

Private Sub grdSaida_DblClick()
    On Error Resume Next
    
    If grdSaida.ListItems.Count = 0 Then Exit Sub
    AtualizaTextoSaidas grdSaida.SelectedItem.Index, True
    AtualizaApuracao
End Sub

Public Sub CarregarItens()
    Dim Controle As Object
    Dim Item As cItemDeclaracao
    
    Declaracao.Itens.Limpar
    Declaracao.Notas.Limpar
    'NOTAS DE ENTRADA(SUBSTITUICAO TRIBUTARIA)
    'NOTAS DE SAIDA(SERVICOS PRESTADOS)
    CarregaItensGrid grdSaida, etoSaida
    'APURACAO DO IMPOSTO
    For Each Controle In txtItemDecl
        If Trim(Controle.Text) <> "" And Trim(Controle.Text) <> "0,00" Then
            Set Item = New cItemDeclaracao
            Item.Numero = Controle.Index
            Item.Valor = Nvl(Controle.Text, 0)
            Declaracao.Itens.Adicionar Item
        End If
    Next
End Sub

Private Sub Grid_DblClick()
'ClassGrid.EditaCelula Grid, txtGrig
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then ClassGrid.TeclaPressionada Grid, txtGrig, KeyCode
End Sub

Private Sub txtGrig_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then ClassGrid.TextoKeyDown KeyCode, Grid, txtGrig
End Sub

Private Sub txtGrig_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtGrig_LostFocus()
'    txtGrig = Format(txtGrig, Const_Monetario)
End Sub

Private Sub txtIM_LostFocus()
    
    
    If Trim$(txtIM) <> "" Then
        If Not BuscarContribuinte(txtIM, txtRazao) Then
            Avisa "Contribuinte n�o encontrado" & vbCrLf & "Verifique se todos os dados est�o corretos."
            txtIM = "": txtRazao = ""
            txtIM.SetFocus
        Else
            Dim rs As VSRecordset
            If Bdados.AbreTabela("SELECT TCI_CGC_CPF FROM TAB_CONTRIBUINTE WHERE TCI_IM='" & txtIM & "'", rs) Then
                txtTMInscricao = IIf(IsNull(rs(0)), "", rs(0))
            End If
    '        Set atividade = New VsTEcon.atividade
            On Error Resume Next
            AliqISSQN = Atividade.BuscaAliquotaAtividade(Bdados, txtIM, txtPeriodo, ISSQNFixo)
            txtTMAliq = AliqISSQN * 100
            
            Declaracao.tciAtividade = Atividade.Nome
  '          Modalidade = BuscaModalidadeDeclaracao(txtIM, lblModalidade)
            AtualizaDEC Modalidade
'            TabDec.Tabs(4).Enabled = IIf(Modalidade > 0, True, False)
            Set Atividade = Nothing
            cboTipo_LostFocus
        End If
    End If
End Sub

Private Sub txtItemDecl_Change(Index As Integer)
    Select Case Index
        Case 6, 9, 12
            'BCP
            'txtItemDecl(13) = CDbl(Nvl(txtItemDecl(6), 0))
            'txtItemDecl(14) = Format(TotalImpostoRetidoSaida, Const_Monetario)
        Case 5
            If txtItemDecl(3).Enabled = True Then
                'TotalBaseSaida = CDbl(Nvl(txtItemDecl(5), 0))
            End If
    End Select
End Sub

Private Sub txtItemDecl_LostFocus(Index As Integer)
    Select Case Index
        Case 3, 4
            txtItemDecl(5) = CDbl(Nvl(txtItemDecl(3), 0)) - CDbl(Nvl(txtItemDecl(4), 0))
            CalcularImposto txtItemDecl(3), txtItemDecl(4), txtItemDecl(5), _
                            txtItemDecl(6), Nothing
        Case 7, 8
            txtItemDecl(8) = Nvl(txtItemDecl(8), 0)
            If CDbl(Nvl(txtItemDecl(9), 0)) - CDbl(Nvl(txtItemDecl(8), 0)) >= 0 Then
                txtItemDecl(7) = CDbl(Nvl(txtItemDecl(9), 0)) - CDbl(Nvl(txtItemDecl(8), 0))
            Else
                Avisa "Dados inv�lidos. Valor negativo encontrado." '& Nvl(txtItemDecl(9), 0) & " - " & Nvl(txtItemDecl(8), 0) & " = " & CDbl(CDbl(Nvl(txtItemDecl(9), 0)) - CDbl(Nvl(txtItemDecl(8), 0)))
                txtItemDecl(8).SetFocus
             End If
        Case 10, 11
        txtItemDecl(11) = Nvl(txtItemDecl(11), 0)
        If CDbl(Nvl(txtItemDecl(12), 0)) - CDbl(Nvl(txtItemDecl(11), 0)) >= 0 Then
            txtItemDecl(10) = CDbl(Nvl(txtItemDecl(12), 0)) - CDbl(Nvl(txtItemDecl(11), 0))
        Else
            Avisa "Dados inv�lidos. Valor negativo encontrado." '& Nvl(txtItemDecl(12), 0) & " - " & Nvl(txtItemDecl(11), 0) & " = " & CDbl(CDbl(Nvl(txtItemDecl(12), 0)) - CDbl(Nvl(txtItemDecl(11), 0)))
            txtItemDecl(11).SetFocus
        End If
    End Select
End Sub

Private Sub CalcularImposto(ByRef Total As Object, ByRef ICMS As Object, ByRef Tributavel As Object, ByRef Imposto As Object, ByRef Aliquota As Object)
    Total = Nvl(Trim$(Total.Text), 0)
    ICMS = Nvl(Trim$(ICMS.Text), 0)
   'Aliquota = AliqISSQN * 100
   'Tributavel = Total - ICMS
   'If AliqISSQN > 0 Then
   '    Imposto = Tributavel * AliqISSQN
   'Else
   '    Imposto = ISSQNFixo
   'End If
    
    Tributavel = Total - ICMS
    If AliqISSQN >= 0 Then
        Imposto = (CDbl(txtTMAliq) / 100) * CDbl(txtItemDecl(5))
    Else
        Imposto = ISSQNFixo
    End If
End Sub

Private Sub txtPeriodo_Change()
    
    If Len(Trim(txtPeriodo)) <> 6 Then Exit Sub
'    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
'    If Trim(txtIM) <> "" And cboTipo.ListIndex <> -1 Then PreencheDeclaracao
    AliqISSQN = 0 'Atividade.BuscaAliquotaAtividade(Bdados, txtIM, txtPeriodo, ISSQNFixo)
    If CInt(Left(Trim(txtPeriodo), 2)) > 12 Or CInt(Left(Trim(txtPeriodo), 2)) < 1 Then
        Avisa "Periodo inv�lido."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    cboTipo_LostFocus
    If Trim$(txtTMEmissao) = "" Then Exit Sub
    If Right(Trim(txtTMEmissao), 7) <> Trim(txtPeriodo) Then
        Avisa "Periodo da NF incompativel com periodo da declarac�o."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtPeriodo_LostFocus()
    If Len(Trim(txtPeriodo)) <> 6 Then Exit Sub
    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
End Sub


Private Function BuscarContribuinte(ByRef Inscricao As Object, Optional ByRef Nome As Object, Optional ByRef Endereco As Object, _
                    Optional ByRef Bairro As Object, Optional ByRef Cep As Object, Optional ByRef Cidade As Object, Optional ByRef Uf As Object) As Boolean
    Dim Im As Boolean
    Im = False
    If Trim(Inscricao) = "" Then Exit Function
    Inscricao.Text = Edita.TiraTudo(Inscricao.Text)
    If Len(Inscricao.Text) = 10 Then Im = True
    FormataRegistro Inscricao
    If Trim(Inscricao) = "" Then Exit Function
    
    Dim Sql As String, rs As VSRecordset
    Sql = "SELECT tci_im, TCI_CGC_CPF,tci_nome, tci_logradouro, tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade, tci_UF " & _
            ",TAE_NOME FROM TAB_CONTRIBUINTE LEFT JOIN TAB_ATIVIDADE_ECONOMICA ON TCI_TAE_CAE = TAE_CAE WHERE 1=1"
    If Im Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        Sql = Sql & " AND TCI_IM='" & Inscricao & "'"
    Else
        Sql = Sql & " AND TCI_CGC_CPF='" & Inscricao & "'"
    End If
    
    If Bdados.AbreTabela(Sql, rs) Then
        If Im Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            Inscricao = "" & rs!tci_im
        Else
            Inscricao = "" & rs!TCI_CGC_CPF
        End If
        If Not Nome Is Nothing Then Nome = "" & rs!tci_nome
        If Not Endereco Is Nothing Then Endereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & ", " & rs!tci_numero & " " & rs!tci_complemento
        If Not Bairro Is Nothing Then Bairro = "" & rs!tci_bairro
        If Not Cep Is Nothing Then Cep = "" & rs!tci_cep
        If Not Cidade Is Nothing Then Cidade = "" & rs!tci_cidade
        If Not Uf Is Nothing Then Uf = "" & rs!tci_UF
        With Declaracao
            .tciNome = "" & rs!tci_nome
            .tciEndereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & ", " & rs!tci_numero & " " & rs!tci_complemento
            .tciBairro = "" & rs!tci_bairro
            .tciCEP = "" & rs!tci_cep
            .tciCidade = "" & rs!tci_cidade
            .tciUF = "" & rs!tci_UF
            .tciEndereco = .tciEndereco & " " & .tciBairro & " " & .tciCidade & "-" & rs!tci_UF
            .tciAtividade = "" & rs!TAE_NOME
        End With
        BuscarContribuinte = True
    End If
    Bdados.FechaTabela rs
End Function

Private Sub PrepararGrid(Nome_Grid As Object, IndiceInicial As Byte)
    Nome_Grid.ColumnHeaders.Clear
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Inscricao": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Nota": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Emiss�o": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Valor da Nota": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Sujeito ICMS": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Tributavel": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Imposto Devido": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Imposto Retido": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Saldo Devedor": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Cancelada": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Aliquota": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "AIDF": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Tomador": IndiceInicial = IndiceInicial + 1
    
    
    
End Sub

Private Sub txtTMAliq_Change()
 Calc_ISSQN
End Sub
Private Sub txtTMAliq_LostFocus()
 Calc_ISSQN
End Sub
Private Sub txtTMEmissao_LostFocus()
    If Trim$(txtTMEmissao) = "" Then Exit Sub
    If Trim(txtPeriodo) = "" Then Exit Sub
    If Right(Trim(txtTMEmissao), 7) <> Trim(txtPeriodo) Then
        Avisa "Periodo da NF incompativel com periodo da declarac�o."
        txtTMEmissao.SetFocus
    End If
End Sub

Private Sub txtTMIcms_LostFocus()
    txtTMIcms = Nvl(Trim$(txtTMIcms), 0)
    txtTMSaldo = CDbl(Nvl(txtTMValor, 0)) - CDbl(txtTMIcms)
    If AliqISSQN > 0 Then
        txtTMImpostoDevido = txtTMSaldo * AliqISSQN
    Else
        txtTMImpostoDevido = ISSSTFixo
    End If
End Sub

Private Sub txtTMImpostoDevido_Change()
    txtTMSaldoDevedor = CDbl(Nvl(txtTMImpostoDevido, 0)) - CDbl(Nvl(txtTMImpostoRetido, 0))
End Sub

Private Sub txtTMImpostoRetido_Change()
    txtTMSaldoDevedor = CDbl(Nvl(txtTMImpostoDevido, 0)) - CDbl(Nvl(txtTMImpostoRetido, 0))
End Sub

Private Sub txtTMImpostoRetido_LostFocus()
    txtTMImpostoRetido = Nvl(Trim$(txtTMImpostoRetido), 0)
End Sub

Private Sub txtTMInscricao_LostFocus()
    
        BuscarContribuinte txtTMInscricao
End Sub

Private Sub txtTMValor_LostFocus()
    txtTMSaldo = CDbl(Nvl(txtTMValor, 0)) - CDbl(Nvl(txtTMIcms, 0))
    If AliqISSQN > 0 Then
        txtTMImpostoDevido = txtTMSaldo * AliqISSQN
    Else
        txtTMImpostoDevido = ISSSTFixo
    End If
End Sub
Private Sub Calc_ISSQN()
    'AliqISSQN = (txtItemDecl(3) - ((CCur(txtTMAliq) * txtItemDecl(3)) / 100))
    'txtItemDecl_LostFocus 3
    If txtTMSaldo = "" Then txtTMSaldo = 0
    If CCur(txtTMSaldo) > 0 Then
        txtTMImpostoDevido = CCur(txtTMSaldo * (txtTMAliq / 100))
    End If
End Sub
Private Sub Pega_taxas()
    Dim i As Integer
    Dim pos As Integer
    String_Taxas = ""
    Total_Taxas = 0
    For i = 1 To Grdtaxas.ListItems.Count
        If Grdtaxas.ListItems(i).Checked Then
            pos = InStr(Grdtaxas.ListItems(i).SubItems(1), "-") - 1
            If String_Taxas = "" Then
                String_Taxas = String_Taxas & " [ " & Left(Grdtaxas.ListItems(i).SubItems(1), pos) & " ]" & " - " & Format(Grdtaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            Else
                String_Taxas = String_Taxas & ", [ " & Left(Grdtaxas.ListItems(i).SubItems(1), pos) & " ]" & " - " & Format(Grdtaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            End If
            Total_Taxas = Total_Taxas + CCur(Grdtaxas.ListItems(i).SubItems(2))
        End If
    Next
End Sub

