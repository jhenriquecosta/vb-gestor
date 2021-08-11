VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TDNT105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TDNT105"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "TDNT105.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   42
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TDNT105.frx":08CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   43
      Top             =   6945
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   345
         Left            =   9105
         TabIndex        =   52
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   6225
         TabIndex        =   37
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   7185
         TabIndex        =   38
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   10065
         TabIndex        =   40
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   8145
         TabIndex        =   39
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TDNT105.frx":29ED
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1725
      Left            =   60
      TabIndex        =   45
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   675
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   3043
      Altura          =   1905
      Caption         =   " Dados do Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   300
         Left            =   10110
         TabIndex        =   3
         Top             =   360
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   6675
         TabIndex        =   2
         Top             =   360
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtFolhaConsulta 
         Height          =   285
         Left            =   6840
         TabIndex        =   5
         Top             =   1335
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   503
         Caption         =   "Folha"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   0
         MaxLen          =   4
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtLivroConsulta 
         Height          =   285
         Left            =   8715
         TabIndex        =   6
         Top             =   1335
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   503
         Caption         =   "Livro"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   0
         MaxLen          =   50
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRegistroConsulta 
         Height          =   285
         Left            =   4815
         TabIndex        =   4
         Top             =   1335
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   503
         Caption         =   "Registro"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   0
         MaxLen          =   8
         AgruparValores  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   3525
         TabIndex        =   1
         Top             =   375
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   30
         TabIndex        =   41
         Top             =   690
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   503
         Caption         =   "Nome/Razão Social"
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   0
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   510
         TabIndex        =   0
         Tag             =   "Insc. Municipal"
         Top             =   375
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   0
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   900
         TabIndex        =   46
         Top             =   1005
         Width           =   9630
         _ExtentX        =   16986
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   0
         CorTexto        =   4194304
      End
   End
   Begin ActiveTabs.SSActiveTabs tabNotificacao 
      Height          =   4500
      Left            =   45
      TabIndex        =   47
      Top             =   2415
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   7938
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
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
      Tabs            =   "TDNT105.frx":2D07
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4110
         Index           =   0
         Left            =   30
         TabIndex        =   48
         Top             =   30
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   7250
         _Version        =   131082
         TabGuid         =   "TDNT105.frx":2D8C
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   4065
            Left            =   15
            TabIndex        =   51
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   15
            Width           =   10830
            _ExtentX        =   19103
            _ExtentY        =   7170
            Altura          =   1905
            Caption         =   " Dívida Ativa"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL CboStatus 
               Height          =   510
               Left            =   8145
               TabIndex        =   53
               ToolTipText     =   "STATUS DAT"
               Top             =   285
               Width           =   2640
               _ExtentX        =   4657
               _ExtentY        =   900
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtCargo 
               Height          =   480
               Left            =   5445
               TabIndex        =   36
               Top             =   3495
               Width           =   5310
               _ExtentX        =   9366
               _ExtentY        =   847
               Caption         =   "Cargo"
               Text            =   ""
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   200
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtAutoridade 
               Height          =   480
               Left            =   75
               TabIndex        =   35
               Top             =   3495
               Width           =   5370
               _ExtentX        =   9472
               _ExtentY        =   847
               Caption         =   "Autoridade"
               Text            =   ""
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   200
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNAuto 
               Height          =   480
               Left            =   8580
               TabIndex        =   34
               Top             =   3015
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   847
               Caption         =   "Nº Auto"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   10
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtProcesso 
               Height          =   480
               Left            =   75
               TabIndex        =   33
               Top             =   3015
               Width           =   8520
               _ExtentX        =   15028
               _ExtentY        =   847
               Caption         =   "Processo"
               Text            =   ""
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   200
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNProcesso 
               Height          =   480
               Left            =   4230
               TabIndex        =   32
               Top             =   2490
               Width           =   6540
               _ExtentX        =   11536
               _ExtentY        =   847
               Caption         =   "Nº Processo"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   100
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtUsuario 
               Height          =   480
               Left            =   75
               TabIndex        =   31
               Top             =   2490
               Width           =   4170
               _ExtentX        =   7355
               _ExtentY        =   847
               Caption         =   "Usuário"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   1
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNTida 
               Height          =   480
               Left            =   9465
               TabIndex        =   30
               Top             =   1965
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   847
               Caption         =   "Nº TIDA"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   8
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDataTida 
               Height          =   480
               Left            =   8175
               TabIndex        =   29
               Top             =   1965
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   847
               Caption         =   "Data TIDA"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   10
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNCDA 
               Height          =   480
               Left            =   6855
               TabIndex        =   28
               Top             =   1965
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   847
               Caption         =   "Nº CDA"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   8
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDataCda 
               Height          =   480
               Left            =   5535
               TabIndex        =   27
               Top             =   1965
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   847
               Caption         =   "Data CDA"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   10
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNMalic 
               Height          =   480
               Left            =   4215
               TabIndex        =   26
               Top             =   1965
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   847
               Caption         =   "Nº MALIC"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   8
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDataMalic 
               Height          =   480
               Left            =   2850
               TabIndex        =   25
               Top             =   1965
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   847
               Caption         =   "Data MALIC"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   10
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNMacal 
               Height          =   480
               Left            =   1455
               TabIndex        =   24
               Top             =   1965
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   847
               Caption         =   "Nº MACAL"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   8
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDataMacal 
               Height          =   480
               Left            =   75
               TabIndex        =   23
               Top             =   1965
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   847
               Caption         =   "Data MACAL"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   10
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtInscData 
               Height          =   480
               Left            =   9465
               TabIndex        =   22
               Top             =   1410
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   847
               Caption         =   "Insc. Data"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   10
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtLivro 
               Height          =   480
               Left            =   9435
               TabIndex        =   14
               Top             =   825
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   847
               Caption         =   "Livro"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   50
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDividaTotal 
               Height          =   480
               Left            =   8175
               TabIndex        =   21
               Top             =   1410
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   847
               Caption         =   "Dívida Total"
               Text            =   ""
               Formato         =   5
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   20
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCorrecao 
               Height          =   480
               Left            =   6855
               TabIndex        =   20
               Top             =   1410
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   847
               Caption         =   "Correção"
               Text            =   ""
               Formato         =   5
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   20
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtMulta 
               Height          =   480
               Left            =   5535
               TabIndex        =   19
               Top             =   1410
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   847
               Caption         =   "Multa"
               Text            =   ""
               Formato         =   5
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   20
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtJuros 
               Height          =   480
               Left            =   4215
               TabIndex        =   18
               Top             =   1410
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   847
               Caption         =   "Juros"
               Text            =   ""
               Formato         =   5
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   20
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDividaOriginal 
               Height          =   480
               Left            =   2850
               TabIndex        =   17
               Top             =   1410
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   847
               Caption         =   "Dívida Original"
               Text            =   ""
               Formato         =   5
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   20
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtRefDivida 
               Height          =   480
               Left            =   1455
               TabIndex        =   16
               Top             =   1410
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   847
               Caption         =   "Ref. Dívida"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   4
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtFolha 
               Height          =   480
               Left            =   8145
               TabIndex        =   13
               Top             =   825
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   847
               Caption         =   "Folha"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   4
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtAnoDivida 
               Height          =   480
               Left            =   75
               TabIndex        =   15
               Top             =   1410
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   847
               Caption         =   "Ano Dívida"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   4
               AgruparValores  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboImposto 
               Height          =   510
               Left            =   60
               TabIndex        =   12
               Top             =   810
               Width           =   8100
               _ExtentX        =   14288
               _ExtentY        =   900
               Caption         =   "Imposto"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboNaturezaTributo 
               Height          =   510
               Left            =   5910
               TabIndex        =   11
               Top             =   285
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   900
               Caption         =   "Natureza doTributo"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtEdital 
               Height          =   480
               Left            =   4425
               TabIndex        =   10
               Top             =   300
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   847
               Caption         =   "Edital"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   4
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtInscricao 
               Height          =   480
               Left            =   2970
               TabIndex        =   9
               Top             =   300
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   847
               Caption         =   "Inscrição"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   20
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtObrigacao 
               Height          =   480
               Left            =   1515
               TabIndex        =   8
               Top             =   300
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   847
               Caption         =   "Obrigação"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   8
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtRegistro 
               Height          =   480
               Left            =   60
               TabIndex        =   7
               Top             =   300
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   847
               Caption         =   "Registro"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               MaxLen          =   8
               AgruparValores  =   0   'False
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4110
         Left            =   -99969
         TabIndex        =   49
         Top             =   30
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   7250
         _Version        =   131082
         TabGuid         =   "TDNT105.frx":2DB4
         Begin VTOcx.grdVISUAL grdNotifica 
            Height          =   3705
            Left            =   0
            TabIndex        =   50
            Top             =   60
            Width           =   10830
            _ExtentX        =   19103
            _ExtentY        =   6535
            CorBorda        =   16384
            Caption         =   "Dívida Ativa"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
         End
      End
   End
End
Attribute VB_Name = "TDNT105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscar_Click()
     CarregaDivida
     tabNotificacao.Tabs(1).Selected = True
End Sub

Private Sub cmdExcluir_Click()
    Dim condicao As String
    If grdNotifica.ListItems.Count < 1 Then Exit Sub
    condicao = "TDA_REGISTRO = " & txtRegistro & " AND TDA_TOC_COD_OBRIGACAO = " & txtObrigacao & " AND TDA_INSCRICAO = " & txtInscricao & " AND TDA_TIPO_DIVIDA = 2"
    If txtRegistro <> "" Then
        If Confirma("Deseja excluir registro?", "Excluir?") Then
            If Bdados.DeletaDados(" TAB_DIVIDA_ATIVA", condicao) Then
                Avisa "Dados Excluidos com Sucesso"
                limparCampos
                CarregaDivida
                tabNotificacao.Tabs(1).Selected = True
            End If
        End If
    Else
        Avisa "Selecione um Registro"
    End If
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub
Private Sub txtImovel_LostFocus()
    Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtIm.SetFocus
        End If
    End If
End Sub

Private Sub grdNotifica_DblClick()
        If grdNotifica.ListItems.Count < 1 Then Exit Sub
        tabNotificacao.Tabs(2).Selected = True
        txtRegistro = grdNotifica.SelectedItem
        txtObrigacao = grdNotifica.SelectedItem.SubItems(1)
        txtInscricao = grdNotifica.SelectedItem.SubItems(2)
        txtEdital = grdNotifica.SelectedItem.SubItems(4)
        cboNaturezaTributo.SetarLinha grdNotifica.SelectedItem.SubItems(29), 1
        cboImposto.SetarLinha grdNotifica.SelectedItem.SubItems(30), 1
        txtFolha = grdNotifica.SelectedItem.SubItems(7)
        txtLivro = grdNotifica.SelectedItem.SubItems(8)
        txtAnoDivida = grdNotifica.SelectedItem.SubItems(9)
        txtRefDivida = grdNotifica.SelectedItem.SubItems(10)
        txtDividaOriginal = grdNotifica.SelectedItem.SubItems(11)
        txtJuros = grdNotifica.SelectedItem.SubItems(12)
        txtMulta = grdNotifica.SelectedItem.SubItems(13)
        txtCorrecao = grdNotifica.SelectedItem.SubItems(14)
        txtDividaTotal = grdNotifica.SelectedItem.SubItems(15)
        txtInscData = grdNotifica.SelectedItem.SubItems(16)
        txtDataMacal = grdNotifica.SelectedItem.SubItems(17)
        txtNMacal = grdNotifica.SelectedItem.SubItems(18)
        txtDataMalic = grdNotifica.SelectedItem.SubItems(19)
        txtNMalic = grdNotifica.SelectedItem.SubItems(20)
        txtDataCda = grdNotifica.SelectedItem.SubItems(21)
        txtNCDA = grdNotifica.SelectedItem.SubItems(22)
        txtDataTida = grdNotifica.SelectedItem.SubItems(23)
        txtNTida = grdNotifica.SelectedItem.SubItems(24)
        txtUsuario = grdNotifica.SelectedItem.SubItems(25)
        txtNProcesso = grdNotifica.SelectedItem.SubItems(26)
        txtProcesso = grdNotifica.SelectedItem.SubItems(27)
        txtNAuto = grdNotifica.SelectedItem.SubItems(28)
        txtAutoridade = grdNotifica.SelectedItem.SubItems(31)
        txtCargo = grdNotifica.SelectedItem.SubItems(32)
        cboStatus.SetarLinha grdNotifica.SelectedItem.SubItems(33), 1
End Sub



Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim campos As String
    Dim Valores As String
    Dim condicao As String
    txtIm.Tag = ""
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
   
    campos = " TDA_REGISTRO ,TDA_TOC_COD_OBRIGACAO ,TDA_INSCRICAO ,"
    campos = campos & " TDA_TED_TCOD_EDITAL ,TDA_NATUREZA_TRIBUTO ,"
    campos = campos & " TDA_TIP_COD_IMPOSTO ,TDA_FOLHA ,TDA_LIVRO,TDA_ANO_DIVIDA,"
    campos = campos & " TDA_REFERENCIA_DIVIDA ,TDA_DIVIDA_ORIGINAL ,TDA_JUROS,"
    campos = campos & " TDA_MULTA ,TDA_CORRECAO ,TDA_DIVIDA_TOTAL ,TDA_DATA_INSCRICAO,TDA_MACAL_DATA,"
    campos = campos & " TDA_MACAL_NUMERO ,TDA_MALIC_DATA ,TDA_MALIC_NUMERO ,TDA_CDA_DATA ,TDA_CDA_NUMERO ,"
    campos = campos & " TDA_TIDA_DATA ,TDA_TIDA_NUMERO ,TDA_TUS_COD_USUARIO ,TDA_NUM_PROCESSO,"
    campos = campos & " TDA_PROCESSO , TDA_TAI_NUM_AUTO,TDA_AUTORIDADE,TDA_CARGO,TDA_TIPO_DIVIDA,TDA_STATUS"
    
    Valores = Bdados.PreparaValor(txtRegistro, txtObrigacao, txtInscricao, txtEdital, cboNaturezaTributo.Coluna(1).Valor, _
              cboImposto.Coluna(1).Valor, txtFolha, txtLivro, txtAnoDivida, txtRefDivida, _
              txtDividaOriginal, txtJuros, txtMulta, txtCorrecao, txtDividaTotal, _
              txtInscData, txtDataMacal, txtNMacal, txtDataMalic, txtNMalic, _
              txtDataCda, txtNCDA, txtDataTida, txtNTida, txtUsuario, txtNProcesso, _
              txtProcesso, txtNAuto, txtAutoridade, txtCargo, 2, cboStatus.Coluna(1).Valor)
    
    condicao = "TDA_REGISTRO = " & txtRegistro & " AND TDA_TOC_COD_OBRIGACAO = " & txtObrigacao & " AND TDA_INSCRICAO = " & txtInscricao & " AND TDA_TIPO_DIVIDA = 2"
    
    If Bdados.AtualizaDados("TAB_DIVIDA_ATIVA", Valores, campos, condicao) Then
        Avisa "Dados gravados com sucesso"
        limparCampos
        CarregaDivida
        tabNotificacao.Tabs(1).Selected = True
    End If
End Sub
    

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     cboImposto.Preencher Bdados, "SELECT     tip_nome_imposto, tip_cod_imposto FROM Tab_Imposto"
     cboNaturezaTributo.Preencher Bdados, "SELECT   TGE_NOME , TGE_CODIGO  FROM  VIS_NATUREZA"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
    cboStatus.PreencherGeral Bdados, cboStatus.ToolTipText
End Sub

Private Sub CarregaDivida()
    Dim Sql As String
    
    Sql = "select TDA_REGISTRO as Registro,"
    Sql = Sql & " TDA_TOC_COD_OBRIGACAO AS Obrigação,"
    Sql = Sql & " TDA_INSCRICAO as Insc_Cadastral ,"
    Sql = Sql & " TDA_TCI_IM as Insc_Municipal ,"
    Sql = Sql & " TDA_TED_TCOD_EDITAL as Edital,"
    Sql = Sql & " TGE_NOME as Tributo,"
    Sql = Sql & " tip_nome_imposto as Imposto,"
    Sql = Sql & " TDA_FOLHA as Folha,"
    Sql = Sql & " TDA_LIVRO as Livro,"
    Sql = Sql & " TDA_ANO_DIVIDA as Ano_Dívida,"
    Sql = Sql & " TDA_REFERENCIA_DIVIDA as Referência_Dívida,"
    Sql = Sql & " TDA_DIVIDA_ORIGINAL as Dívida_Original,"
    Sql = Sql & " TDA_JUROS as Juros,"
    Sql = Sql & " TDA_MULTA as Multa,"
    Sql = Sql & " TDA_CORRECAO as Correção,"
    Sql = Sql & " TDA_DIVIDA_TOTAL as Dívida_Total,"
    Sql = Sql & " TDA_DATA_INSCRICAO as Data_Inscrição,"
    Sql = Sql & " TDA_MACAL_DATA as Data_Macal,"
    Sql = Sql & " TDA_MACAL_NUMERO as Número_Macal,"
    Sql = Sql & " TDA_MALIC_DATA as Data_Malic,"
    Sql = Sql & " TDA_MALIC_NUMERO as Número_Malic ,"
    Sql = Sql & " TDA_CDA_DATA as Data_CDA ,"
    Sql = Sql & " TDA_CDA_NUMERO as Número_CDA,"
    Sql = Sql & " TDA_TIDA_DATA as Data_Tida ,"
    Sql = Sql & " TDA_TIDA_NUMERO as Número_Tida,"
    Sql = Sql & " TDA_TUS_COD_USUARIO as Usuário,"
    Sql = Sql & " TDA_NUM_PROCESSO as Número_Processo ,"
    Sql = Sql & " TDA_PROCESSO as Processo,"
    Sql = Sql & " TDA_TAI_NUM_AUTO As Número_Auto,"
    Sql = Sql & " TDA_NATUREZA_TRIBUTO as Natureza_tributo,"
    Sql = Sql & " TDA_TIP_COD_IMPOSTO as Código_Imposto,  "
    Sql = Sql & " TDA_AUTORIDADE AS Autoridade,"
    Sql = Sql & " TDA_CARGO AS Cargo,"
    Sql = Sql & " TDA_STATUS AS Status"
    Sql = Sql & " From tab_divida_ativa,  "
    Sql = Sql & " VIS_NATUREZA,  "
    Sql = Sql & " Tab_Imposto  "
    Sql = Sql & " WHERE TGE_CODIGO = TDA_NATUREZA_TRIBUTO"
    Sql = Sql & " AND tip_cod_imposto = TDA_TIP_COD_IMPOSTO AND TDA_TIPO_DIVIDA = 2"
    
    If txtImovel <> "" Then Sql = Sql & " AND TDA_INSCRICAO = " & txtImovel
    If txtIm <> "" Then Sql = Sql & " AND TDA_TCI_IM = " & txtIm
    If txtRegistroConsulta <> "" Then Sql = Sql & " AND TDA_REGISTRO = " & txtRegistroConsulta
    If txtFolhaConsulta <> "" Then Sql = Sql & " AND TDA_FOLHA = " & txtFolhaConsulta
    If txtLivroConsulta <> "" Then Sql = Sql & " AND TDA_LIVRO = " & txtLivroConsulta
    
    If Not grdNotifica.Preencher(Bdados, Sql) Then Avisa "Busca sem Resultados."
End Sub

Private Sub limparCampos()
       
        txtRegistro = ""
        txtObrigacao = ""
        txtInscricao = ""
        txtEdital = ""
        cboNaturezaTributo = ""
        cboImposto = ""
        txtFolha = ""
        txtLivro = ""
        txtAnoDivida = ""
        txtRefDivida = ""
        txtDividaOriginal = ""
        txtJuros = ""
        txtMulta = ""
        txtCorrecao = ""
        txtDividaTotal = ""
        txtInscData = ""
        txtDataMacal = ""
        txtNMacal = ""
        txtDataMalic = ""
        txtNMalic = ""
        txtDataCda = ""
        txtNCDA = ""
        txtDataTida = ""
        txtNTida = ""
        txtUsuario = ""
        txtNProcesso = ""
        txtProcesso = ""
        txtNAuto = ""
End Sub
