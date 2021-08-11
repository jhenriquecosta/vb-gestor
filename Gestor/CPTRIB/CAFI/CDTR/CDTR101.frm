VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form CDTR101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDTR101"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs ssact 
      Height          =   4950
      Left            =   0
      TabIndex        =   58
      Top             =   1440
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   8731
      _Version        =   131082
      TabCount        =   4
      TagVariant      =   ""
      Tabs            =   "CDTR101.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   4560
         Left            =   -99969
         TabIndex        =   68
         Top             =   360
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   8043
         _Version        =   131082
         TabGuid         =   "CDTR101.frx":00EB
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   8085
            Left            =   0
            TabIndex        =   69
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   0
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   14261
            Altura          =   1905
            Caption         =   " Ponto/ Posto"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VTOcx.cmdVISUAL cmdLista 
               Height          =   345
               Left            =   9120
               TabIndex        =   74
               Top             =   360
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   609
               Caption         =   "Lista"
               Acao            =   4
               CorBorda        =   32768
               CorFrente       =   16384
               CorFoco         =   14737632
            End
            Begin VTOcx.cmdVISUAL cmdFicha 
               Height          =   345
               Left            =   8040
               TabIndex        =   73
               Top             =   360
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   609
               Caption         =   "Ficha"
               Acao            =   4
               CorBorda        =   32768
               CorFrente       =   16384
               CorFoco         =   14737632
            End
            Begin VTOcx.cmdVISUAL cmdBuscar 
               Height          =   345
               Left            =   6960
               TabIndex        =   72
               Top             =   360
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   609
               Caption         =   "Buscar"
               Acao            =   5
               CorBorda        =   32768
               CorFrente       =   16384
               CorFoco         =   14737632
            End
            Begin VTOcx.txtVISUAL txtNomeConsulta 
               Height          =   285
               Left            =   120
               TabIndex        =   71
               Top             =   360
               Width           =   6765
               _ExtentX        =   11933
               _ExtentY        =   503
               Caption         =   "Nome"
               Text            =   ""
               CorRotulo       =   4210752
               CorTexto        =   4194304
            End
            Begin VTOcx.grdVISUAL grd 
               Height          =   3735
               Left            =   120
               TabIndex        =   70
               Top             =   840
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   6588
               Caption         =   "Mototaxistas"
            End
            Begin VB.Image Image2 
               BorderStyle     =   1  'Fixed Single
               DragMode        =   1  'Automatic
               Height          =   3480
               Left            =   7320
               Stretch         =   -1  'True
               Top             =   840
               Width           =   2760
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   4560
         Left            =   -99969
         TabIndex        =   62
         Top             =   360
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   8043
         _Version        =   131082
         TabGuid         =   "CDTR101.frx":0113
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   1845
            Left            =   0
            TabIndex        =   63
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   120
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   3254
            Altura          =   1905
            Caption         =   " Dados da CNH"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VTOcx.txtVISUAL txtDataPrimHabilitacao 
               Height          =   480
               Left            =   3720
               TabIndex        =   31
               Top             =   360
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   847
               Caption         =   "Data 1. Habilitação"
               Text            =   ""
               TipoLetras      =   0
               Formato         =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtCatCnh 
               Height          =   480
               Left            =   2880
               TabIndex        =   30
               Top             =   360
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   847
               Caption         =   "Cat."
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtOrgaoCurso 
               Height          =   480
               Left            =   3720
               TabIndex        =   35
               Top             =   1320
               Width           =   6435
               _ExtentX        =   11351
               _ExtentY        =   847
               Caption         =   ""
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtOrgaoEmissaoCnh 
               Height          =   480
               Left            =   120
               TabIndex        =   34
               Top             =   1320
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   847
               Caption         =   "Orgão Emissor / UF"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtRegHab 
               Height          =   480
               Left            =   120
               TabIndex        =   32
               Top             =   840
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   847
               Caption         =   "Registro"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtCnh 
               Height          =   480
               Left            =   120
               TabIndex        =   29
               Top             =   360
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   847
               Caption         =   "Cnh"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VB.CheckBox chkCusoEspecializado 
               Caption         =   "Curso Especializado"
               Height          =   195
               Left            =   3720
               TabIndex        =   33
               Top             =   1080
               Width           =   2055
            End
         End
         Begin VTOcx.fraVISUAL fraVeiculo 
            Height          =   2520
            Left            =   0
            TabIndex        =   64
            Top             =   2040
            Width           =   10440
            _ExtentX        =   18415
            _ExtentY        =   4445
            Altura          =   1905
            Caption         =   " Dados do Veículo"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VTOcx.txtVISUAL txtPropriedade 
               Height          =   480
               Left            =   6120
               TabIndex        =   67
               Top             =   1920
               Width           =   4275
               _ExtentX        =   7541
               _ExtentY        =   847
               Caption         =   "Categoria/ Propriedade"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtBairroProprietario 
               Height          =   480
               Left            =   6120
               TabIndex        =   50
               Top             =   1380
               Width           =   4275
               _ExtentX        =   7541
               _ExtentY        =   847
               Caption         =   "Bairro"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtCidadeProprietario 
               Height          =   480
               Left            =   120
               TabIndex        =   51
               Top             =   1920
               Width           =   5955
               _ExtentX        =   10504
               _ExtentY        =   847
               Caption         =   "Cidade"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtNumeroProprietario 
               Height          =   480
               Left            =   5325
               TabIndex        =   49
               Tag             =   "Licenciamento"
               Top             =   1380
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   847
               Caption         =   "Número"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtLicenciamentoVeiculo 
               Height          =   480
               Left            =   135
               TabIndex        =   43
               Tag             =   "Modelo"
               Top             =   840
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   847
               Caption         =   "Licenciamento"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtRegistroVeiculo 
               Height          =   480
               Left            =   3675
               TabIndex        =   45
               Tag             =   "Município"
               Top             =   840
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   847
               Caption         =   "Registro"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtPotencia 
               Height          =   480
               Left            =   9480
               TabIndex        =   42
               Tag             =   "Marca"
               Top             =   300
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   847
               Caption         =   "Pot. Motor"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtEndeProprietario 
               Height          =   480
               Left            =   105
               TabIndex        =   48
               Tag             =   "Ano"
               Top             =   1380
               Width           =   5220
               _ExtentX        =   9208
               _ExtentY        =   847
               Caption         =   "Endereço"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtVeiculo 
               Height          =   480
               Left            =   135
               TabIndex        =   36
               Tag             =   "Veículo"
               Top             =   300
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   847
               Caption         =   "Veículo"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtModelo 
               Height          =   480
               Left            =   3675
               TabIndex        =   38
               Tag             =   "Modelo"
               Top             =   300
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   847
               Caption         =   "Modelo"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtMarca 
               Height          =   480
               Left            =   2145
               TabIndex        =   37
               Tag             =   "Marca"
               Top             =   300
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   847
               Caption         =   "Marca"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtChassi 
               Height          =   480
               Left            =   7365
               TabIndex        =   41
               Tag             =   "Chassi"
               Top             =   300
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   847
               Caption         =   "Chassi"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtPlaca 
               Height          =   480
               Left            =   6105
               TabIndex        =   40
               Tag             =   "Placa"
               Top             =   300
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   847
               Caption         =   "Placa"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtAnoFabric 
               Height          =   480
               Left            =   5325
               TabIndex        =   39
               Tag             =   "Ano"
               Top             =   300
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   847
               Caption         =   "Ano Fab."
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtMunicipioVeiculo 
               Height          =   480
               Left            =   2145
               TabIndex        =   44
               Tag             =   "Município"
               Top             =   840
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   847
               Caption         =   "Município"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtEstadoVeiculo 
               Height          =   480
               Left            =   5325
               TabIndex        =   46
               Tag             =   "Licenciamento"
               Top             =   840
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   847
               Caption         =   "UF"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtProprietarioVeiculo 
               Height          =   480
               Left            =   6105
               TabIndex        =   47
               Top             =   840
               Width           =   4275
               _ExtentX        =   7541
               _ExtentY        =   847
               Caption         =   "Proprietario"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4560
         Left            =   -99969
         TabIndex        =   61
         Top             =   360
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   8043
         _Version        =   131082
         TabGuid         =   "CDTR101.frx":013B
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   2520
            Left            =   0
            TabIndex        =   65
            Top             =   120
            Width           =   10440
            _ExtentX        =   18415
            _ExtentY        =   4445
            Altura          =   1905
            Caption         =   " Informações Profissionais"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VTOcx.txtVISUAL txtLocalAtividade 
               Height          =   480
               Left            =   4125
               TabIndex        =   20
               Tag             =   "Licenciamento"
               Top             =   840
               Width           =   6180
               _ExtentX        =   10901
               _ExtentY        =   847
               Caption         =   "Local"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtGrauEscolar 
               Height          =   480
               Left            =   4125
               TabIndex        =   18
               Tag             =   "Ano"
               Top             =   300
               Width           =   6180
               _ExtentX        =   10901
               _ExtentY        =   847
               Caption         =   "Grau Escolar"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtProfissao 
               Height          =   480
               Left            =   135
               TabIndex        =   17
               Tag             =   "Veículo"
               Top             =   300
               Width           =   3930
               _ExtentX        =   6932
               _ExtentY        =   847
               Caption         =   "Profissão"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtEndAtividade 
               Height          =   480
               Left            =   105
               TabIndex        =   21
               Tag             =   "Ano"
               Top             =   1380
               Width           =   5220
               _ExtentX        =   9208
               _ExtentY        =   847
               Caption         =   "Endereço"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtAtividade 
               Height          =   480
               Left            =   120
               TabIndex        =   19
               Tag             =   "Município"
               Top             =   840
               Width           =   3930
               _ExtentX        =   6932
               _ExtentY        =   847
               Caption         =   "Atividade que Execerce"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtNumeroEndAtividade 
               Height          =   480
               Left            =   5325
               TabIndex        =   22
               Tag             =   "Licenciamento"
               Top             =   1380
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   847
               Caption         =   "Número"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtCidadeAtividade 
               Height          =   480
               Left            =   120
               TabIndex        =   24
               Top             =   1920
               Width           =   10275
               _ExtentX        =   18124
               _ExtentY        =   847
               Caption         =   "Cidade"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtBairroAtividade 
               Height          =   480
               Left            =   6120
               TabIndex        =   23
               Top             =   1380
               Width           =   4275
               _ExtentX        =   7541
               _ExtentY        =   847
               Caption         =   "Bairro"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   1725
            Left            =   0
            TabIndex        =   66
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   2640
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   3043
            Altura          =   1905
            Caption         =   " Vinculo Empregatício"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VTOcx.txtVISUAL txtEmpresaNome 
               Height          =   480
               Left            =   2160
               TabIndex        =   28
               Tag             =   "Ano"
               Top             =   960
               Width           =   8220
               _ExtentX        =   14499
               _ExtentY        =   847
               Caption         =   "Empresa Particular"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtEmpresaInicio 
               Height          =   480
               Left            =   120
               TabIndex        =   27
               Top             =   960
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   847
               Caption         =   "Data Início"
               Text            =   ""
               TipoLetras      =   0
               Formato         =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtPublicoNome 
               Height          =   480
               Left            =   2160
               TabIndex        =   26
               Tag             =   "Ano"
               Top             =   360
               Width           =   8220
               _ExtentX        =   14499
               _ExtentY        =   847
               Caption         =   "Orgão Público"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtPublicoInicio 
               Height          =   480
               Left            =   120
               TabIndex        =   25
               Top             =   360
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   847
               Caption         =   "Data Início"
               Text            =   ""
               TipoLetras      =   0
               Formato         =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4560
         Left            =   30
         TabIndex        =   59
         Top             =   360
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   8043
         _Version        =   131082
         TabGuid         =   "CDTR101.frx":0163
         Begin VTOcx.fraVISUAL fraProPrietario 
            Height          =   4845
            Left            =   0
            TabIndex        =   60
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   120
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   8546
            Altura          =   1905
            Caption         =   " Informações Pessoais"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VTOcx.txtVISUAL txtMae 
               Height          =   480
               Left            =   5160
               TabIndex        =   16
               Top             =   3360
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   847
               Caption         =   "Mae"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtPai 
               Height          =   480
               Left            =   120
               TabIndex        =   15
               Top             =   3360
               Width           =   4995
               _ExtentX        =   8811
               _ExtentY        =   847
               Caption         =   "Pai"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtNomePessoa 
               Height          =   480
               Left            =   120
               TabIndex        =   1
               Top             =   360
               Width           =   4995
               _ExtentX        =   8811
               _ExtentY        =   847
               Caption         =   "Nome"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtNascimento 
               Height          =   480
               Left            =   5160
               TabIndex        =   2
               Top             =   360
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   847
               Caption         =   "Data Nascimento"
               Text            =   ""
               TipoLetras      =   0
               Formato         =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtCpf 
               Height          =   480
               Left            =   120
               TabIndex        =   3
               Top             =   840
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   847
               Caption         =   "Cpf"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtIdentidade 
               Height          =   480
               Left            =   2400
               TabIndex        =   4
               Top             =   840
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   847
               Caption         =   "Identidade"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtOrgaoEmissorRg 
               Height          =   480
               Left            =   5160
               TabIndex        =   5
               Top             =   840
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   847
               Caption         =   "Orgão Emissor"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtEnderecoPessoa 
               Height          =   480
               Left            =   120
               TabIndex        =   6
               Top             =   1320
               Width           =   4155
               _ExtentX        =   7329
               _ExtentY        =   847
               Caption         =   "Endereço"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtBairroPessoa 
               Height          =   480
               Left            =   5160
               TabIndex        =   8
               Top             =   1320
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   847
               Caption         =   "Bairro"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtNumeroEndereco 
               Height          =   480
               Left            =   4320
               TabIndex        =   7
               Top             =   1320
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   847
               Caption         =   "Número"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtCidadePessoa 
               Height          =   480
               Left            =   120
               TabIndex        =   9
               Top             =   1800
               Width           =   4995
               _ExtentX        =   8811
               _ExtentY        =   847
               Caption         =   "Cidade"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtEstadoPessoa 
               Height          =   480
               Left            =   5160
               TabIndex        =   10
               Top             =   1800
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   847
               Caption         =   "Estado"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtEmail 
               Height          =   480
               Left            =   120
               TabIndex        =   11
               Top             =   2280
               Width           =   4995
               _ExtentX        =   8811
               _ExtentY        =   847
               Caption         =   "Email"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtTelefone 
               Height          =   480
               Left            =   5160
               TabIndex        =   12
               Top             =   2280
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   847
               Caption         =   "Telefone"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtNaturalidade 
               Height          =   480
               Left            =   120
               TabIndex        =   13
               Top             =   2760
               Width           =   4995
               _ExtentX        =   8811
               _ExtentY        =   847
               Caption         =   "Naturalidade"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               MaxLen          =   100
            End
            Begin VTOcx.txtVISUAL txtEstadoCivil 
               Height          =   480
               Left            =   5160
               TabIndex        =   14
               Top             =   2760
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   847
               Caption         =   "Estado Civil"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               MaxLen          =   100
            End
            Begin VB.Image Image1 
               BorderStyle     =   1  'Fixed Single
               DragMode        =   1  'Automatic
               Height          =   2880
               Left            =   7500
               Stretch         =   -1  'True
               Top             =   360
               Width           =   2760
            End
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   1138
      Icone           =   "CDTR101.frx":018B
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   56
      Top             =   6435
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   7440
         TabIndex        =   52
         Top             =   90
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
         Left            =   9420
         TabIndex        =   54
         Top             =   90
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
         Left            =   8430
         TabIndex        =   53
         Top             =   90
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
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   765
      Left            =   0
      TabIndex        =   57
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1349
      Altura          =   1905
      Caption         =   " Ponto/ Posto"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtNomePosto 
         Height          =   285
         Left            =   45
         TabIndex        =   0
         Top             =   375
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   503
         Caption         =   "Nome"
         Text            =   ""
         CorRotulo       =   4210752
         CorTexto        =   4194304
      End
   End
End
Attribute VB_Name = "CDTR101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cad As DMTransCadastro
Dim cod As String
Dim Rpt As VSRelatorio


Private Sub LimpaCampo()
    
End Sub

Private Sub cmdBuscar_Click()
    inicio
    Set cad = New DMTransCadastro
    If cad.PreencherGrid(grd, txtNomeConsulta) Then
    Else
        Mensagem "Consulta sem resultado"
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    inicio
End Sub

Private Sub cmdLista_Click()
    Set Rpt = New VSRelatorio
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TListaMototaxista.rpt") Then Exit Sub
        .Visualizar
    End With
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub inicio()
    cod = 0
    Image1.Picture = LoadPicture("\\ntserver\Publica\FOTOSDMTRANS\0.jpg")
    Image2.Picture = LoadPicture("\\ntserver\Publica\FOTOSDMTRANS\0.jpg")
    txtCidadeAtividade = "CODO"
    txtCidadeProprietario = "CODO"
    txtEstadoCivil = "SOLTEIRO"
    txtEstadoPessoa = "MA"
    txtEstadoVeiculo = "MA"
    txtMunicipioVeiculo = "CODO"
    txtPropriedade = "PARTICULAR"
End Sub
Private Sub cmdSalvar_Click()
   ' If Not Edita.CriticaCampos(Me) Then Exit Sub
    'Screen.MousePointer = 11
    Set cad = New DMTransCadastro
    If cod = 0 Then
        cod = Imposto.GeraNumNota(1, 102)
    End If
    With cad
            .Codigo = cod
            .POSTO = txtNomePosto
            .CNOME = txtNomePessoa
            .CDATANASCIMENTO = Bdados.Converte(Format(txtNascimento, "dd/mm/yyyy"), TCDataHora)
            .CCPF = ":" & Replace(CStr(txtCpf), ":", "") & ""
            .CIDENTIDADE = txtIdentidade
            .CSSP = txtOrgaoEmissorRg
            .CENDERECO = txtEnderecoPessoa
            .CNUMERO = txtNumeroEndereco
            .CBAIRRO = txtBairroPessoa
            .CCIDADE = txtCidadePessoa
            .CESTADO = txtEstadoPessoa
            .CEMAIL = txtEmail
            .CTELEFONE = txtTelefone
            .CNATURALIDADE = txtNaturalidade
            .CESTADOCIVIL = txtEstadoCivil
            .CPAI = txtPai
            .CMAE = txtMae
            
            .PPROFISSAO = txtProfissao
            .PGRAU = txtGrauEscolar
            .PATIVIDADE = txtAtividade
            .PLOCAL = txtLocalAtividade
            .PCIDADE = txtCidadeAtividade
            .PENDERECO = txtEndAtividade
            .PNUMERO = txtNumeroEndAtividade
            .PBAIRRO = txtBairroAtividade
            .PPUBLICOINICIO = Bdados.Converte(Format(txtPublicoInicio, "dd/mm/yyyy"), TCDataHora)
            .PPUBLICONOME = txtPublicoNome
            .PEMPRESAINICIO = Bdados.Converte(Format(txtEmpresaInicio, "dd/mm/yyyy"), TCDataHora)
            .PEMPRESANOME = txtEmpresaNome
            .CVCNH = txtCnh
            .CVCATEGORIA = txtCatCnh
            .CVDATAHABILITACAO = Bdados.Converte(Format(txtDataPrimHabilitacao, "dd/mm/yyyy"), TCDataHora)
            .CVREGISTRO = txtRegHab
            .CVCURSO = chkCusoEspecializado.Value
            .CVORGAOEMISSORCURSO = txtOrgaoCurso
            .CVORGAOEMISSOR = txtOrgaoEmissaoCnh
            .CVVEICULO = txtVeiculo
            .CVMARCA = txtMarca
            .CVMODELO = txtModelo
            .CVANOFAB = txtAnoFabric
            .CVPLACA = txtPlaca
            .CVCHASSI = txtChassi
            .CVPOTENCIA = txtPotencia
            .CVREGVEICULO = txtRegistroVeiculo
            .CVCIDADEREGISTRO = txtMunicipioVeiculo
            .CVLICENCIAMENTO = txtLicenciamentoVeiculo
            .CVUFLICENCIAMENTO = txtEstadoVeiculo
            .CVPROPRIETARIO = txtProprietarioVeiculo
            .CVENDERECO = txtEndeProprietario
            .CVNUMERO = txtNumeroProprietario
            .CVBAIRRO = txtBairroProprietario
            .CVCIDADEPROPRIETARIO = txtCidadeProprietario
            .CVPROPRIEDADE = txtPropriedade
            
        If .Salvar = True Then
            Avisa "Dados Salvos com Sucesso."
            LimpaCampo
        End If
        Set Rpt = New VSRelatorio
        With Rpt
            If Not .DefinirArquivo(Bdados, App.Path + "\TFichaMototaxista.rpt") Then Exit Sub
            If cod > 0 Then
                .Selecao = "{TAB_BCP_DMTRANS_CADASTRO.CODIGO} = '" & cod & "'"
                .Formulas "DN", Format(txtNascimento, "DD/MM/YYYY")
                .Formulas "DE", Format(txtEmpresaInicio, "DD/MM/YYYY")
                .Formulas "DP", Format(txtPublicoInicio, "DD/MM/YYYY")
                .Formulas "DH", Format(txtDataPrimHabilitacao, "DD/MM/YYYY")
                .Formulas "DM", Format(Now, "DD/MM/YYYY")
                Dim Sql As String
                Dim rs As VSRecordset
                Sql = "SELECT TUS_NOME,TUS_TSE_MATRICULA FROM TAB_USUARIO WHERE TUS_COD_USUARIO = '" & AplicacoesVTFuncoes.Usuario & "'"
                If Bdados.AbreTabela(Sql, rs) Then
                    .Formulas "F", "" & rs(1).Value & " - " & rs(0).Value
                End If
                .Visualizar
            End If
        End With
    End With
    
    'Screen.MousePointer = 0
End Sub

Private Sub cmdVISUAL1_Click()
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     inicio
      ' pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.ico;*.wmf"
     txtPropriedade = "PARTICULAR"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
       
     End If
     
End Sub
Private Sub grd_DblClick()
    On Error GoTo err
    Call cmdLimpar_Click
    cod = grd.ListItems(grd.SelectedItem.Index)
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "select * from tab_bcp_dmtrans_cadastro where codigo = '" & cod & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtAnoFabric = IIf(IsNull(rs!CVANOFAB), "", rs!CVANOFAB)
        txtAtividade = IIf(IsNull(rs!PATIVIDADE), "", rs!PATIVIDADE)
        
        txtBairroAtividade = IIf(IsNull(rs!PBAIRRO), "", rs!PBAIRRO)
        txtBairroPessoa = IIf(IsNull(rs!CBAIRRO), "", rs!CBAIRRO)
        txtBairroProprietario = IIf(IsNull(rs!CVBAIRRO), "", rs!CVBAIRRO)
        txtCatCnh = IIf(IsNull(rs!CVCATEGORIA), "", rs!CVCATEGORIA)
        txtChassi = IIf(IsNull(rs!CVCHASSI), "", rs!CVCHASSI)
        txtCidadeAtividade = IIf(IsNull(rs!PCIDADE), "", rs!PCIDADE)
        txtCidadePessoa = IIf(IsNull(rs!CCIDADE), "", rs!CCIDADE)
        txtCidadeProprietario = IIf(IsNull(rs!CVCIDADEPROPRIETARIO), "", rs!CVCIDADEPROPRIETARIO)
        txtCnh = IIf(IsNull(rs!CVCNH), "", rs!CVCNH)
        txtCpf = IIf(IsNull(rs!CCPF), "", rs!CCPF)
        'txtDataPrimHabilitacao = IIf(IsNull(rs!CVDATAHABILITACAO), "", CDate(rs!CVDATAHABILITACAO))
        txtEmail = IIf(IsNull(rs!CEMAIL), "", rs!CEMAIL)
        If IsNull(rs!CVDATAHABILITACAO) Then
            txtDataPrimHabilitacao = ""
        Else
            txtDataPrimHabilitacao = CDate(rs!CVDATAHABILITACAO)
        End If
        If IsNull(rs!PEMPRESAINICIO) Then
            txtEmpresaInicio = ""
        Else
            txtEmpresaInicio = CDate(rs!PEMPRESAINICIO)
        End If
        txtEmpresaNome = IIf(IsNull(rs!PEMPRESANOME), "", rs!PEMPRESANOME)
        txtEndAtividade = IIf(IsNull(rs!PENDERECO), "", rs!PENDERECO)
        txtEndeProprietario = IIf(IsNull(rs!CVENDERECO), "", rs!CVENDERECO)
        txtEnderecoPessoa = IIf(IsNull(rs!CENDERECO), "", rs!CENDERECO)
        txtEstadoCivil = IIf(IsNull(rs!CESTADOCIVIL), "", rs!CESTADOCIVIL)
        txtEstadoPessoa = IIf(IsNull(rs!CESTADO), "", rs!CESTADO)
        txtEstadoVeiculo = IIf(IsNull(rs!CVUFLICENCIAMENTO), "", rs!CVUFLICENCIAMENTO)
        txtGrauEscolar = IIf(IsNull(rs!pgrauescolar), "", rs!pgrauescolar)
        txtIdentidade = IIf(IsNull(rs!CIDENTIDADE), "", rs!CIDENTIDADE)
        txtLicenciamentoVeiculo = IIf(IsNull(rs!CVLICENCIAMENTO), "", rs!CVLICENCIAMENTO)
        txtLocalAtividade = IIf(IsNull(rs!PLOCAL), "", rs!PLOCAL)
        txtMae = IIf(IsNull(rs!CMAE), "", rs!CMAE)
        txtMarca = IIf(IsNull(rs!CVMARCA), "", rs!CVMARCA)
        txtModelo = IIf(IsNull(rs!CVMODELO), "", rs!CVMODELO)
        txtMunicipioVeiculo = IIf(IsNull(rs!CVCIDADEREGISTRO), "", rs!CVCIDADEREGISTRO)
        txtNascimento = IIf(IsNull(rs!CDATANASCIMENTO), "", CDate(rs!CDATANASCIMENTO))
        txtNaturalidade = IIf(IsNull(rs!CNATURALIDADE), "", rs!CNATURALIDADE)
        txtNomePessoa = IIf(IsNull(rs!CNOME), "", rs!CNOME)
        txtNomePosto = IIf(IsNull(rs!POSTO), "", rs!POSTO)
        txtNumeroEndAtividade = IIf(IsNull(rs!PNUMERO), "", rs!PNUMERO)
        txtNumeroEndereco = IIf(IsNull(rs!CNUMERO), "", rs!CNUMERO)
        txtNumeroProprietario = IIf(IsNull(rs!CVNUMERO), "", rs!CVNUMERO)
        txtOrgaoCurso = IIf(IsNull(rs!CVORGAOEMISSORCURSO), "", rs!CVORGAOEMISSORCURSO)
        txtOrgaoEmissaoCnh = IIf(IsNull(rs!CVORGAOEMISSOR), "", rs!CVORGAOEMISSOR)
        txtOrgaoEmissorRg = IIf(IsNull(rs!CSSP), "", rs!CSSP)
        txtPai = IIf(IsNull(rs!CPAI), "", rs!CPAI)
        txtPlaca = IIf(IsNull(rs!CVPLACA), "", rs!CVPLACA)
        txtPotencia = IIf(IsNull(rs!CVPOTENCIA), "", rs!CVPOTENCIA)
        txtProfissao = IIf(IsNull(rs!PPROFISSAO), "", rs!PPROFISSAO)
        txtPropriedade = IIf(IsNull(rs!CVPROPRIEDADE), "", rs!CVPROPRIEDADE)
        txtProprietarioVeiculo = IIf(IsNull(rs!CVPROPRIETARIO), "", rs!CVPROPRIETARIO)
        
        If IsNull(rs!PPUBLICOINICIO) Then
            txtPublicoInicio = ""
        Else
            txtPublicoInicio = CDate(rs!PPUBLICOINICIO)
        End If
        txtPublicoNome = IIf(IsNull(rs!PPUBLICONOME), "", rs!PPUBLICONOME)
        txtRegHab = IIf(IsNull(rs!CVREGISTRO), "", rs!CVREGISTRO)
        txtRegistroVeiculo = IIf(IsNull(rs!CVREGVEICULO), "", rs!CVREGVEICULO)
        txtTelefone = IIf(IsNull(rs!CTELEFONE), "", rs!CTELEFONE)
        txtVeiculo = IIf(IsNull(rs!CVVEICULO), "", rs!CVVEICULO)
        
        Image1.Picture = LoadPicture("\\ntserver\Publica\FOTOSDMTRANS\" & CInt(Right(cod, 5)) & ".jpg")
        Image2.Picture = LoadPicture("\\ntserver\Publica\FOTOSDMTRANS\" & CInt(Right(cod, 5)) & ".jpg")
        
    End If
err:
    'Mensagem "Cadastro sem foto"
End Sub
