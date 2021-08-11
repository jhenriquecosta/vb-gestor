VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form REES107 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " REES107"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   7605
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   820
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   5925
         TabIndex        =   9
         Top             =   105
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   4815
         TabIndex        =   8
         Top             =   105
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         Caption         =   "&Buscar"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   7020
         TabIndex        =   3
         Top             =   105
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   8115
         TabIndex        =   2
         Top             =   105
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1020
      Left            =   30
      TabIndex        =   4
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   645
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   1799
      Altura          =   1905
      Caption         =   " Dados do Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2760
         TabIndex        =   7
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
         Left            =   3165
         TabIndex        =   6
         Top             =   375
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   503
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Tag             =   "Insc. Municipal"
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   285
         Left            =   450
         TabIndex        =   5
         Top             =   690
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   503
         Caption         =   "Endere�o"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin ActiveTabs.SSActiveTabs tabRegime 
      Height          =   4425
      Left            =   30
      TabIndex        =   10
      Tag             =   "Documento gerencial"
      Top             =   3180
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   7805
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      TagVariant      =   ""
      Tabs            =   "REES107.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel11 
         Height          =   4020
         Left            =   -99969
         TabIndex        =   11
         Top             =   30
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7091
         _Version        =   131082
         TabGuid         =   "REES107.frx":00F1
         Begin VTOcx.fraVISUAL fraVISUAL11 
            Height          =   1605
            Left            =   15
            TabIndex        =   12
            Top             =   3045
            Width           =   9045
            _ExtentX        =   15954
            _ExtentY        =   2831
            Altura          =   1905
            Caption         =   " Autoridade Fiscal"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtrespConclu 
               Height          =   480
               Left            =   45
               TabIndex        =   15
               Tag             =   "Respons�vel"
               Top             =   315
               Width           =   4980
               _ExtentX        =   8784
               _ExtentY        =   847
               Caption         =   "Respons�vel"
               Text            =   ""
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   50
            End
            Begin VTOcx.txtVISUAL txtMatrRespiConclui 
               Height          =   480
               Left            =   5025
               TabIndex        =   14
               Tag             =   "Matr�cula"
               Top             =   315
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   847
               Caption         =   "Matr�cula"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   15
            End
            Begin VTOcx.txtVISUAL txtDataConcl 
               Height          =   480
               Left            =   7065
               TabIndex        =   13
               Tag             =   "Data"
               Top             =   315
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   847
               Caption         =   "Data"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   15
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL12 
            Height          =   2865
            Left            =   0
            TabIndex        =   16
            Top             =   30
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   5054
            Altura          =   1905
            Caption         =   " Conclus�o"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VB.TextBox txtConclusao 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   2040
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               Top             =   810
               Width           =   9060
            End
            Begin VTOcx.cboVISUAL cboStatus 
               Height          =   315
               Left            =   60
               TabIndex        =   17
               Tag             =   "UF"
               Top             =   390
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorRotulo       =   4210752
               CorTexto        =   4194304
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4035
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7117
         _Version        =   131082
         TabGuid         =   "REES107.frx":0119
         Begin ActiveTabs.SSActiveTabs tabInstrucao 
            Height          =   4050
            Left            =   -30
            TabIndex        =   20
            Tag             =   "Documento gerencial"
            Top             =   0
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   7144
            _Version        =   131082
            TabCount        =   2
            TagVariant      =   ""
            Tabs            =   "REES107.frx":0141
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel10 
               Height          =   3660
               Left            =   -99969
               TabIndex        =   21
               Top             =   360
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   6456
               _Version        =   131082
               TabGuid         =   "REES107.frx":01DF
               Begin VTOcx.fraVISUAL fraVISUAL7 
                  Height          =   3405
                  Left            =   0
                  TabIndex        =   22
                  Top             =   0
                  Width           =   9120
                  _ExtentX        =   16087
                  _ExtentY        =   6006
                  Altura          =   1905
                  Caption         =   " Despacho"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Borda           =   0
                  Begin VTOcx.fraVISUAL fraVISUAL10 
                     Height          =   855
                     Left            =   60
                     TabIndex        =   24
                     Top             =   2535
                     Width           =   9045
                     _ExtentX        =   15954
                     _ExtentY        =   1508
                     Altura          =   1905
                     Caption         =   " Autoridade Fiscal"
                     CorTexto        =   16777215
                     CorFaixa        =   32768
                     CorFundo        =   -2147483633
                     Ocultavel       =   0   'False
                     Begin VTOcx.txtVISUAL txtDataCont 
                        Height          =   480
                        Left            =   7110
                        TabIndex        =   27
                        Tag             =   "Data"
                        Top             =   315
                        Width           =   1890
                        _ExtentX        =   3334
                        _ExtentY        =   847
                        Caption         =   "Data"
                        Text            =   ""
                        Enabled         =   0   'False
                        Formato         =   0
                        Restricao       =   2
                        AlinhamentoRotulo=   1
                        CorRotulo       =   4210752
                        CorTexto        =   4194304
                        MaxLen          =   15
                     End
                     Begin VTOcx.txtVISUAL txtRespCont 
                        Height          =   480
                        Left            =   90
                        TabIndex        =   26
                        Tag             =   "Respons�vel"
                        Top             =   315
                        Width           =   4980
                        _ExtentX        =   8784
                        _ExtentY        =   847
                        Caption         =   "Respons�vel"
                        Text            =   ""
                        Enabled         =   0   'False
                        AlinhamentoRotulo=   1
                        CorRotulo       =   4210752
                        CorTexto        =   4194304
                        MaxLen          =   50
                     End
                     Begin VTOcx.txtVISUAL txtCPFCont 
                        Height          =   480
                        Left            =   5070
                        TabIndex        =   25
                        Tag             =   "Matr�cula"
                        Top             =   315
                        Width           =   2040
                        _ExtentX        =   3598
                        _ExtentY        =   847
                        Caption         =   "CPF"
                        Text            =   ""
                        Enabled         =   0   'False
                        Formato         =   1
                        Restricao       =   2
                        AlinhamentoRotulo=   1
                        CorRotulo       =   4210752
                        CorTexto        =   4194304
                        MaxLen          =   15
                     End
                  End
                  Begin VB.TextBox txtDespCont 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   2205
                     Left            =   30
                     MaxLength       =   4000
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   23
                     Tag             =   "Declara��o Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel9 
               Height          =   3510
               Left            =   -99969
               TabIndex        =   28
               Top             =   360
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   6191
               _Version        =   131082
               TabGuid         =   "REES107.frx":0207
               Begin VTOcx.fraVISUAL fraVISUAL6 
                  Height          =   3285
                  Left            =   0
                  TabIndex        =   29
                  Top             =   0
                  Width           =   9120
                  _ExtentX        =   16087
                  _ExtentY        =   5794
                  Altura          =   1905
                  Caption         =   " Declara��o Fiscal"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Borda           =   0
                  Begin VB.TextBox Text2 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   2940
                     Left            =   30
                     MaxLength       =   4000
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   30
                     Tag             =   "Declara��o Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
               Height          =   3660
               Left            =   30
               TabIndex        =   31
               Top             =   360
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   6456
               _Version        =   131082
               TabGuid         =   "REES107.frx":022F
               Begin VTOcx.fraVISUAL fraVISUAL5 
                  Height          =   3405
                  Left            =   0
                  TabIndex        =   32
                  Top             =   -30
                  Width           =   9120
                  _ExtentX        =   16087
                  _ExtentY        =   6006
                  Altura          =   1905
                  Caption         =   " Despacho"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Borda           =   0
                  Begin VTOcx.fraVISUAL fraVISUAL9 
                     Height          =   855
                     Left            =   45
                     TabIndex        =   34
                     Top             =   2535
                     Width           =   9045
                     _ExtentX        =   15954
                     _ExtentY        =   1508
                     Altura          =   1905
                     Caption         =   " Autoridade Fiscal"
                     CorTexto        =   16777215
                     CorFaixa        =   32768
                     CorFundo        =   -2147483633
                     Ocultavel       =   0   'False
                     Begin VTOcx.txtVISUAL txtDataAutor 
                        Height          =   480
                        Left            =   7125
                        TabIndex        =   37
                        Tag             =   "Data"
                        Top             =   300
                        Width           =   1890
                        _ExtentX        =   3334
                        _ExtentY        =   847
                        Caption         =   "Data"
                        Text            =   ""
                        Enabled         =   0   'False
                        Formato         =   0
                        Restricao       =   2
                        AlinhamentoRotulo=   1
                        CorRotulo       =   4210752
                        CorTexto        =   4194304
                        MaxLen          =   15
                     End
                     Begin VTOcx.txtVISUAL txtRespAutor 
                        Height          =   480
                        Left            =   105
                        TabIndex        =   36
                        Tag             =   "Respons�vel"
                        Top             =   300
                        Width           =   4980
                        _ExtentX        =   8784
                        _ExtentY        =   847
                        Caption         =   "Respons�vel"
                        Text            =   ""
                        Enabled         =   0   'False
                        AlinhamentoRotulo=   1
                        CorRotulo       =   4210752
                        CorTexto        =   4194304
                        MaxLen          =   50
                     End
                     Begin VTOcx.txtVISUAL txtMatricula 
                        Height          =   480
                        Left            =   5085
                        TabIndex        =   35
                        Tag             =   "Matr�cula"
                        Top             =   300
                        Width           =   2040
                        _ExtentX        =   3598
                        _ExtentY        =   847
                        Caption         =   "Matr�cula"
                        Text            =   ""
                        Enabled         =   0   'False
                        Restricao       =   2
                        AlinhamentoRotulo=   1
                        CorRotulo       =   4210752
                        CorTexto        =   4194304
                        MaxLen          =   15
                     End
                  End
                  Begin VB.TextBox txtDespAutor 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   2205
                     Left            =   30
                     MaxLength       =   4000
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   33
                     Tag             =   "Declara��o Fiscal"
                     Top             =   315
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel8 
               Height          =   3270
               Left            =   -99969
               TabIndex        =   38
               Top             =   30
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   5768
               _Version        =   131082
               TabGuid         =   "REES107.frx":0257
               Begin VTOcx.grdVISUAL grdVISUAL1 
                  Height          =   3180
                  Left            =   15
                  TabIndex        =   39
                  Top             =   90
                  Width           =   9105
                  _ExtentX        =   16060
                  _ExtentY        =   5609
                  CorBorda        =   32768
                  Caption         =   "Processos em Andamento"
                  CorTitulo       =   32768
                  CorCaption      =   16777215
                  CorDica         =   32768
                  CheckBox        =   -1  'True
                  MarcaUnico      =   -1  'True
               End
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4020
         Left            =   -99969
         TabIndex        =   40
         Top             =   30
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7091
         _Version        =   131082
         TabGuid         =   "REES107.frx":027F
         Begin ActiveTabs.SSActiveTabs tabInstauracao 
            Height          =   2985
            Left            =   -15
            TabIndex        =   41
            Tag             =   "Documento gerencial"
            Top             =   -15
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   5265
            _Version        =   131082
            TabCount        =   4
            TagVariant      =   ""
            Tabs            =   "REES107.frx":02A7
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
               Height          =   2595
               Left            =   -99969
               TabIndex        =   42
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES107.frx":03A6
               Begin VTOcx.fraVISUAL fraVISUAL1 
                  Height          =   2595
                  Left            =   0
                  TabIndex        =   43
                  Top             =   0
                  Width           =   9120
                  _ExtentX        =   16087
                  _ExtentY        =   4577
                  Altura          =   1905
                  Caption         =   " Nota Fiscal"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Borda           =   0
                  Begin VB.TextBox txtNota 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   2295
                     Left            =   30
                     MaxLength       =   4000
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   44
                     Tag             =   "Nota Fiscal"
                     Top             =   285
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
               Height          =   2595
               Left            =   30
               TabIndex        =   45
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES107.frx":03CE
               Begin VTOcx.fraVISUAL fraVISUAL2 
                  Height          =   2595
                  Left            =   15
                  TabIndex        =   46
                  Top             =   0
                  Width           =   9120
                  _ExtentX        =   16087
                  _ExtentY        =   4577
                  Altura          =   1905
                  Caption         =   " Livro Fiscal (Modelos Diferentes)"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Borda           =   0
                  Begin VB.TextBox txtLivro 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   2295
                     Left            =   30
                     MaxLength       =   4000
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   47
                     Tag             =   "Livro Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
               Height          =   2595
               Left            =   -99969
               TabIndex        =   48
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES107.frx":03F6
               Begin VTOcx.fraVISUAL fraVISUAL3 
                  Height          =   2625
                  Left            =   -15
                  TabIndex        =   49
                  Top             =   -30
                  Width           =   9120
                  _ExtentX        =   16087
                  _ExtentY        =   4630
                  Altura          =   1905
                  Caption         =   " Declara��o Fiscal"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Borda           =   0
                  Begin VB.TextBox txtDeclaracao 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   2325
                     Left            =   30
                     MaxLength       =   4000
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   50
                     Tag             =   "Declara��o Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
               Height          =   2595
               Left            =   -99969
               TabIndex        =   51
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES107.frx":041E
               Begin VTOcx.fraVISUAL fraVISUAL4 
                  Height          =   3285
                  Left            =   0
                  TabIndex        =   52
                  Top             =   0
                  Width           =   9120
                  _ExtentX        =   16087
                  _ExtentY        =   5794
                  Altura          =   1905
                  Caption         =   " Documento Gerencial"
                  CorTexto        =   16777215
                  CorFaixa        =   32768
                  CorFundo        =   -2147483633
                  Ocultavel       =   0   'False
                  Borda           =   0
                  Begin VB.TextBox txtDocumento 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00800000&
                     Height          =   2295
                     Left            =   30
                     MaxLength       =   4000
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   53
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL8 
            Height          =   855
            Left            =   30
            TabIndex        =   54
            Top             =   3000
            Width           =   9060
            _ExtentX        =   15981
            _ExtentY        =   1508
            Altura          =   1905
            Caption         =   " Dados do Respons�vel"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtResp 
               Height          =   480
               Left            =   75
               TabIndex        =   56
               Tag             =   "Respons�vel"
               Top             =   315
               Width           =   5850
               _ExtentX        =   10319
               _ExtentY        =   847
               Caption         =   "Respons�vel:"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   50
            End
            Begin VTOcx.txtVISUAL txtCPF 
               Height          =   480
               Left            =   5925
               TabIndex        =   55
               Tag             =   "CPF Respons�vel"
               Top             =   315
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   847
               Caption         =   "CPF"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   1
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               MaxLen          =   15
            End
         End
      End
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   1710
      Left            =   15
      TabIndex        =   57
      Top             =   1695
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   3016
      CorBorda        =   32768
      Caption         =   "Processos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1138
      Icone           =   "REES107.frx":0446
   End
   Begin VTOcx.txtVISUAL txtCodigo 
      Height          =   480
      Left            =   1350
      TabIndex        =   59
      Tag             =   "CPF Respons�vel"
      Top             =   1020
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   847
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
      Restricao       =   2
      AlinhamentoRotulo=   1
      CorRotulo       =   4210752
      MaxLen          =   15
   End
End
Attribute VB_Name = "REES107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    If txtIm <> "" Then
        carregaProcesso txtIm
    Else
        carregaProcesso
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdDados.ListItems.Clear
    txtDataConcl = Date
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub



Private Sub cmdSalvar_Click()
     ' status = 1 - aberto TipoProcesso Regime Especial = 3
    Dim camposPr As String
    Dim ValoresPr As String
    Dim Condicao As String
   
   
    If txtCodigo = "" Then
        Avisa "Selecione um processo"
        Exit Sub
    End If
    
     If cboStatus.ListIndex = -1 Then
        Avisa "Status deve ser informado"
        Exit Sub
    End If
     If txtConclusao = "" Then
        Avisa "Conclus�o deve ser informado"
        Exit Sub
    End If
     If txtrespConclu = "" Then
        Avisa "Respons�vel deve ser informado"
        Exit Sub
    End If
       If txtMatrRespiConclui = "" Then
        Avisa "Matricula deve ser informado"
        Exit Sub
    End If
       If txtDataConcl = "" Then
        Avisa "Data deve ser informado"
        Exit Sub
    End If
    Condicao = " TRE_TPR_NUMERO_PROCESSO = '" & txtCodigo & "'"
    camposPr = " TRE_STATUS_CONCLUSAO,TRE_FUNCIONARIO_CONCLUSAO,TRE_MATRICULA_FUNC_CONCLUSAO,TRE_DESCRICAO_CONCLUSAO,TRE_DATA_CONCLUSAO "
    ValoresPr = Bdados.PreparaValor(cboStatus.Coluna(1).VALOR, txtrespConclu, txtMatrRespiConclui, txtConclusao, txtDataConcl)
    If Bdados.AtualizaDados("TAB_REGIME_ESPECIAL", ValoresPr, camposPr, Condicao) Then
        AteraStatusProcesso
        limparcampos
        carregaProcesso
    End If
  
End Sub

Private Sub AteraStatusProcesso()
     ' status = 1 - aberto TipoProcesso Regime Especial = 3
    Dim camposPr As String
    Dim ValoresPr As String
    Dim Condicao As String
    Dim status As Integer
    
    status = 2
    
    Condicao = " TPR_NUMERO_PROCESSO = '" & txtCodigo & "'"
    camposPr = " TPR_STATUS"
    ValoresPr = Bdados.PreparaValor(status)
    If Bdados.AtualizaDados("TAB_PROCESSO", ValoresPr, camposPr, Condicao) Then
        Avisa "Dados Salvos com Sucesso"
    End If
  
End Sub


Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
     txtDataConcl = Date
    cboStatus.Preencher Bdados, "select * from vis_status_regime_especial where TGE_CODIGO<>3"
End Sub
Private Sub limparcampos()
        txtCodigo = ""
        txtLivro = ""
        txtNota = ""
        txtDeclaracao = ""
        txtDocumento = ""
        txtResp = ""
        txtCPF = ""
        txtDespAutor = ""
        txtRespAutor = ""
        txtMatricula = ""
        txtDataAutor = ""
        txtDespCont = ""
        txtRespCont = ""
        txtCPFCont = ""
        txtDataCont = ""
        cboStatus.ListIndex = -1
        txtConclusao = ""
        txtrespConclu = ""
        txtMatrRespiConclui = ""
        txtDataConcl = ""
        txtDataConcl = Date
End Sub
Private Sub grdDados_dblClick()
    If grdDados.ListItems.Count >= 1 Then
        txtCodigo = grdDados.SelectedItem
        txtLivro = grdDados.SelectedItem.SubItems(4)
        txtNota = grdDados.SelectedItem.SubItems(5)
        txtDeclaracao = grdDados.SelectedItem.SubItems(6)
        txtDocumento = grdDados.SelectedItem.SubItems(7)
        txtResp = grdDados.SelectedItem.SubItems(8)
        txtCPF = grdDados.SelectedItem.SubItems(9)
        txtDespAutor = grdDados.SelectedItem.SubItems(10)
        txtRespAutor = grdDados.SelectedItem.SubItems(11)
        txtMatricula = grdDados.SelectedItem.SubItems(12)
        txtDataAutor = grdDados.SelectedItem.SubItems(13)
        txtDespCont = grdDados.SelectedItem.SubItems(14)
        txtRespCont = grdDados.SelectedItem.SubItems(15)
        txtCPFCont = grdDados.SelectedItem.SubItems(16)
        txtDataCont = grdDados.SelectedItem.SubItems(17)
    End If
End Sub

Private Sub carregaProcesso(Optional Im As String)
    Dim sql As String
    
     sql = "select TPR_NUMERO_PROCESSO as Processo, "
     sql = sql & " TPR_INSCRICAO as Inscri��o,"
     sql = sql & " TPR_DESCRICAO_PEDIDO as Descri��o,"
     sql = sql & " TGE_NOME   As Status,"
     sql = sql & " TRE_LIVROS_FISCAIS_MODELOS ,"
     sql = sql & " TRE_DESCRICAO_NOTA_FISCAL ,"
     sql = sql & " TRE_DESCRICAO_DECLARACAO,"
     sql = sql & " TRE_DESCRICAO_DOCUMENTO_FISCAL,"
     sql = sql & " TPR_PEDIDO_REPR_PREPOSTO,"
     sql = sql & " TPR_PEDIDO_REPR_PREPOSTO_CPF,"
     sql = sql & " TPR_FUNCIONARIO_DESPACHO ,"
     sql = sql & " TPR_FUNCIONARIO_NOME,"
     sql = sql & " TPR_FUNCIONARIO_MATRICULA,"
     sql = sql & " TPR_FUNCIONARIO_DATA_VISTO,"
     sql = sql & " TPR_INSTRUCAO_PASSIVO_DESPACHO,"
     sql = sql & " TPR_INST_PASSIVO_REPR_PREPOSTO,"
     sql = sql & " TPR_INST_PAS_REPR_PREPOSTO_CPF,"
     sql = sql & " TPR_INSTRUCAO_PASSIVO_DATA"
     sql = sql & " From tab_processo,"
     sql = sql & "  vis_status_Processo ,"
     sql = sql & " tab_regime_especial"
     sql = sql & " Where TPR_TIPO_PROCESSO = 3"
     sql = sql & " And TPR_STATUS = 1"
     sql = sql & " And TPR_STATUS = TGE_CODIGO"
     sql = sql & " And TRE_STATUS_CONCLUSAO = 3"
     sql = sql & " and TRE_TPR_NUMERO_PROCESSO =   TPR_NUMERO_PROCESSO "
    
  
    If Im <> "" Then sql = sql & " and TPR_INSCRICAO = '" & Im & "'"
    
    sql = sql & " order by TRE_TPR_NUMERO_PROCESSO"
  
    If Not grdDados.Preencher(Bdados, sql, 1200, 1200, 5500, 1500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) Then
        Avisa "Busca sem resultados"
    End If
    

End Sub
