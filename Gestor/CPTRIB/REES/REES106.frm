VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form REES106 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " REES106"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   7635
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   820
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   5925
         TabIndex        =   10
         Top             =   90
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
         TabIndex        =   4
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   330
         Left            =   4815
         TabIndex        =   3
         Top             =   90
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         Caption         =   "&Imprimir"
         Acao            =   4
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
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1138
      Icone           =   "REES106.frx":0000
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1020
      Left            =   30
      TabIndex        =   6
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   690
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   503
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin ActiveTabs.SSActiveTabs tabRegime 
      Height          =   4410
      Left            =   45
      TabIndex        =   11
      Tag             =   "Documento gerencial"
      Top             =   3225
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   7779
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      Tabs            =   "REES106.frx":031A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel11 
         Height          =   4020
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7091
         _Version        =   131082
         TabGuid         =   "REES106.frx":040B
         Begin VTOcx.fraVISUAL fraVISUAL11 
            Height          =   1605
            Left            =   15
            TabIndex        =   13
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
            Begin VTOcx.txtVISUAL txtDataConcl 
               Height          =   480
               Left            =   7065
               TabIndex        =   16
               Tag             =   "Data"
               Top             =   315
               Width           =   1950
               _ExtentX        =   3440
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
            Begin VTOcx.txtVISUAL txtMatrRespiConclui 
               Height          =   480
               Left            =   5025
               TabIndex        =   15
               Tag             =   "Matrícula"
               Top             =   315
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   847
               Caption         =   "Matrícula"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   15
            End
            Begin VTOcx.txtVISUAL txtrespConclu 
               Height          =   480
               Left            =   45
               TabIndex        =   14
               Tag             =   "Responsável"
               Top             =   315
               Width           =   4980
               _ExtentX        =   8784
               _ExtentY        =   847
               Caption         =   "Responsável"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   50
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL12 
            Height          =   2865
            Left            =   0
            TabIndex        =   17
            Top             =   30
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   5054
            Altura          =   1905
            Caption         =   " Conclusão"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VTOcx.cboVISUAL cboStatus 
               Height          =   315
               Left            =   60
               TabIndex        =   19
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
               Enabled         =   0   'False
            End
            Begin VB.TextBox txtConclusao 
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
               Height          =   2040
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               Top             =   825
               Width           =   9060
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4020
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7091
         _Version        =   131082
         TabGuid         =   "REES106.frx":0433
         Begin ActiveTabs.SSActiveTabs tabInstrucao 
            Height          =   4050
            Left            =   -30
            TabIndex        =   21
            Tag             =   "Documento gerencial"
            Top             =   0
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   7144
            _Version        =   131082
            TabCount        =   2
            Tabs            =   "REES106.frx":045B
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel10 
               Height          =   3660
               Left            =   30
               TabIndex        =   22
               Top             =   360
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   6456
               _Version        =   131082
               TabGuid         =   "REES106.frx":04F9
               Begin VTOcx.fraVISUAL fraVISUAL7 
                  Height          =   3405
                  Left            =   0
                  TabIndex        =   23
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
                     TabIndex        =   28
                     Tag             =   "Declaração Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
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
                     Begin VTOcx.txtVISUAL txtCPFCont 
                        Height          =   480
                        Left            =   5070
                        TabIndex        =   27
                        Tag             =   "Matrícula"
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
                     Begin VTOcx.txtVISUAL txtRespCont 
                        Height          =   480
                        Left            =   90
                        TabIndex        =   26
                        Tag             =   "Responsável"
                        Top             =   315
                        Width           =   4980
                        _ExtentX        =   8784
                        _ExtentY        =   847
                        Caption         =   "Responsável"
                        Text            =   ""
                        Enabled         =   0   'False
                        AlinhamentoRotulo=   1
                        CorRotulo       =   4210752
                        CorTexto        =   4194304
                        MaxLen          =   50
                     End
                     Begin VTOcx.txtVISUAL txtDataCont 
                        Height          =   480
                        Left            =   7110
                        TabIndex        =   25
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
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel9 
               Height          =   3510
               Left            =   -99969
               TabIndex        =   29
               Top             =   360
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   6191
               _Version        =   131082
               TabGuid         =   "REES106.frx":0521
               Begin VTOcx.fraVISUAL fraVISUAL6 
                  Height          =   3285
                  Left            =   0
                  TabIndex        =   30
                  Top             =   0
                  Width           =   9120
                  _ExtentX        =   16087
                  _ExtentY        =   5794
                  Altura          =   1905
                  Caption         =   " Declaração Fiscal"
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
                     TabIndex        =   31
                     Tag             =   "Declaração Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
               Height          =   3660
               Left            =   30
               TabIndex        =   32
               Top             =   360
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   6456
               _Version        =   131082
               TabGuid         =   "REES106.frx":0549
               Begin VTOcx.fraVISUAL fraVISUAL5 
                  Height          =   3405
                  Left            =   0
                  TabIndex        =   33
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
                     TabIndex        =   38
                     Tag             =   "Declaração Fiscal"
                     Top             =   315
                     Width           =   9060
                  End
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
                     Begin VTOcx.txtVISUAL txtMatricula 
                        Height          =   480
                        Left            =   5085
                        TabIndex        =   37
                        Tag             =   "Matrícula"
                        Top             =   315
                        Width           =   2040
                        _ExtentX        =   3598
                        _ExtentY        =   847
                        Caption         =   "Matrícula"
                        Text            =   ""
                        Enabled         =   0   'False
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
                        Tag             =   "Responsável"
                        Top             =   300
                        Width           =   4980
                        _ExtentX        =   8784
                        _ExtentY        =   847
                        Caption         =   "Responsável"
                        Text            =   ""
                        Enabled         =   0   'False
                        AlinhamentoRotulo=   1
                        CorRotulo       =   4210752
                        CorTexto        =   4194304
                        MaxLen          =   50
                     End
                     Begin VTOcx.txtVISUAL txtDataAutor 
                        Height          =   480
                        Left            =   7125
                        TabIndex        =   35
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
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel8 
               Height          =   3270
               Left            =   -99969
               TabIndex        =   39
               Top             =   30
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   5768
               _Version        =   131082
               TabGuid         =   "REES106.frx":0571
               Begin VTOcx.grdVISUAL grdVISUAL1 
                  Height          =   3180
                  Left            =   15
                  TabIndex        =   40
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
         Left            =   30
         TabIndex        =   41
         Top             =   30
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7091
         _Version        =   131082
         TabGuid         =   "REES106.frx":0599
         Begin ActiveTabs.SSActiveTabs tabInstauracao 
            Height          =   2985
            Left            =   -45
            TabIndex        =   42
            Tag             =   "Documento gerencial"
            Top             =   -15
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   5265
            _Version        =   131082
            TabCount        =   4
            Tabs            =   "REES106.frx":05C1
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
               Height          =   2595
               Left            =   30
               TabIndex        =   43
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES106.frx":06C0
               Begin VTOcx.fraVISUAL fraVISUAL1 
                  Height          =   2595
                  Left            =   0
                  TabIndex        =   44
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
                     TabIndex        =   45
                     Tag             =   "Nota Fiscal"
                     Top             =   285
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
               Height          =   2595
               Left            =   30
               TabIndex        =   46
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES106.frx":06E8
               Begin VTOcx.fraVISUAL fraVISUAL2 
                  Height          =   2595
                  Left            =   15
                  TabIndex        =   47
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
                     TabIndex        =   48
                     Tag             =   "Livro Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
               Height          =   2595
               Left            =   30
               TabIndex        =   49
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES106.frx":0710
               Begin VTOcx.fraVISUAL fraVISUAL3 
                  Height          =   2625
                  Left            =   -15
                  TabIndex        =   50
                  Top             =   -30
                  Width           =   9120
                  _ExtentX        =   16087
                  _ExtentY        =   4630
                  Altura          =   1905
                  Caption         =   " Declaração Fiscal"
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
                     TabIndex        =   51
                     Tag             =   "Declaração Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
               Height          =   2595
               Left            =   30
               TabIndex        =   52
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES106.frx":0738
               Begin VTOcx.fraVISUAL fraVISUAL4 
                  Height          =   3285
                  Left            =   0
                  TabIndex        =   53
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
                     TabIndex        =   54
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL8 
            Height          =   855
            Left            =   30
            TabIndex        =   55
            Top             =   3000
            Width           =   9060
            _ExtentX        =   15981
            _ExtentY        =   1508
            Altura          =   1905
            Caption         =   " Dados do Responsável"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtCPF 
               Height          =   480
               Left            =   5925
               TabIndex        =   57
               Tag             =   "CPF Responsável"
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
            Begin VTOcx.txtVISUAL txtResp 
               Height          =   480
               Left            =   75
               TabIndex        =   56
               Tag             =   "Responsável"
               Top             =   315
               Width           =   5850
               _ExtentX        =   10319
               _ExtentY        =   847
               Caption         =   "Responsável:"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   50
            End
         End
      End
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   1695
      Left            =   15
      TabIndex        =   58
      Top             =   1695
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   2990
      CorBorda        =   32768
      Caption         =   "Processos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
End
Attribute VB_Name = "REES106"
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
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
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
         cboStatus.Preencher Bdados, "select * from vis_status_regime_especial "
End Sub

Private Sub grdDados_dblClick()
    If grdDados.ListItems.Count >= 1 Then
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
        cboStatus.SetarLinha grdDados.SelectedItem.SubItems(18), 1
        txtrespConclu = grdDados.SelectedItem.SubItems(19)
        txtMatrRespiConclui = grdDados.SelectedItem.SubItems(20)
        txtConclusao = grdDados.SelectedItem.SubItems(21)
        txtDataConcl = grdDados.SelectedItem.SubItems(22)
    End If
End Sub

Private Sub carregaProcesso(Optional Im As String)
    Dim sql As String
    
     sql = "select TPR_NUMERO_PROCESSO as Processo, "
     sql = sql & " TPR_INSCRICAO as Inscrição,"
     sql = sql & " TPR_DESCRICAO_PEDIDO as Descrição,"
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
     sql = sql & " TPR_INSTRUCAO_PASSIVO_DATA,"
     sql = sql & " TRE_STATUS_CONCLUSAO,"
     sql = sql & " TRE_FUNCIONARIO_CONCLUSAO,"
     sql = sql & " TRE_MATRICULA_FUNC_CONCLUSAO,"
     sql = sql & " TRE_DESCRICAO_CONCLUSAO,"
     sql = sql & " TRE_DATA_CONCLUSAO"
     sql = sql & " From tab_processo,"
     sql = sql & "  vis_status_Processo ,"
     sql = sql & " tab_regime_especial"
     sql = sql & " Where TPR_TIPO_PROCESSO = 3"
     sql = sql & " And TPR_STATUS = TGE_CODIGO"
     sql = sql & " and TRE_TPR_NUMERO_PROCESSO =   TPR_NUMERO_PROCESSO "
    
  
    If Im <> "" Then sql = sql & " and TPR_INSCRICAO = '" & Im & "'"
    
    sql = sql & " order by TRE_TPR_NUMERO_PROCESSO"
  
    If Not grdDados.Preencher(Bdados, sql, 1200, 1200, 5500, 1500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) Then
        Avisa "Busca sem resultados"
    End If
    

End Sub
