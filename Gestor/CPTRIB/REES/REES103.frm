VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form REES103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REES103"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   6
      Top             =   6180
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   8115
         TabIndex        =   5
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   330
         Left            =   6990
         TabIndex        =   4
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Caption         =   "&Imprimir"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   1138
      Icone           =   "REES103.frx":0000
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1020
      Left            =   30
      TabIndex        =   8
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   660
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   1799
      Altura          =   1905
      Caption         =   " Dados do Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   285
         Left            =   450
         TabIndex        =   2
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
         Enabled         =   0   'False
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3165
         TabIndex        =   3
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
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   375
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
         Enabled         =   0   'False
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin ActiveTabs.SSActiveTabs tabRegime 
      Height          =   4410
      Left            =   45
      TabIndex        =   9
      Tag             =   "Documento gerencial"
      Top             =   1740
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   7779
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      Tabs            =   "REES103.frx":031A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel11 
         Height          =   4020
         Left            =   -99969
         TabIndex        =   10
         Top             =   30
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7091
         _Version        =   131082
         TabGuid         =   "REES103.frx":040B
         Begin VTOcx.fraVISUAL fraVISUAL11 
            Height          =   1605
            Left            =   15
            TabIndex        =   11
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
            Begin VTOcx.txtVISUAL txtMatrRespiConclui 
               Height          =   480
               Left            =   5025
               TabIndex        =   13
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
            Begin VTOcx.txtVISUAL txtDataConcl 
               Height          =   480
               Left            =   7065
               TabIndex        =   12
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
         End
         Begin VTOcx.fraVISUAL fraVISUAL12 
            Height          =   2865
            Left            =   0
            TabIndex        =   15
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
               TabIndex        =   17
               Top             =   825
               Width           =   9060
            End
            Begin VTOcx.cboVISUAL cboStatus 
               Height          =   315
               Left            =   60
               TabIndex        =   16
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
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4020
         Left            =   -99969
         TabIndex        =   18
         Top             =   30
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7091
         _Version        =   131082
         TabGuid         =   "REES103.frx":0433
         Begin ActiveTabs.SSActiveTabs tabInstrucao 
            Height          =   4050
            Left            =   -30
            TabIndex        =   19
            Tag             =   "Documento gerencial"
            Top             =   0
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   7144
            _Version        =   131082
            TabCount        =   2
            Tabs            =   "REES103.frx":045B
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel10 
               Height          =   3660
               Left            =   -99969
               TabIndex        =   20
               Top             =   360
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   6456
               _Version        =   131082
               TabGuid         =   "REES103.frx":04F9
               Begin VTOcx.fraVISUAL fraVISUAL7 
                  Height          =   3405
                  Left            =   0
                  TabIndex        =   21
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
                     TabIndex        =   23
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
                        TabIndex        =   26
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
                        TabIndex        =   25
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
                     Begin VTOcx.txtVISUAL txtCPFCont 
                        Height          =   480
                        Left            =   5070
                        TabIndex        =   24
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
                     TabIndex        =   22
                     Tag             =   "Declaração Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel9 
               Height          =   3510
               Left            =   -99969
               TabIndex        =   27
               Top             =   360
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   6191
               _Version        =   131082
               TabGuid         =   "REES103.frx":0521
               Begin VTOcx.fraVISUAL fraVISUAL6 
                  Height          =   3285
                  Left            =   0
                  TabIndex        =   28
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
                     TabIndex        =   29
                     Tag             =   "Declaração Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
               Height          =   3660
               Left            =   30
               TabIndex        =   30
               Top             =   360
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   6456
               _Version        =   131082
               TabGuid         =   "REES103.frx":0549
               Begin VTOcx.fraVISUAL fraVISUAL5 
                  Height          =   3405
                  Left            =   0
                  TabIndex        =   31
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
                     TabIndex        =   33
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
                        TabIndex        =   36
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
                        TabIndex        =   35
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
                     Begin VTOcx.txtVISUAL txtMatricula 
                        Height          =   480
                        Left            =   5085
                        TabIndex        =   34
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
                     TabIndex        =   32
                     Tag             =   "Declaração Fiscal"
                     Top             =   315
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel8 
               Height          =   3270
               Left            =   -99969
               TabIndex        =   37
               Top             =   30
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   5768
               _Version        =   131082
               TabGuid         =   "REES103.frx":0571
               Begin VTOcx.grdVISUAL grdVISUAL1 
                  Height          =   3180
                  Left            =   15
                  TabIndex        =   38
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
         TabIndex        =   39
         Top             =   30
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7091
         _Version        =   131082
         TabGuid         =   "REES103.frx":0599
         Begin ActiveTabs.SSActiveTabs tabInstauracao 
            Height          =   2985
            Left            =   -45
            TabIndex        =   40
            Tag             =   "Documento gerencial"
            Top             =   -15
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   5265
            _Version        =   131082
            TabCount        =   4
            Tabs            =   "REES103.frx":05C1
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
               Height          =   2595
               Left            =   -99969
               TabIndex        =   41
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES103.frx":06C0
               Begin VTOcx.fraVISUAL fraVISUAL1 
                  Height          =   2595
                  Left            =   0
                  TabIndex        =   42
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
                     TabIndex        =   43
                     Tag             =   "Nota Fiscal"
                     Top             =   285
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
               Height          =   2595
               Left            =   30
               TabIndex        =   44
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES103.frx":06E8
               Begin VTOcx.fraVISUAL fraVISUAL2 
                  Height          =   2595
                  Left            =   15
                  TabIndex        =   45
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
                     TabIndex        =   46
                     Tag             =   "Livro Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
               Height          =   2595
               Left            =   -99969
               TabIndex        =   47
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES103.frx":0710
               Begin VTOcx.fraVISUAL fraVISUAL3 
                  Height          =   2625
                  Left            =   -15
                  TabIndex        =   48
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
                     TabIndex        =   49
                     Tag             =   "Declaração Fiscal"
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
            Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
               Height          =   2595
               Left            =   -99969
               TabIndex        =   50
               Top             =   360
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   4577
               _Version        =   131082
               TabGuid         =   "REES103.frx":0738
               Begin VTOcx.fraVISUAL fraVISUAL4 
                  Height          =   3285
                  Left            =   0
                  TabIndex        =   51
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
                     TabIndex        =   52
                     Top             =   300
                     Width           =   9060
                  End
               End
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL8 
            Height          =   855
            Left            =   30
            TabIndex        =   53
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
            Begin VTOcx.txtVISUAL txtResp 
               Height          =   480
               Left            =   75
               TabIndex        =   55
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
            Begin VTOcx.txtVISUAL txtCPF 
               Height          =   480
               Left            =   5925
               TabIndex        =   54
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
         End
      End
   End
End
Attribute VB_Name = "REES103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GeraCod As New ContaCorrente


Private Sub cmdLimpar_Click()
    LimpaCampos Me
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

Private Sub Form_Activate()
   carregaResp REES103.Tag
End Sub

Private Sub carregaResp(Num As String)
    Dim sql As String
    Dim rs As VSRecordset
    Dim ins As String
    Dim status As String
     sql = "select TPR_INSCRICAO ,"
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
     sql = sql & " Where TPR_NUMERO_PROCESSO = '" & Num & "'"
     sql = sql & " And TPR_STATUS = TGE_CODIGO"
     sql = sql & " and TRE_TPR_NUMERO_PROCESSO =   TPR_NUMERO_PROCESSO "
          
    If Bdados.AbreTabela(sql, rs) Then
    
            ins = "" & rs!TPR_INSCRICAO
            If ins <> "" Then
                txtIm = ins
                txtIm_LostFocus
            End If
            status = "" & rs!TRE_STATUS_CONCLUSAO
            
            txtLivro = "" & rs!TRE_LIVROS_FISCAIS_MODELOS
            txtNota = "" & rs!TRE_DESCRICAO_NOTA_FISCAL
            txtDeclaracao = "" & rs!TRE_DESCRICAO_DECLARACAO
            txtDocumento = "" & rs!TRE_DESCRICAO_DOCUMENTO_FISCAL
            txtResp = "" & rs!TPR_PEDIDO_REPR_PREPOSTO
            txtCPF = "" & rs!TPR_PEDIDO_REPR_PREPOSTO_CPF
            txtDespAutor = "" & rs!TPR_FUNCIONARIO_DESPACHO
            txtRespAutor = "" & rs!TPR_FUNCIONARIO_NOME
            txtMatricula = "" & rs!TPR_FUNCIONARIO_MATRICULA
            txtDataAutor = "" & rs!TPR_FUNCIONARIO_DATA_VISTO
            txtDespCont = "" & rs!TPR_INSTRUCAO_PASSIVO_DESPACHO
            txtRespCont = "" & rs!TPR_INST_PASSIVO_REPR_PREPOSTO
            txtCPFCont = "" & rs!TPR_INST_PAS_REPR_PREPOSTO_CPF
            txtDataCont = "" & rs!TPR_INSTRUCAO_PASSIVO_DATA
            cboStatus.SetarLinha CInt(status), 1
            txtrespConclu = "" & rs!TRE_FUNCIONARIO_CONCLUSAO
            txtMatrRespiConclui = "" & rs!TRE_MATRICULA_FUNC_CONCLUSAO
            txtConclusao = "" & rs!TRE_DESCRICAO_CONCLUSAO
            txtDataConcl = "" & rs!TRE_DATA_CONCLUSAO
            
    End If

End Sub

