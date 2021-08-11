VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{467EEF11-5281-4102-AFD3-AD54F754C329}#1.5#0"; "VTControles.ocx"
Object = "{741D44DD-BF8E-4BC8-85FF-338C9BF39DFB}#1.0#0"; "Cabecalho.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Begin VB.Form RREC101 
   BackColor       =   &H00FBEDE8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RREC101"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   Icon            =   "RREC101.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2430
      Top             =   -345
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RREC101.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RREC101.frx":0BE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   5505
      Left            =   15
      TabIndex        =   21
      Top             =   645
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   9710
      _Version        =   131082
      BackColor       =   16510440
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseImageList    =   -1  'True
      Tabs            =   "RREC101.frx":14BE
      ImageList       =   "ImageList1"
      Images          =   "RREC101.frx":154B
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5085
         Left            =   30
         TabIndex        =   23
         Top             =   30
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   8969
         _Version        =   131082
         TabGuid         =   "RREC101.frx":157D
         Begin VB.Frame FraParcela 
            BackColor       =   &H00FBEDE8&
            BorderStyle     =   0  'None
            Height          =   360
            Left            =   4470
            TabIndex        =   18
            Top             =   1770
            Width           =   1125
            Begin VTOcx.txtVISUAL txtParcela 
               Height          =   285
               Left            =   0
               TabIndex        =   5
               Top             =   30
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   503
               Caption         =   "Parcela"
               Text            =   ""
               MaxLen          =   8
               PictureFundo    =   "RREC101.frx":15A5
            End
         End
         Begin VB.Frame FraValor 
            BackColor       =   &H00FBEDE8&
            BorderStyle     =   0  'None
            Height          =   1560
            Left            =   7470
            TabIndex        =   30
            Top             =   3105
            Width           =   2685
            Begin VB.TextBox txtValor 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1440
               MaxLength       =   20
               TabIndex        =   10
               Tag             =   "Valor"
               ToolTipText     =   "Valor"
               Top             =   45
               Width           =   1185
            End
            Begin Threed.SSPanel lblEscola 
               Height          =   225
               Index           =   1
               Left            =   960
               TabIndex        =   31
               Top             =   75
               Width           =   450
               _ExtentX        =   794
               _ExtentY        =   397
               _Version        =   196610
               ForeColor       =   0
               BackColor       =   16510440
               Windowless      =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Valor"
               BorderWidth     =   1
               BevelOuter      =   0
               AutoSize        =   1
               Alignment       =   0
               RoundedCorners  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDesconto 
               Height          =   285
               Left            =   600
               TabIndex        =   11
               Top             =   345
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   503
               Caption         =   "Desconto"
               Text            =   ""
               Formato         =   5
               ValorPadrao     =   "0"
               MaxLen          =   8
               PictureFundo    =   "RREC101.frx":15C1
            End
            Begin VTOcx.txtVISUAL txtValorAreceber 
               Height          =   285
               Left            =   15
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   645
               Width           =   2610
               _ExtentX        =   4604
               _ExtentY        =   503
               Caption         =   "Valor a Receber"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               ValorPadrao     =   "0"
               MaxLen          =   8
               PictureFundo    =   "RREC101.frx":15DD
            End
            Begin VTOcx.txtVISUAL txtValorPAgo 
               Height          =   285
               Left            =   105
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   960
               Width           =   2520
               _ExtentX        =   4445
               _ExtentY        =   503
               Caption         =   "Valor Recebido"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               ValorPadrao     =   "0"
               MaxLen          =   8
               PictureFundo    =   "RREC101.frx":15F9
            End
            Begin VTOcx.txtVISUAL txtSaldoDevedor 
               Height          =   285
               Left            =   120
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   1260
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               Caption         =   "Saldo Devedor"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               ValorPadrao     =   "0"
               MaxLen          =   8
               PictureFundo    =   "RREC101.frx":1615
            End
         End
         Begin VB.Frame FraServico 
            BackColor       =   &H00FBEDE8&
            BorderStyle     =   0  'None
            Height          =   360
            Left            =   705
            TabIndex        =   29
            Top             =   1395
            Width           =   9450
            Begin VTOcx.cboVISUAL cboServiço 
               Height          =   315
               Left            =   0
               TabIndex        =   3
               Tag             =   "Serviço / Curso"
               Top             =   0
               Width           =   9420
               _ExtentX        =   16616
               _ExtentY        =   556
               Caption         =   "Serviço"
               Text            =   ""
               TipoCampo       =   ""
               PictureFundo    =   "RREC101.frx":1631
            End
         End
         Begin VB.Frame FraAluno 
            BackColor       =   &H00FBEDE8&
            BorderStyle     =   0  'None
            Height          =   405
            Left            =   810
            TabIndex        =   28
            Top             =   345
            Width           =   9360
            Begin VTOcx.cboVISUAL CboAluno 
               Height          =   315
               Left            =   60
               TabIndex        =   1
               Tag             =   "Aluno"
               Top             =   60
               Width           =   9270
               _ExtentX        =   16351
               _ExtentY        =   556
               Caption         =   "Aluno"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorRotulo       =   0
               TipoCampo       =   ""
               PictureFundo    =   "RREC101.frx":164D
            End
         End
         Begin VTOcx.txtVISUAL txtDataVencimento 
            Height          =   315
            Left            =   5610
            TabIndex        =   6
            Top             =   1800
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   556
            Caption         =   "Dt.Vencimento"
            Text            =   ""
            Formato         =   0
            PictureFundo    =   "RREC101.frx":1669
         End
         Begin VB.TextBox txtOBS 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   570
            Left            =   1410
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            ToolTipText     =   "Observações"
            Top             =   2520
            Width           =   8715
         End
         Begin VB.TextBox txtData 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Data"
            ToolTipText     =   "Data"
            Top             =   1800
            Width           =   1275
         End
         Begin VB.TextBox txtDesc 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   585
            Left            =   1395
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Tag             =   "Histórico"
            ToolTipText     =   "Histórico"
            Top             =   750
            Width           =   8730
         End
         Begin Threed.SSPanel lblEscola 
            Height          =   225
            Index           =   2
            Left            =   135
            TabIndex        =   24
            Top             =   1860
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   397
            _Version        =   196610
            ForeColor       =   0
            BackColor       =   16510440
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Dt.Documento"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lblEscola 
            Height          =   225
            Index           =   4
            Left            =   270
            TabIndex        =   25
            Top             =   2490
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   397
            _Version        =   196610
            ForeColor       =   0
            BackColor       =   16510440
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Observações"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lblEscola 
            Height          =   225
            Index           =   3
            Left            =   600
            TabIndex        =   26
            Top             =   720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            _Version        =   196610
            ForeColor       =   0
            BackColor       =   16510440
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Histórico"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin VTOcx.cboVISUAL cboTipo 
            Height          =   315
            Left            =   645
            TabIndex        =   8
            Tag             =   "Serviço"
            Top             =   2160
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   556
            Caption         =   "Contabil"
            Text            =   ""
            TipoCampo       =   ""
            PictureFundo    =   "RREC101.frx":1685
         End
         Begin VTOcx.cmdVISUAL cmdDetalhes 
            Height          =   285
            Left            =   4575
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2175
            Visible         =   0   'False
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            Caption         =   "..."
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtMatricula 
            Height          =   285
            Left            =   8175
            TabIndex        =   7
            Top             =   1785
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   503
            Caption         =   "Matricula"
            Text            =   ""
            Enabled         =   0   'False
            PictureFundo    =   "RREC101.frx":16A1
         End
         Begin VTOcx.cboVISUAL cboEscola 
            Height          =   315
            Left            =   810
            TabIndex        =   0
            Tag             =   "Aluno"
            Top             =   45
            Width           =   9315
            _ExtentX        =   16431
            _ExtentY        =   556
            Caption         =   "Escola"
            Text            =   ""
            AutoFocaliza    =   0   'False
            CorRotulo       =   0
            TipoCampo       =   ""
            PictureFundo    =   "RREC101.frx":16BD
         End
         Begin VTOcx.txtVISUAL txtAcrescimo 
            Height          =   285
            Left            =   7980
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   4680
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   503
            Caption         =   "Acrescimo"
            Text            =   ""
            Enabled         =   0   'False
            Formato         =   5
            ValorPadrao     =   "0"
            MaxLen          =   8
            PictureFundo    =   "RREC101.frx":16D9
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5085
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   8969
         _Version        =   131082
         TabGuid         =   "RREC101.frx":16F5
         Begin VTOcx.cmdVISUAL cmdImiprimir 
            Height          =   405
            Left            =   6120
            TabIndex        =   48
            Top             =   4635
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   714
            Caption         =   "&Imprimir Boletos"
            Acao            =   4
            CorFundo        =   12648447
         End
         Begin VTOcx.cmdVISUAL cmdMudarStatusAqruivo 
            Height          =   405
            Left            =   7950
            TabIndex        =   46
            Top             =   4635
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   714
            Caption         =   "&Mudar Status Arquivo"
            Acao            =   1
            CorFundo        =   12640511
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FBEDE8&
            Caption         =   "Selecionar Todos"
            Height          =   240
            Left            =   105
            TabIndex        =   44
            Top             =   4755
            Width           =   1830
         End
         Begin VTOcx.grdVISUAL Grid 
            Height          =   3180
            Left            =   60
            TabIndex        =   32
            Top             =   1785
            Width           =   10110
            _ExtentX        =   17833
            _ExtentY        =   5609
            Caption         =   "Relação de recebimentos Encontrados"
            CorDica         =   255
            OcultarRodape   =   -1  'True
            CheckBox        =   -1  'True
            PictureBarra    =   "RREC101.frx":171D
         End
         Begin VTOcx.fraVISUAL txt 
            Height          =   1725
            Left            =   45
            TabIndex        =   33
            Top             =   30
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   3043
            Altura          =   1905
            Caption         =   " Consultar Por:"
            CorTexto        =   0
            CorFaixa        =   12632256
            CorFundo        =   16510440
            Ocultavel       =   0   'False
            BackStyle       =   0
            Picture         =   "RREC101.frx":1739
            Picture2        =   "RREC101.frx":6B23
            Begin VTOcx.txtVISUAL txtFim 
               Height          =   285
               Left            =   2715
               TabIndex        =   38
               Top             =   1005
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   503
               Caption         =   "Até"
               Text            =   ""
               Formato         =   0
               PictureFundo    =   "RREC101.frx":6B3F
            End
            Begin VTOcx.txtVISUAL txtInicio 
               Height          =   285
               Left            =   165
               TabIndex        =   37
               Top             =   1005
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   503
               Caption         =   "Vencimento"
               Text            =   ""
               Formato         =   0
               PictureFundo    =   "RREC101.frx":6B5B
            End
            Begin VTOcx.cboVISUAL cboRepresentante_Financeiro 
               Height          =   315
               Left            =   90
               TabIndex        =   41
               Tag             =   "Representante Financeiro"
               Top             =   1350
               Width           =   9840
               _ExtentX        =   17357
               _ExtentY        =   556
               Caption         =   "Representante Financeiro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               TipoLetras      =   0
               TipoCampo       =   ""
               PictureFundo    =   "RREC101.frx":6B77
            End
            Begin VTOcx.txtVISUAL txtNumero 
               Height          =   285
               Left            =   480
               TabIndex        =   34
               Top             =   315
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   503
               Caption         =   "Número"
               Text            =   ""
               PictureFundo    =   "RREC101.frx":6B93
            End
            Begin VTOcx.cboVISUAL cboStatus 
               Height          =   315
               Left            =   7005
               TabIndex        =   40
               Top             =   990
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   556
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorRotulo       =   0
               TipoCampo       =   ""
               PictureFundo    =   "RREC101.frx":6BAF
            End
            Begin VTOcx.txtVISUAL txtNotaConsulta 
               Height          =   285
               Left            =   4845
               TabIndex        =   39
               Top             =   1020
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   503
               Caption         =   "Matricula"
               Text            =   ""
               PictureFundo    =   "RREC101.frx":6BCB
            End
            Begin VTOcx.cboVISUAL cboAlunoConsulta 
               Height          =   315
               Left            =   2430
               TabIndex        =   35
               Top             =   315
               Width           =   7485
               _ExtentX        =   13203
               _ExtentY        =   556
               Caption         =   "Aluno"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorRotulo       =   0
               TipoCampo       =   ""
               PictureFundo    =   "RREC101.frx":6BE7
            End
            Begin VTOcx.cboVISUAL cboServicoConsulta 
               Height          =   315
               Left            =   510
               TabIndex        =   36
               Top             =   645
               Width           =   9420
               _ExtentX        =   16616
               _ExtentY        =   556
               Caption         =   "Serviço"
               Text            =   ""
               TipoCampo       =   ""
               PictureFundo    =   "RREC101.frx":6C03
            End
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   1138
      Icone           =   "RREC101.frx":6C1F
      ImagemFundo     =   "RREC101.frx":6F39
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   19
      Top             =   6180
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   900
      CorFundo        =   -2147483633
      ImagemFundo     =   "RREC101.frx":1AE93
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   2220
         TabIndex        =   47
         Top             =   75
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         Caption         =   "&Boleto"
         Acao            =   4
      End
      Begin VTOcx.cmdVISUAL cmdGegarConta 
         Height          =   375
         Left            =   3150
         TabIndex        =   45
         Top             =   75
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "&Gerar Débitos Material"
         Acao            =   3
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   5460
         TabIndex        =   43
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Cancelar "
         Acao            =   2
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   6600
         TabIndex        =   42
         Top             =   75
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
      End
      Begin VTOcx.cmdVISUAL cmdGravar 
         Height          =   375
         Left            =   7620
         TabIndex        =   15
         Top             =   75
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
      End
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   8640
         TabIndex        =   16
         Top             =   75
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   661
         Caption         =   "&Novo"
         Acao            =   1
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9525
         TabIndex        =   17
         Top             =   75
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
      End
   End
   Begin ActiveResizer.SSResizer SSResizer2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196610
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   10380
      DesignHeight    =   6690
   End
End
Attribute VB_Name = "RREC101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Vbruto As String, Vinss As String, Viss As String, Virrf As String, Voutros As String
Private Trans As String
Private Ordem As String
Private Sub cboConta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboContaB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub


'Private Sub cmdBoleto_Click()
'    ImprimeDam Grid.SelectedItem, CStr(cboAluno.Coluna(0).Valor), DebitoNormal, , txtParcela, txtValor, 0, 0, txtValor, , txtDataVencimento, 0, Debito, CStr(cboServiço.Coluna(0).Valor)
'End Sub

Private Sub Check1_Click()
    Dim i As Integer
    
    For i = 1 To Grid.ListItems.Count
        Grid.ListItems(i).Checked = Check1.Value
    Next
End Sub

Private Sub cmdBuscar_Click()
    Dim sql As String
'    Dim RS  As VSRecordset
'    Dim rsa  As VSRecordset
'    Dim rsb  As VSRecordset
'    Dim campos As String
'    Dim valores As String
'
'    sql = "Select * from TAB_TEMP_MOV_FINANCEIRO"
'    If Bdados.AbreTabela(sql, RS) Then
'        Do Until RS.EOF
'            sql = "Select * from tab_conta_receber where TRC_MOVIMENTO_FIn = " & RS.Fields("numero")
'            If Bdados.AbreTabela(sql, rsa) Then
'                Do Until rsa.EOF
'                    campos = "TBR_TCR_CODIGO,"
'                    campos = campos & "TBR_ORDEM,"
'                    campos = campos & "TBR_OPERACAO,"
'                    campos = campos & "TBR_VALOR_PAGO,"
'                    campos = campos & "TBR_MULTA,"
'                    campos = campos & "TBR_JUROS,"
'                    campos = campos & "TBR_DESCONTO,"
'                    campos = campos & "TBR_SUB_TOTAL,"
'                    campos = campos & "TBR_FORMA_PAGAMENTO,"
'                    campos = campos & "TBR_DATA_PAGAMENTO,"
'                    campos = campos & "TBR_TCB_CONTA"
'                    valores = Bdados.PreparaValor(rsa.Fields("TCR_COD_CONTA"), 1, 1, (RS.Fields("valor") - rsa.Fields("tcr_multa")), rsa.Fields("tcr_multa"), 0, 0, RS.Fields("VALOR"), RS.Fields("FORMA_PAGAMENTO"), RS.Fields("data_baixa"), 1)
'                    Bdados.InsereDados "TAB_BAIXA_RECEBIMENTO", valores, campos
'                    rsa.MoveNext
'                Loop
'            End If
'            RS.MoveNext
'        Loop
'    End If
'    Exit Sub
 '   If Not Valida_Empresa(cboEmp) Then Exit Sub
    sql = "Select * from VIS_CONTA_receber where 1 =1 "
    If cboAlunoConsulta.Text <> "" Then
        sql = sql & " AND Aluno = '" & cboAlunoConsulta.Text & "'"
    End If
    If txtNumero <> "" Then
        sql = sql & " and código = " & txtNumero
    End If
    If cboServicoConsulta <> "" Then
        sql = sql & " AND [Curso / Serviço] = '" & cboServicoConsulta.Text & "'"
    End If
    
    
    If txtNotaConsulta <> "" Then
        sql = sql & " and Matricula = " & txtNotaConsulta
    End If
    If cboStatus <> "" Then
        sql = sql & " and Status = '" & cboStatus.Text & "'"
    End If
    If cboRepresentante_Financeiro.Text <> "" Then
        sql = sql & " and CodRepFin  = " & cboRepresentante_Financeiro.Coluna(0).Valor
    End If
    If txtInicio <> "" And txtFim <> "" Then
        sql = sql & " and Vencimento >= " & Bdados.Converte(txtInicio, TCDataHora)
        sql = sql & " and Vencimento <= " & Bdados.Converte(txtFim, TCDataHora)
    ElseIf txtInicio <> "" And txtFim = "" Then
        sql = sql & " and Vencimento >= " & Bdados.Converte(txtInicio, TCDataHora)
        sql = sql & " and Vencimento <= " & Bdados.Converte(txtInicio, TCDataHora)
    End If
    'Sql = Sql & " order by TCR_descricao"
    If Grid.Preencher(Bdados, sql) Then
        'Grid.Mensagem = "Tot Original : " & Format(Grid.Colunas(9).Soma, "currency") & "| Tot Desconto : " & Format(Grid.Colunas(10).Soma, "currency") & "| Tot a Pagar : " & Format(Grid.Colunas(11).Soma, "currency") & "| Total Pago : " & Format(Grid.Colunas(12).Soma, "currency") & "| Saldo Devedor : " & Format(Grid.Colunas(13).Soma, "currency")
    Else
        Avisa "Consulta sem resultados."
        'Grid.Mensagem = "Nenhum registro encontrado."
    End If
End Sub

Private Function GravaRetencao(Trans As String) As Boolean
    Dim Val As String
    If Vbruto = "" Then
        GravaRetencao = True
        Exit Function
    End If
    
    Val = Bdados.PreparaValor(Trans, Vbruto, Vinss, Viss, Virrf, Voutros)
    GravaRetencao = Bdados.GravaDados("TAB_RETENCAO", Val, "TRE_COD_RETENCAO,TRE_BRUTO,TRE_INSS,TRE_ISS,TRE_IRRF,TRE_OUTROS", "TRE_COD_RETENCAO=" & Trans)
End Function

Private Sub cmdExcluir_Click()
    Dim i As Integer
    
    If Grid.ListItems.Count >= 1 Then
        If Confirma("Confirma o cancelamento do(s) recebimeto(s) selecionado(s)") Then
            For i = 1 To Grid.ListItems.Count
                If Grid.ListItems(i).Checked Then
                    If Bdados.GravaDados("TAB_CONTA_RECEBER", Bdados.PreparaValor(esrCancelado), "TCR_STATUS", "TCR_COD_CONTA = " & Grid.ListItems(i)) Then
                        Call Bdados.GravaDados("TAB_CONTA_RECEBER", Bdados.PreparaValor(0, 0, 0, 0, 0, 0, 0), "TCR_VALOR,TCR_DESCONTO,TCR_VALOR_APAGAR,TCR_VALOR_PAGO,TCR_SALDO_DEVEDOR,TCR_JUROS,TCR_MULTA", "TCR_COD_CONTA = " & Grid.ListItems(i))
                    End If
                End If
            Next
            Avisa "Valores zerados com sucesso."
        End If
    End If
End Sub

Private Sub cmdGegarConta_Click()
    Load REC101a
    REC101a.Show 1
End Sub

Private Sub cmdImiprimir_Click()
    Dim i As Integer
    
    For i = 1 To Grid.ListItems.Count
        If Grid.ListItems(i).Checked Then
            Imprimir_Boleto Grid.ListItems(i), eBBBanco_Bradesco, ediIpressora
        End If
    Next
End Sub

Private Sub CmdImprimir_Click()
    Imprimir_Boleto Grid.SelectedItem, eBBBanco_Bradesco
End Sub

Private Sub cmdMudarStatusAqruivo_Click()
    Dim i As Integer
    If Grid.ListItems.Count >= 1 Then
        For i = 1 To Grid.ListItems.Count
            If Grid.ListItems(i).Checked Then
                If PegaStatusArquivoDebito(Grid.ListItems(i)) = esadGerado Then
                    Call MudaStatusArquivoDebito(Grid.ListItems(i), esadNaoGerado)
                Else
                    Call MudaStatusArquivoDebito(Grid.ListItems(i), esadGerado)
                End If
                
                If PegaStatusArquivoDebitoBradesco(Grid.ListItems(i)) = esadGerado Then
                    Call MudaStatusArquivoDebitoBradesco(Grid.ListItems(i), esadNaoGerado)
                Else
                    Call MudaStatusArquivoDebitoBradesco(Grid.ListItems(i), esadGerado)
                End If
                
            End If
        Next
        Avisa "Operação concluída com sucesso"
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdNovo_Click()
    LimpaCampos Me
    Trans = ""
    TabDados.Tabs(2).Selected = True
    FraAluno.Enabled = True
    FraParcela.Enabled = True
    FraValor.Enabled = True
    FraServico.Enabled = True
    CboAluno.SetFocus
End Sub

Private Sub SSActiveTabs1_BeforeTabClick(ByVal NewTab As ActiveTabs.SSTab, ByVal Cancel As ActiveTabs.SSReturnBoolean)

End Sub

Private Sub SSActiveTabs1_TabClick(ByVal NewTab As ActiveTabs.SSTab)
'    if t
End Sub

Private Sub Grid_DblClick()
    If Grid.ListItems.Count >= 1 Then
        If Bdados.AbreTabela("Select * from tab_conta_receber where TCR_COD_CONTA = " & Grid.SelectedItem) Then
            CboAluno.SetarLinha "" & Bdados.Tabela("TCR_TAN_ALUNO")
            cboServiço.SetarLinha "" & Bdados.Tabela("TCR_curso")
            txtData = "" & Bdados.Tabela("TCR_DATA")
            txtOBS = "" & Bdados.Tabela("TCR_OBS")
            cboEscola.SetarLinha "" & Bdados.Tabela("TCR_ESCOLA")
            txtMatricula = "" & Bdados.Tabela("TCR_MATRICULA")
            txtDesc = "" & Bdados.Tabela("TCR_DESCRICAO")
            cboTipo.SetarLinha "" & Bdados.Tabela("TCR_TTI_COD_TIPO")
            txtParcela = "" & Bdados.Tabela("TCR_PARCELA")
            txtDesconto = "" & Bdados.Tabela("tcr_desconto")
            txtSaldoDevedor = "" & Bdados.Tabela("tcr_saldo_devedor")
            txtValorPago = "" & Bdados.Tabela("tcr_valor_pago")
            txtValorAreceber = "" & Bdados.Tabela("tcr_valor_apagar")
            txtValor = "" & Bdados.Tabela("TCR_valor")
            txtValor = FormataTexto(txtValor, Monetario)
            txtDataVencimento = "" & Bdados.Tabela("TCR_VENCIMENTO")
            txtAcrescimo = ((CCur(txtValorPago) - CCur(txtValorAreceber)) * -1) - CCur(txtSaldoDevedor)
            TabDados.Tabs(2).Selected = True
            Trans = Grid.SelectedItem
            If Bdados.Tabela("TCR_STATUS") = esrQuitado Then
                FraValor.Enabled = False
            Else
                FraValor.Enabled = True
            End If
            Ordem = "" & Bdados.Tabela("TCR_TSP_ORDEM")
            
        End If
    End If
End Sub

Private Sub TabDados_TabClick(ByVal NewTab As ActiveTabs.SSTab)
    If TabDados.Tabs(1).Selected Then
        Grid.SetFocus
    Else
        If FraAluno.Enabled Then
            cboEscola.SetFocus
        Else
            txtDesc.SetFocus
        End If
    End If
End Sub

Private Sub txtData_LostFocus()
    If Trim(txtData) = "" Then
        txtData = Format(Date, "dd/mm/yyyy")
    Else
        If Not IsDate(txtData) _
        Then txtData = FormataTexto(txtData, Data)
    End If
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDescConsulta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDtPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDesconto_LostFocus()
    CalcSUb
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Function Atualizar() As Boolean
    On Error GoTo TRATA
    Dim Valores As String
    Dim Campos As String
    Dim Conta As String
    Dim Valor As String
    Dim Saldo As String
    Dim rs As VSRecordset
    Dim Status As String
    
    
    Dim Acao As String
    If Trans <> "" Then
        Trans = Grid.SelectedItem
        If Confirma("Confirma alteração?") Then
            
        Else
            Exit Function
        End If
        Acao = "Alteração"
    Else
        Trans = GeraCorrelativo(ecContaReceber)
        Campos = "TCR_COD_CONTA,"
        Valores = Bdados.PreparaValor(Trans)
        Acao = "Cadastro"
    End If
    
    
    
    Valor = txtValor
    
    
    Conta = GeraContaReceber(IIf(Trans <> "", Trans, ""), txtMatricula, CStr(CboAluno.Coluna(0).Valor), CStr(cboServiço.Coluna(0).Valor), txtDesc, txtData, txtValor, 0, 0, txtDesconto, txtValorAreceber, txtOBS, CStr(cboTipo.Coluna(0).Valor), esrAberto, Aplica.Usuario, txtParcela, txtValor, txtSaldoDevedor, txtDataVencimento, 0, esadNaoGerado, CStr(cboEscola.Coluna(0).Valor), txtValorPago, Year(txtData))
    
    Bdados.AbreTrans
    If Conta <> "" Then
        Bdados.GravaTrans
        Informa "Dados salvos com sucesso."
        'Gravo o Log...
        Campos = "TCR_COD_CONTA,TCR_DESCRICAO,TCR_DATA,TCR_VALOR,TCR_OBS,TCR_TTI_COD_TIPO,TCR_STATUS,TCR_TUS_COD_USUARIO,TCR_PARCELA,TCR_VENCIMENTO,TCR_TAN_ALUNO,TCR_CURSO,TCR_DATA_LOG,TCR_HORA,TCR_ACAO"
        Valores = Bdados.PreparaValor(Trans, Trim(txtDesc), txtData, Valor, Trim(txtOBS), cboTipo.Coluna(0).Valor, esrAberto, Aplica.Usuario, txtParcela, Bdados.Converte(txtDataVencimento, TCDataHora), CboAluno.Coluna(0).Valor, cboServiço.Coluna(0).Valor, Date, Bdados.Converte(Time, tctexto), Acao)
        Call Bdados.GravaDados("TAB_CONTA_RECEBER_LOG", Valores, Campos, "TCR_COD_CONTA = " & Trans & " AND TCR_HORA =   '" & Time & "'")
        Atualizar = True
        If txtMatricula <> "" Then
            If PegaTipoServicoCurso(CStr(cboServiço.Coluna(0).Valor)) = eTipoContabilCurso Then
                'Atualizo a TAB_PARCELA_TURMA
                Call Bdados.GravaDados("TAB_PARCELA_TURMA", Bdados.PreparaValor(Bdados.Converte(txtDataVencimento, TCDataHora), Day(txtDataVencimento), txtValorAreceber), "TPT_VENCIMENTO,TPT_DIA_PAGAMENTO,TPT_VALOR", "TPT_TCR_COD_CONTA = " & Trans)
            ElseIf PegaTipoServicoCurso(CStr(cboServiço.Coluna(0).Valor)) = eTipoContabilServico Then
                'Atualizo a TAB_SERVICO_PARCELA
                If Bdados.AbreTabela("Select count(*) from TAB_SERVICO_PARCELA where TSP_TMS_TCU_SERVICO = " & cboServiço.Coluna(0).Valor) Then
                    If Bdados.Tabela(0) > 0 Then
                        Call Bdados.GravaDados("TAB_SERVICO_PARCELA", Bdados.PreparaValor(Bdados.Converte(txtDataVencimento, TCDataHora)), "TSP_VENCIMENTO", "TSP_TMA_MATRICULA = " & txtMatricula & " and TSP_ORDEM = " & Ordem & " and TSP_TMS_TCU_SERVICO = " & cboServiço.Coluna(0).Valor)
                    End If
                    Bdados.Tabela.Fechar
                End If
            End If
        End If
    Else
        Bdados.CancelaTrans
        Erro "Erro ao Gravar. Informações não gravadas."
    End If

    Exit Function
TRATA:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Function

Private Sub cmdGravar_Click()
    On Error GoTo TRATA
    Dim Cheque As String
    
    
    If Not IsDate(txtData) Then
        Avisa "Data Inválida."
        txtData.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = 11
    cboRepresentante_Financeiro.Tag = ""
    If CriticaCampos(Me) Then
        If Atualizar Then
            cmdNovo_Click
        End If
        
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
TRATA:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo TRATA
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    AtualizaTipo cboTipo
    cboRepresentante_Financeiro.Preencher Bdados, "Select trp_codigo,trp_nome  + '  /  ' + trp_doc from tab_representante", 1
    CboAluno.Preencher Bdados, "SELECT TAN_CODIGO,TAN_NOME FROM TAB_ALUNOS", 1
    cboAlunoConsulta.Preencher Bdados, "SELECT TAN_CODIGO,TAN_NOME FROM TAB_ALUNOS", 1
    cboServiço.Preencher Bdados, "SELECT tcu_codigo,tcu_nome FROM TAB_CURSOS", 1
    cboServicoConsulta.Preencher Bdados, "SELECT tcu_codigo,tcu_nome FROM TAB_CURSOS", 1
    cboEscola.Preencher Bdados, "Select tes_codigo,tes_nome from tab_escola", 1
    cboStatus.PreencherGeral Bdados, "STATUS COTA"
    Screen.MousePointer = 0
'    Call cmdBuscar_Click
    Exit Sub
TRATA:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtValor_LostFocus()
    txtValor = FormataTexto(txtValor, Monetario, True)
    If txtValor = "" Then
        txtValor = "0,00"
    End If
    CalcSUb
End Sub
Private Sub CalcSUb()
    If txtDesconto = "" Then
        txtDesconto = 0
    End If
    If txtSaldoDevedor = "" Then
        txtSaldoDevedor = 0
    End If
    txtValorAreceber = CCur(txtValor) - CCur(txtDesconto)
    txtSaldoDevedor = txtValorAreceber
  
    If txtValorPago = "" Then
        txtValorPago = 0
    End If
End Sub
