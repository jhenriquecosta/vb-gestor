VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TRPT401 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000016&
   Caption         =   "TRPT401"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs tabRelatorios 
      Height          =   4560
      Left            =   30
      TabIndex        =   40
      Top             =   690
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   8043
      _Version        =   131082
      BackColor       =   -2147483626
      TabCount        =   4
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
      Tabs            =   "TRPT401.frx":0000
      Images          =   "TRPT401.frx":00EB
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4140
         Index           =   1
         Left            =   30
         TabIndex        =   41
         Top             =   30
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   7303
         _Version        =   131082
         TabGuid         =   "TRPT401.frx":1311
         Begin VTOcx.grdVISUAL grdRelatorios 
            Height          =   3810
            Left            =   60
            TabIndex        =   58
            Top             =   60
            Width           =   10155
            _ExtentX        =   17912
            _ExtentY        =   4339
            Caption         =   "Relatórios Gerenciais"
            CorTitulo       =   32768
            CorCaption      =   16777215
            OcultarRodape   =   -1  'True
            CheckBox        =   -1  'True
            MarcaUnico      =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4140
         Left            =   30
         TabIndex        =   42
         Top             =   30
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   7303
         _Version        =   131082
         TabGuid         =   "TRPT401.frx":1339
         Begin VTOcx.txtVISUAL txtPeriodoFinal 
            Height          =   315
            Left            =   495
            TabIndex        =   5
            Tag             =   "Data Inicial"
            Top             =   1605
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            Caption         =   "Per. Final"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   6
         End
         Begin VTOcx.cboVISUAL cboSituacaoTributo 
            Height          =   315
            Left            =   540
            TabIndex        =   8
            Top             =   2700
            Visible         =   0   'False
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   556
            Caption         =   "Situação"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtParcela 
            Height          =   315
            Left            =   660
            TabIndex        =   7
            Tag             =   "Data Inicial"
            Top             =   2325
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "Parcela"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   1
         End
         Begin VTOcx.txtVISUAL txtNumDocumento 
            Height          =   315
            Left            =   690
            TabIndex        =   6
            Tag             =   "Data Inicial"
            Top             =   1965
            Visible         =   0   'False
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            Caption         =   "Nº Doc"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   8
         End
         Begin VTOcx.txtVISUAL txtPeriodoInicial 
            Height          =   315
            Left            =   375
            TabIndex        =   4
            Tag             =   "Data Inicial"
            Top             =   1245
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            Caption         =   "Per. Inicial"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   6
         End
         Begin VTOcx.cboVISUAL cboSiglaTributo 
            Height          =   315
            Left            =   690
            TabIndex        =   3
            Top             =   885
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            Caption         =   "Tributo"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboAgenteArrecadador 
            Height          =   315
            Left            =   5865
            TabIndex        =   9
            Top             =   435
            Visible         =   0   'False
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   556
            Caption         =   "Agente"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtDtInicialArrecadacao 
            Height          =   315
            Left            =   2880
            TabIndex        =   10
            Tag             =   "Data Inicial"
            Top             =   1245
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   556
            Caption         =   "Data Inicial (Pagamento)"
            Text            =   ""
            Formato         =   0
            Restricao       =   2
            MaxLen          =   10
         End
         Begin VTOcx.txtVISUAL txtDtFinalArrecadacao 
            Height          =   315
            Left            =   3000
            TabIndex        =   11
            Tag             =   "Data Final"
            Top             =   1605
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   556
            Caption         =   "Data Final (Pagamento)"
            Text            =   ""
            Formato         =   0
            Restricao       =   2
            MaxLen          =   10
         End
         Begin VTOcx.txtVISUAL txtTop 
            Height          =   300
            Left            =   8370
            TabIndex        =   14
            Top             =   2460
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Caption         =   "Top"
            Text            =   ""
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.cboVISUAL cboBairroNovo 
            Height          =   315
            Left            =   4860
            TabIndex        =   17
            Top             =   3180
            Visible         =   0   'False
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   556
            Caption         =   "Bairro"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Editavel        =   -1  'True
         End
         Begin VTOcx.cboVISUAL cboLogra 
            Height          =   315
            Left            =   6735
            TabIndex        =   16
            Top             =   2820
            Visible         =   0   'False
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            AutoFocaliza    =   0   'False
            Editavel        =   -1  'True
         End
         Begin VTOcx.cboVISUAL cboTipoLogra 
            Height          =   315
            Left            =   4410
            TabIndex        =   15
            Top             =   2820
            Visible         =   0   'False
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   556
            Caption         =   "Logradouro"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Editavel        =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtValorFinal 
            Height          =   315
            Left            =   7710
            TabIndex        =   13
            Tag             =   "Data Inicial"
            Top             =   2100
            Visible         =   0   'False
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            Caption         =   "Valor. Final"
            Text            =   ""
            Formato         =   5
            Restricao       =   2
         End
         Begin VTOcx.txtVISUAL txtValorInicial 
            Height          =   315
            Left            =   7590
            TabIndex        =   12
            Tag             =   "Data Inicial"
            Top             =   1740
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            Caption         =   "Valor. Inicial"
            Text            =   ""
            Formato         =   5
            Restricao       =   2
         End
         Begin VTOcx.txtVISUAL txtAnoConstrucao 
            Height          =   315
            Left            =   1635
            TabIndex        =   60
            Tag             =   "Data Inicial"
            Top             =   3420
            Visible         =   0   'False
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "Ano"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   4
            Mascara         =   "0000"
         End
         Begin VTOcx.txtVISUAL txtICImovel 
            Height          =   315
            Left            =   1080
            TabIndex        =   61
            Tag             =   "Data Inicial"
            Top             =   3060
            Visible         =   0   'False
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            Caption         =   "IC"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   14
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":. Localização"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   6
            Left            =   5235
            TabIndex        =   59
            Top             =   2550
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":. Imprimir Por:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   3
            Left            =   690
            TabIndex        =   53
            Top             =   600
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relatório"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   675
            TabIndex        =   48
            Top             =   90
            Width           =   690
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   30
            Picture         =   "TRPT401.frx":1361
            Top             =   30
            Width           =   240
         End
         Begin VB.Label lblRelatorio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ARRECADAÇÃO DA RECEITA PRÓPRIA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   240
            Index           =   0
            Left            =   660
            TabIndex        =   47
            Top             =   270
            Width           =   3960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":. Arrecadação"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   5
            Left            =   5865
            TabIndex        =   45
            Top             =   135
            Visible         =   0   'False
            Width           =   1425
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   4140
         Left            =   30
         TabIndex        =   43
         Top             =   30
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   7303
         _Version        =   131082
         TabGuid         =   "TRPT401.frx":21A3
         Begin VTOcx.cboVISUAL cboAtividadeContribuinte 
            Height          =   315
            Left            =   300
            TabIndex        =   2
            Top             =   2040
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   556
            Caption         =   "Atividade"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtRazao 
            Height          =   315
            Left            =   570
            TabIndex        =   54
            Top             =   1305
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   556
            Caption         =   "Razão"
            Text            =   ""
            Enabled         =   0   'False
         End
         Begin VTOcx.txtVISUAL txtEndereco 
            Height          =   315
            Left            =   300
            TabIndex        =   55
            Top             =   1665
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   556
            Caption         =   "Endereço"
            Text            =   ""
            Enabled         =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
            Height          =   315
            Left            =   2910
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   915
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   ""
            Acao            =   5
         End
         Begin VTOcx.txtVISUAL txtImovel 
            Height          =   300
            Left            =   5640
            TabIndex        =   1
            Top             =   945
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   529
            Caption         =   "Cadastro do Imóvel"
            Text            =   ""
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.cmdVISUAL cmdBuscaImovel 
            Height          =   315
            Left            =   9120
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   930
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   ""
            Acao            =   5
         End
         Begin VTOcx.txtVISUAL txtIm 
            Height          =   300
            Left            =   60
            TabIndex        =   0
            Top             =   930
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   529
            Caption         =   "Contribuinte"
            Text            =   ""
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relatório"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   675
            TabIndex        =   50
            Top             =   90
            Width           =   690
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   30
            Picture         =   "TRPT401.frx":21CB
            Top             =   30
            Width           =   240
         End
         Begin VB.Label lblRelatorio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ARRECADAÇÃO DA RECEITA PRÓPRIA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   240
            Index           =   1
            Left            =   660
            TabIndex        =   49
            Top             =   270
            Width           =   3960
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   4140
         Left            =   30
         TabIndex        =   44
         Top             =   30
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   7303
         _Version        =   131082
         TabGuid         =   "TRPT401.frx":300D
         Begin VTOcx.txtVISUAL txtValorVenalFim 
            Height          =   315
            Left            =   7155
            TabIndex        =   28
            Tag             =   "Data Inicial"
            Top             =   1455
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            Caption         =   "até"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
         End
         Begin VTOcx.txtVISUAL txtValorVenalInicio 
            Height          =   315
            Left            =   4635
            TabIndex        =   27
            Tag             =   "Data Inicial"
            Top             =   1455
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            Caption         =   "Valor Venal"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
         End
         Begin VTOcx.cboVISUAL cboConservacaoImovel 
            Height          =   315
            Left            =   4500
            TabIndex        =   26
            Top             =   1080
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            Caption         =   "Conservação"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboEstruturaImovel 
            Height          =   315
            Left            =   4845
            TabIndex        =   25
            Top             =   735
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            Caption         =   "Estrutura"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboTipologiaImovel 
            Height          =   315
            Left            =   705
            TabIndex        =   24
            Top             =   3540
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   556
            Caption         =   "Tipologia"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboPadraoImovel 
            Height          =   315
            Left            =   870
            TabIndex        =   23
            Top             =   3195
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   556
            Caption         =   "Padrão"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboDestinacaoImovel 
            Height          =   315
            Left            =   525
            TabIndex        =   22
            Top             =   2850
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   556
            Caption         =   "Destinação"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboUsoImovel 
            Height          =   315
            Left            =   1170
            TabIndex        =   21
            Top             =   2505
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            Caption         =   "Uso"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboOcupacaoImovel 
            Height          =   315
            Left            =   660
            TabIndex        =   20
            Top             =   2160
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            Caption         =   "Ocupacao"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboTipoImovel 
            Height          =   315
            Left            =   1155
            TabIndex        =   19
            Top             =   1815
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            Caption         =   "Tipo"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboAforado 
            Height          =   315
            Left            =   855
            TabIndex        =   18
            Top             =   1455
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            Caption         =   "Aforado"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtQuadra 
            Height          =   315
            Left            =   6660
            TabIndex        =   30
            Tag             =   "Data Inicial"
            Top             =   2505
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Caption         =   "Quadra"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   4
            Mascara         =   "0000"
         End
         Begin VTOcx.txtVISUAL txtSetor 
            Height          =   315
            Left            =   5145
            TabIndex        =   29
            Tag             =   "Data Inicial"
            Top             =   2505
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "Setor"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   2
            Mascara         =   "00"
         End
         Begin VTOcx.cboVISUAL cboBairro 
            Height          =   315
            Left            =   5055
            TabIndex        =   34
            Top             =   3555
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   556
            Caption         =   "Bairro"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboLogradouro 
            Height          =   315
            Left            =   6960
            TabIndex        =   33
            Top             =   3195
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cboVISUAL cboTipoLogradouro 
            Height          =   315
            Left            =   4635
            TabIndex        =   32
            Top             =   3195
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   556
            Caption         =   "Logradouro"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCodLogradouro 
            Height          =   315
            Left            =   4710
            TabIndex        =   31
            Tag             =   "Data Inicial"
            Top             =   2850
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            Caption         =   "Cód. Logr."
            Text            =   ""
            Restricao       =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relatório"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   675
            TabIndex        =   52
            Top             =   90
            Width           =   690
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   2
            Left            =   30
            Picture         =   "TRPT401.frx":3035
            Top             =   30
            Width           =   240
         End
         Begin VB.Label lblRelatorio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ARRECADAÇÃO DA RECEITA PRÓPRIA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   240
            Index           =   2
            Left            =   660
            TabIndex        =   51
            Top             =   270
            Width           =   3960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":. Localização"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   4
            Left            =   4500
            TabIndex        =   46
            Top             =   2220
            Width           =   1320
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   39
      Top             =   5280
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   900
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8535
         TabIndex        =   36
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   7350
         TabIndex        =   35
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9570
         TabIndex        =   37
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   1138
      Formulario      =   "TRPT401"
      Descricao       =   "Relatórios Gerenciais"
      Icone           =   "TRPT401.frx":3E77
   End
End
Attribute VB_Name = "TRPT401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscaImovel_Click()
      AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo trata
    Dim CodRelatorio As Integer
    Dim Item As Object
    
    For Each Item In grdRelatorios.ListItems
        If Item.Checked Then
            Screen.MousePointer = vbHourglass
            CodRelatorio = Item
            Set Rpt = New VSRelatorio
                If DefinirArquivo(CodRelatorio) Then
                    DefinirCabecalhoRodape CodRelatorio
                    If DefinirFormulas(CodRelatorio) Then
                        If DefinirSelecao(CodRelatorio) Then
                            Rpt.Titulo = Item.SubItems(1)
                            Rpt.Arvore = False
                            Rpt.Visualizar
                        End If
                    End If
                End If
            Set Rpt = Nothing
            Exit For
        End If
    Next
    Screen.MousePointer = vbNormal
    Exit Sub
trata:
    Avisa Err.Description
    Exit Sub
    Resume
    Screen.MousePointer = vbNormal
End Sub

Private Function DefinirArquivo(CodRelatorio As Integer) As Boolean
    Dim VIEW As String
    'DEFINO O TOP DOS RELATÓRIOS...
    '6 MAIORES INADIMPLENTES POR INSCRICAO - VALORES LANÇADOS
    '7 MAIOES INADIMPLENTES  POR IMOVEL - VALORES LANÇADOS
    '8 MAIOES ADIMPLENTES    POR IMOVEL - VALORES LANÇADOS
    '9 MAIOES ADIMPLENTES    POR INSCRICAO VALORES LANÇADOS
    '10 MAIOES ADIMPLENTES    POR INSCRICAO - VALORES ARRECADADOS
    
    If CodRelatorio = 1 Or CodRelatorio = 2 Then
        DefinirArquivo = Rpt.DefinirArquivo(Bdados, App.Path + "\Obrigacao_Resumo.rpt")
        Exit Function
    ElseIf CodRelatorio = 6 Then
        Call Bdados.Executa("DROP VIEW VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_ABERTO")
            VIEW = " CREATE VIEW  VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_ABERTO AS"
            If txtTop <> "" Then
                VIEW = VIEW & " SELECT top " & txtTop & " TOC_INSCRICAO AS Contribuinte,TAB_CONTRIBUINTE.TCI_NOME as Razao,TOC_TIP_COD_IMPOSTO as Tributo,"
            Else
                VIEW = VIEW & " SELECT TOC_INSCRICAO AS Contribuinte,TAB_CONTRIBUINTE.TCI_NOME as Razao,TOC_TIP_COD_IMPOSTO as Tributo,"
            End If
            VIEW = VIEW & " SUM(TOTAL) As TOTAL"
            VIEW = VIEW & " From VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS, TAB_CONTRIBUINTE"
            VIEW = VIEW & " WHERE TAB_CONTRIBUINTE.tci_im = TOC_INSCRICAO  and (TOC_STATUS_OBRIGACAO  = 2 OR TOC_STATUS_OBRIGACAO  = 4 OR"
            VIEW = VIEW & " TOC_STATUS_OBRIGACAO  = 5 OR"
            VIEW = VIEW & " TOC_STATUS_OBRIGACAO  = 9)"
            
            If cboSiglaTributo.ListIndex >= 0 Then
                VIEW = VIEW & " and TOC_TIP_COD_IMPOSTO  = '" & cboSiglaTributo.Coluna(1).Valor & "'"
            End If
            
            If Trim$(txtPeriodoInicial) <> "" Then
                VIEW = VIEW & " and Right(Toc_Periodo,4) >= " & txtPeriodoInicial
            End If
            If Trim$(txtPeriodoFinal) <> "" Then
                VIEW = VIEW & " and Right(Toc_Periodo,4) <= " & txtPeriodoFinal
            End If
            
            If txtValorInicial <> "" Then
                VIEW = VIEW & " and TOTAL >= " & Bdados.Converte(txtValorInicial, TCMonetario)
            End If
            
            If txtValorFinal <> "" Then
                VIEW = VIEW & " and TOTAL <= " & Bdados.Converte(txtValorFinal, TCMonetario)
            End If
            
            If cboTipoLogra.ListIndex >= 0 Then
                VIEW = VIEW & " AND tci_logradouro LIKE '%" & cboTipoLogra.Text & "%'"
            End If
            
            If cboLogra.ListIndex >= 0 Then
                VIEW = VIEW & " AND tci_nome_logradouro LIKE '%" & cboLogra.Text & "%'"
            End If
            
            If cboBairroNovo.ListIndex >= 0 Then
                VIEW = VIEW & " AND tci_bairro LIKE '%" & cboBairroNovo.Text & "%'"
            End If
            
            VIEW = VIEW & " GROUP BY TOC_INSCRICAO,TOC_TIPO_INSCRICAO,"
            If txtTop <> "" Then
                VIEW = VIEW & " TOC_TIP_COD_IMPOSTO , TAB_CONTRIBUINTE.TCI_NOME order by 4 desc"
            Else
                VIEW = VIEW & " TOC_TIP_COD_IMPOSTO , TAB_CONTRIBUINTE.TCI_NOME"
            End If
            If Not Bdados.Executa(VIEW) Then
                Avisa "Erro ao criar view VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_ABERTO"
            End If

    ElseIf CodRelatorio = 7 Then
        Call Bdados.Executa("DROP VIEW VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_ABERTO")
            VIEW = " CREATE VIEW VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_ABERTO AS"
            If txtTop <> "" Then
                VIEW = VIEW & " SELECT top  " & txtTop & " TOC_INSCRICAO AS Contribuinte,TAB_CONTRIBUINTE.TCI_NOME AS Razao,TOC_TIP_COD_IMPOSTO as Tributo,"
            Else
                VIEW = VIEW & " SELECT TOC_INSCRICAO AS Contribuinte,TAB_CONTRIBUINTE.TCI_NOME AS Razao,TOC_TIP_COD_IMPOSTO as Tributo,"
            End If
                     
            VIEW = VIEW & " SUM(TOTAL) As TOTAL"
            VIEW = VIEW & " From VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS, TAB_CONTRIBUINTE,VIS_IMOVEL"
            VIEW = VIEW & " Where TOC_TIPO_INSCRICAO = 1 AND tim_ic = TOC_INSCRICAO"
            VIEW = VIEW & " AND VIS_IMOVEL.TIM_TCI_IM = TCI_IM"
            VIEW = VIEW & " AND (TOC_STATUS_OBRIGACAO  = 2 OR  TOC_STATUS_OBRIGACAO  = 4 OR TOC_STATUS_OBRIGACAO  = 5 OR TOC_STATUS_OBRIGACAO  = 9)"
            
            If cboSiglaTributo.ListIndex >= 0 Then
                VIEW = VIEW & " and TOC_TIP_COD_IMPOSTO  = '" & cboSiglaTributo.Coluna(1).Valor & "'"
            End If
            
            If Trim$(txtPeriodoInicial) <> "" Then
                VIEW = VIEW & " and Toc_Periodo >= " & txtPeriodoInicial
            End If
            If Trim$(txtPeriodoFinal) <> "" Then
                VIEW = VIEW & " and Toc_Periodo <= " & txtPeriodoFinal
            End If
            
            If txtValorInicial <> "" Then
                VIEW = VIEW & " and TOTAL >= " & Bdados.Converte(txtValorInicial, TCMonetario)
            End If
            
            If txtValorFinal <> "" Then
                VIEW = VIEW & " and TOTAL <= " & Bdados.Converte(txtValorFinal, TCMonetario)
            End If
            
            If cboTipoLogra.ListIndex >= 0 Then
                VIEW = VIEW & " AND TTL_NOME LIKE '%" & cboTipoLogra.Text & "%'"
            End If
        
            If cboLogra.ListIndex >= 0 Then
                VIEW = VIEW & " AND tlg_nome LIKE '%" & cboLogra.Text & "%'"
            End If
            If cboBairroNovo.ListIndex >= 0 Then
                VIEW = VIEW & " and  TBA_NOME like '%" & cboBairroNovo.Text & "%'"
            End If
            
            VIEW = VIEW & " GROUP BY TOC_INSCRICAO , TOC_TIPO_INSCRICAO, TOC_TIP_COD_IMPOSTO, VIS_IMOVEL.TIM_TCI_IM, Toc_Periodo, TAB_CONTRIBUINTE.TCI_NOME"
            If txtTop <> "" Then
                VIEW = VIEW & "  order by 4 desc"
            End If
            If Not Bdados.Executa(VIEW) Then
                Avisa "Erro ao criar view VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_ABERTO"
            End If
        
    ElseIf CodRelatorio = 8 Then
        Call Bdados.Executa("DROP VIEW VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_PAGO")
         VIEW = "CREATE VIEW VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_PAGO AS"
         If txtTop <> "" Then
             VIEW = VIEW & " SELECT TOP " & txtTop & " TOC_INSCRICAO AS Contribuinte,TAB_CONTRIBUINTE.TCI_NOME  AS Razao,TOC_TIP_COD_IMPOSTO as Tributo,"
        Else
             VIEW = VIEW & " SELECT  TOC_INSCRICAO AS Contribuinte,TAB_CONTRIBUINTE.TCI_NOME  AS Razao,TOC_TIP_COD_IMPOSTO as Tributo,"
        End If
         VIEW = VIEW & " SUM(TOTAL) As TOTAL"
         VIEW = VIEW & " From VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS, TAB_CONTRIBUINTE,VIS_IMOVEL"
         VIEW = VIEW & " Where TOC_TIPO_INSCRICAO = 1 AND tim_ic = TOC_INSCRICAO"
         VIEW = VIEW & " AND vis_imovel.TIM_TCI_IM = TCI_IM"
         VIEW = VIEW & " AND TOC_STATUS_OBRIGACAO  = 3"
         
         If cboSiglaTributo.ListIndex >= 0 Then
             VIEW = VIEW & " and TOC_TIP_COD_IMPOSTO  = '" & cboSiglaTributo.Coluna(1).Valor & "'"
         End If
         
         If Trim$(txtPeriodoInicial) <> "" Then
             VIEW = VIEW & " and Toc_Periodo >= " & txtPeriodoInicial
         End If
         If Trim$(txtPeriodoFinal) <> "" Then
             VIEW = VIEW & " and Toc_Periodo <= " & txtPeriodoFinal
         End If
            
         
         If Trim$(txtPeriodoInicial) <> "" Then
             VIEW = VIEW & " and Toc_Periodo >= " & txtPeriodoInicial
         End If
         
         If Trim$(txtPeriodoFinal) <> "" Then
             VIEW = VIEW & " and Toc_Periodo <= " & txtPeriodoFinal
         End If
        
         If txtValorInicial <> "" Then
             VIEW = VIEW & " and TOTAL >= " & Bdados.Converte(txtValorInicial, TCMonetario)
         End If
        
         If txtValorFinal <> "" Then
             VIEW = VIEW & " and TOTAL <= " & Bdados.Converte(txtValorFinal, TCMonetario)
         End If
        
         If cboTipoLogra.ListIndex >= 0 Then
             VIEW = VIEW & " AND TTL_NOME LIKE '%" & cboTipoLogra.Text & "%'"
         End If
    
         If cboLogra.ListIndex >= 0 Then
             VIEW = VIEW & " AND tlg_nome LIKE '%" & cboLogra.Text & "%'"
         End If
         
         If cboBairroNovo.ListIndex >= 0 Then
             VIEW = VIEW & " and  TBA_NOME like '%" & cboBairroNovo.Text & "%'"
         End If
        
         VIEW = VIEW & " GROUP BY TOC_INSCRICAO , TOC_TIPO_INSCRICAO, TOC_TIP_COD_IMPOSTO, vis_imovel.TIM_TCI_IM, Toc_Periodo, TAB_CONTRIBUINTE.TCI_NOME  "
         If txtTop <> "" Then
            VIEW = VIEW & "  order by 4 desc"
         End If
        If Not Bdados.Executa(VIEW) Then
            Avisa "Erro ao criar view VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS"
        End If
    ElseIf CodRelatorio = 9 Then
        Call Bdados.Executa("DROP VIEW VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_PAGO")
            VIEW = "CREATE VIEW VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_PAGO AS"
            If txtTop <> "" Then
                VIEW = VIEW & " SELECT top " & txtTop & " TOC_INSCRICAO AS Contribuinte,TAB_CONTRIBUINTE.TCI_NOME as Razao,TOC_TIP_COD_IMPOSTO as Tributo,"
            Else
                VIEW = VIEW & " SELECT  TOC_INSCRICAO AS Contribuinte,TAB_CONTRIBUINTE.TCI_NOME as Razao,TOC_TIP_COD_IMPOSTO as Tributo,"
            End If
            VIEW = VIEW & " SUM(TOTAL) As TOTAL"
            VIEW = VIEW & " From VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS,TAB_CONTRIBUINTE"
            VIEW = VIEW & " Where TAB_CONTRIBUINTE.tci_im = TOC_INSCRICAO and  TOC_TIPO_INSCRICAO = 2"
            VIEW = VIEW & " AND TOC_STATUS_OBRIGACAO  = 3"
            
            If cboSiglaTributo.ListIndex >= 0 Then
                VIEW = VIEW & " and TOC_TIP_COD_IMPOSTO  = '" & cboSiglaTributo.Coluna(1).Valor & "'"
            End If
            
            If Trim$(txtPeriodoInicial) <> "" Then
                VIEW = VIEW & " and Right(Toc_Periodo,4) >= " & txtPeriodoInicial
            End If
            If Trim$(txtPeriodoFinal) <> "" Then
                VIEW = VIEW & " and Right(Toc_Periodo,4) <= " & txtPeriodoFinal
            End If
            
            If txtValorInicial <> "" Then
                 VIEW = VIEW & " and TOTAL >= " & Bdados.Converte(txtValorInicial, TCMonetario)
            End If
        
            If txtValorFinal <> "" Then
                 VIEW = VIEW & " and TOTAL <= " & Bdados.Converte(txtValorFinal, TCMonetario)
            End If
        
           If cboTipoLogra.ListIndex >= 0 Then
               VIEW = VIEW & " AND tci_logradouro LIKE '%" & cboTipoLogra.Text & "%'"
           End If
        
           If cboLogra.ListIndex >= 0 Then
               VIEW = VIEW & " AND tci_nome_logradouro LIKE '%" & cboLogra.Text & "%'"
           End If
         
           If cboBairroNovo.ListIndex >= 0 Then
               VIEW = VIEW & " AND tci_bairro LIKE '%" & cboBairroNovo.Text & "%'"
           End If
            
            VIEW = VIEW & " GROUP BY TOC_TIP_COD_IMPOSTO,TOC_INSCRICAO,TAB_CONTRIBUINTE.TCI_NOME,Toc_Periodo "
            If txtTop <> "" Then
                VIEW = VIEW & " order by 4 desc"
            End If
        If Not Bdados.Executa(VIEW) Then
            Avisa "Erro ao criar view VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_PAGO"
        End If
    ElseIf CodRelatorio = 10 Then
            Call Bdados.Executa("DROP VIEW VIS_ARRE_CONTRIB_CONTRIBUINTE_STATUS_PAGO")
            VIEW = "CREATE VIEW VIS_ARRE_CONTRIB_CONTRIBUINTE_STATUS_PAGO AS "
            If txtTop <> "" Then
                VIEW = VIEW & " SELECT top " & txtTop & " TDR_INSCRICAO AS CONTRIBUINTE,"
                VIEW = VIEW & " TCI_NOME AS NOME,"
                VIEW = VIEW & " TDR_TIP_COD_IMPOSTO AS TRIBUTO,"
            Else
                VIEW = VIEW & " SELECT  TDR_INSCRICAO AS CONTRIBUINTE,"
                VIEW = VIEW & " TCI_NOME AS NOME,"
                VIEW = VIEW & " TDR_TIP_COD_IMPOSTO AS TRIBUTO,"
            End If
            VIEW = VIEW & " Sum (TDR_VALOR_REAL_PAGO) AS TOTAL"
            VIEW = VIEW & " From TAB_DARM_RECEBIDO, TAB_CONTRIBUINTE" 'FROM
            VIEW = VIEW & " Where TDR_TIPO_INSCRICAO = 2 And TCI_IM = TDR_INSCRICAO" ' WHERE
            
            
            If cboSiglaTributo.ListIndex >= 0 Then
                VIEW = VIEW & " and TDR_TIP_COD_IMPOSTO  = '" & cboSiglaTributo.Coluna(1).Valor & "'"
            End If
                     
            If txtDtInicialArrecadacao <> "" And txtDtFinalArrecadacao <> "" Then
                VIEW = VIEW & " and TDR_DATA_PAGAMENTO >=  " & Bdados.Converte(txtDtInicialArrecadacao, TCDataHora) & " aNd TDR_DATA_PAGAMENTO <= " & Bdados.Converte(txtDtFinalArrecadacao, TCDataHora)
            ElseIf txtDtInicialArrecadacao <> "" And txtDtFinalArrecadacao = "" Then
                VIEW = VIEW & " and TDR_DATA_PAGAMENTO >=  " & Bdados.Converte(txtDtInicialArrecadacao, TCDataHora) & " aNd TDR_DATA_PAGAMENTO <= " & Bdados.Converte(txtDtInicialArrecadacao, TCDataHora)
            End If
                     
            If Trim$(txtPeriodoInicial) <> "" Then
                VIEW = VIEW & " and TDR_PERIODO >= " & txtPeriodoInicial
            End If
            If Trim$(txtPeriodoFinal) <> "" Then
                VIEW = VIEW & " and TDR_PERIODO <= " & txtPeriodoFinal
            End If
            
            If txtValorInicial <> "" Then
                VIEW = VIEW & " and TDR_VALOR_REAL_PAGO >= " & Bdados.Converte(txtValorInicial, TCMonetario)
            End If
        
            If txtValorFinal <> "" Then
                 VIEW = VIEW & " and TDR_VALOR_REAL_PAGO <= " & Bdados.Converte(txtValorFinal, TCMonetario)
            End If
        
           If cboTipoLogra.ListIndex >= 0 Then
               VIEW = VIEW & " AND tci_logradouro LIKE '%" & cboTipoLogra.Text & "%'"
           End If
        
           If cboLogra.ListIndex >= 0 Then
               VIEW = VIEW & " AND tci_nome_logradouro LIKE '%" & cboLogra.Text & "%'"
           End If
         
           If cboBairroNovo.ListIndex >= 0 Then
               VIEW = VIEW & " AND tci_bairro LIKE '%" & cboBairroNovo.Text & "%'"
           End If
            
           VIEW = VIEW & " GROUP BY TDR_INSCRICAO,TCI_NOME,TDR_TIP_COD_IMPOSTO,TDR_DATA_PAGAMENTO" 'GROUP BY
           If txtTop <> "" Then
                VIEW = VIEW & " order by 4 desc"
           End If
        If Not Bdados.Executa(VIEW) Then
            Avisa "Erro ao criar view VIEW VIS_ARRE_CONTRIB_CONTRIBUINTE_STATUS_PAGO"
        End If
        
        ElseIf CodRelatorio = 11 Then
            Call Bdados.Executa("DROP VIEW VIS_ARRE_IMOVEL_CONTRIBUINTE_STATUS_PAGO")
            VIEW = "CREATE VIEW  VIS_ARRE_IMOVEL_CONTRIBUINTE_STATUS_PAGO AS "
            If txtTop <> "" Then
                VIEW = VIEW & " SELECT top " & txtTop & " TDR_INSCRICAO AS CONTRIBUINTE,"
                VIEW = VIEW & " TCI_NOME AS NOME,"
                VIEW = VIEW & " TDR_TIP_COD_IMPOSTO AS TRIBUTO,"
            Else
                VIEW = VIEW & " SELECT  TDR_INSCRICAO AS CONTRIBUINTE,"
                VIEW = VIEW & " TCI_NOME AS NOME,"
                VIEW = VIEW & " TDR_TIP_COD_IMPOSTO AS TRIBUTO,"
            End If
            VIEW = VIEW & " Sum (TDR_VALOR_REAL_PAGO) AS TOTAL"
            VIEW = VIEW & " From TAB_DARM_RECEBIDO,VIS_IMOVEL" 'FROM
            VIEW = VIEW & " WHERE TDR_TIPO_INSCRICAO = 1"
            VIEW = VIEW & " AND TIM_IC = TDR_INSCRICAO" ' WHERE
            
            If cboSiglaTributo.ListIndex >= 0 Then
                VIEW = VIEW & " and TDR_TIP_COD_IMPOSTO  = '" & cboSiglaTributo.Coluna(1).Valor & "'"
            End If
                     
            If txtDtInicialArrecadacao <> "" And txtDtFinalArrecadacao <> "" Then
                VIEW = VIEW & " and TDR_DATA_PAGAMENTO >=  " & Bdados.Converte(txtDtInicialArrecadacao, TCDataHora) & " aNd TDR_DATA_PAGAMENTO <= " & Bdados.Converte(txtDtFinalArrecadacao, TCDataHora)
            ElseIf txtDtInicialArrecadacao <> "" And txtDtFinalArrecadacao = "" Then
                VIEW = VIEW & " and TDR_DATA_PAGAMENTO >=  " & Bdados.Converte(txtDtInicialArrecadacao, TCDataHora) & " aNd TDR_DATA_PAGAMENTO <= " & Bdados.Converte(txtDtInicialArrecadacao, TCDataHora)
            End If
                     
            If Trim$(txtPeriodoInicial) <> "" Then
                VIEW = VIEW & " and TDR_PERIODO >= " & txtPeriodoInicial
            End If
            If Trim$(txtPeriodoFinal) <> "" Then
                VIEW = VIEW & " and TDR_PERIODO <= " & txtPeriodoFinal
            End If
            
            If txtValorInicial <> "" Then
                VIEW = VIEW & " and TDR_VALOR_REAL_PAGO >= " & Bdados.Converte(txtValorInicial, TCMonetario)
            End If
        
            If txtValorFinal <> "" Then
                 VIEW = VIEW & " and TDR_VALOR_REAL_PAGO <= " & Bdados.Converte(txtValorFinal, TCMonetario)
            End If
        
           If cboTipoLogra.ListIndex >= 0 Then
               VIEW = VIEW & " AND tci_logradouro LIKE '%" & cboTipoLogra.Text & "%'"
           End If
        
           If cboLogra.ListIndex >= 0 Then
               VIEW = VIEW & " AND tci_nome_logradouro LIKE '%" & cboLogra.Text & "%'"
           End If
         
           If cboBairroNovo.ListIndex >= 0 Then
               VIEW = VIEW & " AND tci_bairro LIKE '%" & cboBairroNovo.Text & "%'"
           End If
            
           VIEW = VIEW & " GROUP BY TDR_INSCRICAO,TCI_NOME,TDR_TIP_COD_IMPOSTO,TDR_DATA_PAGAMENTO" 'GROUP BY
           If txtTop <> "" Then
                VIEW = VIEW & " order by 4 desc"
           End If
        If Not Bdados.Executa(VIEW) Then
            Avisa "Erro ao criar view VIEW VIEW VIS_ARRE_IMOVEL_CONTRIBUINTE_STATUS_PAGO"
        End If
    End If
    DefinirArquivo = Rpt.DefinirArquivo(Bdados, App.Path + "\TRPT401" & CodRelatorio & ".rpt")
End Function

Private Sub DefinirCabecalhoRodape(CodRelatorio As Integer)
    On Error Resume Next
    Select Case CodRelatorio
        Case 1 To 10
            Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            Rpt.Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), "TRPT401." & CodRelatorio, Aplicacoes.Usuario
        Case 11, 12
            Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            Rpt.Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), "TRPT401." & CodRelatorio, Aplicacoes.Usuario
    End Select
End Sub

Private Function DefinirFormulas(CodRelatorio As Integer) As Boolean
    Dim ValorFormula As String
    
    DefinirFormulas = True
    'Rpt.LimparFormulas
    ValorFormula = ""
    
    Select Case CodRelatorio
        Case 1, 2
            Rpt.Formulas "VTTitulo", grdRelatorios.SelectedItem.SubItems(1)
            Rpt.Formulas "VTSubtitulo", Util.ParseString(cboSiglaTributo, " - ", 1)
        Case 6, 7, 8, 9, 10, 11 'IMPOSTOS 100/TODOS MAIORES
            'If cboSiglaTributo = "" Then
            '    Erro "Informe o tributo."
            '    DefinirFormulas = False
            'Else
            Rpt.Formulas "VTTitulo", grdRelatorios.SelectedItem.SubItems(1)
            Rpt.Formulas "VTSubtitulo", Util.ParseString(cboSiglaTributo, " - ", 1)
            If Trim$(txtPeriodoInicial) <> "" Then
                ValorFormula = txtPeriodoInicial
            End If
            If Trim$(txtPeriodoFinal) <> "" Then
                ValorFormula = ValorFormula & IIf(Len(ValorFormula) > 0, " - ", "") & txtPeriodoFinal
            End If
            If ValorFormula = "" Then
                If Trim$(txtDtInicialArrecadacao) <> "" Then
                    ValorFormula = txtDtInicialArrecadacao
                End If
                If Trim$(txtDtFinalArrecadacao) <> "" Then
                    ValorFormula = ValorFormula & IIf(Len(ValorFormula) > 0, " - ", "") & txtDtFinalArrecadacao
                End If
            End If
            Rpt.Formulas "VTPeriodo", ValorFormula
            'End If
            
        
        Case 3, 4, 8, 9 'IPTU 100/TODOS MAIORES
            Rpt.Formulas "VTTitulo", grdRelatorios.SelectedItem.SubItems(1)
         '   Rpt.Formulas "VTSubtitulo", ""
            If Trim$(txtPeriodoInicial) <> "" Then
                ValorFormula = txtPeriodoInicial
            End If
            If Trim$(txtPeriodoFinal) <> "" Then
                ValorFormula = ValorFormula & IIf(Len(ValorFormula) > 0, " - ", "") & txtPeriodoFinal
            End If
            If ValorFormula = "" Then
                If Trim$(txtDtInicialArrecadacao) <> "" Then
                    ValorFormula = txtDtInicialArrecadacao
                End If
                If Trim$(txtDtFinalArrecadacao) <> "" Then
                    ValorFormula = ValorFormula & IIf(Len(ValorFormula) > 0, " - ", "") & txtDtFinalArrecadacao
                End If
            End If
            Rpt.Formulas "VTPeriodo", ValorFormula
            
        Case 5 'ARRECADACAO PROPRIA (CONTABIL)
            Rem Rpt.Formulas "VTTitulo", grdRelatorios.SelectedItem.SubItems(1)
            If Trim$(txtDtInicialArrecadacao) <> "" Then
                Rem Rpt.Formulas "VTSubtitulo", "Exercício " & Year(txtDtInicialArrecadacao)
                ValorFormula = txtDtInicialArrecadacao
            End If
            If Trim$(txtDtFinalArrecadacao) <> "" Then
                ValorFormula = ValorFormula & IIf(Len(ValorFormula) > 0, " - ", "") & txtDtFinalArrecadacao
            End If
            If ValorFormula = "" Then
                If Trim$(txtPeriodoInicial) <> "" Then
                    Rpt.Formulas "VTSubtitulo", "Exercício " & txtPeriodoInicial
                    ValorFormula = txtPeriodoInicial
                End If
                If Trim$(txtPeriodoFinal) <> "" Then
                    ValorFormula = ValorFormula & IIf(Len(ValorFormula) > 0, " - ", "") & txtPeriodoFinal
                End If
            End If
    End Select
End Function

Private Function DefinirSelecao(CodRelatorio As Integer) As Boolean
    Dim Filtro As String
    DefinirSelecao = True
    
    Filtro = ""
    Select Case CodRelatorio
        Case 1 'Inadimplentes
            Filtro = "{VIS_OBRIGACAO_CONTRIBUINTE.TOC_STATUS_OBRIGACAO} = 2 "
            If cboSiglaTributo <> "" Then
                Filtro = Filtro & " AND {VIS_OBRIGACAO_CONTRIBUINTE.TOC_TIP_COD_IMPOSTO} ='" & cboSiglaTributo.Coluna(1).Valor & "'"
            End If
            If Trim$(txtPeriodoInicial) <> "" And Trim$(txtPeriodoFinal) <> "" Then
                Filtro = Filtro & " AND ({VIS_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} >= " & txtPeriodoInicial & " AND {VIS_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} <= " & txtPeriodoFinal & ") "
            End If
        Case 2 'Adimplentes
            Filtro = "{VIS_OBRIGACAO_CONTRIBUINTE.TOC_STATUS_OBRIGACAO} = 3 "
            If cboSiglaTributo <> "" Then
                Filtro = Filtro & " AND {VIS_OBRIGACAO_CONTRIBUINTE.TOC_TIP_COD_IMPOSTO} ='" & cboSiglaTributo.Coluna(1).Valor & "'"
            End If
            If Trim$(txtPeriodoInicial) <> "" And Trim$(txtPeriodoFinal) <> "" Then
                Filtro = Filtro & " AND ({VIS_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} >= " & txtPeriodoInicial & " AND {VIS_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} <= " & txtPeriodoFinal & ") "
            End If
            If Trim$(txtDtInicialArrecadacao) <> "" And Trim$(txtDtFinalArrecadacao) <> "" Then
                Filtro = Filtro & " AND ({Tab_Darm_Recebido.tdr_data_pagamento} >= Date (" & Year(txtDtInicialArrecadacao) & "," & Month(txtDtInicialArrecadacao) & "," & Day(txtDtInicialArrecadacao) & ") AND {Tab_Darm_Recebido.tdr_data_pagamento} <= Date (" & Year(txtDtFinalArrecadacao) & "," & Month(txtDtFinalArrecadacao) & "," & Day(txtDtFinalArrecadacao) & "))"
            End If
        Case 5 'ARRECADACAO PROPRIA (CONTABIL)
            If Trim$(txtDtInicialArrecadacao) <> "" Then
                Filtro = " {Tab_Darm_Recebido.tdr_data_pagamento} >= Date (" & Year(txtDtInicialArrecadacao) & "," & Month(txtDtInicialArrecadacao) & "," & Day(txtDtInicialArrecadacao) & ")"
            End If
            If Trim$(txtDtFinalArrecadacao) <> "" Then
                Filtro = Filtro & IIf(Filtro <> "", " AND ", "") & "{Tab_Darm_Recebido.tdr_data_pagamento} <= Date (" & Year(txtDtFinalArrecadacao) & "," & Month(txtDtFinalArrecadacao) & "," & Day(txtDtFinalArrecadacao) & ")"
            End If
            If cboAgenteArrecadador <> "" Then
                Filtro = Filtro & IIf(Filtro <> "", " AND ", "") & "{TAB_LOTE_PAGAMENTO.TLP_TAR_COD_AGENTE} =" & cboAgenteArrecadador.Coluna(1).Valor
            End If
            If Trim$(txtPeriodoInicial) <> "" Then
                Filtro = Filtro & IIf(Filtro <> "", " AND ", "") & "{Tab_Darm_Recebido.tdr_periodo} >= " & txtPeriodoInicial
            End If
            If Trim$(txtPeriodoFinal) <> "" Then
                Filtro = Filtro & IIf(Filtro <> "", " AND ", "") & "{Tab_Darm_Recebido.tdr_periodo} <= " & txtPeriodoFinal
            End If
        End Select
    
    If Filtro <> "" Then
        Rpt.Selecao = Filtro
    End If
End Function

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me

End Sub

Private Sub cmdPesquisaInscricao_Click()
       AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub PreencherRelatorios()
    Dim Sql As String
    
    Sql = "SELECT TGE_CODIGO AS Codigo, TGE_NOME as Relatorio " & _
        " FROM TAB_GERAL " & _
        " WHERE TGE_CODIGO>0 AND " & _
            " TGE_TIPO = (SELECT TGE_TIPO" & _
                            " FROM TAB_GERAL" & _
                            " WHERE TGE_CODIGO=0 AND" & _
                                " TGE_NOME ='RELATORIOS GERENCIAIS TRPT401')" & _
        " ORDER BY TGE_NOME"
    grdRelatorios.Preencher Bdados, Sql
End Sub

Private Sub Form_Activate()
    tabRelatorios.Tabs(3).Enabled = False
    tabRelatorios.Tabs(4).Enabled = False
    
End Sub

Private Sub Form_Load()
    PreencherRelatorios
    
    PrepararTributo
    PrepararContribuinte
    PrepararImovel
    PrepararLocalizacao
    PrepararArrecadacao
    
End Sub

Private Sub grdRelatorios_Click()
    Dim lbl As Label
    If Not grdRelatorios.SelectedItem Is Nothing Then
        grdRelatorios.MarcarTodos False
        grdRelatorios.SelectedItem.Checked = True
        For Each lbl In lblRelatorio
            lbl = grdRelatorios.SelectedItem.SubItems(1)
        Next
        If grdRelatorios.SelectedItem = 11 Then txtDtFinalArrecadacao = Now
        'fraFiltro.Caption = ":. " & grdRelatorios.SelectedItem.SubItems(1)
    End If
End Sub

Private Sub grdRelatorios_DblClick()
    Dim lbl As Label
    If Not grdRelatorios.SelectedItem Is Nothing Then
        tabRelatorios.Tabs(2).Selected = True
        For Each lbl In lblRelatorio
            lbl = grdRelatorios.SelectedItem.SubItems(1)
        Next
        cboSiglaTributo.SetFocus
        'fraFiltro.Caption = ":. " & grdRelatorios.SelectedItem.SubItems(1)
    End If
End Sub

Private Sub PrepararTributo()
    Dim Sql As String
    
    Sql = "SELECT TIP_SIGLA_IMPOSTO " & Bdados.Concatena & "' - '" & Bdados.Concatena & " TIP_COD_IMPOSTO, TIP_COD_IMPOSTO" & _
        " FROM TAB_IMPOSTO" & _
        " ORDER BY TIP_SIGLA_IMPOSTO"
    cboSiglaTributo.Preencher Bdados, Sql

    cboSituacaoTributo.AddItem ""
    cboSituacaoTributo.AddItem "PAGO"
    cboSituacaoTributo.AddItem "NÃO PAGO"
End Sub

Private Sub PrepararContribuinte()
    Dim Sql As String
    
    Sql = "SELECT DISTINCT(tae_nome) " & _
            " FROM Tab_Atividade_Economica" & _
            " ORDER BY tae_nome"
    cboAtividadeContribuinte.Preencher Bdados, Sql
End Sub

Private Sub PrepararImovel()
    Dim Sql As String, OrderBy As String
    Dim CodGrupo As String
    
    cboAforado.AddItem ""
    cboAforado.AddItem "SIM"
    cboAforado.AddItem "NÃO"
    
    cboTipoImovel.AddItem ""
    cboTipoImovel.AddItem "PREDIAL"
    cboTipoImovel.AddItem "TERRITORIAL"
    
    Sql = "Select tco_descricao_componente " & _
        " From Tab_Componente_Avancado " & _
        " Where tco_grupo = "
    OrderBy = " order by tco_cod_componente asc"
    
    CodGrupo = 1
    cboOcupacaoImovel.Preencher Bdados, Sql & CodGrupo & OrderBy

    CodGrupo = 16
    cboUsoImovel.Preencher Bdados, Sql & CodGrupo & OrderBy

    CodGrupo = 11
    cboDestinacaoImovel.Preencher Bdados, Sql & CodGrupo & OrderBy

    CodGrupo = 12
    cboPadraoImovel.Preencher Bdados, Sql & CodGrupo & OrderBy

    CodGrupo = 9
    cboTipologiaImovel.Preencher Bdados, Sql & CodGrupo & OrderBy

    CodGrupo = 10
    cboEstruturaImovel.Preencher Bdados, Sql & CodGrupo & OrderBy

    CodGrupo = 13
    cboConservacaoImovel.Preencher Bdados, Sql & CodGrupo & OrderBy
End Sub

Private Sub PrepararLocalizacao()
    Dim Sql As String
    
    Sql = "Select DISTINCT(ttl_nome),TTL_COD_TIP_LOGR From Tab_Tipo_Logr"
    cboTipoLogradouro.Preencher Bdados, Sql
    cboTipoLogra.Preencher Bdados, Sql
    
    Sql = "Select DISTINCT(tlg_nome),tlg_cod_logradouro From Tab_Logradouro where tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
    cboLogradouro.Preencher Bdados, Sql
    
    cboLogra.Preencher Bdados, Sql
    
    Sql = "Select DISTINCT(tba_nome),tba_cod_bairro From Tab_Bairro where TBA_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio
    cboBairro.Preencher Bdados, Sql
    
    cboBairroNovo.Preencher Bdados, Sql
End Sub

Private Sub PrepararArrecadacao()
    Dim Sql As String
    
    Sql = "Select tar_nome_agente,tar_cod_agente " & _
        " from tab_agente_arrecadador " & _
        " where tar_ativo =0"
    cboAgenteArrecadador.Preencher Bdados, Sql
End Sub

Private Sub txtIm_LostFocus()
  Dim Ic As String
    If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        If Len(txtIm) = 10 Or Len(txtIm) = 11 Then
            Ic = Imposto.FormataInscricao(txtIm, InscContrib)
        Else
            Ic = txtIm
        End If
    Else
        Ic = txtIm
    End If
    If Trim(txtIm) <> "" Then
        txtIm = BuscaContribuinte(Ic, txtRazao, txtEndereco)
        If Trim(txtIm) = "" Then
            Avisa "Inscricão não encontrada"
            txtIm.SetFocus
        End If
    End If
    
End Sub

Private Sub txtImovel_LostFocus()
 Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtImovel.SetFocus
        End If
    End If
End Sub

