VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TRTT101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRTT101"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   5610
      Left            =   15
      TabIndex        =   5
      Top             =   675
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   9895
      _Version        =   131082
      TabCount        =   3
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
      TagVariant      =   ""
      Tabs            =   "TRTT101.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   5220
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9208
         _Version        =   131082
         TabGuid         =   "TRTT101.frx":00B8
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1215
            Left            =   120
            TabIndex        =   21
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   105
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   2143
            Altura          =   1905
            Caption         =   " Dados do Processo"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Enabled         =   0   'False
            Begin VTOcx.txtVISUAL txtEntradaMensagem 
               Height          =   285
               Left            =   7770
               TabIndex        =   27
               Top             =   375
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   503
               Caption         =   "Dt Entrada"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboSubAcaoMensagem 
               Height          =   315
               Left            =   5085
               TabIndex        =   26
               Top             =   720
               Width           =   5025
               _ExtentX        =   8864
               _ExtentY        =   556
               Caption         =   "SubAção"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Enabled         =   0   'False
            End
            Begin VTOcx.cboVISUAL CboAcaoMensagem 
               Height          =   315
               Left            =   480
               TabIndex        =   25
               Top             =   720
               Width           =   4530
               _ExtentX        =   7990
               _ExtentY        =   556
               Caption         =   "Ação"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Enabled         =   0   'False
            End
            Begin VTOcx.txtVISUAL txtProcesso 
               Height          =   300
               Left            =   150
               TabIndex        =   24
               Top             =   375
               Width           =   2910
               _ExtentX        =   5133
               _ExtentY        =   529
               Caption         =   "Processo"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               Requerido       =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   1755
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   1350
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   3096
            Altura          =   1905
            Caption         =   " Dados do Lançamento"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Enabled         =   0   'False
            Begin VTOcx.txtVISUAL txtDamMensagem 
               Height          =   300
               Left            =   210
               TabIndex        =   39
               Top             =   405
               Width           =   2760
               _ExtentX        =   4868
               _ExtentY        =   529
               Caption         =   "Número DAM"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               Requerido       =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtValor 
               Height          =   300
               Left            =   7860
               TabIndex        =   33
               Tag             =   "Periodo Inicial"
               Top             =   1380
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   529
               Caption         =   "Valor"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               Requerido       =   0   'False
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cboVISUAL CboTributoMensagem 
               Height          =   315
               Left            =   3030
               TabIndex        =   32
               Tag             =   "Tributo"
               Top             =   405
               Width           =   6810
               _ExtentX        =   12012
               _ExtentY        =   556
               Caption         =   "Tributo"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
               Enabled         =   0   'False
            End
            Begin VTOcx.txtVISUAL txtPeriodoMensagem 
               Height          =   300
               Left            =   120
               TabIndex        =   31
               Tag             =   "Periodo Inicial"
               Top             =   1380
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   529
               Caption         =   "Periodo Inicial"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               Requerido       =   0   'False
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtEnderecoMensagem 
               Height          =   300
               Left            =   540
               TabIndex        =   30
               Top             =   1050
               Width           =   9285
               _ExtentX        =   16378
               _ExtentY        =   529
               Caption         =   "Endereço"
               Text            =   ""
               Enabled         =   0   'False
               Requerido       =   0   'False
               AgruparValores  =   0   'False
            End
            Begin VTOcx.cboVISUAL CboStatusMensagem 
               Height          =   315
               Left            =   3000
               TabIndex        =   29
               Tag             =   "Tributo"
               ToolTipText     =   "STATUS OBRIGACAO"
               Top             =   1380
               Width           =   4425
               _ExtentX        =   7805
               _ExtentY        =   556
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
               Enabled         =   0   'False
            End
            Begin VTOcx.txtVISUAL txtRazaoMensagem 
               Height          =   300
               Left            =   240
               TabIndex        =   28
               Top             =   735
               Width           =   9585
               _ExtentX        =   16907
               _ExtentY        =   529
               Caption         =   "Nome/Razão"
               Text            =   ""
               Enabled         =   0   'False
               Requerido       =   0   'False
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   1995
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   3135
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   3519
            Altura          =   1905
            Caption         =   " Restituição"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VB.TextBox txtMotivo 
               Appearance      =   0  'Flat
               Height          =   1125
               Left            =   1395
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   37
               Top             =   735
               Width           =   8400
            End
            Begin VTOcx.cboVISUAL cboTipo 
               Height          =   315
               Left            =   3615
               TabIndex        =   35
               ToolTipText     =   "TIPO RESTITUICAO"
               Top             =   360
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   556
               Caption         =   "Tipo"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtValorRestituicao 
               Height          =   300
               Left            =   6660
               TabIndex        =   36
               Top             =   375
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   529
               Caption         =   "Valor Restituição"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtData 
               Height          =   315
               Left            =   945
               TabIndex        =   34
               Top             =   360
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   556
               Caption         =   "Data"
               Text            =   ""
               Formato         =   0
            End
            Begin VB.Label Label1 
               Caption         =   "Motivo"
               Height          =   285
               Left            =   870
               TabIndex        =   38
               Top             =   675
               Width           =   480
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5220
         Left            =   -99969
         TabIndex        =   7
         Top             =   30
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9208
         _Version        =   131082
         TabGuid         =   "TRTT101.frx":00E0
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   2400
            Left            =   105
            TabIndex        =   40
            Top             =   45
            Width           =   10260
            _ExtentX        =   18098
            _ExtentY        =   4233
            Altura          =   1905
            Caption         =   " Consulta de Lançamentos"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483644
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtImLancamento 
               Height          =   300
               Left            =   600
               TabIndex        =   54
               Top             =   690
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   529
               Caption         =   "Inscricão"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cboVISUAL cboImposto 
               Height          =   315
               Left            =   765
               TabIndex        =   53
               Tag             =   "Tributo"
               Top             =   330
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   556
               Caption         =   "Tributo"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDAM 
               Height          =   300
               Left            =   7110
               TabIndex        =   52
               Top             =   690
               Width           =   2760
               _ExtentX        =   4868
               _ExtentY        =   529
               Caption         =   "Número DAM"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cmdVISUAL cmdVISUAL2 
               Height          =   300
               Left            =   6480
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   690
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   529
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtImovel 
               Height          =   300
               Left            =   3285
               TabIndex        =   50
               Top             =   690
               Width           =   3180
               _ExtentX        =   5609
               _ExtentY        =   529
               Caption         =   "Cadastro do Imóvel"
               Text            =   ""
               Requerido       =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtPeriodoInicial 
               Height          =   300
               Left            =   5550
               TabIndex        =   49
               Tag             =   "Periodo Inicial"
               Top             =   2025
               Width           =   3105
               _ExtentX        =   5477
               _ExtentY        =   529
               Caption         =   "Periodo(dd/mm/aaaa)"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               Requerido       =   0   'False
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtPeriodoFinal 
               Height          =   300
               Left            =   8700
               TabIndex        =   48
               Top             =   2010
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               Caption         =   ""
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               Requerido       =   0   'False
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtExercicioFinal 
               Height          =   300
               Left            =   2640
               TabIndex        =   47
               Tag             =   "Periodo Final"
               Top             =   2040
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   529
               Caption         =   "Periodo Final"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtExercicioInicial 
               Height          =   300
               Left            =   180
               TabIndex        =   46
               Tag             =   "Periodo Inicial"
               Top             =   2040
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   529
               Caption         =   "Periodo Inicial"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
               Height          =   300
               Left            =   2880
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   690
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   529
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtEnderecoLancamento 
               Height          =   300
               Left            =   585
               TabIndex        =   44
               Top             =   1350
               Width           =   9285
               _ExtentX        =   16378
               _ExtentY        =   529
               Caption         =   "Endereço"
               Text            =   ""
               Enabled         =   0   'False
               Requerido       =   0   'False
               AgruparValores  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboStatus 
               Height          =   315
               Left            =   5475
               TabIndex        =   43
               Tag             =   "Tributo"
               ToolTipText     =   "STATUS OBRIGACAO"
               Top             =   1680
               Width           =   4410
               _ExtentX        =   7779
               _ExtentY        =   556
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
            End
            Begin VTOcx.cboVISUAL cboRestricao 
               Height          =   315
               Left            =   585
               TabIndex        =   42
               Tag             =   "Tributo"
               ToolTipText     =   "RESTRICAO DAM"
               Top             =   1680
               Width           =   4875
               _ExtentX        =   8599
               _ExtentY        =   556
               Caption         =   "Restrição"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
            End
            Begin VTOcx.txtVISUAL txtRazao 
               Height          =   300
               Left            =   285
               TabIndex        =   41
               Top             =   1020
               Width           =   9585
               _ExtentX        =   16907
               _ExtentY        =   529
               Caption         =   "Nome/Razão"
               Text            =   ""
               Enabled         =   0   'False
               Requerido       =   0   'False
            End
         End
         Begin VTOcx.grdVISUAL lstObrig 
            Height          =   2700
            Left            =   90
            TabIndex        =   20
            Top             =   2475
            Width           =   10290
            _ExtentX        =   18150
            _ExtentY        =   4763
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
            CheckBox        =   -1  'True
            MarcaUnico      =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5220
         Left            =   -99969
         TabIndex        =   6
         Top             =   30
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9208
         _Version        =   131082
         TabGuid         =   "TRTT101.frx":0108
         Begin VTOcx.fraVISUAL fraProPrietario 
            Height          =   1860
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   105
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   3281
            Altura          =   1905
            Caption         =   " Consulta de Processos"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtNomeContrib 
               Height          =   285
               Left            =   3285
               TabIndex        =   17
               Top             =   390
               Width           =   6600
               _ExtentX        =   11642
               _ExtentY        =   503
               Caption         =   ""
               Text            =   ""
               CorRotulo       =   0
               CorTexto        =   4194304
            End
            Begin VTOcx.txtVISUAL txtIm 
               Height          =   285
               Left            =   210
               TabIndex        =   16
               Tag             =   "Insc. Municipal"
               Top             =   390
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   503
               Caption         =   "Insc. Municipal"
               Text            =   ""
               Restricao       =   2
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtEndereco 
               Height          =   300
               Left            =   690
               TabIndex        =   15
               Top             =   705
               Width           =   9195
               _ExtentX        =   16219
               _ExtentY        =   529
               Caption         =   "Endereço"
               Text            =   ""
               Enabled         =   0   'False
               Requerido       =   0   'False
               CorRotulo       =   0
               CorTexto        =   4194304
            End
            Begin VTOcx.cmdVISUAL cmdBUsca 
               Height          =   285
               Left            =   2955
               TabIndex        =   14
               Top             =   390
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Caption         =   ""
               Acao            =   5
               CorBorda        =   32768
               CorFrente       =   16384
               CorFoco         =   14737632
            End
            Begin VTOcx.cboVISUAL cboAcao 
               Height          =   315
               Left            =   1065
               TabIndex        =   13
               Top             =   1035
               Visible         =   0   'False
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   556
               Caption         =   "Ação"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDataEntrada 
               Height          =   285
               Left            =   5130
               TabIndex        =   12
               Top             =   1050
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   503
               Caption         =   "Dt Entrada"
               Text            =   ""
               Formato         =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtEntrega 
               Height          =   285
               Left            =   7560
               TabIndex        =   11
               Top             =   1035
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   503
               Caption         =   "Dt Entrega"
               Text            =   ""
               Formato         =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboSubAcao 
               Height          =   315
               Left            =   735
               TabIndex        =   10
               Top             =   1395
               Visible         =   0   'False
               Width           =   9150
               _ExtentX        =   16140
               _ExtentY        =   556
               Caption         =   "SubAção"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
         End
         Begin VTOcx.grdVISUAL grdDados 
            Height          =   2895
            Left            =   105
            TabIndex        =   18
            Top             =   2010
            Width           =   10245
            _ExtentX        =   18071
            _ExtentY        =   5106
            CorBorda        =   32768
            Caption         =   "Processos"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
            CheckBox        =   -1  'True
            MarcaUnico      =   -1  'True
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   1138
      Icone           =   "TRTT101.frx":0130
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   6345
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   6675
         TabIndex        =   19
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL CmdGravar 
         Height          =   375
         Left            =   7665
         TabIndex        =   4
         Top             =   90
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Gravar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9555
         TabIndex        =   3
         Top             =   90
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8610
         TabIndex        =   2
         Top             =   90
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
End
Attribute VB_Name = "TRTT101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obrig As New Obrigacao

Private Sub cboAcao_Click()
    Dim Sql As String
    
    Sql = " SELECT  TPP_NOME_PARAMETRO ,TPP_CODIGO_PARAMETRO"
    Sql = Sql & " From TAB_PARAMETRO_PROTOCOLO"
    Sql = Sql & " Where (TPP_CODIGO_PARAMETRO <> 0)"
    Sql = Sql & " and  tpp_tipo_parametro =  " & cboAcao.Coluna(1).Valor
    Sql = Sql & " ORDER BY TPP_CODIGO_PARAMETRO "
    
    cboSubAcao.Preencher Bdados, Sql
End Sub

Private Sub CboAcaoMensagem_Click()
    Dim Sql As String
    
    Sql = " SELECT  TPP_NOME_PARAMETRO ,TPP_CODIGO_PARAMETRO"
    Sql = Sql & " From TAB_PARAMETRO_PROTOCOLO"
    Sql = Sql & " Where (TPP_CODIGO_PARAMETRO <> 0)"
    Sql = Sql & " and  tpp_tipo_parametro =  " & CboAcaoMensagem.Coluna(1).Valor
    Sql = Sql & " ORDER BY TPP_CODIGO_PARAMETRO "
    
    cboSubAcaoMensagem.Preencher Bdados, Sql

End Sub

Private Sub cmdBUsca_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdBuscar_Click()
    Dim Sql As String
    
    Dim Inscri As String


    If TabDados.Tabs(1).Selected Then
        Sql = " SELECT TAB_PROTOCOLO.TPR_PROCESSO AS Processo, TAB_PARAMETRO_PROTOCOLO.TPP_NOME_PARAMETRO AS Ação,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_REQUERENTE AS Insc_Municiapal, TAB_PROTOCOLO.TPR_NOME_REQUERENTE AS Requerente,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_ENDERECO AS Endereço, TAB_PROTOCOLO.TPR_DATA_ENTRADA AS Data_Entrada,"
        Sql = Sql & " VIS_STATUS_PROTOCOLO.TGE_NOME AS Status, VIS_STATUS_FINAL_PROTOCOLO.TGE_NOME AS Homologação,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_DATA_HOMOLOGACAO AS Data_Homologação, TAB_PROTOCOLO.TPR_USUARIO_HOMOLOGACAO AS Func_Homologação,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_MOTIVO_HOMOLOGACAO AS Obs_Homologação, TAB_PROTOCOLO.TPR_DATA_ENTREGA AS Data_Entrega,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_TUS_USUARIO AS Usuário, TAB_PROTOCOLO.TPR_RESPONSAVEL AS Responsável,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_ASSUNTO AS Assunto, TAB_PROTOCOLO.TPR_ACAO AS Cod_Ação, TAB_PROTOCOLO.TPR_SUBACAO AS Cod_SubAção"
        Sql = Sql & " FROM TAB_PROTOCOLO LEFT OUTER JOIN"
        Sql = Sql & " VIS_STATUS_PROTOCOLO ON TAB_PROTOCOLO.TPR_STATUS = VIS_STATUS_PROTOCOLO.TGE_CODIGO LEFT OUTER JOIN"
        Sql = Sql & " TAB_PARAMETRO_PROTOCOLO ON TAB_PROTOCOLO.TPR_ACAO = TAB_PARAMETRO_PROTOCOLO.TPP_TIPO_PARAMETRO AND"
        Sql = Sql & " TAB_PROTOCOLO.TPR_SUBACAO = TAB_PARAMETRO_PROTOCOLO.TPP_CODIGO_PARAMETRO LEFT OUTER JOIN"
        Sql = Sql & " VIS_STATUS_FINAL_PROTOCOLO ON TAB_PROTOCOLO.TPR_STATUS_HOMOLOGACAO = VIS_STATUS_FINAL_PROTOCOLO.TGE_CODIGO"
        Sql = Sql & " Where (1 = 1)"
        
        If txtIm <> "" Then Sql = Sql & " and TPR_REQUERENTE = " & txtIm
        If txtDataEntrada <> "" Then Sql = Sql & " and TPR_DATA_ENTRADA = " & Bdados.Converte(txtDataEntrada, TCDataHora)
        If txtEntrega <> "" Then Sql = Sql & " and TPR_DATA_ENTREGA = " & Bdados.Converte(txtEntrega, TCDataHora)
        If cboAcao.ListIndex <> -1 Then Sql = Sql & " and TPR_ACAO = " & cboAcao.Coluna(1).Valor
        If cboSubAcao.ListIndex <> -1 Then Sql = Sql & " and TPR_SUBACAO = " & cboSubAcao.Coluna(1).Valor
        If txtNomeContrib <> "" Then Sql = Sql & " AND TPR_NOME_REQUERENTE LIKE '%" & txtNomeContrib & "%'"
        Sql = Sql & " order  by TPR_PROCESSO"
        If Not grdDados.Preencher(Bdados, Sql) Then Avisa "Busca sem resultados "
    ElseIf TabDados.Tabs(2).Selected Then
        Inscri = txtImLancamento
        If Not Obrig.MostraObrigacaoGerada(lstObrig, CStr(cboImposto.Coluna(0).Valor), Inscri, _
                CInt(cboRestricao.Coluna(1).Valor), CInt(cboStatus.Coluna(1).Valor), txtPeriodoInicial, txtPeriodoFinal, _
                txtExercicioInicial, txtExercicioFinal, , txtImovel, , IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False), txtDam) Then
                Avisa "Nenhum registro encontrado."
             cboImposto.SetFocus
        End If
    End If
End Sub

Private Sub CmdGravar_Click()
    Dim Valores         As String
    Dim Campos          As String
    Dim Condicao        As String
    Dim NUMERO          As String
    Dim CONTA As New ContaCorrente
    NUMERO = CONTA.GeraCodPagamento(51)
        
    Campos = "TRT_NUMERO,TRT_TOC_COD_OBRIGACAO,TRT_TPR_PROTOCOLO,TRT_DATA,TRT_VALOR_RESTITUIDO,TRT_MOTIVO,TRT_TUS_COD_USUARIO,TRT_TIPO"
    Valores = Bdados.PreparaValor(Bdados.Converte(NUMERO, tctexto), Bdados.Converte(txtDamMensagem, tctexto), Bdados.Converte(txtProcesso, tctexto), Bdados.Converte(txtData, TCDataHora), Bdados.Converte(txtValorRestituicao, TCMonetario), txtMotivo, AplicacoesVTFuncoes.Usuario, cboTipo.Coluna(1).Valor)
    
    If Bdados.InsereDados("TAB_RESTITUICAO", Valores, Campos) Then
        If cboTipo.Coluna(1).Valor = 1 Then
            'ALTERO NA TAB_OBRIGACAO_CONTRIBUINTE
            Bdados.GravaDados "TAB_OBRIGACAO_CONTRIBUINTE", etsCreditoOriginalAberto, "TOC_STATUS_OBRIGACAO", "TOC_COD_OBRIGACAO = '" & txtDamMensagem & "'"
        End If
        cmdLimpar_Click
        Avisa "Lançamento restituido com sucesso."
    End If
    
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtImLancamento
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL2_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     
    Obrig.PreencheComboTributo cboImposto, False, etcTodos
    Obrig.PreencheComboTributo CboTributoMensagem, False, etcTodos
    
    cboStatus.PreencherGeral Bdados, cboStatus.ToolTipText
    CboStatusMensagem.PreencherGeral Bdados, cboStatus.ToolTipText
    
    cboRestricao.PreencherGeral Bdados, cboRestricao.ToolTipText
    cboAcao.Preencher Bdados, "SELECT TPP_NOME_PARAMETRO , TPP_TIPO_PARAMETRO FROM TAB_PARAMETRO_PROTOCOLO where TPP_CODIGO_PARAMETRO = 0  ORDER BY TPP_TIPO_PARAMETRO"
    CboAcaoMensagem.Preencher Bdados, "SELECT TPP_NOME_PARAMETRO , TPP_TIPO_PARAMETRO FROM TAB_PARAMETRO_PROTOCOLO where TPP_CODIGO_PARAMETRO = 0  ORDER BY TPP_TIPO_PARAMETRO"
    
    cboTipo.PreencherGeral Bdados, cboTipo.ToolTipText
End Sub

Private Sub grdDados_ItemCheck(ByVal Item As MSComctlLib.IListItem)
    If Item.Checked Then
        txtProcesso = grdDados.SelectedItem
        CboAcaoMensagem.SetarLinha grdDados.SelectedItem.SubItems(15), 1
        CboAcaoMensagem_Click
        cboSubAcaoMensagem.SetarLinha grdDados.SelectedItem.SubItems(16), 1
        txtEntradaMensagem = grdDados.SelectedItem.SubItems(5)
    End If
End Sub

Private Sub lstObrig_ItemCheck(ByVal Item As MSComctlLib.IListItem)
    If Item.Checked Then
        CboTributoMensagem.SetarLinha Imposto.BuscaCodImposto(lstObrig.SelectedItem.SubItems(3))
        txtRazaoMensagem = txtRazao
        txtEnderecoMensagem = txtEnderecoLancamento
        CboStatusMensagem = lstObrig.SelectedItem.SubItems(9)
        txtPeriodoMensagem = lstObrig.SelectedItem.SubItems(4)
        txtValor = Format(lstObrig.SelectedItem.SubItems(6), Const_Monetario)
        txtDamMensagem = lstObrig.SelectedItem
    End If
End Sub

Private Sub txtImLancamento_LostFocus()
    Dim Ic As String
    If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        If Len(txtImLancamento) = 10 Or Len(txtImLancamento) = 11 Then
            Ic = Imposto.FormataInscricao(txtImLancamento, InscContrib)
        Else
            Ic = txtImLancamento
        End If
    Else
        Ic = txtImLancamento
    End If
    If Trim(txtImLancamento) <> "" Then
        txtImLancamento = BuscaContribuinte(Ic, txtRazao, txtEnderecoLancamento)
        If Trim(txtImLancamento) = "" Then
            Avisa "Inscricão não encontrada"
            txtImLancamento.SetFocus
        End If
    End If
End Sub

Private Sub txtImovel_LostFocus()
    Dim Ic As String
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEnderecoLancamento, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtImovel.SetFocus
        End If
    End If
End Sub

Private Sub txtValorRestituicao_LostFocus()
'    If cboTipo.Coluna(1).Valor = 1 Then   'INTEGRAL THEN
'        If Nvl(txtValorRestituicao, 0) <> Nvl(txtValor, 0) Then
'            Avisa "O Valor da restituição não pode ser diferente do valor lançado."
'            txtValorRestituicao.SetFocus
'        End If
'    ElseIf cboTipo.Coluna(1).Valor = 2 Then 'PARCIAL
'        If Nvl(txtValorRestituicao, 0) > Nvl(txtValor, 0) Then
'            Avisa "O Valor da restituição não pode ser maior que o  valor lançado."
'            txtValorRestituicao.SetFocus
'        End If
'    End If
    
End Sub
