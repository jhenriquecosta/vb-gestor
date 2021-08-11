VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TRPT402A 
   BackColor       =   &H80000016&
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs tabRelatorios 
      Height          =   6060
      Left            =   30
      TabIndex        =   36
      Top             =   690
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   10689
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
      TagVariant      =   ""
      Tabs            =   "TRPT402A.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5670
         Index           =   1
         Left            =   -99969
         TabIndex        =   37
         Top             =   30
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   10001
         _Version        =   131082
         TabGuid         =   "TRPT402A.frx":007E
         Begin VTOcx.grdVISUAL grdRelatorios 
            Height          =   5790
            Left            =   60
            TabIndex        =   0
            Top             =   60
            Width           =   10155
            _ExtentX        =   17912
            _ExtentY        =   4339
            CorFundo        =   -2147483633
            Caption         =   "Relatórios Operacionais"
            CorTitulo       =   32768
            CorCaption      =   16777215
            OcultarRodape   =   -1  'True
            MarcaUnico      =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5670
         Index           =   0
         Left            =   30
         TabIndex        =   38
         Top             =   30
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   10001
         _Version        =   131082
         TabGuid         =   "TRPT402A.frx":00A6
         Begin VTOcx.fraVISUAL fraFiltro 
            Height          =   5475
            Left            =   90
            TabIndex        =   39
            Top             =   90
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   9657
            Altura          =   1905
            Caption         =   " "
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483626
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtPeriodoFinal 
               Height          =   315
               Left            =   870
               TabIndex        =   3
               Tag             =   "Data Inicial"
               Top             =   1290
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   556
               Caption         =   "Per. Final"
               Text            =   ""
               Restricao       =   2
               MaxLen          =   6
            End
            Begin VTOcx.cboVISUAL cboAgenteArrecadador 
               Height          =   315
               Left            =   6810
               TabIndex        =   29
               Top             =   3090
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   556
               Caption         =   "Agente"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtQuadra 
               Height          =   315
               Left            =   8460
               TabIndex        =   28
               Tag             =   "Data Inicial"
               Top             =   2010
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
               Left            =   6810
               TabIndex        =   27
               Tag             =   "Data Inicial"
               Top             =   2010
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
               Left            =   6750
               TabIndex        =   26
               Top             =   1650
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   556
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboLogradouro 
               Height          =   315
               Left            =   6780
               TabIndex        =   25
               Top             =   1290
               Width           =   3285
               _ExtentX        =   5794
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboTipoLogradouro 
               Height          =   315
               Left            =   6750
               TabIndex        =   24
               Top             =   930
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   556
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCodLogradouro 
               Height          =   315
               Left            =   6840
               TabIndex        =   23
               Tag             =   "Data Inicial"
               Top             =   570
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   556
               Caption         =   "Cód. Logr."
               Text            =   ""
               Restricao       =   2
            End
            Begin VTOcx.txtVISUAL txtValorVenalFim 
               Height          =   315
               Left            =   4560
               TabIndex        =   22
               Tag             =   "Data Inicial"
               Top             =   4890
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
               Left            =   3360
               TabIndex        =   21
               Tag             =   "Data Inicial"
               Top             =   4530
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
               Left            =   3240
               TabIndex        =   20
               Top             =   4170
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   556
               Caption         =   "Conservação"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboEstruturaImovel 
               Height          =   315
               Left            =   3240
               TabIndex        =   19
               Top             =   3810
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   556
               Caption         =   "Estrutura"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboTipologiaImovel 
               Height          =   315
               Left            =   3240
               TabIndex        =   18
               Top             =   3450
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   556
               Caption         =   "Tipologia"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboPadraoImovel 
               Height          =   315
               Left            =   3240
               TabIndex        =   17
               Top             =   3090
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   556
               Caption         =   "Padrão"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboDestinacaoImovel 
               Height          =   315
               Left            =   3240
               TabIndex        =   16
               Top             =   2730
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   556
               Caption         =   "Destinação"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboUsoImovel 
               Height          =   315
               Left            =   3240
               TabIndex        =   15
               Top             =   2370
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   556
               Caption         =   "Uso"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboOcupacaoImovel 
               Height          =   315
               Left            =   3240
               TabIndex        =   14
               Top             =   2010
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   556
               Caption         =   "Ocupacao"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboTipoImovel 
               Height          =   315
               Left            =   3240
               TabIndex        =   13
               Top             =   1650
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   556
               Caption         =   "Tipo"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL cboAforado 
               Height          =   315
               Left            =   3990
               TabIndex        =   12
               Top             =   1290
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   556
               Caption         =   "Aforado"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtAnoConstrucao 
               Height          =   315
               Left            =   3300
               TabIndex        =   11
               Tag             =   "Data Inicial"
               Top             =   930
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   556
               Caption         =   "Ano Construção"
               Text            =   ""
               Restricao       =   2
               MaxLen          =   4
               Mascara         =   "0000"
            End
            Begin VTOcx.txtVISUAL txtICImovel 
               Height          =   315
               Left            =   3270
               TabIndex        =   10
               Tag             =   "Data Inicial"
               Top             =   570
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   556
               Caption         =   "IC"
               Text            =   ""
               Restricao       =   2
               MaxLen          =   14
               Mascara         =   "000000000000000"
            End
            Begin VTOcx.cboVISUAL cboSituacaoTributo 
               Height          =   315
               Left            =   630
               TabIndex        =   6
               Top             =   2370
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   556
               Caption         =   "Situação"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtParcela 
               Height          =   315
               Left            =   750
               TabIndex        =   5
               Tag             =   "Data Inicial"
               Top             =   2010
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
               Left            =   420
               TabIndex        =   4
               Tag             =   "Data Inicial"
               Top             =   1650
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               Caption         =   "Documento"
               Text            =   ""
               Restricao       =   2
               MaxLen          =   8
            End
            Begin VTOcx.cboVISUAL cboAtividadeContribuinte 
               Height          =   315
               Left            =   60
               TabIndex        =   9
               Top             =   4110
               Width           =   3105
               _ExtentX        =   5477
               _ExtentY        =   556
               Caption         =   "Atv"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtPeriodoInicial 
               Height          =   315
               Left            =   750
               TabIndex        =   2
               Tag             =   "Data Inicial"
               Top             =   930
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   556
               Caption         =   "Per. Inicial"
               Text            =   ""
               Restricao       =   2
               MaxLen          =   6
            End
            Begin VTOcx.txtVISUAL txtRazaoContribuinte 
               Height          =   315
               Left            =   90
               TabIndex        =   8
               Tag             =   "Data Inicial"
               Top             =   3750
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   556
               Caption         =   "Razão"
               Text            =   ""
               MaxLen          =   11
            End
            Begin VTOcx.txtVISUAL txtIMContribuinte 
               Height          =   315
               Left            =   390
               TabIndex        =   7
               Tag             =   "Data Inicial"
               Top             =   3390
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               Caption         =   "IM"
               Text            =   ""
               Restricao       =   2
               MaxLen          =   11
               Mascara         =   "00000000-00"
            End
            Begin VTOcx.txtVISUAL txtDtInicialArrecadacao 
               Height          =   315
               Left            =   7260
               TabIndex        =   30
               Tag             =   "Data Inicial"
               Top             =   3450
               Width           =   2745
               _ExtentX        =   4842
               _ExtentY        =   556
               Caption         =   "Data Inicial"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               MaxLen          =   10
            End
            Begin VTOcx.txtVISUAL txtDtFinalArrecadacao 
               Height          =   315
               Left            =   7380
               TabIndex        =   31
               Tag             =   "Data Final"
               Top             =   3810
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   556
               Caption         =   "Data Final"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               MaxLen          =   10
            End
            Begin VTOcx.cboVISUAL cboSiglaTributo 
               Height          =   315
               Left            =   90
               TabIndex        =   1
               Top             =   570
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   556
               Caption         =   "Sigla"
               Text            =   ""
               AutoFocaliza    =   0   'False
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
               Left            =   6810
               TabIndex        =   45
               Top             =   2850
               Width           =   1425
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
               Left            =   6810
               TabIndex        =   44
               Top             =   330
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":. Imóvel"
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
               Left            =   3300
               TabIndex        =   42
               Top             =   330
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":. Tributo"
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
               Index           =   2
               Left            =   60
               TabIndex        =   41
               Top             =   330
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":. Contribuinte"
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
               Index           =   1
               Left            =   60
               TabIndex        =   40
               Top             =   3120
               Width           =   1380
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   3  'Dot
               Index           =   1
               X1              =   6720
               X2              =   6720
               Y1              =   420
               Y2              =   5340
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   3  'Dot
               Index           =   0
               X1              =   3180
               X2              =   3180
               Y1              =   420
               Y2              =   5340
            End
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   34
      Top             =   6855
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   900
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7320
         TabIndex        =   43
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
         Left            =   8370
         TabIndex        =   32
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
         TabIndex        =   35
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
      TabIndex        =   33
      Top             =   0
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   1138
      Formulario      =   "TREL402"
      Descricao       =   "Relatórios Gerenciais"
      Icone           =   "TRPT402A.frx":00CE
   End
End
Attribute VB_Name = "TRPT402A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    On Error GoTo trata
    Dim CodRelatorio As Integer
        
    If Not grdRelatorios.SelectedItem Is Nothing Then
        Screen.MousePointer = vbHourglass
        CodRelatorio = grdRelatorios.SelectedItem
        Set Rpt = New VSRelatorio
            If DefinirArquivo(CodRelatorio) Then
                DefinirCabecalhoRodape CodRelatorio
                If DefinirFormulas(CodRelatorio) Then
                    If DefinirSelecao(CodRelatorio) Then
                        Rpt.Titulo = grdRelatorios.SelectedItem.SubItems(1)
                        Rpt.Arvore = False
                        Rpt.Visualizar
                    End If
                End If
            End If
        Set Rpt = Nothing
    End If
trata:
    Screen.MousePointer = vbNormal
End Sub

Private Function DefinirArquivo(CodRelatorio As Integer) As Boolean
    DefinirArquivo = Rpt.DefinirArquivo(Bdados, App.Path + "\TBCP799" & CodRelatorio & ".rpt")
    'Select Case CodRelatorio
     '   Case 1, 2, 3, 4
            'DefinirArquivo = Rpt.DefinirArquivo(Bdados, App.Path + "\TRPT401" & CodRelatorio & ".rpt")
    'End Select
End Function

Private Sub DefinirCabecalhoRodape(CodRelatorio As Integer)
    'Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            '
    'Select Case CodRelatorio
     '   Case 1, 2, 3, 4
      '      Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            'Rpt.Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), "TRPT401." & CodRelatorio, Aplicacoes.Usuario
    'End Select
End Sub

Private Function DefinirFormulas(CodRelatorio As Integer) As Boolean
    DefinirFormulas = True
    Rpt.LimparFormulas
    Rpt.Formulas "CLIENTE", Temp.PegaParametro(Bdados, "CLIENTE")
    Select Case CodRelatorio
        'TAB_GERAL TIPO = 799
        Case 1
                   
    End Select
End Function

Private Function DefinirSelecao(CodRelatorio As Integer) As Boolean
    Dim Filtro As String
    DefinirSelecao = True
    
    Select Case CodRelatorio
        'TAB_GERAL TIPO = 799
        
        Case 1
            If cboSiglaTributo = "" Then
                Erro "Informe o tributo."
                DefinirSelecao = False
            Else
                Filtro = "{Tab_Geracao_Tributo.tgt_tip_cod_imposto} ='" & cboSiglaTributo.Coluna(1).Valor & "'"
                If Trim$(txtPeriodoInicial) <> "" Then
                    Filtro = Filtro & " and {Tab_Geracao_Tributo.tgt_periodo} >=" & txtPeriodoInicial
                End If
                If Trim$(txtPeriodoFinal) <> "" Then
                    Filtro = Filtro & " and {Tab_Geracao_Tributo.tgt_periodo} <=" & txtPeriodoFinal
                End If
            End If
            
        Case 3 'LANCAMENTO IPTU 100 MAIROES
            Filtro = "{Tab_Geracao_Tributo.tgt_tim_ic} <> ''"
            Filtro = Filtro & " and ({Tab_Geracao_Tributo.tgt_tip_cod_imposto} = '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "' or {Tab_Geracao_Tributo.tgt_tip_cod_imposto} = '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "')"
            If Trim$(txtPeriodoInicial) <> "" Then
                Filtro = Filtro & " and {Tab_Geracao_Tributo.tgt_periodo} >=" & txtPeriodoInicial
            End If
            If Trim$(txtPeriodoFinal) <> "" Then
                Filtro = Filtro & " and {Tab_Geracao_Tributo.tgt_periodo} <=" & txtPeriodoFinal
            End If
        
        Case 4 'ARRECADACAO IPTU 100 MAIORES
            Filtro = "{Tab_Darm_Recebido.tdr_tim_ic} <> ''"
            Filtro = Filtro & " and ({Tab_Darm_Recebido.tdr_tip_cod_imposto} = '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "' or {Tab_Darm_Recebido.tdr_tip_cod_imposto} = '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "')"
            If Trim$(txtPeriodoInicial) <> "" Then
                Filtro = Filtro & " and {Tab_Darm_Recebido.tdr_periodo} >=" & txtPeriodoInicial
            End If
            If Trim$(txtPeriodoFinal) <> "" Then
                Filtro = Filtro & " and {Tab_Darm_Recebido.tdr_periodo} <=" & txtPeriodoFinal
            End If
    End Select
    
    If Filtro <> "" Then
        Rpt.Selecao = Filtro
    End If
End Function
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
                                " TGE_NOME ='RELATORIOS GERENCIAIS TREL402')" & _
        " ORDER BY TGE_NOME"
    grdRelatorios.Preencher Bdados, Sql
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
    If Not grdRelatorios.SelectedItem Is Nothing Then
        fraFiltro.Caption = ":. " & grdRelatorios.SelectedItem.SubItems(1)
    End If
End Sub

Private Sub grdRelatorios_DblClick()
    If Not grdRelatorios.SelectedItem Is Nothing Then
        tabRelatorios.Tabs(2).Selected = True
        fraFiltro.Caption = ":. " & grdRelatorios.SelectedItem.SubItems(1)
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
    
    Sql = "Select DISTINCT(tlg_nome),tlg_cod_logradouro From Tab_Logradouro where tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
    cboLogradouro.Preencher Bdados, Sql
    
    Sql = "Select DISTINCT(tba_nome),tba_cod_bairro From Tab_Bairro where TBA_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio
    cboBairro.Preencher Bdados, Sql
End Sub

Private Sub PrepararArrecadacao()
    Dim Sql As String
    
    Sql = "Select tar_nome_agente " & _
        " from tab_agente_arrecadador " & _
        " where tar_ativo =0"
    cboAgenteArrecadador.Preencher Bdados, Sql
End Sub

Private Function VToA(Arg As String) As String
VToA = Replace(CDbl(Arg), ",", ".")
End Function
