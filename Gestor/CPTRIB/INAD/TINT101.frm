VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TINT101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TINT101"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   3600
      Left            =   15
      TabIndex        =   11
      Top             =   2430
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   6350
      _Version        =   131082
      ForeColor       =   0
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "TINT101.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3210
         Index           =   1
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5662
         _Version        =   131082
         TabGuid         =   "TINT101.frx":0089
         Begin VB.TextBox txtTexto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   3120
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   45
            Width           =   8685
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3210
         Left            =   -99969
         TabIndex        =   16
         Top             =   30
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5662
         _Version        =   131082
         TabGuid         =   "TINT101.frx":00B1
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Selecionar Todos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   45
            TabIndex        =   18
            Top             =   2925
            Width           =   2655
         End
         Begin VTOcx.grdVISUAL grdDoc 
            Height          =   3150
            Left            =   -15
            TabIndex        =   17
            Top             =   15
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   5556
            Caption         =   "Documentos Solicitados"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
            CheckBox        =   -1  'True
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   6135
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         Top             =   105
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7605
         TabIndex        =   6
         Top             =   105
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdEmitir 
         Height          =   375
         Left            =   5340
         TabIndex        =   4
         Top             =   105
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Emitir"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   1138
      Icone           =   "TINT101.frx":00D9
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1725
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   660
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   3043
      Altura          =   1905
      Caption         =   " Dados do Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   285
         Left            =   420
         TabIndex        =   2
         Tag             =   "Período Inicial"
         Top             =   1425
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         Caption         =   "Período"
         Text            =   ""
         Formato         =   0
         MaxLen          =   10
         MinLen          =   10
      End
      Begin VTOcx.txtVISUAL txtPeriodoFinal 
         Height          =   285
         Left            =   2265
         TabIndex        =   3
         Tag             =   "Período Final"
         Top             =   1425
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         Caption         =   "a"
         Text            =   ""
         Formato         =   0
         MaxLen          =   10
         MinLen          =   10
      End
      Begin VTOcx.txtVISUAL txtCgc 
         Height          =   300
         Left            =   2940
         TabIndex        =   7
         Tag             =   "CPF/CNPJ"
         Top             =   405
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   529
         Caption         =   "CPF/CNPJ"
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         MaxLen          =   20
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   300
         Left            =   2580
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   405
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   285
         TabIndex        =   0
         Tag             =   "Inscrição"
         Top             =   405
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   529
         Caption         =   "Inscricao"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   20
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   315
         Left            =   270
         TabIndex        =   9
         Top             =   1080
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   556
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   540
         TabIndex        =   8
         Top             =   735
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   556
         Caption         =   "Razão"
         Text            =   ""
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "TINT101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private correlativo As New ContaCorrente
Dim SelecaoRpt As String

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    DesmarcaGrid
    Texto
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub

Private Sub cmdEmitir_Click()
    Dim campos As String
    Dim camposdoc As String
    Dim Valores As String
    Dim valoresdoc As String
    Dim Codigo As String
    Dim i As Integer
    Dim marcou As Boolean
    
    If Not CriticaCampos(Me) Then Exit Sub
       For i = 1 To grdDoc.ListItems.Count
             If (grdDoc.ListItems(i).Checked) Then
                marcou = True
             End If
        Next
    If Not marcou Then Avisa "Selecione os Documentos": Exit Sub
    Codigo = correlativo.GeraCodPagamento(68)
    
    campos = "  TIN_CODIGO,TIN_IM,TIN_DATA_EMISSAO,TIN_PERIODO_INICIAL,TIM_PERIODO_FINAL"
    Valores = Bdados.PreparaValor(Codigo, txtIm, Date, txtPeriodoInicial, txtPeriodoFinal)
    If Bdados.InsereDados("TAB_INTIMACAO", Valores, campos) Then
         camposdoc = "TII_COD_INTIMACAO,TII_COD_DOCUMENTO"
         For i = 1 To grdDoc.ListItems.Count
             If (grdDoc.ListItems(i).Checked) Then
                valoresdoc = Bdados.PreparaValor(Codigo, grdDoc.ListItems(i).SubItems(1))
                Bdados.InsereDados "TAB_ITEM_INTIMACAO", valoresdoc, camposdoc
             End If
        Next
        SelecaoRpt = "{TAB_INTIMACAO.TIN_CODIGO}= " & Codigo
        If Confirma("Intimação Emitida com Sucesso", "Deseja Imprimir?") Then
            Screen.MousePointer = 11
            
                With RPT
                    If Not .DefinirArquivo(Bdados, App.Path & "\TIntimacao.rpt") Then
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                    If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SMTU"), Temp.PegaParametro(Bdados, "SMTUSETOR")
                    Else
                        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                    End If
                    .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                    .Selecao = SelecaoRpt
                    .Titulo = "Ficha Cadastral"
                    .Arvore = False
                    .Visualizar
                    DoEvents
                End With
            End If
            Set RPT = Nothing
            Screen.MousePointer = 0

        End If
        cmdLimpar_Click
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    grdDoc.Preencher Bdados, "select TDI_DOCUMENTO as Documento,TDI_CODIGO from TAB_DOCUMENTOS_INTIMACAO", 6000, 0
    Texto
End Sub

Private Sub Texto()
    Dim Rs As VSRecordset
    Dim Sql As String
    Sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = 'TERMO INTIMACAO'"
    If Bdados.AbreTabela(Sql, Rs) Then
        txtTexto = "" & Rs!TPT_TEXTO
    End If
End Sub

Private Sub Check1_Click()
    grdDoc.MarcarTodos Check1
End Sub

Private Sub txtIm_LostFocus()
    Dim Rs As VSRecordset
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, txtCgc, etiContribuinte)
    If Bdados.AbreTabela("select tci_cgc_cpf from tab_contribuinte t where tci_im = '" & txtIm & "'", Rs) Then
       txtCgc = "" & Rs!TCI_CGC_CPF
    End If
End Sub
Private Sub txtPeriodoFinal_LostFocus()
    If txtPeriodoFinal = "" Then Exit Sub
    If CDate(txtPeriodoFinal) < CDate(txtPeriodoInicial) Then
        Avisa "Período final não pode ser menor que período incicial"
        txtPeriodoFinal.SetFocus
    End If
    
End Sub

Private Sub DesmarcaGrid()
Dim i As Integer
    For i = 1 To grdDoc.ListItems.Count
         If (grdDoc.ListItems(i).Checked) Then
            grdDoc.ListItems(i).Checked = False
         End If
    Next
End Sub
