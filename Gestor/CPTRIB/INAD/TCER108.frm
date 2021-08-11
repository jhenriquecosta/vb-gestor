VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles1.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCER108 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCER108"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   60
      ScaleHeight     =   570
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCER108.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   1138
      Icone           =   "TCER108.frx":2123
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   7140
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL CmdImprimir 
         Height          =   375
         Left            =   5655
         TabIndex        =   11
         Top             =   90
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   661
         Caption         =   "Imprimir"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7725
         TabIndex        =   13
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
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8670
         TabIndex        =   14
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
      Begin VTOcx.cmdVISUAL cmdEmitir 
         Height          =   375
         Left            =   6780
         TabIndex        =   12
         Top             =   90
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Emitir"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2670
      Left            =   45
      TabIndex        =   18
      Top             =   705
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   4710
      Altura          =   1905
      Caption         =   " Período de Entrega"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboCertidao 
         Height          =   315
         Left            =   420
         TabIndex        =   0
         Top             =   375
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   556
         Caption         =   "Certidão"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtObs 
         Height          =   510
         Left            =   150
         TabIndex        =   9
         Top             =   2070
         Width           =   9210
         _ExtentX        =   16245
         _ExtentY        =   900
         Caption         =   "Observacão"
         Text            =   ""
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   300
         Left            =   8985
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   750
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   5550
         TabIndex        =   3
         Top             =   750
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   300
         Left            =   3330
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   750
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   300
         Left            =   75
         TabIndex        =   5
         Top             =   1080
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   529
         Caption         =   "Nome/Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   390
         TabIndex        =   1
         Top             =   750
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   529
         Caption         =   "Inscricão"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtFinalidade 
         Height          =   300
         Left            =   315
         TabIndex        =   7
         Tag             =   "Finalidade"
         Top             =   1740
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   529
         Caption         =   "Finalidade"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtValidade 
         Height          =   300
         Left            =   7380
         TabIndex        =   8
         Tag             =   "Validade"
         Top             =   1740
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "Validade"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   0
         Restricao       =   2
         MaxLen          =   10
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1215
         TabIndex        =   6
         Top             =   1410
         Width           =   8130
      End
   End
   Begin ActiveTabs.SSActiveTabs tabCND 
      Height          =   3690
      Left            =   45
      TabIndex        =   10
      Top             =   3390
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   6509
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
      Tabs            =   "TCER108.frx":2B65
      Images          =   "TCER108.frx":2BE3
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3270
         Index           =   0
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "TCER108.frx":387C
         Begin VB.TextBox txtTexto 
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
            Height          =   3180
            Left            =   45
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   30
            Width           =   9390
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3270
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "TCER108.frx":38A4
         Begin VTOcx.grdVISUAL grdCPND 
            Height          =   3180
            Left            =   15
            TabIndex        =   22
            Top             =   90
            Width           =   9420
            _ExtentX        =   16616
            _ExtentY        =   5609
            CorBorda        =   32768
            Caption         =   "Certidões emitidas"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
   End
End
Attribute VB_Name = "TCER108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conta As ContaCorrente
Dim CodCertidao As String
Dim InscricaoCad As String, InscricaoMun As String
Dim Inscricao As String
Dim Tipo As Integer
Dim condicao As String
Dim emitir As Boolean

Private Sub BuscaCertidao()
    Dim Sql As String
    
    Sql = "select TCG_COD_NEGATIVA as Código,"
    Sql = Sql & " TCG_DATA_EMISSAO as Data_Emissão,"
    Sql = Sql & " TCG_FINALIDADE as Finalidade,"
    Sql = Sql & " TCG_VALIDADE as Validade,"
    Sql = Sql & " TCG_OBSERVACAO As Observação "
    Sql = Sql & " From tab_certidao_generica"
    Sql = Sql & " where TCG_TCI_INSCRICAO = '" & Inscricao & "'"
    
    If condicao <> "" Then Sql = Sql & " and " & condicao
    grdCPND.Preencher Bdados, Sql
End Sub
Sub ImprimeCertidao(CodCertidao As String)
    Dim RELAT As VSRelatorio
    Dim Im As String, Ic As String
    Dim Filtro As String
    Dim Sql As String
    Set RELAT = New VSRelatorio

    With RELAT
        If Not .DefinirArquivo(Bdados, App.Path + "\TCertidaoGenerica.rpt") Then Exit Sub
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            .Selecao = "{TAB_CERTIDAO_GENERICA.TCG_COD_NEGATIVA} = " & CodCertidao
            '.Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
            .Titulo = cboCertidao
            .Formulas "VT_CIDADE", AplicacoesVTFuncoes.Municipio
            .Arvore = False
            .CopiasDetalhes = 2
            .Visualizar
    End With
    Set RELAT = Nothing
End Sub

Private Sub cboCertidao_Click()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim dias As Integer
    
    Sql = "select TCE_CODIGO,"
    Sql = Sql & " TCE_VALIDADE,"
    Sql = Sql & " TCE_TEXTO "
    Sql = Sql & " From tab_tipo_certidao "
    Sql = Sql & " Where TCE_CODIGO = " & cboCertidao.Coluna(1).Valor
    
    If Bdados.AbreTabela(Sql, rs) Then
        dias = CInt("" & rs!TCE_VALIDADE)
        txtValidade = DateAdd("D", dias, Date)
        txtTexto = "" & rs!TCE_TEXTO
    End If
    condicao = "TCG_TIPO_CERTIDAO = " & cboCertidao.Coluna(1).Valor
    If txtIm <> "" Or txtImovel <> "" Then
        BuscaCertidao
        condicao = ""
        
    End If
End Sub

Private Sub cmdEmitir_Click()
    Dim Valores As String
    Dim campos As String
    If txtIm = "" And txtImovel = "" Or txtIm <> "" And txtImovel <> "" Then
        Util.Avisa "Informe Inscr.Municipal ou Cadastral."
        txtIm.SetFocus
        Exit Sub
    End If
    If Not CriticaCampos(Me) Then Exit Sub
    If Not Util.Confirma("Confirma a emissão da certidão") Then Exit Sub
        
    CodCertidao = Conta.GeraCodPagamento("37")
    
    campos = "TCG_COD_NEGATIVA,TCG_TCI_INSCRICAO,TCG_DATA_EMISSAO,TCG_FINALIDADE,TCG_VALIDADE,TCG_TUS_COD_USUARIO,TCG_TIPO,TCG_OBSERVACAO,TCG_TIPO_CERTIDAO "
    Valores = Bdados.PreparaValor(CodCertidao, Bdados.Converte(Inscricao, tctexto), Format(Date, "DD/MM/YYYY"), txtFinalidade, Format(txtValidade, "DD/MM/YYYY"), AplicacoesVTFuncoes.Usuario, Tipo, txtObs, cboCertidao.Coluna(1).Valor)
    
    If Bdados.InsereDados("TAB_CERTIDAO_GENERICA", Valores, campos) Then
        Avisa "Certidão emitida com sucesso"
        ImprimeCertidao CodCertidao
        cmdLimpar_Click
    End If
    

End Sub

Private Sub cmdImprimir_Click()
    If grdCPND.ListItems.Count >= 1 Then
        ImprimeCertidao grdCPND.SelectedItem
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    grdCPND.ListItems.Clear
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub




Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Set Conta = New ContaCorrente
    cboCertidao.Preencher Bdados, "select tce_nome,tce_codigo from tab_tipo_certidao"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Set Conta = Nothing
End Sub

Private Sub txtIm_LostFocus()
    Dim RetNome As String
    Dim Doc As String
    
    If Trim(txtIm) = "" Then Exit Sub
    txtIm.AgruparValores = False
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, Doc, etiContribuinte)
    InscricaoMun = txtIm: InscricaoCad = ""
     Inscricao = txtIm
     Tipo = 2
    Call BuscaCertidao
   
End Sub

Private Sub txtImovel_LostFocus()
   Dim RetNome As String
   Dim Doc As String
    
    If Trim(txtImovel) = "" Then Exit Sub
    txtImovel.AgruparValores = False
    txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, Doc, etiImovel)
    InscricaoMun = "": InscricaoCad = txtImovel
    Inscricao = txtImovel
    Tipo = 1
    Call BuscaCertidao
    
End Sub

Private Function CodAtividade(Contribuinte As String) As String
    Dim Sql As String
    Sql = "Select tci_tae_cae from tab_contribuinte where tci_im = '" & Contribuinte & "'"
    If Bdados.AbreTabela(Sql) Then
        CodAtividade = "" & Bdados.Tabela("tci_tae_cae")
    End If
End Function
Private Function PegaDoc(Contribuinte As String) As String
    Dim Sql As String
    Sql = "Select tci_cgc_cpf from tab_contribuinte where tci_im = '" & Contribuinte & "'"
    If Bdados.AbreTabela(Sql) Then
        PegaDoc = "" & Bdados.Tabela("tci_cgc_cpf")
    End If
End Function
