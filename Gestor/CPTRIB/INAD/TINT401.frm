VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TINT401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TINT401"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   3
      Top             =   6810
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   820
      Begin VTOcx.cmdVISUAL cmdRelA 
         Height          =   330
         Left            =   5805
         TabIndex        =   12
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   7935
         TabIndex        =   2
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   6930
         TabIndex        =   1
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   1138
      Icone           =   "TINT401.frx":0000
   End
   Begin VTOcx.grdVISUAL grdInt 
      Height          =   2115
      Left            =   15
      TabIndex        =   5
      Top             =   2190
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   3731
      CorBorda        =   32768
      Caption         =   "Intimação"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.grdVISUAL grdDoc 
      Height          =   2415
      Left            =   15
      TabIndex        =   6
      Top             =   4350
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   4260
      CorBorda        =   32768
      Caption         =   "Documentos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1440
      Left            =   45
      TabIndex        =   7
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   675
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   2540
      Altura          =   1905
      Caption         =   " Dados do Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   540
         TabIndex        =   11
         Top             =   735
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   556
         Caption         =   "Razão"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   315
         Left            =   270
         TabIndex        =   10
         Top             =   1080
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   556
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
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
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   300
         Left            =   2580
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   405
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
      End
      Begin VTOcx.txtVISUAL txtCgc 
         Height          =   300
         Left            =   2940
         TabIndex        =   8
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
   End
End
Attribute VB_Name = "TINT401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelecaoRpt As String

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdInt.ListItems.Clear
    grdDoc.ListItems.Clear
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub

Private Sub cmdRelA_Click()
          If grdDoc.ListItems.Count < 1 Then Exit Sub
          
          With Rpt
            If Not .DefinirArquivo(Bdados, App.Path & "\TIntimacao.rpt") Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
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
    
    Set Rpt = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub

Private Sub grdInt_DblClick()
    If grdInt.ListItems.Count < 1 Then Exit Sub
    CarregaItemIntimacao
    SelecaoRpt = "{TAB_INTIMACAO.TIN_CODIGO}= " & grdInt.SelectedItem
End Sub

Private Sub txtIm_LostFocus()
    Dim rs As VSRecordset
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, txtCgc, etiContribuinte)
   If Bdados.AbreTabela("select tci_cgc_cpf from tab_contribuinte t where tci_im = '" & txtIm & "'", rs) Then
       txtCgc = "" & rs!TCI_CGC_CPF
    End If
    CarregaIntimacao
End Sub

Private Sub CarregaIntimacao()
    Dim Sql As String
    
    Sql = "select  TIN_CODIGO as Código,TIN_IM as Inscrição,TCI_NOME as Contribuinte , TIN_DATA_EMISSAO as Data_Emissão,TIN_PERIODO_INICIAL as Período_Inicial, TIM_PERIODO_FINAL as Período_Final  from tab_intimacao , tab_contribuinte where tci_im = TIN_IM and TIN_IM = " & txtIm
    grdInt.Preencher Bdados, Sql
End Sub
Private Sub CarregaItemIntimacao()
    Dim Sql As String
    
    Sql = "select TII_COD_INTIMACAO as Código,TDI_DOCUMENTO as Documento,TII_DATA_ENTREGA as Data_Entrega from tab_item_intimacao,TAB_DOCUMENTOS_INTIMACAO WHERE    TII_COD_DOCUMENTO =TDI_CODIGO AND TII_COD_INTIMACAO = " & grdInt.SelectedItem
    grdDoc.Preencher Bdados, Sql
End Sub
