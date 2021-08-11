VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TINT103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TINT103"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   9
      Top             =   7095
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   820
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   5910
         TabIndex        =   6
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
         Icone           =   "TINT103.frx":0000
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   315
         Left            =   7950
         TabIndex        =   8
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   6945
         TabIndex        =   7
         Top             =   105
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
      TabIndex        =   10
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   1138
      Icone           =   "TINT103.frx":031A
   End
   Begin VTOcx.grdVISUAL grdInt 
      Height          =   1905
      Left            =   30
      TabIndex        =   11
      Top             =   2190
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   3360
      CorBorda        =   32768
      Caption         =   "Intimação"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.grdVISUAL grdDoc 
      Height          =   2385
      Left            =   15
      TabIndex        =   12
      Top             =   4155
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   4207
      CorBorda        =   32768
      Caption         =   "Documentos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1440
      Left            =   45
      TabIndex        =   13
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
         TabIndex        =   3
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
         TabIndex        =   4
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
      Begin VTOcx.txtVISUAL txtCgc 
         Height          =   300
         Left            =   2940
         TabIndex        =   2
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
   Begin VTOcx.txtVISUAL txtData 
      Height          =   300
      Left            =   6285
      TabIndex        =   5
      Tag             =   "Data Entrega"
      Top             =   6645
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   529
      Caption         =   "Data Entrega"
      Text            =   ""
      Formato         =   0
      Restricao       =   2
      MaxLen          =   10
      RetirarMascara  =   0   'False
   End
End
Attribute VB_Name = "TINT103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdInt.ListItems.Clear
    grdDoc.ListItems.Clear
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim camposdoc As String
    Dim valoresdoc As String
    Dim Condicao As String
    Dim i As Integer
        
    Dim marcou As Boolean
    
       For i = 1 To grdDoc.ListItems.Count
             If (grdDoc.ListItems(i).Checked) Then
                marcou = True
             End If
        Next
    If Not marcou Then Avisa "Selecione os Documentos": Exit Sub
        
    If Not CriticaCampos(Me) Then Exit Sub
    If grdDoc.ListItems.Count < 1 Then Exit Sub
    
    valoresdoc = Bdados.PreparaValor(txtData)
    camposdoc = "TII_DATA_ENTREGA"
    For i = 1 To grdDoc.ListItems.Count
          If (grdDoc.ListItems(i).Checked) Then
            
             Condicao = "TII_COD_INTIMACAO = " & grdDoc.ListItems(i) & " and tii_cod_documento= " & grdDoc.ListItems(i).SubItems(3)
             Bdados.AtualizaDados "TAB_ITEM_INTIMACAO", valoresdoc, camposdoc, Condicao
          End If
     Next
     Avisa "Intimação recebida com sucesso"
     grdInt_DblClick
    
End Sub



Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision

End Sub

Private Sub grdInt_DblClick()
    If grdInt.ListItems.Count < 1 Then Exit Sub
    CarregaItemIntimacao
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
    
    Sql = "select TII_COD_INTIMACAO as Código,TDI_DOCUMENTO as Documento,TII_DATA_ENTREGA as Data_Entrega,tii_cod_documento from tab_item_intimacao,TAB_DOCUMENTOS_INTIMACAO WHERE    TII_COD_DOCUMENTO =TDI_CODIGO AND TII_COD_INTIMACAO = " & grdInt.SelectedItem
    grdDoc.Preencher Bdados, Sql, 1500, 5000, 1700, 0
End Sub
