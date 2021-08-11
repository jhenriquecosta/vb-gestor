VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form THOM102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THOM102"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   1138
      Icone           =   "THOM102.frx":0000
   End
   Begin VTOcx.grdVISUAL grdInfra 
      Height          =   3345
      Left            =   30
      TabIndex        =   10
      Top             =   2460
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   5900
      CorBorda        =   32768
      Caption         =   "Infrações"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin Cabecalho.rodVISUAL rod 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   11
      Top             =   5850
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   820
      CorFrente       =   0
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   7290
         TabIndex        =   5
         Top             =   60
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   0
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   9240
         TabIndex        =   7
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   0
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10215
         TabIndex        =   8
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   0
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   8265
         TabIndex        =   6
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   32768
         CorFrente       =   0
         CorFoco         =   14737632
      End
   End
   Begin VTOcx.fraVISUAL fraInfra 
      Height          =   1725
      Left            =   45
      TabIndex        =   12
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   690
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   3043
      Altura          =   1905
      Caption         =   " Dados da Infração"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtAgravante 
         Height          =   285
         Left            =   8505
         TabIndex        =   3
         Tag             =   "Valor"
         Top             =   1350
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   503
         Caption         =   "Agravante (%)"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtValorUfm 
         Height          =   285
         Left            =   6030
         TabIndex        =   2
         Tag             =   "Valor"
         Top             =   1350
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   503
         Caption         =   "Valor (UFM)"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtAgravanteUFM 
         Height          =   480
         Left            =   3615
         TabIndex        =   9
         Tag             =   "Agravate"
         Top             =   -570
         Visible         =   0   'False
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   847
         Caption         =   "Agravante (UFM%)"
         Text            =   ""
         Formato         =   5
         Restricao       =   2
         AlinhamentoRotulo=   1
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtArtigo 
         Height          =   285
         Left            =   1245
         TabIndex        =   4
         Tag             =   "Artigo"
         Top             =   1035
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   503
         Caption         =   "Artigo"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   765
         TabIndex        =   13
         Top             =   405
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   503
         Caption         =   "Nº Infração"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   8
      End
      Begin VTOcx.txtVISUAL txtDescricao 
         Height          =   285
         Left            =   915
         TabIndex        =   0
         Tag             =   "Descricao"
         Top             =   720
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   503
         Caption         =   "Descricao"
         Text            =   ""
         TipoLetras      =   0
      End
      Begin VTOcx.cboVISUAL cboGravidade 
         Height          =   510
         Left            =   285
         TabIndex        =   1
         Top             =   -570
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   900
         Caption         =   "Gravidade"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
   End
End
Attribute VB_Name = "THOM102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Codigo As New ContaCorrente
Dim CodigoInfracao As String
Private Sub cmdExcluir_Click()
    If grdInfra.ListItems.Count >= 1 Then
        If txtCodigo <> "" Then
            If Util.Confirma("Deseja excluir a infração?", "Excluir Infração?") = True Then
                If Bdados.DeletaDados("TAB_INFRACAO", "TIN_COD_INFRACAO=" & grdInfra.SelectedItem) Then
                    Avisa "Infração excluida com sucesso."
                    LimpaCampos Me
                    PreencherGrid
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    CodigoInfracao = ""
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim rs As VSRecordset
    
    txtAgravanteUFM.Tag = ""
    
    If CriticaCampos(Me) = False Then Exit Sub
    
    If (CodigoInfracao = "") Then CodigoInfracao = Codigo.GeraCodPagamento(94)
    
    If Not CriticaCampos(Me) Then Exit Sub
    Campos = "  TIN_COD_INFRACAO,TIN_DESCRICAO_INFRACAO,TIN_TGR_COD_GRAVIDADE,TIN_VALOR_UFM,TIN_AGRAVANTE_UFM,TIN_ARTIGO,TIN_REFERENCIA"
    Valores = Bdados.PreparaValor(CodigoInfracao, txtDescricao, cboGravidade.Coluna(1).Valor, txtValorUfm, Bdados.Converte(txtAgravante, TCMonetario), txtArtigo, Bdados.Converte(txtCodigo, tctexto))
    If Bdados.GravaDados("TAB_INFRACAO", Valores, Campos, "TIN_COD_INFRACAO =" & CodigoInfracao) Then
        Informa "Infração gravada com sucesso!"
        PreencherGrid
        cmdLimpar_Click
    End If
End Sub



Private Sub Form_Load()
      PreencherGrid
      cabVISUAL1.Exibir Bdados, Me.Name, App.Path
      rod.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
      cboGravidade.Preencher Bdados, "select * from TAB_GRAVIDADE_INFRACAO"
End Sub

Private Sub PreencherGrid()
    Dim Sql As String
    
    Sql = " SELECT TIN_COD_INFRACAO as Código ,tin_referencia as Infração,"
    Sql = Sql & " tin_descricao_infracao as Descrição,"
    Sql = Sql & " tin_valor_ufm As VALOR, tin_artigo As Artigo,TIN_AGRAVANTE_UFM AS Agravante"
    Sql = Sql & " From TAB_INFRACAO"
    
    grdInfra.Preencher Bdados, Sql
End Sub

Private Sub grdInfra_dblClick()
    If grdInfra.ListItems.Count >= 1 Then
        CodigoInfracao = grdInfra.SelectedItem
        txtCodigo = grdInfra.SelectedItem.SubItems(1)
        txtDescricao = grdInfra.SelectedItem.SubItems(2)
        txtValorUfm = grdInfra.SelectedItem.SubItems(3)
        txtArtigo = grdInfra.SelectedItem.SubItems(4)
        txtAgravante = grdInfra.SelectedItem.SubItems(5)
    End If

End Sub

