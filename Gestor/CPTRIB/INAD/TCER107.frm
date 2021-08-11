VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCER107 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCER107"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   1138
      Icone           =   "TCER107.frx":0000
   End
   Begin VTOcx.grdVISUAL grdDoc 
      Height          =   2445
      Left            =   75
      TabIndex        =   4
      Top             =   4185
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   4313
      CorBorda        =   32768
      Caption         =   "Certidões"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin Cabecalho.rodVISUAL rod 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   10
      Top             =   6720
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   820
      CorFrente       =   0
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   6060
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
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7980
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
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7020
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
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   5100
         TabIndex        =   5
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   0
         CorFoco         =   14737632
      End
   End
   Begin VTOcx.fraVISUAL fraInfra 
      Height          =   3420
      Left            =   60
      TabIndex        =   11
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   690
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   6033
      Altura          =   1905
      Caption         =   " Dados da Infração"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
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
         Height          =   2235
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   3
         Tag             =   "Texto"
         Top             =   1035
         Width           =   8640
      End
      Begin VTOcx.txtVISUAL txtValidade 
         Height          =   480
         Left            =   7335
         TabIndex        =   2
         Tag             =   "Validade"
         Top             =   315
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   847
         Caption         =   "Validade (dias)"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   8
      End
      Begin VTOcx.txtVISUAL txtDescricao 
         Height          =   480
         Left            =   1755
         TabIndex        =   1
         Tag             =   "Descrição"
         Top             =   315
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   847
         Caption         =   "Descrição"
         Text            =   ""
         TipoLetras      =   0
         AlinhamentoRotulo=   1
         MaxLen          =   50
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   480
         Left            =   60
         TabIndex        =   0
         Top             =   315
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   847
         Caption         =   "Código"
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   8
      End
      Begin VB.Label Labtext 
         Caption         =   "Texto"
         Height          =   210
         Left            =   75
         TabIndex        =   12
         Top             =   840
         Width           =   1470
      End
   End
End
Attribute VB_Name = "TCER107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private correlativo As New ContaCorrente

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim campos As String
    Dim condicao As String
    
     If CriticaCampos(Me) = False Then Exit Sub
    
    If (txtCodigo = "") Then txtCodigo = correlativo.GeraCodPagamento(2)
    
    campos = " TCE_CODIGO,TCE_NOME,TCE_VALIDADE,TCE_TEXTO "
    Valores = Bdados.PreparaValor(txtCodigo, txtDescricao, txtValidade, txtTexto)
    condicao = " TCE_CODIGO = " & txtCodigo
    If Bdados.GravaDados("TAB_TIPO_CERTIDAO", Valores, campos, condicao) Then
        Informa "Dados gravados com sucesso!"
        PreencherGrid
        cmdLimpar_Click
        txtDescricao.SetFocus
    End If
End Sub


Private Sub cmdExcluir_Click()
    If grdDoc.ListItems.Count >= 1 Then
    If txtCodigo = "" Then Exit Sub
        If Util.Confirma("Deseja excluir a Certidão?", "Excluir Certidão?") = True Then
            If Bdados.DeletaDados("TAB_TIPO_CERTIDAO", "TCE_CODIGO=" & txtCodigo) Then
                Avisa "Certidão excluida com sucesso."
                LimpaCampos Me
                PreencherGrid
            End If
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
      PreencherGrid
      cabVISUAL1.Exibir Bdados, Me.Name, App.Path
      rod.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
End Sub

Private Sub PreencherGrid()
    Dim Sql As String
    
    Sql = "select TCE_CODIGO as Código,TCE_NOME as Certidão,TCE_VALIDADE as Validade,TCE_TEXTO as Texto FROM TAB_TIPO_CERTIDAO"
    
    grdDoc.Preencher Bdados, Sql
End Sub

Private Sub grdDoc_dblClick()
    If grdDoc.ListItems.Count >= 1 Then
        txtCodigo = grdDoc.SelectedItem
        txtDescricao = grdDoc.SelectedItem.SubItems(1)
        txtValidade = grdDoc.SelectedItem.SubItems(2)
        txtTexto = grdDoc.SelectedItem.SubItems(3)
       
    End If

End Sub

