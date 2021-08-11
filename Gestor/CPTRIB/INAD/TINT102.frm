VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TINT102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TINT102"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   1138
      Icone           =   "TINT102.frx":0000
   End
   Begin VTOcx.grdVISUAL grdDoc 
      Height          =   2940
      Left            =   30
      TabIndex        =   7
      Top             =   1590
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   5186
      CorBorda        =   32768
      Caption         =   "Documentos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin Cabecalho.rodVISUAL rod 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   8
      Top             =   4620
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   820
      CorFrente       =   0
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   4380
         TabIndex        =   3
         Top             =   60
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
         Left            =   6330
         TabIndex        =   5
         Top             =   60
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
         Left            =   5355
         TabIndex        =   4
         Top             =   60
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
         Left            =   3405
         TabIndex        =   2
         Top             =   45
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
      Height          =   825
      Left            =   45
      TabIndex        =   9
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   690
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1455
      Altura          =   1905
      Caption         =   " Dados Documentos"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtDocumento 
         Height          =   480
         Left            =   1740
         TabIndex        =   6
         Tag             =   "Descricao"
         Top             =   330
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   847
         Caption         =   "Documento"
         Text            =   ""
         TipoLetras      =   0
         AlinhamentoRotulo=   1
         MaxLen          =   100
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   480
         Left            =   45
         TabIndex        =   0
         Top             =   330
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
   End
End
Attribute VB_Name = "TINT102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private correlativo As New ContaCorrente

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    
     If CriticaCampos(Me) = False Then Exit Sub
    
    If (txtCodigo = "") Then txtCodigo = correlativo.GeraCodPagamento(1)
    
    Campos = " TDI_CODIGO ,TDI_DOCUMENTO "
    Valores = Bdados.PreparaValor(txtCodigo, txtDocumento)
    Condicao = " TDI_CODIGO = " & txtCodigo
    If Bdados.GravaDados("TAB_DOCUMENTOS_INTIMACAO", Valores, Campos, Condicao) Then
        Informa "Dados gravados com sucesso!"
        PreencherGrid
        cmdLimpar_Click
        
    End If
End Sub


Private Sub cmdExcluir_Click()
    If grdDoc.ListItems.Count >= 1 Then
    If txtCodigo = "" Then Exit Sub
        If Util.Confirma("Deseja excluir a infração?", "Excluir Infração?") = True Then
            If Bdados.DeletaDados("TAB_DOCUMENTOS_INTIMACAO", "TDI_CODIGO=" & txtCodigo) Then
                Avisa "Documento excluido com sucesso."
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
    
    Sql = "select TDI_CODIGO AS Código,TDI_DOCUMENTO AS Documento FROM TAB_DOCUMENTOS_INTIMACAO"
    
    grdDoc.Preencher Bdados, Sql
End Sub

Private Sub grdDoc_dblClick()
    If grdDoc.ListItems.Count >= 1 Then
        txtCodigo = grdDoc.SelectedItem
        txtDocumento = grdDoc.SelectedItem.SubItems(1)
       
    End If

End Sub

