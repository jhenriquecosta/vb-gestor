VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TMPU109 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMPOSTO DIGITAL"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TMPU109.frx":0000
   ScaleHeight     =   5820
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   8
      Top             =   5295
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdExclui 
         Height          =   375
         Left            =   3300
         TabIndex        =   4
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   2295
         TabIndex        =   3
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   5310
         TabIndex        =   6
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Sair"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   4305
         TabIndex        =   5
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.grdVISUAL grdPonto 
      Height          =   3450
      Left            =   60
      TabIndex        =   2
      Top             =   1845
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   6085
      CorBorda        =   16711680
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   16711680
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1080
      Left            =   45
      TabIndex        =   7
      Top             =   720
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   1905
      Altura          =   1905
      Caption         =   " Ponto de Recepção"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtDescPonto 
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Tag             =   "Descrição"
         Top             =   705
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   503
         Caption         =   "Descrição"
         Text            =   ""
         Restricao       =   1
         ValorMaximo     =   100
         MaxLen          =   50
         MinLen          =   1
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   330
         TabIndex        =   0
         Tag             =   "Código"
         Top             =   375
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   503
         Caption         =   "Código"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         MaxLen          =   10
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1138
      Icone           =   "TMPU109.frx":0342
   End
End
Attribute VB_Name = "TMPU109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExclui_Click()
    If grdPonto.SelectedItem Is Nothing Then
        Util.Avisa "Selecione um Bairro."
    Else
        If Util.Confirma("Deseja Excluir" & grdPonto.SelectedItem & "?") Then
            If Bdados.DeletaDados("TAB_BAIRRO", "TBA_COD_BAIRRO =" & txtCodigo) Then
                Informa "Transação completada."
                cmdLimpar_Click
            End If
        End If
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtDescPonto.SetFocus
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim rs As VSRecordset
    Dim Sql As String
    Dim Valores As String
    Dim Campos As String
    Dim Grupo As Byte
    Dim NomeGrupo As String

    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    Valores = Bdados.PreparaValor(txtCodigo, txtDescPonto)
    Campos = "TBA_COD_BAIRRO,TBA_NOME,TBA_TMU_COD_MUNICIPIO"
    Bdados.GravaDados "TAB_BAIRRO", Valores, Campos, "TBA_COD_BAIRRO= " & txtCodigo
    grdPonto.Preencher Bdados, "SELECT TBA_COD_BAIRRO AS CODIGO,TBA_NOME AS BAIRRO FROM TAB_BAIRRO ", 1100, 4500
    Screen.MousePointer = 0
    Avisa "Registro gravado com sucesso."
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    
    grdPonto.Preencher Bdados, "SELECT TBA_COD_BAIRRO AS CODIGO,TBA_NOME AS BAIRRO FROM TAB_BAIRRO ", 1100, 4500
    AtualizaCabecalho grdPonto
End Sub


Private Sub grdPonto_DblClick()
    DoEvents
    txtCodigo = grdPonto.SelectedItem
    txtCodigo_LostFocus
End Sub

Private Sub txtCodigo_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtCodigo) = "" Then Exit Sub
    Sql = "SELECT TBA_NOME FROM TAB_BAIRRO where TBA_COD_BAIRRO =" & txtCodigo
    If Bdados.AbreTabela(Sql, rs) Then
        txtDescPonto = rs!TBA_NOME
    End If
End Sub
