VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCIS104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TCIS104.frx":0000
   ScaleHeight     =   5820
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   30
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCIS104.frx":0342
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   9
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
         CorBorda        =   8421504
         CorFrente       =   16384
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
         CorBorda        =   8421504
         CorFrente       =   16384
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
         CorBorda        =   8421504
         CorFrente       =   16384
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
         CorBorda        =   8421504
         CorFrente       =   16384
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
      CorBorda        =   32768
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
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
      CorFaixa        =   32768
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
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1138
      Icone           =   "TCIS104.frx":2465
   End
End
Attribute VB_Name = "TCIS104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ponto As New ePontoRecepcao

Private Sub cmdExclui_Click()
    If grdPonto.SelectedItem Is Nothing Then
        Util.Avisa "Selecione um Ponto."
    Else
        If Util.Confirma("Deseja Excluir" & grdPonto.SelectedItem & "?") Then
            If Ponto.Excluir(grdPonto.SelectedItem) Then
                Informa "Transação completada."
                cmdLimpar_Click
            End If
        End If
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    Ponto.PreencherGrd grdPonto
    txtDescPonto.SetFocus
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
'    Dim Rs As VSRecordset
'    Dim Sql As String
'    Dim Valores As String
'    Dim Campos As String
'    Dim Grupo As Byte
'    Dim NomeGrupo As String
'
'    If Not Edita.CriticaCampos(Me) Then Exit Sub
'    Screen.MousePointer = 11
'    Valores = Bdados.PreparaValor(txtCodigo, txtDescAtiv)
'    Campos = "TRA_COD_Ponto,TRA_NOME_Ponto"
'    Bdados.GravaDados "TAB_Ponto_ATIVIDADE", Valores, Campos, "TRA_COD_Ponto = " & txtCodigo
'    lstAtv.Preencher Bdados, "SELECT TRA_COD_Ponto AS CODIGO,TRA_NOME_Ponto AS DESCRICAO FROM TAB_Ponto_ATIVIDADE ", 1100, 4500
    If Edita.CriticaCampos(Me) Then
        With Ponto
            .CodPonto = txtCodigo
            .NomePonto = txtDescPonto
            If .Gravar Then
                Util.Avisa "Transação Finalizada."
                cmdLimpar_Click
            End If
        End With
    End If
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set Ponto = New ePontoRecepcao
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    Ponto.PreencherGrd grdPonto
    'lstAtv.Preencher Bdados, "SELECT TRA_COD_Ponto AS CODIGO,TRA_NOME_Ponto AS DESCRICAO FROM TAB_Ponto_ATIVIDADE ", 1100, 4500
    AtualizaCabecalho grdPonto
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Ponto = Nothing
End Sub

Private Sub grdPonto_DblClick()
    DoEvents
    txtCodigo = grdPonto.SelectedItem
    txtCodigo_LostFocus
End Sub

Private Sub txtCodigo_LostFocus()
    Dim sql As String
    Dim rs As VSRecordset
    If Trim(txtCodigo) = "" Then Exit Sub
    If Ponto.Buscar(txtCodigo) Then
        txtDescPonto = Ponto.NomePonto
    End If
End Sub
