VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{EA761AE1-8FDE-4340-8E6D-420E99B0C363}#1.0#0"; "VTControles.ocx"
Begin VB.Form CTRN401 
   BackColor       =   &H00FFF5EC&
   Caption         =   "CTRN401"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cboVISUAL cboSistema 
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   720
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   556
      Caption         =   "Sistema"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   4050
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   820
      CorFundo        =   14737632
      CorFrente       =   8421504
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7800
         TabIndex        =   1
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   4210752
         CorFrente       =   4210752
         CorFundo        =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1138
      Icone           =   "CTRN401.frx":0000
   End
   Begin VTOcx.grdVISUAL grdGrupo 
      Height          =   2865
      Left            =   60
      TabIndex        =   3
      Top             =   1110
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   4339
      CorBorda        =   12632064
      Caption         =   "Grupos"
      CorTitulo       =   12632064
      CorCaption      =   16777215
      CorDica         =   8388608
   End
   Begin VTOcx.grdVISUAL grdConsulta 
      Height          =   3285
      Left            =   3900
      TabIndex        =   4
      Top             =   690
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4339
      CorBorda        =   12632064
      Caption         =   "Consultas"
      CorTitulo       =   12632064
      CorCaption      =   16777215
      CorDica         =   8388608
   End
End
Attribute VB_Name = "CTRN401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub PreencherSistemas()
    Dim sql As String
    
    sql = "SELECT DISTINCT TSI_COD_SISTEMA FROM TAB_SISTEMA, TAB_GRUPO_TRANSFERENCIA WHERE TSI_COD_SISTEMA=TGT_TSI_COD_SISTEMA ORDER BY TSI_COD_SISTEMA "
    cboSistema.Preencher Bdados, sql
End Sub

Private Sub PreencherGrupos(CodSistema As String)
    Dim sql As String
    
    sql = "SELECT TGT_COD_GRUPO AS Codigo, TGT_GRUPO AS Grupo FROM TAB_GRUPO_TRANSFERENCIA WHERE TGT_TSI_COD_SISTEMA='" & CodSistema & "'"
    grdGrupo.Preencher Bdados, sql
End Sub

Private Sub PreencherConsultas(CodGrupo As String)
    Dim sql As String
    
    sql = "SELECT TTT_COD_CONSULTA AS Codigo, TTT_CONSULTA AS Consulta, TTT_LIMPAR_DESTINO as Limpar FROM TAB_TABELA_TRANSFERENCIA WHERE TTT_TGT_COD_GRUPO=" & CodGrupo
    grdConsulta.Preencher Bdados, sql
End Sub

Private Sub ExcluirGrupo(CodGrupo As String)
    Bdados.DeletaDados "TAB_TABELA_TRANSFERENCIA", "TTT_TGT_COD_GRUPO=" & CodGrupo
    Bdados.DeletaDados "TAB_GRUPO_TRANSFERENCIA", "TGT_COD_GRUPO=" & CodGrupo
End Sub

Private Sub ExcluirConsulta(CodGrupo As String, CodConsulta As String)
    Bdados.DeletaDados "TAB_TABELA_TRANSFERENCIA", "TTT_TGT_COD_GRUPO=" & CodGrupo & " AND TTT_COD_CONSULTA=" & CodConsulta
End Sub
Private Sub cboSistema_Click()
    PreencherGrupos cboSistema
    grdConsulta.Preencher Bdados, ""
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    PreencherSistemas
End Sub

Private Sub grdConsulta_DblClick()
    If Not grdConsulta.SelectedItem Is Nothing Then
        If Util.Confirma("Excluir " & grdConsulta.SelectedItem.SubItems(1) & "?") Then
            ExcluirConsulta grdGrupo.SelectedItem, grdConsulta.SelectedItem
            PreencherConsultas grdGrupo.SelectedItem
        End If
    End If
End Sub

Private Sub grdGrupo_Click()
    If Not grdGrupo.SelectedItem Is Nothing Then
        PreencherConsultas grdGrupo.SelectedItem
    End If
End Sub

Private Sub grdGrupo_DblClick()
    If Not grdGrupo.SelectedItem Is Nothing Then
        If Util.Confirma("Excluir " & grdGrupo.SelectedItem.SubItems(1) & "?") Then
            ExcluirGrupo grdGrupo.SelectedItem
            PreencherGrupos cboSistema
        End If
    End If
End Sub
