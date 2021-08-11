VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "CABECALHO.OCX"
Begin VB.Form TMPU106 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   ControlBox      =   0   'False
   Icon            =   "TMPU106.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCaixa 
      Interval        =   100
      Left            =   990
      Top             =   5190
   End
   Begin Threed.SSFrame fra 
      Height          =   3825
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   1290
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6747
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Componentes"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Tag             =   "Nome da série"
         ToolTipText     =   "Data de nascimento do aluno"
         Top             =   210
         Width           =   7935
      End
      Begin VB.ListBox lstExist 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3000
         ItemData        =   "TMPU106.frx":08CA
         Left            =   120
         List            =   "TMPU106.frx":08CC
         TabIndex        =   2
         Top             =   630
         Width           =   7935
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   735
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   540
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   1296
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Grupo de Componente"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.ComboBox cboGrupo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "TMPU106.frx":08CE
         Left            =   150
         List            =   "TMPU106.frx":08D0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   7905
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   1138
      Icone           =   "TMPU106.frx":08D2
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   6900
      TabIndex        =   4
      ToolTipText     =   "Deseja sair?"
      Top             =   5190
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TMPU106.frx":0BEC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU106.frx":0C08
      Caption         =   "Sai&r"
      ButtonStyle     =   3
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdAjuda 
      Height          =   435
      Left            =   90
      TabIndex        =   6
      ToolTipText     =   "Ajuda"
      Top             =   5190
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      MousePointer    =   14
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TMPU106.frx":0C24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU106.frx":0C40
      Caption         =   "?"
      ButtonStyle     =   3
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdGravar 
      Height          =   435
      Left            =   5400
      TabIndex        =   3
      ToolTipText     =   "Salvar Informações"
      Top             =   5190
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TMPU106.frx":0C5C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU106.frx":0C78
      Caption         =   "&Gravar"
      ButtonStyle     =   3
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   435
      Left            =   3900
      TabIndex        =   5
      ToolTipText     =   "Novo"
      Top             =   5190
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TMPU106.frx":0C94
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU106.frx":0CB0
      Caption         =   "&Novo"
      ButtonStyle     =   3
      PictureAlignment=   6
   End
End
Attribute VB_Name = "TMPU106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OP As String

Private Sub cboGrupo_Click()
    If Trim(cboGrupo.Text) <> "" Then AtualizaLista
End Sub

Private Sub cboGrupo_GotFocus()
    txtEdit = ""
    lstExist.Clear
End Sub

Private Sub cboGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstExist.SetFocus
End Sub

Private Sub cboGrupo_LostFocus()
    If cboGrupo.Text = "" Then
        OP = ""
    Else
        OP = "N"
    End If
End Sub

Private Function GeraCodComponente(CodGrupo As Integer) As String
    Dim Sql As String
    Dim RS1 As VSRecordset
    Sql = "SELECT MAX(tcl_cod_componente) FROM TAB_COMPONENTE_LOGRADOURO where tcl_grupo=" & CodGrupo
    If Bdados.AbreTabela(Sql, RS1) Then
        If RS1(0) <> "" Then
            GeraCodComponente = CInt(RS1(0)) + 1
        Else
            GeraCodComponente = "1"
        End If
    End If
    Bdados.FechaTabela RS1
End Function

Private Sub cmdGravar_Click()
    On Error GoTo trata
    Dim COD As String
    
    If cboGrupo.ListIndex = -1 Then Exit Sub
    Screen.MousePointer = 11
    If OP = "A" Then
        COD = BuscaCodigo("Select tcl_cod_componente from TAB_COMPONENTE_LOGRADOURO where tcl_descricao_componente ='" & Mid(lstExist.Text, 2 + InStr(1, lstExist, "-"))) & "'"
        If Trim(txtEdit) <> "" Then
            If Bdados.AtualizaDados("TAB_COMPONENTE_LOGRADOURO", _
             txtEdit, "tcl_descricao_componente", "tcl_cod_componente = " & COD) Then
                Informa "Registro Atualizado."
            Else
                Erro "Erro ao Atualizar."
            End If
        Else
            If lstExist.ListCount > 0 Then
                If Confirma("Deseja mesmo Excluir: " & lstExist & ".") Then
                    If Bdados.DeletaDados("TAB_COMPONENTE_LOGRADOURO", "tcl_cod_componente = " & COD) Then
                        Informa "Registro Excluido."
                    Else
                        Erro "Erro ao apagar."
                    End If
                End If
            End If
        End If
        Call cmdNovo_Click
    
    ElseIf OP = "N" Then
        If Trim(txtEdit) <> "" Then
            COD = GeraCodComponente(BuscaCodigo("Select tcl_grupo from TAB_COMPONENTE_LOGRADOURO where tcl_descricao_componente ='" & Mid(lstExist.Text, 2 + InStr(1, lstExist, "-"))) & "'")
            Informa "Novo Registro Gravado."
        Else
            Avisa "Entrada Inválida."
        End If
        Call cmdNovo_Click
    Else
        Informa "Selecione um Grupo"
        cboGrupo.ListIndex = -1
        cboGrupo.SetFocus
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdNovo_Click()
    On Error GoTo trata
    txtEdit = ""
    cboGrupo.ListIndex = -1
    cboGrupo.SetFocus
trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo trata

    cabVisual.Exibir Bdados, Cod_Form, App.Path
    
    Call AtualizaCombo(Bdados, cboGrupo, "SELECT Tgc_NOME FROM TAB_GRUPO_COMPONENTE")
    
    Screen.MousePointer = 0
    Exit Sub
trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub lstExist_DblClick()
    txtEdit = Mid(lstExist.Text, 2 + InStr(1, lstExist, "-"))
    txtEdit.SetFocus
    OP = "A"
End Sub

Private Sub lstExist_GotFocus()
    txtEdit.Text = ""
End Sub

Private Sub lstExist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lstExist_DblClick
End Sub

Private Sub AtualizaLista()
    On Error GoTo trata
    Dim RS As VSRecordset
    Screen.MousePointer = 11
    
    lstExist.Clear

    Bdados.FechaTabela RS
    Screen.MousePointer = 0
    
    Exit Sub
trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtEdit_GotFocus()
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit)
    If txtEdit = "" Then OP = "N"
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    'KeyAscii = Maiuscula(KeyAscii)
    If KeyAscii = 13 Then Call cmdGravar_Click
End Sub
