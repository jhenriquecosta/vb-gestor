VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form PMUS401 
   BackColor       =   &H00DDF1FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PMUS401.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   900
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Nome do Usuário"
      Top             =   1560
      Width           =   6315
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   900
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Código do Usuário"
      Top             =   1140
      Width           =   2010
   End
   Begin VB.ListBox lstUs 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3345
      ItemData        =   "PMUS401.frx":08CA
      Left            =   75
      List            =   "PMUS401.frx":08D1
      TabIndex        =   2
      Top             =   2280
      Width           =   7140
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   1349
      CorFundo        =   14545407
      CorFrente       =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Resultado de Pesquisa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   60
      TabIndex        =   9
      Top             =   1965
      Width           =   7170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   345
      TabIndex        =   8
      Top             =   1590
      Width           =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Usuário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   825
      Width           =   7170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   255
      TabIndex        =   6
      Top             =   1215
      Width           =   570
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   6165
      TabIndex        =   5
      ToolTipText     =   "Sair"
      Top             =   5715
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PMUS401.frx":08DC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PMUS401.frx":08F8
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdConsultar 
      Height          =   435
      Left            =   3945
      TabIndex        =   3
      ToolTipText     =   "Localizar Usuário"
      Top             =   5715
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PMUS401.frx":0914
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PMUS401.frx":0930
      Caption         =   "&Consultar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdLimpar 
      Height          =   435
      Left            =   5055
      TabIndex        =   4
      ToolTipText     =   "Limpar Campos"
      Top             =   5715
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PMUS401.frx":094C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PMUS401.frx":0968
      Caption         =   "&Limpar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
End
Attribute VB_Name = "PMUS401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConsultar_Click()
    On Error GoTo Trata
    Dim RS As Object
    Dim SQL As String
    Dim JA As Boolean
    
    JA = False
    lstUs.Clear
    
    SQL = "SELECT TUS_COD_USUARIO, TUS_NOME FROM TAB_USUARIO "
    
    If Trim(txtCodigo) <> "" Then
        If JA Then
            SQL = SQL & " AND TUS_COD_USUARIO = '" & (Trim(txtCodigo)) & "' "
        Else
            SQL = SQL & " WHERE TUS_COD_USUARIO = '" & (Trim(txtCodigo)) & "' "
            JA = True
        End If
    End If
    If Trim(txtNome) <> "" Then
        If JA Then
            SQL = SQL & " AND TUS_NOME LIKE '%" & (Trim(txtNome)) & "%' "
        Else
            SQL = SQL & " WHERE TUS_NOME LIKE '%" & (Trim(txtNome)) & "%' "
            JA = True
        End If
    End If
    SQL = SQL & " ORDER BY TUS_NOME"

    If Bdados.AbreTabela(SQL, RS) Then
        Do Until RS.EOF
            lstUs.AddItem CStr((RS(0)) & " - " & (RS(1)))
            RS.MoveNext
        Loop
    Else
        Util.Informa "Usuário não encontrado."
    End If
    DoEvents
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    lstUs.Clear
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Dim RS As Object
    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"
    lstUs.Clear
    Screen.MousePointer = 0
    If Bdados.AbreTabela("SELECT TUS_COD_USUARIO, TUS_NOME FROM TAB_USUARIO", RS) Then
        Do Until RS.EOF
            lstUs.AddItem CStr(RS(0)) & " - " & (RS(1))
            RS.MoveNext
        Loop
    End If
    Bdados.FechaTabela RS
    DoEvents
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
