VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form PMUS301 
   BackColor       =   &H00DDF1FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PMUS301.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
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
      Top             =   1155
      Width           =   2010
   End
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Nome do Usuário"
      Top             =   1545
      Width           =   6330
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   855
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
      TabIndex        =   5
      Top             =   1245
      Width           =   570
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   6195
      TabIndex        =   4
      ToolTipText     =   "Deseja sair?"
      Top             =   1980
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PMUS301.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PMUS301.frx":08E6
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdSalvar 
      Height          =   435
      Left            =   3930
      TabIndex        =   2
      ToolTipText     =   "Excluir"
      Top             =   1980
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PMUS301.frx":0902
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PMUS301.frx":091E
      Caption         =   "&Excluir"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdCancelar 
      Height          =   435
      Left            =   5070
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   1980
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PMUS301.frx":093A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PMUS301.frx":0956
      Caption         =   "&Cancelar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
End
Attribute VB_Name = "PMUS301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GravaDados(COD As String, Nome As String)
    On Error GoTo Trata
    
    Bdados.abretrans
    
    If Bdados.DeletaDados("TAB_ACESSO_USUARIO", "TAU_TUS_COD_USUARIO = '" & COD & "'") And Bdados.DeletaDados("TAB_USUARIO", "TUS_COD_USUARIO = '" & COD & "'") Then
        Bdados.gravatrans
        Call Util.Informa("Registro excluido com sucesso.")
    Else
        Bdados.cancelatrans
    End If
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If

End Sub

Private Sub cmdCancelar_Click()
    txtCodigo = ""
    txtNome = ""
    txtCodigo.SetFocus
    Screen.MousePointer = 0
End Sub

Private Sub cmdSalvar_Click()
    On Error GoTo Trata
    
    If Trim(txtCodigo) = "" Then
        Util.Informa "Selecione um Usuário."
        txtCodigo.SetFocus
        Exit Sub
    End If
    
    If Util.Confirma("Deseja mesmo excluir o usuário " & (txtNome) & " ?") Then
        Screen.MousePointer = 11
        Call GravaDados(txtCodigo, txtNome)
        Call cmdCancelar_Click
    End If
    Screen.MousePointer = 0

Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Trata
    
    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"

    Screen.MousePointer = 0

    Exit Sub
    
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
        
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtCodigo_GotFocus()
    Call cmdCancelar_Click
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    On Error GoTo Trata
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And Trim(txtCodigo) <> "" Then
        txtNome = UCase(Seguranca.ExisteUsuario(Bdados, txtCodigo))
        If Trim(txtNome) = "" Then
            Util.Informa "Usuário '" & txtCodigo & "' não Cadastrado."
            Call cmdCancelar_Click
        Else
            txtNome.SetFocus
        End If
    End If
    DoEvents
    
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
