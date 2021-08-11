VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form PMUS201 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PMUS201.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
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
      Left            =   885
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Código do Usuário"
      Top             =   1020
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
      Left            =   885
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Nome do Usuário"
      Top             =   1440
      Width           =   6330
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   1138
      Icone           =   "PMUS201.frx":08CA
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   405
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   714
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdCancelar 
      Height          =   405
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   714
      Caption         =   "&Cancelar"
      Acao            =   9
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   405
      Left            =   3240
      TabIndex        =   4
      Top             =   1800
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   714
      Caption         =   "&Sair"
      Acao            =   7
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   390
      TabIndex        =   7
      Top             =   1485
      Width           =   405
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
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
      Left            =   0
      TabIndex        =   6
      Top             =   720
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
      Top             =   1110
      Width           =   570
   End
End
Attribute VB_Name = "PMUS201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GravaDados(COD As String)
    On Error GoTo Trata
    Dim Val As String
    Dim S As String
    S = Temp.PegaParametro(BDados, "SENHA INICIAL")
    Val = BDados.PreparaValor(COD, Seguranca.Criptografa(S))
    
    Call BDados.AtualizaDados("TAB_USUARIO", Val, _
    "TUS_COD_USUARIO,TUS_SENHA", "TUS_COD_USUARIO = '" & COD & "'")
    Call Util.Informa("Senha atualizada com sucesso.")
    
    Exit Sub
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If

End Sub

Private Sub cmdCancelar_Click()
    txtCodigo = ""
    txtNome = ""
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
    Screen.MousePointer = 0
End Sub

Private Sub cmdSalvar_Click()
    On Error GoTo Trata
    
    If Trim(txtCodigo) = "" Then
        Util.Informa "Selecione um Usuário."
        txtCodigo.SetFocus
    Else
        If Util.Confirma("Deseja mesmo atualizar a senha do usuário " & (txtNome) & " para a senha padrão '" & Temp.PegaParametro(BDados, "SENHA INICIAL") & "' ?") Then
            Screen.MousePointer = 11
            Call GravaDados(txtCodigo)
            Call cmdCancelar_Click
        End If
    End If
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdVISUAL1_Click()

End Sub

Private Sub Form_Load()
    On Error GoTo Trata
    
 '   Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"

    Screen.MousePointer = 0

    Exit Sub
    
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
        
    End If
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub txtCodigo_GotFocus()
    Call cmdCancelar_Click
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    On Error GoTo Trata
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    On Error GoTo Trata

    If Trim(txtCodigo) <> "" Then
        txtNome = UCase(Seguranca.ExisteUsuario(BDados, txtCodigo))
        If Trim(txtNome) = "" Then
            Util.Informa "Usuário '" & txtCodigo & "' não Cadastrado."
            Call cmdCancelar_Click
        Else
            txtCodigo.Enabled = False
            txtNome.SetFocus
        End If
    End If
    DoEvents
    
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
