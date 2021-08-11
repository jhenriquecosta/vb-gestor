VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.1#0"; "CABECALHO.OCX"
Begin VB.Form CSEG101 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFF5EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CSEG101.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1665
      TabIndex        =   3
      Top             =   2250
      Width           =   4800
   End
   Begin VB.TextBox txtSerial 
      Alignment       =   1  'Right Justify
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
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   2
      ToolTipText     =   "Duração da Atualização"
      Top             =   1815
      Width           =   1050
   End
   Begin VB.TextBox txtDuracao 
      Alignment       =   1  'Right Justify
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
      Left            =   3660
      MaxLength       =   3
      TabIndex        =   1
      ToolTipText     =   "Duração da Atualização"
      Top             =   1815
      Width           =   810
   End
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
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
      Left            =   1665
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "Data da Atualização"
      Top             =   1815
      Width           =   1065
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   1349
      CorFundo        =   16774636
      CorFrente       =   12632064
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destinho do arquivo"
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
      Left            =   105
      TabIndex        =   11
      Top             =   2295
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duração"
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
      Index           =   2
      Left            =   4740
      TabIndex        =   10
      Top             =   1890
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duração"
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
      Index           =   1
      Left            =   2970
      TabIndex        =   9
      Top             =   1890
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Atualização"
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
      Index           =   0
      Left            =   105
      TabIndex        =   8
      Top             =   1875
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Configurações do Disco"
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
      Left            =   75
      TabIndex        =   7
      Top             =   1455
      Width           =   6375
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste Teste "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1110
      TabIndex        =   6
      Top             =   795
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   5520
      TabIndex        =   5
      ToolTipText     =   "Deseja sair?"
      Top             =   2760
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CSEG101.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CSEG101.frx":08E6
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdGerar 
      Height          =   435
      Left            =   4485
      TabIndex        =   4
      ToolTipText     =   "Gerar disco"
      Top             =   2760
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CSEG101.frx":0902
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CSEG101.frx":091E
      Caption         =   "&Gerar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
End
Attribute VB_Name = "CSEG101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtData.SetFocus
End Sub

Private Sub cmdGerar_Click()
    On Error GoTo Trata
    
    Screen.MousePointer = 11
    If Instala.GeraArquivoAtualizador(txtData, txtDuracao, txtSerial, Temp.PegaParametro(Bdados, "MUNICIPIO"), Mid(Drive1.Drive, 1, 2)) Then Util.informa "Disco gerado."
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Trata
    Dim sql As String
    Dim RSs As Object
    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\"
    txtData = Format(Date, "dd/mm/yyyy")
    txtDuracao = Format("60", "000")
    txtSerial = Format(CStr(CLng(Instala.PegaSerialDisco(Bdados)) + 1), "00000")
    Screen.MousePointer = 0
    sql = "SELECT * FROM TAB_ATUALIZACAO"
    If Bdados.AbreTabela(sql, RSs) Then
        lblStatus = "Atualizado: " & RSs(1) & _
        ". Vencimento: " & DateAdd("d", RSs(2), RSs(1)) & " (" & RSs(2) & " dias). " & _
        "Restante: " & DateDiff("d", Date, DateAdd("d", RSs(2), RSs(1))) & " dias."
    End If
    Bdados.FechaTabela RSs
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtData_LostFocus()
    If Not IsDate(txtData) Then txtData = Format(txtData, "##/##/####")

End Sub

Private Sub txtDuracao_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtSerial_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtDuracao_LostFocus()
    txtDuracao = Format(txtDuracao, "000")
End Sub

Private Sub txtSerial_LostFocus()
    txtSerial = Format(txtSerial, "00000")
End Sub
