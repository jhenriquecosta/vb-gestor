VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.1#0"; "CABECALHO.OCX"
Begin VB.Form PTBS401 
   BackColor       =   &H00DDF1FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PTBS401.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstExist 
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
      Height          =   3930
      ItemData        =   "PTBS401.frx":08CA
      Left            =   60
      List            =   "PTBS401.frx":08CC
      TabIndex        =   1
      Top             =   1635
      Width           =   4980
   End
   Begin VB.TextBox txtEdit 
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
      Left            =   510
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "Data de Nascimento"
      ToolTipText     =   "Data de nascimento do aluno"
      Top             =   930
      Width           =   4530
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   1349
      CorFundo        =   14545407
      CorFrente       =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Left            =   90
      TabIndex        =   6
      Top             =   990
      Width           =   360
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Tabela de Tipos"
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
      TabIndex        =   5
      Top             =   1350
      Width           =   4980
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4065
      TabIndex        =   3
      ToolTipText     =   "Deseja sair?"
      Top             =   5670
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PTBS401.frx":08CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PTBS401.frx":08EA
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdGravar 
      Height          =   435
      Left            =   3015
      TabIndex        =   2
      ToolTipText     =   "Salvar Informações"
      Top             =   5670
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PTBS401.frx":0906
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PTBS401.frx":0922
      Caption         =   "&Gravar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   435
      Left            =   1965
      TabIndex        =   4
      ToolTipText     =   "Novo"
      Top             =   5670
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PTBS401.frx":093E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PTBS401.frx":095A
      Caption         =   "&Novo"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
End
Attribute VB_Name = "PTBS401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OP As String

Private Sub cmdGravar_Click()
    On Error GoTo Trata
    Dim I As String
    Dim COD As String
    Screen.MousePointer = 11
    If OP = "A" Then
        I = InStr(1, lstExist.Text, "-", vbTextCompare)
        COD = Mid(lstExist.Text, 1, I - 1)
        If Trim(txtEdit) <> "" Then
            If Bdados.AtualizaDados("TAB_TIPO_LOGR", Bdados.preparavalor(txtEdit), _
            "TTL_NOME", "TTL_COD_TIP_LOGR = " & COD) Then Util.Informa "Registro Atualizado."
        Else
            If Util.Confirma("Deseja mesmo Excluir: " & lstExist & ".") Then
                If Bdados.DeletaDados("TAB_TIPO_LOGR", "TTL_COD_TIP_LOGR = " & COD) Then Util.Informa "Registro Excluido."
            End If
        End If
        Call cmdNovo_Click
        AtualizaLista
    ElseIf OP = "N" Then
        If Trim(txtEdit) <> "" Then
            lstExist.ListIndex = lstExist.ListCount - 1
            I = InStr(1, lstExist.Text, "-", vbTextCompare)
            If I = 0 Then
                COD = 1
            Else
                COD = Mid(lstExist.Text, 1, I - 1) + 1
            End If
            If Bdados.InsereDados("TAB_TIPO_LOGR", _
            Bdados.preparavalor(COD, txtEdit), "TTL_COD_TIP_LOGR,TTL_NOME") Then
                AtualizaLista
                Util.Informa "Novo Registro Gravado."
            End If
        Else
            Util.Avisa "Entrada Inválida."
        End If
        Call cmdNovo_Click
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdNovo_Click()
    On Error GoTo Trata
    OP = "N"
    txtEdit = ""
    txtEdit.SetFocus
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Trata

    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"
    AtualizaLista
    OP = "N"
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

Private Sub lstExist_DblClick()
    Dim I As Integer
    I = InStr(1, lstExist.Text, "-", vbTextCompare)
    txtEdit = Mid(lstExist.Text, I + 2)
    txtEdit.SetFocus
    OP = "A"
End Sub

Private Sub lstExist_GotFocus()
    txtEdit.Text = ""
End Sub

Private Sub AtualizaLista()
    On Error GoTo Trata
    Dim RS As Object
    Screen.MousePointer = 11
    lstExist.Clear
    If Bdados.AbreTabela("SELECT * FROM TAB_TIPO_LOGR", RS) Then
        Do Until RS.EOF
            If Not IsNull(RS(0)) Then
                If Trim(RS(0)) <> "" Then
                    lstExist.AddItem RS(0) & " - " & RS(1)
                End If
            End If
            RS.MoveNext
        Loop
    End If
    Bdados.FechaTabela RS
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtEdit_GotFocus()
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdGravar_Click
End Sub
