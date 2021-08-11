VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form PMUS501 
   BackColor       =   &H00DDF1FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PMUS501.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPar 
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
      Tag             =   "Parâmetro"
      Top             =   1155
      Width           =   2010
   End
   Begin VB.TextBox txtDes 
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
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Descrição"
      Top             =   1575
      Width           =   7230
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   2850
      Left            =   60
      TabIndex        =   4
      Top             =   2820
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   5027
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   1349
      CorFundo        =   14545407
      CorFrente       =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Tabela de Usuários"
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
      TabIndex        =   10
      Top             =   2535
      Width           =   8085
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
      TabIndex        =   9
      Top             =   1605
      Width           =   405
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
      TabIndex        =   8
      Top             =   840
      Width           =   8085
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
      TabIndex        =   7
      Top             =   1230
      Width           =   570
   End
   Begin Threed.SSCommand cmdGravar 
      Height          =   435
      Left            =   6165
      TabIndex        =   2
      Top             =   1980
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PMUS501.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PMUS501.frx":08E6
      Caption         =   "&Salvar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdDeletar 
      Height          =   435
      Left            =   7200
      TabIndex        =   3
      Top             =   1980
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PMUS501.frx":0902
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PMUS501.frx":091E
      Caption         =   "&Apagar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   435
      Left            =   6075
      TabIndex        =   5
      Top             =   5760
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PMUS501.frx":093A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PMUS501.frx":0956
      Caption         =   "&Novo"
      ButtonStyle     =   4
      PictureAlignment=   6
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   7155
      TabIndex        =   6
      ToolTipText     =   "Deseja sair?"
      Top             =   5760
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PMUS501.frx":0972
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PMUS501.frx":098E
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
End
Attribute VB_Name = "PMUS501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDeletar_Click()
    If txtPar.Enabled = False And Trim(txtPar) <> "" Then
        If Util.Confirma("Deseja mesmo apagar " & txtPar & "?") Then
            
            Bdados.abretrans
            If Bdados.DeletaDados("TAB_ACESSO_USUARIO", "TAU_TUS_COD_USUARIO = '" & txtPar & "'") And Bdados.DeletaDados("TAB_USUARIO", "TUS_COD_USUARIO = '" & txtPar & "'") Then
                Bdados.gravatrans
                Util.Informa "Usuário apagado."
            Else
                Bdados.cancelatrans
                Util.Avisa "Usuário não apagado."
            End If
        End If
        
    Else
        Util.Avisa "Selecione um Usuário."
    End If
    Call cmdNovo_Click
End Sub

Private Sub cmdNovo_Click()
    txtPar.Enabled = True
    txtPar = ""
    txtDes = ""
    AtualizaG
    
    txtPar.SetFocus
End Sub

Private Sub AtualizaG()
    Call Util.MontaGrid(Bdados, Grid, "SELECT TUS_COD_USUARIO AS Código, TUS_NOME AS Nome FROM TAB_USUARIO", 2000, 5500)
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Util.OrdenaGrid(Grid, ColumnHeader)
End Sub

Private Sub Grid_DblClick()
    On Error GoTo Trata
    txtDes.SetFocus
    
    txtPar = Grid.SelectedItem.Text
    txtDes = Grid.SelectedItem.ListSubItems.Item(1).Text
    txtDes.SetFocus
    txtPar.Enabled = False
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Grid_DblClick
End Sub

Private Sub txtPar_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Sub Atualizar(Codigo As String)
    On Error GoTo Trata
    Dim Valores As String
    Dim Campos As String
    
    Campos = "TUS_COD_USUARIO,TUS_NOME,TUS_SENHA,TUS_ATIVO"
    Valores = Bdados.PreparaValor(txtPar, txtDes, Seguranca.Criptografa(Temp.PegaParametro(Bdados, "SENHA INICIAL")), -1)
    
    Call Bdados.GravaDados("TAB_USUARIO", Valores, Campos, _
    "TUS_COD_USUARIO = '" & txtPar & "'")

    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo Trata
    
    If Edita.CriticaCampos(Me) Then
        If Util.Confirma("Deseja salvar os dados de " & txtPar & "?") Then
            Screen.MousePointer = 11
            Call Atualizar(txtPar)
            Call cmdNovo_Click
            Util.Informa "Operação realizada."
        End If
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo Trata
    
    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"
    
    AtualizaG
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

Private Sub txtPar_LostFocus()
    On Error GoTo Trata
    
    Dim I As Integer
    txtPar = Trim(txtPar)
    For I = 1 To Grid.ListItems.Count
        If Grid.ListItems(I).Text = Trim(txtPar) Then
            txtDes.Text = Grid.ListItems(I).ListSubItems(1).Text
            Exit For
        End If
    Next
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        If Err.Number = 35600 Then Exit Sub
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub
