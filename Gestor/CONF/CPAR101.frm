VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.1#0"; "CABECALHO.OCX"
Begin VB.Form CPAR101 
   BackColor       =   &H00FFF5EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CPAR101.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7350
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
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "Parâmetro"
      Top             =   960
      Width           =   6090
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
      Left            =   1185
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Descrição"
      Top             =   1380
      Width           =   6090
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   1349
      CorFundo        =   16774636
      CorFrente       =   12632064
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   3225
      Left            =   90
      TabIndex        =   4
      Top             =   2625
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   5689
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Parâmetros"
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
      Left            =   90
      TabIndex        =   9
      Top             =   2340
      Width           =   7185
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
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
      Left            =   405
      TabIndex        =   8
      Top             =   1425
      Width           =   690
   End
   Begin Threed.SSCommand cmdGravar 
      Height          =   435
      Left            =   5175
      TabIndex        =   2
      Top             =   1815
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CPAR101.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CPAR101.frx":08E6
      Caption         =   "&Salvar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdDeletar 
      Height          =   435
      Left            =   6270
      TabIndex        =   3
      Top             =   1815
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CPAR101.frx":0902
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CPAR101.frx":091E
      Caption         =   "&Apagar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parâmetro"
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
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   1035
      Width           =   915
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   435
      Left            =   5010
      TabIndex        =   5
      Top             =   5940
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CPAR101.frx":093A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CPAR101.frx":0956
      Caption         =   "&Novo"
      ButtonStyle     =   4
      PictureAlignment=   6
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   6195
      TabIndex        =   6
      ToolTipText     =   "Deseja sair?"
      Top             =   5940
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CPAR101.frx":0972
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CPAR101.frx":098E
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
End
Attribute VB_Name = "CPAR101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDeletar_Click()
    If txtPar.Enabled = False And Trim(txtPar) <> "" Then
        If Util.Confirma("Deseja mesmo apagar " & txtPar & "?") Then
            If Bdados.DeletaDados("TAB_PARAMETRO", "TPR_TIPO='" & txtPar & "'") Then
                Util.informa "Parâmentro apagado."
            Else
                Util.avisa "Parâmetro não apagado."
            End If
        End If
    Else
        Util.avisa "Selecione um registro gravado."
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
    Call Util.MontaGrid(Bdados, Grid, "SELECT TPR_TIPO AS Parmâmetro, TPR_DESCRICAO AS Descrição FROM TAB_PARAMETRO", 2500, 4300)
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Util.OrdenaGrid(Grid, ColumnHeader)
End Sub

Private Sub Grid_DblClick()
    On Error GoTo Trata
    txtDes.SetFocus
    
    txtPar = Grid.SelectedItem.Text
    txtDes = Grid.SelectedItem.SubItems(1)
    
    txtPar.Enabled = False
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtPar_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Sub Atualizar(CODIGO As String)
    On Error GoTo Trata
    Dim valores As String
    Dim campos As String
    
    campos = "TPR_TIPO,TPR_DESCRICAO"
    valores = Bdados.PreparaValor(txtPar, txtDes)
    
    Call Bdados.GravaDados("TAB_PARAMETRO", valores, campos, _
    "TPR_TIPO = '" & txtPar & "'")

    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.avisa "Erro: " & Err.Number & " - " & Err.Description & "."
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
            Util.informa "Operação realizada."
        End If
    End If
    
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
    
    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"
    
    AtualizaG
    Screen.MousePointer = 0
    
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

Private Sub txtPar_LostFocus()
    Dim I As Integer
    txtPar = Trim(txtPar)
    For I = 1 To Grid.ListItems.Count
        If Grid.ListItems(I).Text = Trim(txtPar) Then
            txtDes.Text = Grid.ListItems(I).ListSubItems.Item(1).Text
            Exit For
        End If
    Next
    
End Sub
