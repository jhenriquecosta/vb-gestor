VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form PTBS201 
   BackColor       =   &H00DDF1FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PTBS201.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMunicipio 
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
      Left            =   1065
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1140
      Width           =   6225
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
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "Descrição"
      Top             =   1950
      Width           =   6225
   End
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
      Left            =   1065
      MaxLength       =   5
      TabIndex        =   1
      Tag             =   "Parâmetro"
      Top             =   1530
      Width           =   1380
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   3105
      Left            =   75
      TabIndex        =   3
      Top             =   2640
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5477
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
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   1349
      CorFundo        =   14545407
      CorFrente       =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Município"
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
      Index           =   4
      Left            =   195
      TabIndex        =   11
      Top             =   1215
      Width           =   795
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Tabela de Bairros"
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
      Top             =   2355
      Width           =   7230
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
      Left            =   570
      TabIndex        =   9
      Top             =   1995
      Width           =   405
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Bairro"
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
      Top             =   810
      Width           =   7230
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
      Left            =   420
      TabIndex        =   7
      Top             =   1620
      Width           =   570
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   435
      Left            =   4110
      TabIndex        =   6
      Top             =   5835
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      BackColor       =   12632256
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Novo"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdGravar 
      Height          =   435
      Left            =   5190
      TabIndex        =   4
      Top             =   5835
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      BackColor       =   12632256
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Gravar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   6270
      TabIndex        =   5
      Top             =   5835
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      BackColor       =   12632256
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Sair"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
End
Attribute VB_Name = "PTBS201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CodMunicSelec As Integer

Private Sub cboMunicipio_Click()
    CodMunicSelec = Bdados.BuscaCodigo("SELECT TMU_COD_MUNICIPIO FROM TAB_MUNICIPIO WHERE TMU_NOME='" & cboMunicipio & "'")
    AtualizaG
End Sub


Private Sub cmdNovo_Click()
    txtPar.Enabled = True
    txtPar = Bdados.BuscaCodigo("SELECT MAX(TBA_COD_BAIRRO) FROM TAB_BAIRRO WHERE TBA_TMU_COD_MUNICIPIO=" & CodMunicSelec) + 1
    txtDes = ""

    txtPar.SetFocus
End Sub

Private Sub AtualizaG()
    Call Util.MontaGrid(Bdados, Grid, "SELECT TBA_COD_BAIRRO AS Código, TBA_NOME AS Bairro FROM TAB_BAIRRO WHERE TBA_TMU_COD_MUNICIPIO=" & CodMunicSelec, 800, 5800)
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Util.OrdenaGrid(Grid, ColumnHeader)
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Grid_DblClick
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

Private Sub txtDes_GotFocus()
    Edita.SelecionaTexto txtDes
End Sub

Private Sub txtPar_GotFocus()
    Edita.SelecionaTexto txtPar
End Sub

Private Sub txtPar_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, 1)
End Sub

Sub Atualizar(Codigo As String)
    On Error GoTo Trata
    Dim Valores As String
    Dim Campos As String
    
    Campos = "TBA_TMU_COD_MUNICIPIO,TBA_COD_BAIRRO,TBA_NOME"
    Valores = Bdados.PreparaValor(CodMunicSelec, txtPar, txtDes)
    
    Call Bdados.GravaDados("TAB_BAIRRO", Valores, Campos, "TBA_COD_BAIRRO = " & txtPar & " AND TBA_TMU_COD_MUNICIPIO=" & CodMunicSelec)

    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo Trata
    
    If txtPar.Enabled = False And Trim(txtDes) = "" Then
        If Util.Confirma("Deseja mesmo apagar " & txtPar & "?") Then
            If Bdados.DeletaDados("TAB_BAIRRO", "TBA_TMU_COD_MUNICIPIO=" & CodMunicSelec & " AND TBA_COD_BAIRRO=" & txtPar) Then
                Util.Informa "Bairro apagado."
                AtualizaG
                Call cmdNovo_Click
            End If
        End If
    Else
        If Edita.CriticaCampos(Me) Then
            If Util.Confirma("Deseja salvar os dados de " & txtPar & "?") Then
                Screen.MousePointer = 11
                Call Atualizar(txtPar)
                txtPar = ""
                txtDes = ""
                Util.Informa "Operação realizada."
                cboMunicipio_Click
                txtPar.Enabled = True
                txtPar.SetFocus
            End If
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
    Edita.AtualizaCombo Bdados, cboMunicipio, "SELECT TMU_NOME FROM TAB_MUNICIPIO ORDER BY TMU_NOME"
    If cboMunicipio.ListCount Then cboMunicipio.ListIndex = 0
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
            txtDes.Text = Grid.ListItems(I).ListSubItems.Item(1).Text
            Exit For
        End If
    Next
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        If Err.Number = 35600 Then Resume Next
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If

End Sub
