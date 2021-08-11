VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.1#0"; "CABECALHO.OCX"
Begin VB.Form CAPL201 
   BackColor       =   &H00FFF5EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CAPL201.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSis 
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
      ItemData        =   "CAPL201.frx":08CA
      Left            =   1050
      List            =   "CAPL201.frx":08CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Sistema"
      Top             =   1965
      Width           =   1395
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
      Height          =   660
      Left            =   1050
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2370
      Width           =   6180
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
      Left            =   1050
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Nome"
      Top             =   1545
      Width           =   3945
   End
   Begin VB.TextBox txtCod 
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
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Código"
      Top             =   1125
      Width           =   1395
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   2610
      Left            =   60
      TabIndex        =   6
      Top             =   3480
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   4604
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
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   1349
      CorFundo        =   16774636
      CorFrente       =   12632064
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
      Index           =   3
      Left            =   270
      TabIndex        =   14
      Top             =   2430
      Width           =   690
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Módulo selecionado"
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
      TabIndex        =   13
      Top             =   810
      Width           =   7185
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema"
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
      Left            =   405
      TabIndex        =   12
      Top             =   2025
      Width           =   555
   End
   Begin Threed.SSCommand cmdDeletar 
      Height          =   435
      Left            =   6270
      TabIndex        =   5
      Top             =   1830
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL201.frx":08CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL201.frx":08EA
      Caption         =   "&Apagar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdGravar 
      Height          =   435
      Left            =   5205
      TabIndex        =   4
      Top             =   1830
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL201.frx":0906
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL201.frx":0922
      Caption         =   "&Salvar"
      ButtonStyle     =   4
      PictureAlignment=   6
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
      Index           =   1
      Left            =   390
      TabIndex        =   11
      Top             =   1200
      Width           =   570
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
      Index           =   0
      Left            =   525
      TabIndex        =   10
      Top             =   1605
      Width           =   405
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   6225
      TabIndex        =   8
      ToolTipText     =   "Deseja sair?"
      Top             =   6165
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL201.frx":093E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL201.frx":095A
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   435
      Left            =   5130
      TabIndex        =   7
      Top             =   6165
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL201.frx":0976
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL201.frx":0992
      Caption         =   "&Novo"
      ButtonStyle     =   4
      PictureAlignment=   6
      ShapeSize       =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Tabela de Módulos"
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
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   3195
      Width           =   7185
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "CAPL201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COD As String

Private Sub cmdDeletar_Click()
    If COD <> "" And Trim(txtCod) <> "" Then
        If Util.Confirma("Deseja mesmo apagar '" & txtCod & "'?") Then
            If Bdados.DeletaDados("TAB_MODULO", "TMO_COD_MODULO='" & txtCod & "'") Then
                Util.informa "Registro apagado."
            Else
                Util.avisa "Registro não apagado."
            End If
        End If
    Else
        Util.avisa "Selecione um registro gravado."
    End If
    Call cmdNovo_Click
End Sub

Private Sub cmdNovo_Click()
    COD = ""
    Edita.LimpaCampos Me
    AtualizaG
    
    txtCod.SetFocus
End Sub

Private Sub AtualizaG()
    Call Util.MontaGrid(Bdados, Grid, _
    "SELECT TMO_COD_MODULO AS Código, TMO_NOME as Nome, TMO_TSI_COD_SISTEMA AS Sistema, TMO_DESCR AS Descrição FROM TAB_MODULO", 800, 2500, 800, 2600)
    
    Call Edita.AtualizaCombo(Bdados, cboSis, "SELECT TSI_COD_SISTEMA FROM TAB_SISTEMA")
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Util.OrdenaGrid(Grid, ColumnHeader)
End Sub


Private Sub Grid_DblClick()
    On Error GoTo Trata
    txtCod.SetFocus
    
    txtCod = Grid.SelectedItem.Text
    txtNome = Grid.SelectedItem.SubItems(1)
    cboSis.ListIndex = Edita.ListIndexDe(cboSis, Grid.SelectedItem.SubItems(2))
    txtDes = Grid.SelectedItem.SubItems(3)
    
    COD = Grid.SelectedItem.Text
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Atualizar(CODIGO As String)
    On Error GoTo Trata
    Dim valores As String
    Dim campos As String
    
    If COD = "" Then COD = CODIGO
    
    campos = "TMO_COD_MODULO,TMO_NOME,TMO_TSI_COD_SISTEMA,TMO_DESCR"
    valores = Bdados.PreparaValor(txtCod, txtNome, cboSis, txtDes)
    
    Call Bdados.GravaDados("TAB_MODULO", valores, campos, _
    "TMO_COD_MODULO = '" & COD & "'")

    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo Trata
    txtCod = Trim(txtCod)
    If Edita.CriticaCampos(Me) Then
        If IsNumeric(txtCod) Or Len(txtCod) < 4 Then
            Util.avisa "Código deve conter 4 caracteres."
        Else
            If Util.Confirma("Deseja salvar os dados de '" & txtCod & "'?") Then
                Screen.MousePointer = 11
                Call Atualizar(txtCod)
                Call cmdNovo_Click
                Util.informa "Operação realizada."
            End If
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

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Grid_DblClick
End Sub

Private Sub txtCod_LostFocus()
    On Error GoTo Trata
    Dim I As Integer
    txtCod = Trim(txtCod)
    If COD = "" Then
        For I = 1 To Grid.ListItems.Count
            If Grid.ListItems(I).Text = Trim(txtCod) Then
                txtNome.Text = Grid.ListItems(I).ListSubItems.Item(1).Text
                cboSis.ListIndex = Edita.ListIndexDe(cboSis, Grid.ListItems(I).ListSubItems.Item(2).Text)
                txtDes.Text = Grid.ListItems(I).ListSubItems.Item(3).Text
                COD = txtCod
                Exit For
            End If
        Next
    End If
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        If Err.Number = 35600 Then Resume Next

        Util.avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
        
    End If
End Sub

