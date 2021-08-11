VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "CABECALHO.OCX"
Begin VB.Form CAPL301 
   BackColor       =   &H00FFF5EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CAPL301.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMod 
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
      ItemData        =   "CAPL301.frx":08CA
      Left            =   1050
      List            =   "CAPL301.frx":08CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Módulo"
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
      MaxLength       =   60
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
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Código"
      Top             =   1125
      Width           =   1395
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   3165
      Left            =   60
      TabIndex        =   6
      Top             =   3480
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   5583
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
   Begin Threed.SSCommand cmdImprimir 
      Height          =   435
      Left            =   5190
      TabIndex        =   15
      Top             =   6705
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL301.frx":08CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL301.frx":08EA
      Caption         =   "&Imprimir"
      ButtonStyle     =   4
      PictureAlignment=   6
      ShapeSize       =   1
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
      Caption         =   " Formulário selecionado"
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
      Caption         =   "Módulo"
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
      Left            =   435
      TabIndex        =   12
      Top             =   2025
      Width           =   510
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
      MouseIcon       =   "CAPL301.frx":0906
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL301.frx":0922
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
      MouseIcon       =   "CAPL301.frx":093E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL301.frx":095A
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
      Top             =   6705
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL301.frx":0976
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL301.frx":0992
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   435
      Left            =   4140
      TabIndex        =   7
      Top             =   6705
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL301.frx":09AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL301.frx":09CA
      Caption         =   "&Novo"
      ButtonStyle     =   4
      PictureAlignment=   6
      ShapeSize       =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Tabela de Formulários"
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
Attribute VB_Name = "CAPL301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COD As String

Private Sub cmdDeletar_Click()
    If COD <> "" And Trim(txtCod) <> "" Then
        If Util.Confirma("Deseja mesmo apagar '" & cboMod & txtCod & "'?") Then
            If Bdados.DeletaDados("TAB_FORMULARIO", "TFO_COD_FORMULARIO='" & txtCod & "' AND TFO_TMO_COD_MODULO='" & cboMod & "'") Then
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

Private Sub cmdImprimir_Click()
    Screen.MousePointer = vbHourglass
    With Relatorio
        .DefinirArquivo Bdados, App.Path & "\" & Me.Name & ".rpt"
        .Formulas "VT_Sistema", "'" & Temp.PegaParametro(Bdados, "SISTEMA") & "'"
        .Formulas "VT_Descricao", "'" & Temp.PegaParametro(Bdados, "DESCRICAO") & "'"
        .Visualizar
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdNovo_Click()
    COD = ""
    Edita.LimpaCampos Me
    AtualizaG
    
    txtCod.SetFocus
End Sub

Private Sub AtualizaG()
    Call Util.MontaGrid(Bdados, Grid, _
    "SELECT TFO_COD_FORMULARIO AS Código, TFO_NOME as Nome, TFO_TMO_COD_MODULO AS Módulo, TFO_DESCR AS Descrição FROM TAB_FORMULARIO", 800, 2800, 800, 2400)
    
    Call Edita.AtualizaCombo(Bdados, cboMod, "SELECT TMO_COD_MODULO FROM TAB_MODULO")
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Util.OrdenaGrid(Grid, ColumnHeader)
End Sub


Private Sub Grid_DblClick()
    On Error GoTo Trata
    txtCod.SetFocus
    
    txtCod = Grid.SelectedItem.Text
    txtNome = Grid.SelectedItem.SubItems(1)
    cboMod.ListIndex = Edita.ListIndexDe(cboMod, Grid.SelectedItem.SubItems(2))
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
    KeyAscii = Edita.AceitaDig(KeyAscii, 1)
End Sub

Private Sub Atualizar(CODIGO As String)
    On Error GoTo Trata
    Dim valores As String
    Dim campos As String
    
    If COD = "" Then COD = CODIGO
    
    campos = "TFO_COD_FORMULARIO,TFO_NOME,TFO_TMO_COD_MODULO,TFO_DESCR"
    valores = Bdados.PreparaValor(txtCod, txtNome, cboMod, txtDes)
    
    Call Bdados.GravaDados("TAB_FORMULARIO", valores, campos, _
    "TFO_COD_FORMULARIO='" & txtCod & "' AND TFO_TMO_COD_MODULO='" & cboMod & "'")

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
        If (Not IsNumeric(txtCod)) Or Len(txtCod) < 3 Then
            Util.avisa "Código deve conter 3 números."
        Else
            If Util.Confirma("Deseja salvar os dados de '" & cboMod & txtCod & "'?") Then
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

