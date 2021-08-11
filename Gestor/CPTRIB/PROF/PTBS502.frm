VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form PTBS502 
   BackColor       =   &H00DDF1FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PTBS502.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboGer 
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
      ItemData        =   "PTBS502.frx":08CA
      Left            =   885
      List            =   "PTBS502.frx":08CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Gerência"
      Top             =   1530
      Width           =   8055
   End
   Begin VB.ComboBox cboMun 
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
      ItemData        =   "PTBS502.frx":08CE
      Left            =   885
      List            =   "PTBS502.frx":08D0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "Municipio"
      Top             =   1140
      Width           =   3915
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1349
      CorFundo        =   14545407
      CorFrente       =   255
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   3090
      Left            =   60
      TabIndex        =   2
      Top             =   2205
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   5450
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
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Tabela de Município - Gerências"
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
      Top             =   1920
      Width           =   8865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gerência"
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
      Left            =   135
      TabIndex        =   8
      Top             =   1590
      Width           =   630
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Município e Gerência de Desenvolvimento Regional"
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
      Top             =   810
      Width           =   8865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Município"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1215
      Width           =   645
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   7950
      TabIndex        =   4
      Top             =   5385
      Width           =   975
      _ExtentX        =   1720
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
   Begin Threed.SSCommand cmdGravar 
      Height          =   435
      Left            =   6900
      TabIndex        =   3
      Top             =   5385
      Width           =   975
      _ExtentX        =   1720
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
   Begin Threed.SSCommand cmdNovo 
      Height          =   435
      Left            =   5850
      TabIndex        =   5
      Top             =   5385
      Width           =   975
      _ExtentX        =   1720
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
End
Attribute VB_Name = "PTBS502"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboMun_LostFocus()
    Dim I As Integer
    Dim JA As Boolean
    JA = False
    If cboMun.Text <> "" Then
        For I = 1 To Grid.ListItems.Count
            If cboMun.Text = Grid.ListItems.Item(I).ListSubItems.Item(1).Text Then
                cboGer.ListIndex = Edita.ListIndexDe(cboGer, Grid.ListItems.Item(I).Text)
                JA = True
                Exit For
            End If
        Next
    End If
    If Not JA Then cboGer.ListIndex = 0
    
End Sub

Private Sub cmdNovo_Click()
    cboGer.ListIndex = -1
    cboMun.ListIndex = -1
    DoEvents
    AtualizaG
    cboMun.SetFocus
End Sub

Private Sub AtualizaG()
    Dim SQL As String
    
    SQL = "SELEct  TGR_NOME AS Gerência, TMU_NOME as Município " & _
    "FROM tab_gerencia_municipio,TAB_GERENCIA,TAB_MUNICIPIO " & _
    "WHERE TGM_TGR_COD_GERENCIA=TGR_COD_GERENCIA AND TGM_TMU_COD_MUNICIPIO=TMU_COD_MUNICIPIO ORDER BY TGR_NOME,TMU_NOME ASC "
    
    Call Util.MontaGrid(Bdados, Grid, SQL, 5000, 3000)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Util.OrdenaGrid(Grid, ColumnHeader)
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Grid_DblClick
End Sub


Private Sub Grid_DblClick()
    On Error GoTo Trata
    
    
    cboMun.ListIndex = Edita.ListIndexDe(cboMun, Grid.SelectedItem.ListSubItems.Item(1).Text)
    cboGer.ListIndex = Edita.ListIndexDe(cboGer, Grid.SelectedItem.Text)
    
    
    cboGer.SetFocus
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Sub Atualizar()
    On Error GoTo Trata
    Dim Valores As String
    Dim Campos As String
    Dim M As String
    
    M = CodigoDe(Municipio, cboMun)
    
    Campos = "TGM_TGR_COD_GERENCIA,TGM_TMU_COD_MUNICIPIO"
    Valores = Bdados.PreparaValor(CodigoDe(Gerencia, cboGer), M)
    
    Call Bdados.GravaDados("TAB_GERENCIA_MUNICIPIO", Valores, Campos, _
    " TGM_TMU_COD_MUNICIPIO=" & M)

    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo Trata
    
    If cboMun.Text <> "" Then
        If cboGer.Text = "" Then
            
            If Util.Confirma("Deseja apagar os dados de " & cboMun & "?") Then
                Screen.MousePointer = 11
                If Not Bdados.DeletaDados("TAB_GERENCIA_MUNICIPIO", "TGM_TMU_COD_MUNICIPIO=" & CodigoDe(Municipio, cboMun)) Then
                    Util.Avisa "Dados não apagados."
                Else
                    Util.Informa "Operação realizada."
                End If
            End If
        Else
    
            If Util.Confirma("Deseja salvar os dados de " & cboMun & "?") Then
                Screen.MousePointer = 11
                Call Atualizar
                Util.Informa "Operação realizada."
            End If

        End If
    Else
        Util.Avisa "Selecione um município."
    End If
    Call cmdNovo_Click
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Bdados.Conexao.Errors(0).Number & " - " & Bdados.Conexao.Errors(0).Description & "."
        Screen.MousePointer = 0
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo Trata
    
    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"
    
    AtualizaG
    AtualizaCombos
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

Private Sub AtualizaCombos()
    Dim SQL As String
    SQL = "SELECT TGR_NOME FROM TAB_GERENCIA ORDER BY TGR_NOME ASC"
    
    Call Edita.AtualizaCombo(Bdados, cboGer, SQL)
    
    cboGer.AddItem ""
    
    SQL = "SELECT TMU_NOME FROM TAB_MUNICIPIO ORDER BY TMU_NOME ASC"
    
    Call Edita.AtualizaCombo(Bdados, cboMun, SQL)
End Sub

