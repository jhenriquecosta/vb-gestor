VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "CABECALHO.OCX"
Begin VB.Form TMPU105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ControlBox      =   0   'False
   Icon            =   "TMPU105.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCaixa 
      Interval        =   100
      Left            =   570
      Top             =   4650
   End
   Begin Threed.SSFrame fra 
      Height          =   1545
      Index           =   1
      Left            =   30
      TabIndex        =   8
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2725
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtDes 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1170
         MaxLength       =   200
         TabIndex        =   1
         Tag             =   "Descrição"
         Top             =   570
         Width           =   5895
      End
      Begin VB.TextBox txtPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1170
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "Parâmetro"
         Top             =   150
         Width           =   1035
      End
      Begin Threed.SSPanel lblEscola 
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Parâmetro:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lblEscola 
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   10
         Top             =   630
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Descrição:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdDeletar 
         Height          =   435
         Left            =   5670
         TabIndex        =   3
         Top             =   990
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   767
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   128
         PictureFrames   =   1
         Windowless      =   -1  'True
         MouseIcon       =   "TMPU105.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TMPU105.frx":08E6
         Caption         =   "&Apagar"
         ButtonStyle     =   3
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdGravar 
         Height          =   435
         Left            =   4170
         TabIndex        =   2
         Top             =   990
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   767
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   128
         PictureFrames   =   1
         Windowless      =   -1  'True
         MouseIcon       =   "TMPU105.frx":0902
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TMPU105.frx":091E
         Caption         =   "&Salvar"
         ButtonStyle     =   3
         PictureAlignment=   6
      End
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   2295
      Left            =   30
      TabIndex        =   4
      Top             =   2310
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4048
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   1138
      Icone           =   "TMPU105.frx":093A
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "ENTER"
      Default         =   -1  'True
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   435
      Left            =   4410
      TabIndex        =   6
      Top             =   4650
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TMPU105.frx":0C54
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU105.frx":0C70
      Caption         =   "&Novo"
      ButtonStyle     =   3
      PictureAlignment=   6
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "Deseja sair?"
      Top             =   4650
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TMPU105.frx":0C8C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU105.frx":0CA8
      Caption         =   "Sai&r"
      ButtonStyle     =   3
      PictureAlignment=   6
   End
End
Attribute VB_Name = "TMPU105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdDeletar_Click()
    If txtPar.Enabled = False And Trim(txtPar) <> "" Then
        If Confirma("Deseja mesmo apagar " & txtPar & "?") Then
            If Bdados.DeletaDados("TAB_GRUPO_COMPONENTE_AVANCADO", "tgc_cod_grupo=" & txtPar) Then
                Informa "Grupo apagado."
            Else
                Avisa "Grupo não apagado."
            End If
        End If
    Else
        Avisa "Selecione um registro gravado."
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
    Call MontaGrid(Bdados, grid, _
    "SELECT tgc_cod_grupo AS Código, tgc_nome AS Descrição FROM TAB_GRUPO_COMPONENTE_AVANCADO", 1400)
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid grid, ColumnHeader
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub grid_DblClick()
    On Error GoTo trata
    txtDes.SetFocus
    
    txtPar = grid.SelectedItem.Text
    txtDes = grid.SelectedItem.SubItems(1)
    
    txtPar.Enabled = False
    
    Exit Sub
trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtPar_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Sub Atualizar(Codigo As String)
    On Error GoTo trata
    Dim Valores As String
    Dim Campos As String
    
    Campos = "tgc_cod_grupo, tgc_nome"
    Valores = Bdados.PreparaValor(txtPar, txtDes)
    
    Call Bdados.GravaDados("TAB_GRUPO_COMPONENTE_AVANCADO", Valores, Campos, _
    "tgc_cod_grupo = " & txtPar)

    Exit Sub
trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo trata
    
    If CriticaCampos(Me) Then
        If Confirma("Deseja salvar os dados de " & txtPar & "?") Then
            Screen.MousePointer = 11
            Call Atualizar(txtPar)
            Call cmdNovo_Click
            Informa "Operação realizada."
        End If
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo trata
    
    cabVisual.Exibir Bdados, Cod_Form, App.Path
    
    AtualizaG
    Screen.MousePointer = 0
    
    Exit Sub
trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtPar_LostFocus()
    Dim i As Integer
    txtPar = Trim(txtPar)
    For i = 1 To grid.ListItems.Count
        If grid.ListItems(i).Text = Trim(txtPar) Then
            txtDes.Text = grid.ListItems(i).ListSubItems(1).Text
            Exit For
        End If
    Next
    
End Sub
