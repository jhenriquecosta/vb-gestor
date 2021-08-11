VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EMAN501 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ControlBox      =   0   'False
   Icon            =   "TMPU501.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
   Begin Threed.SSPanel pan 
      Height          =   465
      Left            =   60
      TabIndex        =   8
      Top             =   30
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   820
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
      BorderWidth     =   1
      Alignment       =   6
      RoundedCorners  =   0   'False
      Begin Threed.SSPanel lblForm 
         Height          =   375
         Left            =   1290
         TabIndex        =   9
         Top             =   60
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   661
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   128
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Grupo de Componentes"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lblModulo 
         Height          =   285
         Left            =   6150
         TabIndex        =   10
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   503
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   128
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
         Caption         =   "EMAN501"
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         AutoSize        =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lblHora 
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   128
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
         Caption         =   "00:00:00"
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         AutoSize        =   3
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "ENTER"
      Default         =   -1  'True
      Height          =   315
      Left            =   180
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin Threed.SSFrame fra 
      Height          =   1545
      Index           =   1
      Left            =   60
      TabIndex        =   13
      Top             =   570
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         MouseIcon       =   "TMPU501.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TMPU501.frx":08E6
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
         MouseIcon       =   "TMPU501.frx":0902
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TMPU501.frx":091E
         Caption         =   "&Salvar"
         ButtonStyle     =   3
         PictureAlignment=   6
      End
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   2415
      Left            =   60
      TabIndex        =   4
      Top             =   2160
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4260
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
      MouseIcon       =   "TMPU501.frx":093A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU501.frx":0956
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
      MouseIcon       =   "TMPU501.frx":0972
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU501.frx":098E
      Caption         =   "Sai&r"
      ButtonStyle     =   3
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdAjuda 
      Height          =   435
      Left            =   60
      TabIndex        =   7
      ToolTipText     =   "Ajuda"
      Top             =   4650
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      MousePointer    =   14
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TMPU501.frx":09AA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU501.frx":09C6
      Caption         =   "?"
      ButtonStyle     =   3
      PictureAlignment=   6
   End
End
Attribute VB_Name = "EMAN501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAjuda_Click()
    Temp.AjudaTemporaria
End Sub

Private Sub cmdDeletar_Click()
    If txtPar.Enabled = False And Trim(txtPar) <> "" Then
        If Confirma("Deseja mesmo apagar " & txtPar & "?") Then
            If Bdados.DeletaDados(BancoSistema, "TAB_GRUPO_COMPONENTE", "TGC_COD_GRUPO=" & txtPar) Then
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
    Call MontaGrid(Grid, BancoSistema, _
    "SELECT TGC_COD_GRUPO AS Código, TGC_NOME AS Descrição FROM TAB_GRUPO_COMPONENTE")
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Grid.SortKey = ColumnHeader.Index - 1 Then
        Grid.SortOrder = Abs(Grid.SortOrder - 1)
    Else
        Grid.SortOrder = lvwAscending
        Grid.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
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
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtPar_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
    KeyAscii = Maiuscula(KeyAscii)
End Sub

Sub Atualizar(Codigo As String)
    On Error GoTo Trata
    Dim Valores As String
    Dim Campos As String
    
    Campos = "TGC_COD_GRUPO, TGC_NOME,TGC_MODELO,TGC_TIPO,TGC_REQUERIDO"
    Valores = Bdados.PreparaValor(txtPar, txtDes, 0, 0, 0)
    
    Call Bdados.GravaDados(BancoSistema, "TAB_GRUPO_COMPONENTE", Campos, Valores, _
    "TGC_COD_GRUPO = " & txtPar)

    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo Trata
    
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
Trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo Trata
    
    Instala.Perfil Me, lblForm, lblModulo, Aplica.Usuario
    
    AtualizaG
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub tmrCaixa_Timer()
    lblHora = Time
End Sub

Private Sub txtPar_LostFocus()
    Dim I As Integer
    txtPar = Trim(txtPar)
    For I = 1 To Grid.ListItems.Count
        If Grid.ListItems(I).Text = Trim(txtPar) Then
            txtDes.Text = Grid.ListItems(I).ListSubItems(1).Text
            Exit For
        End If
    Next
    
End Sub
