VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form GMEN102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GMEN102"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3225
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   3225
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.TreeView Tremenseger 
      Height          =   3690
      Left            =   105
      TabIndex        =   2
      Top             =   480
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   6509
      _Version        =   393217
      Indentation     =   538
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   105
      Picture         =   "GMEN102.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   435
   End
   Begin VB.Image ImgEnviarArquivoAzul 
      Height          =   300
      Left            =   210
      Picture         =   "GMEN102.frx":0490
      Stretch         =   -1  'True
      Top             =   5115
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.Image ImgEnviarArquivoNormal 
      Height          =   300
      Left            =   240
      Picture         =   "GMEN102.frx":0C5C
      Stretch         =   -1  'True
      Top             =   5115
      Width           =   2760
   End
   Begin VB.Image ImgEuGostariaAzul 
      Height          =   360
      Left            =   300
      Picture         =   "GMEN102.frx":1346
      Top             =   4275
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image ImgEuGostariaNormal 
      Height          =   405
      Left            =   300
      Picture         =   "GMEN102.frx":193B
      Top             =   4290
      Width           =   3345
   End
   Begin VB.Image ImgEnviarAzual 
      Height          =   300
      Left            =   225
      Picture         =   "GMEN102.frx":1FB7
      Stretch         =   -1  'True
      Top             =   4710
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.Image ImgEnviarNormal 
      Height          =   300
      Left            =   210
      Picture         =   "GMEN102.frx":272F
      Stretch         =   -1  'True
      Top             =   4710
      Width           =   2760
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   585
      Picture         =   "GMEN102.frx":2DD8
      Stretch         =   -1  'True
      Top             =   60
      Width           =   2505
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   165
      Picture         =   "GMEN102.frx":32DA
      Top             =   6030
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MESSENGER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   1
      Top             =   6375
      Width           =   2745
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GESTOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   6075
      Width           =   2745
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   750
      Left            =   105
      Top             =   5985
      Width           =   3000
   End
   Begin VB.Image Image2 
      Height          =   7695
      Left            =   -15
      Picture         =   "GMEN102.frx":376A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3315
   End
   Begin VB.Menu MnuArquivos 
      Caption         =   "Arquivos"
      Begin VB.Menu IteConfiguracoes 
         Caption         =   "Configugações"
      End
      Begin VB.Menu MnuSair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "GMEN102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim Sql            As String
    Dim RS             As VSRecordset
    Dim ContadorOnline As Integer
    Dim ContadorOfline As Integer
    
    
   'PEGO OS USUÁRIO QUE ESTÃO ONLINE...
    Sql = "SELECT * FROM TAB_USUARIO ORDER BY TUS_ONLINE_OFLINE "
    If BDados.AbreTabela(Sql, RS) Then
        Tremenseger.Nodes.Add , , "OnLine", "OnLine"
        Tremenseger.Nodes.Add , , "OfLine", "OfLine"
        RS.MoveFirst
        Do Until RS.EOF
            If RS.Fields("TUS_ONLINE_OFLINE") = 1 Then
                Tremenseger.Nodes.Add "OnLine", 4, RS.Fields("tus_cod_usuario"), RS.Fields("tus_cod_usuario")
                ContadorOnline = ContadorOnline + 1
            Else
                Tremenseger.Nodes.Add "OfLine", 4, RS.Fields("tus_cod_usuario"), RS.Fields("tus_cod_usuario")
                ContadorOfline = ContadorOfline + 1
            End If
            RS.MoveNext
        Loop
    End If
End Sub



Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgEuGostariaAzul.Visible = False
    ImgEnviarAzual.Visible = False
    ImgEnviarArquivoAzul.Visible = False
End Sub

Private Sub ImgEnviarArquivoNormal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ImgEnviarArquivoAzul.Visible = False Then
        ImgEnviarArquivoAzul.Visible = True
    End If
End Sub

Private Sub ImgEnviarNormal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ImgEnviarAzual.Visible = False Then
        ImgEnviarAzual.Visible = True
    End If
End Sub

Private Sub ImgEuGostariaNormal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ImgEuGostariaAzul.Visible = False Then
        ImgEuGostariaAzul.Visible = True
    End If
End Sub

Private Sub IteConfiguracoes_Click()
    GMEN104.Show 1
End Sub

Private Sub Tremenseger_DblClick()
    If UCase(Tremenseger.SelectedItem.Key) = "ONLINE" Then Exit Sub
    If UCase(Tremenseger.SelectedItem.Key) = "OFLINE" Then Exit Sub
    
    If Tremenseger.SelectedItem.Key <> "Oline" Or Tremenseger.SelectedItem.Key <> "Ofline" Then
        GMEN103.Show
    End If
End Sub
