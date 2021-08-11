VERSION 5.00
Begin VB.Form Dec007 
   BorderStyle     =   0  'None
   Caption         =   "Validar Declaração"
   ClientHeight    =   5220
   ClientLeft      =   1125
   ClientTop       =   1380
   ClientWidth     =   7215
   ControlBox      =   0   'False
   HelpContextID   =   5
   Icon            =   "DEC007.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5220
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox SSPanel1 
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   90
      ScaleHeight     =   225
      ScaleWidth      =   6975
      TabIndex        =   6
      Top             =   960
      Width           =   7035
   End
   Begin VB.PictureBox SSPanel1 
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   90
      ScaleHeight     =   225
      ScaleWidth      =   6975
      TabIndex        =   7
      Top             =   1860
      Width           =   7035
   End
   Begin VB.PictureBox SSPanel1 
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   3
      Left            =   90
      ScaleHeight     =   225
      ScaleWidth      =   6975
      TabIndex        =   8
      Top             =   2910
      Width           =   7035
   End
   Begin VB.PictureBox SSPanel1 
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   90
      ScaleHeight     =   225
      ScaleWidth      =   6975
      TabIndex        =   9
      Top             =   3840
      Width           =   7035
   End
   Begin VB.PictureBox SSPanel1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   90
      ScaleHeight     =   225
      ScaleWidth      =   6585
      TabIndex        =   5
      Top             =   60
      Width           =   6645
   End
   Begin VB.PictureBox SSCommand1 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6780
      ScaleHeight     =   255
      ScaleWidth      =   285
      TabIndex        =   11
      Top             =   60
      Width           =   345
   End
   Begin VB.PictureBox Cmdfinal 
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   5370
      ScaleHeight     =   375
      ScaleWidth      =   1665
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4650
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   5145
      Left            =   30
      Top             =   30
      Width           =   7155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Index           =   4
      Left            =   690
      TabIndex        =   4
      Top             =   4200
      Width           =   6435
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Index           =   3
      Left            =   675
      TabIndex        =   3
      Top             =   3180
      Width           =   6450
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   2
      Left            =   690
      TabIndex        =   2
      Top             =   2100
      Width           =   6435
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1230
      Width           =   6405
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   585
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   6405
   End
   Begin VB.Image ImageErro 
      Height          =   285
      Left            =   3840
      Picture         =   "DEC007.frx":0442
      Top             =   5250
      Width           =   375
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   5070
      Picture         =   "DEC007.frx":0A98
      Top             =   5220
      Width           =   480
   End
   Begin VB.Image ImageOK 
      Height          =   285
      Left            =   4530
      Picture         =   "DEC007.frx":0EDA
      Top             =   5250
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   3
      Left            =   180
      Picture         =   "DEC007.frx":1530
      Top             =   3270
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   4
      Left            =   180
      Picture         =   "DEC007.frx":1B86
      Top             =   4260
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   2
      Left            =   180
      Picture         =   "DEC007.frx":21DC
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   1
      Left            =   180
      Picture         =   "DEC007.frx":2832
      Top             =   1350
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   0
      Left            =   180
      Picture         =   "DEC007.frx":2E88
      Stretch         =   -1  'True
      Top             =   450
      Width           =   375
   End
End
Attribute VB_Name = "Dec007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Cmdfinal_Click()
Dim ret As Integer
    
    query = "Select * from tab_declaracao where tde_num_inscricao = '" & Trim(Str(dec002.msk_inscricao)) & "'"
    query = query & " and tde_periodo_inicial =  '" & g_periodo_inicial & "'"
    query = query & " and tde_substitutiva =  " & Val(g_tipo_decla) & " And tde_estado = 'I'"
    
    Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
    If restauracontr.EOF And restauracontr.BOF Then
        MsgBox "Esta Declaração já esta Finalizada", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
        Unload Me
        Exit Sub
    End If

    ret = MsgBox("Você Deseja Finalizar a Declaração ?", 36, NOME_SISTEMA_DEFAULT)

    If ret = 6 Then
        controle = calcula_controle(g_num_inscricao, g_periodo_inicial, g_tipo_decla) * 100
        query = "update tab_declaracao set "
        query = query & "tde_estado = 'F'"
        query = query & ", tde_controle = '" & controle & "'"
        query = query & " where tde_num_inscricao = '" & Trim(Str(dec002.msk_inscricao)) & "'"
        query = query & " and tde_periodo_inicial = '" & g_periodo_inicial & "'"
        query = query & " and tde_substitutiva = " & Val(g_tipo_decla)
        query = query & " and tde_aliquota_m17 = " & g_aliquota_Maior17
        query = query & " and tde_condicao = " & g_condicao
        
        If Not EjecutarQuery(query) Then
           MsgBox "Não foi possivel finalizar Declaração, favor verificar e tentar novamente", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           Unload Me
        Else
           MsgBox "Finalização Completada", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           dec002.Situacao.Caption = "Finalizada"
           Unload Me
        End If
    End If

End Sub

Private Sub Form_Activate()
Dim a
For a = 0 To 4
  If Trim(Label3(a).Caption) <> "OK" Then
        Cmdfinal.Visible = False
  End If
Next

'query = "Select * from tab_declaracao where tde_num_inscricao = '" & Trim(Str(dec002.msk_inscricao)) & "'"
'query = query & " and tde_periodo_inicial =  '" & g_periodo_inicial & "'"
'query = query & " and tde_substitutiva =  " & Val(g_tipo_decla) & " And tde_estado = 'I'"
'Set restauracontr = dbgim.OpenRecordset(query, dbOpenDynaset)
'If restauracontr.EOF And restauracontr.BOF Then
'    Cmdfinal.Visible = False
'    MsgBox g_num_cgc_cpf, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii <> 0 Then
  Unload Me
End If
End Sub

Private Sub Form_Load()

CentrarFormu Me
End Sub


Private Sub Label3_Change(Index As Integer)
Dim erro As Variant

If Trim(Label3(Index).Caption) <> "OK" Then
   Label3(Index).ForeColor = &HC0&
   Label3(Index).FontSize = 8
   
Else
   Label3(Index).ForeColor = &H800000
   Label3(Index).FontBold = True
   Label3(Index).FontSize = 20
  
End If
   
End Sub

Private Sub SSCommand1_Click()
 Unload Me
End Sub


