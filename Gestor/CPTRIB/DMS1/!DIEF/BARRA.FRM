VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   2985
   ClientTop       =   1515
   ClientWidth     =   6690
   Height          =   6345
   Left            =   2925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   1170
   Width           =   6810
   Begin VB.PictureBox Picture2 
      Height          =   435
      Left            =   1770
      Picture         =   "BARRA.frx":0000
      ScaleHeight     =   375
      ScaleMode       =   0  'User
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   240
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2775
      ScaleHeight     =   675
      ScaleWidth      =   1845
      TabIndex        =   11
      Top             =   2475
      Width           =   1905
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   960
         Picture         =   "BARRA.frx":0442
         Top             =   45
         Width           =   510
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   60
         Picture         =   "BARRA.frx":0884
         Stretch         =   -1  'True
         Top             =   150
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   1800
      Left            =   435
      ScaleHeight     =   1740
      ScaleWidth      =   5970
      TabIndex        =   0
      Top             =   3270
      Width           =   6030
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4425
         Picture         =   "BARRA.frx":0B8E
         ScaleHeight     =   180
         ScaleMode       =   0  'User
         ScaleWidth      =   67.5
         TabIndex        =   20
         Top             =   1050
         Width           =   180
      End
      Begin VB.PictureBox pic_sair 
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   4485
         Picture         =   "BARRA.frx":0FD0
         ScaleHeight     =   8.467
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   7.673
         TabIndex        =   6
         Top             =   165
         Width           =   495
      End
      Begin VB.PictureBox pic_ajuda 
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   3510
         Picture         =   "BARRA.frx":12DA
         ScaleHeight     =   8.467
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   7.673
         TabIndex        =   5
         Top             =   165
         Width           =   495
      End
      Begin VB.PictureBox pic_diretorio 
         AutoSize        =   -1  'True
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   2520
         Picture         =   "BARRA.frx":171C
         ScaleHeight     =   8.467
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   8.467
         TabIndex        =   4
         Top             =   165
         Width           =   540
      End
      Begin VB.PictureBox pic_ferramentas 
         FillStyle       =   0  'Solid
         Height          =   570
         Left            =   1560
         Picture         =   "BARRA.frx":1B5E
         ScaleHeight     =   8.996
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   8.202
         TabIndex        =   3
         Top             =   165
         Width           =   525
      End
      Begin VB.PictureBox pic_declaracao 
         FillStyle       =   0  'Solid
         Height          =   570
         Left            =   840
         Picture         =   "BARRA.frx":1FA0
         ScaleHeight     =   8.996
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   8.731
         TabIndex        =   2
         Top             =   165
         Width           =   555
      End
      Begin VB.PictureBox pic_cadastro 
         FillStyle       =   0  'Solid
         Height          =   570
         Left            =   0
         Picture         =   "BARRA.frx":23E2
         ScaleHeight     =   8.996
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   10.054
         TabIndex        =   1
         Top             =   165
         Width           =   630
      End
      Begin VB.PictureBox cmd_sair 
         Height          =   600
         Left            =   2385
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   14
         Top             =   885
         Width           =   600
      End
      Begin VB.PictureBox cmd_cad 
         AutoSize        =   -1  'True
         Height          =   600
         Left            =   390
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   13
         Top             =   915
         Width           =   600
      End
      Begin VB.PictureBox cmd_ferr 
         AutoSize        =   -1  'True
         Height          =   600
         Left            =   1365
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   12
         Top             =   945
         Width           =   600
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   285
         Left            =   4395
         TabIndex        =   19
         Top             =   990
         Width           =   285
         _version        =   65536
         _extentx        =   503
         _extenty        =   503
         _stockprops     =   78
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   600
         Left            =   3510
         TabIndex        =   15
         Top             =   960
         Width           =   600
         _version        =   65536
         _extentx        =   1058
         _extenty        =   1058
         _stockprops     =   78
         autosize        =   2
         picture         =   "BARRA.frx":2824
      End
   End
   Begin VB.Image Image11 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2865
      Picture         =   "BARRA.frx":2C76
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2340
      Picture         =   "BARRA.frx":30B8
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1830
      Picture         =   "BARRA.frx":31BA
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1260
      Picture         =   "BARRA.frx":35FC
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   735
      Picture         =   "BARRA.frx":3A3E
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   240
      Picture         =   "BARRA.frx":3E80
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   360
   End
   Begin Threed.SSRibbon SSRibbon2 
      Height          =   510
      Left            =   180
      TabIndex        =   18
      Top             =   1275
      Width           =   6390
      _version        =   65536
      _extentx        =   11271
      _extenty        =   900
      _stockprops     =   65
      backcolor       =   12632256
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   5070
      Picture         =   "BARRA.frx":42C2
      Top             =   255
      Width           =   300
   End
   Begin VB.Image Image6 
      Height          =   390
      Left            =   4245
      Top             =   285
      Width           =   405
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3465
      Picture         =   "BARRA.frx":43C4
      Top             =   270
      Width           =   300
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   990
      Picture         =   "BARRA.frx":44C6
      Stretch         =   -1  'True
      Top             =   180
      Width           =   435
   End
   Begin Threed.SSRibbon SSRibbon1 
      Height          =   675
      Left            =   735
      TabIndex        =   16
      Top             =   90
      Width           =   5145
      _version        =   65536
      _extentx        =   9075
      _extenty        =   1191
      _stockprops     =   65
      backcolor       =   12632256
      groupallowallup =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sair"
      Height          =   225
      Index           =   3
      Left            =   4935
      TabIndex        =   10
      Top             =   5100
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Diretório"
      Height          =   225
      Index           =   2
      Left            =   2835
      TabIndex        =   9
      Top             =   5145
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Declaração"
      Height          =   225
      Index           =   1
      Left            =   1260
      TabIndex        =   8
      Top             =   5130
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cadastro"
      Height          =   225
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   5100
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Option Explicit

Private Sub imgsalvar_Click()
 MsgBox " vou sair"
 Unload Me
 
End Sub


Private Sub imgsalvardesabilitado_Click()
MsgBox "chamou manut"
End Sub

Private Sub imgsalvarnormal_Click()
MsgBox "chamou declaracao"
Label1.Visible = False
End Sub

Private Sub imgsalvarnormal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Visible = False
End Sub


Private Sub imgsalvarnormal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label1.Visible = True

If X > 345 Or X < 0 Then
   Label1.Visible = False
End If
If Y > 300 Or Y < 0 Then
   Label1.Visible = False
End If

End Sub


Private Sub imgsalvarnormal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Visible = False

End Sub


Private Sub imgsalvarpressionado_Click()
MsgBox "chamou diretorioo"
End Sub

Private Sub picbarra_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Visible = False
End Sub


Private Sub cmd_cad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(0).Visible = True
Label1(1).Visible = False
Label1(2).Visible = False
Label1(3).Visible = False

End Sub


Private Sub cmd_ferr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(0).Visible = False
Label1(1).Visible = True
Label1(2).Visible = False
Label1(3).Visible = False

End Sub


Private Sub cmd_sair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(0).Visible = False
Label1(1).Visible = False
Label1(2).Visible = False
Label1(3).Visible = True

End Sub


Private Sub Picture1_Click()
MsgBox "chamou o cadastro"
End Sub


Private Sub Picture1_LostFocus()
  Label1.Visible = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Visible = True

'MsgBox "valor de x - " & X & "valor de Y - " & Y
'If X >= 7.1 Then 'And Y >= 7.1 Then
'   Label1.Visible = False
'   Exit Sub
'End If

'If X >= 7.2 And Y <= 0 Then
'   Label1.Visible = False
'End If
'If X <= 0.3 Then '1 And Y <= 0 Then
'   Label1.Visible = False
'End If

'If X <= 0 And Y >= 8.2 Then
'   Label1.Visible = False
'End If


End Sub


Private Sub SSCommand1_Click()
MsgBox "cadastro 1 "
End Sub


Private Sub pic_cadastro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label1(0).Visible = True
Label1(1).Visible = False
Label1(2).Visible = False
Label1(3).Visible = False


End Sub


Private Sub pic_declaracao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(0).Visible = False
Label1(1).Visible = True
Label1(2).Visible = False
Label1(3).Visible = False

End Sub


Private Sub pic_diretorio_Click()
Label1(0).Visible = False
Label1(1).Visible = False
Label1(2).Visible = True
Label1(3).Visible = False

End Sub


Private Sub pic_sair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(0).Visible = False
Label1(1).Visible = False
Label1(2).Visible = False
Label1(3).Visible = True

End Sub


Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(0).Visible = False
Label1(1).Visible = False
Label1(2).Visible = False
Label1(3).Visible = False

End Sub


Private Sub SSCommand1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(0).Visible = True
Label1(1).Visible = False
Label1(2).Visible = False
Label1(3).Visible = False
End Sub


Private Sub SSCommand2_Click()

End Sub


Private Sub SSCommand2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(0).Visible = True
Label1(1).Visible = False
Label1(2).Visible = False
Label1(3).Visible = False

End Sub


