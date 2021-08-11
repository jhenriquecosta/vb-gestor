VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "m"
   ClientHeight    =   5940
   ClientLeft      =   990
   ClientTop       =   2940
   ClientWidth     =   9840
   Height          =   6345
   Left            =   930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9840
   Top             =   2595
   Width           =   9960
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   5280
   End
   Begin VB.Label Label4 
      Caption         =   "Label2"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   6255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "mmmmm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Frase As String, Frase3 As String
Dim Pos3 As Integer

Private Sub Form_Load()
    Frase = " Senhor Contribuinte: Esta informação é muito importante e o senhor deve ler atentamente antes de fazer qualquer besteira"
    Pos3 = 0
    Frase3 = Frase + Space(30)
    label3 = Frase3
End Sub

Private Sub Timer1_Timer()
    label3 = Mid(Frase3, 1, Pos3)
    Pos3 = Pos3 + 1
    If Pos3 = Len(Frase3) Then Pos3 = 1
End Sub
