VERSION 5.00
Begin VB.Form Dec010 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   855
   ClientTop       =   2895
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2640
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00800000&
      Height          =   2565
      Left            =   30
      ScaleHeight     =   2505
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   30
      Width           =   7575
      Begin VB.PictureBox SSPanel2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1155
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   7215
         TabIndex        =   2
         Top             =   810
         Width           =   7275
      End
      Begin VB.PictureBox SSCommand1 
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   6120
         ScaleHeight     =   375
         ScaleWidth      =   1365
         TabIndex        =   4
         Top             =   2100
         Width           =   1425
      End
      Begin VB.PictureBox Check 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   210
         ScaleHeight     =   225
         ScaleWidth      =   4395
         TabIndex        =   3
         Top             =   2220
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atenção"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   2550
         TabIndex        =   1
         Top             =   60
         Width           =   2265
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   2625
      Left            =   30
      Top             =   0
      Width           =   7605
   End
End
Attribute VB_Name = "Dec010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
   CentrarFormu Me
End Sub

Private Sub SSCommand1_Click()
If Check.Value = True Then
    gArq = "\Dief.ini"
    nFile = FreeFile(0)
    Open App.Path + gArq For Output As nFile
    Print #nFile, "nMostra:False"
    Close #nFile
End If
DoEvents
Unload Me
End Sub


