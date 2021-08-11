VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.UserControl PrbBarraProgresso 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   KeyPreview      =   -1  'True
   ScaleHeight     =   315
   ScaleWidth      =   3900
   Begin Threed.SSPanel SPainel 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   556
      _Version        =   196610
      ForeColor       =   16777215
      Windowless      =   -1  'True
      Caption         =   "SSPanel1"
      FloodType       =   1
      FloodColor      =   12582912
      RoundedCorners  =   0   'False
      Begin VB.Label lblPercentual 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   1740
         TabIndex        =   1
         Top             =   60
         Width           =   315
      End
   End
End
Attribute VB_Name = "PrbBarraProgresso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Values As Integer
Public Property Get Value() As Integer
    If bRegistrado Then AutoTAB = m_Values
End Property
Public Property Let Value(Value As Integer)
    If bRegistrado Then
        m_Values = Value
        PropertyChanged "Value"
    End If
End Property

Private Sub UserControl_Resize()
    SPainel.Width = Width
    SPainel.Height = Height
    lblPercentual.Left = (SPainel.Width / 2) - 200
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", 0)
End Sub
