VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "CABECALHO.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   2670
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   926
      CorFundo        =   12632256
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   1138
      Icone           =   "Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim bd As New VSClass.VSDados
    Show
    If bd.AbreBanco(SQLServer, "VTSERWEB", "SA", "VCLA", "VTFROT") Then
        cabVISUAL1.Exibir bd, "BMEC403", App.Path
        rodVISUAL1.Exibir bd, "BMEC403"
    End If
End Sub
