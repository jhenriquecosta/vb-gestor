VERSION 5.00
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TPRT110 
   Caption         =   "CAPITULAÇÃO LEGAL"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cmdVISUAL cmdImprimir 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   1920
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Imprimir"
      Acao            =   4
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdCancelar 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Sair"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.txtVISUAL txtI 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Tag             =   "C"
      Top             =   120
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   529
      Caption         =   "Infração             "
      Text            =   ""
      Requerido       =   0   'False
   End
   Begin VTOcx.txtVISUAL txtCL 
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Tag             =   "C"
      Top             =   480
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   529
      Caption         =   "Capitulação Legal"
      Text            =   ""
      Requerido       =   0   'False
   End
   Begin VTOcx.txtVISUAL txtJ 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Tag             =   "C"
      Top             =   840
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   529
      Caption         =   "Juros Aplicado     "
      Text            =   ""
      Requerido       =   0   'False
   End
   Begin VTOcx.txtVISUAL txtM 
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Tag             =   "C"
      Top             =   1200
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   529
      Caption         =   "Multa Aplicada     "
      Text            =   ""
      Requerido       =   0   'False
   End
   Begin VTOcx.txtVISUAL txtAm 
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Tag             =   "C"
      Top             =   1560
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   529
      Caption         =   "At. Monetária      "
      Text            =   ""
      Requerido       =   0   'False
   End
End
Attribute VB_Name = "TPRT110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private r As VSRelatorio
Private Processo As Long
Dim os As New ordemServico

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
      os.ImprimeAutoInfracao Processo, CStr(txtI), CStr(txtCL), CStr(txtJ), CStr(txtM), False, CInt(0), CStr(""), CStr(""), CStr(""), CStr(""), CStr(""), CStr("")
        'r.visualizar
    ''End If
     Unload Me
End Sub
Public Sub carregar(codOs As Long)
    Processo = codOs
    'Set r = rpt
    txtJ = "Art. 10 §3º da L.C. 086/2008"
    txtM = "0,33% dia. Art. 250;251 e 252 da L.C. 086/2008 e alterações da Lei 153/2011"
    txtCL = ""
    txtAm = "Art. 10 §2º e Art. 13 p/único da L.C. 086/2008"
    Me.Show
End Sub

