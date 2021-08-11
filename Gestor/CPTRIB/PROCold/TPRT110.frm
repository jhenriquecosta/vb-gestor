VERSION 5.00
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TPRT110 
   Caption         =   "CAPITULAÇÃO LEGAL"
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cmdVISUAL cmdImprimir 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   1560
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
      Top             =   1560
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
      Width           =   7560
      _ExtentX        =   13335
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
      Width           =   7560
      _ExtentX        =   13335
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
      Width           =   7545
      _ExtentX        =   13309
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
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   529
      Caption         =   "Multa Aplicada     "
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
    txtJ = " 1% a.m.  § 3º Inciso I do Art. 372 da L.C. 001/2011"
    txtM = "0,33% dia até 30% - Alínea " & "a" & "do Inciso III do Art. 373 da L.C. 001/2011"
    txtCL = "SELIC - Inciso I do art. 372 da L.C. 001/2011"
    Me.Show
End Sub

