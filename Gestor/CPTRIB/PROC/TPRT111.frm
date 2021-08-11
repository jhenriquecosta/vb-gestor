VERSION 5.00
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TPRT111 
   Caption         =   "TERMO DE PARCELAMENTO"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3360
      Width           =   9015
   End
   Begin VB.TextBox txt4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2760
      Width           =   9015
   End
   Begin VB.TextBox txt3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   9015
   End
   Begin VB.TextBox txt2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   9015
   End
   Begin VB.TextBox txt1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   9015
   End
   Begin VTOcx.cmdVISUAL cmdImprimir 
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   4080
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
      Left            =   6720
      TabIndex        =   8
      Top             =   4080
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Sair"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.txtVISUAL txtParcelas 
      Height          =   480
      Left            =   8400
      TabIndex        =   1
      Tag             =   "A"
      Top             =   180
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   847
      Caption         =   "Parcelas"
      Text            =   ""
      TipoLetras      =   0
      Restricao       =   2
      AlinhamentoRotulo=   1
   End
   Begin VTOcx.txtVISUAL txtref 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   529
      Caption         =   ""
      Text            =   ""
      Requerido       =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Referente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Documentos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "TPRT111"
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
    os.ImprimeAutoInfracao Processo, "", "", "", "", True, CInt(CStr(txtParcelas)), CStr(txtref), CStr(txt1), CStr(txt2), CStr(txt3), CStr(txt4), CStr(txt5)
     Unload Me
End Sub
Public Sub carregar(codOs As Long)
    Processo = codOs
    txt1 = "Cópia simples do CNPJ/MF, do Contrato Social e da última alteração no Contrato Social da Confitente Devedora"
    txt2 = "Cópia simples do CPF/MF, da Carteira de Identidade e do comprovante de endereço do sócios da Confitente Devedora"
    txt3 = "Cópia do Levantamento de ISSQN do débito objeto do presente termo"
    txtref = "ISSQN (Imposto sobre Serviço de Qualquer Natureza)"
    txtParcelas = 12
    Me.Show
End Sub

