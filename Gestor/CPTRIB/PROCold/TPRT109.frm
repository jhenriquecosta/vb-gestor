VERSION 5.00
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TPRT109 
   Caption         =   "ATUALIZAÇÃO DE PROCESSO"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPrazo 
      Caption         =   "Check1"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VTOcx.cboVISUAL cboTIPO 
      Height          =   510
      Left            =   360
      TabIndex        =   0
      Tag             =   "C"
      Top             =   120
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   900
      Caption         =   "Etapa  "
      Text            =   ""
      AutoFocaliza    =   0   'False
      Requerido       =   0   'False
      Alinhamento     =   1
   End
   Begin VTOcx.cmdVISUAL cmdImprimir 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Atualizar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdCancelar 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1800
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Sair"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.txtVISUAL txtInicio 
      Height          =   480
      Left            =   360
      TabIndex        =   2
      Tag             =   "A"
      Top             =   1200
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   847
      Caption         =   "Data Ciência (Início)"
      Text            =   ""
      TipoLetras      =   0
      Formato         =   0
      AlinhamentoRotulo=   1
   End
   Begin VTOcx.txtVISUAL txtDias 
      Height          =   480
      Left            =   2400
      TabIndex        =   3
      Tag             =   "A"
      Top             =   1200
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   847
      Caption         =   "N. Dias (Prazo)"
      Text            =   ""
      TipoLetras      =   0
      AlinhamentoRotulo=   1
   End
   Begin VB.Label lbl 
      Caption         =   "Prazo para o processo!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "TPRT109"
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
     If Len(cboTIPO.Text) = 0 Then
        Mensagem "Informe a proxima etapa do processo"
        Exit Sub
     End If
     If os.atualizaProcesso(Processo, CInt(cboTIPO.Coluna(1).Valor), CDate(txtInicio), CInt(txtDias), chkPrazo.Value) Then
        'r.visualizar
     End If
     Unload Me
End Sub
Public Sub carregar(codOs As Long)
    os.PreencheCombo cboTIPO
    Processo = codOs
    txtInicio = Format(Now, "DD/MM/YYYY")
    txtDias = 0
    Me.Show
End Sub

