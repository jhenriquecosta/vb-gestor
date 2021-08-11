VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TOBR407 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TOBR405"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   885
      Left            =   30
      TabIndex        =   4
      Top             =   690
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1561
      Altura          =   1905
      Caption         =   " Pagamento"
      CorTexto        =   0
      CorFaixa        =   12632256
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   510
         Left            =   3300
         TabIndex        =   7
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   900
         Caption         =   "Valor Pago(R$)"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
         CorFundo        =   14737632
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtData 
         Height          =   510
         Left            =   1680
         TabIndex        =   6
         Tag             =   "Periodo Final"
         Top             =   300
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   900
         Caption         =   "Data Pagamento"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
         CorFundo        =   14737632
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtNumeroProcesso 
         Height          =   510
         Left            =   60
         TabIndex        =   5
         Tag             =   "Periodo Final"
         Top             =   300
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   900
         Caption         =   "No. Lote"
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
         CorFundo        =   14737632
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   3765
         TabIndex        =   2
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   1138
      Icone           =   "TOBR407.frx":0000
   End
   Begin VTOcx.grdVISUAL GrdTaxas 
      Height          =   1620
      Left            =   60
      TabIndex        =   3
      Top             =   6345
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   2858
      Caption         =   "Taxas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      CheckBox        =   -1  'True
   End
End
Attribute VB_Name = "TOBR407"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim String_Taxas As String
Dim Total_Taxas  As String
Dim NovaData As String

Private Sub cmd_Click(Index As Integer)
    Dim Cobranca As New VSCobranca
 
        Select Case Index
            Case 1
                Unload Me
            Case 2
                Avisa "Indisponível no momento."
        End Select
End Sub

Private Sub Form_Activate()
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "SELECT TDR_TLP_COD_LOTE,TDR_VALOR_REAL_PAGO,TDR_DATA_PAGAMENTO FROM TAB_DARM_RECEBIDO  " & _
        " WHERE TDR_TGT_COD_PAGAMENTO =" & Me.Tag
    If Bdados.AbreTabela(Sql, rs) Then
        txtData = "" & rs!TDR_DATA_PAGAMENTO
        txtValor = "" & rs!TDR_VALOR_REAL_PAGO
        txtNumeroProcesso = "" & rs!TDR_TLP_COD_LOTE
    End If
    
End Sub

