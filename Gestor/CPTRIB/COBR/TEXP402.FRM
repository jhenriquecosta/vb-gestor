VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TEXP402 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TEXP402.frx":0000
   ScaleHeight     =   5040
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TEXP402.frx":0342
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   795
      Left            =   3105
      TabIndex        =   10
      Top             =   735
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1402
      Altura          =   1905
      Caption         =   " Geração"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtDtVence 
         Height          =   300
         Left            =   2715
         TabIndex        =   2
         Tag             =   "Data Vencimento"
         Top             =   360
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   529
         Caption         =   "Data Vencimento"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
      End
      Begin VTOcx.txtVISUAL txtContrib 
         Height          =   300
         Left            =   150
         TabIndex        =   1
         Top             =   375
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         Caption         =   "Contribuinte"
         Text            =   ""
         TipoLetras      =   0
         Formato         =   8
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   11
      Top             =   4500
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   375
         Left            =   6660
         TabIndex        =   5
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdRecupera 
         Height          =   375
         Left            =   5310
         TabIndex        =   4
         Top             =   105
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   661
         Caption         =   "&Recuperar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7845
         TabIndex        =   6
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Sair"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   795
      Left            =   90
      TabIndex        =   9
      Top             =   735
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   1402
      Altura          =   1905
      Caption         =   " Consulta"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtExtrato 
         Height          =   300
         Left            =   390
         TabIndex        =   0
         Tag             =   "Número do Extrato"
         Top             =   375
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   529
         Caption         =   "Nº Extrato"
         Text            =   ""
         Restricao       =   2
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   930
      Width           =   375
   End
   Begin VTOcx.grdVISUAL lstAtv 
      Height          =   2880
      Left            =   90
      TabIndex        =   3
      Top             =   1605
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   5080
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   1138
      Icone           =   "TEXP402.frx":2465
   End
End
Attribute VB_Name = "TEXP402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Busca()
    Dim Sql As String
    If Trim(txtExtrato) = "" Then Exit Sub
    Sql = "Select TPE_TGT_COD_PAGAMENTO as DOCUMENTO,TGT_IM AS IM,TIP_SIGLA_IMPOSTO AS TRIBUTO,TGT_PERIODO AS PERIODO, TPE_SUB_VALOR AS VALOR " & _
        " FROM TAB_PAGAMENTO_EXTRATO LEFT JOIN TAB_IMPOSTO ON  TPE_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO " & _
        "LEFT JOIN TAB_GERACAO_TRIBUTO ON TPE_TGT_COD_PAGAMENTO = TGT_COD_PAGAMENTO" & _
        " where TPE_COD_PAGAMENTO_EXTRATO =" & txtExtrato & " ORDER BY TGT_PERIODO DESC"
    If Not lstAtv.Preencher(Bdados, Sql, 1500, 0, 1000, 0, 1000) Then
        Avisa "Extrato não encontrado."
        txtExtrato.SetFocus
        Exit Sub
    End If
    txtContrib = lstAtv.SelectedItem.SubItems(1)
    lstAtv.Mensagem = "Total do extrato R$ " & Format(lstAtv.Colunas(5).Soma, Const_Monetario)
    txtDtVence.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdRecupera_Click()
    Dim Sql As String
    Dim CCorrente As New ContaCorrente
    Dim rs As VSRecordset
    Dim Conta As New ContaCorrente
    Dim CodPagamento As Double
    Dim i As Integer
    
    On Error Resume Next
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
      
    Conta.GeraPagamento txtContrib, "", Const_Extrato, Right(Format(Date, "DD/MM/YYYY"), 4) & _
                    Mid(Format(Date, "DD/MM/YYYY"), 4, 2), txtDtVence, CDbl(lstAtv.Colunas(5).Soma), 0, 0, _
                    txtExtrato, 0, 0, 0, , EtcCreditoTributario
    Screen.MousePointer = 0
    Informa "Extrato recuperado com sucesso."
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    Edita.LimpaCampos Me
    lstAtv.Preencher Bdados, ""
    txtExtrato.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    AtualizaCabecalho lstAtv
    If UCase(Temp.PegaParametro(Bdados, "TIPO INSCRICAO")) = "REDUZIDA" Then
        txtContrib.Formato = formNenhum
    Else
        txtContrib.Formato = formDocumento
    End If
End Sub

Private Sub txtExtrato_LostFocus()
    If Trim(txtExtrato) = "" Then Exit Sub
    Busca
End Sub
