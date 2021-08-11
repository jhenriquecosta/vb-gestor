VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TNOT403 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TNOT403.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.ComboBox cboTipo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      ItemData        =   "TNOT403.frx":2123
      Left            =   6285
      List            =   "TNOT403.frx":2130
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   660
      Width           =   1845
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   11
      Top             =   6825
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   900
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdImprime 
         Height          =   375
         Left            =   7335
         TabIndex        =   7
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9600
         TabIndex        =   9
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   8550
         TabIndex        =   8
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TNOT403.frx":2152
   End
   Begin VTOcx.txtVISUAL txtRefInicio 
      Height          =   300
      Left            =   4200
      TabIndex        =   4
      Tag             =   "Periodo Inicial"
      Top             =   1890
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   529
      Caption         =   "Data"
      Text            =   ""
      Formato         =   0
      Restricao       =   2
   End
   Begin VTOcx.txtVISUAL txtRefFim 
      Height          =   300
      Left            =   6015
      TabIndex        =   5
      Tag             =   "Periodo Final"
      Top             =   1890
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      Caption         =   "até"
      Text            =   ""
      Formato         =   0
      Restricao       =   2
   End
   Begin VTOcx.grdVISUAL grdNotifica 
      Height          =   4290
      Left            =   90
      TabIndex        =   12
      Top             =   2430
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   7567
      Caption         =   "Notificacoes emitidas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VTOcx.cmdVISUAL cmdParcela 
      Height          =   375
      Left            =   8970
      TabIndex        =   6
      Top             =   1980
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      Caption         =   "&Consultar"
      Acao            =   5
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin Threed.SSPanel lbl 
      Height          =   240
      Index           =   0
      Left            =   5400
      TabIndex        =   13
      Top             =   675
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   423
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   -2147483626
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tipo"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   5
      RoundedCorners  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtNotInicial 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Tag             =   "Periodo Inicial"
      Top             =   1860
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   529
      Caption         =   "Notificação nº"
      Text            =   ""
      MaxLen          =   8
   End
   Begin VTOcx.txtVISUAL txtNotFinal 
      Height          =   300
      Left            =   2445
      TabIndex        =   3
      Tag             =   "Periodo Final"
      Top             =   1860
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   529
      Caption         =   "até"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   8
   End
   Begin VTOcx.txtVISUAL txtRazao 
      Height          =   315
      Left            =   540
      TabIndex        =   14
      Top             =   1080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   556
      Caption         =   "Razão"
      Text            =   ""
      Enabled         =   0   'False
   End
   Begin VTOcx.txtVISUAL txtEndereco 
      Height          =   315
      Left            =   270
      TabIndex        =   15
      Top             =   1410
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   556
      Caption         =   "Endereço"
      Text            =   ""
      Enabled         =   0   'False
   End
   Begin VTOcx.txtVISUAL TXTIM 
      Height          =   300
      Left            =   300
      TabIndex        =   0
      Tag             =   "Inscrição Cadastral"
      Top             =   720
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   529
      Caption         =   "Inscricao"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   20
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
      Height          =   315
      Left            =   3060
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   720
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
   Begin VB.Menu mnuNotifica 
      Caption         =   "."
      Visible         =   0   'False
      Begin VB.Menu mnuEmitir 
         Caption         =   "&Emitir notificação ..."
      End
      Begin VB.Menu mnuCancelamento 
         Caption         =   "&Cancelar Notificação"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "TNOT403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim Sql As String

Private Sub cboDest_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub cboImposto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdCancela_Click()
    Dim rs As VSRecordset
    Dim Sql As String
    txtIm.Enabled = True
    cmdParcela.Enabled = True
    Edita.LimpaCampos Me
    
End Sub

Private Sub cmdEnter_Click()
'    SendKeys "{TAB}"
End Sub

Private Sub cmdImprime_Click()
    On Error GoTo Trata
    Dim i As Integer
    Dim ImAnterior As String
    Dim SelecaoRpt As String
    Dim Conta As New ContaCorrente
    Dim CodPagamento  As Double
    Dim Valores As String
    Dim Campos As String
    Dim Cobranca As New VSCobranca
    
    Screen.MousePointer = 11
    '1.
    ImprimirNotificacao txtIm, txtRefInicio, txtRefFim, cboTipo.ListIndex, Nvl(Trim(txtNotInicial), 0), Nvl(Trim(txtNotFinal), 0)
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Avisa Err.Description
        Exit Sub
        Err.Clear
    End If
End Sub
Private Sub cmdParcela_Click()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Notific As New Notificacao
    Notific.ExibirNotificacoes grdNotifica, txtIm, txtRefInicio, txtRefFim, cboTipo.ListIndex, Nvl(Trim(txtNotInicial), 0), Nvl(Trim(txtNotFinal), 0)
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub grdNotifica_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not grdNotifica.SelectedItem Is Nothing Then
        If Button = 2 Then
            mnuEmitir.Caption = "Emitir Notificação nº " & grdNotifica.SelectedItem
            mnuEmitir.Tag = grdNotifica.SelectedItem.SubItems(1) & "|" & grdNotifica.SelectedItem.SubItems(2) & "|" & grdNotifica.SelectedItem.SubItems(3)
            mnuCancelamento.Caption = "Cancelar Notificação nº " & grdNotifica.SelectedItem
            mnuCancelamento.Enabled = IIf(UCase(Aplicacoes.Usuario) = "LUCIANA" Or UCase(Aplicacoes.Usuario) = "ANDRE", True, False)
            Me.PopupMenu mnuNotifica
        End If
    End If
End Sub

Private Sub mnuCancelamento_Click()
    If Util.Confirma("Cancelar a notificação " & grdNotifica.SelectedItem & "?") Then
        If Bdados.DeletaDados("TAB_PAGAMENTO_NOTIFICACAO", "TPN_TNO_COD_NOTIFICACAO = " & grdNotifica.SelectedItem) Then
            If Bdados.DeletaDados("TAB_GERACAO_TRIBUTO", "TGT_COD_PAGAMENTO =" & grdNotifica.SelectedItem) Then
                If Bdados.DeletaDados("TAB_NOTIFICACAO", "TNT_COD_NOTIFICACAO = " & grdNotifica.SelectedItem) Then
                    Util.Mensagem "Notíficação nº " & grdNotifica.SelectedItem & " excluida."
                    Edita.LimpaCampos Me
                    txtRefInicio.SetFocus
                    Call cmdParcela_Click
                End If
            End If
        End If
    End If
End Sub

Private Sub mnuEmitir_Click()
    ImprimirNotificacao grdNotifica.SelectedItem
End Sub

Private Sub txtic_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Notific As New Notificacao
    Notific.ExibirNotificacoes grdNotifica, txtIm
    
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        'KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
    End If
End Sub

Private Sub ImprimirNotificacao(Optional Im As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String, Optional Tipo As Integer, Optional SeqInicial As Double, Optional SeqFinal As Double)
    On Error GoTo Trata
    Dim Cobranca As New VSCobranca
    Dim SelecaoRpt As String
    Dim CONDICAO As String
    Screen.MousePointer = 11
    CONDICAO = ""
    If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TNotifEmitidas.rpt") Then Exit Sub
    Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
    With Rpt
        If Tipo > 0 Then
            CONDICAO = " {Tab_Notificacao.TNT_TIPO_NOTIFICACAO} =" & Tipo
        End If
        If Trim$(Im) <> "" Then
            CONDICAO = CONDICAO & IIf(Trim(CONDICAO) = "", "", " and ") & " {Tab_Notificacao.TNT_INSCRICAO}='" & Im & "'"
        End If
        If Trim$(PeriodoInicial) <> "" And Trim$(PeriodoFinal) <> "" Then
            CONDICAO = CONDICAO & IIf(Trim(CONDICAO) = "", "", " and ") & "  ({Tab_Notificacao.TNT_DT_EMISSAO} in  " & _
                    "Date (" & Year(PeriodoInicial) & "," & Month(PeriodoInicial) & "," & Day(PeriodoInicial) & ") to Date " & _
                    "(" & Year(PeriodoFinal) & "," & Month(PeriodoFinal) & "," & Day(PeriodoFinal) & "))"
        End If
        If SeqInicial > 0 Then
            CONDICAO = CONDICAO & IIf(Trim(CONDICAO) = "", "", " and ") & " {Tab_Notificacao.tnt_cod_notificacao} >= " & SeqInicial
        End If
        If SeqFinal > 0 Then
            CONDICAO = CONDICAO & IIf(Trim(CONDICAO) = "", "", " and ") & " {Tab_Notificacao.tnt_cod_notificacao} <= " & SeqFinal
        End If
        SelecaoRpt = CONDICAO
        .Selecao = SelecaoRpt
        
        .Arvore = False
        .Visualizar
    End With
    Set Rpt = Nothing
    Screen.MousePointer = 0
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Avisa Err.Description
        Resume
        Err.Clear
    End If

End Sub

Private Sub txtIM_LostFocus()
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco)
End Sub


