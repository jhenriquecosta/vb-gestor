VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TNOT104 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs tabNotificacao 
      Height          =   3600
      Left            =   45
      TabIndex        =   9
      Top             =   2685
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6350
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "TNOT104.frx":0000
      Images          =   "TNOT104.frx":00B1
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3180
         Index           =   0
         Left            =   -99969
         TabIndex        =   13
         Top             =   30
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   5609
         _Version        =   131082
         TabGuid         =   "TNOT104.frx":0D52
         Begin VTOcx.grdVISUAL lstNot 
            Height          =   3060
            Left            =   45
            TabIndex        =   15
            Top             =   60
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   5398
            CorFundo        =   -2147483633
            Caption         =   "Débitos em Aberto"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3180
         Index           =   1
         Left            =   -99969
         TabIndex        =   14
         Top             =   30
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   5609
         _Version        =   131082
         TabGuid         =   "TNOT104.frx":0D7A
         Begin VB.TextBox txtTexto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   3105
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   30
            Width           =   8415
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3180
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   5609
         _Version        =   131082
         TabGuid         =   "TNOT104.frx":0DA2
         Begin VTOcx.grdVISUAL grdNotifica 
            Height          =   3060
            Left            =   45
            TabIndex        =   8
            Top             =   60
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   5398
            CorFundo        =   -2147483633
            Caption         =   "Notificacoes emitidas"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   6360
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   873
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdImprime 
         Height          =   375
         Left            =   5565
         TabIndex        =   5
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdParcela 
         Height          =   375
         Left            =   4590
         TabIndex        =   4
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Gerar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7830
         TabIndex        =   7
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
         Left            =   6780
         TabIndex        =   6
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
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   60
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   0.953
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   1.005
      TabIndex        =   11
      Top             =   15
      Visible         =   0   'False
      Width           =   570
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TNOT104.frx":0DCA
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   570
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   1138
      Icone           =   "TNOT104.frx":2EED
   End
   Begin VTOcx.txtVISUAL txtPeriodoInicial 
      Height          =   285
      Left            =   195
      TabIndex        =   2
      Top             =   2265
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   503
      Caption         =   "Período"
      Text            =   ""
      MaxLen          =   4
      MinLen          =   4
   End
   Begin VTOcx.txtVISUAL txtPeriodoFinal 
      Height          =   285
      Left            =   1590
      TabIndex        =   3
      Top             =   2265
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   503
      Caption         =   "a"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   4
      MinLen          =   4
   End
   Begin VTOcx.txtVISUAL txtRazao 
      Height          =   315
      Left            =   315
      TabIndex        =   18
      Top             =   1230
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   556
      Caption         =   "Razão"
      Text            =   ""
      Enabled         =   0   'False
   End
   Begin VTOcx.txtVISUAL txtEndereco 
      Height          =   315
      Left            =   45
      TabIndex        =   19
      Top             =   1560
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   556
      Caption         =   "Endereço"
      Text            =   ""
      Enabled         =   0   'False
   End
   Begin VTOcx.txtVISUAL txtInscricao 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Tag             =   "Inscrição Cadastral"
      Top             =   915
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   529
      Caption         =   "Inscricao"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   20
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.cboVISUAL cboImposto 
      Height          =   315
      Left            =   225
      TabIndex        =   1
      Top             =   1905
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   556
      Caption         =   "Tributo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdPesquisaIM 
      Height          =   315
      Left            =   2340
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   900
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
         Caption         =   "&Imprimir notificação ..."
      End
   End
End
Attribute VB_Name = "TNOT104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim Sql As String
Dim CodPagamento  As Double

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
    cboImposto.Enabled = True
    txtInscricao.Enabled = True
    cmdParcela.Enabled = True
    Edita.LimpaCampos Me
    lstNot.ListItems.Clear
    lstNot.Mensagem = ""
    grdNotifica.Preencher Bdados, ""
    grdNotifica.Mensagem = ""
    CodPagamento = 0
    Sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = 'NOTIFICACAO LANCAMENTO'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtTexto = "" & rs!TPT_TEXTO
    End If
    cboImposto.SetFocus
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
    Dim Valores As String
    Dim Campos As String
    Dim Cobranca As New VSCobranca
    Dim ValorFinal As Double
    
  
    Screen.MousePointer = 11
    '1.
    Bdados.GravaDados "TAB_PARAMETRO_TEXTO", Bdados.PreparaValor(txtTexto), "TPT_TEXTO", "TPT_PARAMETRO = 'NOTIFICACAO LANCAMENTO'"
    
    '2.
    If CodPagamento = 0 Then
        CodPagamento = Conta.GeraCodPagamento(EtsNotificacao)
        Campos = "TPN_TNO_COD_NOTIFICACAO,TPN_TOC_COD_OBRIGACAO,TPN_SUB_VALOR,TPN_TIP_COD_IMPOSTO"
        For i = 1 To lstNot.ListItems.Count
            Valores = Bdados.PreparaValor(CodPagamento, lstNot.ListItems(i).Text, Bdados.Converte(lstNot.ListItems(i).SubItems(4), TCDuplo), lstNot.ListItems(i).SubItems(6))
            Bdados.InsereDados "TAB_PAGAMENTO_NOTIFICACAO", Valores, Campos
        Next
        If Not (lstNot.SelectedItem Is Nothing) Then
            ValorFinal = Format(lstNot.Colunas(5).Soma, Const_Monetario)
        Else
            If grdNotifica.ListItems.Count > 0 Then ValorFinal = grdNotifica.Colunas(4).Soma
        End If
        Conta.GeraPagamento txtInscricao, "", Const_Notificacao, Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), Date, CDbl(ValorFinal), 0, 0, CStr(CodPagamento), 0, 0, 0, , EtcCreditoTributario
        
        '3.
        Valores = Bdados.PreparaValor(CodPagamento, txtInscricao, Bdados.Converte(Format(Date, "DD/MM/YYYY"), TCDataHora), Bdados.Converte(Format(Date, "DD/MM/YYYY"), TCDataHora), ValorFinal, Aplicacoes.Usuario, 1, 1)
        Campos = "TNT_COD_NOTIFICACAO,TNT_INSCRICAO,TNT_DT_EMISSAO,TNT_VENCIMENTO,TNT_VALOR_NOTIFICACAO,TNT_TUS_COD_USUARIO,TNT_TIPO_NOTIFICACAO,TNT_TIPO"
        Bdados.InsereDados "TAB_NOTIFICACAO", Valores, Campos
    End If
    '4.
    'ImprimirNotificacao txtInscricao , Trim(Mid(cboImposto, Edita.PosPic(cboImposto, "-") + 1)), CStr(ValorFinal), Date, CodPagamento, cboDest.ListIndex, "", txtPeriodoInicial, txtPeriodoFinal
    ImAnterior = ""
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Avisa Err.Description
        Err.Clear
    End If
End Sub

Private Sub cmdPesquisaIM_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscricao
    If txtInscricao <> "" Then
        TXTINSCRICAO_LostFocus
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Call Edita.AtualizaCombo(Bdados, cboImposto, "Select TIP_sigla_IMPOSTO " & Bdados.Concatena & "  ' - '  " & Bdados.Concatena & "  tip_nome_imposto  From TAB_IMPOSTO")
    Dim rs As VSRecordset
    Dim Sql As String
    Sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = 'NOTIFICACAO LANCAMENTO'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtTexto = "" & rs!TPT_TEXTO
    End If
End Sub


Private Sub lstNot_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstNot, ColumnHeader
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub grdNotifica_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not grdNotifica.SelectedItem Is Nothing Then
        If Button = 2 Then
            mnuEmitir.Caption = "Emitir notificação nº " & grdNotifica.SelectedItem
            mnuEmitir.Tag = grdNotifica.SelectedItem.SubItems(1) & "|" & grdNotifica.SelectedItem.SubItems(2) & "|" & grdNotifica.SelectedItem.SubItems(3)
            Me.PopupMenu mnuNotifica
        End If
    End If

End Sub

Private Sub mnuEmitir_Click()
    ImprimirNotificacao grdNotifica.SelectedItem
End Sub



Private Sub txtIc_Change()
    CodPagamento = 0
End Sub

Private Sub TXTINSCRICAO_Change()
    CodPagamento = 0
End Sub

Private Sub TXTINSCRICAO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        'KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
    End If
End Sub

Private Sub TXTINSCRICAO_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    
    If Trim(txtInscricao) = "" Then Exit Sub
    txtInscricao = BuscaContribuinte(txtInscricao, txtRazao, txtEndereco)
    
End Sub

Private Sub ExibirNotificacoes(Optional Im As String, Optional Ic As String)
    Dim Sql As String
    Dim Condicao As String
    
    If Trim$(Im) <> "" Then
        Condicao = "TNT_INSCRICAO='" & Im & "'"
    End If
     Sql = "SELECT TNT_COD_NOTIFICACAO AS Numero, " & _
            " TNT_DT_EMISSAO as Emissao, " & _
            " TNT_VENCIMENTO as Vencimento, " & _
            FuncaoReal("TNT_VALOR_NOTIFICACAO") & " as Valor" & _
        " FROM TAB_NOTIFICACAO "

    If Condicao <> "" Then
        Sql = Sql & " WHERE " & Condicao
    End If
    Sql = Sql & " ORDER BY TNT_VENCIMENTO"
    grdNotifica.Preencher Bdados, Sql
    If grdNotifica.ListItems.Count > 0 Then
        If grdNotifica.ListItems.Count > 0 Then
            grdNotifica.Mensagem = "Total : " & Format(grdNotifica.Colunas(4).Soma, Const_Monetario)
        Else
            grdNotifica.Mensagem = ""
        End If
    End If
    tabNotificacao.Tabs(1).Selected = True
End Sub

Private Sub ImprimirNotificacao(Numero As String)
    On Error GoTo Trata
    Dim Cobranca As New VSCobranca
    Dim SelecaoRpt As String
    
    Screen.MousePointer = 11
    If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TNotifLancto.rpt") Then Exit Sub
    Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
    With Rpt
        '.Formulas "VT_Cidade", Aplicacoes.Municipio
        SelecaoRpt = "{Tab_Notificacao.TNT_COD_NOTIFICACAO} = " & Numero
        .Selecao = SelecaoRpt
        
        .Titulo = "Extrato de Lançamento de Créditos"
        .Arvore = False
        .Visualizar
    End With
    Set Rpt = Nothing
    Screen.MousePointer = 0
    Avisa "Impressão concluída."
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Avisa Err.Description
        Resume
        Err.Clear
    End If

End Sub




