VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TNOT103 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   25
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TNOT103.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.cboVISUAL cboTributo 
      Height          =   315
      Left            =   765
      TabIndex        =   0
      Top             =   675
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   556
      Caption         =   "Tributo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin ActiveTabs.SSActiveTabs tabCND 
      Height          =   3330
      Left            =   1350
      TabIndex        =   18
      Top             =   2400
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   5874
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
      Tabs            =   "TNOT103.frx":2123
      Images          =   "TNOT103.frx":21D4
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   2910
         Index           =   0
         Left            =   -99969
         TabIndex        =   19
         Top             =   30
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   5133
         _Version        =   131082
         TabGuid         =   "TNOT103.frx":2E6D
         Begin VB.TextBox txtTexto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   2835
            Left            =   45
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   30
            Width           =   8070
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   2910
         Index           =   1
         Left            =   -99969
         TabIndex        =   20
         Top             =   30
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   5133
         _Version        =   131082
         TabGuid         =   "TNOT103.frx":2E95
         Begin VTOcx.grdVISUAL lstNot 
            Height          =   2880
            Left            =   30
            TabIndex        =   22
            Top             =   30
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   4339
            CorFundo        =   -2147483633
            Caption         =   "Débitos em aberto"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   2910
         Left            =   30
         TabIndex        =   23
         Top             =   30
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   5133
         _Version        =   131082
         TabGuid         =   "TNOT103.frx":2EBD
         Begin VTOcx.grdVISUAL grdCND 
            Height          =   2865
            Left            =   30
            TabIndex        =   24
            Top             =   45
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   4339
            CorFundo        =   -2147483633
            Caption         =   "Certidões Positivas com efeito de Negativas emitidas"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
   End
   Begin VTOcx.txtVISUAL txtValidade 
      Height          =   300
      Left            =   7560
      TabIndex        =   6
      Tag             =   "Validade"
      Top             =   1725
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   529
      Caption         =   "Validade"
      Text            =   ""
      Formato         =   0
      Restricao       =   2
      MaxLen          =   10
   End
   Begin VTOcx.txtVISUAL txtFinalidade 
      Height          =   300
      Left            =   480
      TabIndex        =   5
      Tag             =   "Finalidade"
      Top             =   1725
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   529
      Caption         =   "Finalidade"
      Text            =   ""
   End
   Begin VB.TextBox txtIc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   1395
      TabIndex        =   3
      Top             =   1380
      Width           =   1905
   End
   Begin VB.TextBox txtEndereco 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Height          =   315
      Left            =   3690
      TabIndex        =   16
      Top             =   1380
      Width           =   5865
   End
   Begin VB.TextBox txtContrib 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Height          =   315
      Left            =   2970
      TabIndex        =   14
      Top             =   1035
      Width           =   6585
   End
   Begin VB.TextBox txtIm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   1395
      TabIndex        =   1
      Top             =   1035
      Width           =   1185
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   5820
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   873
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdParcela 
         Height          =   375
         Left            =   6840
         TabIndex        =   9
         Top             =   75
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Emitir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8805
         TabIndex        =   11
         Top             =   75
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   7800
         TabIndex        =   10
         Top             =   75
         Width           =   975
         _ExtentX        =   1720
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
      TabIndex        =   12
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TNOT103.frx":2EE5
   End
   Begin Threed.SSPanel lbl 
      Height          =   270
      Index           =   15
      Left            =   30
      TabIndex        =   15
      Top             =   1050
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   476
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
      Caption         =   "Insc. Municipal"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   4
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel lbl 
      Height          =   270
      Index           =   5
      Left            =   0
      TabIndex        =   17
      Top             =   1395
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   476
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
      Caption         =   "Insc. Cadastral"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   4
      RoundedCorners  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtRefInicio 
      Height          =   300
      Left            =   690
      TabIndex        =   7
      Tag             =   "Periodo Inicial"
      Top             =   2055
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   529
      Caption         =   "Periodo"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   6
   End
   Begin VTOcx.txtVISUAL txtRefFim 
      Height          =   300
      Left            =   2745
      TabIndex        =   8
      Tag             =   "Periodo Final"
      Top             =   2055
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      Caption         =   "até"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   6
   End
   Begin VTOcx.cmdVISUAL cmdPesquisaIM 
      Height          =   315
      Left            =   2610
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1030
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.cmdVISUAL cmdPesquisaIC 
      Height          =   315
      Left            =   3330
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1380
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
   Begin VB.Menu mnuCND 
      Caption         =   "."
      Visible         =   0   'False
      Begin VB.Menu mnuEmitir 
         Caption         =   "&Imprimir certidão ..."
      End
   End
End
Attribute VB_Name = "TNOT103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim sql As String
Dim Cgc As String

Sub ImprimeCertidao(CodCertidao As String)
    Dim a As Byte
    Dim sql As String
    Dim rs As VSRecordset
    
'    If Temp.PegaParametro(Bdados, "MODELO CERTIDAO") = "2" Then
'        Sql = "SELECT * FROM VIS_CERTIDAO_NEGATIVA"
'        Sql = Sql & " WHERE TCN_COD_NEGATIVA = " & CodCertidao
'        VisualizarActiveReport AR_CND, Bdados, Sql
'        Exit Sub
'    End If
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TCN.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Selecao = "{TAB_CERTIDAO_NEGATIVA.TCN_COD_NEGATIVA} = " & CodCertidao
        .Titulo = "Certidão Positiva com Efeito de Negativa de Débitos"
        .Arvore = False
        .Visualizar
         Bdados.FechaTabela rs
    End With
End Sub

Private Sub cmdCancela_Click()
    Dim rs As VSRecordset
    Dim sql As String
    Edita.LimpaCampos Me
    grdCND.Preencher Bdados, ""
    tabCND.Tabs(1).Selected = True
    cboTributo.SetFocus
    lstNot.ListItems.Clear
    lstNot.Mensagem = ""
    sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = 'CERTIDAO NEGATIVA'"
    If Bdados.AbreTabela(sql, rs) Then
        txtTexto = "" & rs!TPT_TEXTO
    End If
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdParcela_Click()
    Dim Conta As New ContaCorrente
    Dim CodImp As String
    Dim rs As VSRecordset
    Dim RsConta As VSRecordset
    Dim CodArr As String
    Dim CodLanc As String
    Dim Valores As String
    Dim CodCertidao As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
    If Trim(txtIm.Text) = "" Or txtContrib.Text = "" Then
        Util.Avisa "Informe uma Inscrição Municipal válida"
        txtIm.SetFocus
        Exit Sub
    End If
    
    If Trim(txtIc.Text) <> "" And txtEndereco.Text = "" Then
        Util.Avisa "Informe uma Inscrição Cadastral válida"
        txtIc.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Bdados.GravaDados "TAB_PARAMETRO_TEXTO", Bdados.PreparaValor(txtTexto), "TPT_TEXTO", "TPT_PARAMETRO = 'CERTIDAO NEGATIVA'"
    
    If Trim(txtIm) <> "" Then
        CodArr = " tdr_im ='" & txtIm & "'"
        CodLanc = " tgt_im ='" & txtIm & "'"
    End If
    If Trim(txtIc) <> "" Then
        CodArr = " tdr_tim_ic ='" & txtIc & "'"
        CodLanc = " tgt_tim_ic ='" & txtIc & "'"
    End If
    sql = "SELECT tgt_cod_pagamento AS Documento, tgt_tip_cod_imposto AS Imposto ,tgt_periodo as Periodo,tgt_Valor_tributo as Valor" & _
        ", TGT_DATA_VENCIMENTO AS Vencimento FROM TAB_GERACAO_TRIBUTO " & _
        " where tgt_COD_PAGAMENTO NOT IN ( SELECT TDR_tgt_cod_pagamento FROM TAB_DARM_RECEBIDO WHERE " & CodArr _
        & " ) AND " & CodLanc & " AND TGT_TIP_COD_IMPOSTO <> '" & Const_Extrato
    sql = sql & "' AND tgt_tip_cod_imposto <>'" & Const_Notificacao & "'"
    
'*-*-*-*-*/-*CERTIDAO NEGATIVA - IMPRESSAO
    If Bdados.AbreTabela(sql, rs) Then
        If Not Confirma("Contribuinte inadimplente. Confirma impressão de Certidão negativa.") Then
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    CodCertidao = Conta.GeraCodPagamento("37")
    DoEvents
    Valores = Bdados.PreparaValor(CodCertidao, txtIm, Format(Date, "DD/MM/YYYY"), Bdados.Converte(txtIc, tctexto), txtFinalidade, txtValidade, txtRefInicio, txtRefFim, Aplicacoes.Usuario, 30, cboTributo.Coluna(0).Valor)
    If Bdados.InsereDados("TAB_CERTIDAO_NEGATIVA", Valores, "TCN_COD_NEGATIVA,TCN_TCI_IM,TCN_DATA_NEGATIVA,TCN_TIM_IC,TCN_FINALIDADE, TCN_VALIDADE, TCN_PERIODO_INICIAL, TCN_PERIODO_FINAL, TCN_TUS_COD_USUARIO, TCN_TIPO, TCN_TIP_COD_IMPOSTO") Then
        cmdCancela_Click
        If Util.Confirma("Certidão emitida com sucesso. Deseja imprimir?") Then
            ImprimeCertidao CodCertidao
        End If
    End If
    lstNot.Preencher Bdados, sql, 1400
    If lstNot.ListItems.Count > 0 Then lstNot.Mensagem = "Total da dívida: R$" & Format(lstNot.Colunas(4).Soma, Const_Monetario)
    tabCND.Tabs(3).Selected = True
    Bdados.FechaTabela rs
    Screen.MousePointer = 0
End Sub

Private Sub cmdPesquisaIC_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtIc
End Sub

Private Sub cmdPesquisaIM_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    AtualizaCabecalho lstNot
    Dim Obrig As New Obrigacao
    Dim sql As String
    Dim rs As VSRecordset
    sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = 'CPND'"
    If Bdados.AbreTabela(sql, rs) Then
        txtTexto = "" & rs!TPT_TEXTO
    End If
    txtValidade = DateAdd("d", 120, Now)
    Obrig.PreencheComboTributo cboTributo, False
End Sub

Private Sub lstNot_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstNot, ColumnHeader
End Sub

Private Sub grdCND_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not grdCND.SelectedItem Is Nothing Then
        If Button = 2 Then
            mnuEmitir.Caption = "Imprimir CPND " & grdCND.SelectedItem
            Me.PopupMenu mnuCND
        End If
    End If
End Sub

Private Sub mnuEmitir_Click()
    ImprimeCertidao Util.ParseString(mnuEmitir.Caption, "CPND", 2)
End Sub

Private Sub txtIc_Change()
    txtEndereco.Text = ""
End Sub

Private Sub txtic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtic_LostFocus()
    

    CarregaEnderecoImovel txtIc, txtEndereco, txtIm
    txtIm_LostFocus
    'BuscarCertidoes txtIm, txtIc
    
End Sub

Private Sub txtIm_Change()
    txtContrib.Text = ""
    If Me.ActiveControl Is txtIm Then
        txtIc.Text = ""
    End If
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
    End If
End Sub

Private Sub txtIm_LostFocus()
    Dim sql As String
    Dim rs As VSRecordset
    
    If Trim(txtIm) = "" Then Exit Sub
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIm = Imposto.FormataInscricao(txtIm, InscContrib)
    End If
    sql = "Select tci_nome,tci_logradouro,tci_nome_logradouro," & _
    " tci_numero,tci_complemento,tci_bairro,tci_cidade,tci_uf,tci_cgc_cpf FROM tab_Contribuinte where tci_im='" & txtIm & _
    "' and tci_tsc_cod_sit_cad=1"
    If Bdados.AbreTabela(sql, rs) Then
        txtContrib = "" & rs!tci_nome
        Cgc = "" & rs!TCI_CGC_CPF
        BuscarCertidoes txtIm, txtIc
    Else
        Avisa "Contribuinte não Cadastrado."
        txtIm.SetFocus
    End If
    Bdados.FechaTabela rs
End Sub

Public Sub BuscarCertidoes(Optional Im As String, Optional Ic As String)
    Dim sql As String, Condicao As String
    sql = "SELECT TCN_COD_NEGATIVA AS Numero, " & _
            " TCN_FINALIDADE as Finalidade, " & _
            " TCN_PERIODO_INICIAL as [Per. Inicial], " & _
            " TCN_PERIODO_FINAL as [Per. Final], " & _
            " TCN_VALIDADE as Validade, " & _
            " TCN_DATA_NEGATIVA as Emissao " & _
        " FROM TAB_CERTIDAO_NEGATIVA"
    Condicao = "TCN_TIPO = 30" 'CPND
    If Trim$(Im) <> "" Then
        Condicao = Condicao + " AND TCN_TCI_IM='" & Im & "'"
    End If
    If Trim$(Ic) <> "" Then
        Condicao = Condicao & IIf(Condicao = "", "", " AND ") & "TCN_TIM_IC='" & Ic & "'"
    End If
    If Condicao <> "" Then
        sql = sql & " WHERE " & Condicao
    End If
    sql = sql & " ORDER BY TCN_VALIDADE DESC"
    grdCND.Preencher Bdados, sql
    tabCND.Tabs(1).Selected = True
End Sub
