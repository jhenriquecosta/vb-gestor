VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TOBR106 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tcob103"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdLancamento 
      Height          =   2535
      Left            =   60
      TabIndex        =   10
      Top             =   2400
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   4471
      CorBorda        =   16777215
      Caption         =   "Valores Lançados"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   128
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   8
      Top             =   5280
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8265
         TabIndex        =   5
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdGera 
         Height          =   375
         Left            =   7290
         TabIndex        =   4
         Top             =   90
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         Caption         =   "&Gerar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9330
         TabIndex        =   6
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8700
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   7
      Top             =   5730
      Visible         =   0   'False
      Width           =   795
   End
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   1875
      Left            =   0
      TabIndex        =   9
      Top             =   570
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   3307
      Caption         =   "Opções de Filtro"
      Descricao       =   "Fornece critérios para identificar os contribuintes"
      corFaixa        =   -2147483633
      corFundo        =   -2147483633
      Icone           =   "TOBR106.frx":0000
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.txtVISUAL txtExercicio 
         Height          =   315
         Left            =   5400
         TabIndex        =   16
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "Exercício"
         Text            =   ""
         TipoLetras      =   0
         Restricao       =   2
         MaxLen          =   4
      End
      Begin VTOcx.cmdVISUAL CmdFinal 
         Height          =   300
         Left            =   8070
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   690
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.cmdVISUAL cmdPesqim 
         Height          =   300
         Left            =   3600
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   660
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.cboVISUAL cboAtividade 
         Height          =   315
         Left            =   870
         TabIndex        =   3
         Top             =   1440
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   556
         Caption         =   "Atividade"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboGrupoAtividade 
         Height          =   315
         Left            =   30
         TabIndex        =   2
         Top             =   1050
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   556
         Caption         =   "Grupo de atividade"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIMFinal 
         Height          =   315
         Left            =   5520
         TabIndex        =   1
         Top             =   660
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         Caption         =   "IM Final"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtIMInicial 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   660
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         Caption         =   "IM Inicial"
         Text            =   ""
      End
   End
   Begin Threed.SSPanel pnlProgresso 
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   4950
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   397
      _Version        =   196610
      ForeColor       =   -2147483645
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      FloodType       =   1
      FloodColor      =   -2147483646
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1138
      Icone           =   "TOBR106.frx":1D0A
   End
   Begin VB.Label lblContribuinte 
      AutoSize        =   -1  'True
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   4740
      TabIndex        =   13
      Top             =   4980
      Width           =   45
   End
   Begin VB.Label lblGuias 
      AutoSize        =   -1  'True
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   4080
      TabIndex        =   12
      Top             =   4980
      Width           =   45
   End
End
Attribute VB_Name = "TOBR106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imprime As Boolean

Sub GeraAlvara(pInscricao As String, pCgcCpf As String, pRazao As String, pEndereco As String, pAtividade As String, pPeriodo As Integer, pTaxaServico As Double, pNomeImposto As String, ByVal pDataVenc As String, pCodImposto As String)
    Dim Valores As String
    Dim Campos As String
    Dim CodPagamento As String
    Dim Juros As Double
    Dim Multa As Double
    Dim Imposto As New VSImposto
    Dim Conta As New ContaCorrente
    Dim ValorAlvara As Double
    Dim Rpt As New VSRelatorio
    Dim Cobranca As New VSCobranca
    Dim TotalImposto As Double
    Dim Obrig As New Obrigacao
    Dim Calculo As New CalculoObrigacao
    On Error Resume Next
  
    CodPagamento = Obrig.CriaObrigacao(pCodImposto, CStr(pPeriodo), CStr(pPeriodo), pInscricao, , , etsSubstitui, pDataVenc)
    If Imprime Then
        If Obrig.BuscaDetalheObrigacao(CodPagamento, etiContribuinte) Then
            Cobranca.ImprimeDam Rpt, CodPagamento, Obrig.obContribuinte, pRazao, "" & pCgcCpf, pEndereco, "", "", _
                    pCodImposto, "Alvará", pNomeImposto, CStr(pPeriodo), 0, 1, pDataVenc, 0, Obrig.obValorObrigacao, _
                    CStr(Multa), CStr(Juros), 0, "0", pAtividade, "", PicBarra, , , , , , , , , , , tdiImpressora, etdNormal
        End If
    End If
    DoEvents
End Sub

Private Sub cboGrupoAtividade_Click()
    cboAtividade.Preencher Bdados, "SELECT  tae_nome, tae_cae FROM TAB_ATIVIDADE_ECONOMICA where tae_tga_cod_grupo=" & cboGrupoAtividade.Coluna(1).Valor & " or tae_tga_cod_grupo=0 order by tae_nome"
End Sub

Private Sub cmdEnter_Click()
    'SendKeys "{tab}"
End Sub

Private Sub CmdFinal_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIMFinal
End Sub

Private Sub cmdGera_Click()
    Dim Sql As String, Filtro As String
    Dim rs As VSRecordset
    Dim RSPago As VSRecordset
    Dim NomeImposto As String
    Dim DataVenc As String
    Dim CodImposto As String
    Dim Gerar As Boolean
    Dim Obrig As New Obrigacao
    Dim i, t As Integer
    On Error Resume Next
    If Not Util.Confirma("Gerar os Alvarás informados?") Then Exit Sub
    grdLancamento.Preencher Bdados, ""
    Screen.MousePointer = 11
        
    Imprime = Confirma("Deseja imprimir os lançamentos gerados?")
    pnlProgresso.Visible = True
    lblGuias.Visible = True
    lblContribuinte.Visible = True
    
    Sql = "Select tip_nome_imposto,tip_cod_imposto from tab_imposto where tip_sigla_imposto ='" & Imposto.NomeTributo(ttr_ALVARA) & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        NomeImposto = rs!tip_nome_imposto
        CodImposto = rs!tip_cod_imposto
    Else
        Avisa "Imposto não identificado."
        pnlProgresso.Visible = False
        lblGuias.Visible = False
        lblContribuinte.Visible = False
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    DataVenc = Imposto.BuscaDataVencimento(CodImposto, txtExercicio)
    Sql = "Select " & _
                " tci_im," & _
                " tae_nome," & _
                " tci_inicio_atividade," & _
                " tci_nome,tci_cgc_cpf," & _
                " tci_logradouro," & _
                " tci_nome_logradouro," & _
                " tci_numero," & _
                " tci_complemento," & _
                " tci_bairro," & _
                " tci_cep," & _
                " tci_cidade," & _
                " tci_UF " & _
            " from tab_contribuinte," & _
                " tab_atividade_economica " & _
            " WHERE tci_tae_cae=tae_cae " & _
                " and tci_tae_cae > 0 " '& _
    Sql = Sql & " AND TCI_CIDADE = '" & AplicacoesVTFuncoes.municipio & "'"
    Sql = Sql & " AND TCI_CIDADE = '" & AplicacoesVTFuncoes.municipio & "'"
    Filtro = ""
    If Trim$(txtIMInicial) <> "" Then Filtro = Filtro & " and tci_im >= '" & Imposto.FormataInscricao(Trim(txtIMInicial), InscContrib) & "'"
    If Trim$(txtIMFinal) <> "" Then Filtro = Filtro & " and tci_im <= '" & Imposto.FormataInscricao(Trim(txtIMFinal), InscContrib) & "'"
    If cboGrupoAtividade <> "" Then Filtro = Filtro & " and tae_tga_cod_grupo=" & cboGrupoAtividade.Coluna(1).Valor
    If cboAtividade <> "" Then Filtro = Filtro & " and tae_cae=" & cboAtividade.Coluna(1).Valor
    Sql = Sql & Filtro
    Sql = Sql & " order by tci_im"
        
    If Bdados.AbreTabela("select count(*) from tab_contribuinte,tab_atividade_economica where tci_tae_cae=tae_cae and tci_tae_cae > 0 AND TCI_CIDADE = '" & AplicacoesVTFuncoes.municipio & "'" & Filtro, rs) Then
        t = rs(0)
    End If
    Bdados.FechaTabela rs
    
    i = 0
    If Bdados.AbreTabela(Sql, rs) Then
        rs.MoveFirst
        Do
            GeraAlvara "" & rs!TCI_IM, "" & rs!TCI_CGC_CPF, "" & rs!tci_nome, "" & rs!tci_logradouro & " " & _
                rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_BAIRRO & _
                " - CEP " & rs!tci_cep & " " & rs!tci_cidade & " " & rs!tci_UF, rs!tae_nome, txtExercicio, _
                0, NomeImposto, DataVenc, CodImposto
            
            i = i + 1
            lblGuias = i & "/" & t
            lblContribuinte = "  " & rs!TCI_IM & " " & rs!tci_nome
            pnlProgresso.FloodPercent = (i / t) * 100

            DoEvents
            rs.MoveNext
            DoEvents
        Loop While Not rs.EOF
        Informa "Geração Concluída."
        Sql = " SELECT TOC_COD_OBRIGACAO AS Documento,Toc_inscricao as Inscrição,"
        Sql = Sql & " toc_periodo as Período,toc_data_vencimento as Vencimento,"
        Sql = Sql & " toc_valor_obrigacao As Valor"
        Sql = Sql & " From TAB_OBRIGACAO_CONTRIBUINTE"
        Sql = Sql & " where toc_tip_cod_imposto = " & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) & " and toc_periodo = '" & txtExercicio & "'"
        If Trim(txtIMInicial) <> "" And Trim(txtIMFinal) <> "" Then
            Sql = Sql & " and TOC_INSCRICAO  >=   '" & Imposto.FormataInscricao(txtIMInicial, InscContrib) & "' and TOC_INSCRICAO  <= '" & Imposto.FormataInscricao(txtIMFinal, InscContrib) & "'"
        End If
        grdLancamento.Preencher Bdados, Sql
    Else
        Util.Informa "Não foi possível selecionar os alvarás as serem gerados," & vbCrLf & "Verifique se o cadastro correspondente a essa incrição está correto."
    End If
    
    pnlProgresso.Visible = False
    lblGuias.Visible = False
    lblContribuinte.Visible = False
    
    Screen.MousePointer = 0
    rs.Fechar
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtExercicio.SetFocus
End Sub

Private Sub cmdPesq_Click(Index As Integer)

End Sub

Private Sub cmdPesqim_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIMInicial
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    txtExercicio = Format(Now, "yyyy")
    cboGrupoAtividade.Preencher Bdados, "SELECT tga_nome,tga_cod_grupo FROM tab_grupo_atividade order by tga_nome"
    pnlProgresso.Visible = False
    lblGuias.Visible = False
    lblContribuinte.Visible = False
End Sub


Private Sub txtExercicioOld_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIMFinal_LostFocus()
    If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIMFinal = Imposto.FormataInscricao(txtIMFinal, InscContrib)
    End If
End Sub

Private Sub txtIMInicial_LostFocus()
    If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIMInicial = Imposto.FormataInscricao(txtIMInicial, InscContrib)
    End If
''    Call Imposto.FormataInscricao(txtIMInicial, InscContrib)
    If Not AplicacoesVTFuncoes.Usuario = "PETROLINA" Then
        If Trim$(txtIMFinal) = "" Then
            txtIMFinal = Imposto.FormataInscricao(txtIMInicial, InscContrib)
        End If
    End If
End Sub
