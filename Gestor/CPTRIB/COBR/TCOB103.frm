VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCOB103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tcob103"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCOB103.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.grdVISUAL grdLancamento 
      Height          =   2535
      Left            =   60
      TabIndex        =   12
      Top             =   2400
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   4471
      CorBorda        =   32768
      Caption         =   "Valores Lançados"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   128
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   10
      Top             =   5280
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7785
         TabIndex        =   6
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdGera 
         Height          =   375
         Left            =   6810
         TabIndex        =   5
         Top             =   90
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         Caption         =   "&Gerar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8850
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
      TabIndex        =   8
      Top             =   5730
      Visible         =   0   'False
      Width           =   795
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   1138
      Icone           =   "TCOB103.frx":2123
   End
   Begin VTOcx.txtVISUAL txtExercicio 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   750
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "Exercício"
      Text            =   ""
      TipoLetras      =   0
      Restricao       =   2
      MaxLen          =   4
   End
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   1875
      Left            =   1560
      TabIndex        =   11
      Top             =   570
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   3307
      Caption         =   "Opções de Filtro"
      Descricao       =   "Fornece critérios para identificar os contribuintes"
      corFaixa        =   -2147483633
      corFundo        =   -2147483633
      Icone           =   "TCOB103.frx":243D
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.cmdVISUAL CmdFinal 
         Height          =   300
         Left            =   7470
         TabIndex        =   18
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
         Left            =   4200
         TabIndex        =   17
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
         Left            =   1470
         TabIndex        =   4
         Top             =   1440
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   556
         Caption         =   "Atividade"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboGrupoAtividade 
         Height          =   315
         Left            =   630
         TabIndex        =   3
         Top             =   1050
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         Caption         =   "Grupo de atividade"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIMFinal 
         Height          =   315
         Left            =   4920
         TabIndex        =   2
         Top             =   660
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         Caption         =   "IM Final"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtIMInicial 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
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
      TabIndex        =   13
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   4980
      Width           =   45
   End
End
Attribute VB_Name = "TCOB103"
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
    ' CALCULA O VALOR DO ALVARÁ PARA O CONTRIBUINTE INFORMADO
    ValorAlvara = Imposto.CalculaAlvara(pInscricao, pPeriodo, pTaxaServico, pNomeImposto, pDataVenc, CodPagamento, "")
    'TMPBALSAS
    Call Conta.CriaContaContribuinte(CodPagamento)
    'TMPBALSAS
    Multa = 0
    Juros = 0
    TotalImposto = CDbl(ValorAlvara) + CDbl(Multa) + CDbl(Juros)
    If Imprime Then
        Cobranca.ImprimeDam Rpt, CodPagamento, pInscricao, pRazao, "" & pCgcCpf, pEndereco, "", "", _
                pCodImposto, Imposto.NomeTributo(ttr_ALVARA), pNomeImposto, CStr(pPeriodo), 0, 1, pDataVenc, 0, CStr(ValorAlvara), _
                CStr(Multa), CStr(Juros), 0, "0", pAtividade, "", PicBarra, , , , , , , , , , , tdiImpressora
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
    
    Dim i, t As Integer
    
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
                " and tci_tae_cae > 0 "
    Filtro = ""
    If Trim$(txtIMInicial) <> "" Then Filtro = Filtro & " and tci_im >= '" & Imposto.FormataInscricao(Trim(txtIMInicial), InscContrib) & "'"
    If Trim$(txtIMFinal) <> "" Then Filtro = Filtro & " and tci_im <= '" & Imposto.FormataInscricao(Trim(txtIMFinal), InscContrib) & "'"
    If cboGrupoAtividade <> "" Then Filtro = Filtro & " and tae_tga_cod_grupo=" & cboGrupoAtividade.Coluna(1).Valor
    If cboAtividade <> "" Then Filtro = Filtro & " and tae_cae=" & cboAtividade.Coluna(1).Valor
    Sql = Sql & Filtro
    Sql = Sql & " order by tci_im"
        
    If Bdados.AbreTabela("select count(*) from tab_contribuinte,tab_atividade_economica where tci_tae_cae=tae_cae and tci_tae_cae > 0" & Filtro, rs) Then
        t = rs(0)
    End If
    Bdados.FechaTabela rs
    
    i = 0
    If Bdados.AbreTabela(Sql, rs) Then
        rs.MoveFirst
        Do
            'If Not Bdados.AbreTabela("Select tdr_tgt_cod_pagamento from  tab_darm_recebido where tdr_im = '" & Rs!TCI_IM & "' and tdr_periodo =" & txtExercicio & " and tdr_tip_cod_imposto = '" & CodImposto & "'", RSPago) Then
                GeraAlvara rs!TCI_IM, rs!TCI_CGC_CPF, rs!tci_nome, rs!tci_logradouro & " " & _
                rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_BAIRRO & _
                " - CEP " & rs!tci_cep & " " & rs!tci_cidade & " " & rs!tci_UF, rs!tae_nome, txtExercicio, _
                0, NomeImposto, DataVenc, CodImposto
            'End If
            
            i = i + 1
            lblGuias = i & "/" & t
            lblContribuinte = rs!TCI_IM & " " & rs!tci_nome
            pnlProgresso.FloodPercent = (i / t) * 100

            DoEvents
            rs.MoveNext
            DoEvents
        Loop While Not rs.EOF
        Informa "Geração Concluída."
        
        Sql = "SELECT tgt_im as IM, " & _
                            " tci_nome as Razão, " & _
                            " tae_nome as Atividade, " & _
                            " tae_valor as Valor, " & _
                            " tae_desc_fator as Multiplicador, " & _
                            " tci_fator_alvara as Fator, " & _
                            FuncaoReal("tgt_valor_tributo") & " as Tributo " & _
                " FROM TAB_GERACAO_TRIBUTO, TAB_CONTRIBUINTE, TAB_ATIVIDADE_ECONOMICA" & _
                " WHERE tgt_im=tci_im " & _
                    " and tae_cae=tci_tae_cae" & _
                    " and tgt_tip_cod_imposto=" & CodImposto & _
                    " and tgt_periodo=" & txtExercicio
        Sql = Sql & Filtro
        Sql = Sql & " order by tgt_im"
        If grdLancamento.Preencher(Bdados, Sql) Then
            grdLancamento.Mensagem = "Soma : " & Format$(grdLancamento.Colunas(7).Soma, "currency") & " x Menor : " & Format$(grdLancamento.Colunas(7).Min, "currency") & " x Maior : " & Format$(grdLancamento.Colunas(7).Max, "currency") & " x Média : " & Format$(grdLancamento.Colunas(7).Media, "currency")
        End If

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
    cabVisual.Exibir Bdados, Me.Name, App.Path
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
   ' txtIMFinal = Imposto.FormataInscricao(txtIMFinal, InscContrib)
End Sub

Private Sub txtIMInicial_LostFocus()
'    txtIMInicial = Imposto.FormataInscricao(txtIMInicial, InscContrib)
''    Call Imposto.FormataInscricao(txtIMInicial, InscContrib)
    If Not AplicacoesVTFuncoes.Usuario = "PETROLINA" Then
        If Trim$(txtIMFinal) = "" Then
            txtIMFinal = Imposto.FormataInscricao(txtIMInicial, InscContrib)
        End If
    End If
End Sub
