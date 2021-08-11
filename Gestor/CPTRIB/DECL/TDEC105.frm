VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TDEC105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DECL101"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TDEC105.frx":0000
   ScaleHeight     =   8220
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   27
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TDEC105.frx":0342
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.PictureBox TabDec 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4440
      Left            =   30
      ScaleHeight     =   4380
      ScaleWidth      =   11250
      TabIndex        =   8
      Top             =   3270
      Width           =   11310
      Begin VB.PictureBox SSActiveTabPanel1 
         Height          =   4050
         Index           =   0
         Left            =   30
         ScaleHeight     =   3990
         ScaleWidth      =   11190
         TabIndex        =   9
         Top             =   30
         Width           =   11250
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   3885
            Left            =   30
            TabIndex        =   14
            Top             =   60
            Width           =   11190
            _ExtentX        =   19738
            _ExtentY        =   6853
            _Version        =   393216
            GridLines       =   2
         End
         Begin VB.TextBox txtGrig 
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   9810
            TabIndex        =   15
            Text            =   "txtGrig"
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.PictureBox SSActiveTabPanel1 
         Height          =   4050
         Index           =   1
         Left            =   -99969
         ScaleHeight     =   3990
         ScaleWidth      =   11190
         TabIndex        =   17
         Top             =   30
         Width           =   11250
         Begin VTOcx.cmdVISUAL cmdRemover 
            Height          =   330
            Left            =   9375
            TabIndex        =   18
            Top             =   630
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   582
            Caption         =   "Remover Item"
            Acao            =   2
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtISS 
            Height          =   480
            Left            =   9945
            TabIndex        =   19
            Tag             =   "ISS"
            Top             =   3510
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   847
            Caption         =   "ISS Devido"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtTotalNota 
            Height          =   480
            Left            =   8520
            TabIndex        =   20
            Tag             =   "Total da Nota"
            Top             =   3510
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   847
            Caption         =   "Total(R$)"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.cmdVISUAL cmdInclui 
            Height          =   330
            Left            =   7545
            TabIndex        =   21
            Top             =   630
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   582
            Caption         =   "Adicionar Item"
            Acao            =   1
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtValorTotal 
            Height          =   480
            Left            =   3150
            TabIndex        =   22
            Top             =   510
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   847
            Caption         =   "Valor Total(R$)"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtValorUnitario 
            Height          =   480
            Left            =   1410
            TabIndex        =   23
            Top             =   510
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   847
            Caption         =   "Valor Unitário(R$)"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtQnt 
            Height          =   480
            Left            =   345
            TabIndex        =   24
            Top             =   510
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   847
            Caption         =   "Quantidade"
            Text            =   ""
            Restricao       =   2
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtDescServico 
            Height          =   480
            Left            =   330
            TabIndex        =   25
            Top             =   0
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   847
            Caption         =   "Descrição de Serviço"
            Text            =   ""
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.grdVISUAL grdNota 
            Height          =   2760
            Left            =   330
            TabIndex        =   26
            Top             =   1020
            Width           =   10830
            _ExtentX        =   19103
            _ExtentY        =   4868
            CorBorda        =   32768
            Caption         =   "Itens"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
            OcultarRodape   =   -1  'True
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   7
      Top             =   7695
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdFinaliza 
         Height          =   375
         Left            =   6720
         TabIndex        =   3
         Top             =   90
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   661
         Caption         =   "Finalizar Declaracão"
         Acao            =   1
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmLimpar 
         Height          =   375
         Left            =   8820
         TabIndex        =   4
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4770
         TabIndex        =   2
         Top             =   90
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Salvar Declaracão"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10110
         TabIndex        =   5
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
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
      TabIndex        =   6
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   1138
      Icone           =   "TDEC105.frx":2465
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1080
      Left            =   30
      TabIndex        =   10
      Top             =   660
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   1905
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   330
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   556
         Caption         =   "Insc.Municipal"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisar 
         Height          =   345
         Left            =   10740
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   660
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   609
         Caption         =   ""
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   285
         Left            =   300
         TabIndex        =   0
         Top             =   690
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         Caption         =   "Período"
         Text            =   ""
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   4710
         TabIndex        =   11
         Top             =   330
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         Caption         =   "Razão Social"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   315
         Left            =   4410
         TabIndex        =   1
         Top             =   690
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   556
         Caption         =   "Tipo Declaracão"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin VTOcx.grdVISUAL grdDec 
      Height          =   1725
      Left            =   60
      TabIndex        =   16
      Top             =   1770
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   3043
      Caption         =   "Declaracões"
      CorTitulo       =   5346129
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "mnuGeral"
      Visible         =   0   'False
      Begin VB.Menu mnuDeletar 
         Caption         =   "Deletar"
      End
   End
End
Attribute VB_Name = "TDEC105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declaracao As VsTFuncoes.cDeclaracao
Private GerarIM As Boolean
Private AliqISSQN As Double, ISSQNFixo As Double
Private AliqISSST As Double, ISSSTFixo As Double

Private TotalImpostoST As Double
Private TotalBaseST As Double
Private TotalImpostoDevidoSaida As Double
Private TotalImpostoRetidoSaida As Double
Private TotalBaseSaida As Double
Private TotalICMSSujeito As Double
Private DeduzValores As Boolean
Private ContribuinteEndereco As String
Private ContribuinteAtividade As String
Dim Notas() As New NotaFiscal
Dim Modalidade As Integer
Dim ClassGrid As New grdEditavel

Private Sub FormataRegistro(ByRef Inscricao As Object)
    Select Case Len(Inscricao.Text)
        Case 10
            Inscricao.Text = Imposto.FormataInscricao(Inscricao.Text, InscContrib)
        Case 11
            Inscricao.Text = Edita.FormataTexto(Inscricao, Cpf)
        Case 14
            Inscricao.Text = Edita.FormataTexto(Inscricao, Cgc)
    End Select
End Sub

Private Sub cmdPesquisar_Click()
    Declaracao.CarregaGrid grdDec, txtIM, txtPeriodo, CInt(cboTipo.Coluna(1).Valor)
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If txtIM = "" Then Exit Sub
    
    Declaracao.Im = txtIM
    Declaracao.Periodo = txtPeriodo
    Declaracao.Data = Format(Date, "dd/mm/yyyy")
    Declaracao.Origem = orgSistema
    Declaracao.Recepcao = Date
    Declaracao.VERSAO = Nvl(Temp.PegaParametro(Bdados, "VERSAO DEC"), 2)
    Declaracao.Tipo = cboTipo.Coluna(1).Valor
    
    Declaracao.BaseGeral = TotalBaseSaida + TotalBaseST
    Declaracao.Status = decAberta
    If Declaracao.Gravar() Then
        Avisa "Declaração gravada com sucesso."
        txtIM.SetFocus
    End If
End Sub

Private Sub cmLimpar_Click()
    Edita.LimpaCampos Me
    txtIM.SetFocus
End Sub

Private Sub Form_Load()
    Dim Sql As String
    
    cabVISUAL1.Exibir Bdados, Me.Tag, App.Path
    rodVISUAL1.Exibir Bdados, Me.Tag
    Set Imposto = New VsTFuncoes.VSImposto
    DeduzValores = True
    Set Declaracao = New cDeclaracao
    cboTipo.PreencherGeral Bdados, "TIPO DECLARACAO"
    Sql = "SELECT TCD_COD_CAMPO as Item ,TCD_CAMPO as Descricao, ' ' as Valor FROM " & _
        "TAB_CONTEUDO_DECLARACAO WHERE TCD_TMD_COD_DECLARACAO = " & 1
    ClassGrid.CarregaGrid Grid, Sql
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Declaracao = Nothing
End Sub

Private Sub Grid_DblClick()
ClassGrid.EditaCelula Grid, txtGrig
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then ClassGrid.TeclaPressionada Grid, txtGrig, KeyCode
End Sub

Private Sub txtGrig_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then ClassGrid.TextoKeyDown KeyCode, Grid, txtGrig
End Sub

Private Sub txtGrig_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtGrig_LostFocus()
    txtGrig = Format(txtGrig, Const_Monetario)
End Sub

Private Sub txtIM_LostFocus()
    Dim atividade As VsTEcon.atividade
    
    If Trim$(txtIM) <> "" Then
        If Not BuscarContribuinte(txtIM, txtRazao) Then
            Avisa "Contribuinte não encontrado."
            txtIM = "": txtRazao = ""
            txtIM.SetFocus
        Else
            Set atividade = New VsTEcon.atividade
            AliqISSQN = atividade.BuscaAliquotaAtividade(Bdados, txtIM, txtPeriodo, ISSQNFixo)
            
            Declaracao.tciAtividade = atividade.Nome
'            If Len(Trim(txtPeriodo)) = 7 And cboTipo.ListIndex <> -1 Then PreencheDeclaracao
            Modalidade = BuscaModalidadeDeclaracao(txtIM)
            TabDec.Tabs(4).Enabled = IIf(Modalidade > 0, True, False)
            Set atividade = Nothing
        End If
    End If
End Sub


Private Sub txtPeriodo_Change()
    Dim atividade As New VsTEcon.atividade
    
    
    If Len(Trim(txtPeriodo)) <> 7 Then Exit Sub
'    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
    AliqISSQN = atividade.BuscaAliquotaAtividade(Bdados, txtIM, txtPeriodo, ISSQNFixo)
    If CInt(Left(Trim(txtPeriodo), 2)) > 12 Or CInt(Left(Trim(txtPeriodo), 2)) < 1 Then
        Avisa "Periodo inválido."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    
    
End Sub

Private Sub txtPeriodo_LostFocus()
    If Len(Trim(txtPeriodo)) <> 6 Then Exit Sub
    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
End Sub


Private Function BuscarContribuinte(ByRef Inscricao As Object, Optional ByRef Nome As Object, Optional ByRef Endereco As Object, _
                    Optional ByRef Bairro As Object, Optional ByRef Cep As Object, Optional ByRef Cidade As Object, Optional ByRef Uf As Object) As Boolean
    Dim Im As Boolean
    Im = False
    If Trim(Inscricao) = "" Then Exit Function
    Inscricao.Text = Edita.TiraTudo(Inscricao.Text)
    If Len(Inscricao.Text) = 10 Then Im = True
    FormataRegistro Inscricao
    If Trim(Inscricao) = "" Then Exit Function
    
    Dim Sql As String, rs As VSRecordset
    Sql = "SELECT tci_im, TCI_CGC_CPF,tci_nome, tci_logradouro, tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade, tci_UF " & _
            ",TAE_NOME FROM TAB_CONTRIBUINTE,TAB_ATIVIDADE_ECONOMICA WHERE TCI_TAE_CAE = TAE_CAE"
    If Im Then
        Sql = Sql & " AND TCI_IM='" & Inscricao & "'"
    Else
        Sql = Sql & " AND TCI_CGC_CPF='" & Inscricao & "'"
    End If
    
    If Bdados.AbreTabela(Sql, rs) Then
        If Im Then
            Inscricao = "" & rs!tci_im
        Else
            Inscricao = "" & rs!TCI_CGC_CPF
        End If
        If Not Nome Is Nothing Then Nome = "" & rs!tci_nome
        If Not Endereco Is Nothing Then Endereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & ", " & rs!tci_numero & " " & rs!tci_complemento
        If Not Bairro Is Nothing Then Bairro = "" & rs!tci_bairro
        If Not Cep Is Nothing Then Cep = "" & rs!tci_cep
        If Not Cidade Is Nothing Then Cidade = "" & rs!tci_cidade
        If Not Uf Is Nothing Then Uf = "" & rs!tci_UF
        With Declaracao
            .tciNome = "" & rs!tci_nome
            .tciEndereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & ", " & rs!tci_numero & " " & rs!tci_complemento
            .tciBairro = "" & rs!tci_bairro
            .tciCEP = "" & rs!tci_cep
            .tciCidade = "" & rs!tci_cidade
            .tciUF = "" & rs!tci_UF
            .tciEndereco = .tciEndereco & " " & .tciBairro & " " & .tciCidade & "-" & rs!tci_UF
            .tciAtividade = rs!TAE_NOME
        End With
        BuscarContribuinte = True
    End If
    Bdados.FechaTabela rs
End Function

