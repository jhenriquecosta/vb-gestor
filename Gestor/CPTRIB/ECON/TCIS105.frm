VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCIS105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TDEC108"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TCIS105.frx":0000
   ScaleHeight     =   2295
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCIS105.frx":0342
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   8
      Top             =   1770
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdFinaliza 
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   90
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         Caption         =   "Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmLimpar 
         Height          =   375
         Left            =   7770
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   690
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
         Left            =   9060
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
      TabIndex        =   7
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1138
      Icone           =   "TCIS105.frx":2465
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1050
      Left            =   30
      TabIndex        =   9
      Top             =   660
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   1852
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtVISUAL1 
         Height          =   285
         Left            =   6960
         TabIndex        =   13
         Top             =   690
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   503
         Caption         =   "Valor Mensal Estimado"
         Text            =   ""
         Formato         =   5
         Restricao       =   2
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboProcedimento 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   690
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   556
         Caption         =   "Procedimento"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
         Height          =   315
         Left            =   3090
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   330
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtData 
         Height          =   285
         Left            =   5100
         TabIndex        =   2
         Top             =   690
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         Caption         =   "Data"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   3450
         TabIndex        =   10
         Top             =   330
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   315
         Left            =   45
         TabIndex        =   0
         Top             =   330
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         Caption         =   "Insc.Municipal"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "mnuGeral"
      Visible         =   0   'False
      Begin VB.Menu mnuDeletar 
         Caption         =   "Deletar"
      End
   End
End
Attribute VB_Name = "TCIS105"
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
Dim String_Taxas As String
Dim Total_Taxas As Double
Dim atividade As New VsTEcon.atividade

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

Private Sub CmdConsultaContribuinte_Click()
    'blnConsultaIM = True
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIM
    'blnConsultaIM = False
End Sub

Private Sub cmdFinaliza_Click()
    Dim NumDec As String
    Dim Controle As Control
    Dim Item As New cItemDeclaracao
        
    If Not Edita.CriticaCampos(Me) Then Exit Sub
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Sql As String
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    'rodVISUAL1.Exibir Bdados, Me.Tag
    Set Imposto = New VsTFuncoes.VSImposto
    DeduzValores = True
    Set Declaracao = New cDeclaracao
'    TabDec.Tabs(4).Enabled = False
    Sql = "SELECT TCD_COD_CAMPO as Item ,TCD_CAMPO as Descricao, ' ' as Valor FROM " & _
        "TAB_CONTEUDO_DECLARACAO WHERE TCD_TMD_COD_DECLARACAO = " & 1
'    ClassGrid.CarregaGrid Grid, Sql
    Sql = "Select tip_cod_imposto,tip_nome_imposto from tab_imposto where tip_sigla_imposto like 'ISS%'"
    cboProcedimento.PreencherGeral Bdados, "PROCEDIMENTO ESTIMATIVA"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Declaracao = Nothing
End Sub

Private Sub txtIM_LostFocus()
    If Trim$(txtIM) <> "" Then
        If Not BuscarContribuinte(txtIM, txtRazao) Then
            Avisa "Contribuinte não encontrado" & vbCrLf & "Verifique se todos os dados estão corretos."
            txtIM = "": txtRazao = ""
            txtIM.SetFocus
    
        End If
    End If
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
    If Im Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        Sql = Sql & " AND TCI_IM='" & Inscricao & "'"
    Else
        Sql = Sql & " AND TCI_CGC_CPF='" & Inscricao & "'"
    End If
    
    If Bdados.AbreTabela(Sql, rs) Then
        If Im Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
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

