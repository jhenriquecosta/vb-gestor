VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TDEC106 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DECL101"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TDEC106.frx":0000
   ScaleHeight     =   4860
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   5
      Top             =   4335
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmLimpar 
         Height          =   375
         Left            =   7140
         TabIndex        =   3
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   5100
         TabIndex        =   2
         Top             =   90
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   661
         Caption         =   "&Aceitar Declaracão"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8430
         TabIndex        =   4
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1080
      Left            =   30
      TabIndex        =   6
      Top             =   660
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   1905
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
         Height          =   315
         Left            =   3030
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
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   315
         Left            =   120
         TabIndex        =   0
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
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   285
         Left            =   300
         TabIndex        =   1
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
         Left            =   3390
         TabIndex        =   7
         Top             =   330
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
      End
   End
   Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
      Height          =   3690
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   6509
      _Version        =   131082
      TabGuid         =   "TDEC106.frx":0342
      Begin VB.TextBox txtGrig 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   9810
         TabIndex        =   9
         Text            =   "txtGrig"
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblTexto 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"TDEC106.frx":036A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   2025
         Left            =   0
         TabIndex        =   10
         Top             =   1320
         Width           =   9585
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   1138
      Icone           =   "TDEC106.frx":0475
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "mnuGeral"
      Visible         =   0   'False
      Begin VB.Menu mnuDeletar 
         Caption         =   "Deletar"
      End
   End
End
Attribute VB_Name = "TDEC106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIM
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Not Confirma("Confirma a apresentacão da declaracão negativa de movimentos " & _
                "para o contribuinte " & txtRazao & " , Inscricão " & txtIM & _
                ", no período de " & txtPeriodo & "?") Then
        Exit Sub
    End If
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If txtIM = "" Then Exit Sub
    
    
    Declaracao.Im = txtIM
    Declaracao.Periodo = txtPeriodo
    Declaracao.Data = Format(Date, "dd/mm/yyyy")
    Declaracao.Origem = orgSistema
    Declaracao.Recepcao = Date
    Declaracao.Versao = Nvl(Temp.PegaParametro(Bdados, "VERSAO DEC"), 2)
    Declaracao.Tipo = decNegativa
    Declaracao.BaseGeral = 0
    Declaracao.Status = decFinalizada
    
    If Declaracao.Gravar() Then
        Declaracao.Finalizar False, , , decNegativa, 0
        Avisa "Declaração gravada com sucesso."
        txtIM.SetFocus
    End If
End Sub

Private Sub cmLimpar_Click()
    Edita.LimpaCampos Me
    txtIM.SetFocus
End Sub

Private Sub Form_Load()
    
    'cabVISUAL1.Bdados , Me.Tag, App.Path
    'rodVISUAL1.Exibir Bdados, Me.Tag
    Set Imposto = New VsTFuncoes.VSImposto
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Declaracao = Nothing
End Sub

Private Sub txtIM_LostFocus()
    
    If Trim$(txtIM) <> "" Then
        If Not BuscarContribuinte(txtIM, txtRazao) Then
            Avisa "Contribuinte não encontrado."
            txtIM = "": txtRazao = ""
            txtIM.SetFocus
        End If
    End If
End Sub


Private Sub txtPeriodo_Change()
        
    If Len(Trim(txtPeriodo)) <> 7 Then Exit Sub
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
'    If Len(Inscricao.Text) = 10 Then Im = True
    FormataRegistro Inscricao
    If Trim(Inscricao) = "" Then Exit Function
    
    Dim Sql As String, rs As VSRecordset
    Sql = "SELECT tci_im, TCI_CGC_CPF,tci_nome, tci_logradouro, tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade, tci_UF " & _
            ",TAE_NOME FROM TAB_CONTRIBUINTE LEFT JOIN TAB_ATIVIDADE_ECONOMICA ON TCI_TAE_CAE = TAE_CAE WHERE 1=1"
'    If Im Then
        Sql = Sql & " AND TCI_IM='" & Inscricao & "'"
'    Else
'        Sql = Sql & " AND TCI_CGC_CPF='" & Inscricao & "'"
'    End If
    
    If Bdados.AbreTabela(Sql, rs) Then
'        If Im Then
            Inscricao = "" & rs!tci_im
'        Else
'            Inscricao = "" & rs!TCI_CGC_CPF
'        End If
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
            .tciAtividade = "" & rs!TAE_NOME
        End With
        BuscarContribuinte = True
    End If
    Bdados.FechaTabela rs
End Function

