VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TMCO101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   6210
      Left            =   60
      TabIndex        =   24
      Top             =   705
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   10954
      Caption         =   "Dados do Contribuinte"
      Descricao       =   "Salva, Exclui e altera informações"
      corFaixa        =   32768
      Icone           =   "TMCO101.frx":0000
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.grdVISUAL grdContribuinte 
         Height          =   3195
         Left            =   105
         TabIndex        =   20
         Top             =   2955
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   5636
         CorBorda        =   32768
         Caption         =   "Contribuintes"
         CorTitulo       =   32768
         CorCaption      =   16777215
         CorDica         =   32768
      End
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   2175
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   3836
         Altura          =   1905
         Caption         =   " Dados do Contribuinte"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtInicio 
            Height          =   285
            Left            =   8850
            TabIndex        =   13
            Top             =   1800
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   503
            Caption         =   "Inicio Atv."
            Text            =   ""
            Formato         =   0
            Restricao       =   2
            AgruparValores  =   0   'False
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.cboVISUAL cboAtividade 
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Tag             =   "Atividade Economica"
            Top             =   1770
            Width           =   8550
            _ExtentX        =   15081
            _ExtentY        =   556
            Caption         =   "Atividade Economica"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   330
            Left            =   10515
            TabIndex        =   18
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtNomeContrib 
            Height          =   285
            Left            =   3405
            TabIndex        =   1
            Tag             =   "Nome/Razão Social"
            Top             =   390
            Width           =   7065
            _ExtentX        =   12462
            _ExtentY        =   503
            Caption         =   "Nome/Razão Social"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtFantasia 
            Height          =   285
            Left            =   3810
            TabIndex        =   3
            Top             =   720
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   503
            Caption         =   "Nome Fantasia"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtNomeLogrContrib 
            Height          =   285
            Left            =   3420
            TabIndex        =   5
            Top             =   1065
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   503
            Caption         =   ""
            Text            =   ""
         End
         Begin VTOcx.cboVISUAL cboUf 
            Height          =   315
            Left            =   9840
            TabIndex        =   11
            Top             =   1425
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            Caption         =   "UF"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtNumero 
            Height          =   285
            Left            =   7020
            TabIndex        =   6
            Top             =   1065
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Caption         =   "Nº"
            Text            =   ""
            Restricao       =   2
         End
         Begin VTOcx.txtVISUAL txtComp 
            Height          =   285
            Left            =   7965
            TabIndex        =   7
            Top             =   1065
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   503
            Caption         =   "Compl."
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtCep 
            Height          =   285
            Left            =   8115
            TabIndex        =   10
            Top             =   1425
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            Caption         =   "CEP"
            Text            =   ""
            Formato         =   4
            Restricao       =   2
            AgruparValores  =   0   'False
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtMunic 
            Height          =   285
            Left            =   3975
            TabIndex        =   9
            Top             =   1425
            Width           =   3990
            _ExtentX        =   7038
            _ExtentY        =   503
            Caption         =   "Município"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtIm 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Tag             =   "Ins. Municipal"
            Top             =   390
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   503
            Caption         =   "Ins. Municipal"
            Text            =   ""
            Formato         =   8
            Restricao       =   2
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCpfCgc 
            Height          =   285
            Left            =   450
            TabIndex        =   2
            Tag             =   "CPF/CNPJ"
            Top             =   720
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   503
            Caption         =   "CPF/CNPJ"
            Text            =   ""
            Restricao       =   2
         End
         Begin VTOcx.cboVISUAL cboTipoLogrContrib 
            Height          =   315
            Left            =   300
            TabIndex        =   4
            Top             =   1065
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   556
            Caption         =   "Logradouro"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtBairro 
            Height          =   285
            Left            =   765
            TabIndex        =   8
            Top             =   1425
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   503
            Caption         =   "Bairro"
            Text            =   ""
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      TabIndex        =   22
      Top             =   6945
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1058
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   7725
         TabIndex        =   16
         Top             =   135
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   8910
         TabIndex        =   14
         Top             =   135
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10095
         TabIndex        =   17
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6570
         TabIndex        =   15
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.TextBox txtFatorFixo 
      Height          =   285
      Left            =   8670
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   1200
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1138
      Icone           =   "TMCO101.frx":08DA
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2745
      TabIndex        =   23
      Top             =   4365
      Width           =   375
   End
End
Attribute VB_Name = "TMCO101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cadastro As VSImposto
Dim eContribuinte As eContribuinte
Private Boletim As TipoBoletim
Dim Contribuinte As VsContribuinte

Private Sub cmdSalvar_Click()
On Error Resume Next
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub

    With eContribuinte
        .Im = txtIm
        .Nome = txtNomeContrib
        .Logradouro = cboTipoLogrContrib
        .NomeLogradouro = txtNomeLogrContrib
        .Numero = txtNumero
        .Complemento = txtComp
        .Bairro = txtBairro
        .Cidade = txtMunic
        .Cep = txtCep
        .Uf = cboUf
        .CgcCpf = txtCpfCgc
        .Fantasia = txtFantasia
        .CodAtividade = cboAtividade.Coluna(1).Valor
        .CodSitCadastral = 1
        .TipoContribuinte = 1
        .InicioAtividade = txtInicio
        If .Salvar = True Then
            Informa "Dados Gravados com sucesso."
            cmdLimpar_Click
        End If
    End With
End Sub

Private Sub cmdEnter_Click()
        SendKeys "{Tab}"
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo trata
    
    If Confirma("Deseja excluir o contribuinte?") Then
        'crítaca pra verificar se existem imoveis
    
        Screen.MousePointer = 11
        If eContribuinte.Excluir(txtIm) Then
            Informa "Registro eliminado com sucesso."
            cmdLimpar_Click
        Else
            Informa "Registro não pode ser eliminado."
        End If
        Screen.MousePointer = 0
    End If
Exit Sub
trata:
    Erro ("erro ao excluir")
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtIm.Enabled = True
    txtCpfCgc.SetFocus
    grdContribuinte.ListItems.Clear
End Sub

Private Sub cmdOpcao_Click()
    If eContribuinte.PreencherGrd(grdContribuinte, txtCpfCgc, txtNomeContrib, 1) = False Then
        Util.Avisa ("Nenhum contribuinte encontrado.")
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Contribuinte = Nothing
    Set eContribuinte = Nothing
End Sub

Private Sub grdContribuinte_dblClick()
    txtIm = grdContribuinte.SelectedItem
    txtIm_LostFocus
End Sub

Private Sub Form_Load()
    
    Set eContribuinte = New eContribuinte
    Set Cadastro = New VSImposto
    Set Contribuinte = New VsContribuinte
    Dim Atividade As New Atividade
    eContribuinte.PreencherComboTipoLogr cboTipoLogrContrib
    cboUf.PreencherGeral Bdados, "UF"
    
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, App.Major, App.Minor, App.Revision
    
    Boletim = tbo_Territorial
    Atividade.PreencherCboAtiv cboAtividade
'    AtualizaCabecalho grdContribuinte
End Sub

Private Sub txtIm_LostFocus()
    If Trim(txtIm) = "" Then Exit Sub
'    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or Me.ActiveControl.ToolTipText = "Sair" Or _
'        Me.ActiveControl.ToolTipText = "Limpar" Then Exit Sub
'    If Len(txtCpfCgc) = 11 Then
'        txtCpfCgc.Formato = formCPF
'    ElseIf Len(txtCpfCgc) = 14 And Mid(txtCpfCgc, 4, 1) <> "." Then
'        txtCpfCgc.MaxLen = 20
'        txtCpfCgc.Formato = formCGC
'    ElseIf Trim(txtCpfCgc) <> "" And Len(txtCpfCgc) <> 18 And Mid(txtCpfCgc, 4, 1) <> "." Then
'        Call Util.Informa("Número de CNPJ ou CPF inválido.")
'        txtCpfCgc.SetFocus
'    Else
'        txtCpfCgc = Edita.TiraPic(txtCpfCgc, ".")
'        txtCpfCgc = Edita.TiraPic(txtCpfCgc, "-")
'    End If
    With eContribuinte
        If .Buscar(txtIm, , False) Then
            txtNomeContrib = .Nome
            cboTipoLogrContrib.SetarLinha .Logradouro, 0
            txtNomeLogrContrib = .NomeLogradouro
            txtNumero = .Numero
            txtComp = .Complemento
            txtBairro = .Bairro
            txtCep = .Cep
            txtMunic = .Cidade
            If Trim(.Uf) <> "" Then cboUf.SetarLinha .Uf, 0
            If Trim(.CodAtividade) <> "" Then cboAtividade.SetarLinha .CodAtividade, 1
            txtCpfCgc = .CgcCpf
            txtFantasia = .Fantasia
            txtInicio = Format(.InicioAtividade, "DD/MM/YYYY")
            txtIm.Enabled = False
        Else
            Dim CPFAux As String
            CPFAux = txtCpfCgc
            Edita.LimpaCampos Me
            txtCpfCgc = CPFAux
        End If
    End With
    txtCpfCgc.Formato = formNenhum
End Sub
