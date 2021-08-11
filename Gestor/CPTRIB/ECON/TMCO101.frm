VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TMCO101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   6990
      Left            =   60
      TabIndex        =   23
      Top             =   705
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   12330
      Caption         =   "Dados do Contribuinte"
      Descricao       =   "Salva, Exclui e altera informações"
      corFaixa        =   16711680
      Icone           =   "TMCO101.frx":0000
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.grdVISUAL grdContribuinte 
         Height          =   3615
         Left            =   105
         TabIndex        =   16
         Top             =   3195
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   6376
         CorBorda        =   16711680
         Caption         =   "Contribuintes"
         CorTitulo       =   16711680
         CorCaption      =   16777215
         CorDica         =   16711680
      End
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   2385
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   4207
         Altura          =   1905
         Caption         =   " Dados do Proprietário"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtImAnterior 
            Height          =   285
            Left            =   8160
            TabIndex        =   25
            Top             =   1740
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   503
            Caption         =   "Ins. Anterior"
            Text            =   ""
            Formato         =   8
            Restricao       =   2
            AgruparValores  =   0   'False
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtEmail 
            Height          =   285
            Left            =   3390
            TabIndex        =   14
            Top             =   1740
            Width           =   4600
            _ExtentX        =   8123
            _ExtentY        =   503
            Caption         =   "e-mail"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL TXTTelefone 
            Height          =   285
            Left            =   540
            TabIndex        =   13
            Top             =   1740
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   503
            Caption         =   "Telefone"
            Text            =   ""
         End
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   330
            Left            =   10515
            TabIndex        =   2
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
            TabIndex        =   4
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
            TabIndex        =   6
            Top             =   1065
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   503
            Caption         =   ""
            Text            =   ""
         End
         Begin VTOcx.cboVISUAL cboUf 
            Height          =   315
            Left            =   9960
            TabIndex        =   12
            Top             =   1395
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            Caption         =   "UF"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtNumero 
            Height          =   285
            Left            =   7020
            TabIndex        =   7
            Top             =   1065
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Caption         =   "Nº"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtComp 
            Height          =   285
            Left            =   7965
            TabIndex        =   8
            Top             =   1065
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   503
            Caption         =   "Compl."
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtCep 
            Height          =   285
            Left            =   8535
            TabIndex        =   11
            Top             =   1395
            Width           =   1425
            _ExtentX        =   2514
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
            Left            =   4575
            TabIndex        =   10
            Top             =   1425
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   503
            Caption         =   "Município"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtIm 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   390
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   503
            Caption         =   "Ins. Municipal"
            Text            =   ""
            Formato         =   8
            Restricao       =   2
            AgruparValores  =   0   'False
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCpfCgc 
            Height          =   285
            Left            =   450
            TabIndex        =   3
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
            TabIndex        =   5
            Top             =   1065
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   556
            Caption         =   "Logradouro"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Editavel        =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtBairro 
            Height          =   285
            Left            =   765
            TabIndex        =   9
            Top             =   1425
            Width           =   3645
            _ExtentX        =   6429
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
      TabIndex        =   21
      Top             =   7575
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1058
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   8925
         TabIndex        =   18
         Top             =   135
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   8388608
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   6600
         TabIndex        =   19
         Top             =   135
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8388608
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10095
         TabIndex        =   20
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8388608
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7770
         TabIndex        =   17
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8388608
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VB.TextBox txtFatorFixo 
      Height          =   285
      Left            =   8670
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2745
      TabIndex        =   22
      Top             =   4365
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1138
      Icone           =   "TMCO101.frx":08DA
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
Dim InscricaoAuxiliar As String
Private Sub cmdSalvar_Click()
On Error Resume Next
    'GLEYSON - BCP - CORRIGIR MIGRACAO
    'nao use constantemente, o nome ja diz
  '  criarInscricao
  '  Exit Sub
    
    Dim InscricaoMunicipal As String
    
    Dim Conta As New ContaCorrente
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If Trim(txtIm) = "" And txtIm.Enabled = False Then
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
            InscricaoMunicipal = Conta.GeraCodPagamento("CADASTRO ECONOMICO")
            InscricaoAuxiliar = Cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
        Else
            InscricaoMunicipal = Cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
            InscricaoAuxiliar = ""
    End If
    Else
        InscricaoMunicipal = txtIm
    End If
    
    With eContribuinte
        .Im = InscricaoMunicipal
        .Obs = ""
        .CodAtividade = 7499300
        .AreaEstabelecimento = 0
        If Trim(InscricaoAuxiliar) <> "" Then .ImAuxiliar = InscricaoAuxiliar
        .Nome = txtNomeContrib
        .Logradouro = cboTipoLogrContrib
        .NomeLogradouro = txtNomeLogrContrib
        .Numero = txtNumero
        .Complemento = txtComp
        .Bairro = txtBairro
        .Cidade = txtMunic
        .Cep = txtCep
        .DataCadastro = Bdados.Converte(Date, TCDataHora)
        .DataModificacao = Bdados.Converte(Date, TCDataHora)
        .InicioAtividade = Bdados.Converte(Date, TCDataHora)
        .Uf = cboUf
        .CgcCpf = txtCpfCgc
        .Telefone = TXTTelefone
        .Email = txtEmail
        .Fantasia = txtFantasia
        .CodSitCadastral = 1
        If txtImAnterior = "" Then
            .ImAuxiliar = 0
            .ImAnterior = 0
        Else
            .ImAuxiliar = txtImAnterior
            .ImAnterior = txtImAnterior
        End If
        .Rg = 0
            .ImAuxiliar = 0
            'BCP
            .ImAnterior = 0
            .AreaEstabelecimento = 0
            .Obs = ""
           ' .Crc = txtCrc
            'FIM BCP
            .TipoCadastro = 1
            If txtFantasia = "" Then
                .Fantasia = txtNomeContrib
            Else
                .Fantasia = txtFantasia
            End If
            .InicioPrestacaoServico = Bdados.Converte(Date, TCDataHora)
            .SituacaoAlvara = 1
            .Matriz_Filial = 1
            .VariavelAnuncio = 0
            .CodGrupo = 1
            .CodSitCadastral = 1
            .CodNatureza = 1
            .CodAtivPoder = 1
            .Estabelecido = 1
            .DataCadastro = Bdados.Converte(Date, TCDataHora)
            .Nome_Tela = Me.Caption
            .CodUsuario = AplicacoesVTFuncoes.Usuario
            .GrupoAtividade = 1
            .InicioAtividade = Bdados.Converte(Date, TCDataHora)
            .TipoContribuinte = 1
            .TipoRecolhimentoIss = 1
            .Ruc = 0
            .FatorAlvara = 1
            .Conselho = 1
            .Registro = 1
            .ImovelProprio = 0
            .NumEmpregado = 0
            .PorteEmpresa = 1
            .CodAtividadeSec = 7499300
            .CodAtividadeTerc = 7499300
            .NivelEscolar = 1
            .Protocolo = 0
            .CodRamo = 0
            .CNH = 0
            .Categoria = 0
            .Autorizacao = 0
            .PontoRecepcao = 0
            'Salva dados do contribuinte
        If .SalvarReduzido = True Then
            Informa "Dados Gravados com sucesso."
            cmdLimpar_Click
        End If
    End With
    Call Util.Informa("Registro gravado com sucesso. Inscricão Municipal Gerada Nº: " & InscricaoMunicipal & ".")
End Sub

Private Sub cmdEnter_Click()
        SendKeys "{Tab}"
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo trata
    If Trim(txtIm) = "" Then
        Informa "Informe o contribuinte."
        txtIm.Enabled = True
        txtIm.SetFocus
        Exit Sub
    End If
    If Confirma("Deseja excluir o contribuinte?") Then
        'crítaca pra verificar se existem imoveis
        If eContribuinte.VerificaTEMImovel(txtIm) = True Then
            Call Util.Informa("Contribuinte possui imóveis.")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Screen.MousePointer = 11
        'crítica pra verificar se existe debito
        If eContribuinte.VerificaTEMDebito(txtIm) = True Then
            Call Util.Informa("Contribuinte possui lançamento de tributos.")
            Screen.MousePointer = 0
            Exit Sub
        End If
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
    txtIm.SetFocus
    txtCep = Temp.PegaParametro(Bdados, "CEP MUNICIPIO")
    grdContribuinte.ListItems.Clear
End Sub

Private Sub cmdOpcao_Click()
    If eContribuinte.PreencherGrd(grdContribuinte, txtIm, txtNomeContrib, , txtCpfCgc.Text) = False Then
        Util.Avisa ("Nenhum contribuinte encontrado.")
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
'   criarInscricao

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Contribuinte = Nothing
    Set eContribuinte = Nothing
End Sub

Private Sub grdContribuinte_dblClick()
    txtIm = grdContribuinte.SelectedItem
    txtIM_LostFocus
End Sub

Private Sub Form_Load()
    
    Set eContribuinte = New eContribuinte
    Set Cadastro = New VSImposto
    Set Contribuinte = New VsContribuinte
    
    eContribuinte.PreencherComboTipoLogr cboTipoLogrContrib
    cboUf.PreencherGeral Bdados, "UF"
    
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, App.Major, App.Minor, App.Revision
    
    Boletim = tbo_Territorial
    AtualizaCabecalho grdContribuinte
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
        txtIm.Formato = formNenhum
    End If
End Sub

Private Sub txtCpfCgc_LostFocus()
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or Me.ActiveControl.ToolTipText = "Sair" Or _
        Me.ActiveControl.ToolTipText = "Limpar" Then Exit Sub
    If Len(txtCpfCgc) = 11 Then
        txtCpfCgc.Formato = formCPF
    ElseIf Len(txtCpfCgc) = 14 And Mid(txtCpfCgc, 4, 1) <> "." Then
        txtCpfCgc.MaxLen = 20
        txtCpfCgc.Formato = formCGC
    ElseIf Trim(txtCpfCgc) <> "" And Len(txtCpfCgc) <> 18 And Mid(txtCpfCgc, 4, 1) <> "." Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
        txtCpfCgc.SetFocus
    Else
        txtCpfCgc = Edita.TiraPic(txtCpfCgc, ".")
        txtCpfCgc = Edita.TiraPic(txtCpfCgc, "-")
    End If
    If Cadastro.VerificaEmpresaAntiga(txtCpfCgc, txtNomeContrib) = 1 Then
        If Not Util.Confirma("Já existe uma empresa cadastrada com o mesmo CNPJ/CPF. Confirma cadastro.") Then
            txtCpfCgc.SetFocus
            Exit Sub
        End If
    End If
    txtCpfCgc.Formato = formNenhum
End Sub

Private Sub txtIM_LostFocus()
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIm) <> "" Then
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            If InStr(1, txtIm, "-") = 0 Then txtIm = Cadastro.FormataInscricao(txtIm, InscContrib)
        Else
            txtIm.Formato = formNenhum
        End If
        With eContribuinte
        
            If .Buscar(txtIm, , False) Then
                txtNomeContrib = .Nome
                cboTipoLogrContrib.SetarLinha .Logradouro, 0
                txtNomeLogrContrib = .NomeLogradouro
                txtNumero = .Numero
                txtComp = .Complemento
                txtBairro = .Bairro
                InscricaoAuxiliar = .ImAuxiliar
                TXTTelefone = .Telefone
                txtEmail = .Email
                'txtImAnterior = .ImAuxiliar
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
                    txtCep = Temp.PegaParametro(Bdados, "CEP MUNICIPIO")
                End If
                txtCep = .Cep
                txtMunic = .Cidade
                If Trim(.Uf) <> "" Then cboUf.SetarLinha .Uf, 0
                txtCpfCgc = .CgcCpf
                txtFantasia = .Fantasia
                txtIm.Enabled = True
                txtCpfCgc_LostFocus
            Else
                If Not Util.Confirma("Contribuinte não cadastrado. Deseja cadastrá-lo?") Then
                    txtIm = ""
                    txtNomeContrib.SetFocus
                Else
                    
                End If
            End If
        End With
    Else
        'txtIm = cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
        txtIm.Enabled = False
    End If
End Sub
Private Sub criarInscricao()
    Dim Rs As VSRecordset, rsUp As VSRecordset
    Dim Conta As New ContaCorrente
    
    Dim novaIm As String
    If Bdados.AbreTabela("SELECT TCI_IM FROM TAB_CONTRIBUINTE", Rs) Then
       Do While Rs.EOF = False
            novaIm = Cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
            txtIm.Text = novaIm
            
            Bdados.Executa ("UPDATE TAB_CONTRIBUINTE SET TCI_IM='" & novaIm & "' WHERE TCI_IM='" & Rs(0) & "'")
            Rs.MoveNext
            DoEvents
        Loop
    End If
    Util.Informa "Fim..."
End Sub

