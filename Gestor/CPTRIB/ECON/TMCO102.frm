VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TMCO102 
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
      Height          =   6810
      Left            =   60
      TabIndex        =   20
      Top             =   705
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   12012
      Caption         =   "Dados do Contribuinte"
      Descricao       =   "Salva, Exclui e altera informações"
      corFaixa        =   16711680
      Icone           =   "TMCO102.frx":0000
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.fraVISUAL fraVISUAL2 
         Height          =   1785
         Left            =   120
         TabIndex        =   22
         Top             =   2550
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   3149
         Altura          =   1905
         Caption         =   " Contribuinte Válido"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtBairroNova 
            Height          =   285
            Left            =   765
            TabIndex        =   34
            Top             =   1425
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   503
            Caption         =   "Bairro"
            Text            =   ""
         End
         Begin VTOcx.cboVISUAL cboTipoLogrContribNova 
            Height          =   315
            Left            =   300
            TabIndex        =   33
            Top             =   1065
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   556
            Caption         =   "Logradouro"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Editavel        =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtCpfCgcNova 
            Height          =   285
            Left            =   450
            TabIndex        =   32
            Top             =   720
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   503
            Caption         =   "CPF/CNPJ"
            Text            =   ""
            Restricao       =   2
         End
         Begin VTOcx.txtVISUAL txtimNova 
            Height          =   285
            Left            =   120
            TabIndex        =   31
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
         Begin VTOcx.txtVISUAL txtMunicNova 
            Height          =   285
            Left            =   4575
            TabIndex        =   30
            Top             =   1425
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   503
            Caption         =   "Município"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtCepNova 
            Height          =   285
            Left            =   8535
            TabIndex        =   29
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
         Begin VTOcx.txtVISUAL txtCompNova 
            Height          =   285
            Left            =   7965
            TabIndex        =   28
            Top             =   1065
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   503
            Caption         =   "Compl."
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtNumeroNova 
            Height          =   285
            Left            =   7020
            TabIndex        =   27
            Top             =   1065
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Caption         =   "Nº"
            Text            =   ""
         End
         Begin VTOcx.cboVISUAL cboUfNova 
            Height          =   315
            Left            =   9960
            TabIndex        =   26
            Top             =   1395
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            Caption         =   "UF"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtNomeLogrContribNova 
            Height          =   285
            Left            =   3420
            TabIndex        =   25
            Top             =   1065
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   503
            Caption         =   ""
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtFantasiaNova 
            Height          =   285
            Left            =   3810
            TabIndex        =   24
            Top             =   720
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   503
            Caption         =   "Nome Fantasia"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtNomeContribNova 
            Height          =   285
            Left            =   3405
            TabIndex        =   23
            Top             =   390
            Width           =   7485
            _ExtentX        =   13203
            _ExtentY        =   503
            Caption         =   "Nome/Razão Social"
            Text            =   ""
         End
      End
      Begin VTOcx.grdVISUAL grdContribuinte 
         Height          =   2535
         Left            =   105
         TabIndex        =   14
         Top             =   4395
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   4471
         CorBorda        =   16711680
         Caption         =   "Contribuintes"
         CorTitulo       =   16711680
         CorCaption      =   16777215
         CorDica         =   16711680
         OcultarRodape   =   -1  'True
      End
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   1785
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   3149
         Altura          =   1905
         Caption         =   " Contribuinte a ser Excluído"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
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
         Begin VTOcx.txtVISUAL txtImAntiga 
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
      TabIndex        =   18
      Top             =   7575
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1058
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   8160
         TabIndex        =   16
         Top             =   135
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         Caption         =   "&Unificar registros"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
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
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6960
         TabIndex        =   15
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VB.TextBox txtFatorFixo 
      Height          =   285
      Left            =   8670
      TabIndex        =   13
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
      TabIndex        =   19
      Top             =   4365
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1138
      Icone           =   "TMCO102.frx":08DA
   End
End
Attribute VB_Name = "TMCO102"
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
    
    If Confirma("Confirma a reuinificação dos lançamentos da inscrição nº " & txtImAntiga & " para a inscrição nº " & txtimNova & "?") Then
        Screen.MousePointer = 11
        Bdados.AtualizaDados "TAB_GERACAO_TRIBUTO", Bdados.PreparaValor(Bdados.Converte(txtimNova, tctexto)), "TGT_INSCRICAO", "TGT_INSCRICAO='" & txtImAntiga & "' AND TGT_TIPO_INSCRICAO = 2"
        Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(Bdados.Converte(txtimNova, tctexto)), "TOC_INSCRICAO", "TOC_INSCRICAO='" & txtImAntiga & "' AND TOC_TIPO_INSCRICAO = 2"
        Bdados.AtualizaDados "TAB_CONTA_CONTRIBUINTE", Bdados.PreparaValor(Bdados.Converte(txtimNova, tctexto)), "TCC_INSCRICAO", "TCC_INSCRICAO='" & txtImAntiga & "'  AND TCC_TIPO_INSCRICAO = 2"
        Bdados.AtualizaDados "TAB_DARM_RECEBIDO", Bdados.PreparaValor(Bdados.Converte(txtimNova, tctexto)), "TDR_INSCRICAO", "TDR_INSCRICAO='" & txtImAntiga & "'  AND TDR_TIPO_INSCRICAO = 2"
        Bdados.AtualizaDados "TAB_PARCELAMENTO", Bdados.PreparaValor(Bdados.Converte(txtimNova, tctexto)), "TPA_INSCRICAO", "TPA_INSCRICAO ='" & txtImAntiga & "'  AND TPA_TIPO_INSCRICAO = 2"
        Bdados.AtualizaDados "TAB_COTAS_PARCELAMENTO", Bdados.PreparaValor(Bdados.Converte(txtimNova, tctexto)), "TCP_INSCRICAO", "TCP_INSCRICAO ='" & txtImAntiga & "'"
        Bdados.AtualizaDados "TAB_IMOVEL", Bdados.PreparaValor(Bdados.Converte(txtimNova, tctexto)), "TIM_TCI_IM", "TIM_TCI_IM ='" & txtImAntiga & "'"
        Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "TCC_INSCRICAO ='" & txtImAntiga & "'"
        Bdados.DeletaDados "TAB_CONTRIBUINTE", "TCI_IM = '" & txtImAntiga & "'"
        Screen.MousePointer = 0
        Avisa "Reunificação concluída."
    End If
End Sub

Private Sub cmdEnter_Click()
        SendKeys "{Tab}"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtImAntiga.Enabled = True
    txtImAntiga.SetFocus
    grdContribuinte.ListItems.Clear
    txtimNova.Enabled = True
End Sub

Private Sub cmdOpcao_Click()
    If eContribuinte.PreencherGrd(grdContribuinte, txtImAntiga, txtNomeContrib, , txtCpfCgc.Text) = False Then
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
    txtImAntiga = grdContribuinte.SelectedItem
    txtImAntiga_LostFocus
End Sub

Private Sub Form_Load()
    
    Set eContribuinte = New eContribuinte
    Set Cadastro = New VSImposto
    Set Contribuinte = New VsContribuinte
    
    eContribuinte.PreencherComboTipoLogr cboTipoLogrContrib
    cboUf.PreencherGeral Bdados, "UF"
    
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, App.Major, App.Minor, App.Revision
    
    Boletim = tbo_Territorial
    AtualizaCabecalho grdContribuinte
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
        txtImAntiga.Formato = formNenhum
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
    txtCpfCgc.Formato = formNenhum
End Sub



Private Sub txtCpfCgcNova_LostFocus()
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or Me.ActiveControl.ToolTipText = "Sair" Or _
        Me.ActiveControl.ToolTipText = "Limpar" Then Exit Sub
    If Len(txtCpfCgcNova) = 11 Then
        txtCpfCgcNova.Formato = formCPF
    ElseIf Len(txtCpfCgcNova) = 14 And Mid(txtCpfCgcNova, 4, 1) <> "." Then
        txtCpfCgcNova.MaxLen = 20
        txtCpfCgcNova.Formato = formCGC
    ElseIf Trim(txtCpfCgcNova) <> "" And Len(txtCpfCgcNova) <> 18 And Mid(txtCpfCgcNova, 4, 1) <> "." Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
        txtCpfCgcNova.SetFocus
    Else
        txtCpfCgcNova = Edita.TiraPic(txtCpfCgcNova, ".")
        txtCpfCgcNova = Edita.TiraPic(txtCpfCgcNova, "-")
    End If
    txtCpfCgcNova.Formato = formNenhum
End Sub


Private Sub txtImAntiga_LostFocus()
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtImAntiga) <> "" Then
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            txtImAntiga = Cadastro.FormataInscricao(txtImAntiga, InscContrib)
        Else
            txtImAntiga.Formato = formNenhum
        End If
        With eContribuinte
        
            If .Buscar(txtImAntiga, , False) Then
                txtNomeContrib = .Nome
                cboTipoLogrContrib.SetarLinha .Logradouro, 0
                txtNomeLogrContrib = .NomeLogradouro
                txtNumero = .Numero
                txtComp = .Complemento
                txtBairro = .Bairro
                InscricaoAuxiliar = .ImAuxiliar
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
                    txtCep = Temp.PegaParametro(Bdados, "CEP MUNICIPIO")
                End If
                txtCep = .Cep
                txtMunic = .Cidade
                If Trim(.Uf) <> "" Then cboUf.SetarLinha .Uf, 0
                txtCpfCgc = .CgcCpf
                txtFantasia = .Fantasia
                txtImAntiga.Enabled = True
                txtCpfCgc_LostFocus
            Else
                If Not Util.Confirma("Contribuinte não cadastrado. Deseja cadastrá-lo?") Then
                    txtImAntiga = ""
                    txtNomeContrib.SetFocus
                Else
                    
                End If
            End If
        End With
    Else
        'txtImAntiga = cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
        txtImAntiga.Enabled = False
    End If
End Sub

Private Sub txtimNova_LostFocus()
    If Trim(txtimNova) = "" Then Exit Sub
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtimNova) <> "" Then
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            txtimNova = Cadastro.FormataInscricao(txtimNova, InscContrib)
        Else
            txtimNova.Formato = formNenhum
        End If
        With eContribuinte
        
            If .Buscar(txtimNova, , False) Then
                txtNomeContribNova = .Nome
                cboTipoLogrContribNova.SetarLinha .Logradouro, 0
                txtNomeLogrContribNova = .NomeLogradouro
                txtNumeroNova = .Numero
                txtCompNova = .Complemento
                txtBairroNova = .Bairro
                InscricaoAuxiliar = .ImAuxiliar
                If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
                    txtCepNova = Temp.PegaParametro(Bdados, "CEP MUNICIPIO")
                End If
                txtCepNova = .Cep
                txtMunicNova = .Cidade
                If Trim(.Uf) <> "" Then cboUfNova.SetarLinha .Uf, 0
                txtCpfCgcNova = .CgcCpf
                txtFantasiaNova = .Fantasia
                txtimNova.Enabled = True
                txtCpfCgcNova_LostFocus
            Else
                If Not Util.Confirma("Contribuinte não cadastrado. Deseja cadastrá-lo?") Then
                    txtimNova = ""
                    txtNomeContrib.SetFocus
                Else
                    
                End If
            End If
        End With
    Else
        'txtimNova = cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
        txtimNova.Enabled = False
    End If
End Sub
