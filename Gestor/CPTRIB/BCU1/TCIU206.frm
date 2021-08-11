VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU206 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   6810
      Left            =   60
      TabIndex        =   15
      Top             =   705
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   12012
      Caption         =   "Dados do Imóvel"
      Descricao       =   "Salva, Exclui e altera informações"
      corFaixa        =   16711680
      Icone           =   "TCIU206.frx":0000
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.fraVISUAL fraVISUAL2 
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   2310
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   2566
         Altura          =   1905
         Caption         =   " Imóvel Válido"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtTipoLogrNova 
            Height          =   285
            Left            =   450
            TabIndex        =   26
            Top             =   720
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   503
            Caption         =   "Logradouro"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtBairroNova 
            Height          =   285
            Left            =   885
            TabIndex        =   24
            Top             =   1065
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   503
            Caption         =   "Bairro"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtimNova 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   390
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   503
            Caption         =   "Insc. Cadastral"
            Text            =   ""
            Restricao       =   2
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCepNova 
            Height          =   285
            Left            =   6885
            TabIndex        =   22
            Top             =   1035
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
            TabIndex        =   21
            Top             =   705
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   503
            Caption         =   "Compl."
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtNumeroNova 
            Height          =   285
            Left            =   7020
            TabIndex        =   20
            Top             =   705
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Caption         =   "Nº"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtNomeLogrContribNova 
            Height          =   285
            Left            =   3510
            TabIndex        =   19
            Top             =   705
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   503
            Caption         =   ""
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtNomeContribNova 
            Height          =   285
            Left            =   4275
            TabIndex        =   18
            Top             =   390
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   503
            Caption         =   "Proprietário"
            Text            =   ""
         End
      End
      Begin VTOcx.grdVISUAL grdContribuinte 
         Height          =   2925
         Left            =   105
         TabIndex        =   9
         Top             =   3825
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   5159
         CorBorda        =   16711680
         Caption         =   "Contribuintes"
         CorTitulo       =   16711680
         CorCaption      =   16777215
         CorDica         =   16711680
         OcultarRodape   =   -1  'True
      End
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   1455
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   2566
         Altura          =   1905
         Caption         =   " Imóvel a ser Excluído"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtTipoLogr 
            Height          =   285
            Left            =   360
            TabIndex        =   25
            Top             =   720
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   503
            Caption         =   "Logradouro"
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
            Left            =   4215
            TabIndex        =   1
            Top             =   390
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   503
            Caption         =   "Proprietário"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtNomeLogrContrib 
            Height          =   285
            Left            =   3420
            TabIndex        =   3
            Top             =   735
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   503
            Caption         =   ""
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtNumero 
            Height          =   285
            Left            =   7020
            TabIndex        =   4
            Top             =   735
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Caption         =   "Nº"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtComp 
            Height          =   285
            Left            =   7965
            TabIndex        =   5
            Top             =   735
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   503
            Caption         =   "Compl."
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtCep 
            Height          =   285
            Left            =   6885
            TabIndex        =   7
            Top             =   1065
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
         Begin VTOcx.txtVISUAL txtImAntiga 
            Height          =   285
            Left            =   30
            TabIndex        =   0
            Top             =   390
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   503
            Caption         =   "Insc. Cadastral"
            Text            =   ""
            Restricao       =   2
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtBairro 
            Height          =   285
            Left            =   795
            TabIndex        =   6
            Top             =   1095
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
      TabIndex        =   13
      Top             =   7575
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1058
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   8160
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   10
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
      TabIndex        =   8
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
      TabIndex        =   14
      Top             =   4365
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1138
      Icone           =   "TCIU206.frx":08DA
   End
End
Attribute VB_Name = "TCIU206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As VSImposto
Dim Imovel As New BCI
Private Boletim As TipoBoletim

Dim InscricaoAuxiliar As String
Private Sub cmdSalvar_Click()
On Error Resume Next
    
    If Confirma("Confirma a reunificação dos lançamentos da inscrição nº " & txtImAntiga & " para a inscrição nº " & txtimNova & "?") Then
        Screen.MousePointer = 11
        Bdados.AtualizaDados "TAB_GERACAO_TRIBUTO", Bdados.PreparaValor(Bdados.Converte(Trim(txtimNova), tctexto)), "TGT_INSCRICAO", "TGT_INSCRICAO='" & txtImAntiga & "'" ' AND TGT_TIPO_INSCRICAO = 2"
        Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(Bdados.Converte(Trim(txtimNova), tctexto)), "TOC_INSCRICAO", "TOC_INSCRICAO='" & txtImAntiga & "'" ' AND TOC_TIPO_INSCRICAO = 2"
        Bdados.AtualizaDados "TAB_CONTA_CONTRIBUINTE", Bdados.PreparaValor(Bdados.Converte(Trim(txtimNova), tctexto)), "TCC_INSCRICAO", "TCC_INSCRICAO='" & txtImAntiga & "'" '  AND TCC_TIPO_INSCRICAO = 2"
        Bdados.AtualizaDados "TAB_DARM_RECEBIDO", Bdados.PreparaValor(Bdados.Converte(Trim(txtimNova), tctexto)), "TDR_INSCRICAO", "TDR_INSCRICAO='" & txtImAntiga & "'" '  AND TDR_TIPO_INSCRICAO = 2"
        Bdados.AtualizaDados "TAB_PARCELAMENTO", Bdados.PreparaValor(Bdados.Converte(Trim(txtimNova), tctexto)), "TPA_INSCRICAO", "TPA_INSCRICAO ='" & txtImAntiga & "'" '  AND TPA_TIPO_INSCRICAO = 2"
        Bdados.AtualizaDados "TAB_COTAS_PARCELAMENTO", Bdados.PreparaValor(Bdados.Converte(Trim(txtimNova), tctexto)), "TCP_INSCRICAO", "TCP_INSCRICAO ='" & txtImAntiga & "'"
        
        Bdados.DeletaDados "TAB_CONTA_CONTRIBUINTE", "TCC_INSCRICAO ='" & txtImAntiga & "'"
        Bdados.DeletaDados "TAB_DETALHE_IMOVEL", "TDI_TIM_IC= '" & txtImAntiga & "'"
        Bdados.DeletaDados "TAB_IMOVEL", "TIM_IC= '" & txtImAntiga & "'"
        Screen.MousePointer = 0
        Avisa "Reunificação concluída."
        cmdLimpar_Click
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
    Dim Sql As String
    Sql = "select tim_ic as  [Cad Imobiliária],tim_ic_auxiliar as [Insc Imobiliária],"
    Sql = Sql & " tim_tci_im as [Cad Contribuinte],tci_nome as Contribuinte,TTL_NOME as Logradouro ,"
    Sql = Sql & " tlg_nome as Endereco,TBA_NOME as Bairro,tim_numero as Número "
    Sql = Sql & " From vis_imovel where 1 = 1"
    If Trim(txtImAntiga) <> "" Then
        Sql = Sql & " and tim_ic = '" & Trim(txtImAntiga) & "'"
    End If
    If Trim(txtNomeContrib) <> "" Then
        Sql = Sql & " and tci_nome like '%" & Trim(txtNomeContrib) & "%'"
    End If
    If Not grdContribuinte.Preencher(Bdados, Sql) Then
        Avisa "Nenhum registro encontrado."
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub grdContribuinte_dblClick()
    txtImAntiga = grdContribuinte.SelectedItem
    txtImAntiga_LostFocus
End Sub

Private Sub Form_Load()
    
    Set cadastro = New VSImposto
    
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, App.Major, App.Minor, App.Revision
    
    Boletim = tbo_Territorial
    AtualizaCabecalho grdContribuinte
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
        txtImAntiga.Formato = formNenhum
    End If
End Sub

Private Sub txtImAntiga_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtImAntiga) <> "" Then
        With Imovel
            Sql = "Select TTL_NOME,tlg_nome,tim_complemento,TBA_NOME,tci_nome,tim_numero,tim_cep from vis_imovel where TIM_IC ='" & txtImAntiga & "'"
            If Bdados.AbreTabela(Sql, Rs) Then
                txtTipoLogr = "" & Rs!TTL_NOME
                txtNomeLogrContrib = "" & Rs!tlg_nome
                txtComp = "" & Rs!tim_complemento
                txtBairro = "" & Rs!TBA_NOME
                txtNomeContrib = "" & Rs!tci_nome
                txtNumero = "" & Rs!tim_numero
                txtCep = "" & Rs!tim_cep
            Else
               Avisa "Imóvel não cadastrado."
            End If
        End With
    Else
        txtImAntiga.Enabled = False
    End If
End Sub

Private Sub txtimNova_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtimNova) <> "" Then
        With Imovel
            Sql = "Select TTL_NOME,tlg_nome,tim_complemento,TBA_NOME,tci_nome,tim_numero,tim_cep from vis_imovel where TIM_IC ='" & txtimNova & "'"
            If Bdados.AbreTabela(Sql, Rs) Then
                txtTipoLogrNova = "" & Rs!TTL_NOME
                txtNomeLogrContribNova = "" & Rs!tlg_nome
                txtCompNova = "" & Rs!tim_complemento
                txtBairroNova = "" & Rs!TBA_NOME
                txtNomeContribNova = "" & Rs!tci_nome
                txtNumeroNova = "" & Rs!tim_numero
                txtCepNova = "" & Rs!tim_cep
                txtimNova.Enabled = False
            Else
               Avisa "Imóvel não cadastrado."
            End If
        End With
    Else
        txtImAntiga.Enabled = False
    End If
End Sub
