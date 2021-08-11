VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TATV401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TATV401.frx":0000
   ScaleHeight     =   7095
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   5850
      Left            =   45
      TabIndex        =   5
      Top             =   690
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   10319
      Caption         =   "Atividades Econômicas"
      Descricao       =   "Consulta de atividades econômicas, informações gerais"
      corFaixa        =   16711680
      Icone           =   "TATV401.frx":0342
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.grdVISUAL grdAtividade 
         Height          =   3510
         Left            =   90
         TabIndex        =   11
         Top             =   2205
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   6191
         CorBorda        =   16711680
         Caption         =   "Atividades"
         CorTitulo       =   16711680
         CorCaption      =   16777215
         CorDica         =   16711680
      End
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   1440
         Left            =   105
         TabIndex        =   6
         Top             =   720
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   2540
         Altura          =   1905
         Caption         =   " Opções de Filtro"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtDescAtiv 
            Height          =   285
            Left            =   150
            TabIndex        =   10
            Top             =   690
            Width           =   8580
            _ExtentX        =   15134
            _ExtentY        =   503
            Caption         =   "Nome Atividade"
            Text            =   ""
            Restricao       =   1
            ValorMaximo     =   100
            MaxLen          =   50
            MinLen          =   1
         End
         Begin VTOcx.cboVISUAL cboGrupoAtiv 
            Height          =   315
            Left            =   90
            TabIndex        =   9
            Top             =   1035
            Width           =   5625
            _ExtentX        =   9922
            _ExtentY        =   556
            Caption         =   "Grupo Atividade"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCodAtiv 
            Height          =   285
            Left            =   885
            TabIndex        =   8
            Top             =   345
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   503
            Caption         =   "Código"
            Text            =   ""
            Restricao       =   2
         End
         Begin VTOcx.cboVISUAL cboEstimativa 
            Height          =   315
            Left            =   6360
            TabIndex        =   7
            Top             =   1035
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            Caption         =   "Estimado"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   4
      Top             =   6570
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   5460
         TabIndex        =   1
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdBusca 
         Height          =   375
         Left            =   4290
         TabIndex        =   0
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdImprime 
         Height          =   375
         Left            =   6615
         TabIndex        =   2
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7770
         TabIndex        =   3
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Sair"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1138
      Icone           =   "TATV401.frx":065C
   End
End
Attribute VB_Name = "TATV401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim atividade As atividade
Public CodAtividade As Long


'Private Sub cboGrupoAtiv_Click()
'    cboAtividade.Visible = IIf(cboGrupoAtiv.ListIndex = 4, True, False)
'End Sub

Private Sub cmdBusca_Click()
'    Dim Sql As String
'    Dim RsPref As VSRecordset
'    Dim RsCTM As VSRecordset
'    Dim Anterior As String
'    Sql = "Select tae_cae as Código, tae_nome as Atividade,tga_nome as Grupo," & Bdados.Converte("tae_valor", TCDuplo) & " as [Valor(R$)]," & _
'        "tae_desc_fator as Fator from Tab_Atividade_Economica, Tab_Grupo_Atividade where " & _
'        " tae_tga_cod_grupo = tga_cod_grupo"
'    If Trim(cboGrupoAtiv) <> "" Then Sql = Sql & " and tga_nome='" & cboGrupoAtiv & "'"
'    If Trim(txtDescAtiv) <> "" Then Sql = Sql & " and (tae_nome like '" & txtDescAtiv & "%' or tae_nome like '% " & txtDescAtiv & "%')"
'    grdAtividade.Preencher Bdados, Sql, 1400
atividade.PreencheGrid grdAtividade, CStr(cboGrupoAtiv.Coluna(1).Valor), txtCodAtiv, CStr(cboEstimativa.Coluna(1).Valor), txtDescAtiv
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtCodAtiv.SetFocus
    grdAtividade.ListItems.Clear
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdImprime_Click()
    Dim Aux As Byte
    Dim Formula As String
    Dim Paginas As Integer
    Dim SelecaoRpt As String
    With Rpt
            If Not .DefinirArquivo(Bdados, App.Path + "\TAtividades.rpt") Then Exit Sub
            '.Connect = Bdados.BDSistema.Connect
            If Trim(cboGrupoAtiv) <> "" Then
                SelecaoRpt = "{Tab_Grupo_Atividade.tga_nome} ='" & cboGrupoAtiv & "'"
                Formula = "Filtro ='" & cboGrupoAtiv
                Aux = 1
            End If
            If Trim(txtDescAtiv) <> "" Then
                SelecaoRpt = SelecaoRpt & IIf(Aux = 1, " and ", "") & " {Tab_Atividade_Economica.tae_nome} like '" & txtDescAtiv & "%' or {Tab_Atividade_Economica.tae_nome}  like '% " & txtDescAtiv & "%'"
                If Aux = 1 Then
                    Formula = Formula & " e Nome Atividade = " & txtDescAtiv
                Else
                    Formula = "Filtro =' e Nome Atividade =" & txtDescAtiv
                End If
            End If
            .Selecao = SelecaoRpt
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
            .Arvore = False
            .Visualizar
    End With
End Sub

Private Sub Form_Activate()
    Dim Sql As String
    
    atividade.PreencheGrid grdAtividade
    'sql = "Select tae_cae as Código, tae_nome as Atividade, tga_nome as Grupo, tae_valor as [Valor(R$)], tae_desc_fator as Fator from Tab_Atividade_Economica, Tab_Grupo_Atividade where tae_tga_cod_grupo = tga_cod_grupo"
   ' stAtv..Preencher Bdados, sql, 1400
        
    AtualizaCabecalho grdAtividade
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set atividade = New atividade
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    atividade.PreencheCombo cboGrupoAtiv, iaGrupoAtividade
    cboEstimativa.PreencherGeral Bdados, "SIM OU NÃO"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set atividade = Nothing
End Sub

Private Sub grdAtividade_DblClick()
    'codAtividade = grdAtividade.ListItems(grdAtividade.SelectedItem.Index)
End Sub

