VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TATV104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TATV104"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TATV104.frx":0000
   ScaleHeight     =   7095
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   5865
      Left            =   0
      TabIndex        =   1
      Top             =   645
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   10345
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      TagVariant      =   ""
      Tabs            =   "TATV104.frx":0342
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5475
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   9657
         _Version        =   131082
         TabGuid         =   "TATV104.frx":03BF
         Begin VTOcx.grdVISUAL GrdTaxa 
            Height          =   4305
            Left            =   135
            TabIndex        =   10
            Top             =   1095
            Width           =   8640
            _ExtentX        =   15240
            _ExtentY        =   7594
            CorBorda        =   16711680
            Caption         =   "Atividades"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
         Begin VTOcx.cboVISUAL CboTaxa 
            Height          =   315
            Left            =   195
            TabIndex        =   11
            Top             =   645
            Width           =   3720
            _ExtentX        =   6562
            _ExtentY        =   556
            Caption         =   "Taxa"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtValor 
            Height          =   285
            Left            =   6255
            TabIndex        =   12
            Top             =   660
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   503
            Caption         =   "Valor"
            Text            =   ""
            Formato         =   5
            Restricao       =   2
         End
         Begin VTOcx.cmdVISUAL cmdMais 
            Height          =   345
            Left            =   8010
            TabIndex        =   13
            Top             =   645
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   609
            Caption         =   ""
            Acao            =   1
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL CmdMenos 
            Height          =   345
            Left            =   8385
            TabIndex        =   14
            Top             =   645
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   609
            Caption         =   ""
            Acao            =   2
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtValorUFM 
            Height          =   285
            Left            =   3975
            TabIndex        =   21
            Tag             =   "Valor"
            Top             =   675
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   503
            Caption         =   "Valor(UFM)"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            MaxLen          =   10
         End
         Begin VB.Label LblAtividade 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Left            =   165
            TabIndex        =   20
            Top             =   150
            Width           =   8550
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5475
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   9657
         _Version        =   131082
         TabGuid         =   "TATV104.frx":03E7
         Begin VTOcx.grdVISUAL grdAtividade 
            Height          =   3960
            Left            =   0
            TabIndex        =   3
            Top             =   1500
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   6985
            CorBorda        =   16711680
            Caption         =   "Atividades"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1440
            Left            =   15
            TabIndex        =   4
            Top             =   0
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
               TabIndex        =   8
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
               TabIndex        =   7
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
               TabIndex        =   6
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
               TabIndex        =   5
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
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   6570
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4350
         TabIndex        =   16
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   5490
         TabIndex        =   17
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdBusca 
         Height          =   375
         Left            =   3210
         TabIndex        =   15
         Top             =   105
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
         Left            =   6630
         TabIndex        =   18
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
         TabIndex        =   19
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
      TabIndex        =   22
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   1138
      Icone           =   "TATV104.frx":040F
   End
End
Attribute VB_Name = "TATV104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim atividade As atividade

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
    grdAtividade.ListItems.Clear
    GrdTaxa.ListItems.Clear
End Sub

Private Sub cmdMais_Click()
    Dim Index    As Integer
    Dim PosTaxa  As Integer
    Dim Contador As Integer
    
    If CboTaxa.ListIndex = -1 Then
        Util.Avisa "Selecione Taxa."
        CboTaxa.SetFocus
        Exit Sub
    End If
    
    If txtValor = "" Then
        Util.Avisa "Informe valor."
        txtValor.SetFocus
        Exit Sub
    End If
    
    If txtValorUFM = "" Then
        Util.Avisa "Informe o valor em UFM."
        txtValorUFM.SetFocus
        Exit Sub
    End If
    
    'Checo se a taxa já foi inseria.
    
    
    For Contador = 1 To GrdTaxa.ListItems.Count
        PosTaxa = InStr(GrdTaxa.ListItems(Contador), " - ") - 1
        If Left(GrdTaxa.ListItems(Contador), PosTaxa) = CStr(CboTaxa.Coluna(0).Valor) Then
            Util.Avisa "Taxa já foi inserida na tabela."
            CboTaxa.SetFocus
            Exit Sub
        End If
    Next
    
    
    Index = GrdTaxa.ListItems.Count + 1
    
    GrdTaxa.ListItems.Add Index, , CboTaxa.Coluna(0).Valor & " - " & CboTaxa.Text
    GrdTaxa.ListItems(Index).SubItems(1) = txtValorUFM
    GrdTaxa.ListItems(Index).SubItems(2) = txtValor
    CboTaxa.ListIndex = CboTaxa.ListIndex + 1
    txtValor = ""
    txtValorUFM = ""
    txtValor.SetFocus
End Sub

Private Sub CmdMenos_Click()
    If GrdTaxa.ListItems.Count >= 1 Then
        GrdTaxa.ListItems.Remove GrdTaxa.SelectedItem.Index
    End If
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

Private Sub cmdSalvar_Click()
    Dim Valores  As String
    Dim Campos   As String
    Dim Contador As Integer
    Dim Foi      As Integer
    Dim PosTaxa  As Integer
    
     
    If Bdados.DeletaDados("TAB_TAXA_ATIVIDADE", " TTA_TAE_CODIGO = '" & grdAtividade.SelectedItem & "'") Then
        For Contador = 1 To GrdTaxa.ListItems.Count
              Campos = "TTA_TAE_CODIGO, TTA_CODIGO_TAXA,TTA_VALOR,TTA_VALOR_REAL"
              PosTaxa = InStr(GrdTaxa.ListItems(Contador), " - ") - 1
              Valores = Bdados.PreparaValor(grdAtividade.SelectedItem, Left(GrdTaxa.ListItems(Contador), PosTaxa), GrdTaxa.ListItems(Contador).SubItems(1), GrdTaxa.ListItems(Contador).SubItems(2))
              If Bdados.InsereDados("tab_taxa_atividade", Valores, Campos) Then
                Foi = Foi + 1
              End If
        Next
        If Foi = GrdTaxa.ListItems.Count Then
            Util.Avisa "Operação concluída com sucesso."
            cmdLimpar_Click
        Else
            Util.Avisa "Erro ao gravar taxa. "
        End If
    Else
        Util.Avisa "Erro ao excluir taxas."
    End If
End Sub

Private Sub Form_Activate()
    Dim Sql As String
    
    atividade.PreencheGrid grdAtividade
    'sql = "Select tae_cae as Código, tae_nome as Atividade, tga_nome as Grupo, tae_valor as [Valor(R$)], tae_desc_fator as Fator from Tab_Atividade_Economica, Tab_Grupo_Atividade where tae_tga_cod_grupo = tga_cod_grupo"
   ' stAtv..Preencher Bdados, sql, 1400
    
    CboTaxa.Preencher Bdados, "SELECT * FROM TAB_IMPOSTO WHERE TIP_SIGLA_IMPOSTO LIKE '%TF%'", 1
    AtualizaCabecalho grdAtividade
    GrdTaxa.ColumnHeaders.Clear
    GrdTaxa.ColumnHeaders.Add , , "Taxa", 1000
    GrdTaxa.ColumnHeaders.Add , , "Valor UFM", 2000
    GrdTaxa.ColumnHeaders.Add , , "Valor Real", 2000
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
    If grdAtividade.ListItems.Count >= 1 Then
        Dim Sql As String
        Sql = " SELECT TTA_CODIGO_TAXA " & Bdados.Concatena & "' - '" & Bdados.Concatena & " tip_sigla_imposto AS Taxa ,tta_valor as Valor_UFM,tta_valor_real as VaLor_Real  "
        Sql = Sql & " FROM TAB_TAXA_ATIVIDADE,tab_imposto "
        Sql = Sql & " where  tip_cod_imposto = TTA_CODIGO_TAXA and   "
        Sql = Sql & " tta_tae_codigo = '" & grdAtividade.SelectedItem & "'"
        GrdTaxa.Preencher Bdados, Sql
        SSActiveTabs1.Tabs(2).Selected = True
        LblAtividade = grdAtividade.SelectedItem & " - " & grdAtividade.SelectedItem.SubItems(1)
    Else
        LblAtividade = ""
    End If
End Sub

Private Sub GrdTaxa_DblClick()
    If GrdTaxa.ListItems.Count >= 1 Then
        Dim PosTaxa  As Integer
        Dim Contador As Integer
        
        PosTaxa = InStr(GrdTaxa.SelectedItem, " - ") - 1
        CboTaxa.SetarLinha Left(GrdTaxa.SelectedItem, PosTaxa)
        txtValorUFM = GrdTaxa.SelectedItem.SubItems(1)
        txtValor = GrdTaxa.SelectedItem.SubItems(2)
        GrdTaxa.ListItems.Remove GrdTaxa.SelectedItem.Index
    End If
End Sub

Private Sub txtValor_LostFocus()
    If txtValor = "" Then Exit Sub
    txtValorUFM = Calcula_UFM(txtValor, Converete_UFM)
End Sub

Private Sub txtValorUFM_LostFocus()
   If txtValorUFM = "" Then Exit Sub
   txtValor = Calcula_UFM(txtValorUFM, Converete_Real)
End Sub
