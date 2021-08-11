VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form PLCO101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLCO101"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   7
      Top             =   6855
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   847
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   345
         Left            =   6180
         TabIndex        =   12
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   5190
         TabIndex        =   4
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   8175
         TabIndex        =   6
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   7170
         TabIndex        =   5
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   1138
      Icone           =   "PLCO101.frx":0000
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   3585
      Left            =   45
      TabIndex        =   9
      Top             =   3210
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6324
      CorBorda        =   32768
      Caption         =   "Conta/Subconta"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1365
      Left            =   60
      TabIndex        =   10
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   1785
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   2408
      Altura          =   1905
      Caption         =   " Dados do Proprietário"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtContaMae 
         Height          =   480
         Left            =   5205
         TabIndex        =   3
         Tag             =   "                    "
         Top             =   825
         Visible         =   0   'False
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   847
         Caption         =   "Conta"
         Text            =   ""
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   50
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   510
         Left            =   285
         TabIndex        =   2
         Tag             =   "Tipo"
         ToolTipText     =   "783 - TIPO CONTA"
         Top             =   795
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   900
         Caption         =   "Tipo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   4210752
      End
      Begin VTOcx.txtVISUAL txtDescricao 
         Height          =   480
         Left            =   2415
         TabIndex        =   1
         Tag             =   "Descrição"
         Top             =   300
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   847
         Caption         =   "Descrição"
         Text            =   ""
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   300
      End
      Begin VTOcx.txtVISUAL txtContaSubconta 
         Height          =   480
         Left            =   255
         TabIndex        =   0
         Tag             =   "Conta/Subconta"
         Top             =   300
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   847
         Caption         =   "Conta/subconta"
         Text            =   ""
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   50
      End
   End
   Begin VTOcx.txtVISUAL txtCodigo 
      Height          =   480
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   847
      Caption         =   ""
      Text            =   ""
      Requerido       =   0   'False
      AlinhamentoRotulo=   1
      CorRotulo       =   16384
      CorTexto        =   4194304
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1080
      Left            =   60
      TabIndex        =   13
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   675
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   1905
      Altura          =   1905
      Caption         =   " Dados do Proprietário"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   465
         TabIndex        =   17
         Top             =   750
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   75
         TabIndex        =   16
         Tag             =   "Insc. Municipal"
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3165
         TabIndex        =   15
         Top             =   375
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   503
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Top             =   375
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
End
Attribute VB_Name = "PLCO101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GeraCod As New ContaCorrente

Private Sub cboTipo_Click()
    If cboTipo.Coluna(1).Valor = 2 Then
        txtContaMae.Visible = True
    ElseIf cboTipo.Coluna(1).Valor = 1 Then
        txtContaMae.Visible = False
    End If
End Sub

Private Sub cmdExcluir_Click()
    Dim condicao As String
    If grdDados.ListItems.Count < 1 Then Exit Sub
    condicao = "TPC_COD_CONTA = '" & txtCodigo & "'"
    If txtCodigo <> "" Then
        If Confirma("Deseja excluir registro?", "Excluir?") Then
            If Bdados.DeletaDados("TAB_PLANO_CONTAS", condicao) Then
                Avisa "Dados Excluidos com Sucesso"
                
                If grdDados.ListItems.Count <= 1 Then
                    cmdLimpar_Click
                    CarregaContas
                Else
                    CarregaContas
                    Limpa
                End If
            End If
        End If
    Else
        Avisa "Selecione um Registro"
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    CarregaContas
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub



Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
    If Len(txtIm) >= 10 Then txtIm.Formato = formDoisDigitos
    grdDados.ListItems.Clear
    CarregaContas
    txtIm.Formato = formNenhum
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim campos As String
    Dim valores As String
    Dim condicao As String
    Dim Codigo As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If txtCodigo = "" Then
        Codigo = CStr(GeraCod.GeraCodPagamento(70))
    Else
        Codigo = txtCodigo
    End If
       
    campos = " TPC_COD_CONTA,TPC_CONTA,TPC_DESCRICAO,TPC_TIPO,TPC_COD_CONTA_MAE,TPC_TCI_IM"
    valores = Bdados.PreparaValor(Codigo, txtContaSubconta, txtDescricao, cboTipo.Coluna(1).Valor, txtContaMae, txtIm)
    condicao = "TPC_COD_CONTA = '" & Codigo & "'"
    If Bdados.GravaDados("TAB_PLANO_CONTAS", valores, campos, condicao) Then
        Avisa "Dados Salvos com Sucesso"
        Limpa
        txtContaSubconta.SetFocus
        CarregaContas
        txtContaMae.Visible = False
        
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
    cboTipo.PreencherGeral Bdados, "TIPO CONTA"
    CarregaContas
End Sub

Private Sub Limpa()
    txtCodigo = ""
    txtContaSubconta = ""
    txtDescricao = ""
    cboTipo.ListIndex = -1
    txtContaMae = ""
End Sub

Private Sub CarregaContas()
    Dim sql As String
    Dim condicao As String
    Limpa
    If txtIm <> "" Then
        condicao = " and TPC_TCI_IM =  '" & txtIm & "'"
    End If
    
    sql = "select TCI_NOME as Banco,"
    sql = sql & " TGE_NOME as Tipo,"
    sql = sql & " TPC_CONTA as Conta_SubConta,"
    sql = sql & " TPC_COD_CONTA_MAE as Conta,"
    sql = sql & " TPC_DESCRICAO as Descrição,"
    sql = sql & " TPC_TIPO,"
    sql = sql & " TPC_TCI_IM,"
    sql = sql & " TPC_COD_CONTA"
    sql = sql & " from tab_plano_contas,"
    sql = sql & " vis_tipo_conta , TAB_CONTRIBUINTE"
    sql = sql & " Where tge_codigo = TPC_TIPO"
    sql = sql & " and TCI_IM  = TPC_TCI_IM "
    
    If condicao <> "" Then
        sql = sql & condicao & "order by TCI_NOME,TPC_COD_CONTA_MAE,TPC_CONTA"
    Else
        sql = sql & "order by TCI_NOME,TPC_COD_CONTA_MAE,TPC_CONTA"
    End If
    grdDados.Preencher Bdados, sql, 2000, 1500, 2000, 1500, 3000, 0, 0, 0
End Sub

Private Sub grdDados_DblClick()
    If grdDados.ListItems.Count >= 1 Then
        txtCodigo = grdDados.SelectedItem.SubItems(7)
        txtContaSubconta = grdDados.SelectedItem.SubItems(2)
        txtDescricao = grdDados.SelectedItem.SubItems(4)
        cboTipo.SetarLinha grdDados.SelectedItem.SubItems(5), 1
        txtContaMae = grdDados.SelectedItem.SubItems(3)
        If txtIm = "" Then
            txtIm = grdDados.SelectedItem.SubItems(6)
            txtIm_LostFocus
            grdDados_DblClick
        End If
        If cboTipo.Coluna(1).Valor = 2 Then
            txtContaMae.Visible = True
        Else
            txtContaMae.Visible = False
        End If
    End If
End Sub
