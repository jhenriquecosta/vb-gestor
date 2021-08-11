VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TPRT105 
   Caption         =   "TPRT105"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1138
      Icone           =   "TPRT105.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   16
      Top             =   7275
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   5205
         TabIndex        =   11
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   7170
         TabIndex        =   13
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   8160
         TabIndex        =   14
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   6195
         TabIndex        =   12
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1320
      Left            =   45
      TabIndex        =   17
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   705
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   2328
      Altura          =   1905
      Caption         =   " Dados do Processo"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdBUsca 
         Height          =   285
         Left            =   2790
         TabIndex        =   2
         Top             =   645
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   510
         TabIndex        =   4
         Top             =   960
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   0
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   45
         TabIndex        =   1
         Tag             =   "Insc. Municipal"
         Top             =   645
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   503
         Caption         =   "Insc. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   0
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtNomeContrib 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   645
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         Caption         =   "Nome"
         Text            =   ""
         CorRotulo       =   0
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtProcesso 
         Height          =   285
         Left            =   270
         TabIndex        =   0
         Top             =   330
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
         Caption         =   "Nº Processo"
         Text            =   ""
         TipoLetras      =   0
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1845
      Left            =   45
      TabIndex        =   18
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   5415
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   3254
      Altura          =   1905
      Caption         =   " Observação"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VB.TextBox txtObservacao 
         Appearance      =   0  'Flat
         Height          =   1515
         Left            =   30
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "TPRT105.frx":031A
         Top             =   300
         Width           =   8940
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1635
      Left            =   45
      TabIndex        =   19
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   3750
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2884
      Altura          =   1905
      Caption         =   " Dados da Movimentação"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtFuncionario 
         Height          =   480
         Left            =   840
         TabIndex        =   8
         Tag             =   "Funcionário"
         Top             =   1080
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   847
         Caption         =   "Funcionário "
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   0
         CorTexto        =   4194304
         MaxLen          =   120
      End
      Begin VTOcx.cboVISUAL cboOrigem 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Tag             =   "Origem"
         Top             =   390
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   556
         Caption         =   "Origem"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Editavel        =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtDataMovimento 
         Height          =   480
         Left            =   7125
         TabIndex        =   9
         Tag             =   "Data Movimento"
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   847
         Caption         =   "Data Movimento"
         Text            =   ""
         Formato         =   0
         AlinhamentoRotulo=   1
         AgruparValores  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboDestino 
         Height          =   315
         Left            =   165
         TabIndex        =   7
         Tag             =   "Destino"
         Top             =   750
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   556
         Caption         =   "Destino"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Editavel        =   -1  'True
      End
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   1905
      Left            =   30
      TabIndex        =   5
      Top             =   2055
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   3360
      CorBorda        =   32768
      Caption         =   "Processos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
      MarcaUnico      =   -1  'True
   End
End
Attribute VB_Name = "TPRT105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private checou As Boolean

Private Sub cmdBuscar_Click()
    Buscar
End Sub

Private Sub Buscar()
     Dim Sql As String
    Dim aux As String
    
    Sql = Sql & " select TPR_PROCESSO as Processo ,"
    Sql = Sql & " TPP_NOME_PARAMETRO as Ação ,"
    Sql = Sql & " TPR_REQUERENTE as  Insc_Municiapal,"
    Sql = Sql & " TPR_NOME_REQUERENTE as Requerente,"
    Sql = Sql & " TPR_ENDERECO as Endereço,"
    Sql = Sql & " TPR_DATA_ENTRADA as Data_Entrada,"
    Sql = Sql & " TPR_DATA_ENTREGA as Data_Entrega,"
    Sql = Sql & " TPR_TUS_USUARIO as Usuário,"
    Sql = Sql & " TPR_RESPONSAVEL as Responsável,"
    Sql = Sql & " TPR_ASSUNTO as Assunto,"
    Sql = Sql & " TGE_NOME as Status,"
    Sql = Sql & " TPR_ACAO as Cod_Ação,"
    Sql = Sql & " TPR_SUBACAO as Cod_SubAção"
    Sql = Sql & " From TAB_PROTOCOLO, TAB_PARAMETRO_PROTOCOLO,VIS_STATUS_PROTOCOLO"
    Sql = Sql & " Where TPR_ACAO = TPP_TIPO_PARAMETRO"
    Sql = Sql & " and TPR_SUBACAO = TPP_CODIGO_PARAMETRO"
    Sql = Sql & " and TPR_STATUS = " & spAberto
     Sql = Sql & " and TPR_STATUS = TGE_CODIGO"
    
    aux = Trim(Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4))
    If txtIm <> "" Then Sql = Sql & " and TPR_REQUERENTE = " & txtIm
    If txtProcesso <> "" Then Sql = Sql & " and TPR_PROCESSO = '" & aux & "'"
    If txtNomeContrib <> "" Then Sql = Sql & " AND TPR_NOME_REQUERENTE LIKE '%" & txtNomeContrib & "%'"
     Sql = Sql & " order  by TPR_PROCESSO"
    If Not grdDados.Preencher(Bdados, Sql) Then
        Avisa "Busca sem resultados "
    Else
        txtDataMovimento = Date
    End If
End Sub

Private Sub cmdSalvar_Click()
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    Dim Conta As New ContaCorrente
    Dim Movimento As String
    Dim aux As String
   
    If grdDados.ListItems.Count < 1 Then Exit Sub
    If Not checou Then
        Avisa "Selecione um processo"
        Exit Sub
    End If
    Movimento = Conta.GeraCodPagamento(32)
    aux = Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4)
    campos = "TMP_TPR_PROTOCOLO, TMP_MOVIMENTO, TMP_ORIGEM, TMP_DESTINO, TMP_DATA, TMP_USUARIO, TMP_OBS,TMP_FUNCIONARIO_DESTINO"
    valores = Bdados.PreparaValor(aux, Movimento, cboOrigem.Coluna(1).Valor, cboDestino.Coluna(1).Valor, txtDataMovimento, AplicacoesVTFuncoes.Usuario, txtObservacao, txtFuncionario)
    If Bdados.InsereDados(" TAB_MOVIMENTO_PROTOCOLO", valores, campos) Then
        Util.Avisa "Processo Movimentado com Sucesso"
        LimpaCampos Me
        carregaCbo
    End If
    checou = False
End Sub



Private Sub cmdBUsca_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdLimpar_Click()
  LimpaCampos Me
  grdDados.ListItems.Clear
  
End Sub

Private Sub limpar()
    txtProcesso = ""
    txtResponsavel = ""
    txtDataEntrada = ""
    txtEntrega = ""
    txtMotivo = ""
    cboAcao.ListIndex = -1
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If TPRT105.Tag <> "" Then
       txtProcesso = TPRT105.Tag
       Buscar
       If grdDados.ListItems.Count >= 1 Then
           grdDados.ListItems.Item(1).Checked = True
           checado
       End If
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    LimpaCampos Me
    carregaCbo
End Sub


Private Sub grdDados_ItemCheck(ByVal Item As MSComctlLib.IListItem)
    checado
End Sub

Private Sub txtIm_LostFocus()
    Dim Ic As String
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtIm) = 10 Or Len(txtIm) = 11 Then
            Ic = Imposto.FormataInscricao(txtIm, InscContrib)
        Else
            Ic = txtIm
        End If
    Else
            Ic = txtIm
    End If
    txtIm = BuscaContribuinte(Ic, txtNomeContrib, txtEndereco)
End Sub

Private Sub txtProcesso_LostFocus()
   txtProcesso = Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4)
End Sub

Private Sub carregaCbo()
    Dim Sql As String
    
    Sql = "select * from tab_setor_protocolo"
    
    cboOrigem.Preencher Bdados, Sql
    cboDestino.Preencher Bdados, Sql
End Sub

Private Sub checado()
     If grdDados.ListItems.Count < 1 Then Exit Sub
    For i = 1 To grdDados.ListItems.Count
        If (grdDados.ListItems(i).Checked) Then
            checou = True
            Exit For
        Else
            checou = False
        End If
        
    Next
       
        txtProcesso = grdDados.SelectedItem
        txtProcesso_LostFocus
        txtIm = grdDados.SelectedItem.SubItems(2)
        txtNomeContrib = grdDados.SelectedItem.SubItems(3)
        txtEndereco = grdDados.SelectedItem.SubItems(4)
End Sub
