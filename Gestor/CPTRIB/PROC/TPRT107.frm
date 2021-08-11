VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TPRT107 
   Caption         =   "TPRT107"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1138
      Icone           =   "TPRT107.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   10
      Top             =   7305
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   6210
         TabIndex        =   18
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   5235
         TabIndex        =   6
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
         Left            =   7185
         TabIndex        =   7
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
         TabIndex        =   8
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1320
      Left            =   45
      TabIndex        =   11
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
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1650
      Left            =   45
      TabIndex        =   12
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   5625
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2910
      Altura          =   1905
      Caption         =   " Homologação"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VB.TextBox txtObservacao 
         Appearance      =   0  'Flat
         Height          =   825
         Left            =   990
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "TPRT107.frx":031A
         Top             =   765
         Width           =   7935
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   420
         TabIndex        =   15
         Tag             =   "Status"
         Top             =   375
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Editavel        =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtDataHomologacao 
         Height          =   285
         Left            =   5820
         TabIndex        =   5
         Tag             =   "Data Homologação"
         Top             =   390
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         Caption         =   "Data Homologação"
         Text            =   ""
         Formato         =   0
         AgruparValores  =   0   'False
      End
      Begin VB.Label obs 
         Caption         =   "Observação"
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   765
         Width           =   945
      End
   End
   Begin VTOcx.grdVISUAL grdMovimento 
      Height          =   2115
      Left            =   30
      TabIndex        =   13
      Top             =   3735
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   3731
      CorBorda        =   32768
      Caption         =   "Movimentos do Processo"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
      MarcaUnico      =   -1  'True
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   1905
      Left            =   30
      TabIndex        =   14
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
      MarcaUnico      =   -1  'True
   End
End
Attribute VB_Name = "TPRT107"
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
    Sql = Sql & " and TPR_STATUS = TGE_CODIGO"
    aux = Trim(Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4))
    If txtIm <> "" Then Sql = Sql & " and TPR_REQUERENTE = " & txtIm
    If txtProcesso <> "" Then Sql = Sql & " and TPR_PROCESSO = '" & aux & "'"
    If txtNomeContrib <> "" Then Sql = Sql & " AND TPR_NOME_REQUERENTE LIKE '%" & txtNomeContrib & "%'"
     Sql = Sql & " order  by TPR_PROCESSO"
    If Not grdDados.Preencher(Bdados, Sql) Then Avisa "Busca sem resultados "
   
End Sub

Private Sub cmdBUsca_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdLimpar_Click()
  LimpaCampos Me
  grdDados.ListItems.Clear
  grdMovimento.ListItems.Clear
  txtProcesso.SetFocus
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

Private Sub cmdSalvar_Click()
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    
    If txtProcesso = "" Then
        Avisa "Selecione um Processo"
        Exit Sub
    End If
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    campos = " TPR_STATUS_HOMOLOGACAO, TPR_MOTIVO_HOMOLOGACAO, TPR_DATA_HOMOLOGACAO, TPR_USUARIO_HOMOLOGACAO, TPR_STATUS"
    valores = Bdados.PreparaValor(cboStatus.Coluna(1).Valor, txtObservacao, txtDataHomologacao, AplicacoesVTFuncoes.Usuario, spFechado) & "'"
    condicao = "TPR_PROCESSO = '" & Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4) & "'"
    If Bdados.AtualizaDados(" TAB_PROTOCOLO", valores, campos, condicao) Then
        Util.Avisa "Processo encerrado com sucesso"
        cmdLimpar_Click
    End If
   
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    LimpaCampos Me
    cboStatus.Preencher Bdados, "SELECT *  FROM VIS_STATUS_FINAL_PROTOCOLO"
End Sub


Private Sub grdDados_DblClick()
     If grdDados.ListItems.Count < 1 Then Exit Sub
        Dim Sql As String
        Dim sqlMov As String
        Dim rs As VSRecordset
        
        Sql = "SELECT TAB_PROTOCOLO.TPR_PROCESSO,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_REQUERENTE,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_NOME_REQUERENTE,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_ENDERECO"
        Sql = Sql & " FROM  TAB_PROTOCOLO"
        Sql = Sql & " where  TAB_PROTOCOLO.TPR_PROCESSO = '" & grdDados.SelectedItem & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtProcesso = "" & rs!TPR_PROCESSO
            txtProcesso_LostFocus
            txtIm = "" & rs!TPR_REQUERENTE
            txtNomeContrib = "" & rs!TPR_NOME_REQUERENTE
            txtEndereco = "" & rs!TPR_ENDERECO
                    
        End If
        sqlMov = sqlMov & " SELECT tAB_MOVIMENTO_PROTOCOLO.TMP_MOVIMENTO AS Movimento, "
        sqlMov = sqlMov & " TAB_MOVIMENTO_PROTOCOLO.TMP_USUARIO AS Func_Origem,"
        sqlMov = sqlMov & " tab_setor_protocolo.tpp_nome_parametro AS Origem, "
        sqlMov = sqlMov & " TAB_MOVIMENTO_PROTOCOLO.TMP_FUNCIONARIO_DESTINO AS Func_Destino,"
        sqlMov = sqlMov & " tab_setor_protocolo_1.tpp_nome_parametro AS Destino, "
        sqlMov = sqlMov & " TAB_MOVIMENTO_PROTOCOLO.TMP_DATA AS Data,"
        sqlMov = sqlMov & " TAB_MOVIMENTO_PROTOCOLO.TMP_OBS AS Observação"
        sqlMov = sqlMov & " FROM TAB_MOVIMENTO_PROTOCOLO INNER JOIN"
        sqlMov = sqlMov & " tab_setor_protocolo ON TAB_MOVIMENTO_PROTOCOLO.TMP_ORIGEM = tab_setor_protocolo.tpp_codigo_parametro INNER JOIN"
        sqlMov = sqlMov & " tab_setor_protocolo tab_setor_protocolo_1 ON TAB_MOVIMENTO_PROTOCOLO.TMP_DESTINO = tab_setor_protocolo_1.tpp_codigo_parametro"
        sqlMov = sqlMov & " where TMP_TPR_PROTOCOLO = '" & Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4) & "'"
        If Not grdMovimento.Preencher(Bdados, sqlMov) Then Avisa "Não Existe Movimento para o Processo " & Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4)

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


