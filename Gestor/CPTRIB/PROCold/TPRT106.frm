VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TPRT106 
   Caption         =   "TPRT106"
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
      TabIndex        =   14
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1138
      Icone           =   "TPRT106.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   15
      Top             =   7305
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   6195
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
         TabIndex        =   12
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
         TabIndex        =   13
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
      TabIndex        =   16
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
      TabIndex        =   17
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
         Enabled         =   0   'False
         Height          =   1515
         Left            =   30
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "TPRT106.frx":031A
         Top             =   300
         Width           =   8940
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1650
      Left            =   45
      TabIndex        =   18
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   3720
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2910
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
         Top             =   1095
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   847
         Caption         =   "Funcionário "
         Text            =   ""
         Enabled         =   0   'False
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
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtDataMovimento 
         Height          =   480
         Left            =   7065
         TabIndex        =   9
         Tag             =   "Data Movimento"
         Top             =   1095
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   847
         Caption         =   "Data Movimento"
         Text            =   ""
         Enabled         =   0   'False
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
         Enabled         =   0   'False
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
      MarcaUnico      =   -1  'True
   End
End
Attribute VB_Name = "TPRT106"
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
    
    Sql = Sql & " SELECT TAB_MOVIMENTO_PROTOCOLO.TMP_MOVIMENTO AS Movimento,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_PROCESSO AS Processo,"
    Sql = Sql & "  VIS_SUBACAO_PROTOCOLO.TPP_NOME_PARAMETRO AS Ação,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_REQUERENTE AS Requerente,"
    Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO.TMP_USUARIO AS Func_Origem,"
    Sql = Sql & " tab_setor_protocolo.tpp_nome_parametro AS Origem,"
    Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO.TMP_FUNCIONARIO_DESTINO AS Func_Destino,"
    Sql = Sql & " tab_setor_protocolo_1.tpp_nome_parametro AS Destino"
    Sql = Sql & " FROM TAB_PROTOCOLO INNER JOIN"
    Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO ON"
    Sql = Sql & " TAB_PROTOCOLO.TPR_PROCESSO = TAB_MOVIMENTO_PROTOCOLO.TMP_TPR_PROTOCOLO"
    Sql = Sql & " INNER JOIN VIS_SUBACAO_PROTOCOLO ON"
    Sql = Sql & " TAB_PROTOCOLO.TPR_SUBACAO = VIS_SUBACAO_PROTOCOLO.tpp_codigo_parametro"
    Sql = Sql & " AND TAB_PROTOCOLO.TPR_ACAO = VIS_SUBACAO_PROTOCOLO.TPP_TIPO_PARAMETRO"
    Sql = Sql & " INNER JOIN tab_setor_protocolo ON"
    Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO.TMP_ORIGEM = tab_setor_protocolo.tpp_codigo_parametro"
    Sql = Sql & " INNER JOIN tab_setor_protocolo tab_setor_protocolo_1 ON"
    Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO.TMP_DESTINO = tab_setor_protocolo_1.tpp_codigo_parametro"
    Sql = Sql & " Where 1=1"
    aux = Trim(Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4))
    If txtIm <> "" Then Sql = Sql & " and TPR_REQUERENTE = " & txtIm
    If txtProcesso <> "" Then Sql = Sql & " and TPR_PROCESSO = '" & aux & "'"
    If txtNomeContrib <> "" Then Sql = Sql & " AND TPR_NOME_REQUERENTE LIKE '%" & txtNomeContrib & "%'"
     
    If Not grdDados.Preencher(Bdados, Sql) Then Avisa "Busca sem resultados "
   
End Sub

Private Sub cmdBUsca_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdLimpar_Click()
  LimpaCampos Me
  grdDados.ListItems.Clear
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

Private Sub Form_Load()
    Dim Sql As String
    
    checou = False
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    LimpaCampos Me
      
    Sql = "select * from tab_setor_protocolo"
    
    cboOrigem.Preencher Bdados, Sql
    cboDestino.Preencher Bdados, Sql
End Sub


Private Sub grdDados_DblClick()
     If grdDados.ListItems.Count < 1 Then Exit Sub
        Dim Sql As String
        Dim rs As VSRecordset
        
        Sql = Sql & " SELECT TAB_PROTOCOLO.TPR_PROCESSO,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_REQUERENTE,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_NOME_REQUERENTE,"
        Sql = Sql & " TAB_PROTOCOLO.TPR_ENDERECO,"
        Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO.TMP_DATA,"
        Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO.TMP_OBS,"
        Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO.TMP_FUNCIONARIO_DESTINO,"
        Sql = Sql & " tab_setor_protocolo.tpp_nome_parametro AS Origem,"
        Sql = Sql & " tab_setor_protocolo_1.tpp_nome_parametro AS Destino"
        Sql = Sql & " FROM TAB_PROTOCOLO INNER JOIN"
        Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO ON"
        Sql = Sql & " TAB_PROTOCOLO.TPR_PROCESSO = TAB_MOVIMENTO_PROTOCOLO.TMP_TPR_PROTOCOLO"
        Sql = Sql & " INNER JOIN tab_setor_protocolo tab_setor_protocolo_1 ON"
        Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO.TMP_DESTINO = tab_setor_protocolo_1.tpp_codigo_parametro"
        Sql = Sql & " INNER JOIN tab_setor_protocolo ON"
        Sql = Sql & " TAB_MOVIMENTO_PROTOCOLO.TMP_ORIGEM = tab_setor_protocolo.tpp_codigo_parametro"
        Sql = Sql & " where  TAB_MOVIMENTO_PROTOCOLO.TMP_MOVIMENTO = " & grdDados.SelectedItem
        If Bdados.AbreTabela(Sql, rs) Then
            txtProcesso = "" & rs!TPR_PROCESSO
            txtProcesso_LostFocus
            txtIm = "" & rs!TPR_REQUERENTE
            txtNomeContrib = "" & rs!TPR_NOME_REQUERENTE
            txtEndereco = "" & rs!TPR_ENDERECO
            cboOrigem.Text = "" & rs!Origem
            cboDestino.Text = "" & rs!Destino
            txtDataMovimento = "" & rs!TMP_DATA
            txtObservacao = "" & rs!TMP_OBS
            txtFuncionario = "" & rs!TMP_FUNCIONARIO_DESTINO
           
        End If
    

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



