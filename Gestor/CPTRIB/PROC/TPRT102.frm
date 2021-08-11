VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TPRT102 
   Caption         =   "TPRT102"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1138
      Icone           =   "TPRT102.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   17
      Top             =   7605
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   5145
         TabIndex        =   12
         Top             =   90
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
         Left            =   7110
         TabIndex        =   14
         Top             =   90
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
         Left            =   8100
         TabIndex        =   15
         Top             =   90
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
         Left            =   6120
         TabIndex        =   13
         Top             =   90
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
      Left            =   30
      TabIndex        =   18
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   660
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   2328
      Altura          =   1905
      Caption         =   " Dados do Processo"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
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
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1845
      Left            =   30
      TabIndex        =   19
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   5685
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   3254
      Altura          =   1905
      Caption         =   " Solicitação"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VB.TextBox txtMotivo 
         Appearance      =   0  'Flat
         Height          =   1515
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "TPRT102.frx":031A
         Top             =   300
         Width           =   8940
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1860
      Left            =   30
      TabIndex        =   20
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   3750
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3281
      Altura          =   1905
      Caption         =   " Dados do Processo"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboSubAcao 
         Height          =   315
         Left            =   525
         TabIndex        =   7
         Top             =   750
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   556
         Caption         =   "SubAção"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtResponsavel 
         Height          =   285
         Left            =   210
         TabIndex        =   8
         Top             =   1095
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   503
         Caption         =   "Responsável"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtEntrega 
         Height          =   285
         Left            =   6765
         TabIndex        =   10
         Top             =   1410
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         Caption         =   "Dt Entrega"
         Text            =   ""
         Formato         =   0
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtDataEntrada 
         Height          =   285
         Left            =   4590
         TabIndex        =   9
         Top             =   1410
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         Caption         =   "Dt Entrada"
         Text            =   ""
         Formato         =   0
         AgruparValores  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboAcao 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   390
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   556
         Caption         =   "Ação"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   1905
      Left            =   15
      TabIndex        =   5
      Top             =   2025
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   3360
      CorBorda        =   32768
      Caption         =   "Processos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
End
Attribute VB_Name = "TPRT102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rpt As New VSRelatorio

Private Sub cmdBUsca_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub



Private Sub cmdLimpar_Click()
   LimpaCampos Me
   grdDados.ListItems.Clear
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim valores   As String
    Dim campos   As String
    Dim condicao As String
    Dim Conta As New ContaCorrente
    Dim Processo As String
    Dim Entrega As String
    Dim rs As VSRecordset
    If txtProcesso = "" Then
        Avisa "Selecione um Processo"
        Exit Sub
    End If
    
    If txtEntrega = "" Then
        Entrega = "null"
    Else
        Entrega = Bdados.Converte(txtEntrega, TCDataHora)
    End If
    Processo = Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4)
    condicao = " tpr_processo  = '" & Processo & "'"
    campos = "TPR_ACAO,TPR_SUBACAO,TPR_DATA_ENTRADA,TPR_DATA_ENTREGA,TPR_TUS_USUARIO,TPR_RESPONSAVEL,TPR_ASSUNTO"
    valores = Bdados.PreparaValor(cboAcao.Coluna(1).Valor, cboSubAcao.Coluna(1).Valor, Bdados.Converte(txtDataEntrada, TCDataHora), Entrega, Bdados.Converte(AplicacoesVTFuncoes.Usuario, tctexto), Bdados.Converte(txtResponsavel, tctexto), Bdados.Converte(txtMotivo, tctexto))
    If Bdados.GravaDados("TAB_PROTOCOLO", valores, campos, condicao) Then
        Util.Avisa "Processo atualizado com sucesso" & vbCrLf & "Nº Processo " & Processo
        If Confirma("Deseja imprimir a ficha de Processo.") = False Then GoTo Vai
            With Rpt
                If Not .DefinirArquivo(Bdados, App.Path + "\TProcesso.rpt") Then Exit Sub
                .Titulo = "Ficha de Processo"
                .Formulas "vt_texto_cabecalho", Temp.PegaParametro(Bdados, "TEXTO PROCESSO")
                .Formulas "VT_DATA", AplicacoesVTFuncoes.municipio & " -  " & Temp.PegaParametro(Bdados, "ESTADO CLIENTE") & "  " & FormatDateTime(Date, vbLongDate)
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                'TPR_ACAO,TPR_SUBACAO
                .SubRelatorio = "TSubProcesso.rpt"
                If Bdados.AbreTabela("Select TPR_ACAO,TPR_SUBACAO from tab_Protocolo where TPR_PROCESSO = '" & Processo & "'", rs) Then
                    .Selecao = "{VIS_ACAO_PROTOCOLO.TPP_TIPO_PARAMETRO} = " & rs.Fields("TPR_ACAO") & " and  {VIS_SUBACAO_PROTOCOLO.TPP_CODIGO_PARAMETRO} = " & rs.Fields("TPR_SUBACAO")
                End If
                .SubRelatorio = ""
                .Selecao = "{TAB_PROTOCOLO.TPR_PROCESSO} = '" & Processo & "'"
                .Visualizar
            End With
            ImprimirFicha
Vai:
        Limpa
        Buscar
    End If
    
End Sub
Private Sub ImprimirFicha()
    Dim Selecao As String
    Dim i As Integer
    Dim rs As VSRecordset
    
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TProcessoFicha.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Formulas "vt_texto_cabecalho", "FICHA DE PROTOCOLO"
        .Formulas "VT_DATA", AplicacoesVTFuncoes.municipio & " -  " & Temp.PegaParametro(Bdados, "ESTADO CLIENTE") & "  " & FormatDateTime(Date, vbLongDate)
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Titulo = "Ficha de Processo"
        .SubRelatorio = "TSubProcesso.rpt"
        'TPR_ACAO,TPR_SUBACAO
        .Selecao = "{VIS_ACAO_PROTOCOLO.TPP_TIPO_PARAMETRO} = " & cboAcao.Coluna(1).Valor & " and  {VIS_SUBACAO_PROTOCOLO.TPP_CODIGO_PARAMETRO} = " & cboSubAcao.Coluna(1).Valor
        .SubRelatorio = ""
        .Selecao = Selecao
        
        .Imprimir
    End With
End Sub

Private Sub cmdBuscar_Click()
    Buscar
End Sub

Private Sub Buscar()
    Dim Sql As String
    Dim aux As String
    Sql = Sql & " SELECT TAB_PROTOCOLO.TPR_PROCESSO AS Processo, TAB_PARAMETRO_PROTOCOLO.TPP_NOME_PARAMETRO AS Ação, "
    Sql = Sql & " TAB_PROTOCOLO.TPR_REQUERENTE AS Insc_Municiapal, TAB_PROTOCOLO.TPR_NOME_REQUERENTE AS Requerente,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_ENDERECO AS Endereço, TAB_PROTOCOLO.TPR_DATA_ENTRADA AS Data_Entrada,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_DATA_ENTREGA AS Data_Entrega,VIS_STATUS_PROTOCOLO.TGE_NOME AS Status,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_RESPONSAVEL AS Responsável,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_ASSUNTO AS Assunto, TAB_PROTOCOLO.TPR_ACAO AS Cod_Ação, TAB_PROTOCOLO.TPR_SUBACAO AS Cod_SubAção"
    Sql = Sql & " FROM TAB_PROTOCOLO LEFT OUTER JOIN"
    Sql = Sql & " VIS_STATUS_PROTOCOLO ON TAB_PROTOCOLO.TPR_STATUS = VIS_STATUS_PROTOCOLO.TGE_CODIGO LEFT OUTER JOIN"
    Sql = Sql & " TAB_PARAMETRO_PROTOCOLO ON TAB_PROTOCOLO.TPR_ACAO = TAB_PARAMETRO_PROTOCOLO.TPP_TIPO_PARAMETRO AND"
    Sql = Sql & " TAB_PROTOCOLO.TPR_SUBACAO = TAB_PARAMETRO_PROTOCOLO.TPP_CODIGO_PARAMETRO LEFT OUTER JOIN"
    Sql = Sql & " VIS_STATUS_FINAL_PROTOCOLO ON TAB_PROTOCOLO.TPR_STATUS_HOMOLOGACAO = VIS_STATUS_FINAL_PROTOCOLO.TGE_CODIGO"
    Sql = Sql & " where TPR_STATUS = " & spAberto
  
    
    aux = Trim(Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4))
    If txtIm <> "" Then Sql = Sql & " and TPR_REQUERENTE = " & txtIm
    If txtProcesso <> "" Then Sql = Sql & " and TPR_PROCESSO = '" & aux & "'"
    If txtNomeContrib <> "" Then Sql = Sql & " AND TPR_NOME_REQUERENTE LIKE '%" & txtNomeContrib & "%'"
     Sql = Sql & " order  by TPR_PROCESSO"
    If Not grdDados.Preencher(Bdados, Sql) Then Avisa "Busca sem resultados "
End Sub

Private Sub cboAcao_Click()
    Dim Sql As String
    If cboAcao.ListIndex <> -1 Then
        Sql = " SELECT  TPP_NOME_PARAMETRO ,TPP_CODIGO_PARAMETRO"
        Sql = Sql & " From TAB_PARAMETRO_PROTOCOLO"
        Sql = Sql & " Where (TPP_CODIGO_PARAMETRO <> 0)"
        Sql = Sql & " and  tpp_tipo_parametro =  " & cboAcao.Coluna(1).Valor
        Sql = Sql & " ORDER BY TPP_CODIGO_PARAMETRO "
        cboSubAcao.Preencher Bdados, Sql
   
    End If
End Sub

Private Sub cboAcao_GotFocus()
     cboSubAcao.ListIndex = -1
End Sub

Private Sub cboAcao_LostFocus()
    If cboAcao.ListIndex = -1 Then cboSubAcao.Clear
End Sub

Private Sub Form_Load()
    txtEntrega = Date
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    LimpaCampos Me
    cboAcao.Preencher Bdados, "SELECT TPP_NOME_PARAMETRO , TPP_TIPO_PARAMETRO FROM TAB_PARAMETRO_PROTOCOLO where TPP_CODIGO_PARAMETRO = 0 and TPP_TIPO_PARAMETRO <>999 ORDER BY TPP_TIPO_PARAMETRO"
End Sub
Private Sub Limpa()
        txtProcesso = ""
        txtResponsavel = ""
        txtDataEntrada = ""
        txtEntrega = ""
        txtMotivo = ""
        cboAcao.ListIndex = -1
        cboSubAcao.ListIndex = -1
End Sub

Private Sub grdDados_DblClick()
        If grdDados.ListItems.Count < 1 Then Exit Sub
        txtProcesso = grdDados.SelectedItem
        txtProcesso_LostFocus
        txtIm = grdDados.SelectedItem.SubItems(2)
        txtNomeContrib = grdDados.SelectedItem.SubItems(3)
        txtEndereco = grdDados.SelectedItem.SubItems(4)
        txtResponsavel = grdDados.SelectedItem.SubItems(8)
        txtDataEntrada = grdDados.SelectedItem.SubItems(5)
        txtEntrega = grdDados.SelectedItem.SubItems(6)
        txtMotivo = grdDados.SelectedItem.SubItems(9)
        cboAcao.SetarLinha grdDados.SelectedItem.SubItems(10), 1
        cboAcao_Click
        cboSubAcao.SetarLinha grdDados.SelectedItem.SubItems(11), 1
        
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
