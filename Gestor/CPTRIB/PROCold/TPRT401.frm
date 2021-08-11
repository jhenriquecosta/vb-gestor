VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TPRT401 
   Caption         =   "TPRT401"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   1138
      Icone           =   "TPRT401.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   15
      Top             =   6870
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL CmdImprimirListagem 
         Height          =   330
         Left            =   2775
         TabIndex        =   17
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         Caption         =   "&Listagem"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   330
         Left            =   3930
         TabIndex        =   9
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   6045
         TabIndex        =   11
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
         Left            =   7035
         TabIndex        =   12
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
         Left            =   8025
         TabIndex        =   13
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdFicha 
         Height          =   330
         Left            =   5055
         TabIndex        =   10
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "&Ficha"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   2175
      Left            =   45
      TabIndex        =   16
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   690
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   3836
      Altura          =   1905
      Caption         =   " Dados do Processo"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtProcesso 
         Height          =   285
         Left            =   255
         TabIndex        =   18
         Tag             =   "Insc. Municipal"
         Top             =   435
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         Caption         =   "Nº Protocolo"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   0
         AgruparValores  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboSubAcao 
         Height          =   315
         Left            =   570
         TabIndex        =   5
         Top             =   1755
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   556
         Caption         =   "SubAção"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEntrega 
         Height          =   285
         Left            =   6270
         TabIndex        =   7
         Top             =   2340
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   503
         Caption         =   "Dt Entrega"
         Text            =   ""
         Formato         =   0
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtDataEntrada 
         Height          =   285
         Left            =   6240
         TabIndex        =   6
         Top             =   1785
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   503
         Caption         =   "Dt Entrada"
         Text            =   ""
         Formato         =   0
         AgruparValores  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboAcao 
         Height          =   315
         Left            =   900
         TabIndex        =   4
         Top             =   1395
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   556
         Caption         =   "Ação"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdBUsca 
         Height          =   285
         Left            =   2805
         TabIndex        =   1
         Top             =   750
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
         Left            =   525
         TabIndex        =   3
         Top             =   1065
         Width           =   7995
         _ExtentX        =   14102
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
         Left            =   60
         TabIndex        =   0
         Tag             =   "Insc. Municipal"
         Top             =   750
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
         Left            =   3150
         TabIndex        =   2
         Top             =   750
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   503
         Caption         =   ""
         Text            =   ""
         CorRotulo       =   0
         CorTexto        =   4194304
      End
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   3840
      Left            =   45
      TabIndex        =   8
      Top             =   2940
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   6773
      CorBorda        =   32768
      Caption         =   "Processos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
End
Attribute VB_Name = "TPRT401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SelecaoListagem As String

Private Sub cmdBuscar_Click()
    Buscar
End Sub

Private Sub Buscar()
    Dim Sql As String
       
    Sql = " SELECT TAB_PROTOCOLO.TPR_PROCESSO AS Processo, TAB_PARAMETRO_PROTOCOLO.TPP_NOME_PARAMETRO AS Ação,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_REQUERENTE AS Insc_Municiapal, TAB_PROTOCOLO.TPR_NOME_REQUERENTE AS Requerente,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_ENDERECO AS Endereço, TAB_PROTOCOLO.TPR_DATA_ENTRADA AS Data_Entrada,"
    Sql = Sql & " VIS_STATUS_PROTOCOLO.TGE_NOME AS Status, VIS_STATUS_FINAL_PROTOCOLO.TGE_NOME AS Homologação,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_DATA_HOMOLOGACAO AS Data_Homologação, TAB_PROTOCOLO.TPR_USUARIO_HOMOLOGACAO AS Func_Homologação,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_MOTIVO_HOMOLOGACAO AS Obs_Homologação, TAB_PROTOCOLO.TPR_DATA_ENTREGA AS Data_Entrega,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_TUS_USUARIO AS Usuário, TAB_PROTOCOLO.TPR_RESPONSAVEL AS Responsável,"
    Sql = Sql & " TAB_PROTOCOLO.TPR_ASSUNTO AS Assunto, TAB_PROTOCOLO.TPR_ACAO AS Cod_Ação, TAB_PROTOCOLO.TPR_SUBACAO AS Cod_SubAção"
    Sql = Sql & " FROM TAB_PROTOCOLO LEFT OUTER JOIN"
    Sql = Sql & " VIS_STATUS_PROTOCOLO ON TAB_PROTOCOLO.TPR_STATUS = VIS_STATUS_PROTOCOLO.TGE_CODIGO LEFT OUTER JOIN"
    Sql = Sql & " TAB_PARAMETRO_PROTOCOLO ON TAB_PROTOCOLO.TPR_ACAO = TAB_PARAMETRO_PROTOCOLO.TPP_TIPO_PARAMETRO AND"
    Sql = Sql & " TAB_PROTOCOLO.TPR_SUBACAO = TAB_PARAMETRO_PROTOCOLO.TPP_CODIGO_PARAMETRO LEFT OUTER JOIN"
    Sql = Sql & " VIS_STATUS_FINAL_PROTOCOLO ON TAB_PROTOCOLO.TPR_STATUS_HOMOLOGACAO = VIS_STATUS_FINAL_PROTOCOLO.TGE_CODIGO"
    Sql = Sql & " Where (1 = 1)"

    SelecaoListagem = "1 = 1"
    If txtProcesso <> "" Then
        Sql = Sql & " and TPR_PROCESSO = '" & txtProcesso & "'"
        SelecaoListagem = SelecaoListagem & " and {TAB_PROTOCOLO.TPR_PROCESSO} = '" & txtProcesso & "'"
    End If
    If txtIm <> "" Then
        Sql = Sql & " and TPR_REQUERENTE = " & txtIm
        SelecaoListagem = SelecaoListagem & " and {TAB_PROTOCOLO.TPR_REQUERENTE} = " & txtIm
    End If
    If txtDataEntrada <> "" Then
        Sql = Sql & " and TPR_DATA_ENTRADA = " & Bdados.Converte(txtDataEntrada, TCDataHora)
        SelecaoListagem = SelecaoListagem & " and {TAB_PROTOCOLO.TPR_DATA_ENTRADA} = #" & txtDataEntrada & "#"
    End If
    If cboAcao.ListIndex <> -1 Then
        Sql = Sql & " and TPR_ACAO = " & cboAcao.Coluna(1).Valor
        SelecaoListagem = SelecaoListagem & " and {TAB_PROTOCOLO.TPR_ACAO} = '" & cboAcao.Coluna(1).Valor & "'"
    End If
    If cboSubAcao.ListIndex <> -1 Then
        Sql = Sql & " and TPR_SUBACAO = " & cboSubAcao.Coluna(1).Valor
        SelecaoListagem = SelecaoListagem & " and {TAB_PROTOCOLO.TPR_SUBACAO} = " & cboSubAcao.Coluna(1).Valor
    End If
    If txtNomeContrib <> "" Then
        Sql = Sql & " AND TPR_NOME_REQUERENTE LIKE '%" & txtNomeContrib & "%'"
        SelecaoListagem = SelecaoListagem & " AND {TAB_PROTOCOLO.TPR_NOME_REQUERENTE} LIKE '%" & txtNomeContrib & "%'"
    End If
    Sql = Sql & " order  by TPR_PROCESSO"
    If Not grdDados.Preencher(Bdados, Sql) Then Avisa "Busca sem resultados "
End Sub




Private Sub cmdBUsca_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdFicha_Click()
    Dim Selecao As String
    Dim i As Integer
    If grdDados.ListItems.Count >= 1 Then
        For i = 1 To grdDados.ListItems.Count
            If grdDados.ListItems(i).Selected Then
                Selecao = "{TAB_PROTOCOLO.TPR_PROCESSO} = '" & grdDados.ListItems(i) & "'"
            End If
        Next
    End If
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TProcessoFicha.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Formulas "VT_DATA", AplicacoesVTFuncoes.municipio & " -  " & Temp.PegaParametro(Bdados, "ESTADO CLIENTE") & "  " & FormatDateTime(Date, vbLongDate)
        .Titulo = "Cartão do Protocolo"
        .SubRelatorio = "TSubProcesso.rpt"
        'TPR_ACAO,TPR_SUBACAO
        .Selecao = "{VIS_ACAO_PROTOCOLO.TPP_TIPO_PARAMETRO} = " & grdDados.SelectedItem.SubItems(15) & " and  {VIS_SUBACAO_PROTOCOLO.TPP_CODIGO_PARAMETRO} = " & grdDados.SelectedItem.SubItems(16)
        .SubRelatorio = ""
        .Selecao = Selecao
        .Visualizar
        
    End With
End Sub

Private Sub cmdImprimir_Click()
    Dim Selecao As String
    Dim i As Integer
    Dim rs As VSRecordset
    If grdDados.ListItems.Count >= 1 Then
        For i = 1 To grdDados.ListItems.Count
            If grdDados.ListItems(i).Selected Then
                Selecao = "{TAB_PROTOCOLO.TPR_PROCESSO} = '" & grdDados.ListItems(i) & "'"
            End If
        Next
    End If
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TProcesso.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Formulas "vt_texto_cabecalho", Temp.PegaParametro(Bdados, "TEXTO PROCESSO")
        .Formulas "VT_DATA", AplicacoesVTFuncoes.municipio & " -  " & Temp.PegaParametro(Bdados, "ESTADO CLIENTE") & "  " & FormatDateTime(Date, vbLongDate)
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Titulo = "Ficha de Processo"
        .SubRelatorio = "TSubProcesso.rpt"
        'TPR_ACAO,TPR_SUBACAO
        If Bdados.AbreTabela("Select TPR_ACAO,TPR_SUBACAO from tab_Protocolo where TPR_PROCESSO = '" & grdDados.SelectedItem & "'", rs) Then
            .Selecao = "{VIS_ACAO_PROTOCOLO.TPP_TIPO_PARAMETRO} = " & rs.Fields("TPR_ACAO") & " and  {VIS_SUBACAO_PROTOCOLO.TPP_CODIGO_PARAMETRO} = " & rs.Fields("TPR_SUBACAO")
        End If
        .SubRelatorio = ""
        .Selecao = Selecao
        .Visualizar
        
    End With
    
End Sub

Private Sub CmdImprimirListagem_Click()
    
    Dim i As Integer
    Dim rs As VSRecordset
        
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TListagemProtocolo.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Selecao = SelecaoListagem
        .Visualizar
        
    End With
End Sub

Private Sub cmdLimpar_Click()
    txtIm = ""
    txtDataEntrada = ""
    txtEntrega = ""
    cboAcao.ListIndex = -1
    txtIm.SetFocus
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    LimpaCampos Me
    cboAcao.Preencher Bdados, "SELECT TPP_NOME_PARAMETRO , TPP_TIPO_PARAMETRO FROM TAB_PARAMETRO_PROTOCOLO where TPP_CODIGO_PARAMETRO = 0  ORDER BY TPP_TIPO_PARAMETRO"
    txtProcesso.Restricao = restrNenhuma
End Sub

Private Sub cboAcao_Click()
    Dim Sql As String
    
    Sql = " SELECT  TPP_NOME_PARAMETRO ,TPP_CODIGO_PARAMETRO"
    Sql = Sql & " From TAB_PARAMETRO_PROTOCOLO"
    Sql = Sql & " Where (TPP_CODIGO_PARAMETRO <> 0)"
    Sql = Sql & " and  tpp_tipo_parametro =  " & cboAcao.Coluna(1).Valor
    Sql = Sql & " ORDER BY TPP_CODIGO_PARAMETRO "
    
    cboSubAcao.Preencher Bdados, Sql
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
