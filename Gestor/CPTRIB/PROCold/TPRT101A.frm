VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONT~1.OCX"
Begin VB.Form TPRT101A 
   Caption         =   "TPRT101"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1138
      Icone           =   "TPRT101A.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   14
      Top             =   6870
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   7035
         TabIndex        =   11
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
         TabIndex        =   12
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
         Left            =   6045
         TabIndex        =   10
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
      Height          =   1185
      Left            =   30
      TabIndex        =   15
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   660
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   2090
      Altura          =   1905
      Caption         =   " Dados do Requerente"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   465
         TabIndex        =   3
         Top             =   750
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Requerido       =   0   'False
         CorRotulo       =   0
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   0
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtNomeContrib 
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   390
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         Caption         =   "Nome"
         Text            =   ""
         CorRotulo       =   0
         CorTexto        =   4194304
      End
      Begin VTOcx.cmdVISUAL cmdBUsca 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
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
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1860
      Left            =   30
      TabIndex        =   16
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   1875
      Width           =   8970
      _ExtentX        =   15822
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   8
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
         Left            =   4605
         TabIndex        =   7
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
         TabIndex        =   4
         Top             =   390
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   556
         Caption         =   "Ação"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   3015
      Left            =   30
      TabIndex        =   17
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   3795
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   5318
      Altura          =   1905
      Caption         =   " Solicitação"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VB.TextBox txtMotivo 
         Appearance      =   0  'Flat
         Height          =   2685
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "TPRT101A.frx":031A
         Top             =   300
         Width           =   8910
      End
   End
End
Attribute VB_Name = "TPRT101A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rpt As New VSRelatorio

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

Private Sub cmdBUsca_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdLimpar_Click()
    txtIm = ""
    txtNomeContrib = ""
    txtEndereco = ""
    txtResponsavel = ""
    txtEndereco = ""
    txtDataEntrada = ""
    txtEntrega = ""
    txtMotivo = ""
    cboAcao.ListIndex = -1
    txtIm.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
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
        .Selecao = "{VIS_ACAO_PROTOCOLO.TPP_TIPO_PARAMETRO} = '" & cboAcao.Coluna(1).Valor & "' and  {VIS_SUBACAO_PROTOCOLO.TPP_CODIGO_PARAMETRO} = '" & cboSubAcao.Coluna(1).Valor & "'"
        .SubRelatorio = ""
        .Selecao = Selecao
        .Imprimir
    End With
    
End Sub
Private Sub cmdSalvar_Click()
    Dim valores   As String
    Dim campos   As String
    Dim condicao As String
    Dim Conta As New ContaCorrente
    Dim Processo As String
    Dim Entrega As String
    Dim rs As VSRecordset
    If UCase(AplicacoesVTFuncoes.municipio) = "PETROLINA" Then
        Processo = Conta.GeraCodPagamento(53) & "/" & Year(Date)
    Else
        Processo = Conta.GeraCodPagamento(53)
    End If
      
    
    If txtEntrega = "" Then
        Entrega = "null"
    Else
        Entrega = txtEntrega
    End If
    campos = "TPR_PROCESSO,TPR_ACAO,TPR_SUBACAO,TPR_REQUERENTE,TPR_DATA_ENTRADA,TPR_DATA_ENTREGA,TPR_TUS_USUARIO,TPR_RESPONSAVEL,TPR_ASSUNTO,TPR_ENDERECO,TPR_NOME_REQUERENTE,TPR_STATUS"
    valores = Bdados.PreparaValor(Processo, cboAcao.Coluna(1).Valor, cboSubAcao.Coluna(1).Valor, txtIm, Bdados.Converte(txtDataEntrada, TCDataHora), Entrega, AplicacoesVTFuncoes.Usuario, txtResponsavel, txtMotivo, txtEndereco, txtNomeContrib, spAberto)
    If Bdados.InsereDados("TAB_PROTOCOLO", valores, campos) Then
        Util.Avisa "Processo Aberto com sucesso" & vbCrLf & "Nº Processo " & Processo
        If Confirma("Deseja imprimir a ficha de processo?") Then
            With Rpt
                If Not .DefinirArquivo(Bdados, App.Path + "\TProcesso.rpt") Then Exit Sub
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                .Formulas "VT_DATA", AplicacoesVTFuncoes.municipio & " -  " & Temp.PegaParametro(Bdados, "ESTADO CLIENTE") & "  " & FormatDateTime(Date, vbLongDate)
                .Formulas "vt_texto_cabecalho", Temp.PegaParametro(Bdados, "TEXTO PROCESSO")
                .Titulo = "Ficha de Processo"
                .SubRelatorio = "TSubProcesso.rpt"
                'TPR_ACAO,TPR_SUBACAO
                If Bdados.AbreTabela("Select TPR_ACAO,TPR_SUBACAO from tab_Protocolo where TPR_PROCESSO = '" & Processo & "'", rs) Then
                    Dim strSelecao As String
                    strSelecao = "{VIS_ACAO_PROTOCOLO.TPP_TIPO_PARAMETRO} = '" & rs.Fields("TPR_ACAO") & "' and  {VIS_SUBACAO_PROTOCOLO.TPP_CODIGO_PARAMETRO} = '" & rs.Fields("TPR_SUBACAO") & "'"
                    .Selecao = strSelecao
                End If
                .SubRelatorio = ""
                .Selecao = "{TAB_PROTOCOLO.TPR_PROCESSO} = '" & Processo & "'"
                .Visualizar
                
            End With
            ImprimirFicha
        End If
        LimpaCampos Me
        Load TPRT105
        TPRT105.Tag = Processo
        TPRT105.Show
    End If
End Sub

Private Sub Form_Load()
    txtEntrega = Date
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    LimpaCampos Me
    cboAcao.Preencher Bdados, "SELECT TPP_NOME_PARAMETRO , TPP_TIPO_PARAMETRO FROM TAB_PARAMETRO_PROTOCOLO where TPP_CODIGO_PARAMETRO = 0 and TPP_TIPO_PARAMETRO <>999  ORDER BY TPP_TIPO_PARAMETRO"
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
