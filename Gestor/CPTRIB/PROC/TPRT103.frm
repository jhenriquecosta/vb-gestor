VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TPRT103 
   Caption         =   "TPRT103"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1138
      Icone           =   "TPRT103.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   17
      Top             =   7635
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   5205
         TabIndex        =   12
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
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   330
         Left            =   6195
         TabIndex        =   13
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   32768
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1320
      Left            =   45
      TabIndex        =   18
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
         Caption         =   "Endere?o"
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
         Caption         =   "N? Processo"
         Text            =   ""
         TipoLetras      =   0
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1845
      Left            =   45
      TabIndex        =   19
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   5730
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   3254
      Altura          =   1905
      Caption         =   " Solicita??o"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VB.TextBox txtMotivo 
         Appearance      =   0  'Flat
         Height          =   1515
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "TPRT103.frx":031A
         Top             =   300
         Width           =   8940
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1860
      Left            =   45
      TabIndex        =   20
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   3795
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3281
      Altura          =   1905
      Caption         =   " Dados do Processo"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboAcao 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   390
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   556
         Caption         =   "A??o"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Enabled         =   0   'False
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
         Enabled         =   0   'False
         Formato         =   0
         AgruparValores  =   0   'False
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
         Enabled         =   0   'False
         Formato         =   0
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtResponsavel 
         Height          =   285
         Left            =   210
         TabIndex        =   8
         Top             =   1095
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   503
         Caption         =   "Respons?vel"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.cboVISUAL cboSubAcao 
         Height          =   315
         Left            =   525
         TabIndex        =   7
         Top             =   750
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   556
         Caption         =   "SubA??o"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Enabled         =   0   'False
      End
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   1905
      Left            =   30
      TabIndex        =   5
      Top             =   2070
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
Attribute VB_Name = "TPRT103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    Buscar
End Sub

Private Sub Buscar()
    Dim Sql As String
    Dim aux As String
    Sql = Sql & " select TPR_PROCESSO as Processo ,"
    Sql = Sql & " TPP_NOME_PARAMETRO as A??o ,"
    Sql = Sql & " TPR_REQUERENTE as  Insc_Municiapal,"
    Sql = Sql & " TPR_NOME_REQUERENTE as Requerente,"
    Sql = Sql & " TPR_ENDERECO as Endere?o,"
    Sql = Sql & " TPR_DATA_ENTRADA as Data_Entrada,"
    Sql = Sql & " TPR_DATA_ENTREGA as Data_Entrega,"
    Sql = Sql & " TPR_TUS_USUARIO as Usu?rio,"
    Sql = Sql & " TPR_RESPONSAVEL as Respons?vel,"
    Sql = Sql & " TPR_ASSUNTO as Assunto,"
    Sql = Sql & " TGE_NOME as Status,"
    Sql = Sql & " TPR_ACAO as Cod_A??o,"
    Sql = Sql & " TPR_SUBACAO as Cod_SubA??o"
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
    If Not grdDados.Preencher(Bdados, Sql) Then Avisa "Busca sem resultados "
End Sub

Private Sub cmdExcluir_Click()
    Dim Proc As String
    Proc = Left(txtProcesso, 8) & "/" & Right(txtProcesso, 4)
    If txtProcesso <> "" Then
        If Confirma("Confirma a exclusao do processo N? " & txtProcesso) Then
            If Bdados.DeletaDados("TAB_PROTOCOLO", "tpr_processo = '" & Proc & "'") Then
                Avisa "Processo apagado com sucesso."
                Limpa
                Buscar
            End If
        End If
    End If
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

Private Sub Form_Load()
   cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    LimpaCampos Me
    cboAcao.Preencher Bdados, "SELECT TPP_NOME_PARAMETRO , TPP_TIPO_PARAMETRO FROM TAB_PARAMETRO_PROTOCOLO where TPP_CODIGO_PARAMETRO = 0 and TPP_TIPO_PARAMETRO <>999 ORDER BY TPP_TIPO_PARAMETRO"
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
        cboAcao.SetarLinha grdDados.SelectedItem.SubItems(11), 1
        cboAcao_Click
        cboSubAcao.SetarLinha grdDados.SelectedItem.SubItems(12), 1
       
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


