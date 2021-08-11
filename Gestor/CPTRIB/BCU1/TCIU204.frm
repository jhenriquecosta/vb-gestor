VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU204 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU204"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   795
      Left            =   30
      TabIndex        =   22
      Top             =   3150
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1402
      Altura          =   1905
      Caption         =   " Dimensões"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtAreaConstruida 
         Height          =   315
         Left            =   5940
         TabIndex        =   11
         Top             =   390
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         Caption         =   "Área Total Construída"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
      End
      Begin VTOcx.txtVISUAL txtAreaLote 
         Height          =   315
         Left            =   3750
         TabIndex        =   10
         Top             =   390
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         Caption         =   "Área Lote"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   20
      Top             =   6390
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   2385
         TabIndex        =   14
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Imprimir"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL CmdBuscar 
         Height          =   375
         Left            =   3555
         TabIndex        =   13
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7065
         TabIndex        =   16
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8220
         TabIndex        =   17
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4725
         TabIndex        =   12
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   5895
         TabIndex        =   15
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1470
      TabIndex        =   19
      Top             =   -570
      Width           =   375
   End
   Begin VTOcx.grdVISUAL grid 
      Height          =   2625
      Left            =   45
      TabIndex        =   18
      Top             =   3960
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   4630
      CorBorda        =   16711680
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   16711680
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   2475
      Left            =   30
      TabIndex        =   21
      Top             =   660
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4366
      Altura          =   1905
      Caption         =   " Dados Adicionais da Localização"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtSecao 
         Height          =   315
         Left            =   3600
         TabIndex        =   9
         Top             =   2085
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "Seção"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtLote 
         Height          =   315
         Left            =   780
         TabIndex        =   7
         Top             =   2085
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "Lote"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtQuadra 
         Height          =   315
         Left            =   2070
         TabIndex        =   8
         Top             =   2085
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "Quadra"
         Text            =   ""
      End
      Begin VTOcx.cboVISUAL cboLoteamento 
         Height          =   315
         Left            =   135
         TabIndex        =   6
         Top             =   1725
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   556
         Caption         =   "Loteamento"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtNumero 
         Height          =   285
         Left            =   7935
         TabIndex        =   4
         Top             =   1035
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Caption         =   "Nº"
         Text            =   ""
      End
      Begin VTOcx.cboVISUAL CboLogradouro 
         Height          =   315
         Left            =   2805
         TabIndex        =   3
         Top             =   1020
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtDescricao 
         Height          =   285
         Left            =   315
         TabIndex        =   1
         Top             =   690
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   503
         Caption         =   "Descrição"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   570
         TabIndex        =   0
         Top             =   360
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   503
         Caption         =   "Código"
         Text            =   ""
      End
      Begin VTOcx.cboVISUAL CboBairro 
         Height          =   315
         Left            =   615
         TabIndex        =   5
         Top             =   1380
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   556
         Caption         =   "Bairro"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipLogra 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   1020
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   556
         Caption         =   "Endereço"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1138
      Icone           =   "TCIU204.frx":0000
   End
End
Attribute VB_Name = "TCIU204"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cadastro As VSImposto
Private Selecao As String

Private Sub cboVISUAL1_Click()

End Sub

Private Sub cmdBuscar_Click()
    Dim Sql As String
    
    Selecao = " 1 = 1"
    Sql = "SELECT * FROM vis_edificio where 1 = 1 "
        
    If txtDescricao <> "" Then
        Sql = Sql & " and Descrição like '%" & txtDescricao & "%'"
        Selecao = Selecao & " and {VIS_EDIFICIO.DESCRIÇÃO} LIKE '*" & txtDescricao & "*'"
    End If
    If cboTipLogra.ListIndex <> -1 Then
        Sql = Sql & " and CodTipoLogra = '" & cboTipLogra.Coluna(0).Valor & "'"
        Selecao = Selecao & " and {VIS_EDIFICIO.CodTipoLogra} = " & cboTipLogra.Coluna(0).Valor
    End If
    
    If CboLogradouro.ListIndex <> -1 Then
        Sql = Sql & " and codlogradouro = '" & CboLogradouro.Coluna(0).Valor & "'"
        Selecao = Selecao & " and {VIS_EDIFICIO.codlogradouro} = " & CboLogradouro.Coluna(0).Valor
    End If
    
    If txtNumero <> "" Then
        Sql = Sql & " and Número = '" & txtNumero & "'"
        Selecao = Selecao & " and {VIS_EDIFICIO.Número} = '" & txtNumero & "'"
    End If
    
    If cboBairro.ListIndex <> -1 Then
        Sql = Sql & " and CodBairro = '" & cboBairro.Coluna(0).Valor & "'"
        Selecao = Selecao & " and {VIS_EDIFICIO.CodBairro} = " & cboBairro.Coluna(0).Valor
    End If
    grid.Preencher Bdados, Sql, 1000, 4000, 3000, 1000, 3000, 0, 0, 0
    
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdExcluir_Click()
    If Confirma("Confirma a exclusão do registro selecionado?", "Aviso") = True Then
        If Bdados.DeletaDados("TAB_EDIFICIO", "TED_COD_EDIFICIO = " & txtCodigo) Then
            Util.Avisa "Registro excluído com sucesso."
            cmdLimpar_Click
            cmdBuscar_Click
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim Rpt As New VSRelatorio
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path & "\Tedificio.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Selecao = Selecao
        .Visualizar
    End With
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Campos = "TED_COD_EDIFICIO,TED_DESCRICAO,TED_TBA_COD_BAIRRO,TED_TLG_COD_LOGRADOURO,TED_NUMERO," & _
        "TED_TLO_COD_LOTEAMENTO,TED_LOTE,TED_QUADRA,TED_SECAO,TED_AREA_LOTE,TED_AREA_CONSTRUIDA,TED_TTL_COD_TIPO_LOGRA"
    Valores = Bdados.PreparaValor(txtCodigo, txtDescricao, cboBairro.Coluna(0).Valor, CboLogradouro.Coluna(0).Valor, _
        txtNumero, Nvl(CStr(cboLoteamento.Coluna(0).Valor), 0), txtLote, txtQuadra, txtSecao, _
        txtAreaLote, txtAreaConstruida, cboTipLogra.Coluna(0).Valor)
    If Bdados.GravaDados("TAB_EDIFICIO", Valores, Campos, "TED_COD_EDIFICIO = " & txtCodigo) Then
        Informa "Transação completada."
        cmdLimpar_Click
        cmdBuscar_Click
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    cboTipLogra.Preencher Bdados, "Select TTL_COD_TIP_LOGR, TTL_NOME  From Tab_Tipo_Logr order by ttl_nome asc", 1
    cboBairro.Preencher Bdados, "select tba_cod_bairro,tba_nome from tab_bairro where tba_tmu_cod_municipio = " & Aplicacoes.Codigo_Municipio, 1
    CboLogradouro.Preencher Bdados, "Select tlg_cod_logradouro,tlg_nome  from tab_logradouro ", 1
    cboLoteamento.Preencher Bdados, "Select TLO_COD_LOTEAMENTO,TLO_DESCRICAO FROM TAB_LOTEAMENTO", 1
    If Bdados.AbreTabela("select max(ted_cod_edificio) as Total from tab_edificio") Then
        If Not IsNull(Bdados.Tabela(0)) Then
            txtCodigo = Bdados.Tabela(0) + 1
        Else
            txtCodigo = 1
        End If
    End If
    txtCodigo.Enabled = False
End Sub
Private Sub grid_DblClick()
    If grid.ListItems.Count >= 1 Then
        txtCodigo = grid.SelectedItem
        txtCodigo_LostFocus
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    Dim Rs As VSRecordset
    Dim Sql As String
    
    If txtCodigo = "" Then Exit Sub
    
    Sql = "Select * from tab_edificio where TED_COD_EDIFICIO = '" & txtCodigo & "'"
   If Bdados.AbreTabela(Sql, Rs, Dinamico) Then
        txtDescricao = Rs.Fields("TED_DESCRICAO")
        
        'pego o tipo de logradouro
        If Bdados.AbreTabela("select tlg_ttl_cod_tip_logr from tab_logradouro where tlg_cod_logradouro = '" & Rs.Fields("TED_TLG_COD_LOGRADOURO") & "'") Then
            cboTipLogra.SetarLinha Bdados.Tabela(0), 0
        End If
        cboBairro.SetarLinha Rs.Fields("TED_TBA_COD_BAIRRO")
        txtNumero = "" & Rs.Fields("TED_NUMERO")
        cboTipLogra.SetarLinha "" & Rs.Fields("TED_TTL_COD_TIPO_LOGRA")
        CboLogradouro.SetarLinha Rs.Fields("TED_TLG_COD_LOGRADOURO")
        txtCodigo.Enabled = False
        cboLoteamento.SetarLinha "" & Rs!TED_TLO_COD_LOTEAMENTO, 0
        txtLote = "" & Rs!TED_LOTE
        txtQuadra = "" & Rs!TED_QUADRA
        txtSecao = "" & Rs!TED_SECAO
        txtAreaConstruida = "" & Rs!TED_AREA_CONSTRUIDA
        txtAreaLote = "" & Rs!TED_AREA_LOTE
   Else
        txtCodigo.Enabled = True
        txtNumero = ""
        txtDescricao = ""
        cboBairro.ListIndex = -1
        CboLogradouro.ListIndex = -1
        cboLoteamento.ListIndex = -1
        txtLote = ""
        txtQuadra = ""
        txtSecao = ""
        txtAreaConstruida = ""
        txtAreaLote = ""
   End If
    
End Sub
