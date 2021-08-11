VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form THOM401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THOM401"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   ControlBox      =   0   'False
   Icon            =   "THOM401.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   6
      Top             =   6555
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL CmdImprimir 
         Height          =   345
         Left            =   4920
         TabIndex        =   12
         Top             =   150
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "Imprimir"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   6030
         TabIndex        =   3
         Top             =   150
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   609
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   7125
         TabIndex        =   4
         Top             =   150
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   8115
         TabIndex        =   5
         Top             =   150
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1138
      Icone           =   "THOM401.frx":08CA
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1830
      Left            =   30
      TabIndex        =   9
      Top             =   690
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   3228
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   825
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   556
         Caption         =   "Razão"
         Text            =   ""
         Enabled         =   0   'False
         CorFundo        =   -2147483644
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   315
         Left            =   750
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1185
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   556
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         CorFundo        =   -2147483644
      End
      Begin VTOcx.txtVISUAL txtInscricao 
         Height          =   300
         Left            =   765
         TabIndex        =   0
         Tag             =   "Inscrição Cadastral"
         Top             =   465
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   529
         Caption         =   "Inscricao"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   -2147483644
         MaxLen          =   20
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3075
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   450
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
      End
   End
   Begin VTOcx.grdVISUAL grdInfra 
      Height          =   1890
      Left            =   15
      TabIndex        =   11
      Top             =   4845
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   3334
      CorBorda        =   32768
      Caption         =   "Infrações"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.grdVISUAL grdAutos 
      Height          =   2535
      Left            =   15
      TabIndex        =   8
      Top             =   2565
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   4471
      CorBorda        =   32768
      Caption         =   "Autos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
End
Attribute VB_Name = "THOM401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscar_Click()

 On Error GoTo TrataErro
        
    Screen.MousePointer = 11
    preencherGridAuto
    Screen.MousePointer = 0
    
    Exit Sub
TrataErro:
    Util.Erro Err.Description
    Exit Sub
    Resume

End Sub

Private Sub CmdImprimir_Click()
    Dim Rpt As New VSRelatorio
    Dim CondRelatorio As String
    
    With Rpt
        
        If Not .DefinirArquivo(Bdados, App.Path & "\TAutoInfracao.rpt") Then Exit Sub
        .Selecao = "{TAB_AUTO_INFRACAO.TAI_COD_AUTO} = '" & grdAutos.SelectedItem & "'"
        .Formulas "VT_ESTADO", Temp.PegaParametro(Bdados, "ESTADO")
        .Formulas "VT_PREFEITURA", Temp.PegaParametro(Bdados, "CLIENTE")
        .Formulas "VT_SECRETARIA", Temp.PegaParametro(Bdados, "SEMFAZ")
        .Formulas "VT_SETOR", Temp.PegaParametro(Bdados, "GAF")
        .Visualizar
    
    End With
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdAutos.ListItems.Clear
    grdInfra.ListItems.Clear
End Sub


Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscricao
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
'AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
     cabVISUAL.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
       
End Sub





Private Sub grdAutos_DblClick()
    If grdAutos.ListItems.Count >= 1 Then
         PreencherGridInfra
    End If
End Sub

Private Sub txtInscricao_LostFocus()
    If txtInscricao = "" Then Exit Sub
    txtInscricao = BuscaContribuinte(txtInscricao, txtRazao, txtEndereco)
End Sub
Private Sub preencherGridAuto()
    Dim Sql As String
    Dim Inscricao As String
    
        Sql = "select TAI_COD_AUTO as Número,"
        Sql = Sql & "TAI_TCI_IM as IM,"
        Sql = Sql & "TAI_DATA_VENCIMENTO as Vencimento,"
        Sql = Sql & "TAI_VALOR_AUTO as Valor,"
        Sql = Sql & "TAI_VALOR_AGRAVANTE as Agravante,"
        Sql = Sql & "TAI_VALOR_TOTAL as Total,"
        Sql = Sql & "TAI_OBS as OBS "
        Sql = Sql & "FROM  TAB_AUTO_INFRACAO WHERE 1 =1"
        If txtInscricao <> "" Then
            Sql = Sql & " AND TAI_TCI_IM = '" & txtInscricao & "'"
        End If
    
      If Not grdAutos.Preencher(Bdados, Sql) Then
            Util.Avisa "Consulta sem resultados."
        End If
End Sub

Private Sub PreencherGridInfra()
    Dim Sql As String
    
    Sql = " SELECT TIN_COD_INFRACAO AS Infração,"
    Sql = Sql & " tin_descricao_infracao as Descrição,"
    Sql = Sql & " tin_valor_ufm As VALOR, tin_artigo As Artigo"
    Sql = Sql & " From Tab_Infracao, tab_infracao_auto"
    Sql = Sql & " Where TIA_INFRACAO = TIN_COD_INFRACAO"
    Sql = Sql & " and tia_cod_auto   = '" & grdAutos.SelectedItem & "'"
    
    If Not grdInfra.Preencher(Bdados, Sql) Then
            Util.Avisa "Consulta sem resultados."
        End If
    
End Sub
