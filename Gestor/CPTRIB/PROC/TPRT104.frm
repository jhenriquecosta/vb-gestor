VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TPRT104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TPRT104"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "TPRT104.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   9
      Top             =   6150
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   714
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   315
         Left            =   5160
         TabIndex        =   8
         Top             =   75
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
      TabIndex        =   10
      Top             =   0
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   1138
      Icone           =   "TPRT104.frx":08CA
   End
   Begin ActiveTabs.SSActiveTabs tabGeral 
      Height          =   5370
      Left            =   60
      TabIndex        =   11
      Top             =   705
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   9472
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "TPRT104.frx":0BE4
      Images          =   "TPRT104.frx":0C62
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4950
         Index           =   2
         Left            =   -99969
         TabIndex        =   12
         Top             =   30
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   8731
         _Version        =   131082
         TabGuid         =   "TPRT104.frx":12B5
         Begin VTOcx.txtVISUAL txtValor 
            Height          =   315
            Left            =   1530
            TabIndex        =   5
            Top             =   4500
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtCodValor 
            Height          =   315
            Left            =   60
            TabIndex        =   4
            Top             =   4500
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "Valor"
            Text            =   ""
            Restricao       =   2
         End
         Begin VTOcx.cmdVISUAL cmdSalvarValor 
            Height          =   405
            Left            =   5250
            TabIndex        =   6
            Top             =   4470
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   32768
         End
         Begin VTOcx.cmdVISUAL cmdExcluirValor 
            Height          =   405
            Left            =   5640
            TabIndex        =   7
            Top             =   4470
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   32768
         End
         Begin VTOcx.grdVISUAL grdValores 
            Height          =   4335
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   4339
            CorBorda        =   32768
            Caption         =   "Valores"
            CorTitulo       =   32768
            CorCaption      =   16777215
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4950
         Index           =   1
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   8731
         _Version        =   131082
         TabGuid         =   "TPRT104.frx":12DD
         Begin VTOcx.txtVISUAL txtTabela 
            Height          =   315
            Left            =   1530
            TabIndex        =   1
            Top             =   4500
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtCodTabela 
            Height          =   315
            Left            =   60
            TabIndex        =   0
            Top             =   4500
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "Tabela"
            Text            =   ""
            Restricao       =   2
         End
         Begin VTOcx.cmdVISUAL cmdSalvarTabela 
            Height          =   405
            Left            =   5250
            TabIndex        =   2
            Top             =   4470
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   32768
         End
         Begin VTOcx.cmdVISUAL cmdExcluirTabela 
            Height          =   405
            Left            =   5640
            TabIndex        =   3
            Top             =   4470
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   32768
         End
         Begin VTOcx.grdVISUAL grdTabelas 
            Height          =   4335
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   4339
            CorBorda        =   32768
            Caption         =   "Tabelas"
            CorTitulo       =   32768
            CorCaption      =   16777215
         End
      End
   End
End
Attribute VB_Name = "TPRT104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcluirTabela_Click()
    Dim condicao As String
    If grdValores.ListItems.Count < 1 Then Exit Sub
    condicao = "TPP_TIPO_PARAMETRO = " & txtCodTabela
    If txtCodTabela <> "" Then
        If Confirma("Deseja excluir registro?", "Excluir?") Then
            If Bdados.DeletaDados("TAB_PARAMETRO_PROTOCOLO", condicao) Then
                Avisa "Registro Excluidos com Sucesso"
                LmpCampoTab
                carregaTabela
            End If
        End If
    Else
        Avisa "Selecione um Registro"
    End If
End Sub

Private Sub cmdExcluirValor_Click()
    Dim condicao As String
    If grdValores.ListItems.Count < 1 Then Exit Sub
    condicao = "TPP_TIPO_PARAMETRO = " & grdTabelas.SelectedItem & " and TPP_CODIGO_PARAMETRO = " & txtCodValor
    If txtCodValor <> "" Then
        If Confirma("Deseja excluir registro?", "Excluir?") Then
            If Bdados.DeletaDados("TAB_PARAMETRO_PROTOCOLO", condicao) Then
                Avisa "Registro Excluidos com Sucesso"
                LmpValor
                carregaValor
            End If
        End If
    Else
        Avisa "Selecione um Registro"
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
    carregaTabela
End Sub
Private Sub cmdSalvarTabela_Click()
    Dim campos As String
    Dim valores As String
    Dim condicao As String
    Dim Codigo As String
    
    If txtCodTabela = "" Or txtTabela = "" Then Exit Sub
    
    campos = " TPP_TIPO_PARAMETRO,TPP_CODIGO_PARAMETRO,TPP_NOME_PARAMETRO"
    valores = Bdados.PreparaValor(txtCodTabela, 0, txtTabela)
    condicao = "TPP_TIPO_PARAMETRO = " & txtCodTabela
    If Bdados.GravaDados("TAB_PARAMETRO_PROTOCOLO", valores, campos, condicao) Then
        'Avisa "Dados Salvos com Sucesso"
        carregaTabela
        LmpCampoTab
        txtCodTabela.SetFocus
    End If
End Sub

Private Sub cmdSalvarValor_Click()
    Dim campos As String
    Dim valores As String
    Dim condicao As String
    Dim Codigo As String
    
    If txtCodValor = "" Or txtValor = "" Then Exit Sub
    
    campos = " TPP_TIPO_PARAMETRO,TPP_CODIGO_PARAMETRO,TPP_NOME_PARAMETRO"
    valores = Bdados.PreparaValor(grdTabelas.SelectedItem, txtCodValor, txtValor)
    condicao = "TPP_TIPO_PARAMETRO = " & grdTabelas.SelectedItem & " and TPP_CODIGO_PARAMETRO = " & txtCodValor
    If Bdados.GravaDados("TAB_PARAMETRO_PROTOCOLO", valores, campos, condicao) Then
        'Avisa "Dados Salvos com Sucesso"
        carregaValor
        LmpValor
        txtCodValor.SetFocus
    End If
End Sub

Private Sub carregaTabela()
    Dim Sql As String
    Sql = "SELECT TPP_TIPO_PARAMETRO as Código, TPP_NOME_PARAMETRO as Descrição "
    Sql = Sql & " FROM TAB_PARAMETRO_PROTOCOLO"
    Sql = Sql & " where TPP_CODIGO_PARAMETRO = 0"
    Sql = Sql & " ORDER BY TPP_TIPO_PARAMETRO"
    
    grdTabelas.Preencher Bdados, Sql
End Sub
Private Sub LmpCampoTab()
    txtCodTabela = ""
    txtTabela = ""
End Sub
Private Sub carregaValor()
    Dim Sql As String
    Sql = " SELECT  TPP_CODIGO_PARAMETRO AS Código, TPP_NOME_PARAMETRO AS Descrição"
    Sql = Sql & " From TAB_PARAMETRO_PROTOCOLO"
    Sql = Sql & " Where (TPP_CODIGO_PARAMETRO <> 0)"
    Sql = Sql & " and  tpp_tipo_parametro =  " & grdTabelas.SelectedItem
    Sql = Sql & " ORDER BY TPP_CODIGO_PARAMETRO "
    
    grdValores.Preencher Bdados, Sql
End Sub
Private Sub LmpValor()
    txtCodValor = ""
    txtValor = ""
End Sub

Private Sub grdTabelas_dblClick()
    If grdTabelas.ListItems.Count < 1 Then Exit Sub
    carregaValor
    tabGeral.Tabs(2).Selected = True
    txtCodValor.SetFocus
    LmpValor
End Sub

Private Sub grdTabelas_Click()
    If grdTabelas.ListItems.Count < 1 Then Exit Sub
    txtCodTabela = grdTabelas.SelectedItem
    txtTabela = grdTabelas.SelectedItem.SubItems(1)
End Sub
Private Sub grdValores_Click()
    If grdValores.ListItems.Count < 1 Then Exit Sub
    txtCodValor = grdValores.SelectedItem
    txtValor = grdValores.SelectedItem.SubItems(1)
End Sub

