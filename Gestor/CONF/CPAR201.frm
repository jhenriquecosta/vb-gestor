VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form CPAR201 
   BackColor       =   &H00FFF5EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CPAR201.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   20
      Top             =   6210
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   5370
         TabIndex        =   21
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   0
      End
   End
   Begin ActiveTabs.SSActiveTabs tabGeral 
      Height          =   5370
      Left            =   60
      TabIndex        =   13
      Top             =   720
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   9472
      _Version        =   131082
      TabCount        =   3
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
      Tabs            =   "CPAR201.frx":08CA
      Images          =   "CPAR201.frx":097C
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4950
         Index           =   0
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   8731
         _Version        =   131082
         TabGuid         =   "CPAR201.frx":0FCF
         Begin VTOcx.txtVISUAL txtSistema 
            Height          =   315
            Left            =   1530
            TabIndex        =   1
            Top             =   4500
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtFaixaSistema 
            Height          =   315
            Left            =   60
            TabIndex        =   0
            Top             =   4500
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "Sistema"
            Text            =   ""
            Restricao       =   2
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdSalvarSistema 
            Height          =   405
            Left            =   5250
            TabIndex        =   2
            Top             =   4470
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdExcluirSistema 
            Height          =   405
            Left            =   5640
            TabIndex        =   3
            Top             =   4470
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.grdVISUAL grdSistemas 
            Height          =   4335
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   4339
            CorFundo        =   -2147483633
            Caption         =   "Sistemas"
            CorTitulo       =   4210688
            CorCaption      =   16777215
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4950
         Index           =   2
         Left            =   -99969
         TabIndex        =   16
         Top             =   30
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   8731
         _Version        =   131082
         TabGuid         =   "CPAR201.frx":0FF7
         Begin VTOcx.txtVISUAL txtValor 
            Height          =   315
            Left            =   1530
            TabIndex        =   9
            Top             =   4500
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtCodValor 
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   4500
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "Valor"
            Text            =   ""
            Restricao       =   2
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdSalvarValor 
            Height          =   405
            Left            =   5250
            TabIndex        =   10
            Top             =   4470
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdExcluirValor 
            Height          =   405
            Left            =   5640
            TabIndex        =   11
            Top             =   4470
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.grdVISUAL grdValores 
            Height          =   4335
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   4339
            CorFundo        =   -2147483633
            Caption         =   "Valores"
            CorTitulo       =   4210688
            CorCaption      =   16777215
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4950
         Index           =   1
         Left            =   -99969
         TabIndex        =   18
         Top             =   30
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   8731
         _Version        =   131082
         TabGuid         =   "CPAR201.frx":101F
         Begin VTOcx.txtVISUAL txtTabela 
            Height          =   315
            Left            =   1530
            TabIndex        =   5
            Top             =   4500
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtCodTabela 
            Height          =   315
            Left            =   60
            TabIndex        =   4
            Top             =   4500
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "Tabela"
            Text            =   ""
            Restricao       =   2
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdSalvarTabela 
            Height          =   405
            Left            =   5250
            TabIndex        =   6
            Top             =   4470
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdExcluirTabela 
            Height          =   405
            Left            =   5640
            TabIndex        =   7
            Top             =   4470
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.grdVISUAL grdTabelas 
            Height          =   4335
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   4339
            CorFundo        =   -2147483633
            Caption         =   "Tabelas"
            CorTitulo       =   4210688
            CorCaption      =   16777215
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   1138
      Icone           =   "CPAR201.frx":1047
   End
End
Attribute VB_Name = "CPAR201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intTabelaFaixas As Integer
Private intFaixaSistema As Integer
Private intCodTabela As Integer

Private Sub cmdExcluirSistema_Click()
    If ExcluirGeral(intTabelaFaixas, Util.Nvl(txtFaixaSistema, 0), txtSistema) Then
        PrepararSistema
    End If
End Sub

Private Sub cmdExcluirTabela_Click()
    If ExcluirGeral(Util.Nvl(txtCodTabela, 0), 0, txtTabela) Then
        PrepararTabela
    End If
End Sub

Private Sub cmdExcluirValor_Click()
    If ExcluirGeral(intCodTabela, Util.Nvl(txtCodValor, 0), txtValor) Then
        PrepararValor
    End If
End Sub

Private Sub cmdSalvarSistema_Click()
    If GravarGeral(intTabelaFaixas, Util.Nvl(txtFaixaSistema, 0), txtSistema) Then
        PrepararSistema
    End If
End Sub

Private Sub cmdSalvarTabela_Click()
    If GravarGeral(Util.Nvl(txtCodTabela, 0), 0, txtTabela) Then
        PrepararTabela
    End If
End Sub

Private Sub cmdSalvarValor_Click()
    If GravarGeral(intCodTabela, Util.Nvl(txtCodValor, 0), txtValor) Then
        PrepararValor
    End If
End Sub

Private Sub Form_Load()
    SetarTabelaFaixas
    PreencherSistemas
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub grdSistemas_Click()
    If Not grdSistemas.SelectedItem Is Nothing Then
        With grdSistemas.SelectedItem
            intFaixaSistema = .Text
            txtFaixaSistema = .Text
            txtSistema = .SubItems(1)
            grdTabelas.Caption = "Tabelas : " & .SubItems(1)
            PreencherTabelas CInt(.Text)
        End With
    End If
End Sub

Private Sub grdSistemas_DblClick()
    tabGeral.Tabs(2).Selected = True
    txtCodTabela.SetFocus
End Sub

Private Sub grdTabelas_Click()
    If Not grdTabelas.SelectedItem Is Nothing Then
        With grdTabelas.SelectedItem
            intCodTabela = .Text
            txtCodTabela = .Text
            txtTabela = .SubItems(1)
            grdValores.Caption = "Valores : " & .SubItems(1)
            PreencherValores CInt(.Text)
        End With
    End If
End Sub

Private Sub grdTabelas_DblClick()
    tabGeral.Tabs(3).Selected = True
    txtCodValor.SetFocus
End Sub

Private Sub grdValores_Click()
    If Not grdValores.SelectedItem Is Nothing Then
        With grdValores.SelectedItem
            txtCodValor = .Text
            txtValor = .SubItems(1)
        End With
    End If
End Sub

Private Sub PreencherSistemas()
    Dim sql As String
    
    sql = "SELECT TGE_CODIGO as Faixa, TGE_NOME AS Sistema" & _
            " FROM TAB_GERAL" & _
            " WHERE TGE_CODIGO>0 AND" & _
                " TGE_TIPO=" & intTabelaFaixas & _
            " ORDER BY TGE_CODIGO"
    grdSistemas.Preencher Bdados, sql
End Sub

Private Sub PreencherTabelas(FaixaSistema As Integer)
    Dim sql As String
    
    sql = "SELECT TGE_TIPO as Codigo, TGE_NOME AS Tabela" & _
            " FROM TAB_GERAL" & _
            " WHERE TGE_CODIGO=0 AND" & _
                " TGE_TIPO>=" & FaixaSistema & " AND " & _
                " TGE_TIPO<" & FaixaSistema + 100 & _
            " ORDER BY TGE_TIPO"
    grdTabelas.Preencher Bdados, sql
End Sub

Private Sub PreencherValores(CodTabela As Integer)
    Dim sql As String
    
    sql = "SELECT TGE_CODIGO as Codigo, TGE_NOME AS Valor" & _
            " FROM TAB_GERAL" & _
            " WHERE TGE_CODIGO>0 AND" & _
                " TGE_TIPO=" & CodTabela & _
            " ORDER BY TGE_TIPO"
    grdValores.Preencher Bdados, sql
End Sub

Private Sub SetarTabelaFaixas()
    Dim sql As String
    
    sql = "SELECT TGE_TIPO" & _
        " FROM TAB_GERAL" & _
        " WHERE TGE_CODIGO=0 AND" & _
            " TGE_NOME='PROPRIETARIO TABELA'"
    intTabelaFaixas = Bdados.BuscaCodigo(sql)
End Sub

Private Function GravarGeral(Tipo As Integer, Codigo As Integer, Nome As String) As Boolean
    Dim Campos As String, Valores As String
    
    If Trim$(Nome) = "" Then Exit Function
    
    Campos = "TGE_TIPO,TGE_CODIGO,TGE_NOME"
    Valores = Bdados.PreparaValor(Tipo, Codigo, Nome)
    GravarGeral = Bdados.GravaDados("TAB_GERAL", Valores, Campos, "TGE_TIPO=" & Tipo & " AND TGE_CODIGO=" & Codigo)
End Function

Private Function ExcluirGeral(Tipo As Integer, Codigo As Integer, Nome As String) As Boolean
    If Trim$(Nome) = "" Then Exit Function
    
    If Util.Confirma("Apagar " & Nome & " ?") Then
        ExcluirGeral = Bdados.DeletaDados("TAB_GERAL", "TGE_TIPO=" & Tipo & " AND TGE_CODIGO=" & Codigo)
    End If
End Function

Private Sub PrepararSistema()
    PreencherSistemas
    grdTabelas.Preencher Bdados, ""
    grdValores.Preencher Bdados, ""
    txtFaixaSistema = ""
    txtSistema = ""
    txtFaixaSistema.SetFocus
End Sub

Private Sub PrepararTabela()
    PreencherTabelas intFaixaSistema
    grdValores.Preencher Bdados, ""
    txtCodTabela = ""
    txtTabela = ""
    txtCodTabela.SetFocus
End Sub

Private Sub PrepararValor()
    PreencherValores intCodTabela
    txtCodValor = ""
    txtValor = ""
    txtCodValor.SetFocus
End Sub
