VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TNAV301 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   1138
      Icone           =   "TNAV301.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      TabIndex        =   18
      Top             =   5760
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   1058
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7845
         TabIndex        =   7
         Top             =   135
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   6645
         TabIndex        =   6
         Top             =   135
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9030
         TabIndex        =   8
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   4995
      Left            =   75
      TabIndex        =   9
      Top             =   705
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   8811
      Caption         =   "Nota Fiscal Avulsa"
      Descricao       =   "Eliminação de Notas Fiscais"
      corFaixa        =   16711680
      Icone           =   "TNAV301.frx":031A
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin ActiveTabs.SSActiveTabs tabNota 
         Height          =   4095
         Left            =   120
         TabIndex        =   5
         Top             =   750
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   7223
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
         Tabs            =   "TNAV301.frx":0BF4
         Images          =   "TNAV301.frx":0C80
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
            Height          =   3675
            Left            =   30
            TabIndex        =   12
            Top             =   30
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   6482
            _Version        =   131082
            TabGuid         =   "TNAV301.frx":1C40
            Begin VTOcx.txtVISUAL txtPeriodo 
               Height          =   480
               Left            =   4545
               TabIndex        =   14
               Top             =   3165
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   847
               Caption         =   "Periodo"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtTotalNota 
               Height          =   480
               Left            =   5770
               TabIndex        =   15
               Tag             =   "Total da Nota"
               Top             =   3165
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   847
               Caption         =   "Total da Nota"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtBaseCalc 
               Height          =   480
               Left            =   7235
               TabIndex        =   16
               Tag             =   "Base de Cálculo"
               Top             =   3165
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   847
               Caption         =   "Base de Cálculo"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtISS 
               Height          =   480
               Left            =   8700
               TabIndex        =   17
               Tag             =   "ISS"
               Top             =   3165
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   847
               Caption         =   "ISS Devido"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.grdVISUAL grdItem 
               Height          =   3300
               Left            =   30
               TabIndex        =   13
               Top             =   75
               Width           =   9780
               _ExtentX        =   17251
               _ExtentY        =   5821
               CorBorda        =   32768
               Caption         =   "Itens"
               CorTitulo       =   32768
               CorCaption      =   16777215
               CorDica         =   32768
               OcultarRodape   =   -1  'True
            End
         End
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
            Height          =   3675
            Left            =   30
            TabIndex        =   10
            Top             =   30
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   6482
            _Version        =   131082
            TabGuid         =   "TNAV301.frx":1C68
            Begin VTOcx.grdVISUAL grdNota 
               Height          =   2670
               Left            =   60
               TabIndex        =   4
               Top             =   1065
               Width           =   9735
               _ExtentX        =   17171
               _ExtentY        =   4710
               CorBorda        =   16711680
               Caption         =   "Notas Fiscais"
               CorTitulo       =   16711680
               CorCaption      =   16777215
               CorDica         =   16711680
               OcultarRodape   =   -1  'True
            End
            Begin VTOcx.fraVISUAL fraVISUAL1 
               Height          =   900
               Left            =   75
               TabIndex        =   11
               Top             =   90
               Width           =   9705
               _ExtentX        =   17119
               _ExtentY        =   1588
               Altura          =   1905
               Caption         =   " Opções de Busca"
               CorTexto        =   16777215
               CorFaixa        =   16711680
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.txtVISUAL txtBoleto 
                  Height          =   480
                  Left            =   1550
                  TabIndex        =   1
                  Top             =   330
                  Width           =   1400
                  _ExtentX        =   2461
                  _ExtentY        =   847
                  Caption         =   "Boleto"
                  Text            =   ""
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.cmdVISUAL cmdBuscar 
                  Height          =   345
                  Left            =   8580
                  TabIndex        =   3
                  Top             =   480
                  Width           =   1065
                  _ExtentX        =   1879
                  _ExtentY        =   609
                  Caption         =   "&Buscar"
                  Acao            =   5
                  CorBorda        =   16711680
                  CorFrente       =   0
                  CorFundo        =   16777088
               End
               Begin VTOcx.txtVISUAL txtNome 
                  Height          =   480
                  Left            =   3000
                  TabIndex        =   2
                  Top             =   330
                  Width           =   5500
                  _ExtentX        =   9710
                  _ExtentY        =   847
                  Caption         =   "Nome / Motivo Cancelamento"
                  Text            =   ""
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtNumNota 
                  Height          =   480
                  Left            =   100
                  TabIndex        =   0
                  Top             =   330
                  Width           =   1400
                  _ExtentX        =   2461
                  _ExtentY        =   847
                  Caption         =   "Nº Nota Fiscal"
                  Text            =   ""
                  AlinhamentoRotulo=   1
               End
            End
         End
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "TNAV301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NotaAvulsa As cNotaAvulsa
Dim ItemNota As cItemNotaAvulsa

Private Sub cmdBuscar_Click()
    
    NotaAvulsa.PreencherGrid grdNota, txtNumNota, txtNome
End Sub

Private Sub cmdEnter_Click()
            SendKeys "{Tab}"
End Sub

Private Sub cmdExcluir_Click()
    On Error GoTo TRATA
    If grdNota.SelectedItem Is Nothing Then Exit Sub
        'BCP
        If Confirma("Deseja realmente excluir a nota nº " & grdNota.SelectedItem & "?") Then
            If Len(txtBoleto) = 0 Then
                Mensagem "Informe o boleto desta nota"
                Exit Sub
            End If
            NotaAvulsa.Buscar (txtNumNota)
            If txtBoleto <> NotaAvulsa.CodPagamento Then
                Mensagem "O boleto informado não pertence a esta nota fiscal"
                Exit Sub
            End If
        End If
        'fim bcp
            NotaAvulsa.Excluir grdNota.SelectedItem, txtNome ' Apos a consulta o campo nome serve como a observacao
            'BCP ItemNota.Excluir grdNota.SelectedItem
            Avisa "Nota fiscal eliminada."
            cmdLimpar_Click
        
        Exit Sub
TRATA:
    Erro "Erro ao Excluir nota."
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    grdItem.ListItems.Clear
    grdNota.ListItems.Clear
    tabNota.Tabs(1).Selected = True
    txtNumNota.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set NotaAvulsa = New cNotaAvulsa
    Set ItemNota = New cItemNotaAvulsa
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set NotaAvulsa = Nothing
    Set ItemNota = Nothing
End Sub

Private Sub grdNota_dblclick()
    If grdNota.SelectedItem Is Nothing Then Exit Sub
    With NotaAvulsa
        If .Buscar(grdNota.SelectedItem) Then
            txtPeriodo = .Periodo
            txtTotalNota = .ValorNota
            txtBaseCalc = txtTotalNota - .Material
            txtISS = .ValorImposto
            If ItemNota.PreencherGrid(grdItem, grdNota.SelectedItem) = False Then
                Util.Avisa "Nota Fiscal sem Itens."
            End If
            tabNota.Tabs(2).Selected = True
        Else
            Util.Avisa "Nota não encontrada."
        End If
    End With
End Sub

