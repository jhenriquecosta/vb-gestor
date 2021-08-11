VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TIMP107 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs tabTributo 
      Align           =   1  'Align Top
      Height          =   4230
      Left            =   0
      TabIndex        =   3
      Top             =   645
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   7461
      _Version        =   131082
      TabCount        =   1
      CaptionAlignment=   1
      CaptionOrientation=   1
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "TIMP107.frx":0000
      Images          =   "TIMP107.frx":005D
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3810
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   6720
         _Version        =   131082
         TabGuid         =   "TIMP107.frx":09A5
         Begin VTOcx.grdVISUAL grdVISUAL1 
            Height          =   2865
            Left            =   75
            TabIndex        =   5
            Top             =   840
            Width           =   6420
            _ExtentX        =   11324
            _ExtentY        =   5054
            CorTitulo       =   16711680
         End
         Begin VTOcx.txtVISUAL txtData 
            Height          =   315
            Left            =   1080
            TabIndex        =   1
            Tag             =   "Data"
            Top             =   390
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            Caption         =   "Data"
            Text            =   ""
            Formato         =   0
            Restricao       =   2
         End
         Begin VTOcx.txtVISUAL txtValor 
            Height          =   315
            Left            =   4560
            TabIndex        =   2
            Tag             =   "Valor"
            Top             =   390
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            Caption         =   "Valor"
            Text            =   ""
            Restricao       =   3
         End
         Begin VTOcx.cboVISUAL cboUnidade 
            Height          =   315
            Left            =   270
            TabIndex        =   0
            Tag             =   "Unidade Fiscal"
            Top             =   30
            Width           =   6210
            _ExtentX        =   10954
            _ExtentY        =   556
            Caption         =   "Unidade Fiscal"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   1138
      Icone           =   "TIMP107.frx":09CD
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   7
      Top             =   4935
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   10
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Novo"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   9
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   5415
         TabIndex        =   8
         Top             =   120
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
End
Attribute VB_Name = "TIMP107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Tipo As cTipoTributo
Private Categoria As cCategoriaTributo


Dim CodCategoria As Double
Dim CodReceita As Double

Private Sub cboUnidade_Click()
    MontaGrid
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim rs As VSRecordset
    Select Case Index
        Case 0
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Valores = Bdados.PreparaValor(txtData, Bdados.Converte(txtValor, TCDuplo), cboUnidade.Coluna(1).Valor, txtValor)
            Campos = "TMO_DATA,TMO_VALOR,TMO_UNIDADE,tmo_valor_mostra"
            Call Bdados.GravaDados("TAB_MONETARIA", Valores, Campos, "TMO_UNIDADE = " & cboUnidade.Coluna(1).Valor _
                    & " AND TMO_DATA = " & Bdados.Converte(txtData, TCDataHora))
'            grdTipos.Colunas(2).Tipo = tipTexto
            'grdTipos.Preencher Bdados, "Select TMO_DATA as Data,TMO_VALOR_MOSTRA as Valor from Tab_monetaria"
            Call Util.Informa("Transação Completada.")
            MontaGrid
            txtData.Enabled = True
            txtData.SetFocus
        Case 1
            Unload Me
        Case 2
            Edita.LimpaCampos Me
            txtData.Enabled = True
            txtData.SetFocus
    End Select
End Sub
Private Sub MontaGrid()
    Dim Sql As String
    Sql = "SELECT TMO_DATA as Data ,convert(varchar,TMO_VALOR,103) as Valor  FROM TAB_MONETARIA where 1 = 1"
     If cboUnidade.ListIndex <> -1 Then
        Sql = Sql & " and tmo_unidade = '" & cboUnidade.Coluna(1).Valor & "'"
     End If
    grdVISUAL1.Preencher Bdados, Sql
    
End Sub
Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    Set Tipo = New cTipoTributo
    Set Categoria = New cCategoriaTributo
    
    cboUnidade.PreencherGeral Bdados, "UNIDADE FISCAL"
End Sub


Private Sub grdVISUAL1_DblClick()
    If grdVISUAL1.ListItems.Count >= 1 Then
        txtData = grdVISUAL1.SelectedItem
        txtValor = Edita.TrocaPic(grdVISUAL1.SelectedItem.SubItems(1), ".", ",")
    End If
End Sub
