VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form form100MaioresContribuintes 
   BackColor       =   &H00FBEDE8&
   Caption         =   "Contribuintes"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   ForeColor       =   &H00000000&
   Icon            =   "formObrig100MaioresContribuintes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8985
   ScaleMode       =   0  'User
   ScaleWidth      =   12943.47
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView Lst1 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   13785
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Height          =   645
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1138
      Formulario      =   "Relaçao dos Maiores Contribuintes"
      Descricao       =   ""
      Icone           =   "formObrig100MaioresContribuintes.frx":08CA
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   8475
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL cmdRelatorio 
         Height          =   375
         Left            =   9000
         TabIndex        =   4
         Top             =   120
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   661
         Caption         =   "&Relatorio"
         Acao            =   4
         CorBorda        =   16711680
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10320
         TabIndex        =   3
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFundo        =   16777088
      End
   End
End
Attribute VB_Name = "form100MaioresContribuintes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CabLancto()
      
      Lst1.ListItems.Clear
      Lst1.ColumnHeaders.Clear
      Lst1.ColumnHeaders.Add , , "Posição", 1000
      Lst1.ColumnHeaders.Add , , "Contribuinte", 5500
      Lst1.ColumnHeaders.Add , , "SubTotal", 1500
      Lst1.ColumnHeaders.Add , , "Total", 1500
      Lst1.ColumnHeaders.Add , , "Perc (%)", 1500
      
      
      Lst1.ColumnHeaders.Item(3).Alignment = lvwColumnRight
      Lst1.ColumnHeaders.Item(4).Alignment = lvwColumnRight
      Lst1.ColumnHeaders.Item(5).Alignment = lvwColumnRight
      
      Lst1.View = lvwReport
   
End Sub

Sub carregaLista()

    Dim Bd As Connection
    abreConexao Bd
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.PageSize = 27
     
    Set rst = Bd.Execute("exec sp_curva_abc_calculo_final")
    
    Dim vlrPerc As Currency
     CabLancto
     While Not rst.EOF
   
        vlrPerc = (carregaCampo(rst!SubTotal) / carregaCampo(rst!Total)) * 100
   
        Set Lst = Lst1.ListItems.Add(, , carregaCampo(rst!Posicao))
        Lst.SubItems(1) = carregaCampo(rst!Nome)
        Lst.SubItems(2) = Format(carregaCampo(rst!SubTotal), "##,##0.00")
        Lst.SubItems(3) = Format(carregaCampo(rst!Total), "##,##0.00")
        Lst.SubItems(4) = Format(vlrPerc, "##,##0.00")
        rst.MoveNext
    
    Wend
    
    fechaConexao Bd
End Sub

Private Sub cmdRelatorio_Click()
    Dim CondRelatorio As String
    Dim FORMULA As String
    Dim Inscricao As String
    Dim relCons As New relObrig100Maiores
    configRelatorio relCons, "exec sp_curva_abc_calculo_final"
    relCons.Show
End Sub

Private Sub cmdSair_Click()
Unload Me

End Sub

Private Sub Form_Load()
carregaLista
End Sub

