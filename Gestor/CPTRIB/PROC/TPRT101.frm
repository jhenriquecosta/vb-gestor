VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TPRT101 
   Caption         =   "TPRT101"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1138
      Formulario      =   "Ordem de Serviço"
      Descricao       =   "Lançamento das Ações de Fiscalização"
      Icone           =   "TPRT101.frx":0000
   End
   Begin ActiveTabs.SSActiveTabs tabEtapa 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11245
      _Version        =   131082
      TabCount        =   4
      Tabs            =   "TPRT101.frx":1792
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   5985
         Left            =   -99969
         TabIndex        =   5
         Top             =   360
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   10557
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":1873
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5985
         Left            =   -99969
         TabIndex        =   4
         Top             =   360
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   10557
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":189B
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   5985
         Left            =   -99969
         TabIndex        =   3
         Top             =   360
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   10557
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":18C3
         Begin VTOcx.cmdVISUAL cmdVISUAL1 
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   120
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   661
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5985
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   10557
         _Version        =   131082
         TabGuid         =   "TPRT101.frx":18EB
         Begin VTOcx.fraVISUAL fra 
            Height          =   1065
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   1879
            Altura          =   1905
            Caption         =   " Filtro"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboStatus 
               Height          =   510
               Left            =   4920
               TabIndex        =   13
               Tag             =   "C"
               Top             =   360
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   900
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cmdVISUAL cmdBuscar 
               Height          =   375
               Left            =   8160
               TabIndex        =   11
               Top             =   480
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   661
               Caption         =   ""
               Acao            =   5
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cboVISUAL cboFiscal 
               Height          =   510
               Left            =   1440
               TabIndex        =   9
               Tag             =   "Tipo"
               Top             =   360
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   900
               Caption         =   "Fiscal"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cmdVISUAL cmdGerar 
               Height          =   375
               Left            =   9120
               TabIndex        =   8
               Top             =   480
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   661
               Caption         =   "Novo"
               Acao            =   1
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cboVISUAL cboPeriodo 
               Height          =   510
               Left            =   120
               TabIndex        =   7
               Tag             =   "Tipo"
               Top             =   360
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   900
               Caption         =   "Período"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
         End
         Begin VTOcx.grdVISUAL grdOrdens 
            Height          =   4695
            Left            =   0
            TabIndex        =   10
            Top             =   1200
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   8281
            CorBorda        =   16711680
            Caption         =   "Ordem de Serviço"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
      End
   End
End
Attribute VB_Name = "TPRT101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    TPRT201.Show
End Sub

Private Sub cboPeriodo_Click()
    listar
End Sub
Private Sub cboFiscal_Click()
    'listar (True)
End Sub

Public Sub listar()
    Dim Filtro As String
    Filtro = " PERIODO = '" & cboPeriodo.Text & "'"
    If Len(cboFiscal.Text) > 0 Then
        Filtro = Filtro & " AND (FISCAL =  '" & cboFiscal.Text & "' OR FISCAL2 = '" & cboFiscal.Text & "')"
    End If
    If Len(cboStatus.Text) > 0 Then
        Filtro = Filtro & " AND (SITUACAO =  '" & cboStatus.Text & "')"
    End If
    If Not grdOrdens.preencher(Bdados, "SELECT * FROM VIS_BCP_ORDEM_SERVICO WHERE " & Filtro & " ORDER BY SERVICO DESC") Then
            Mensagem "Não existem ordem de serviço para o contribuinte"
    End If

End Sub

Private Sub cmdBuscar_Click()
    If Len(cboPeriodo.Text) = 0 Then
        Mensagem ("Informe um período")
        Exit Sub
    End If
    listar
End Sub

Private Sub cmdGerar_Click()
    Dim Form As New TPRT201
    'Form.OrdemServico = grdOrdens.SelectedItem
    Form.Show
End Sub

Private Sub cmdVISUAL1_Click()
'    Dim Form As New TPRT101OLD
 '   Form.Show
End Sub

Private Sub Form_Load()
    Dim ano As String
    Dim x As Integer
    ano = Format(Now, "YYYY")
    mes = Format(Now, "MM")
    For x = mes To 1 Step -1
        cboPeriodo.AddItem (ano & "-" & Format(x, "00"))
    Next x
     cboFiscal.AddItem ""
    Dim rs As VSRecordset
    
    If Bdados.AbreTabela("SELECT TUS_COD_USUARIO FROM TAB_USUARIO ORDER BY TUS_COD_USUARIO", rs) Then
        Do While Not rs.EOF
            cboFiscal.AddItem rs(0)
            rs.MoveNext
        Loop
    End If
    cboStatus.AddItem "ABERTA"
    cboStatus.AddItem "FISCALIZAÇÃO"
    cboStatus.AddItem "COMPARECIMENTO"
    cboStatus.AddItem "ATENDIMENTO"
    cboStatus.AddItem "EXECUÇÃO"
    cboStatus.AddItem "FINALIZADA"
    cboStatus.AddItem "RE-ABERTA"
End Sub

Private Sub grdOrdens_dblClick()
    Dim Form As New TPRT201
    Form.OrdemServico = grdOrdens.SelectedItem
    Form.Show
    
End Sub
