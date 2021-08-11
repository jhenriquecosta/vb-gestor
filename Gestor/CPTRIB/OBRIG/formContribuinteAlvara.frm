VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form formContribuinteAlvara 
   BackColor       =   &H00FBEDE8&
   Caption         =   "Contribuintes"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15690
   ForeColor       =   &H00000000&
   Icon            =   "formContribuinteAlvara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8985
   ScaleMode       =   0  'User
   ScaleWidth      =   17744.26
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cboVISUAL cboBairro 
      Height          =   315
      Left            =   6480
      TabIndex        =   15
      Tag             =   "Tributo"
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   556
      Caption         =   "Bairro"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Requerido       =   0   'False
   End
   Begin VB.CheckBox chkEmitidos 
      BackColor       =   &H00FBEDE8&
      Caption         =   "Somente contribuinte com alvaras emitidos"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin MSComctlLib.ListView Lst1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7646
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
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   1138
      Formulario      =   "Contribuintes"
      Descricao       =   ""
      Icone           =   "formContribuinteAlvara.frx":08CA
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   8475
      Width           =   15690
      _ExtentX        =   27675
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL cmdRelatorio 
         Height          =   375
         Left            =   13200
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
         Left            =   14520
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
   Begin VTOcx.cboVISUAL cboAtividade 
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Tag             =   "Tributo"
      Top             =   720
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   556
      Caption         =   "Atividade"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Requerido       =   0   'False
   End
   Begin VTOcx.cboVISUAL cboRestricao 
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Tag             =   "Tributo"
      Top             =   1080
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   556
      Caption         =   "Restrição"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Requerido       =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdBuscar 
      Height          =   375
      Left            =   14400
      TabIndex        =   7
      Top             =   1080
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Buscar"
      Acao            =   5
      CorBorda        =   16711680
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdAnterior 
      Height          =   375
      Left            =   13200
      TabIndex        =   9
      Top             =   8040
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "Anterior"
      Acao            =   8
      CorBorda        =   16711680
      CorFundo        =   16777088
      Icone           =   "formContribuinteAlvara.frx":205C
   End
   Begin VTOcx.cmdVISUAL cmdProximo 
      Height          =   375
      Left            =   14400
      TabIndex        =   10
      Top             =   8040
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "Proximo"
      Acao            =   8
      CorBorda        =   16711680
      CorFundo        =   16777088
      Icone           =   "formContribuinteAlvara.frx":237E
   End
   Begin MSComctlLib.ListView Lst2 
      Height          =   4335
      Left            =   11040
      TabIndex        =   11
      Top             =   1560
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7646
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
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
   Begin MSComctlLib.ListView Lst3 
      Height          =   1935
      Left            =   0
      TabIndex        =   12
      Top             =   6000
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   3413
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
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
   Begin VTOcx.txtVISUAL txtCpf 
      Height          =   300
      Left            =   0
      TabIndex        =   13
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   529
      Caption         =   " CPF/CNPJ"
      Text            =   ""
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdPesquisaCpf 
      Height          =   315
      Left            =   2790
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   690
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
End
Attribute VB_Name = "formContribuinteAlvara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As Recordset
Dim sqlRelatorio As String


Sub CabLancto()
      
      Lst1.ListItems.Clear
      Lst1.ColumnHeaders.Clear
      Lst1.ColumnHeaders.Add , , "Inscricao", 1300
      Lst1.ColumnHeaders.Add , , "CpfCnpj", 1800
      Lst1.ColumnHeaders.Add , , "Fantasia", 2500
      Lst1.ColumnHeaders.Add , , "Atividade", 2500
      Lst1.ColumnHeaders.Add , , "Emitido", 2500
      Lst1.ColumnHeaders.Add , , "Não Emitido", 2500
      
      Lst1.View = lvwReport
   
End Sub

Private Sub cmdBuscar_Click()
    
   ' Criterio2 = ""
    Dim sqlRestricao
    Dim sqlAtiv
    Dim sqlEmitido
    Dim sqlPeriodo
    Dim sqlBairro

    CabLancto
    
     
    If cboAtividade <> "" Then sqlAtiv = "Atividade='" & cboAtividade.Text & "'"
    If cboBairro <> "" Then sqlBairro = "Bairro='" & cboBairro.Text & "'"
    
    If chkEmitidos.Value = 1 Then sqlEmitido = "Alvaras > 0"
    If cboRestricao <> "" Then
       
       Restricao = cboRestricao.coluna(1).Valor
       If Restricao = 1 Then 'nao pagos
          sqlRestricao = "(Alvaras_Pago = 0 and alvaras > 0)"
       ElseIf Restricao = 2 Then 'pagos'
          sqlRestricao = "(Alvaras_Pago > 0 and alvaras = alvaras_pago)"
       End If
    
    End If
    'If txtPeriodo = "" Then sqlPeriodo='
    
    
    
    Dim col As New Collection
    col.Add sqlPeriodo
    col.Add sqlRestricao
    col.Add sqlAtiv
    col.Add sqlEmitido
    col.Add sqlBairro
     
    
  '  col.Add sqlImovel
  
    Dim Sql As String
    
    Sql = "SELECT * FROM VIEW_CONTRIBUINTE_ALVARA"
    
    Sql = montaSqlWhere(Sql, col)
    sqlRelatorio = Sql & " ORDER BY bairro,atividade"
    Sql = Sql & " ORDER BY ALVARA_ULTIMO DESC"
    LoadLancto Sql
End Sub

Sub LoadLancto(Sql As String)


On Error GoTo erros
    Dim mVlr As Currency
    Dim mVlP As Currency
    Dim mVlA As Currency
    Dim Bd As New Connection
    Dim Rs As Recordset
   
  
    abreConexao Bd
    Set rst = New ADODB.Recordset
   
    rst.CursorLocation = adUseClient
    rst.PageSize = 16
    rst.Open Sql, Bd, adOpenStatic, adLockReadOnly, adAsyncFetchNonBlocking
    Set rst.ActiveConnection = Nothing
    
    Call ExibePag(1)
    
    fechaConexao Bd
    
    Exit Sub
erros:
  '  MOSTRAERRO Me.Name

End Sub

Sub ExibePag(pag As Double)
   
 '  If Rst.EOF Then
 '     Util.Informa "Não foram encontrados mais registros no criterio especificado!"
 '     Exit Sub
 '  End If
   Dim Contador As Integer
   Dim mVlr As Integer
   Dim tamanhoPagina As Integer
   Dim Lst
   If pag = 0 Then pag = 1
   rst.AbsolutePage = pag
   tamanhoPagina = rst.PageSize
   Contador = 1
    
    
    
   Lst1.ListItems.Clear
   mVlr = 0
   
   While (Not rst.EOF) And (Contador <= tamanhoPagina)
   
        Set Lst = Lst1.ListItems.Add(, , carregaCampo(rst!Inscricao))
        Lst.SubItems(1) = carregaCampo(rst!CPFCNPJ)
        Lst.SubItems(2) = carregaCampo(rst!Fantasia)
        Lst.SubItems(3) = carregaCampo(rst!Atividade)
        Lst.SubItems(4) = carregaCampo(rst!Periodo)
        Lst.SubItems(5) = carregaCampo(rst!PeriodoAberto)
        
        Contador = Contador + 1
        rst.MoveNext
    Wend

End Sub


Private Sub cmdRelatorio_Click()
    
    Dim relCons As New relAlvarasEmitidos
    configRelatorio relCons, sqlRelatorio
    relCons.Show
    
End Sub

Private Sub cmdSair_Click()
Unload Me

End Sub

Private Sub cmdProximo_Click()
 If rst.AbsolutePage <> -3 Then
      Call ExibePag(rst.AbsolutePage)
      cmdAnterior.Enabled = True
    Else
      Avisa " Não Há Mais Pagina! "
      Call ExibePag(rst.PageCount)
    End If
End Sub
Private Sub cmdAnterior_Click()
    
    If rst.AbsolutePage > 2 And rst.AbsolutePage <> -2 Then
       Call ExibePag(rst.AbsolutePage - 2)
    Else
        If rst.AbsolutePage = -3 Then
           Call ExibePag(rst.PageCount - 1)
        Else
           Avisa " Não há mais pagina"
           Call ExibePag(1)
        End If
    End If
    
End Sub
Private Sub Form_Load()
    CabLancto
    cboRestricao.PreencherGeral Bdados, "RESTRICAO DAM"
    carregaCombo cboAtividade, "select tae_nome from tab_atividade_economica order by tae_nome"
    Call Edita.AtualizaCombo(Bdados, cboBairro, "Select DISTINCT(tba_nome) From Tab_Bairro order by tba_nome")
    
 
End Sub

Private Sub Lst1_ItemClick(ByVal Item As MSComctlLib.ListItem)
      Lst2.ListItems.Clear
      Lst2.ColumnHeaders.Clear
      Lst2.ColumnHeaders.Add , , "Periodo", 1200
      Lst2.ColumnHeaders.Add , , "Impressao", 1200
      Lst2.ColumnHeaders.Add , , "Status", 1200
      Lst2.ColumnHeaders.Add , , "Usuario", 1200
      
      Lst2.View = lvwReport
      
      Lst3.ListItems.Clear
      Lst3.ColumnHeaders.Clear
      Lst3.ColumnHeaders.Add , , "Documento", 1200
      Lst3.ColumnHeaders.Add , , "Periodo", 1200
      Lst3.ColumnHeaders.Add , , "Valor", 1200
      Lst3.ColumnHeaders.Add , , "Geracao", 1200
      Lst3.ColumnHeaders.Add , , "Pagto", 1200
      Lst3.ColumnHeaders.Add , , "Situacao", 2000
      Lst3.ColumnHeaders.Add , , "Obs", 9000
      
      Lst3.View = lvwReport
      
      
      Dim Sql As String
      
      Sql = "select tai_periodo as periodo, tai_data_validade as validade, tai_tus_cod_usuario as usuario, tai_data_impressao as impressao, tai_status as status from vis_alvara_impresso where tai_tci_im='" & Item & "' order by periodo"
      Dim Rs As VSRecordset
      Bdados.AbreTabela Sql, Rs
      Do While Not Rs.EOF
        Set Lst = Lst2.ListItems.Add(, , Rs.Fields("periodo"))
        Lst.SubItems(1) = Rs.Fields("Impressao")
        Lst.SubItems(2) = Rs.Fields("Status")
        Lst.SubItems(3) = Rs.Fields("Usuario")
        
        Rs.MoveNext
        
      Loop
      
       Sql = "select left(periodo,4) as periodo,geracao, documento,valor,observacao,situacao,pagamento from view_obrigacao where imposto='11210101' and  inscricao='" & Item & "'"
    '  Dim rs As VSRecordset
      Bdados.AbreTabela Sql, Rs
      Do While Not Rs.EOF
        Set Lst = Lst3.ListItems.Add(, , Rs.Fields("Documento"))
        
        Lst.SubItems(1) = Rs.Fields("Periodo")
        Lst.SubItems(2) = Format(Rs.Fields("Valor"), "##,##0.00")
        Lst.SubItems(3) = Format(Rs.Fields("geracao"), "dd/mm/yy")
        Lst.SubItems(4) = Format(Rs.Fields("Pagamento"), "dd/mm/YY")
        Lst.SubItems(5) = Rs.Fields("Situacao")
        Lst.SubItems(6) = "" & Rs.Fields("Observacao")
        
        
        Rs.MoveNext
        
      Loop
      
      
           
End Sub

   
