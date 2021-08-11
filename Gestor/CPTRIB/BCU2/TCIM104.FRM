VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.0#0"; "VTControles.ocx"
Begin VB.Form TCIM104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   12
      Top             =   5190
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   5370
         TabIndex        =   8
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4350
         TabIndex        =   7
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6570
         TabIndex        =   9
         Top             =   120
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.grdVISUAL grdGrades 
      Height          =   2985
      Left            =   30
      TabIndex        =   10
      Top             =   1740
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   4339
      CorBorda        =   32768
      CabecalhoEstado =   ""
      CabecalhoCliente=   "Prefeitura Municipal de Imperatriza"
      CabecalhoSecretaria=   "Status de Ditação de Boletins"
      CabecalhoDepartamento=   "Cadastro Técnico Municipal - CTM"
      Caption         =   "Grades Cadastradas"
      CorTitulo       =   32768
      CorCaption      =   16777215
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   1138
      Icone           =   "Tcim104.frx":0000
   End
   Begin VTOcx.cmdVISUAL cmdBuscar 
      Height          =   375
      Left            =   6330
      TabIndex        =   4
      Top             =   1320
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      Caption         =   "&Buscar"
      Acao            =   5
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.txtVISUAL txtConsUsuario 
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Tag             =   "Distrito"
      Top             =   690
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   503
      Caption         =   "Usuário"
      Text            =   ""
   End
   Begin VTOcx.cboVISUAL cboConsStatus 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      Caption         =   "Status"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.cboVISUAL cboStatus 
      Height          =   315
      Left            =   2070
      TabIndex        =   6
      Top             =   4755
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      Caption         =   "Status"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.txtVISUAL txtLote 
      Height          =   285
      Left            =   150
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4770
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      Caption         =   "Lote"
      Text            =   ""
      Enabled         =   0   'False
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtDtInicial 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      Caption         =   "Período"
      Text            =   ""
      Formato         =   0
      Restricao       =   2
   End
   Begin VTOcx.txtVISUAL txtDtFinal 
      Height          =   300
      Left            =   2100
      TabIndex        =   3
      Top             =   1380
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   529
      Caption         =   "a "
      Text            =   ""
      Formato         =   0
      Restricao       =   2
   End
   Begin VTOcx.cmdVISUAL cmdFechar 
      Height          =   375
      Left            =   5070
      TabIndex        =   13
      Top             =   4740
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   661
      Caption         =   "&Fechar quadras"
      Acao            =   1
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCIM104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Capa As New cCapa

Private Sub cmdBuscar_Click()
    Set Capa = New cCapa
    
    Capa.ExibirBatimento grdGrades, txtConsUsuario, cboConsStatus
    Set Capa = Nothing
End Sub

Private Sub cmdFechar_Click()
    Capa.FecharQuadras
    Avisa "Processamento finalizado."
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo Trata
    Dim Rpt As vsclass.VSRelatorio
    Dim Filtro As String
    
    Set Rpt = New vsclass.VSRelatorio
    Screen.MousePointer = 11
    With Rpt
        If .DefinirArquivo(Bdados, App.Path & "\TProdutividade.rpt") Then
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), "TProdutividade", Aplicacoes.Usuario
            .Titulo = "Produtividade dos digitadores"
            
            Filtro = ""
            If Trim$(txtConsUsuario) <> "" Then
                Filtro = Filtro & IIf(Filtro <> "", " and ", "") & " {VIS_LOTE_DIGITACAO.VBT_USUARIO} = '" & txtConsUsuario & "'"
            End If
            If cboConsStatus <> "" Then
                Filtro = Filtro & IIf(Filtro <> "", " and ", "") & " {VIS_LOTE_DIGITACAO.TCL_STATUS} = " & cboConsStatus.Coluna(1).Valor
            End If
            If Trim(txtDtInicial) <> "" Then
                Filtro = Filtro & IIf(Filtro <> "", " and ", "") & " {VIS_LOTE_DIGITACAO.TCL_DATA_ABERTURA} >= Date (" & Year(txtDtInicial) & "," & _
                                Month(txtDtInicial) & "," & Day(txtDtInicial) & ")"
            End If
            If Trim(txtDtFinal) <> "" Then
                Filtro = Filtro & IIf(Filtro <> "", " and ", "") & " {VIS_LOTE_DIGITACAO.TCL_DATA_ABERTURA} <= Date (" & Year(txtDtFinal) & "," & _
                                Month(txtDtFinal) & "," & Day(txtDtFinal) & ")"
            End If
            If Filtro <> "" Then .Selecao = Filtro
            
            .Visualizar
        End If
    End With
    Set Rpt = Nothing
   
    Screen.MousePointer = 0
    Exit Sub
    
Trata:
    Erro Err.Description
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Not Capa Is Nothing Then
        Capa.Status = cboStatus.Coluna(1).Valor
        Capa.Gravar
        Util.Informa "Controle registrado."
        Edita.LimpaCampos Me
        cmdBuscar_Click
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    
    cboStatus.Preencher Bdados, "SELECT TGE_NOME,TGE_CODIGO FROM VIS_STATUS_GRADE_LOTE WHERE TGE_CODIGO<>2"
    cboConsStatus.PreencherGeral Bdados, "STATUS GRADE LOTE"
End Sub

Private Sub grdGrades_Click()
    If Not grdGrades.SelectedItem Is Nothing Then
        If Capa.Buscar(grdGrades.SelectedItem, grdGrades.SelectedItem.SubItems(1), grdGrades.SelectedItem.SubItems(2)) Then
            txtLote = Trim(Capa.Distrito) & "." & Trim(Capa.Setor) & "." & Trim(Capa.Quadra)
            cboStatus = grdGrades.SelectedItem.SubItems(9)
            cboStatus.Enabled = (grdGrades.SelectedItem.SubItems(9) <> "ABERTO")
            If cboStatus.Enabled Then cboStatus.SetFocus
        End If
    End If
End Sub

Private Sub grdGrades_KeyDown(KeyCode As Integer, Shift As Integer)
    Edita.LimpaCampos Me
End Sub

Private Sub grdGrades_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then grdGrades_Click
End Sub
