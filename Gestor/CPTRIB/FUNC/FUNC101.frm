VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form FUNC101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FUNC101"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   7
      Top             =   5730
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   847
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   345
         Left            =   6180
         TabIndex        =   11
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   5205
         TabIndex        =   4
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   8175
         TabIndex        =   6
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   7170
         TabIndex        =   5
         Top             =   105
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   3585
      Left            =   30
      TabIndex        =   8
      Top             =   2115
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   6324
      CorBorda        =   32768
      Caption         =   "Funcionários"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1440
      Left            =   30
      TabIndex        =   9
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   645
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   2540
      Altura          =   1905
      Caption         =   " Dados do Funcionário"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboCargo 
         Height          =   510
         Left            =   105
         TabIndex        =   12
         Tag             =   "Cargo"
         Top             =   780
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   900
         Caption         =   "Cargo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   4210752
      End
      Begin VTOcx.txtVISUAL txtCpf 
         Height          =   480
         Left            =   8325
         TabIndex        =   3
         Tag             =   "CPF"
         Top             =   300
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   847
         Caption         =   "CPF"
         Text            =   ""
         Formato         =   1
         Restricao       =   2
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   15
      End
      Begin VTOcx.cboVISUAL cboLotacao 
         Height          =   510
         Left            =   5025
         TabIndex        =   2
         Tag             =   "Lotação"
         Top             =   780
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   900
         Caption         =   "Lotação"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   4210752
      End
      Begin VTOcx.txtVISUAL txtNome 
         Height          =   480
         Left            =   1770
         TabIndex        =   1
         Tag             =   "Nome"
         Top             =   300
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   847
         Caption         =   "Nome"
         Text            =   ""
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   150
      End
      Begin VTOcx.txtVISUAL txtMatricula 
         Height          =   480
         Left            =   60
         TabIndex        =   0
         Tag             =   "Matrícula"
         Top             =   300
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   847
         Caption         =   "Matrícula"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   20
      End
   End
   Begin VTOcx.txtVISUAL txtCodigo 
      Height          =   480
      Left            =   195
      TabIndex        =   10
      Top             =   885
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   847
      Caption         =   ""
      Text            =   ""
      Requerido       =   0   'False
      AlinhamentoRotulo=   1
      CorRotulo       =   16384
      CorTexto        =   4194304
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1138
      Icone           =   "FUNC101.frx":0000
   End
End
Attribute VB_Name = "FUNC101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private GeraCod As New ContaCorrente

Private Sub cmdExcluir_Click()
    Dim condicao As String
    If grdDados.ListItems.Count < 1 Then Exit Sub
    condicao = "TFU_CODIGO = '" & txtCodigo & "'"
    If txtCodigo <> "" Then
        If Confirma("Deseja excluir registro?", "Excluir?") Then
            If Bdados.DeletaDados("TAB_FUNCIONARIO", condicao) Then
                Avisa "Dados Excluidos com Sucesso"
                Limpa
                CarregaFunc
            End If
        End If
    Else
        Avisa "Selecione um Registro"
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    CarregaFunc
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim campos As String
    Dim valores As String
    Dim condicao As String
    Dim Codigo As String

    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If txtCodigo = "" Then
        Codigo = CStr(GeraCod.GeraCodPagamento(3))
    Else
        Codigo = txtCodigo
    End If

    campos = " TFU_CODIGO,TFU_MATRICULA,TFU_NOME,TFU_CPF,TFU_COD_CARGO,TFU_TLO_COD_LOTACAO"
    valores = Bdados.PreparaValor(txtCodigo, txtMatricula, txtNome, txtCpf, cboCargo.Coluna(1).Valor, cboLotacao.Coluna(1).Valor)
    condicao = "TFU_CODIGO = '" & Codigo & "'"
    If Bdados.GravaDados("TAB_FUNCIONARIO", valores, campos, condicao) Then
        Avisa "Dados Salvos com Sucesso"
        Limpa
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
    cboCargo.Preencher Bdados, "SELECT * FROM TAB_FUNCAO", 1
    cboLotacao.Preencher Bdados, "SELECT tlo_codigo,tlo_descricao FROM vis_lotacao", 1
End Sub

Private Sub Limpa()
    txtCodigo = ""
    txtMatricula = ""
    txtNome = ""
    txtCpf = ""
    cboCargo.ListIndex = -1
    cboLotacao.ListIndex = -1
End Sub

Private Sub CarregaFunc()
    Dim Sql As String
    Dim condicao As String

    Sql = "select TFU_CODIGO as Código,"
    Sql = Sql & " TFU_MATRICULA as Matrícula,"
    Sql = Sql & " TFU_NOME as Funcionário,"
    Sql = Sql & " TFU_CPF as CPF,"
    Sql = Sql & " TFU_DESCRICAO as Cargo,"
    Sql = Sql & " TLO_DESCRICAO  as Lotação,"
    Sql = Sql & " TFU_COD_CARGO ,"
    Sql = Sql & " TFU_TLO_COD_LOTACAO,"
    Sql = Sql & " tab_funcionario,tab_funcao,vis_lotacao"
    Sql = Sql & " Where TFU_CODIGO = TFU_COD_CARGO"
    Sql = Sql & " and TLO_CODIGO  = TFU_TLO_COD_LOTACAO "

    grdDados.Preencher Bdados, Sql, 0, 1500, 3000, 1500, 2000, 2000
End Sub

Private Sub grdDados_DblClick()
    If grdDados.ListItems.Count >= 1 Then

        txtCodigo = grdDados.SelectedItem
        txtMatricula = grdDados.SelectedItem.SubItems(1)
        txtNome = grdDados.SelectedItem.SubItems(2)
        txtCpf = grdDados.SelectedItem.SubItems(3)
        cboCargo.SetarLinha grdDados.SelectedItem.SubItems(6), 1
        cboLotacao.SetarLinha grdDados.SelectedItem.SubItems(7), 1
    End If
End Sub
