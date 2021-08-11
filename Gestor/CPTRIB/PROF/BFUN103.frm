VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.0#0"; "VTControles.ocx"
Begin VB.Form BFUN103 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BFUN103.frx":0000
   ScaleHeight     =   7785
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView treLotacao 
      Height          =   2850
      Left            =   60
      TabIndex        =   12
      Top             =   690
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   5027
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VTOcx.cboVISUAL cboHierarquia 
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   6750
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      Caption         =   "Hierarquia"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.txtVISUAL txtCodigo 
      Height          =   285
      Left            =   600
      TabIndex        =   10
      Tag             =   "Codigo"
      Top             =   5760
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   503
      Caption         =   "Codigo"
      Text            =   ""
      Enabled         =   0   'False
   End
   Begin Cabecalho.rodVISUAL rodVisual 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   11
      Top             =   7260
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   926
      CorFundo        =   12632256
      CorFrente       =   4210752
      Begin VTOcx.cmdVISUAL cmdCancelar 
         Height          =   405
         Left            =   4170
         TabIndex        =   8
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         Caption         =   "&Cancelar"
         Acao            =   9
         CorBorda        =   4210752
         CorFrente       =   4210752
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   405
         Left            =   5310
         TabIndex        =   4
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   714
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   4210752
         CorFrente       =   4210752
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   405
         Left            =   6360
         TabIndex        =   5
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   714
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   4210752
         CorFrente       =   4210752
      End
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   405
         Left            =   3240
         TabIndex        =   7
         Top             =   90
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   714
         Caption         =   "&Novo"
         Acao            =   1
         CorBorda        =   4210752
         CorFrente       =   4210752
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   405
         Left            =   7410
         TabIndex        =   6
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   4210752
         CorFrente       =   4210752
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   1138
      Formulario      =   "[Nome]"
      Descricao       =   "[Descricao]"
      Icone           =   "BFUN103.frx":0342
   End
   Begin VTOcx.txtVISUAL txtSigla 
      Height          =   285
      Left            =   2670
      TabIndex        =   0
      Tag             =   "Sigla"
      Top             =   5760
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      Caption         =   "Sigla"
      Text            =   ""
      MaxLen          =   10
   End
   Begin VTOcx.txtVISUAL txtNome 
      Height          =   285
      Left            =   705
      TabIndex        =   1
      Tag             =   "Nome"
      Top             =   6090
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   503
      Caption         =   "Nome"
      Text            =   ""
      MaxLen          =   80
   End
   Begin VTOcx.txtVISUAL txtResponsavel 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   6420
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   503
      Caption         =   "Responsavel"
      Text            =   ""
      MaxLen          =   80
   End
End
Attribute VB_Name = "BFUN103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dragNode As Node, hilitNode As Node

Private Sub cmdCancelar_Click()
    Edita.LimpaCampos Me
    treLotacao.Refresh
    treLotacao.Enabled = True
    treLotacao.SetFocus
End Sub

Private Sub cmdExcluir_Click()
    On Error Resume Next

    If treLotacao.Enabled Then
        If treLotacao.SelectedItem Is Nothing Then
            Exit Sub
        End If
        If Not treLotacao.SelectedItem.Child Is Nothing Then
            Util.Erro "Escolha uma lotação sem lotações subordinadas."
            Exit Sub
        End If
        If Util.Confirma("Excluir " & treLotacao.SelectedItem & " ?") Then
            If Bdados.DeletaDados("TAB_LOTACAO", "TLO_CODIGO=" & Util.ParseString(treLotacao.SelectedItem.Key, ":", 2)) Then
                Avisa "Lotação apagada com sucesso."
                cmdCancelar_Click
                ExibirArvore
            End If
        End If
    End If
    Exit Sub
Trata:
    Util.Erro ERR.Description
End Sub

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    txtCodigo = BuscarUltimaLotacao
    treLotacao.Enabled = False
    If Not treLotacao.SelectedItem Is Nothing Then
        cboHierarquia = Trim$(Util.ParseString(treLotacao.SelectedItem, "(", 1))
    Else
        cboHierarquia.SetarLinha 0, 1
    End If
    txtSigla.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    On Error GoTo Trata
    If Edita.CriticaCampos(Me) Then
        If txtSigla = cboHierarquia Then
            Erro "Erro na hierarquia informada."
            cboHierarquia.SetFocus
            Exit Sub
        End If
        Dim Valores As String, Campos As String
        Campos = "TLO_CODIGO,TLO_HIERARQ,TLO_SIGLA,TLO_NOME,TLO_RESPONSAVEL"
        Valores = Bdados.PreparaValor(txtCodigo, IIf(cboHierarquia = "", 0, cboHierarquia.Coluna(1).Valor), txtSigla, txtNome, txtResponsavel)
        If Bdados.GravaDados("TAB_LOTACAO", Valores, Campos, "TLO_CODIGO=" & txtCodigo) Then
            Util.Avisa "Lotação gravada com sucesso."
            Edita.LimpaCampos Me
            ExibirArvore
            PreencherHierarquia
        Else
            Util.Erro "Problemas ao gravar."
        End If
    End If
    treLotacao.Enabled = True
    Exit Sub
Trata:
    Erro ERR.Description
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVisual.Exibir Bdados, Me.Name
    ExibirArvore
    PreencherHierarquia
End Sub

Public Sub ExibirArvore()
    On Error GoTo ERR
    Dim rs As VSRecordset
    
    treLotacao.Nodes.Clear
    If Bdados.AbreTabela("SELECT * FROM TAB_LOTACAO ORDER BY TLO_HIERARQ, TLO_CODIGO", rs) Then
        Do While Not rs.EOF
            'Avisa rs!TLO_SIGLA & " (" & rs!TLO_CODIGO & " - PAI : " & Bdados.BuscaCodigo("SELECT TLO_SIGLA FROM TAB_LOTACAO WHERE TLO_CODIGO=" & rs!TLO_HIERARQ) & ")"
            If CInt(rs!tlo_Codigo) = CInt(rs!TLO_HIERARQ) Then
                treLotacao.Nodes.Add , , "CODIGO:" & rs!tlo_Codigo, rs!tlo_sigla & " (" & rs!tlo_Nome & ")"
            Else
                treLotacao.Nodes.Add treLotacao.Nodes("CODIGO:" & rs!TLO_HIERARQ), tvwChild, "CODIGO:" & rs!tlo_Codigo, rs!tlo_sigla & " (" & rs!tlo_Nome & ")"
            End If
            rs.MoveNext
        Loop
        treLotacao.Nodes("CODIGO:0").Expanded = True
    End If
    Bdados.FechaTabela rs
    Exit Sub
ERR:
    If ERR.Number <> 0 Then
        Erro "Problemas ao exibir [" & rs!tlo_sigla & "]"
        treLotacao.Nodes.Add treLotacao.Nodes("CODIGO:" & rs!TLO_HIERARQ), tvwChild, "CODIGO:" & rs!tlo_Codigo, rs!tlo_sigla & " (" & rs!tlo_Nome & ")"
        Resume Next
        Exit Sub
    End If
End Sub

Public Function BuscarUltimaLotacao() As Integer
    On Error GoTo Trata
    Dim rs As VSRecordset
    Dim sql As String
    
    sql = "SELECT MAX(TLO_CODIGO) FROM TAB_LOTACAO"
    If Bdados.AbreTabela(sql, rs) Then
        BuscarUltimaLotacao = rs(0) + 1
    Else
        BuscarUltimaLotacao = 0
    End If
    Bdados.FechaTabela
    Exit Function
Trata:
    Util.Erro ERR.Description
End Function

Private Sub PreencherHierarquia()
    Dim sql As String
    sql = "SELECT TLO_SIGLA, TLO_CODIGO FROM TAB_LOTACAO ORDER BY TLO_SIGLA"
    cboHierarquia.Preencher Bdados, sql
End Sub

Private Sub treLotacao_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo Trata
    Dim rs As VSRecordset
    Dim sql As String

    Edita.LimpaCampos Me
    sql = "SELECT * FROM TAB_LOTACAO WHERE TLO_CODIGO=" & Util.ParseString(Node.Key, ":", 2)
    If Bdados.AbreTabela(sql, rs) Then
        txtCodigo = rs!tlo_Codigo
        txtSigla = "" & rs!tlo_sigla
        txtNome = "" & rs!tlo_Nome
        txtResponsavel = "" & rs!TLO_RESPONSAVEL
        If Not Node.Parent Is Nothing Then
            cboHierarquia = Trim$(Util.ParseString(Node.Parent, "(", 1))
        End If
    End If
    Bdados.FechaTabela rs
    Exit Sub
Trata:
    Util.Erro ERR.Description
    Resume
End Sub
