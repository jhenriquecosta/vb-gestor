VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU107 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU107"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7695
   StartUpPosition =   1  'CenterOwner
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   2325
      Left            =   90
      TabIndex        =   6
      Top             =   2670
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   4101
      CorTitulo       =   16711680
      CorCaption      =   16777215
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   2250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1545
      Width           =   4905
   End
   Begin VTOcx.cboVISUAL cboGrupo 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   810
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   556
      Caption         =   "Grupo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   5070
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   4215
         TabIndex        =   8
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   3045
         TabIndex        =   7
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6540
         TabIndex        =   10
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   5385
         TabIndex        =   9
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.txtVISUAL txtCodigo 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   1200
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   503
      Caption         =   "Código do Componente"
      Text            =   ""
   End
   Begin VTOcx.txtVISUAL txtValor 
      Height          =   285
      Left            =   1755
      TabIndex        =   3
      Top             =   2235
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   503
      Caption         =   "Valor"
      Text            =   ""
   End
   Begin VTOcx.txtVISUAL txtUnidade 
      Height          =   285
      Left            =   3510
      TabIndex        =   4
      Top             =   2250
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      Caption         =   "Unidade"
      Text            =   ""
   End
   Begin VTOcx.txtVISUAL txtFatorCalc 
      Height          =   285
      Left            =   5160
      TabIndex        =   5
      Top             =   2265
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   503
      Caption         =   "Fator Calc."
      Text            =   ""
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1138
      Icone           =   "TCIU107.frx":0000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   1500
      TabIndex        =   12
      Top             =   1620
      Width           =   720
   End
End
Attribute VB_Name = "TCIU107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboGrupo_Click()
    PreencherGrid
End Sub

Private Sub cmdExcluir_Click()
    Dim Condicao  As String
    If txtCodigo.Enabled = False Then
        If Confirma("Deseja excluir esse componente?") Then
            Condicao = "tco_cod_componente = '" & txtCodigo & "' and tco_grupo = '" & cboGrupo.Coluna(0).Valor & "'"
            If Bdados.DeletaDados("tab_componente_avancado", Condicao) Then
                Util.Avisa "Operação concluída com sucesso."
                cmdLimpar_Click
                PreencherGrid
            End If
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    txtCodigo = ""
    txtDescricao = ""
    txtValor = ""
    txtUnidade = ""
    txtFatorCalc = ""
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores  As String
    Dim Campos   As String
    Dim Condicao As String
    Campos = "tco_cod_componente,tco_grupo,tco_descricao_componente,tco_valor,tco_unid_moneta,tco_cod_componente_fator_calc"
    Valores = Bdados.PreparaValor(txtCodigo, cboGrupo.Coluna(0).Valor, Bdados.Converte(txtDescricao, tctexto), txtValor, txtUnidade, txtFatorCalc)
    Condicao = "tco_cod_componente = '" & txtCodigo & "' and tco_grupo = '" & cboGrupo.Coluna(0).Valor & "'"
    If Bdados.GravaDados("tab_componente_avancado", Valores, Campos, Condicao) Then
        Util.Avisa "Operação concluída com sucesso."
        cmdLimpar_Click
        PreencherGrid
    End If
End Sub

Private Sub Form_Load()
    PreencherGrid
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    cboGrupo.Preencher Bdados, "select tgc_cod_grupo,tgc_nome from tab_grupo_componente_avancado", 1
End Sub

Private Sub PreencherGrid()
    Dim Sql As String
    Sql = "select tco_cod_componente as Código ,"
    Sql = Sql & " tco_grupo as Grupo,"
    Sql = Sql & " tco_descricao_componente as Descrição,"
    Sql = Sql & " tco_valor as Valor,tco_unid_moneta as Unidade,"
    Sql = Sql & " tco_cod_componente_fator_calc as  Fator_Calc "
    Sql = Sql & " from tab_componente_avancado "
    Sql = Sql & " where 1 = 1"
    If cboGrupo.ListIndex <> -1 Then
        Sql = Sql & " and  tco_grupo = '" & cboGrupo.Coluna(0).Valor & "'"
    End If
    grdDados.Preencher Bdados, Sql
End Sub

Private Sub grdDados_DblClick()
    If grdDados.ListItems.Count >= 1 Then
        txtCodigo = grdDados.SelectedItem
        cboGrupo.SetarLinha grdDados.SelectedItem.SubItems(1)
        txtCodigo_LostFocus
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    Dim Sql As String
    Dim Rs  As VSRecordset
    
    Sql = "Select * from tab_componente_avancado where tco_cod_componente = '" & txtCodigo & "' and  tco_grupo = '" & cboGrupo.Coluna(0).Valor & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        txtDescricao = "" & Rs.Fields("tco_descricao_componente")
        cboGrupo.SetarLinha "" & Rs.Fields("tco_grupo")
        txtUnidade = "" & Rs.Fields("tco_unid_moneta")
        txtValor = "" & Rs.Fields("tco_valor")
        txtFatorCalc = "" & Rs.Fields("tco_cod_componente_fator_calc")
        txtCodigo.Enabled = False
        cboGrupo.Enabled = False
    Else
        cboGrupo.Enabled = True
        txtCodigo.Enabled = True
        txtDescricao = ""
        txtUnidade = ""
        txtValor = ""
        txtFatorCalc = ""
    End If
    
End Sub
