VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TMPU108 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TMPU105"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   10
      Top             =   3870
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   820
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6600
         TabIndex        =   5
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   7650
         TabIndex        =   3
         Top             =   60
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8655
         TabIndex        =   4
         Top             =   60
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.cboVISUAL cboDescricao 
      Height          =   315
      Left            =   6510
      TabIndex        =   9
      Top             =   3450
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   556
      Caption         =   ""
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Aplicar a todos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   1395
   End
   Begin VTOcx.txtVISUAL txtLogradouro 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   690
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   556
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
   End
   Begin VTOcx.txtVISUAL txtCodLogr 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Tag             =   "Codigo Logradouro"
      Top             =   690
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      Caption         =   "Logradouro"
      Text            =   ""
      Restricao       =   2
   End
   Begin VTOcx.grdVISUAL grdTrecho 
      Height          =   2985
      Left            =   0
      TabIndex        =   7
      Top             =   1050
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   4339
      Caption         =   "Trechos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtValor 
      Height          =   315
      Left            =   5130
      TabIndex        =   2
      Tag             =   "Valor"
      Top             =   3450
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "Valor"
      Text            =   ""
      Restricao       =   3
   End
   Begin VTOcx.grdVISUAL grdGrupo 
      Height          =   2655
      Left            =   5100
      TabIndex        =   1
      Top             =   1050
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   4339
      Caption         =   "Componentes"
      CorTitulo       =   32768
      CorCaption      =   16777215
      OcultarRodape   =   -1  'True
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   1138
      Icone           =   "TMPU108.frx":0000
   End
End
Attribute VB_Name = "TMPU108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDescricao_Click()
    txtValor = ""
    If cboDescricao.ListCount > 0 Then
        txtValor = cboDescricao.Coluna(1).Valor
    End If
End Sub

Private Sub chkTodos_Click()
    grdTrecho.MarcarTodos chkTodos
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    limparLogradouro
    txtCodLogr.SetFocus
End Sub

Private Sub prepararLogradouro(Codigo As String)
    Dim Sql As String
    
    limparLogradouro
    If Trim$(Codigo) <> "" Then
        Sql = "SELECT TTL_NOME " & Bdados.Concatena & "' '" & Bdados.Concatena & "tlg_nome" & _
                " FROM VIS_BVT " & _
                " WHERE tlg_cod_logradouro='" & Codigo & "'"
        If Bdados.AbreTabela(Sql) Then
            txtLogradouro = "" & Bdados.Tabela.Fields(0).Value
            Sql = "SELECT TTC_COD_TRECHO AS Trecho, " & _
                      " TTC_QUADRA as Quadra" & _
                    " FROM TAB_TRECHO " & _
                    " WHERE TTC_TLG_COD_LOGRADOURO='" & Codigo & "'" & _
                    " ORDER BY TTC_SEQ_TRECHO"
            grdTrecho.Preencher Bdados, Sql, 0.5 * grdTrecho.Width, 0.5 * grdTrecho.Width
        Else
            Erro "Logradouro não encontrado."
            txtCodLogr.SetFocus
        End If
        Bdados.FechaTabela
    End If
End Sub

Private Sub limparLogradouro()
    txtLogradouro = ""
    grdTrecho.Preencher Bdados, ""
    chkTodos.Value = vbUnchecked
End Sub

Private Sub cmdSalvar_Click()
    On Error GoTo trata
    Dim Campos As String, Valores As String
    Dim Trecho As Variant
    Dim Ok As Boolean
    
    Dim CodComponente As String, Valor As String
    
    If Edita.CriticaCampos(Me) Then
        '1.
        For Each Trecho In grdTrecho.ListItems
            If Trecho.Checked Then
                Ok = True
                Exit For
            End If
        Next
        
        '2.
        If Ok Then
            If Not grdGrupo.SelectedItem Is Nothing Then
                '3.
                If grdGrupo.SelectedItem <= 20 Then
                    CodComponente = txtValor
                    Valor = txtValor - 1
                Else
                    CodComponente = grdGrupo.SelectedItem
                    Valor = txtValor
                End If
                
                '4.
                Campos = "tdl_tlg_cod_logradouro, tdl_tcl_cod_componente, tdl_valor_item, tdl_tgl_cod_grupo, tdl_num_trecho"
                
                '5.
                For Each Trecho In grdTrecho.ListItems
                    If Trecho.Checked Then
                        Bdados.DeletaDados "tab_detalhe_logradouro", "tdl_tlg_cod_logradouro='" & txtCodLogr & "' and tdl_tgl_cod_grupo=" & grdGrupo.SelectedItem & " and tdl_num_trecho='" & Trecho & "'"
                        
                        Valores = Bdados.PreparaValor(txtCodLogr, CodComponente, Valor, grdGrupo.SelectedItem, Trecho)
                        Bdados.InsereDados "TAB_DETALHE_LOGRADOURO", Valores, Campos
                    End If
                Next
                
                Avisa grdGrupo.SelectedItem.SubItems(1) & " atualizado(a) com sucesso."
                
                If grdGrupo.SelectedItem.Index = grdGrupo.ListItems.Count Then
                    cmdLimpar_Click
                    grdGrupo.ListItems(1).Selected = True
                Else
                    grdGrupo.ListItems(grdGrupo.SelectedItem.Index + 1).Selected = True
                    grdGrupo_Click
                End If
                grdGrupo.SelectedItem.EnsureVisible
            End If
        Else
            Avisa "Informe o trecho."
        End If
        
    End If
    Exit Sub
trata:
    Erro Err.Description
    'Resume
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name
    prepararGrupos
End Sub

Private Sub grdGrupo_Click()
    If Not grdGrupo.SelectedItem Is Nothing Then
        preencherValores grdGrupo.SelectedItem
        txtValor.SetFocus
    End If
End Sub

Private Sub txtCodLogr_LostFocus()
    prepararLogradouro txtCodLogr
End Sub

Private Sub prepararGrupos()
    Dim Sql As String
    
    Sql = "SELECT TGL_COD_GRUPO, TGL_NOME_GRUPO as Componente" & _
            " FROM TAB_GRUPO_DETALHE_LOGRADOURO" & _
            " ORDER BY TGL_COD_GRUPO"
    grdGrupo.Preencher Bdados, Sql, 0, grdGrupo.Width
End Sub

Private Sub preencherValores(Grupo As String)
    Dim Sql As String
    
    cboDescricao.Enabled = True
    cboDescricao.Preencher Bdados, ""
    txtValor = ""
    If Trim$(Grupo) <> "" Then
        If Grupo > 20 Then
            cboDescricao.Enabled = False
        Else
            Sql = "SELECT tcl_descricao_componente, tcl_cod_componente" & _
                " FROM TAB_COMPONENTE_LOGRADOURO" & _
                " WHERE tcl_grupo=" & Grupo & _
                " ORDER BY tcl_descricao_componente"
            cboDescricao.Preencher Bdados, Sql
        End If
    End If
End Sub

Private Sub txtValor_Change()
    If cboDescricao.Enabled Then
        cboDescricao.SetarLinha txtValor, 1
    End If
End Sub

