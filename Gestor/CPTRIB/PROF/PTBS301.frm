VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form PTBS301 
   BackColor       =   &H00DDF1FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PTBS301.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMunicipio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1110
      Width           =   3915
   End
   Begin VB.ComboBox cboTipLogr 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "PTBS301.frx":08CA
      Left            =   1080
      List            =   "PTBS301.frx":08CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Municipio"
      Top             =   1920
      Width           =   1245
   End
   Begin VB.TextBox txtLogr 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   2415
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1920
      Width           =   4245
   End
   Begin VB.ComboBox cboBairro 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "PTBS301.frx":08CE
      Left            =   1080
      List            =   "PTBS301.frx":08D0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Municipio"
      Top             =   1530
      Width           =   3915
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1349
      CorFundo        =   14545407
      CorFrente       =   255
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   3120
      Left            =   75
      TabIndex        =   4
      Top             =   2640
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   5503
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Município"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   195
      TabIndex        =   13
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Tabela de Logradouros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   2355
      Width           =   6615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   570
      TabIndex        =   11
      Top             =   1980
      Width           =   405
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Logradouro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   795
      Width           =   6615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   450
      TabIndex        =   9
      Top             =   1590
      Width           =   510
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   5610
      TabIndex        =   6
      Top             =   5835
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      BackColor       =   12632256
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Sair"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdSalvar 
      Height          =   435
      Left            =   4485
      TabIndex        =   5
      Top             =   5835
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      BackColor       =   12632256
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Gravar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   435
      Left            =   2235
      TabIndex        =   7
      Top             =   5835
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      BackColor       =   12632256
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Novo"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdCancelar 
      Height          =   435
      Left            =   3360
      TabIndex        =   8
      Top             =   5835
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      BackColor       =   12632256
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Cancelar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
End
Attribute VB_Name = "PTBS301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CodMunicSelec As Long
Dim CodigoBairro As Long
Dim Bairro As String
Dim CodigoTipLogr As Long
Dim CodigoLogr As Long
Dim Logr As String
Sub MontaGrids()
    Dim SQL As String
    
    SQL = "SELECT tlg_cod_logradouro as Código, TTL_NOME AS Tipo, tlg_nome as Logradouro" & _
            " FROM TAB_LOGRADOURO ,TAB_TIPO_LOGR where tlg_ttl_cod_TIP_logr = ttl_cod_tip_Logr" & _
            " and tlg_tba_cod_bairro =" & CodigoBairro & _
            " AND tlg_tmu_cod_municipio=" & CodMunicSelec
    Util.MontaGrid Bdados, Grid, SQL, 900, 1200, 4000
End Sub

Private Function GeraCodLogr() As Long
    GeraCodLogr = Bdados.BuscaCodigo("SELECT MAX(tlg_cod_logradouro)+1 FROM TAB_LOGRADOURO")
    If GeraCodLogr = 0 Then GeraCodLogr = 1
End Function

Private Sub cboBairro_Click()
    CodigoBairro = Bdados.BuscaCodigo("select tba_cod_bairro from tab_bairro where TBA_TMU_COD_MUNICIPIO=" & CodMunicSelec & " and tba_nome = '" & cboBairro & "'")
    MontaGrids
End Sub
Private Sub cboMunicipio_Click()
    CodMunicSelec = Bdados.BuscaCodigo("SELECT TMU_COD_MUNICIPIO FROM TAB_MUNICIPIO WHERE TMU_NOME='" & cboMunicipio & "'")
    Edita.AtualizaCombo Bdados, cboBairro, "select tba_nome from tab_bairro WHERE TBA_TMU_COD_MUNICIPIO=" & CodMunicSelec
    Grid.ListItems.Clear
End Sub

Private Sub cboTipLogr_Click()
    If cboTipLogr <> "" Then
        CodigoTipLogr = Bdados.BuscaCodigo("select ttl_cod_tip_logr from tab_tipo_logr where ttl_nome = '" & cboTipLogr & "'")
    End If
End Sub

Private Sub cmdCancelar_Click()
    cboMunicipio.Enabled = True
    cboBairro.Enabled = True
    cboTipLogr.ListIndex = -1
    txtLogr = ""
End Sub

Private Sub cmdNovo_Click()
    cboTipLogr.ListIndex = -1
    txtLogr = ""
    CodigoLogr = -1
    cboTipLogr.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim SQL As String
    Dim Campos As String
    Dim RS As Object
    If Edita.CriticaCampos(Me) Then
        Campos = "tlg_tmu_cod_municipio,tlg_cod_logradouro,tlg_tba_cod_bairro,tlg_ttl_cod_tip_logr,tlg_nome,tlg_secao,tlg_ttr_cod_trecho"
'        CodigoBairro = BuscaCodigo("select tba_cod_bairro from tab_bairro where tba_nome = '" & cboBairro & "'")
'        CodigoTipLogr = BuscaCodigo("select ttl_cod_tip_Logr from tab_tipo_logr where ttl_nome = '" & Me.cboTipLogr & "'")
        If CodigoLogr = -1 Then
            CodigoLogr = GeraCodLogr
            Valores = Bdados.PreparaValor(CodMunicSelec, CodigoLogr, CodigoBairro, CodigoTipLogr, txtLogr, 1, 1)
            Call Bdados.InsereDados("Tab_Logradouro", Valores, Campos)
        ElseIf Trim(txtLogr) = "" Then
            If Util.Confirma("Confirma a exclusão do logradouro " & Logr & "?") Then
                Call Bdados.DeletaDados("Tab_Logradouro", "tlg_cod_logradouro = " & CodigoLogr)
                SQL = "Select tlg_cod_logradouro  from Tab_Logradouro where tlg_cod_logradouro =" & CodigoLogr
                If Bdados.AbreTabela(SQL, RS) Then
                    Call Util.Avisa("Existem registros relacionados com o logradouro " & Logr & ". Exclusão cancelada.")
                    Exit Sub
                End If
                Bdados.FechaTabela RS
            Else
                CodigoLogr = -1
                Exit Sub
                Screen.MousePointer = 0
            End If
        Else
            Valores = Bdados.PreparaValor(CodMunicSelec, CodigoLogr, CodigoBairro, CodigoTipLogr, txtLogr, 1)
            Call Bdados.AtualizaDados("Tab_Logradouro", Valores, Campos, "tlg_tba_cod_bairro = " & CodigoBairro & " and tlg_ttl_cod_tip_logr = " & CodigoTipLogr)
        End If
        Call Util.Informa("Transação completada.")
        MontaGrids
        txtLogr = ""
        cboTipLogr.SetFocus
    End If
    CodigoLogr = -1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"
    Edita.AtualizaCombo Bdados, cboMunicipio, "SELECT TMU_NOME FROM TAB_MUNICIPIO ORDER BY TMU_NOME"
    If cboMunicipio.ListCount Then cboMunicipio.ListIndex = 0
    Edita.AtualizaCombo Bdados, cboTipLogr, "select ttl_nome from Tab_Tipo_Logr"
    CodigoLogr = -1
    Screen.MousePointer = 0
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Util.OrdenaGrid(Grid, ColumnHeader)
End Sub

Private Sub Grid_DblClick()
    On Error GoTo Trata
    If Not Grid.SelectedItem Is Nothing Then
        cboMunicipio.Enabled = False
        cboBairro.Enabled = False
        CodigoLogr = Grid.SelectedItem
        cboTipLogr = Grid.SelectedItem.ListSubItems.Item(1).Text
        txtLogr = Grid.SelectedItem.ListSubItems.Item(2).Text
        Logr = Grid.SelectedItem.ListSubItems.Item(2).Text
    End If
    txtLogr.SetFocus
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Grid_DblClick
End Sub


