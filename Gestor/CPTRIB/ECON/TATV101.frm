VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TATV101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TATV101.frx":0000
   ScaleHeight     =   6855
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLike 
      Height          =   375
      Left            =   45
      TabIndex        =   22
      Top             =   3480
      Width           =   8445
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2800
      Left            =   45
      TabIndex        =   17
      Top             =   705
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   4948
      Altura          =   1905
      Caption         =   " Atividade"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VB.TextBox txtDescAtiv 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "TATV101.frx":0342
         Top             =   1080
         Width           =   7215
      End
      Begin VTOcx.txtVISUAL txtAno 
         Height          =   285
         Left            =   6180
         TabIndex        =   1
         Top             =   300
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   503
         Caption         =   "Ano"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         MaxLen          =   10
      End
      Begin VTOcx.cboVISUAL cboNivel 
         Height          =   315
         Left            =   4965
         TabIndex        =   3
         Top             =   690
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   556
         Caption         =   "Nível de Instrução"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtValorISSFixo 
         Height          =   285
         Left            =   2415
         TabIndex        =   6
         Tag             =   "Valor"
         Top             =   1650
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   503
         Caption         =   "Valor ISS Fixo Anual"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtFator 
         Height          =   285
         Left            =   5610
         TabIndex        =   7
         Top             =   1650
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   503
         Caption         =   "Multiplica Por"
         Text            =   ""
         MaxLen          =   20
      End
      Begin VTOcx.fraVISUAL fraVISUAL2 
         Height          =   705
         Left            =   4230
         TabIndex        =   18
         Top             =   1980
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   1244
         Altura          =   1905
         Caption         =   " Alíquotas"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtAliquotaSPL 
            Height          =   285
            Left            =   1410
            TabIndex        =   11
            Top             =   360
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            Caption         =   "SPL"
            Text            =   ""
            Restricao       =   3
            AlinhamentoTexto=   1
            MaxLen          =   10
         End
         Begin VTOcx.txtVISUAL txtAliquotaTPPC 
            Height          =   285
            Left            =   105
            TabIndex        =   10
            Top             =   345
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            Caption         =   "TPPC"
            Text            =   ""
            Restricao       =   3
            AlinhamentoTexto=   1
            MaxLen          =   10
         End
         Begin VTOcx.txtVISUAL txtAliquotaPJ 
            Height          =   285
            Left            =   2835
            TabIndex        =   12
            Tag             =   "Alíquota PJ"
            Top             =   345
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            Caption         =   "PJ(%)"
            Text            =   ""
            Restricao       =   3
            AlinhamentoTexto=   1
            ValorMaximo     =   100
            MaxLen          =   10
            MinLen          =   1
         End
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   285
         Left            =   75
         TabIndex        =   5
         Tag             =   "Valor"
         Top             =   1650
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   503
         Caption         =   "Val. Alvara"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   420
         TabIndex        =   2
         Tag             =   "Código"
         Top             =   690
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   503
         Caption         =   "Código"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   10
      End
      Begin VTOcx.cboVISUAL cboRamo 
         Height          =   315
         Left            =   510
         TabIndex        =   9
         Top             =   2355
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   556
         Caption         =   "Ramo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboEstimativo 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Tag             =   "Estimativo"
         Top             =   1995
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   556
         Caption         =   "Estimativo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboGrupoAtividade 
         Height          =   315
         Left            =   495
         TabIndex        =   0
         Tag             =   "Grupo"
         Top             =   300
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   556
         Caption         =   "Grupo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   975
      End
   End
   Begin VTOcx.grdVISUAL lstAtv 
      Height          =   2405
      Left            =   45
      TabIndex        =   16
      Top             =   3840
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   4233
      CorBorda        =   16711680
      Caption         =   "Cadastradas"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   16711680
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   19
      Top             =   6330
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   5025
         TabIndex        =   20
         Top             =   105
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
         Left            =   3855
         TabIndex        =   13
         Top             =   105
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
         Left            =   7350
         TabIndex        =   15
         Top             =   105
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
         Left            =   6195
         TabIndex        =   14
         Top             =   105
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
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   1138
      Icone           =   "TATV101.frx":0348
   End
End
Attribute VB_Name = "TATV101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim atividade As atividade

Private Sub cboGrupoAtividade_Click()
    Dim Sql As String
    'Exit Sub
    If cboGrupoAtividade <> "" Then
        atividade.PreencheGrid lstAtv, CStr(cboGrupoAtividade.Coluna(1).Valor)
    Else
        atividade.PreencheGrid lstAtv
    End If
    'txtCodigo = ""
    txtDescAtiv = ""
    txtValor = ""
    txtFator = ""
    txtAliquotaTPPC = ""
    txtAliquotaSPL = ""
    txtAliquotaPJ = ""
End Sub
Private Sub cmdExcluir_Click()
    If lstAtv.SelectedItem Is Nothing Then Exit Sub
    lstAtv_DblClick
    If Util.Confirma("Deseja exluir " & txtDescAtiv & "?") Then
        If atividade.VerificaSeTemCadastro(txtCodigo) Then
            Util.Avisa ("Não é possivel excluir. Existe Contribuintes cadastrados.")
            Exit Sub
        Else
            If atividade.Excluir(txtCodigo) Then
                Util.Informa "Dados atualizados."
                atividade.PreencheGrid lstAtv, "tga_nome='" & cboGrupoAtividade & "'"
                Edita.LimpaCampos Me
            End If
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    atividade.PreencheGrid lstAtv
    cboGrupoAtividade.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim rs As VSRecordset
    Dim Sql As String
    Dim Valores As String
    Dim Campos As String
    Dim Grupo As Byte
    Dim NomeGrupo As String
    Dim Ano As String
    Ano = txtAno
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
        With atividade
            .Codigo = Trim(txtCodigo)
            .Nome = Trim(txtDescAtiv)
            .GrupoCodigo = cboGrupoAtividade.Coluna(1).Valor
            .Nivel = CStr(cboNivel.Coluna(1).Valor)
            .Valor = Trim(txtValor)
            .ValorISSFixoAnual = Nvl(txtValorISSFixo, 0)
            .FatorCodigo = IIf(Trim(txtFator) <> "", 2, 1)
            .FatorDescricao = IIf(Trim(txtFator) <> "", txtFator, "")
            .Estimativo = cboEstimativo.Coluna(1).Valor
            .RamoCodigo = cboRamo.Coluna(1).Valor
            .AliquotaTPPC = Nvl(txtAliquotaTPPC, 0)
            .AliquotaSPL = Nvl(txtAliquotaSPL, 0)
            .AliquotaPJ = txtAliquotaPJ
            .Ano = Nvl(txtAno, 0)
            If .Gravar Then
                'atividade.PreencheGrid lstAtv
                Edita.LimpaCampos Me
                Informa "Atividade Econômica Gravada com Sucesso."
                txtAno = Ano
                'cboGrupoAtividade.SetFocus
            End If
        End With
    Screen.MousePointer = 0
    
'    TCIS101.Tag = txtDescAtiv
'    TCIS102.Tag = txtDescAtiv
'    If Me.Tag <> "" Then
'        Screen.MousePointer = 0
'        Unload Me
'    Else
        
'    End If
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        cboGrupoAtividade.Visible = True
        cboGrupoAtividade.Tag = "Grupo Atividade"
        cboNivel.PreencherGeral Bdados, "NIVEL INSTRUÇÃO"
        cboGrupoAtividade.SetFocus
    End If
End Sub

Private Sub Form_DblClick()
Exit Sub

    Dim Obrig As New Obrigacao
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim Aux As Double
    Dim AuxStr As String
    Dim RsAux As VSRecordset
    Dim CodObrigacao As String
    Dim Conta As New ContaCorrente
    Exit Sub
    txtAliquotaPJ.Caption = 0
    
    Sql = "select cast(numeroidentificador as int) ic,  " & _
        " substring(valorcaracteristicaobjeto,1,len(valorcaracteristicaobjeto)-3)" & _
        " as Valor, textodescricao from vis_tipo_imovel where " & _
        "codigotipocaracteristicaobjetofk = 2 and  valorcaracteristicaobjeto  <> ''"
    
    Sql = "select numeroidentificador ic,  " & _
        " substring(valorcaracteristicaobjeto,1,len(valorcaracteristicaobjeto)-2)" & _
        " as Valor, codigotipocaracteristicaobjetofk from vis_tipo_imovel  " & _
        " where   valorcaracteristicaobjeto  <> ''"
    Dim Im As String
    
    If Bdados.AbreTabela(Sql, rs) Then
        rs.MoveFirst
        Do
            Im = rs!Ic
            Aux = rs!Valor
            AuxStr = rs!codigotipocaracteristicaobjetofk
            Campos = "VALOR"
            Valores = Bdados.PreparaValor(Bdados.Converte(Aux, TCDuplo))
            Bdados.AtualizaDados "vis_tipo_imovel", Valores, Campos, "numeroidentificador ='" & Im & "' and codigotipocaracteristicaobjetofk =" & AuxStr
            txtAliquotaPJ.Caption = txtAliquotaPJ.Caption + 1
            DoEvents
            
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    Avisa "acabouuuuuuuuuu!!!"
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    cboEstimativo.PreencherGeral Bdados, "SIM OU NÃO"
    Set atividade = New atividade
        atividade.PreencheCombo cboGrupoAtividade, iaGrupoAtividade
        atividade.PreencheCombo cboRamo, iaRamo
        atividade.PreencheGrid lstAtv
    
    AtualizaCabecalho lstAtv
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Set atividade = Nothing
End Sub

Private Sub lstAtv_DblClick()
    With atividade
        .Buscar lstAtv.SelectedItem
        cboGrupoAtividade.SetarLinha .GrupoCodigo, 1
        cboEstimativo.SetarLinha .Estimativo, 1
        txtCodigo = .Codigo
        txtDescAtiv = .Nome
        txtValor = Format(.Valor, "standard")
        txtValorISSFixo = .ValorISSFixoAnual
        txtFator = .FatorDescricao
        cboRamo.SetarLinha .RamoCodigo, 1
        txtAliquotaTPPC = .AliquotaTPPC
        txtAliquotaSPL = .AliquotaSPL
        txtAliquotaPJ = .AliquotaPJ
        cboNivel.SetarLinha .Nivel, 1
    End With
    If Trim(txtAno.Text) <> "" Then
        Call txtCodigo_LostFocus
    End If
End Sub

Private Sub txtAno_LostFocus()
    If Trim(txtAno) = "" Or Trim(txtCodigo) = "" Then Exit Sub
    txtCodigo_LostFocus
'    With atividade
'        If Not .Buscar(txtCodigo, , Nvl(txtAno, 0)) Then Exit Sub
'        txtValor = .m_ValoresAtividades.v_ValorAlvara
'        txtValorISSFixo = .m_ValoresAtividades.v_ValorFixoAnualISS
'        cboEstimativo.SetarLinha .Estimativo, 1
'        txtFator = .FatorDescricao
'    End With
End Sub

Private Sub txtCodigo_LostFocus()
If Trim(txtCodigo) = "" Then Exit Sub
    With atividade
        If Not .Buscar(txtCodigo, , Nvl(txtAno, 0)) Then Exit Sub
        cboGrupoAtividade.SetarLinha .GrupoCodigo, 1
        cboEstimativo.SetarLinha .Estimativo, 1
        txtCodigo = .Codigo
        txtDescAtiv = .Nome
        txtValor = .Valor
        cboNivel.SetarLinha .Nivel, 1
        txtFator = .FatorDescricao
        cboRamo.SetarLinha .RamoCodigo, 1
        txtAliquotaTPPC = .AliquotaTPPC
        txtAliquotaSPL = .AliquotaSPL
        txtAliquotaPJ = .AliquotaPJ
        txtValorISSFixo = .ValorISSFixoAnual
    End With
End Sub

Private Sub txtLike_Change()
    atividade.PreencheGrid lstAtv, , , , txtLike
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub
