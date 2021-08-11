VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TATV105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TATV105.frx":0000
   ScaleHeight     =   6855
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TATV105.frx":0342
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2520
      Left            =   45
      TabIndex        =   16
      Top             =   705
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   4445
      Altura          =   1905
      Caption         =   " Atividade"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtSubClassse 
         Height          =   285
         Left            =   5940
         TabIndex        =   24
         Tag             =   "Código"
         Top             =   690
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   503
         Caption         =   "SubClasse"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtClasse 
         Height          =   285
         Left            =   4080
         TabIndex        =   23
         Tag             =   "Código"
         Top             =   690
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         Caption         =   "Classe"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtGrupo 
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Tag             =   "Código"
         Top             =   690
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         Caption         =   "Grupo"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtDivisao 
         Height          =   285
         Left            =   390
         TabIndex        =   21
         Tag             =   "Código"
         Top             =   690
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         Caption         =   "Divisão"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   10
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
      Begin VTOcx.txtVISUAL txtValorISSFixo 
         Height          =   285
         Left            =   2415
         TabIndex        =   4
         Tag             =   "Valor"
         Top             =   1410
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
         TabIndex        =   5
         Top             =   1410
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
         TabIndex        =   17
         Top             =   1740
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   1244
         Altura          =   1905
         Caption         =   " Alíquotas"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtAliquotaSPL 
            Height          =   285
            Left            =   1410
            TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   10
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
         TabIndex        =   3
         Tag             =   "Valor"
         Top             =   1410
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   503
         Caption         =   "Val. Alvara"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtDescAtiv 
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Tag             =   "Descrição"
         Top             =   1065
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   503
         Caption         =   "Descrição"
         Text            =   ""
         MaxLen          =   70
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   420
         TabIndex        =   0
         Tag             =   "Código"
         Top             =   360
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
         TabIndex        =   7
         Top             =   2115
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
         TabIndex        =   6
         Tag             =   "Estimativo"
         Top             =   1755
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   556
         Caption         =   "Estimativo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin VTOcx.grdVISUAL lstAtv 
      Height          =   2970
      Left            =   45
      TabIndex        =   14
      Top             =   3300
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   5239
      CorBorda        =   32768
      Caption         =   "Cadastradas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   1138
      Icone           =   "TATV105.frx":2465
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   18
      Top             =   6330
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   5025
         TabIndex        =   19
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   3855
         TabIndex        =   11
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7350
         TabIndex        =   13
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6195
         TabIndex        =   12
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
End
Attribute VB_Name = "TATV105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim atividade As atividade

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
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Valores As String
    Dim Campos As String
    Dim Grupo As Byte
    Dim NomeGrupo As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
        With atividade
            .Codigo = Trim(txtCodigo)
            .Nome = Trim(txtDescAtiv)
            .GrupoCodigo = cboGrupoAtividade.Coluna(1).Valor
            .Nivel = CStr(cboNivel.Coluna(1).Valor)
            .Valor = Trim(txtValor)
            .ValorISSFixoAnual = Nvl(txtValorISSFixo, 0)
            .FatorCodigo = IIf(Trim(txtFator) <> "", 1, 0)
            .FatorDescricao = IIf(Trim(txtFator) <> "", txtFator, "")
            .Estimativo = cboEstimativo.Coluna(1).Valor
            .RamoCodigo = cboRamo.Coluna(1).Valor
            .AliquotaTPPC = Nvl(txtAliquotaTPPC, 0)
            .AliquotaSPL = Nvl(txtAliquotaSPL, 0)
            .AliquotaPJ = txtAliquotaPJ
            .Ano = Nvl(txtAno, 0)
            If .Gravar Then
                atividade.PreencheGrid lstAtv
                Edita.LimpaCampos Me
                Informa "Atividade Econômica Gravada com Sucesso."
                cboGrupoAtividade.SetFocus
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
    Dim Obrig As New Obrigacao
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim Aux As Double
    Dim RsAux As VSRecordset
    Dim CodObrigacao As String
    Dim Conta As New ContaCorrente
    Exit Sub
    txtAliquotaPJ.Caption = 0
    
    Sql = "Select * FROM TAB_CONTRIBUINTE WHERE TCI_IM IN (SELECT * FROM TAB_IM)"
    
    Dim Im As String
    
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do
            Im = Rs!TCI_IM
            On Error Resume Next
            CodObrigacao = Imposto.GeraInscMunicipal(Right(Date, 1), 11, 1)
            If CodObrigacao = "" Then
                CodObrigacao = Imposto.GeraInscMunicipal(Right(Date, 1), 11, 1)
            End If
            If CodObrigacao <> "" Then
                Valores = Bdados.PreparaValor(CodObrigacao)
                Campos = "TCI_IM"
                Bdados.AtualizaDados "TAB_CONTRIBUINTE", Valores, Campos, "TCI_INSCRICAO_ANTERIOR ='" & Rs!IC_ANTERIOR & "'"
                txtAliquotaPJ.Caption = txtAliquotaPJ.Caption + 1
                DoEvents
            Else
                Campos = "TCI_IM"
            End If
            Rs.MoveNext
        Loop While Not Rs.EOF
    End If
    Avisa "acabouuuuuuuuuu!!!"
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
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

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub
