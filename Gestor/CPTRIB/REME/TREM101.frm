VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.0#0"; "VTControles.ocx"
Begin VB.Form TREM101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
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
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TREM101.frx":0000
   ScaleHeight     =   3210
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox optNumGuias 
      Appearance      =   0  'Flat
      Caption         =   "Primeiras"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1590
      Width           =   1245
   End
   Begin Threed.SSPanel pnlProgresso 
      Height          =   225
      Left            =   1950
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   397
      _Version        =   196610
      ForeColor       =   -2147483645
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      FloodType       =   1
      FloodColor      =   -2147483646
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtExercicio 
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Tag             =   "Exercicio"
      Top             =   690
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   503
      Caption         =   "Exercicio"
      Text            =   ""
      Restricao       =   2
      AlinhamentoTexto=   2
      MaxLen          =   4
   End
   Begin Cabecalho.cabVISUAL cabCabecalho 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   1138
      Formulario      =   "CODIGO"
      Icone           =   "TREM101.frx":0C42
   End
   Begin Cabecalho.rodVISUAL rodRodape 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   6
      Top             =   2685
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   926
      CorFundo        =   -2147483632
      CorFrente       =   -2147483633
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   405
         Left            =   2520
         TabIndex        =   23
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   714
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   4650
         TabIndex        =   4
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdGravar 
         Height          =   405
         Left            =   3570
         TabIndex        =   3
         Top             =   90
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   714
         Caption         =   "&Gravar"
         Acao            =   3
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
   End
   Begin VB.Frame fraResumo 
      Appearance      =   0  'Flat
      Caption         =   "Remessa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1635
      Index           =   3
      Left            =   1950
      TabIndex        =   8
      Top             =   990
      Width           =   3525
      Begin VB.Label lblArquivo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   840
         TabIndex        =   20
         Top             =   1410
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   150
         TabIndex        =   19
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   840
         TabIndex        =   18
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label lblGuias 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   840
         TabIndex        =   17
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   840
         TabIndex        =   16
         Top             =   705
         Width           =   45
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   840
         TabIndex        =   15
         Top             =   465
         Width           =   45
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   345
         TabIndex        =   14
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guias"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   330
         TabIndex        =   13
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   420
         TabIndex        =   12
         Top             =   705
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   375
         TabIndex        =   11
         Top             =   465
         Width           =   360
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   840
         TabIndex        =   10
         Top             =   210
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   9
         Top             =   210
         Width           =   570
      End
   End
   Begin VTOcx.txtVISUAL txtNumGuias 
      Height          =   285
      Left            =   600
      TabIndex        =   21
      Top             =   1830
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
      Restricao       =   2
      AlinhamentoTexto=   2
      MaxLen          =   4
   End
   Begin VTOcx.txtVISUAL txtNumArquivo 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   1200
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      Caption         =   "Em"
      Text            =   ""
      Restricao       =   2
      AlinhamentoTexto=   2
      MaxLen          =   4
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "arquivo(s)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   24
      Top             =   1230
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "guias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1470
      TabIndex        =   22
      Top             =   1845
      Width           =   450
   End
End
Attribute VB_Name = "TREM101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intNumRemessa As Integer
Private intTipoRemessa As Integer
Private CodImposto As String
Private qtdRegistrosArquivo As Double

Private Sub cmdGravar_Click()
    Dim NomeArquivo As String
    Dim Arquivo As IGC702.cRemessa
    Dim i As Integer
    
    If Edita.CriticaCampos(Me) Then
        Screen.MousePointer = vbHourglass
        pnlProgresso.Visible = True
        qtdRegistrosArquivo = 0
        For i = 1 To Nvl(txtNumArquivo, 1)
            'A Remessa
            Set Arquivo = prepararRemessa()
            
            'B Convenente
            prepararConvenente Arquivo
            
            'C Convenios
            prepararConvenio Arquivo
            
            'D Definicao Contribuinte
            prepararTipoContribuinte Arquivo
            
            'E Definicao Recebimento
            prepararTipoRecebimento Arquivo
            
            'F Definicao Objeto
            prepararObjeto Arquivo
            
            'K Receitas
            prepararReceitas Arquivo
            
            'L Opcoes de Pagamento
            prepararCotas Arquivo
            
            'S Guias
            prepararGuias Arquivo
            
            Screen.MousePointer = vbNormal
            NomeArquivo = Format(intNumRemessa, "0000") & txtExercicio & ".rem"
            If Arquivo.Gravar(App.Path & "\REMESSAS\" & NomeArquivo) Then
                lblArquivo = "...\REMESSAS\" & NomeArquivo
                lblGuias = Format(Arquivo.Guias.Quantidade, "#,##0")
                lblValor = Format(Arquivo.Guias.Soma, "currency")
                
                Dim Campos As String, Valores As String
                Campos = "TRI_NUMERO, TRI_DATA_GERACAO, TRI_TIPO, TRI_QUANTIDADE_GUIAS, TRI_VALOR_TOTAL, TRI_EXERCICIO"
                Valores = Bdados.PreparaValor(intNumRemessa, Format(Now, "dd/mm/yyyy"), intTipoRemessa, Arquivo.Guias.Quantidade, Arquivo.Guias.Soma, txtExercicio)
                Bdados.GravaDados "TAB_REMESSA_IPTU", Valores, Campos, "TRI_NUMERO=" & intNumRemessa
                
            End If
        Next
        Avisa "Arquivo(s) gerado(s) com sucesso."
        cmdSair.SetFocus
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    optNumGuias = False
    lblArquivo = ""
    lblNumero = ""
    lblData = ""
    lblTipo = ""
    lblGuias = ""
    lblValor = ""
    pnlProgresso.FloodPercent = 0
    txtExercicio.SetFocus
End Sub

Private Sub cmdSair_Click()
    Screen.MousePointer = vbNormal
    Unload Me
End Sub

Private Sub Form_Load()
    cabCabecalho.Exibir Bdados, Me.Name, App.Path
    rodRodape.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
    txtExercicio = Format(Now, "yyyy")
    txtNumArquivo = "1"
End Sub

Private Function prepararRemessa() As IGC702.cRemessa
    On Error GoTo trata
    Dim Remessa As New IGC702.cRemessa
    Dim rs As VSRecordset
    
    intNumRemessa = Bdados.BuscaCodigo("SELECT MAX(TRI_NUMERO) FROM TAB_REMESSA_IPTU") + 1
    If Bdados.AbreTabela("SELECT * FROM TAB_PARAMETRO_REMESSA", rs) Then
        With Remessa
            .numConvenio = "" & rs!TPR_NUMERO_CONVENIO
            '.Layout = "1.00.01" '"" & rs!TPR_LAYOUT
            .numRemessa = intNumRemessa
            .dataGeracao = Now
            .tipoFormulario = "" & rs!TPR_TIPO_FORMULARIO
            intTipoRemessa = "" & rs!TPR_TIPO_REMESSA
            .tipoRemessa = intTipoRemessa
            .postagemBB = True
            .cnpjConvenente = "" & rs!TPR_CNPJ_CONVENENTE
        End With
        Set prepararRemessa = Remessa
        
        lblNumero = intNumRemessa
        lblData = Format(Now, "dd/mm/yyyy")
        lblTipo = IIf((intTipoRemessa = 1), "Teste", "Produção")

    End If
    Bdados.FechaTabela rs
    Exit Function
    
trata:
    Erro Err.Description
End Function

Private Sub prepararConvenente(ByRef Remessa As cRemessa)
    On Error GoTo trata
    Dim rs As VSRecordset
    
    If Bdados.AbreTabela("SELECT * FROM TAB_PARAMETRO_REMESSA", rs) Then
        With Remessa.Convenente
            .Nome = "" & rs!TPR_CONVENENTE
            .Endereco = "" & rs!TPR_ENDERECO_CONVENENTE
            .CEP = "" & rs!TPR_CEP_CONVENENTE
            .Cidade = "" & rs!TPR_CIDADE_CONVENENTE
            .Bairro = "" & rs!TPR_BAIRRO_CONVENENTE
            .UF = "" & rs!TPR_UF_CONVENENTE
        End With
    End If
    Bdados.FechaTabela rs
    Exit Sub
    
trata:
    Erro Err.Description
End Sub

Private Sub prepararConvenio(ByRef Remessa As cRemessa)
    On Error GoTo trata
    Dim rs As VSRecordset
    Dim Convenio As New IGC702.cConvenio
    
    Remessa.Convenios.Limpar
    If Bdados.AbreTabela("SELECT * FROM TAB_PARAMETRO_REMESSA", rs) Then
        With Convenio
            .segmentoFEBRABAN = segPrefeitura
            .identificacaoFEBRABAN = "" & rs!TPR_IDENTIFICACAO_FEBRABAN
            .Moeda = IGC702.enuMoeda.moeReal
            .numDecimais = 2
            .receberVencidos = "" & rs!TPR_RECEBER_VENCIDO
            '.formatoData = "" & rs!TPR_FORMATO_VENCTO
        End With
        Remessa.Convenios.Adicionar Convenio
    End If
    Bdados.FechaTabela rs
    Exit Sub
    
trata:
    Erro Err.Description
End Sub

Private Sub prepararTipoContribuinte(ByRef Remessa As cRemessa)
    On Error GoTo trata
    
    With Remessa.TipoContribuinte
        .Denominacao = "CONTRIBUINTE"
        .Sigla = "IM"
        .signSigla = "INSCRICAO MUNICIPAL"
        .tipoIdentificador = "A"
        .tamIdentificador = 12
        .mascIdentificador = String(9, "9") & "-" & String(2, "9")
    End With
    Exit Sub
    
trata:
    Erro Err.Description
End Sub

Private Sub prepararTipoRecebimento(ByRef Remessa As cRemessa)
    On Error GoTo trata
    Dim rs As VSRecordset
    
    If Bdados.AbreTabela("SELECT * FROM TAB_PARAMETRO_REMESSA", rs) Then
        With Remessa.TipoRecebimento
            .Nome = "IMPOSTO PREDIAL E TERRITORIAL URBANO"
            .Sigla = "IPTU"
            .unidadeConvenente = "" & rs!TPR_UNIDADE_CONVENENTE
            .denominacaoExercicio = "ANO"
            .Exercicio = txtExercicio
            .Titulo = "COMPOSICAO"
            .identificacaoBarra = IGC702.enuIdentificacaoBarra.idObjeto
            .zerarValor = False
            .checarData = True
            .checarValor = True
        End With
    End If
    Bdados.FechaTabela rs
    Exit Sub
    
trata:
    Erro Err.Description
End Sub

Private Sub prepararReceitas(ByRef Remessa As cRemessa)
    On Error GoTo trata
    Dim Receita As New IGC702.cReceita
    
    Remessa.Receitas.Limpar
    With Receita
        .Codigo = CodImposto
        .Descricao = "IPTU"
        .Aliquota = (Imposto.BuscaAliquota(CodImposto, txtExercicio) * 100) * 100
    End With
    Remessa.Receitas.Adicionar Receita
    Exit Sub
    
trata:
    Erro Err.Description
End Sub

Private Sub prepararCotas(ByRef Remessa As cRemessa)
    On Error GoTo trata
    Dim rs As VSRecordset
    Dim Opcao As New IGC702.cPagamento
    
    Remessa.Pagamentos.Limpar
    With Opcao
        .Denominacao = "COTA UNICA"
        .Tipo = "C"
        .Numero = "0"
        .Vencimento = Imposto.BuscaDataVencimento(CodImposto, txtExercicio)
        .Incidencia = txtExercicio
    End With
    If Bdados.AbreTabela("Select TGE_NOME from tab_geral where TGE_TIPO = 755 and TGE_CODIGO > 0", rs) Then
        Opcao.Desconto = Nvl("" & rs!TGE_NOME, 0)
    End If
    Remessa.Pagamentos.Adicionar Opcao
    Bdados.FechaTabela rs
    Exit Sub
    
trata:
    Erro Err.Description
End Sub

Private Sub prepararGuias(ByRef Remessa As cRemessa)
    On Error GoTo trata
    Dim rs As VSRecordset
    Dim Guia As IGC702.cGuia
    
    Dim i, t As Integer
    
    pnlProgresso.Visible = True
    
    Remessa.Guias.Limpar
    'If qtdRegistrosArquivo = 0 Then
        If optNumGuias Then
            t = Nvl(txtNumGuias, 1)
        Else
            If Bdados.AbreTabela("select count(*) from tab_geracao_tributo where tgt_tip_cod_imposto='" & CodImposto & "' AND tgt_periodo=" & txtExercicio & " AND tgt_valor_tributo>0", rs) Then
                t = rs(0)
            End If
            Bdados.FechaTabela rs
        End If
        qtdRegistrosArquivo = t / Nvl(txtNumArquivo, 1) '+ (t Mod Nvl(txtNumArquivo, 1))
    'End If
    
    i = 0
    If Bdados.AbreTabela("select " & IIf(optNumGuias, "top " & t, "") & " * from tab_geracao_tributo where tgt_tip_cod_imposto='" & CodImposto & "' AND tgt_periodo=" & txtExercicio & " AND tgt_valor_tributo>0 AND tgt_tim_ic NOT IN (SELECT TRD_TIM_IC FROM TAB_DETALHE_REMESSA_IPTU, TAB_REMESSA_IPTU WHERE TRD_TRI_NUMERO=TRI_NUMERO AND TRI_EXERCICIO='" & txtExercicio & "')", rs) Then
        Do While Not rs.EOF
            Set Guia = New IGC702.cGuia
            'S Guia
            With Guia
                .Emissao = "" & rs!tgt_data_geracao
                .Validade = "" & rs!tgt_data_vencimento
                .enderecoCorrespondencia = "O"
            End With
            'T Contribuinte
            setarContribuinte rs!tgt_im, Guia
            
            'U Valores
            With Guia.Valores
                .Valor01 = "" & rs!tgt_valor_tributo
            End With
            
            'V Imovel
            setarObjeto rs!tgt_tim_ic, Guia
            
            'Y Codigo de Barras
            With Guia.Barcode
                .tipoPagamento = "C"
                .numPagamento = 0
                .Vencimento = "" & rs!tgt_data_vencimento
                .Valor = "" & rs!tgt_valor_tributo
'                .gerarLinhaDigitavel rs!tgt_tim_ic
            End With
            
            Remessa.Guias.Adicionar Guia
            
            Dim Campos As String, Valores As String
            Campos = "TRD_TRI_NUMERO,TRD_TIM_IC,TRD_VALOR"
            Valores = Bdados.PreparaValor(intNumRemessa, rs!tgt_tim_ic, rs!tgt_valor_tributo)
            Bdados.GravaDados "TAB_DETALHE_REMESSA_IPTU", Valores, Campos, "TRD_TRI_NUMERO=" & intNumRemessa & " AND TRD_TIM_IC='" & rs!tgt_tim_ic & "'"
            
            
            i = i + 1
            lblGuias = i & "/" & t
            pnlProgresso.FloodPercent = IIf((i * Nvl(txtNumArquivo, 1) / t) * 100 > 100, 100, (i * Nvl(txtNumArquivo, 1) / t) * 100)
            lblValor = Format$(Remessa.Guias.Soma, "currency")
            If i >= qtdRegistrosArquivo Then Exit Do
            DoEvents
            rs.MoveNext
        Loop
    End If
    Bdados.FechaTabela rs
    Exit Sub
    
trata:
    Erro Err.Description
End Sub

Private Sub setarContribuinte(Im As String, ByRef Guia As cGuia)
    On Error GoTo trata
    Dim rs As VSRecordset
    
    If Bdados.AbreTabela("SELECT * FROM TAB_CONTRIBUINTE WHERE TCI_IM='" & Im & "'", rs) Then
        With Guia.Contribuinte
            .Nome = "" & rs!tci_nome
            .Inscricao = Im
            .Tipo = "" & rs!tci_tnj_cod_natureza
            .Cnpj = "" & rs!tci_cgc_cpf
            .Endereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & ", " & rs!tci_numero & " " & rs!tci_complemento
            .CEP = "" & rs!tci_cep
            .Cidade = "" & rs!tci_cidade
            .Bairro = "" & rs!tci_bairro
            .UF = "" & rs!tci_UF
        End With
    End If
    Bdados.FechaTabela rs
    Exit Sub
    
trata:
    Erro Err.Description
    'Resume
End Sub

Private Sub setarObjeto(IC As String, ByRef Guia As cGuia)
    On Error GoTo trata
    Dim rs As VSRecordset
    
    If Bdados.AbreTabela("SELECT * FROM VIS_IMOVEL WHERE TIM_IC='" & IC & "'", rs) Then
        With Guia.IdentImovel
            .Identificacao = Trim$(IC)
            .Localizacao = ("" & rs!TTL_NOME) & " " & ("" & rs!tlg_nome) & ", " & ("" & rs!tim_numero)
            .CEP = "" & rs!tim_cep
            .Cidade = Aplicacoes.Municipio
            .Bairro = "" & rs!TBA_NOME
            .UF = "MA"
            .Complemento = "" & rs!tim_complemento
        End With
        Guia.ValorCaracteristica.Conteudo01 = rs!tim_valor
    End If
    Bdados.FechaTabela rs
    Exit Sub
    
trata:
    Erro Err.Description
End Sub
Private Sub prepararObjeto(ByRef Remessa As cRemessa)
    On Error GoTo trata
    Dim Caracteristica As New cCaracteristica
    
    With Remessa.Imovel
        .Denominacao = "IMÓVEL"
        .Sigla = "IC"
        .Significado = "INSCRICAO CADASTRAL"
        .Tipo = "A"
        .Tamanho = 10
        .Mascara = String(8, "9") & "-" & String(1, "9")
        .tituloCaracteristicas = "CARACTERISTICAS"
        
        Caracteristica.Denominacao = "VALOR VENAL"
        .Caracteristicas.Adicionar Caracteristica
    End With
    Exit Sub
    
trata:
    Erro Err.Description
End Sub

Private Sub optNumGuias_Click()
    txtNumGuias.Enabled = optNumGuias
    If optNumGuias Then
        txtNumGuias.SetFocus
    Else
        txtNumGuias = ""
    End If
End Sub

Private Sub txtNumGuias_LostFocus()
    cmdGravar.SetFocus
End Sub
