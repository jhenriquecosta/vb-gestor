VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCOB403 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   14
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCOB403.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   795
      Left            =   60
      TabIndex        =   7
      Top             =   690
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1402
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         Caption         =   "Pré-definidos"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   8
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7770
         TabIndex        =   3
         Top             =   405
         Width           =   1230
      End
      Begin VB.TextBox txtImovel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   11
         TabIndex        =   1
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtAno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cboBairro 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "TCOB403.frx":2123
         Left            =   2460
         List            =   "TCOB403.frx":2130
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Bairro"
         Top             =   330
         Width           =   5085
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   2460
         TabIndex        =   8
         Top             =   120
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Bairro"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ano"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   1140
         TabIndex        =   12
         Top             =   120
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Imóvel"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.PictureBox lstIptu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5250
      Left            =   75
      ScaleHeight     =   5220
      ScaleWidth      =   9075
      TabIndex        =   4
      Top             =   1545
      Width           =   9105
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2820
      Width           =   375
   End
   Begin VB.Timer tmr 
      Interval        =   10
      Left            =   3300
      Top             =   3510
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2730
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   11
      Top             =   2910
      Visible         =   0   'False
      Width           =   795
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   1138
      Icone           =   "TCOB403.frx":2151
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   8055
      TabIndex        =   6
      Top             =   6855
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdBuscar 
      Height          =   375
      Left            =   6885
      TabIndex        =   5
      Top             =   6870
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Buscar"
      Acao            =   5
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCOB403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As VSImposto

Dim TotalImposto As Double
Dim CodTributo  As String
Dim Pagamento As Double

Public Sub GeraDam(InscMuni As String, RazaoSocial As String, CodPagamento As String, EnderecoImovel As String, DataVencimento As String, Ic As String, CPFCNPJ As String, EnderecoContrib As String, Exercicio As String, Imposto As String, Multa As String, Juros As String, CodImposto As String, ValorMetro As String, BaseDeCalculo As String, AreaConstruida As String, AreaTotal As String, NomeImposto As String, Parcela As String, Taxas As String, Aliquota As String, Desconto As Double)
    Dim a As Byte
    Dim Sql As String
    Dim TotalUnico  As Double
    Dim RsDesconto As VSRecordset
    Static TaxaParcela As Double
    Dim cLSImposto As New VSImposto
    Dim Cobranca As New VSCobranca
    If Not IsNumeric(Parcela) Then
        TotalUnico = Imposto
        TaxaParcela = Taxas
    Else
        Sql = "SELECT * from Tab_Parametro_Parcelamento"
        If Bdados.AbreTabela(Sql, RsDesconto) Then
            TotalUnico = Imposto * RsDesconto!tpp_max_cotas
        End If
    End If
    Bdados.FechaTabela RsDesconto
    
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TDAMBarra.rpt") Then Exit Sub
        .Formulas "inscmunicipal", InscMuni
        .Formulas "nome", RazaoSocial
        .Formulas "documento", CodPagamento
        If Trim(EnderecoImovel) <> "" Then
            .Formulas "localizacao", EnderecoImovel & " - " & UCase(Aplicacoes.Municipio) & " MA'"
        End If
        .Formulas "datavencimento", DataVencimento
        .Formulas "codigoimovel", Ic
        .Formulas "cpf/cnpj", CPFCNPJ
        .Formulas "endereco", EnderecoContrib
        .Formulas "exercicio", IIf(Len(Exercicio) = 4, Exercicio, Left(Exercicio, 2) & "/" & Right(Exercicio, 4))
        .Formulas "ValorTributo", Format(Imposto, Const_Monetario)
        .Formulas "ValorMulta", Format(Multa, Const_Monetario)
        .Formulas "ValorJuros", Format(Juros, Const_Monetario)
        TotalImposto = CDbl(Imposto) + CDbl(Multa) + CDbl(Juros)
        .Formulas "ValorTotal", Format(CDbl(Imposto) + CDbl(Multa) + CDbl(Juros), Const_Monetario)
        CodTributo = CodImposto
        Pagamento = CodPagamento
        .Formulas "CodigoTributo", CodImposto
        .Formulas "OBSERVACAO ", "Valor m2 Terreno(R$): " & Format(ValorMetro, Const_Monetario) & "     -       Valor Venal(R$): " & Format(BaseDeCalculo, Const_Monetario) & Space(41) & " Taxas de Servicos Urbanos(TSU): " & Format(TaxaParcela, Const_Monetario) & "     -      Valor IPTU(R$) : " & Format(TotalUnico, Const_Monetario) & _
        IIf(Not IsNumeric(Parcela) And Desconto > 0, Space(32) & "DESCONTO DE " & Desconto & "% EM COTA ÚNICA'", "")
        
        .Formulas "NUM_NOTAS ", "NÃO RECEBER APÓS DATA DE VENCIMENTO'"
        
        .Formulas "BASECALCULO", Format(BaseDeCalculo, Const_Monetario)
        .Formulas "DESCTRIBUTO", NomeImposto
        If Trim(Parcela) <> "" Then .Formulas "Parcela", Parcela
        .Formulas "PREFEITURA", Temp.PegaParametro(Bdados, "CLIENTE")
        .Formulas "ObsMaterial", "Área Total do Imóvel: " & Format(AreaTotal, Const_Monetario) & "m2     -       Área Total Construída: " & Format(AreaConstruida, Const_Monetario) & "m2'"
'        .Formulas "LinhaDigitavel", Cobranca.GeraCodBarra(CDbl(CodPagamento), BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto ='" & CodImposto & "'"), TotalImposto, PicBarra)
        .Formulas "EMISSAO", cLSImposto.BuscaDataGeracaoDam(CDbl(CodPagamento))
        
        '.Connect = Bdados.BDSistema.Connect
        
        .CopiasDetalhes = 3
        .Imprimir
    End With
    Bdados.FechaTabela RsDesconto
End Sub

Private Sub cmdImprime_Click()
    Dim i As Double
    Dim j As String
    Dim Lista As Object
    
End Sub

Private Sub cmdBuscar_Click()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim RsComp As VSRecordset
    Dim Operador As String
    Dim Aux As Byte
    Dim AreaTotal As String
    Dim AreaConstruida As String
    Dim Desconto As Double
    Dim RsDesconto As VSRecordset
    Dim RsAux As VSRecordset
    Dim ValorMetro As Double
    Dim NomeLogr As String
    Dim Logr As String
    Dim CodigoLogr As String
    Dim Bairro As String
    Dim Cobranca As New VSCobranca
    Dim Conta As New ContaCorrente
    Dim CodImposto As VSImposto
    Const ValorMin As Double = 5
    
    Screen.MousePointer = 11
    Aux = 0
    Sql = "SELECT * FROM VIS_BOLETO "
    If Trim(txtImovel) <> "" Then
        Aux = 1
        Sql = Sql & " where ic = '" & Trim(txtImovel) & "'"
    Else
        If Trim(cboBairro) <> "" Then
            
            Sql = Sql & IIf(Aux = 1, " and ", " where ") & " tba_nome = '" & Trim(cboBairro.Text) & "'"
            Aux = 1
        End If
    End If
    If Trim(txtAno) <> "" Then
        Sql = Sql & IIf(Aux = 1, " and ", " where ") & " periodo = " & Trim(txtAno)
    End If
    Sql = Sql & " ORDER BY IC ASC,Parcela asc"
    If chk Then
        Sql = "SELECT * FROM VIS_BOLETO "
        Sql = Sql & " where ic in (select ic from ic where impresso =false)"
        Sql = Sql & " ORDER BY IC ASC,Parcela asc"
    End If
    If Bdados.AbreTabela(Sql, rs) Then
        MontaGrid Bdados, lstIptu, Sql, 1000, 1200, 1000, 3000, 1500, 1000, 1000, 1000
        DoEvents
        rs.MoveFirst
        Sql = "Select TGE_NOME from tab_geral where TGE_TIPO = 755 and TGE_CODIGO > 0"
        If Bdados.AbreTabela(Sql, RsDesconto) Then
            Desconto = RsDesconto(0)
        End If
        Do While Not rs.EOF
            'If Rs!IC = "31001903-8" Then
            
                Sql = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & rs!Ic & _
                    "' and (tdi_tco_cod_componente=110 or tdi_tco_cod_componente=108)"
                If Bdados.AbreTabela(Sql, RsComp) Then
                    RsComp.MoveFirst
                    Do While Not RsComp.EOF
                        If RsComp(0) = 110 Then
                            AreaTotal = RsComp(1)
                        ElseIf RsComp(0) = 108 Then
                            AreaConstruida = RsComp(1)
                        End If
                        RsComp.MoveNext
                    Loop
                End If
                Bdados.FechaTabela RsComp
                
                Sql = "select tvl_valor  as ValorMetro from TAB_VALOR_TERRENO where tvl_tlg_cod_logradouro='" & rs!tlg_cod_logradouro & "'"
                If Bdados.AbreTabela(Sql, RsAux) Then
                    ValorMetro = RsAux!ValorMetro
                End If
                Bdados.FechaTabela RsAux
                Sql = "select ttl_nome as Logr,TLG_NOME AS Nome  from tab_logradouro,tab_tipo_logr where tlg_cod_logradouro='" & rs!tlg_cod_logradouro & "' and tlg_ttl_cod_tip_logr = ttl_cod_tip_logr and tlg_tmu_cod_municipio =" & Aplicacoes.Codigo_Municipio
                If Bdados.AbreTabela(Sql, RsAux) Then
                    Logr = RsAux!Logr
                    NomeLogr = RsAux!Nome
                End If
                Bdados.FechaTabela RsAux
                'COLINAS
                'Conta.MovimentaContaContribuinte Rs!Contribuinte, Rs!IC, ClsImposto.BuscaCodIptu, Rs!Periodo, IIf(Rs!Parcela = 0, EtcCreditoTributario, EtcParcelamento), Rs!Parcela, Date
                If rs!Imposto >= ValorMin Then
                    GeraDam rs!Im, rs!Contribuinte, rs!CodPago, Logr & " " & NomeLogr & " " & rs!Num & " " & rs!TBA_NOME, _
                    rs!vencimento, rs!Ic, "" & rs!cgc_cpf, rs!Logradouro & " " & rs!nome_logr & " " & rs!Numero & " " & rs!compl & " " & rs!bairro_contrib & " " & rs!Cep & " " & rs!Cidade & " " & rs!Uf, _
                    rs!Periodo, rs!Imposto, rs!Multa, rs!Juros, rs!Sigla, CStr(ValorMetro), rs!valor_venal, AreaConstruida, AreaTotal, rs!nome_imposto, IIf(rs!Parcela = 0, "UNICA", rs!Parcela), rs!taxa, 1, Desconto
                ElseIf rs!Parcela = 0 Then
                    GeraDam rs!Im, rs!Contribuinte, rs!CodPago, Logr & " " & NomeLogr & " " & rs!Num & " " & rs!TBA_NOME, _
                    rs!vencimento, rs!Ic, "" & rs!cgc_cpf, rs!Logradouro & " " & rs!nome_logr & " " & rs!Numero & " " & rs!compl & " " & rs!bairro_contrib & " " & rs!Cep & " " & rs!Cidade & " " & rs!Uf, _
                    rs!Periodo, rs!Imposto, rs!Multa, rs!Juros, rs!Sigla, CStr(ValorMetro), rs!valor_venal, AreaConstruida, AreaTotal, rs!nome_imposto, IIf(rs!Parcela = 0, "UNICA", rs!Parcela), rs!taxa, 1, Desconto
                End If
                AreaTotal = 0
                AreaConstruida = 0
                If chk And rs!Parcela = 0 Then Call Bdados.AtualizaDados("IC", Bdados.PreparaValor(1), "impresso", "ic='" & rs!Ic & "'")
            'End If
            rs.MoveNext
        Loop
    Else
        Avisa "Nenhum Registro encontrado."
    End If
    Screen.MousePointer = 0
    Bdados.FechaTabela rs
    Bdados.FechaTabela RsAux
    Bdados.FechaTabela RsComp
    Avisa "Impressão de Boletos  Finalizada."
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdLimpar_Click()
    Call Edita.LimpaCampos(Me)
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Edita.AtualizaCombo(Bdados, cboBairro, "Select tba_nome From Tab_Bairro where tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio)
    cboBairro.AddItem " "
    Set cadastro = New VSImposto
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtImovel_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtImovel_LostFocus()
    Dim Cobranca As New VSCobranca
    Dim cLSImposto As New VSImposto
    
    txtImovel = cLSImposto.FormataInscricao(txtImovel, InscImovel)
End Sub

