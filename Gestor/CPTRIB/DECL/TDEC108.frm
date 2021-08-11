VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TDEC108 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TDEC108"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   11
      Top             =   3990
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdFinaliza 
         Height          =   375
         Left            =   5670
         TabIndex        =   7
         Top             =   90
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   661
         Caption         =   "Finalizar Declaracão"
         Acao            =   1
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmLimpar 
         Height          =   375
         Left            =   7770
         TabIndex        =   8
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4770
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   690
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Salvar Declaracão"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9060
         TabIndex        =   9
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1080
      Left            =   30
      TabIndex        =   12
      Top             =   660
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   1905
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   630
         TabIndex        =   1
         Top             =   690
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
         Height          =   315
         Left            =   3090
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   330
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   285
         Left            =   8370
         TabIndex        =   2
         Top             =   690
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         Caption         =   "Período"
         Text            =   ""
         Restricao       =   2
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   3450
         TabIndex        =   13
         Top             =   330
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   315
         Left            =   45
         TabIndex        =   0
         Top             =   330
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         Caption         =   "Insc.Municipal"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
   End
   Begin VTOcx.fraVISUAL fraNormal 
      Height          =   1365
      Index           =   5
      Left            =   60
      TabIndex        =   15
      Top             =   1770
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   2408
      Altura          =   1905
      Caption         =   " Resumo das Notas de Saída"
      CorTexto        =   0
      CorFaixa        =   16711680
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtItemDecl 
         Height          =   315
         Index           =   6
         Left            =   6630
         TabIndex        =   17
         Top             =   990
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         Caption         =   "Imposto devido em notas"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
         CorFundo        =   14737632
      End
      Begin VTOcx.txtVISUAL txtItemDecl 
         Height          =   315
         Index           =   5
         Left            =   1350
         TabIndex        =   16
         Top             =   990
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         Caption         =   "Base de Calculo"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
         CorFundo        =   14737632
      End
      Begin VTOcx.txtVISUAL txtItemDecl 
         Height          =   315
         Index           =   3
         Left            =   990
         TabIndex        =   5
         Top             =   660
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         Caption         =   "Valor total em notas"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
         CorFundo        =   14737632
      End
      Begin VTOcx.txtVISUAL txtItemDecl 
         Height          =   315
         Index           =   4
         Left            =   7080
         TabIndex        =   6
         Top             =   660
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         Caption         =   "Total sujeito a ICMS"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
         CorFundo        =   14737632
      End
      Begin VTOcx.txtVISUAL txtItemDecl 
         Height          =   315
         Index           =   2
         Left            =   7980
         TabIndex        =   4
         Top             =   330
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         Caption         =   "Nota Final"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         CorFundo        =   14737632
      End
      Begin VTOcx.txtVISUAL txtItemDecl 
         Height          =   315
         Index           =   1
         Left            =   1770
         TabIndex        =   3
         Top             =   330
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         Caption         =   "Nota Inicial"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         CorFundo        =   14737632
      End
   End
   Begin VTOcx.fraVISUAL fraNormal 
      Height          =   735
      Index           =   2
      Left            =   60
      TabIndex        =   18
      Top             =   3180
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   1296
      Altura          =   1905
      Caption         =   " Resumo de Recolhimento"
      CorTexto        =   0
      CorFaixa        =   16711680
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtItemDecl 
         Height          =   315
         Index           =   13
         Left            =   7440
         TabIndex        =   20
         Top             =   330
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         Caption         =   "Total a Recolher"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
         CorFundo        =   14737632
      End
      Begin VTOcx.txtVISUAL txtItemDecl 
         Height          =   315
         Index           =   100
         Left            =   2040
         TabIndex        =   19
         Top             =   330
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         Caption         =   "Aliquota"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
         CorFundo        =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   1138
      Icone           =   "TDEC108.frx":0000
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "mnuGeral"
      Visible         =   0   'False
      Begin VB.Menu mnuDeletar 
         Caption         =   "Deletar"
      End
   End
End
Attribute VB_Name = "TDEC108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declaracao As VsTFuncoes.cDeclaracao
Private GerarIM As Boolean
Private AliqISSQN As Double, ISSQNFixo As Double
Private AliqISSST As Double, ISSSTFixo As Double

Private TotalImpostoST As Double
Private TotalBaseST As Double
Private TotalImpostoDevidoSaida As Double
Private TotalImpostoRetidoSaida As Double
Private TotalBaseSaida As Double
Private TotalICMSSujeito As Double
Private DeduzValores As Boolean
Private ContribuinteEndereco As String
Private ContribuinteAtividade As String
Dim Notas() As New NotaFiscal
Dim Modalidade As Integer
Dim ClassGrid As New grdEditavel
Dim String_Taxas As String
Dim Total_Taxas As Double
Dim Atividade As Object

Private Sub FormataRegistro(ByRef Inscricao As Object)
    Select Case Len(Inscricao.Text)
        Case 10
            Inscricao.Text = Imposto.FormataInscricao(Inscricao.Text, InscContrib)
        Case 11
            Inscricao.Text = Edita.FormataTexto(Inscricao, Cpf)
        Case 14
            Inscricao.Text = Edita.FormataTexto(Inscricao, Cgc)
    End Select
End Sub

Private Sub AtualizaApuracao()
    On Error Resume Next
    'RESUMO NOTAS SAIDA
    txtItemDecl(3) = TotalBaseSaida
    txtItemDecl(4) = TotalICMSSujeito
    txtItemDecl(5) = TotalBaseSaida - TotalICMSSujeito
    txtItemDecl(6) = TotalImpostoDevidoSaida
    'NOTAS DE ENTRADA
    txtItemDecl(10) = TotalImpostoST
    txtItemDecl(11) = 0
    txtItemDecl(12) = TotalImpostoST
    'NOTAS EMITIDAS(SAIDAS)
    txtItemDecl(7) = TotalImpostoRetidoSaida
    txtItemDecl(8) = 0
    txtItemDecl(9) = TotalImpostoRetidoSaida
    'TOTAL RECOLHIMENTO
    txtItemDecl(13) = TotalImpostoDevidoSaida + TotalImpostoST - TotalImpostoRetidoSaida
    txtItemDecl(100) = (100 * CDbl(txtItemDecl(13))) / (TotalBaseSaida + TotalBaseST)
End Sub

Private Sub IniciaTotalizadores()
    TotalImpostoST = 0
    TotalBaseSaida = 0
    TotalBaseST = 0
    TotalImpostoDevidoSaida = 0
    TotalImpostoRetidoSaida = 0
    TotalICMSSujeito = 0
End Sub

Private Sub CmdConsultaContribuinte_Click()
    'blnConsultaIM = True
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIM
    'blnConsultaIM = False
End Sub

Private Sub cmdFinaliza_Click()
    Dim NumDec As String
    Dim Controle As Control
    Dim Item As New cItemDeclaracao
    If Confirma("Ao finalizar a declaracão, ela só poderá ser modificada através de uma retificadora. Deseja prosseguir?") Then
        If Trim(txtItemDecl(1)) = "" Or Trim(txtItemDecl(2)) = "" Then
            Avisa "Informe nota inicial e nota final."
            txtItemDecl(1).SetFocus
            Exit Sub
        End If
        
        If CDbl(Nvl(Trim(txtItemDecl(13)), 0)) = 0 Then
            If Not Confirma("Valor Devido igual a zero. Prosseguir?") Then
                txtItemDecl(3).SetFocus
                Exit Sub
            End If
        End If
        
        
        If Not Edita.CriticaCampos(Me) Then Exit Sub
        
        If txtIM = "" Then Exit Sub
    
        If Trim(txtItemDecl(1)) = "" Or Trim(txtItemDecl(2)) = "" Then
            Avisa "Informe nota inicial e nota final."
            txtItemDecl(1).SetFocus
            Exit Sub
        End If
        If cboImposto.ListIndex = -1 Then
            Util.Avisa "Selecione o imposto."
            cboImposto.SetFocus
            Exit Sub
        End If
        
        If txtPeriodo = "" Then
            Util.Avisa "Informe o período."
            txtPeriodo.SetFocus
            Exit Sub
        End If
        If CDbl(Nvl(Trim(txtItemDecl(13)), 0)) = 0 Then
            If Not Confirma("Valor Devido igual a zero. Prosseguir?") Then
                txtItemDecl(3).SetFocus
                Exit Sub
            End If
        End If
        'Atualizando a Aupracao...
        For Each Controle In txtItemDecl
            If Trim(Controle.Text) <> "" And Trim(Controle.Text) <> "0,00" Then
                Set Item = New cItemDeclaracao
                Item.Numero = Controle.Index
                Item.Valor = Nvl(Controle.Text, 0)
                Declaracao.Itens.Adicionar Item
            End If
        Next
        
        Declaracao.Im = txtIM
        Declaracao.Periodo = txtPeriodo
        Declaracao.CodTributo = cboImposto.Coluna(0).Valor
        Declaracao.Data = Format(Date, "dd/mm/yyyy")
        Declaracao.Origem = orgSistema
        Declaracao.Recepcao = Date
        Declaracao.Versao = Nvl(Temp.PegaParametro(Bdados, "VERSAO DEC"), 2)
        Declaracao.CodTributo = cboImposto.Coluna(0).Valor
        Declaracao.BaseGeral = TotalBaseSaida + TotalBaseST
        Declaracao.Tipo = decNormal
        Declaracao.Status = decFinalizada
        
        If Declaracao.Gravar() Then
            Avisa "Declaração gravada com sucesso."
            Declaracao.Salvar_Sem_Finalizar , , , decNormal, CDbl(Nvl(txtItemDecl(13), 0)), etsCriaNova
            txtIM.SetFocus
        End If
    End If
    Set Item = Nothing
    Set Declaracao = Nothing
    Set Declaracao = New VsTFuncoes.cDeclaracao
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If txtIM = "" Then Exit Sub
    
    If Trim(txtItemDecl(1)) = "" Or Trim(txtItemDecl(2)) = "" Then
        Avisa "Informe nota inicial e nota final."
        txtItemDecl(1).SetFocus
        Exit Sub
    End If
    If cboImposto.ListIndex = -1 Then
        Util.Avisa "Selecione o imposto."
        cboImposto.SetFocus
        Exit Sub
    End If
    
    If txtPeriodo = "" Then
        Util.Avisa "Informe o período."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    If CDbl(Nvl(Trim(txtItemDecl(13)), 0)) = 0 Then
        If Not Confirma("Valor Devido igual a zero. Prosseguir?") Then
            txtItemDecl(3).SetFocus
            Exit Sub
        End If
    End If
     
    Declaracao.Im = txtIM
    Declaracao.Periodo = txtPeriodo
    Declaracao.CodTributo = cboImposto.Coluna(0).Valor
    Declaracao.Data = Format(Date, "dd/mm/yyyy")
    Declaracao.Origem = orgSistema
    Declaracao.Recepcao = Date
    Declaracao.Versao = Nvl(Temp.PegaParametro(Bdados, "VERSAO DEC"), 2)
    Declaracao.CodTributo = cboImposto.Coluna(0).Valor
    Declaracao.Status = decAberta
    Declaracao.BaseGeral = TotalBaseSaida + TotalBaseST
    
    If Declaracao.Gravar() Then
        Avisa "Declaração gravada com sucesso."
        'Call Pega_taxas
        Declaracao.Salvar_Sem_Finalizar True
        cmLimpar_Click
        txtIM.SetFocus
    End If
End Sub

Private Sub cmLimpar_Click()
    Edita.LimpaCampos Me
'    grdDec.ListItems.Clear
    IniciaTotalizadores
    txtIM.SetFocus
End Sub

Private Sub Form_Load()
    Dim Sql As String
    IniciaTotalizadores
    
    cabVisual1.Exibir Bdados, Me.Name, App.Path
    Set Atividade = CreateObject("VsTEcon.atividade")
    Set Imposto = New VsTFuncoes.VSImposto
    DeduzValores = True
    Set Declaracao = New cDeclaracao
    Sql = "SELECT TCD_COD_CAMPO as Item ,TCD_CAMPO as Descricao, ' ' as Valor FROM " & _
        "TAB_CONTEUDO_DECLARACAO WHERE TCD_TMD_COD_DECLARACAO = " & 1
    Sql = "Select tip_cod_imposto,tip_nome_imposto from tab_imposto where tip_sigla_imposto like 'ISS%'"
    cboImposto.Preencher Bdados, Sql, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Declaracao = Nothing
    Set Atividade = Nothing
    IniciaTotalizadores
End Sub

Private Sub txtIM_LostFocus()
    
    
    If Trim$(txtIM) <> "" Then
        If Not BuscarContribuinte(txtIM, txtRazao) Then
            Avisa "Contribuinte não encontrado" & vbCrLf & "Verifique se todos os dados estão corretos."
            txtIM = "": txtRazao = ""
            txtIM.SetFocus
        Else
'            On Error Resume Next
            On Error GoTo 0
            AliqISSQN = Atividade.BuscaAliquotaAtividade(Bdados, txtIM, txtPeriodo, ISSQNFixo)
            txtItemDecl(100) = AliqISSQN * 100
            
            Declaracao.tciAtividade = Atividade.Nome
            If AliqISSQN = 0 Then
                Avisa "Contribuinte sem atividade economica definida. Aliquota igual a zero."
             End If
        End If
    End If
End Sub

Private Sub txtItemDecl_Change(Index As Integer)
    Select Case Index
        Case 6, 9, 12
            txtItemDecl(13) = CDbl(Nvl(txtItemDecl(6), 0))
        Case 5
            If txtItemDecl(3).Enabled = True Then
                TotalBaseSaida = CDbl(Nvl(txtItemDecl(5), 0))
            End If
        Case 100
            If IsNumeric(txtItemDecl(100)) Then
                AliqISSQN = CDbl(Nvl(txtItemDecl(100), 0)) / 100
                CalcularImposto txtItemDecl(3), txtItemDecl(4), txtItemDecl(5), txtItemDecl(6), txtItemDecl(100)
            End If
            
    End Select
End Sub

Private Sub txtItemDecl_LostFocus(Index As Integer)
    Select Case Index
        Case 3, 4
            txtItemDecl(5) = CDbl(Nvl(txtItemDecl(3), 0)) - CDbl(Nvl(txtItemDecl(4), 0))
            CalcularImposto txtItemDecl(3), txtItemDecl(4), txtItemDecl(5), txtItemDecl(6), txtItemDecl(100)
        Case 7, 8
            txtItemDecl(8) = Nvl(txtItemDecl(8), 0)
            If CDbl(Nvl(txtItemDecl(9), 0)) - CDbl(Nvl(txtItemDecl(8), 0)) >= 0 Then
                txtItemDecl(7) = CDbl(Nvl(txtItemDecl(9), 0)) - CDbl(Nvl(txtItemDecl(8), 0))
            Else
                Avisa "Dados inválidos. Valor negativo encontrado." '& Nvl(txtItemDecl(9), 0) & " - " & Nvl(txtItemDecl(8), 0) & " = " & CDbl(CDbl(Nvl(txtItemDecl(9), 0)) - CDbl(Nvl(txtItemDecl(8), 0)))
                txtItemDecl(8).SetFocus
             End If
        Case 10, 11
        txtItemDecl(11) = Nvl(txtItemDecl(11), 0)
        If CDbl(Nvl(txtItemDecl(12), 0)) - CDbl(Nvl(txtItemDecl(11), 0)) >= 0 Then
            txtItemDecl(10) = CDbl(Nvl(txtItemDecl(12), 0)) - CDbl(Nvl(txtItemDecl(11), 0))
        Else
            Avisa "Dados inválidos. Valor negativo encontrado." '& Nvl(txtItemDecl(12), 0) & " - " & Nvl(txtItemDecl(11), 0) & " = " & CDbl(CDbl(Nvl(txtItemDecl(12), 0)) - CDbl(Nvl(txtItemDecl(11), 0)))
            txtItemDecl(11).SetFocus
        End If
    End Select
End Sub

Private Sub CalcularImposto(ByRef Total As Object, ByRef ICMS As Object, ByRef Tributavel As Object, ByRef Imposto As Object, ByRef Aliquota As Object)
    Total = Nvl(Trim$(Total.Text), 0)
    ICMS = Nvl(Trim$(ICMS.Text), 0)
   'Aliquota = AliqISSQN * 100
   'Tributavel = Total - ICMS
   'If AliqISSQN > 0 Then
   '    Imposto = Tributavel * AliqISSQN
   'Else
   '    Imposto = ISSQNFixo
   'End If
    Tributavel = Total - ICMS
    If AliqISSQN >= 0 Then
        Imposto = Tributavel * AliqISSQN
    Else
        Imposto = ISSQNFixo
    End If
End Sub

Private Sub txtPeriodo_Change()
    If Not IsNumeric(Trim(txtPeriodo)) Then Exit Sub
    If Len(Trim(txtPeriodo)) <> 6 Then Exit Sub
'    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
'    If Trim(txtIM) <> "" And cboTipo.ListIndex <> -1 Then PreencheDeclaracao
'    On Error Resume Next
    
    AliqISSQN = Atividade.BuscaAliquotaAtividade(Bdados, txtIM, txtPeriodo, ISSQNFixo)
    txtItemDecl(100) = AliqISSQN * 100
    If CInt(Left(Trim(txtPeriodo), 2)) > 12 Or CInt(Left(Trim(txtPeriodo), 2)) < 1 Then
        Avisa "Periodo inválido."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtPeriodo_LostFocus()
    If Len(Trim(txtPeriodo)) <> 6 Then Exit Sub
    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
End Sub

Private Function BuscarContribuinte(ByRef Inscricao As Object, Optional ByRef Nome As Object, Optional ByRef Endereco As Object, _
                    Optional ByRef Bairro As Object, Optional ByRef Cep As Object, Optional ByRef Cidade As Object, Optional ByRef Uf As Object) As Boolean
    Dim Im As Boolean
    Im = False
    If Trim(Inscricao) = "" Then Exit Function
    Inscricao.Text = Edita.TiraTudo(Inscricao.Text)
    If Len(Inscricao.Text) = 10 Then Im = True
    FormataRegistro Inscricao
    If Trim(Inscricao) = "" Then Exit Function
    
    Dim Sql As String, rs As VSRecordset
    Sql = "SELECT tci_im, TCI_CGC_CPF,tci_nome, tci_logradouro, tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade, tci_UF " & _
            ",TAE_NOME FROM TAB_CONTRIBUINTE left join TAB_ATIVIDADE_ECONOMICA on TCI_TAE_CAE = TAE_CAE where 1=1"
    If Im Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        Sql = Sql & " AND TCI_IM='" & Inscricao & "'"
    Else
        Sql = Sql & " AND TCI_CGC_CPF='" & Inscricao & "'"
    End If
    
    If Bdados.AbreTabela(Sql, rs) Then
        If Im Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            Inscricao = "" & rs!tci_im
        Else
            Inscricao = "" & rs!TCI_CGC_CPF
        End If
        If Not Nome Is Nothing Then Nome = "" & rs!tci_nome
        If Not Endereco Is Nothing Then Endereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & ", " & rs!tci_numero & " " & rs!tci_complemento
        If Not Bairro Is Nothing Then Bairro = "" & rs!tci_bairro
        If Not Cep Is Nothing Then Cep = "" & rs!tci_cep
        If Not Cidade Is Nothing Then Cidade = "" & rs!tci_cidade
        If Not Uf Is Nothing Then Uf = "" & rs!tci_UF
        With Declaracao
            .tciNome = "" & rs!tci_nome
            .tciEndereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & ", " & rs!tci_numero & " " & rs!tci_complemento
            .tciBairro = "" & rs!tci_bairro
            .tciCEP = "" & rs!tci_cep
            .tciCidade = "" & rs!tci_cidade
            .tciUF = "" & rs!tci_UF
            .tciEndereco = .tciEndereco & " " & .tciBairro & " " & .tciCidade & "-" & rs!tci_UF
            .tciAtividade = "" & rs!TAE_NOME
        End With
        BuscarContribuinte = True
    End If
    Bdados.FechaTabela rs
End Function

