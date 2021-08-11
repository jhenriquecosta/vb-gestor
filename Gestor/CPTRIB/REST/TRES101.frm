VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TRES101 
   Caption         =   "TRES101"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "TRES101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GeraCod As New ContaCorrente

Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub





Private Sub txtCPF_LostFocus()
      If Trim(txtCPF) = "" Then Exit Sub
    
    If txtCPF = "11111111111" Or txtCPF = "111.111.111-11" Or txtCPF = "22222222222" Or txtCPF = "222.222.222-22" Or txtCPF = "33333333333" Or txtCPF = "333.333.333-33" Or txtCPF = "44444444444" Or txtCPF = "444.444.444-44" Or txtCPF = "55555555555" Or txtCPF = "555.555.555-55" Or txtCPF = "66666666666" Or txtCPF = "666.666.666-66" Or txtCPF = "77777777777" Or txtCPF = "777.777.777-77" Or txtCPF = "88888888888" Or txtCPF = "888.888.888-88" Or txtCPF = "99999999999" Or txtCPF = "999.999.999-99" Or txtCPF = "00000000000" Or txtCPF = "000.000.000-00" Or txtCPF = "111.111.111-11" Or txtCPF = "11111111111" Then
        Util.Avisa "Valor do CPF inválido."
        txtCPF.SetFocus
    End If
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub

Private Sub cmdSalvar_Click()
    Dim camposRe As String
    Dim camposPr As String
    Dim ValoresRe As String
    Dim ValoresPr As String
    Dim Descricao As String
    Dim Codigo As Double
    Dim tipoProcesso As Integer
    Dim status As Integer
    
    Descricao = "INSTAURAÇÃO DE REGIME ESPECIAL"
    
    tipoProcesso = 3
    status = 1
    Codigo = GeraCod.GeraCodPagamento(40)
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If txtNota = "" And txtLivro = "" And txtDeclaracao = "" And txtDocumento = "" Then
        Avisa "Identifique um dos documentos"
        Exit Sub
    End If
    ' status_processo = 1 - aberto TipoProcesso Regime Especial = 3
    ' status regime = 3
    camposPr = "TPR_NUMERO_PROCESSO,TPR_INSCRICAO,TPR_DESCRICAO_PEDIDO,TPR_PEDIDO_REPR_PREPOSTO,TPR_PEDIDO_REPR_PREPOSTO_CPF,TPR_PEDIDO_DATA,TPR_TIPO_PROCESSO,TPR_STATUS"
    ValoresPr = Bdados.PreparaValor(Codigo, txtIm, Descricao, txtResp, txtCPF, Date, tipoProcesso, status)
    If Bdados.InsereDados("TAB_PROCESSO", ValoresPr, camposPr) Then
        camposRe = "TRE_TCI_IM,TRE_TPR_NUMERO_PROCESSO,TRE_Descricao_Declaracao,TRE_Descricao_Documento_Fiscal,TRE_Descricao_Nota_Fiscal,TRE_Livros_Fiscais_Modelos,TRE_STATUS_CONCLUSAO"
        ValoresRe = Bdados.PreparaValor(txtIm, Codigo, txtDeclaracao, txtDocumento, txtNota, txtLivro, 3)
        If Bdados.InsereDados("TAB_REGIME_ESPECIAL", ValoresRe, camposRe) Then
            Avisa "Dados gravados com sucesso"
            cmdLimpar_Click
        End If
    End If
End Sub



Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
    
End Sub





