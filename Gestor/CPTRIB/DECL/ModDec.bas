Attribute VB_Name = "ModDec"
Option Explicit
Public Declaracao As New VsTFuncoes.cDeclaracao
Public String_Taxas As String
Public Total_Taxas  As Double
Public Atividade As Object
Type gerador
    Inscricao       As String
    periodo_inicial As String
    tipo_decla      As Integer
    data_geracao    As String
    condicao        As Byte
End Type
Public gera() As gerador
Public Function BuscarContribuinte(ByRef Inscricao As Object, ByRef Nome As Object, Optional ByRef Endereco As Object, _
                    Optional ByRef Bairro As Object, Optional ByRef Cep As Object, Optional ByRef Cidade As Object, Optional ByRef Uf As Object) As Boolean
    Dim Im As Boolean
    
    Im = False
    If Trim(Inscricao) = "" Then Exit Function
    Inscricao.Text = Edita.TiraTudo(Inscricao.Text)
    If Len(Inscricao.Text) = 10 Then Im = True
    FormataRegistro Inscricao
    If Trim(Inscricao) = "" Then Exit Function
    
    Dim Sql As String, Rs As VSRecordset
    Sql = "SELECT tci_im, TCI_CGC_CPF,tci_nome, tci_logradouro, tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade, tci_UF " & _
            " FROM TAB_CONTRIBUINTE"
    If Im Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        Sql = Sql & " WHERE TCI_IM='" & Inscricao & "'"
    Else
        Sql = Sql & " WHERE TCI_CGC_CPF='" & Inscricao & "'"
    End If
    
    If Bdados.AbreTabela(Sql, Rs) Then
        If Im Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            Inscricao = "" & Rs!tci_im
        Else
            Inscricao = "" & Rs!TCI_CGC_CPF
        End If
        Nome = "" & Rs!tci_nome
        If Not Endereco Is Nothing Then Endereco = "" & Rs!tci_logradouro & " " & Rs!tci_nome_logradouro & ", " & Rs!tci_numero & " " & Rs!tci_complemento
        If Not Bairro Is Nothing Then Bairro = "" & Rs!tci_bairro
        If Not Cep Is Nothing Then Cep = "" & Rs!tci_cep
        If Not Cidade Is Nothing Then Cidade = "" & Rs!tci_cidade
        If Not Uf Is Nothing Then Uf = "" & Rs!tci_UF
        
        With Declaracao
            .tciNome = "" & Rs!tci_nome
            .tciEndereco = "" & Rs!tci_logradouro & " " & Rs!tci_nome_logradouro & ", " & Rs!tci_numero & " " & Rs!tci_complemento
            .tciBairro = "" & Rs!tci_bairro
            .tciCEP = "" & Rs!tci_cep
            .tciCidade = "" & Rs!tci_cidade
            .tciUF = "" & Rs!tci_UF
        End With
        BuscarContribuinte = True
    End If
    Bdados.FechaTabela Rs
End Function


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

Public Function BuscaModalidadeDeclaracao(Contribuinte As String, Optional DescricaoModalidade As Object) As Integer
    Dim Sql As String
    Dim Rs As VSRecordset
    If Trim(Contribuinte) = "" Then Exit Function
    Sql = "Select TAE_TMD_COD_DECLARACAO,TMD_DESCRICAO FROM TAB_ATIVIDADE_ECONOMICA,TAB_MODALIDADE_DECLARACAO WHERE " & _
            " TAE_TMD_COD_DECLARACAO = TMD_COD_DECLARACAO AND TAE_CAE =( Select TCI_TAE_CAE FROM " & _
            " tab_contribuinte where tci_cgc_cpf ='" & Contribuinte & "' or tci_im ='" & Contribuinte & "')"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaModalidadeDeclaracao = Nvl("" & Rs(0), 0)
        DescricaoModalidade.Caption = "" & Rs!TMD_DESCRICAO
    End If
    Bdados.FechaTabela Rs
End Function
'Public Function Pega_taxas()
 '   TDEC107.Show 1
'End Function

