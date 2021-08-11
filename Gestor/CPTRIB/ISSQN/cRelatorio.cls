VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'============================================'
'      VsTIss.cRetencao: Usado em TSUB101                                '
'                    Feito por Djeison  -  08/05/2003                               '
'============================================'
Option Explicit
Private m_Nota As cNota
Private m_OBS As String

Public Property Get OBS() As String
    OBS = m_OBS
End Property

Public Property Let OBS(ByVal Value As String)
    m_OBS = Value
End Property

Public Property Get Nota() As cNota
    Set Nota = m_Nota
End Property

Public Property Set Nota(ByVal Value As cNota)
    Set m_Nota = Value
End Property

Private Sub Class_Initialize()
    Set m_Nota = New cNota
End Sub

Private Sub Class_Terminate()
    Set m_Nota = Nothing
End Sub

Public Sub GeraDam(ImToma As String, CGCToma As String, RazaoToma As String, EnderecoToma As String, PicBarra As Object)
    On Error GoTo trata
    Dim Cobranca As New VSCobranca
    Dim a As Byte
    Dim CgcPref As String
    Dim i As Integer
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TDAM_SUBST_Barra.rpt") Then Exit Sub
        .Formulas "InscMunicipal", ImToma  'OK
        CgcPref = Temp.PegaParametro(Bdados, "CGC CLIENTE") 'OK
        .Formulas "CgcPrefeitura", "CNPJ " & Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2) 'OK
        
        .Formulas "NOME_SUBST", CGCToma & " - " & RazaoToma    'OK
        .Formulas "endereco_SUBST", EnderecoToma
        
        .Formulas "documento ", CStr(Nota.Cod_Pagamento)  'OK
        .Formulas "datavencimento", Nota.Data_Venc 'OK
        
        .Formulas "nome_PRESTA", Nota.Nome_Empresa 'OK
        .Formulas "endereco_PRESTA", Nota.Endereco.Endereco 'OK
        .Formulas "ATIVIDADE_PRESTA", Nota.Atividade  'OK
        
        .Formulas "exercicio", IIf(Len(Nota.Periodo_Ref) = 4, Nota.Periodo_Ref, Left(Nota.Periodo_Ref, 2) & "/" & Right(Nota.Periodo_Ref, 4)) 'OK
        .Formulas "ValorTributo", Format(Nota.ISS_Devido, Const_Monetario)  'OK

        .Formulas "ValorMulta", Format(Nota.Multa, Const_Monetario) 'OK
        .Formulas "ValorJuros", Format(Nota.Juros, Const_Monetario)   'OK
        .Formulas "TaxaExpediente", Format(0, Const_Monetario) 'OK
        .Formulas "ValorTotal", Format(CDbl(Nota.ISS_Devido) + CDbl(Nota.Multa) + CDbl(Nota.Juros), Const_Monetario) 'OK
        .Formulas "CodigoTributo", Nota.Cod_Imposto  'OK
        
        .Formulas "IMPOSTO", Nota.Nome_Imposto 'OK
        
        .Formulas "NUM_NOTAS", "Nº da Nota Fiscal: " & Nota.Nota_fiscal
        .Formulas "BASECALCULO", Format(Nota.Base_Calculo, Const_Monetario)  'OK

        Dim PosTraco As Byte
        .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE")) 'OK
        .Formulas "EMISSAO", Imposto.BuscaDataGeracaoDam(Nota.Cod_Pagamento)  'OK
'        .Formulas "LinhaDigitavel", Cobranca.GeraCodBarra(Nota.Cod_Pagamento, Nota.Cod_Imposto, CDbl(Nota.ISS_Devido), PicBarra)
        .Formulas "OBSERVACAO", m_OBS    'OK

        .CopiasDetalhes = 3
        .Titulo = "Documento de Arrecadação Municipal - DAM"
        .Arvore = False
        .Visualizar
    End With
    Set Rpt = Nothing
    Screen.MousePointer = 0
    Informa "Impressão de comprovante(s) finalizada."
    Exit Sub
trata:
    If Err.Number = 20515 Or Err.Number = 3265 Then
        Rpt.Formulas "OBSERVACAO", ""
        Resume
    End If
    Avisa "O DAM não será gerado."
    Avisa Err.Description
    Exit Sub
End Sub

