VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TREL402 
   BackColor       =   &H80000016&
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cboVISUAL cboTributo 
      Height          =   315
      Left            =   5700
      TabIndex        =   8
      Top             =   1950
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Caption         =   "Tributo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1125
      Left            =   5670
      TabIndex        =   4
      Top             =   750
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1984
      Altura          =   1905
      Caption         =   " Período"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   16777215
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtDataFinal 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Tag             =   "Data Final"
         Top             =   720
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         Caption         =   "Final"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         CorFundo        =   16777215
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtDataInicial 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Tag             =   "Data Inicial"
         Top             =   330
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         Caption         =   "Inicial"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         CorFundo        =   16777215
         MaxLen          =   10
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   900
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   5940
         TabIndex        =   7
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7140
         TabIndex        =   3
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   1138
      Formulario      =   "TREL402"
      Descricao       =   "Relatórios Gerenciais"
      Icone           =   "TREL402.frx":0000
   End
   Begin VTOcx.grdVISUAL grdRelatorios 
      Height          =   2700
      Left            =   60
      TabIndex        =   2
      Top             =   750
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   4339
      Caption         =   "Relatórios Gerenciais"
      CorTitulo       =   32768
      CorCaption      =   16777215
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
      MarcaUnico      =   -1  'True
   End
End
Attribute VB_Name = "TREL402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    On Error GoTo trata
    Dim Item As ListItem, SelectionFormula As String
        
    If Edita.CriticaCampos(Me) Then
        SelectionFormula = "{Tab_Geracao_Tributo.tgt_tip_cod_imposto} = '" & cboTributo.Coluna(1).Valor & "' and " & _
                            " {Tab_Geracao_Tributo.tgt_data_geracao} in DateTime (" & Year(txtDataInicial) & "," & Month(txtDataInicial) & "," & Day(txtDataInicial) & ") to " & _
                                    " DateTime (" & Year(txtDataFinal) & "," & Month(txtDataFinal) & "," & Day(txtDataFinal) & ")"
        For Each Item In grdRelatorios.ListItems
            If Item.Checked Then
                With Rpt
                    If Not .DefinirArquivo(Bdados, App.Path + "\" & Me.Name & Item & ".rpt") Then Exit Sub
    '                .Formulas "IM", InscMuni
    '                .Formulas "RAZAOSOCIAL", RazaoSocial
    '                .Formulas "NOMEFANTASIA", NomeFantasia
    '                .Formulas "ATIVIDADE", Atividade
    '                .Formulas "RESTRICOES", Restricoes
    '                .Formulas "OBJETIVO", Finalidade
    '                .Formulas "CPF/CNPJ", CPFCNPJ
    '                .Formulas "ENDERECO", Endereco
    '                .Formulas "BAIRRO", Bairro
    '                .Formulas "CidadeEmpresa", Cidade
    '                .Formulas "CEP", Cep
    '                .Formulas "ESTADO", Uf
    '                .Formulas "VALIDADE", txtValidade
    '                .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
    '                .Formulas "CIDADE", Aplicacoes.Municipio
    '                .Formulas "DEPARTAMENTO", Temp.PegaParametro(Bdados, "SETOR")
    '                .Titulo = Imposto.NomeTributo(ttr_ALVARA)
                    .Titulo = Item.SubItems(1)
                    .Formulas "VTTitulo", Item.SubItems(1)
                    .Formulas "VTSubtitulo", cboTributo
                    .Arvore = False
                    '.SELECAO = SelectionFormula
                    .Visualizar
                End With
                Exit Sub
            End If
        Next
    End If
trata:
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub PreencherRelatorios()
    Dim sql As String
    
    sql = "SELECT TGE_CODIGO AS Codigo, TGE_NOME as Relatorio " & _
        " FROM TAB_GERAL " & _
        " WHERE TGE_CODIGO>0 AND " & _
            " TGE_TIPO = (SELECT TGE_TIPO" & _
                            " FROM TAB_GERAL" & _
                            " WHERE TGE_CODIGO=0 AND" & _
                                " TGE_NOME ='RELATORIOS GERENCIAIS TREL402')" & _
        " ORDER BY TGE_NOME"
    grdRelatorios.Preencher Bdados, sql
End Sub

Private Sub PreencherTributos()
    Dim sql As String
    
    sql = "SELECT TIP_SIGLA_IMPOSTO, TIP_COD_IMPOSTO" & _
        " FROM TAB_IMPOSTO" & _
        " ORDER BY TIP_SIGLA_IMPOSTO"
    cboTributo.Preencher Bdados, sql
End Sub

Private Sub Form_Load()
    PreencherRelatorios
    PreencherTributos
End Sub

Private Sub grdRelatorios_Click()
    Dim Item As ListItem
    
    If Not grdRelatorios.SelectedItem Is Nothing Then
        For Each Item In grdRelatorios.ListItems
            Item.Checked = False
        Next
        grdRelatorios.SelectedItem.Checked = True
    End If
End Sub
