VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TIMP105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TIMP105"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   60
      ScaleHeight     =   570
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TIMP105.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.CheckBox chkMostrarSelecionados 
      Caption         =   "Mostrar somente selecionados"
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   4050
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VTOcx.cmdVISUAL cmdCancelar 
      Height          =   345
      Left            =   5700
      TabIndex        =   4
      Top             =   4080
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      Caption         =   "Cancelar"
      Acao            =   7
      CorBorda        =   8421504
   End
   Begin VTOcx.cmdVISUAL cmdOK 
      Height          =   345
      Left            =   4680
      TabIndex        =   3
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      Caption         =   "OK"
      Acao            =   8
      CorBorda        =   8421504
   End
   Begin VTOcx.grdVISUAL grdImposto 
      Height          =   3345
      Left            =   30
      TabIndex        =   1
      Top             =   690
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5900
      Caption         =   "Selecione os tributos relacionados"
      CheckBox        =   -1  'True
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   1138
      Formulario      =   "Impostos Relacionados"
      Descricao       =   ""
      Icone           =   "TIMP105.frx":2123
   End
End
Attribute VB_Name = "TIMP105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim varCodImposto As String
Dim varAnoImposto As String

Private Sub SelecionaImposto(CodImposto As String, Ano As String)
    Dim i As Integer
    For i = 1 To grdImposto.ListItems.Count
        If grdImposto.ListItems(i).Text = CodImposto Then
            If grdImposto.ListItems(i).SubItems(3) = Ano Then
                grdImposto.ListItems(i).Checked = True
                Exit For
            End If
        End If
    Next
End Sub

Public Sub CarregaTributos(CodImposto As String, Ano As String)
    Dim Sql As String
    Dim rs As VSRecordset
    
    varCodImposto = CodImposto
    varAnoImposto = Ano
    
    Sql = "SELECT tpi_tip_cod_imposto as Codigo, tip_sigla_imposto as Sigla,tip_nome_imposto as Imposto, tpi_ano_imposto as Ano, tpi_valor_taxa_fixa as Valor FROM Tab_Parametro_Imposto, Tab_Imposto WHERE tpi_tip_cod_imposto=tip_cod_imposto"
    Sql = Sql & " ORDER BY tip_sigla_imposto, tpi_ano_imposto"
    
    grdImposto.Preencher Bdados, Sql, 1200, 1300, 1700, 900, 900
    
    Sql = "SELECT * FROM TAB_IMPOSTO_RELACIONADO"
    Sql = Sql & " WHERE TIR_TIP_COD_IMPOSTO_PAI = '" & CodImposto & "'"
    Sql = Sql & " AND TIR_ANO_IMPOSTO_PAI = '" & Ano & "'"
    
    Bdados.AbreTabela Sql, rs
    
    Do Until rs.EOF
        SelecionaImposto rs(0), rs(1)
        rs.MoveNext
    Loop
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim condicao As String
    Dim Valores As String
    Dim Campos As String

    If Not Util.Confirma("Deseja salvar os dados") Then
        Exit Sub
    End If
    
    Dim i As Integer
    
    For i = 1 To grdImposto.ListItems.Count
        condicao = "TIR_TIP_COD_IMPOSTO = '" & grdImposto.ListItems(i).Text _
            & "' AND TIR_ANO_IMPOSTO = '" & grdImposto.ListItems(i).SubItems(3) _
            & "' AND TIR_TIP_COD_IMPOSTO_PAI = '" & varCodImposto _
            & "' AND TIR_ANO_IMPOSTO_PAI = '" & varAnoImposto & "'"
            
        If Not grdImposto.ListItems(i).Checked Then
            Bdados.DeletaDados "TAB_IMPOSTO_RELACIONADO", condicao
        Else
            Campos = "TIR_TIP_COD_IMPOSTO,TIR_ANO_IMPOSTO," _
            & "TIR_TIP_COD_IMPOSTO_PAI,TIR_ANO_IMPOSTO_PAI"
            Valores = Bdados.PreparaValor(grdImposto.ListItems(i).Text, grdImposto.ListItems(i).SubItems(3), varCodImposto, varAnoImposto)
            Bdados.GravaDados "TAB_IMPOSTO_RELACIONADO", Valores, Campos, condicao
        End If
    Next
    Unload Me
End Sub
