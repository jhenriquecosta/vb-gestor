VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TMPU902 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grid 
      Height          =   2460
      Left            =   75
      TabIndex        =   7
      Top             =   1590
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   4339
      Caption         =   "Alíquotas Definidas"
      CorTitulo       =   16711680
      CorCaption      =   16777215
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   195
      Left            =   1890
      TabIndex        =   5
      Top             =   -510
      Width           =   375
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   1
      Left            =   6285
      TabIndex        =   4
      Top             =   4110
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   0
      Left            =   5085
      TabIndex        =   3
      Top             =   4110
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   810
      Left            =   60
      TabIndex        =   6
      Top             =   705
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   1429
      Altura          =   1905
      Caption         =   " Definição"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtAno 
         Height          =   300
         Left            =   5940
         TabIndex        =   2
         Top             =   405
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         Caption         =   "Ano"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   4
         MinLen          =   4
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   300
         Left            =   3952
         TabIndex        =   1
         Top             =   405
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         Caption         =   "Valor"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
      End
      Begin VTOcx.cboVISUAL cboUnidade 
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   398
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         Caption         =   "Aliquota"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   1138
      Icone           =   "TMPU902.frx":0000
   End
   Begin VB.Menu mnuApagar 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuApaga 
         Caption         =   "&Apagar parcela"
      End
   End
End
Attribute VB_Name = "TMPU902"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_Click(Index As Integer)
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim rs As VSRecordset
    Select Case Index
        Case 0
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Valores = Bdados.PreparaValor(grid.SelectedItem.SubItems(2), cboUnidade.Coluna(1).Valor, Bdados.Converte(txtValor, TCDuplo), Bdados.Converte(txtAno, TCInteiro))
            Campos = "TPP_COD_TABELA,TPP_REGISTRO,TPP_VALOR,TPP_ANO_ALIQUOTA"
            If Bdados.GravaDados("Tab_Parametro_Iptu", Valores, Campos, "TPP_REGISTRO= " & cboUnidade.Coluna(1).Valor & " and TPP_COD_TABELA =" & grid.SelectedItem.SubItems(2) & " and TPP_ANO_ALIQUOTA = " & CInt(txtAno)) Then
                Util.Informa ("Alíquota " & cboUnidade & " para " & txtAno & " gravada com sucesso.")
                MontaGrade
            End If
            Edita.LimpaCampos Me
            Bdados.FechaTabela rs
            cboUnidade.SetFocus
        Case 1
            Unload Me
    End Select
End Sub

Sub MontaGrade()
    Dim Sql As String
    Sql = "Select TGE_NOME AS ALIQUOTA,TPP_VALOR AS VALOR ,TPP_COD_TABELA,TPP_ANO_ALIQUOTA AS ANO FROM Tab_Parametro_Iptu," & _
        " TAB_GERAL WHERE TPP_COD_TABELA = (SELECT TPP_COD_TABELA FROM Tab_Parametro_Iptu WHERE " & _
        " TPP_VALOR = 'ALIQUOTAS') AND TPP_REGISTRO > 0 AND TPP_REGISTRO = TGE_CODIGO AND TGE_TIPO = " & _
        " (SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME = 'TIPO ALIQUOTA IPTU')"
    grid.Preencher Bdados, Sql, 4500, 1000, 0, 1000
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Dim Sql As String
    Dim rs As VSRecordset
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    MontaGrade
    cboUnidade.PreencherGeral Bdados, "TIPO ALIQUOTA IPTU"
End Sub

Private Sub grid_DblClick()
    cboUnidade.SetarLinha grid.SelectedItem, 0
    txtValor = grid.SelectedItem.SubItems(1)
    txtAno = grid.SelectedItem.SubItems(3)
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuApaga.Caption = "Apagar Parcela " & grid.SelectedItem
        
        Me.PopupMenu mnuApagar
    End If
End Sub

Private Sub mnuApaga_Click()
    Dim Sql As String
    If Confirma("Confirma a exclusão da parcela " & grid.SelectedItem & "?") Then
        If Bdados.DeletaDados("Tab_Parametro_Parcela_Iptu", "TPP_PARCELA = " & grid.SelectedItem) Then
            Informa "Registro excluído."
            LimpaCampos Me
            MontaGrade
            cboUnidade.SetFocus
        End If
    End If
End Sub
