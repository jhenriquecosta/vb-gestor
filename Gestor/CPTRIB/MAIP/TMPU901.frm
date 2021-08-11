VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TMPU901 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   5010
         TabIndex        =   4
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   6210
         TabIndex        =   5
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.grdVISUAL grid 
      Height          =   2460
      Left            =   75
      TabIndex        =   8
      Top             =   1575
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   4339
      Caption         =   "Lançamentos"
      CorTitulo       =   16711680
      CorCaption      =   16777215
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   765
      Left            =   75
      TabIndex        =   9
      Top             =   735
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1349
      Altura          =   1905
      Caption         =   " Parâmetros"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtDesconto 
         Height          =   300
         Left            =   5370
         TabIndex        =   3
         Top             =   375
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   529
         Caption         =   "Desconto(%)"
         Text            =   ""
         Restricao       =   3
      End
      Begin VTOcx.txtVISUAL txtAno 
         Height          =   300
         Left            =   105
         TabIndex        =   0
         Top             =   345
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Caption         =   "Ano"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   4
         MinLen          =   4
      End
      Begin VTOcx.txtVISUAL txtVence 
         Height          =   300
         Left            =   2835
         TabIndex        =   2
         Top             =   360
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   529
         Caption         =   "Vencimento"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
      End
      Begin VTOcx.txtVISUAL txtParcela 
         Height          =   300
         Left            =   1410
         TabIndex        =   1
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Caption         =   "Parcela"
         Text            =   ""
         Restricao       =   2
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   195
      Left            =   1890
      TabIndex        =   6
      Top             =   -540
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   1138
      Icone           =   "TMPU901.frx":0000
   End
   Begin VB.Menu mnuApagar 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuApaga 
         Caption         =   "&Apagar parcela"
      End
   End
End
Attribute VB_Name = "TMPU901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_Click(Index As Integer)
    Dim Valores As String
    Dim Campos As String
    Dim rs As VSRecordset
    Select Case Index
        Case 0
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Valores = Bdados.PreparaValor(txtParcela, Bdados.Converte(txtVence, TCDataHora), txtDesconto, txtAno)
            Campos = "TPP_PARCELA,TPP_VENCIMENTO,TPP_DESCONTO,TPP_ANO"
            Call Bdados.GravaDados("Tab_Parametro_Parcela_Iptu", Valores, Campos, "TPP_PARCELA = " & txtParcela & " AND TPP_ANO = " & CInt(txtAno))
            Call Util.Informa("Transação Completada.")
            MostraParametros
            Edita.LimpaCampos Me
            Bdados.FechaTabela rs
            txtAno.SetFocus
        Case 1
            Unload Me
    End Select
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Dim rs As VSRecordset
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    MostraParametros
End Sub

Private Sub grid_DblClick()
    txtAno = grid.SelectedItem
    txtParcela = grid.SelectedItem.SubItems(1)
    txtVence = grid.SelectedItem.SubItems(2)
    txtDesconto = grid.SelectedItem.SubItems(3)
    txtAno.SetFocus
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuApaga.Caption = "Apagar " & grid.SelectedItem & " : " & grid.SelectedItem.SubItems(1)
        Me.PopupMenu mnuApagar
    End If
End Sub

Private Sub mnuApaga_Click()
    If Confirma("Confirma a exclusão da parcela " & grid.SelectedItem.SubItems(1) & " do ano " & grid.SelectedItem & "?") Then
        If Bdados.DeletaDados("Tab_Parametro_Parcela_Iptu", "TPP_ANO = " & grid.SelectedItem & " and TPP_PARCELA = " & CInt(grid.SelectedItem.SubItems(1))) Then
            'Informa "Registro excluído."
            LimpaCampos Me
            MostraParametros
            txtParcela.SetFocus
        End If
    End If
End Sub

Private Sub MostraParametros()
Dim Sql As String
    Sql = "Select TPP_ANO as Ano,TPP_PARCELA AS Parcela,TPP_VENCIMENTO AS Vencimento,TPP_DESCONTO AS Desconto FROM Tab_Parametro_Parcela_Iptu" & _
            " order by TPP_ANO, TPP_PARCELA"
    grid.Preencher Bdados, Sql, 1000, 1000, 1500
End Sub
