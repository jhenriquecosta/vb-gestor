VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TMPU903 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grid 
      Height          =   2955
      Left            =   75
      TabIndex        =   9
      Top             =   1560
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   4339
      CorBorda        =   0
      CorTitulo       =   16711680
      CorCaption      =   16777215
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   8
      Top             =   4560
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   5385
         TabIndex        =   4
         Top             =   105
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
         Left            =   6585
         TabIndex        =   5
         Top             =   105
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
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   795
      Left            =   75
      TabIndex        =   7
      Top             =   720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1402
      Altura          =   1905
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtOrdem 
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Top             =   375
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         Caption         =   "Ordem"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   300
         Left            =   6300
         TabIndex        =   3
         Top             =   375
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         Caption         =   "Fator"
         Text            =   ""
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtLimSup 
         Height          =   300
         Left            =   3900
         TabIndex        =   2
         Top             =   375
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   529
         Caption         =   "Faixa Superior"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtLimInf 
         Height          =   300
         Left            =   1500
         TabIndex        =   1
         Top             =   375
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   529
         Caption         =   "Faixa Inferior"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
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
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   1138
      Icone           =   "TMPU903.frx":0000
   End
   Begin VB.Menu mnuApagar 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuApaga 
         Caption         =   "&Apagar parcela"
      End
   End
End
Attribute VB_Name = "TMPU903"
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
            Valores = Bdados.PreparaValor(txtOrdem, txtLimInf, txtLimSup, Bdados.Converte(txtValor, TCDuplo))
            Campos = "TGL_ORDEM,TGL_LIMITE_INFERIOR,TGL_LIMITE_SUPERIOR,TGL_VALOR"
            If Bdados.GravaDados("TAB_GLEBA_LIMITE", Valores, Campos, "TGL_ORDEM= " & CDbl(txtOrdem)) Then
                Call Util.Informa("Transação Completada.")
                Sql = "Select TGL_ORDEM as ORDEM ,TGL_LIMITE_INFERIOR AS FAIXA,TGL_LIMITE_SUPERIOR AS VALOR," & _
                    " TGL_VALOR AS VALOR FROM TAB_GLEBA_LIMITE"
                grid.Preencher Bdados, Sql, 1000
                Edita.LimpaCampos Me
                txtOrdem.SetFocus
            End If
        Case 1
            Unload Me
    End Select
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Dim Sql As String
    Dim rs As VSRecordset
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Sql = "Select TGL_ORDEM as ORDEM ,TGL_LIMITE_INFERIOR AS FAIXA,TGL_LIMITE_SUPERIOR AS VALOR," & _
    " TGL_VALOR AS VALOR FROM TAB_GLEBA_LIMITE"
    grid.Preencher Bdados, Sql

End Sub

Private Sub grid_DblClick()
    txtOrdem = grid.SelectedItem
    txtLimInf = grid.SelectedItem.SubItems(1)
    txtLimSup = grid.SelectedItem.SubItems(2)
    txtValor = grid.SelectedItem.SubItems(3)
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 Then
        mnuApaga.Caption = "Apagar Faixa " & grid.SelectedItem
        
        Me.PopupMenu mnuApagar
    End If
End Sub

Private Sub mnuApaga_Click()
    Dim Sql As String
    If grid.ListItems.Count >= 1 Then
        If Confirma("Confirma a exclusão da faixa " & grid.SelectedItem & "?") Then
            If Bdados.DeletaDados("TAB_GLEBA_LIMITE", "TGL_ORDEM= " & grid.SelectedItem) Then
                Informa "Registro excluído."
                LimpaCampos Me
                Sql = "Select TPP_PARCELA AS PARCELA,TPP_VENCIMENTO AS VENC_PARCELA,TPP_DESCONTO  AS DESCONTO FROM Tab_Parametro_Parcela_Iptu"
                Sql = "Select TGL_ORDEM as ORDEM ,TGL_LIMITE_INFERIOR AS FAIXA,TGL_LIMITE_SUPERIOR AS VALOR," & _
                " TGL_VALOR AS VALOR FROM TAB_GLEBA_LIMITE"
                grid.Preencher Bdados, Sql, 1000
            End If
        End If
    End If
End Sub

