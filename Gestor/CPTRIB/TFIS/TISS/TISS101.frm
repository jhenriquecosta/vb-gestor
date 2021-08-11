VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form TISS101 
   Caption         =   "TISS101"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VTOcx.cmdVISUAL CmdBuscar 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   780
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "Buscar"
      Acao            =   5
   End
   Begin VTOcx.txtVISUAL txtItem 
      Height          =   300
      Left            =   300
      TabIndex        =   4
      Top             =   780
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   529
      Caption         =   "Item"
      Text            =   ""
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   6690
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   714
      Begin VTOcx.cmdVISUAL CmaImprimir 
         Height          =   315
         Left            =   8715
         TabIndex        =   6
         Top             =   75
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "Imprimir"
         Acao            =   4
      End
      Begin VTOcx.cmdVISUAL AmdSair 
         Height          =   315
         Left            =   9825
         TabIndex        =   2
         Top             =   75
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         Caption         =   "Sair"
         Acao            =   7
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   1138
      Icone           =   "TISS101.frx":0000
   End
   Begin MSComctlLib.TreeView TreDados 
      Height          =   5310
      Left            =   15
      TabIndex        =   3
      Top             =   1350
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   9366
      _Version        =   393217
      Indentation     =   538
      LabelEdit       =   1
      Style           =   4
      HotTracking     =   -1  'True
      ImageList       =   "imlMenu"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "TISS101.frx":282A
   End
   Begin VB.Label LblDescricao 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2820
      TabIndex        =   7
      Top             =   780
      Width           =   7965
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "TISS101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MontaTre()
    On Error GoTo trata
    Dim Sql As String
    Dim rs As VSRecordset
    Dim i As Integer
    
    
    Sql = "Select * from TAB_ATIVIDADE_COMPARATIVO where 1 = 1"
    If txtItem <> "" Then
        Sql = Sql & "  and TAC_ITEM_406_CODIGO = '" & txtItem & "'"
    End If
    Sql = Sql & "  order by TAC_ITEM_406_CODIGO"
    TreDados.Nodes.Clear
    If Bdados.AbreTabela(Sql, rs) Then
        rs.MoveFirst
        Do Until rs.EOF
            i = i + 1
            TreDados.Nodes.Add , , rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", rs.Fields("TAC_ITEM_406_CODIGO") & " - " & rs.Fields("TAC_ITEM_406_DESCRICAO")
            If Not IsNull(rs.Fields("TAC_ITEM_A_116_CODIGO")) Then
                If Trim(rs.Fields("TAC_ITEM_A_116_CODIGO")) <> "" Then TreDados.Nodes.Add rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, rs.Fields("TAC_ITEM_A_116_CODIGO") & i & "A", rs.Fields("TAC_ITEM_A_116_CODIGO") & " - " & rs.Fields("TAC_ITEM_A_116_DESCRICAO")
                If Trim(rs.Fields("TAC_ITEM_B_116_CODIGO")) <> "" Then TreDados.Nodes.Add rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, rs.Fields("TAC_ITEM_B_116_CODIGO") & i & "B", rs.Fields("TAC_ITEM_B_116_CODIGO") & " - " & rs.Fields("TAC_ITEM_B_116_DESCRICAO")
                If Trim(rs.Fields("TAC_ITEM_C_116_CODIGO")) <> "" Then TreDados.Nodes.Add rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, rs.Fields("TAC_ITEM_C_116_CODIGO") & i & "C", rs.Fields("TAC_ITEM_C_116_CODIGO") & " - " & rs.Fields("TAC_ITEM_C_116_DESCRICAO")
                If Trim(rs.Fields("TAC_ITEM_D_116_CODIGO")) <> "" Then TreDados.Nodes.Add rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, rs.Fields("TAC_ITEM_D_116_CODIGO") & i & "D", rs.Fields("TAC_ITEM_D_116_CODIGO") & " - " & rs.Fields("TAC_ITEM_D_116_DESCRICAO")
                If Not IsNull(rs.Fields("TAC_ITEM_D_116_CODIGO")) Then
                    If Trim(rs.Fields("TAC_ITEM_D_116_CODIGO")) <> "" Then TreDados.Nodes.Add rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, rs.Fields("TAC_ITEM_D_116_CODIGO") & i & "ADICIONAL", "" & rs.Fields("TAC_ITEM_ADICIONAL_116")
                Else
                    If Trim(rs.Fields("TAC_ITEM_D_116_CODIGO")) <> "" Then TreDados.Nodes.Add rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, rs.Fields("TAC_ITEM_D_116_CODIGO") & i & "ADICIONAL", "" & rs.Fields("TAC_ITEM_ADICIONAL_116")
                End If
                If Trim(rs.Fields("TAC_ITEM_E_116_CODIGO")) <> "" Then TreDados.Nodes.Add rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, rs.Fields("TAC_ITEM_E_116_CODIGO") & i & "E", rs.Fields("TAC_ITEM_E_116_CODIGO") & " - " & rs.Fields("TAC_ITEM_E_116_DESCRICAO")
                If Trim(rs.Fields("TAC_ITEM_F_116_CODIGO")) <> "" Then TreDados.Nodes.Add rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, rs.Fields("TAC_ITEM_F_116_CODIGO") & i & "F", rs.Fields("TAC_ITEM_F_116_CODIGO") & " - " & rs.Fields("TAC_ITEM_F_116_DESCRICAO")
                If Trim(rs.Fields("TAC_ITEM_G_116_CODIGO")) <> "" Then TreDados.Nodes.Add rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, rs.Fields("TAC_ITEM_G_116_CODIGO") & i & "G", rs.Fields("TAC_ITEM_G_116_CODIGO") & " - " & rs.Fields("TAC_ITEM_G_116_DESCRICAO")
                If Trim(rs.Fields("TAC_ITEM_H_116_CODIGO")) <> "" Then TreDados.Nodes.Add rs.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, rs.Fields("TAC_ITEM_H_116_CODIGO") & i & "H", rs.Fields("TAC_ITEM_H_116_CODIGO") & " - " & rs.Fields("TAC_ITEM_H_116_DESCRICAO")
            Else
                If Not IsNull(rs.Fields("TAC_ITEM_A_116_CODIGO")) Then
                    If Trim(rs.Fields("TAC_ITEM_A_116_CODIGO")) <> "" Then TreDados.Nodes.Add , , rs.Fields("TAC_ITEM_A_116_CODIGO") & "FILHO", rs.Fields("TAC_ITEM_A_116_CODIGO") & " - " & rs.Fields("TAC_ITEM_A_116_DESCRICAO")
                End If
            End If
Proximo:
            rs.MoveNext
        Loop
    End If
'    If Not IsNull(RS.Fields("TAC_ITEM_A_116_CODIGO")) Then
'    TreDados.Nodes.Add RS.Fields("TAC_ITEM_406_CODIGO") & "PAI", tvwChild, RS.Fields("TAC_ITEM_A_116_CODIGO") & i & "A", RS.Fields("TAC_ITEM_A_116_CODIGO") & " - " & RS.Fields("TAC_ITEM_A_116_DESCRICAO")
'    TreDados.Nodes.Add RS.Fields("TAC_ITEM_A_116_CODIGO") & i & "A", tvwChild, RS.Fields("TAC_ITEM_B_116_CODIGO") & i & "B", RS.Fields("TAC_ITEM_B_116_CODIGO") & " - " & RS.Fields("TAC_ITEM_B_116_DESCRICAO")
'    TreDados.Nodes.Add RS.Fields("TAC_ITEM_B_116_CODIGO") & i & "B", tvwChild, RS.Fields("TAC_ITEM_C_116_CODIGO") & i & "C", RS.Fields("TAC_ITEM_C_116_CODIGO") & " - " & RS.Fields("TAC_ITEM_C_116_DESCRICAO")
'    TreDados.Nodes.Add RS.Fields("TAC_ITEM_C_116_CODIGO") & i & "C", tvwChild, RS.Fields("TAC_ITEM_D_116_CODIGO") & i & "D", RS.Fields("TAC_ITEM_D_116_CODIGO") & " - " & RS.Fields("TAC_ITEM_D_116_DESCRICAO")
'    If Not IsNull(RS.Fields("TAC_ITEM_D_116_CODIGO")) Then
'    TreDados.Nodes.Add RS.Fields("TAC_ITEM_D_116_CODIGO") & i & "D", tvwChild, RS.Fields("TAC_ITEM_D_116_CODIGO") & i & "ADICIONAL", "" & RS.Fields("TAC_ITEM_ADICIONAL_116")
'    Else
'    TreDados.Nodes.Add RS.Fields("TAC_ITEM_D_116_CODIGO") & i & "D", tvwChild, RS.Fields("TAC_ITEM_D_116_CODIGO") & i & "ADICIONAL", "" & RS.Fields("TAC_ITEM_ADICIONAL_116")
'    End If
'    TreDados.Nodes.Add RS.Fields("TAC_ITEM_D_116_CODIGO") & i & "ADICIONAL", tvwChild, RS.Fields("TAC_ITEM_E_116_CODIGO") & i & "E", RS.Fields("TAC_ITEM_E_116_CODIGO") & " - " & RS.Fields("TAC_ITEM_E_116_DESCRICAO")
'    TreDados.Nodes.Add RS.Fields("TAC_ITEM_E_116_CODIGO") & i & "E", tvwChild, RS.Fields("TAC_ITEM_F_116_CODIGO") & i & "F", RS.Fields("TAC_ITEM_F_116_CODIGO") & " - " & RS.Fields("TAC_ITEM_F_116_DESCRICAO")
'    TreDados.Nodes.Add RS.Fields("TAC_ITEM_F_116_CODIGO") & i & "F", tvwChild, RS.Fields("TAC_ITEM_G_116_CODIGO") & i & "G", RS.Fields("TAC_ITEM_G_116_CODIGO") & " - " & RS.Fields("TAC_ITEM_G_116_DESCRICAO")
'    TreDados.Nodes.Add RS.Fields("TAC_ITEM_G_116_CODIGO") & i & "G", tvwChild, RS.Fields("TAC_ITEM_H_116_CODIGO") & i & "H", RS.Fields("TAC_ITEM_H_116_CODIGO") & " - " & RS.Fields("TAC_ITEM_H_116_DESCRICAO")
'    Else
'    If Not IsNull(RS.Fields("TAC_ITEM_A_116_CODIGO")) Then
'    TreDados.Nodes.Add , , RS.Fields("TAC_ITEM_A_116_CODIGO") & "FILHO", RS.Fields("TAC_ITEM_A_116_CODIGO") & " - " & RS.Fields("TAC_ITEM_A_116_DESCRICAO")
'    End If
'    End If
    Exit Sub
trata:
    GoTo Proximo
End Sub

Private Sub AmdSair_Click()
    Unload Me
End Sub


Private Sub CmaImprimir_Click()
    Dim Rpt As New VSRelatorio
    
    With Rpt
        If .DefinirArquivo(Bdados, App.Path & "\TAtividadeComparativa.rpt") = False Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        If txtItem <> "" Then
            .Selecao = "{TAB_ATIVIDADE_COMPARATIVO.TAC_ITEM_406_CODIGO} = " & txtItem
        End If
        .Visualizar
    End With
End Sub

Private Sub cmdBuscar_Click()
    MontaTre
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    MontaTre
    LblDescricao = ""
End Sub

Private Sub txtItem_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    
    If txtItem = "" Then
        LblDescricao = ""
        Exit Sub
    End If
    Sql = "Select * from TAB_ATIVIDADE_COMPARATIVO where 1 = 1"
    Sql = Sql & "  and TAC_ITEM_406_CODIGO = '" & txtItem & "'"
    If Bdados.AbreTabela(Sql) Then
        LblDescricao = "" & Bdados.Tabela!tac_item_406_descricao
    End If
    
End Sub
