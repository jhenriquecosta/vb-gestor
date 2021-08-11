VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form PCAU301 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PCAU301.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstView 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2955
      ItemData        =   "PCAU301.frx":08CA
      Left            =   75
      List            =   "PCAU301.frx":08E3
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Status de cada usuário"
      Top             =   975
      Width           =   6180
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1138
      Icone           =   "PCAU301.frx":0922
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   405
      Left            =   2760
      TabIndex        =   3
      Top             =   3960
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   714
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdCancelar 
      Height          =   405
      Left            =   3960
      TabIndex        =   4
      Top             =   3960
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   714
      Caption         =   "&Cancelar"
      Acao            =   9
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   405
      Left            =   5160
      TabIndex        =   5
      Top             =   3960
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   714
      Caption         =   "&Sair"
      Acao            =   7
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdAbrir 
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   714
      Caption         =   "&Marcar"
      Acao            =   1
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdAbrir 
      Height          =   405
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   3960
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   714
      Caption         =   "&Desmarcar"
      Acao            =   9
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   " Usuário e suas situações"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   75
      TabIndex        =   6
      Top             =   675
      Width           =   6180
   End
End
Attribute VB_Name = "PCAU301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbrir_Click(Index As Integer)
    On Error GoTo Trata
    Dim I As Integer
    For I = 0 To lstView.ListCount - 1
        lstView.Selected(I) = IIf(Index = 0, True, False)
        lstView.ListIndex = IIf(Index = 0, lstView.ListCount - 1, 0)
        lstView.Refresh
    Next
    
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdCancelar_Click()
    CarregaView lstView
End Sub


Private Sub cmdSalvar_Click()
    Dim I As Integer
    Dim X As Integer
    Dim Us As String
    
    On Error GoTo Trata
    Screen.MousePointer = 11
    
    For I = 0 To lstView.ListCount - 1

        lstView.ListIndex = I
        X = InStr(1, lstView.Text, "-", vbTextCompare)
        Us = Mid(lstView.Text, 1, X - 2)
        
        If Not BDados.AtualizaDados("TAB_USUARIO", _
        BDados.PreparaValor(IIf(lstView.Selected(lstView.ListIndex), "1", "0")), "TUS_ATIVO", _
        "TUS_COD_USUARIO = '" & (Us) & "'") Then
            Util.Erro "Registro não gravado: " & lstView.Text
        End If
        
    Next
    
    Screen.MousePointer = 0
    Util.Informa "Informações salvas com segurança."
    Exit Sub
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
        
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Trata
    
'Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"
    CarregaView lstView
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
        
    End If
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub lstView_Click()
    lstView.Refresh
End Sub

Sub CarregaView(lst As ListBox)
    On Error GoTo Trata
    Dim RS1 As Object
    Dim Views As String
    
    lst.Clear
    Views = "SELECT * FROM TAB_USUARIO"
    If BDados.AbreTabela(Views, RS1) Then
        Do Until RS1.EOF
            If Not IsNull(RS1(0)) Then
                lst.AddItem (RS1!TUS_COD_USUARIO) & " - " & (RS1!TUS_NOME), lst.ListCount
                lst.ListIndex = lst.ListCount - 1
                lst.Selected(lst.ListIndex) = IIf(RS1!TUS_ATIVO, True, False)
                DoEvents
            End If
            RS1.MoveNext
        Loop
    End If
    BDados.FechaTabela RS1
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub
