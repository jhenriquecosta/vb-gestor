VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form PCAU201 
   BackColor       =   &H00DDF1FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PCAU201.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   915
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Código do Usuário"
      Top             =   1110
      Width           =   1425
   End
   Begin VB.TextBox txtNome 
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
      Height          =   315
      Left            =   915
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Nome do Usuário"
      Top             =   1500
      Width           =   5700
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   1349
      CorFundo        =   14545407
      CorFrente       =   255
   End
   Begin MSComctlLib.TreeView Tree 
      CausesValidation=   0   'False
      Height          =   3690
      Left            =   75
      TabIndex        =   2
      Top             =   2205
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   6509
      _Version        =   393217
      Indentation     =   1058
      LabelEdit       =   1
      Style           =   6
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   405
      TabIndex        =   7
      Top             =   1575
      Width           =   405
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Acessos"
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
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   1890
      Width           =   7155
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   " Usuário"
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
      Left            =   60
      TabIndex        =   5
      Top             =   810
      Width           =   7155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   1185
      Width           =   570
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   6180
      TabIndex        =   3
      ToolTipText     =   "Deseja sair?"
      Top             =   5970
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   255
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "PCAU201.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PCAU201.frx":08E6
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
End
Attribute VB_Name = "PCAU201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo Trata
    
    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"
    Screen.MousePointer = 0
    Exit Sub
    
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
        
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub CarregaTree(lst As TreeView, User)
    On Error GoTo Trata
    Dim RS1 As Object
    Dim Rs2 As Object
    Dim RS3 As Object
    Dim Sql As String
    lst.Nodes.Clear
    
    lst.Nodes.Add , tvwFirst, "VSIS", Temp.PegaParametro(BDados, "SISTEMA")
    
    Sql = "SELECT TSI_COD_SISTEMA, TSI_NOME FROM TAB_SISTEMA WHERE " & _
    " TSI_COD_SISTEMA IN (SELECT DISTINCT TAU_TSI_COD_SISTEMA FROM " & _
    " TAB_ACESSO_USUARIO WHERE TAU_TUS_COD_USUARIO = '" & User & "') ORDER BY TSI_NOME"
    
    If BDados.AbreTabela(Sql, RS1) Then
        Do Until RS1.EOF
            lst.Nodes.Add "VSIS", tvwChild, RS1(0), RS1(1)
            
            
            Sql = "SELECT TMO_COD_MODULO, TMO_NOME FROM TAB_MODULO WHERE " & _
            " TMO_COD_MODULO IN (SELECT DISTINCT TAU_TMO_COD_MODULO FROM " & _
            " TAB_ACESSO_USUARIO WHERE TAU_TUS_COD_USUARIO = '" & _
            User & "' AND TAU_TSI_COD_SISTEMA = '" & RS1(0) & "') ORDER BY TMO_NOME"
            
            
            If BDados.AbreTabela(Sql, Rs2) Then
                Do Until Rs2.EOF
                    lst.Nodes.Add CStr(RS1(0)), tvwChild, CStr(Rs2(0)), CStr(Rs2(1))
   
                    Sql = "SELECT TFO_COD_FORMULARIO, TFO_NOME FROM TAB_FORMULARIO WHERE " & _
                    " TFO_TMO_COD_MODULO " & BDados.Concatena & " TFO_COD_FORMULARIO IN (SELECT DISTINCT TAU_TMO_COD_MODULO " & BDados.Concatena & " TAU_TFO_COD_FORMULARIO FROM " & _
                    " TAB_ACESSO_USUARIO WHERE TAU_TUS_COD_USUARIO = '" & _
                     User & "' AND TAU_TMO_COD_MODULO = '" & Rs2(0) & "') ORDER BY TFO_NOME"

                    If BDados.AbreTabela(Sql, RS3) Then
                        Do Until RS3.EOF
                            lst.Nodes.Add CStr(Rs2(0)), tvwChild, CStr(Rs2(0) & RS3(0)), CStr(RS3(1))
                            RS3.MoveNext
                        Loop
                    End If
                    BDados.FechaTabela RS3
                    
                    Rs2.MoveNext
                Loop
            End If
            BDados.FechaTabela Rs2
            RS1.MoveNext
        Loop
    End If
    BDados.FechaTabela RS1
    lst.Nodes.Item("VSIS").Expanded = True
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub
Private Sub Tree_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error GoTo Trata
    
    Node.Checked = Not Node.Checked

    
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtCodigo_GotFocus()
    txtCodigo = ""
    txtNome = ""
    Tree.Nodes.Clear
    txtCodigo.SetFocus
    Screen.MousePointer = 0
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    On Error GoTo Trata
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    DoEvents
    
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtCodigo_LostFocus()

    If txtCodigo = "" Then Exit Sub
    txtNome = UCase(Seguranca.ExisteUsuario(BDados, txtCodigo))
    If Trim(txtNome) = "" Then
        Util.Informa "Usuário não Cadastrado."
        txtCodigo = ""
        txtCodigo.SetFocus
    Else
        Screen.MousePointer = 11
        Tree.Enabled = True
        CarregaTree Tree, (txtCodigo)
        Tree.Refresh
        Tree.SetFocus
        Screen.MousePointer = 0
    End If
End Sub
