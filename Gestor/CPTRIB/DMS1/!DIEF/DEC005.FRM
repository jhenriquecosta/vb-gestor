VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form dec005 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SIAT - Declaração Eletrônica: Pesquisa de Contribuintes"
   ClientHeight    =   5415
   ClientLeft      =   870
   ClientTop       =   1440
   ClientWidth     =   7665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5415
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   3825
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   7575
   End
   Begin VB.Data data_contrib 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1050
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados de Pesquisa:"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7635
      Begin VB.TextBox txt_fantasia 
         Height          =   285
         Left            =   1620
         TabIndex        =   1
         Top             =   750
         Width           =   5715
      End
      Begin VB.TextBox txt_razao 
         Height          =   285
         Left            =   1620
         TabIndex        =   0
         Top             =   360
         Width           =   5715
      End
      Begin VB.PictureBox cmd_sair 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   6120
         ScaleHeight     =   240
         ScaleWidth      =   1155
         TabIndex        =   12
         Tag             =   "sair"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.PictureBox cmd_pesquisar 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3960
         ScaleHeight     =   240
         ScaleWidth      =   945
         TabIndex        =   2
         Tag             =   "sair"
         Top             =   1080
         Width           =   1005
      End
      Begin VB.PictureBox cmd_Cancelar 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5040
         ScaleHeight     =   240
         ScaleWidth      =   945
         TabIndex        =   3
         Tag             =   "cancelar"
         Top             =   1080
         Width           =   1005
      End
      Begin MSMask.MaskEdBox msk_inscricao 
         Height          =   285
         Left            =   1590
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   11
         Format          =   "@@@.@@@.@@@"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Nome Fantasia:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nome Empresarial"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Inscricao Estadual:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Dados de pesquisa:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "dec005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pesquisei As Integer

Private Sub cmd_Cancelar_Click()
On Error Resume Next
gformulario = 0
   g_num_inscricao = ""
   g_nro_inscricao = ""
   Unload Me
  If quem_me_chamou <> "DEC002" Then
  ' dec001.Show
  Else
 '  dec002.Show 1
End If

End Sub

Private Sub cmd_pesquisar_Click()

On Error Resume Next

Dim Campo As Integer
Dim achou As Integer
achou = 0
 Screen.MousePointer = 11
'If Trim(msk_inscricao) <> "" Then
'   campo = 1
'   Call grd_selecao_por_inscricao(data_contrib, tdbg_contrib, Trim(Str(msk_inscricao)), achou)
'Else
   If Trim(txt_razao) <> "" Then
      If Trim(txt_fantasia) <> "" Then
        Campo = 2
        Call grd_selecao_por_razao_e_fantasia(data_contrib, tdbg_contrib, Trim(txt_razao), Trim(txt_fantasia), achou)
      Else
        Campo = 2
        Call grd_selecao_por_razao(data_contrib, tdbg_contrib, Trim(txt_razao), achou)
      End If
   Else
      If Trim(txt_fantasia) <> "" Then
         Campo = 3
         Call grd_selecao_por_fantasia(data_contrib, tdbg_contrib, Trim(txt_fantasia), achou)
      Else
        Call Grd_Selecao(data_contrib, tdbg_contrib, achou)

        'MsgBox "Nenhum Critério de Pesquisa foi informado.", MB_ICONEXCLAMATION, "Atenção!"
        
      '  txt_razao.SetFocus
      '  Exit Sub
      End If
   End If
'End If
If achou = 1 Then
   pesquisei = False
   MsgBox "Nenhum Registro foi encontrado com o Critério informado.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
  If Campo = 1 Then
     msk_inscricao.SetFocus
  ElseIf Campo = 2 Then
     txt_razao.SetFocus
  ElseIf Campo = 3 Then
    txt_fantasia.SetFocus
  End If
   Screen.MousePointer = 0
 Exit Sub
Else
   If achou = 2 Then
      pesquisei = False
      MsgBox "Não existe nenhum Contribuinte registrado.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
      
   Else
      pesquisei = True
   End If
End If
  Screen.MousePointer = 0
End Sub


Private Sub cmd_sair_Click()

On Error Resume Next
gformulario = 2
  If pesquisei = True Then
     If teve_alteracao = True And Trim(dec001!msk_inscricao) <> "" Then
         MsgBox "Existe um cadastro aberto com alterações que não foram salvas. Esta busca não poderá ser efetuada.", 64, NOME_SISTEMA_DEFAULT
     Else
        If tdbg_contrib.SelBookmarks.Count <> 0 Then

       ' If Not data_contrib.EOFAction Then
           g_nro_inscricao = data_contrib.Recordset("tcc_nu_inscricao")
           g_num_inscricao = data_contrib.Recordset("tcc_nu_inscricao")
           g_condicao = data_contrib.Recordset("tcc_tipo_contribuinte")
           If quem_me_chamou <> "DEC002" Then
                Call monta_tela_contribuinte(data_contrib)
                dec001.cmd_imprime_contr.Enabled = True
                dec001.cmd_salvar_contr.Enabled = True
                dec001.cmd_excluir_contr.Enabled = True
           End If
        Else
           MsgBox "Nenhum contribuinte selecionado.", 64, NOME_SISTEMA_DEFAULT
           Exit Sub
        End If
     End If
  End If
Unload Me

If quem_me_chamou <> "DEC002" Then
 '  dec001.Show
Else
'   dec002.Show 1

End If
End Sub



Private Sub Form_Load()
   
   CentrarFormu Me
vim_de_onde = "DEC005"
   Screen.MousePointer = 13
   g_nro_inscricao = ""
   pesquisei = False
   
   'Call Grd_Selecao(data_contrib, tdbg_contrib)
   
   Screen.MousePointer = 0
   tdbg_contrib.MarqueeStyle = 5
   tdbg_contrib.EditActive = False

End Sub


Private Sub msk_inscricao_Change()
   If Len(Trim(msk_inscricao)) > 9 Then
      SendKeys "{BackSpace}"
      Exit Sub
   End If

End Sub

Private Sub msk_inscricao_GotFocus()
   msk_inscricao.SelStart = 0
   msk_inscricao.SelLength = Len(msk_inscricao)

End Sub

Private Sub msk_inscricao_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Or KeyAscii = 9 Then 'Se teclar enter
    If Trim(msk_inscricao) = "" Then
       txt_razao.SetFocus
       Exit Sub
    End If

   cmd_pesquisar.SetFocus
 End If
  If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep

End Sub


Private Sub msk_inscricao_LostFocus()
Dim MyInscricao As String

    If TypeOf Me.ActiveControl Is CommandButton Then
       If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "cancelar" Or Me.ActiveControl.Tag = "buscar" Then
          Exit Sub
       End If
    End If

MyInscricao = msk_inscricao.Text
MyInscricao = tirapictu(MyInscricao, ".")

If Trim(MyInscricao) = "" Then
   Exit Sub
End If

If GenDV(CStr(Mid(MyInscricao, 1, 8)), BASE9) <> Val(Mid(MyInscricao, 9, 1)) Then
   MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
   msk_inscricao.SetFocus
   Exit Sub
End If
 txt_fantasia.Text = ""
 txt_razao.Text = ""
cmd_pesquisar.SetFocus

End Sub


Private Sub tdbg_contrib_Click()
   
     tdbg_contrib.MarqueeStyle = 3
     
    ' tdbg_contrib.SelBookmarks.Add tdbg_contrib.RowBookmark(tdbg_contrib.Row)
     
     pesquisei = True
     
End Sub

Private Sub tdbg_contrib_DblClick()

   pesquisei = True
   cmd_sair_Click
   
End Sub


Private Sub tdbg_contrib_GotFocus()

   Call tdbg_contrib_Click
   'tdbg_contrib.MarqueeStyle = 1

End Sub


Private Sub txt_fantasia_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Or KeyAscii = 9 Then 'Se teclar enter
'    If Trim(txt_fantasia.Text) = "" Then
'       Exit Sub
'    Else
'      cmd_pesquisar.SetFocus
'    End If
' End If
Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If
End Sub

Private Sub txt_fantasia_LostFocus()
msk_inscricao.Text = ""
'txt_razao.Text = ""
'If Trim(txt_fantasia) <> "" Then
'   cmd_pesquisar.SetFocus
'End If
End Sub


Private Sub txt_razao_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 
 If KeyAscii = 39 Then
   KeyAscii = 96
End If
' If KeyAscii = 13 Or KeyAscii = 9 Then 'Se teclar enter
 '   If Trim(txt_razao.Text) = "" Then
 '      txt_fantasia.SetFocus
 '      Exit Sub
 '    Else
  '      cmd_pesquisar.SetFocus
  '   End If
' End If

End Sub

Private Sub txt_razao_LostFocus()
  msk_inscricao.Text = ""
 ' txt_fantasia.Text = ""
 ' If Trim(txt_razao.Text) <> "" Then
 '    cmd_pesquisar.SetFocus
 ' End If
End Sub


