VERSION 5.00
Begin VB.Form dec006 
   Caption         =   "SIAT - Declaração Eletrônica: Pesquisa de Contadores"
   ClientHeight    =   5430
   ClientLeft      =   1380
   ClientTop       =   1335
   ClientWidth     =   7650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5430
   ScaleWidth      =   7650
   Begin VB.Frame Frame2 
      Height          =   3945
      Left            =   0
      TabIndex        =   10
      Top             =   1380
      Width           =   7605
   End
   Begin VB.Data data_contador 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   6150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   780
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados de Pesquisa:"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   30
      Width           =   7515
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   180
         MaxLength       =   32
         TabIndex        =   2
         Top             =   510
         Width           =   7065
      End
      Begin VB.ComboBox cmb_uf_conta 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.TextBox txt_num_crc 
         Height          =   285
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox cmd_sair 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   6000
         ScaleHeight     =   300
         ScaleWidth      =   1155
         TabIndex        =   11
         Tag             =   "sair"
         Top             =   840
         Width           =   1215
      End
      Begin VB.PictureBox cmd_cancelar 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4920
         ScaleHeight     =   300
         ScaleWidth      =   945
         TabIndex        =   4
         Tag             =   "cancelar"
         Top             =   840
         Width           =   1005
      End
      Begin VB.PictureBox cmd_pesquisar 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3840
         ScaleHeight     =   300
         ScaleWidth      =   945
         TabIndex        =   3
         Tag             =   "sair"
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Nome do Contador:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Numero do CRC:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "UF do CRC:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
   End
End
Attribute VB_Name = "dec006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pesquisei As Integer


Private Sub cmb_uf_conta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Or KeyAscii = 9 Then 'Se teclar enter
    If Trim(cmb_uf_conta.Text) = "" Then
       txt_nome.SetFocus
       Exit Sub
     Else
        cmd_pesquisar.SetFocus
     End If
 End If

End Sub


Private Sub cmb_uf_conta_LostFocus()

  txt_num_crc.Text = ""
  txt_nome.Text = ""
  
  
  retorno = ""
  If Trim(cmb_uf_conta.Text) <> "" Then
    retorno = CritUF(cmb_uf_conta.Text)
    If Len(retorno) > 0 Then
       MsgBox "UF do CRC inválida.", 48, NOME_SISTEMA_DEFAULT
       cmb_uf_conta.SetFocus
       Exit Sub
    End If
    cmd_pesquisar.SetFocus

  End If
  
End Sub


Private Sub cmd_Cancelar_Click()
On Error Resume Next

   txt_num_crc = ""
   Unload Me

End Sub

Private Sub cmd_pesquisar_Click()

On Error Resume Next

Dim Campo As Integer
 Dim achou As Integer
 achou = 0
Screen.MousePointer = 11
'If Trim(txt_num_crc) <> "" Then
'   campo = 1
'   Call grd_selecao_por_crc(data_contador, tdbg_contador, Trim(Str(txt_num_crc)), achou)
'Else
'   If Trim(cmb_uf_conta.Text) <> "" Then
'       campo = 2
'      Call grd_selecao_por_uf(data_contador, tdbg_contador, Trim(cmb_uf_conta.Text), achou)
'   Else
      If Trim(txt_nome) <> "" Then
          Campo = 3
         Call grd_selecao_por_nome(data_contador, tdbg_contador, Trim(txt_nome), achou)
      Else
         Call Grd_Selecao_contador(data_contador, tdbg_contador, achou)

        'MsgBox "Nenhum Critério de Pesquisa foi informado.", MB_ICONEXCLAMATION, "Atenção!"
       ' txt_nome.SetFocus
      '  Exit Sub
      End If
  ' End If
'End If

If achou = 1 Then
 pesquisei = False
 MsgBox "Nenhum Registro foi encontrado com o Critério informado.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
 If Campo = 1 Then
   txt_num_crc.SetFocus
 ElseIf Campo = 2 Then
   cmb_uf_conta.SetFocus
 ElseIf Campo = 3 Then
   txt_nome.SetFocus
 End If
 Screen.MousePointer = 0
 Exit Sub
Else
   If achou = 2 Then
      pesquisei = False
      MsgBox "Não existe nenhum Contador registrado.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
   Else
    pesquisei = True
   End If
End If
Screen.MousePointer = 0
End Sub

Private Sub cmd_sair_Click()
Dim i As Integer
On Error Resume Next
    If pesquisei = True Then
      If busca_total_contador = True Then
            If teve_alteracao_contador = True Then
               MsgBox "Existe um cadastro aberto com alterações que não foram salvas. Esta busca não poderá ser efetuada.", 64, NOME_SISTEMA_DEFAULT
            Else
               If tdbg_contador.SelBookmarks.Count <> 0 Then
                    g_nro_contador = data_contador.Recordset("tco_num_crc")
                    Call monta_tela_contador(data_contador)
                    dec001.cmd_imprime_conta.Enabled = True
                    dec001.cmd_salvar_conta.Enabled = True
                    dec001.cmd_excluir_conta.Enabled = True
              Else
                   MsgBox$ "Nenhum contador selecionado.", 48, NOME_SISTEMA_DEFAULT
                   Exit Sub
              End If
           End If
      Else
          If tdbg_contador.SelBookmarks.Count <> 0 Then
             dec001.txt_num_crc = data_contador.Recordset("tco_num_crc") & ""
             For i = 0 To dec001.cmb_uf_crc_conta.ListCount - 1
                  If data_contador.Recordset("cmb_uf_crc_conta") = Trim(dec001.cmb_uf_crc_conta.List(i)) Then
                     dec001.cmb_uf_crc_conta.ListIndex = i 'data_contrib.Recordset("tco_uf_crc") & ""
                     Exit For
                    End If
             Next i
          Else
              MsgBox$ "Nenhum contador selecionado.", 48, NOME_SISTEMA_DEFAULT
              Exit Sub
          End If
      End If
   End If
       
Unload Me
'dec001.Show
End Sub


Private Sub Form_Load()

   CentrarFormu Me

   carComboUF dec006.cmb_uf_conta

   g_nro_contador = ""
   pesquisei = False
   
   'Call Grd_Selecao_contador(data_contador, tdbg_contador)
   tdbg_contador.MarqueeStyle = 5
   tdbg_contador.EditActive = False
End Sub


Private Sub tdbg_contador_Click()
    
    tdbg_contador.MarqueeStyle = 3
 '     tdbg_contador.SelBookmarks.Add tdbg_contador.RowBookmark(tdbg_contador.Row)
    pesquisei = True
    
End Sub

Private Sub tdbg_contador_DblClick()

   pesquisei = True
   cmd_sair_Click

End Sub


Private Sub tdbg_contador_GotFocus()
   Call tdbg_contador_Click
   'tdbg_contador.MarqueeStyle = 1

End Sub


Private Sub txt_nome_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Or KeyAscii = 9 Then 'Se teclar enter
    If Trim(txt_nome.Text) = "" Then
       Exit Sub
    Else
      cmd_pesquisar.SetFocus
    End If
 End If
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 39 Then
   KeyAscii = 96
End If
End Sub


Private Sub txt_num_crc_GotFocus()

 txt_num_crc.SelStart = 0
 txt_num_crc.SelLength = Len(txt_num_crc)
End Sub


Private Sub txt_num_crc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Or KeyAscii = 9 Then 'Se teclar enter
    If Trim(txt_num_crc) = "" Then
       txt_nome.SetFocus
       Exit Sub
    End If

   cmd_pesquisar.SetFocus
 End If
  If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then KeyAscii = 0: Beep

End Sub


Private Sub txt_num_crc_LostFocus()

    If TypeOf Me.ActiveControl Is CommandButton Then
       If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "cancelar" Or Me.ActiveControl.Tag = "buscar" Then
          Exit Sub
       End If
    End If


If Trim(txt_num_crc) = "" Then
   Exit Sub
End If
If Trim(txt_num_crc) = "" Then
      MsgBox "Número de CRC Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
      txt_num_crc.SetFocus
      Exit Sub
   End If
   reta = CRIT_CRC(txt_num_crc)
   If reta <> 0 Then
      MsgBox "CRC inválido.", 48, NOME_SISTEMA_DEFAULT
      txt_num_crc.SetFocus
      Exit Sub
   End If
   
 txt_nome.Text = ""
 
 cmd_pesquisar.SetFocus

End Sub


