VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form dec003 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SIAT - Declaração Eletrônica: Arquivo de Declarações"
   ClientHeight    =   5610
   ClientLeft      =   360
   ClientTop       =   915
   ClientWidth     =   8835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   HelpContextID   =   3
   Icon            =   "DEC003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5610
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1500
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox SSFrame1 
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   6045
      TabIndex        =   13
      Top             =   0
      Width           =   6105
      Begin VB.ComboBox cmb_tipo 
         Height          =   300
         ItemData        =   "DEC003.frx":0442
         Left            =   3240
         List            =   "DEC003.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   945
         Width           =   1605
      End
      Begin VB.ComboBox cmbSit 
         Height          =   300
         ItemData        =   "DEC003.frx":0446
         Left            =   3240
         List            =   "DEC003.frx":0456
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox TxtCRC 
         Height          =   315
         Left            =   1800
         MaxLength       =   7
         TabIndex        =   4
         Top             =   945
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtInsc 
         Height          =   285
         Left            =   210
         TabIndex        =   0
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Format          =   "@@.@@@.@@-@"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskDataGeracao 
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   940
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "@@/@@/@@@@"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskPeriodo 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Tag             =   "periodo"
         Top             =   420
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   7
         Format          =   "@@/@@@@"
         PromptChar      =   "_"
      End
      Begin VB.PictureBox cmdPesquisar 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4965
         ScaleHeight     =   315
         ScaleWidth      =   915
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.PictureBox cmdCancelar 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4950
         ScaleHeight     =   315
         ScaleWidth      =   915
         TabIndex        =   7
         Tag             =   "cancelar"
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Período:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1800
         TabIndex        =   15
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estadual:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   210
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Situação:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3240
         TabIndex        =   16
         Top             =   210
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Data Geração:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   735
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CRC Contador:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1800
         TabIndex        =   18
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Condição:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3240
         TabIndex        =   19
         Top             =   735
         Width           =   720
      End
   End
   Begin VB.PictureBox cmdSair 
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   7500
      ScaleHeight     =   285
      ScaleWidth      =   1125
      TabIndex        =   12
      Tag             =   "sair"
      Top             =   540
      Width           =   1185
   End
   Begin VB.PictureBox cmdDeclara 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   7500
      ScaleHeight     =   285
      ScaleWidth      =   1125
      TabIndex        =   11
      Top             =   150
      Width           =   1185
   End
   Begin VB.PictureBox cmdExclDecl 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   6300
      ScaleHeight     =   285
      ScaleWidth      =   1125
      TabIndex        =   10
      Top             =   930
      Width           =   1185
   End
   Begin VB.PictureBox cmdImpDecl 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   6300
      ScaleHeight     =   285
      ScaleWidth      =   1125
      TabIndex        =   9
      Top             =   540
      Width           =   1185
   End
   Begin VB.PictureBox cmdGerEnvio 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   6300
      ScaleHeight     =   285
      ScaleWidth      =   1125
      TabIndex        =   8
      Top             =   150
      Width           =   1185
   End
End
Attribute VB_Name = "dec003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit








Sub limpa_ventana()
TxtInsc.Text = ""
MskPeriodo.Text = ""
'MskPeriodo.Mask = "##/####"
MskDataGeracao.Text = ""
'MskDataGeracao.Mask = "##/##/####"
TxtCRC.Text = ""
cmdGerEnvio.Enabled = False
cmdImpDecl.Enabled = False
cmdExclDecl.Enabled = False

cmbSit.ListIndex = -1
cmb_tipo.ListIndex = -1

cmdGerEnvio.Enabled = False
cmdImpDecl.Enabled = False

cmdDeclara.Enabled = False
cmdExclDecl.Enabled = False

Data1.RecordSource = "select * from tab_declaracao where tde_estado = 'X'"
Data1.Refresh
Data1.Refresh
End Sub



Private Sub cmb_tipo_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
End Sub


Private Sub cmbSit_KeyPress(KeyAscii As Integer)

   Call PulaCampo(KeyAscii)

End Sub





Private Sub cmdCancelar_Click()
TxtInsc.Text = ""
MskPeriodo.Text = ""
cmbSit.ListIndex = 3
cmb_tipo.ListIndex = 3
MskDataGeracao.Text = ""
TxtCRC.Text = ""

If DBDados.SelBookmarks.Count > 0 Then
    gformulario = 0
    g_num_inscricao = ""
    g_periodo_inicial = ""
    g_tipo_decla = 0
    gformulario = 0
    g_condicao = 0

Else
       gformulario = 1
 End If

End Sub

Private Sub cmdDeclara_Click()

On Error Resume Next

If DBDados.SelBookmarks.Count <> 0 Then
   If DBDados.SelBookmarks.Count > 1 Then
      MsgBox "Só uma Declaração pode ser selecionada.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
      Exit Sub
   Else
     If teve_alteracao_declaracao = True Then
        MsgBox "Existe uma declaração aberta com alterações que não foram salvas. Esta busca não poderá ser efetuada.", 64, NOME_SISTEMA_DEFAULT
        Unload Me
        gformulario = 1 ''' hoje 11/3/98
     Else
        Data1.Recordset.Bookmark = DBDados.SelBookmarks(DBDados.SelBookmarks.Count - 1)
        g_num_inscricao = DBDados.Columns(0).Value
        g_periodo_inicial = DBDados.Columns(2).Value
        g_tipo_decla = Val(DBDados.Columns(7).Value)
        g_condicao = DBDados.Columns(6).Value
        'Unload Me
        gformulario = 1
        If gformaAnte <> 0 Then
           Me.Hide
           gformaAnte = 8
        Else
           Unload Me
           gformaAnte = 0
        End If
        dec002.Show
     
     End If
   
   End If
      
End If
   
End Sub

Private Sub cmdExclDecl_Click()

On Error Resume Next

Dim sel As Integer
Dim a As Integer
Dim Banco As Database
Dim result As Integer

Dim query2 As String
sel = DBDados.SelBookmarks.Count - 1

If DBDados.SelBookmarks.Count <> 0 Then
     resposta = enviar_mensagem("Excluir registro(s) Selecionado(s)? ")
'    result = MsgBox("Excluir registro(s) Selecionado(s)? ", vbYesNo, NOME_SISTEMA_DEFAULT)
    
    If resposta = 1 Then
    
        For a = 0 To sel
            Data1.Recordset.Bookmark = DBDados.SelBookmarks(a)
            DBDados.AllowDelete = True
            query = "Delete * from tab_declaracao  where tde_num_inscricao = '" & DBDados.Columns(0).Value _
            & "' And tde_periodo_inicial = '" & DBDados.Columns(2).Value _
            & "' And Tde_substitutiva = " & Mid(DBDados.Columns(7).Value, 1, 1)
            
            query2 = "Delete * from tab_detalhe_declaracao  where tdd_num_inscricao = '" & DBDados.Columns(0).Value _
            & "' And tdd_periodo_inicial = '" & DBDados.Columns(2).Value _
            & "' And Tdd_substitutiva = " & Mid(DBDados.Columns(7).Value, 1, 1)
                                  
          '  Set Banco = OpenDatabase(App.Path + "\Dief.mdb", dbOpenDynaset)
             If Not EjecutarQuery(query) Then
                Screen.MousePointer = 0
              End If
           ' Banco.Execute query
             If Not EjecutarQuery(query2) Then
                Screen.MousePointer = 0
              End If
          '  Banco.Execute query2
           ' Banco.Close
         Next a
         
    End If
End If
 
limpa_ventana
cmdPesquisar_Click
'Data1.Refresh
'DBDados.Refresh

End Sub

Private Sub cmdGerEnvio_Click()

On Error Resume Next

    Dim laco As Integer
   Dim SelBks As SelBookmarks
 
   Set SelBks = DBDados.SelBookmarks
   
   If DBDados.SelBookmarks.Count <> 0 Then
        
        ReDim gera(DBDados.SelBookmarks.Count)
        
        For laco = 0 To DBDados.SelBookmarks.Count - 1
           If DBDados.Columns(3).Value <> "I" Then
              Data1.Recordset.Bookmark = DBDados.SelBookmarks(laco)
              gera(laco + 1).inscricao = DBDados.Columns(0).Value
              gera(laco + 1).periodo_inicial = DBDados.Columns(2).Value
             ' gera(laco + 1).periodo_final = DBDados.Columns(2).Value
              gera(laco + 1).tipo_decla = DBDados.Columns(7).Value
              gera(laco + 1).data_geracao = DBDados.Columns(4).Value
              gera(laco + 1).condicao = DBDados.Columns(6).Value
           Else
             MsgBox "Não pode selecionar Declaração Incompleta. Selecione nova(s) Declarações.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
             Exit Sub
           End If
         Next
         
         GerEnvio.Show 1
         
         Data1.Refresh
         
   Else
        
        MsgBox "Selecione uma Declaração.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
   
   End If
           
End Sub



Private Sub cmdImpDecl_Click()
    
   Dim laco As Integer
   Dim sel As Integer
   Dim Quant As Integer
   Dim SelBks As SelBookmarks
   Dim result As Integer
   Dim i
   
   Set SelBks = DBDados.SelBookmarks
  
  
         Me.MousePointer = 11
         If DBDados.SelBookmarks.Count <> 0 Then
              ReDim gera(DBDados.SelBookmarks.Count - 1)
              For laco = 0 To DBDados.SelBookmarks.Count - 1
                    Data1.Recordset.Bookmark = DBDados.SelBookmarks(laco)
                    Call Impr_declara(DBDados.Columns(0).Value, DBDados.Columns(1).Value, DBDados.Columns(2).Value, DBDados.Columns(3).Value, DBDados.Columns(4).Value, DBDados.Columns(5).Value, DBDados.Columns(6).Value, DBDados.Columns(7).Value)
              Next
             
         Else
              MsgBox "Selecione uma Inscrição", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
         End If
         Me.MousePointer = 0
  

End Sub

Private Sub cmdPesquisar_Click()

On Error Resume Next

Dim estado As Byte
Dim data_wk As String
 ' query = "Select * From tab_cad_contribuintes, tab_declaracao  where "
   query = "SELECT tab_cad_contribuintes.*, tab_declaracao.* " _
           & "FROM tab_cad_contribuintes , tab_declaracao where"
           
If Trim(MskDataGeracao.ClipText) <> "" Then
   data_wk = Mid(MskDataGeracao.ClipText, 5, 4) & Mid(MskDataGeracao.ClipText, 3, 2) & Mid(MskDataGeracao.ClipText, 1, 2)
End If
           
If Len(TxtInsc.Text) <> 0 Then
   query = query & " tab_cad_contribuintes.tcc_nu_Inscricao = '" & TxtInsc.Text & "' " & " and "
  If Trim(MskPeriodo.ClipText) <> "" Then
     query = query & " tab_declaracao.tde_periodo_inicial = '" & Mid$(MskPeriodo.ClipText, 3, 4) & Mid$(MskPeriodo.ClipText, 1, 2) & "' " & " and "
  End If
  If Trim(cmbSit.Text) <> "" Then
     query = query & " tab_declaracao.tde_Estado = '" & Mid(cmbSit.Text, 1, 1) & "' " & " and "
  End If
  If Trim(MskDataGeracao.ClipText) <> "" Then
      query = query & " tab_declaracao.tde_DA_geracao = '" & data_wk & "' " & " and "
  End If
  If Trim(TxtCRC.Text) <> "" Then
      query = query & " Tab_cad_contribuintes.tcc_num_crc_contador = '" & TxtCRC.Text & "' " & " and "
  End If
  If Trim(cmb_tipo.Text) <> "" Then
        estado = cmb_tipo.ListIndex + 1
        If estado = 3 Then
             query = query & " tab_declaracao.tde_condicao <> 1 AND tab_declaracao.tde_condicao <> 2  AND "
        Else
             query = query & " tab_declaracao.tde_condicao = " & estado & " and "
        End If
    End If
Else
    If Trim(MskPeriodo.ClipText) <> "" Then
       query = query & " tab_declaracao.tde_periodo_inicial = '" & Mid$(MskPeriodo.ClipText, 3, 4) & Mid$(MskPeriodo.ClipText, 1, 2) & "' " & " and "
        If Trim(cmbSit.Text) <> "" Then
           query = query & " tab_declaracao.tde_Estado = '" & Mid(cmbSit.Text, 1, 1) & "' " & " and "
        End If
        If Trim(MskDataGeracao.ClipText) <> "" Then
           query = query & " tab_declaracao.tde_DA_geracao = '" & data_wk & "' " & " and "
        End If
        If Trim(TxtCRC.Text) <> "" Then
           query = query & " Tab_cad_contribuintes.tcc_num_crc_contador = '" & TxtCRC.Text & "' " & " and "
        End If
        If Trim(cmb_tipo.Text) <> "" Then
             estado = cmb_tipo.ListIndex + 1
             If estado = 3 Then
                  query = query & " tab_declaracao.tde_condicao <> 1 AND tab_declaracao.tde_condicao <> 2  AND "
             Else
                  query = query & " tab_declaracao.tde_condicao = " & estado & " and "
             End If
        End If
   Else
      If Trim(cmbSit.Text) <> "" Then
         query = query & " tab_declaracao.tde_Estado = '" & Mid(cmbSit.Text, 1, 1) & "' " & " and "
         If MskDataGeracao.ClipText <> "" Then
            query = query & " tab_declaracao.tde_DA_geracao = '" & data_wk & "' " & " and "
         End If
         If Trim(TxtCRC.Text) <> "" Then
            query = query & " Tab_cad_contribuintes.tcc_num_crc_contador = '" & TxtCRC.Text & "' " & " and "
         End If
         If Trim(cmb_tipo.Text) <> "" Then
             estado = cmb_tipo.ListIndex + 1
             If estado = 3 Then
                  query = query & " tab_declaracao.tde_condicao <> 1 AND tab_declaracao.tde_condicao <> 2  AND "
             Else
                  query = query & " tab_declaracao.tde_condicao = " & estado & " and "
             End If
        End If
     Else
       If Trim(MskDataGeracao.ClipText) <> "" Then
'          query = query & " tab_declaracao.tde_DA_geracao = '" & Format(MskDataGeracao.ClipText, "yyyymmdd") & "' " & " and "
          query = query & " tab_declaracao.tde_DA_geracao = '" & data_wk & "' " & " and "
          If Trim(TxtCRC.Text) <> "" Then
             query = query & " Tab_cad_contribuintes.tcc_num_crc_contador = '" & TxtCRC.Text & "' " & " and "
          End If
          If Trim(cmb_tipo.Text) <> "" Then
             estado = cmb_tipo.ListIndex + 1
             If estado = 3 Then
                  query = query & " tab_declaracao.tde_condicao <> 1 AND tab_declaracao.tde_condicao <> 2  AND "
             Else
                  query = query & " tab_declaracao.tde_condicao = " & estado & " and "
             End If
          End If
      Else
         If Trim(TxtCRC.Text) <> "" Then
             query = query & " Tab_cad_contribuintes.tcc_num_crc_contador = '" & TxtCRC.Text & "' " & " and "
             If Trim(cmb_tipo.Text) <> "" Then
               estado = cmb_tipo.ListIndex + 1
               query = query & "  tab_declaracao.tde_condicao = " & estado & " and "
             End If
         Else
             If Trim(cmb_tipo.Text) <> "" Then
                estado = cmb_tipo.ListIndex + 1
                If estado = 3 Then
                   query = query & " tab_declaracao.tde_condicao <> 1 AND tab_declaracao.tde_condicao <> 2  AND "
                Else
                  query = query & " tab_declaracao.tde_condicao = " & estado & " and "
                End If
         Else
            query = "SELECT tab_cad_contribuintes.*, tab_declaracao.* " _
            & "FROM tab_cad_contribuintes INNER JOIN tab_declaracao ON"
        End If
        End If
      End If
    End If
    End If
  
End If
query = query & " tab_cad_contribuintes.tcc_nu_inscricao = tab_declaracao.tde_num_inscricao"
query = query & " ORDER BY  tab_cad_contribuintes.tcc_nu_inscricao asc, tab_declaracao.tde_periodo_inicial DESC , tab_declaracao.tde_Estado  DESC"

Set PESQ_SNAP = dbgim.OpenRecordset(query, dbOpenSnapshot)
If PESQ_SNAP.EOF And PESQ_SNAP.BOF Then
    Me.Data1.DatabaseName = App.Path + "\Dief.mdb"
    Me.Data1.RecordSource = query
    Data1.Refresh
    DBDados.Refresh
   DBDados.Enabled = False
   MsgBox "Não existe nenhuma Declaração registrada.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
   Exit Sub
 Else
     DBDados.Enabled = True
 End If


Me.Data1.DatabaseName = App.Path + "\Dief.mdb"
Me.Data1.RecordSource = query

Data1.Refresh
DBDados.Refresh

DBDados.SetFocus


End Sub

Private Sub cmdSair_Click()

On Error Resume Next


If gformulario = 2 Then
  If teve_alteracao_declaracao = True Then
     MsgBox "Existe uma declaração aberta com alterações que não foram salvas. Esta busca não poderá ser efetuada.", 64, NOME_SISTEMA_DEFAULT
     gformulario = 1
  Else
     If DBDados.SelBookmarks.Count > 0 Then
       If DBDados.SelBookmarks.Count > 1 Then
          MsgBox "Só uma Declaração pode ser selecionada.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
          Exit Sub
       Else
          If teve_alteracao_declaracao = True Then
             MsgBox "Existe uma declaração aberta com alterações que não foram salvas. Esta busca não poderá ser efetuada.", 64, NOME_SISTEMA_DEFAULT
             Unload Me
          Else
           gformulario = 1
           Data1.Recordset.Bookmark = DBDados.SelBookmarks(DBDados.SelBookmarks.Count - 1)
           g_num_inscricao = DBDados.Columns(0).Value
           g_periodo_inicial = DBDados.Columns(2).Value
           g_tipo_decla = Val(DBDados.Columns(7).Value)
         End If
      End If
     Else
        If DBDados.Enabled = True Then
           MsgBox "Nenhuma Declaração foi selecionada.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
           Exit Sub
        Else
           If g_num_inscricao = "" Then
              gformulario = 0
           Else
              g_condicao = 0
              gformulario = 1
           End If
       End If
     End If
 
      Unload Me
      gformulario = 1
      gformaAnte = 8
  End If
Else
    Unload Me
    gformaAnte = 0
End If
 
   
   
   
End Sub














Private Sub DBDados_Click()
   
    If gformulario = 2 Then
        cmdExclDecl.Enabled = False  'bellyrio em 16/02/98
        cmdDeclara.Enabled = False
        cmdGerEnvio.Enabled = True
        cmdImpDecl.Enabled = True
    Else
     If gformulario = 3 Then
       cmdExclDecl.Enabled = True
       cmdDeclara.Enabled = True
       cmdGerEnvio.Enabled = True
       cmdImpDecl.Enabled = True
     End If
  End If
    
End Sub

Private Sub DBDados_DblClick()
   cmdDeclara_Click
End Sub

Private Sub DBDados_FormatText(ByVal ColIndex As Integer, Value As Variant)
   
   If ColIndex = 4 And Value <> "" Then
      Value = Mid(Value, 7, 2) & "/" & Mid(Value, 5, 2) & "/" & Mid(Value, 1, 4)
   ElseIf ColIndex = 2 And Value <> "" Then
      Value = Mid(Value, 5, 2) & "/" & Mid(Value, 1, 4)
   End If
   
   
'   If ColIndex = 7 And Value = "0" Then
'      Value = "Não"
'   ElseIf ColIndex = 7 And Value <> "0" Then
'      Value = "Sim"
'   End If
  
  If ColIndex = 7 And Value = "0" Then
      Value = "Original"
   ElseIf ColIndex = 7 And Value <> "0" Then
      Value = "Substitutiva" & Value
   End If
  
End Sub

Private Sub DBDados_GotFocus()
   
   DBDados.MarqueeStyle = 2
   
 '  Call DBDados_Click

If gformulario = 0 Then
   gformulario = 3
End If

End Sub

Private Sub DBDados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If DBDados.SelBookmarks.Count = 0 Then
   If gformulario = 2 Then
     cmdDeclara.Enabled = False
     cmdExclDecl.Enabled = False  'bellyrio em 16/02/98
     cmdGerEnvio.Enabled = False  'bellyrio em 16/02/98
     cmdImpDecl.Enabled = False  'bellyrio em 16/02/98
   Else
      cmdDeclara.Enabled = True
   End If
   
Else
  If gformulario = 2 Then
    cmdExclDecl.Enabled = False  'bellyrio em 16/02/98
    cmdGerEnvio.Enabled = False  'bellyrio em 16/02/98
    cmdImpDecl.Enabled = False  'bellyrio em 16/02/98
  Else
   cmdExclDecl.Enabled = True
   cmdGerEnvio.Enabled = True
   cmdImpDecl.Enabled = True
  End If
  cmdDeclara.Enabled = False
 
End If
End Sub


Private Sub DBDados_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub

Private Sub Form_Activate()
'   Me.Top = 1200
'    Me.Left = 0
CentrarFormu Me
End Sub

Private Sub Form_Load()
    Me.Top = 1200
    Me.Left = 0
vim_trata_epp = False
vim_de_onde = "DEC003"
   dec003.Enabled = True
   DBDados.Enabled = False
  ' carComboUF dec003.cmbUF
 '   Call CentrarFormu(Me)
   
      cmb_tipo.Clear
      cmb_tipo.AddItem "Normal       "
      cmb_tipo.AddItem "Microempresa "
      cmb_tipo.AddItem "Pequeno Porte"
      cmb_tipo.AddItem "             "
      
   Me.DBDados.Columns(0).Width = 1200
   Me.DBDados.Columns(1).Width = 2200
   Me.DBDados.Columns(2).Width = 800   '750
   Me.DBDados.Columns(3).Width = 1050 ' 900
   Me.DBDados.Columns(4).Width = 1000
   Me.DBDados.Columns(5).Width = 900
   Me.DBDados.Columns(6).Width = 900
   Me.DBDados.Columns(7).Width = 1800 '1200
   
   
If gformulario = 2 Then
  cmdDeclara.Enabled = False
'  cmdSair.Caption = "Aceita&r"
  cmdExclDecl.Enabled = False  'bellyrio em 16/02/98
  cmdGerEnvio.Enabled = False  'bellyrio em 16/02/98
  cmdImpDecl.Enabled = False  'bellyrio em 16/02/98
Else
  cmdSair.Caption = "Sai&r"
End If
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   dec003.Enabled = False

End Sub


Private Sub Form_Unload(Cancel As Integer)
If MDIdec.Enabled = True Then
    MDIdec.Picture1.SetFocus
End If
End Sub

Private Sub Label6_Click()
TxtCRC.SetFocus
End Sub

Private Sub MskDataGeracao_GotFocus()

  MskDataGeracao.SelStart = 0
  MskDataGeracao.SelLength = Len(MskDataGeracao)
   
End Sub

Private Sub MskDataGeracao_KeyPress(KeyAscii As Integer)
    Call PulaCampo(KeyAscii)
End Sub


Private Sub MskDataGeracao_LostFocus()
Dim data_micro As String
Dim data_inicio As String


If Trim(MskDataGeracao.ClipText) = "" Then
  Exit Sub
End If

If Len(Trim(MskDataGeracao.ClipText)) < 8 Then
   MsgBox "Data Inválida. Data no formato DDMMAAA.", 48, NOME_SISTEMA_DEFAULT
   MskDataGeracao = ""
   MskDataGeracao.SetFocus
  Exit Sub
End If
If Len(Trim(MskDataGeracao.ClipText)) > 8 Then
   MsgBox "Data Inválida. Data no formato DDMMAAA.", 48, NOME_SISTEMA_DEFAULT
   MskDataGeracao = ""
   MskDataGeracao.SetFocus
  Exit Sub
End If


reta = CRIT_Data(MskDataGeracao.ClipText)
If reta <> 0 Then
   MsgBox "Data Inválida. Data no formato DDMMAAA.", 48, NOME_SISTEMA_DEFAULT
   MskDataGeracao = ""
   MskDataGeracao.SetFocus
   Exit Sub
End If

data_inicio = Val(Mid$(MskDataGeracao, 5, 4) & Mid$(MskDataGeracao, 3, 2)) & Mid$(MskDataGeracao, 1, 2)

data_micro = Val(Year(Date)) & Mid$(Date, 4, 2) & Mid$(Date, 1, 2)

If data_inicio > data_micro Then
   MsgBox "Data de Geração não pode ser maior que data atual.", 48, NOME_SISTEMA_DEFAULT
    MskDataGeracao = ""
   MskDataGeracao.SetFocus
   Exit Sub
End If


End Sub

Private Sub MskPeriodo_GotFocus()

   MskPeriodo.SelStart = 0
   MskPeriodo.SelLength = Len(MskPeriodo)

End Sub

Private Sub MskPeriodo_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)

End Sub




Private Sub MskPeriodo_LostFocus()


If Trim(MskPeriodo.ClipText) = "" Then
  Exit Sub
End If

If Val(Len(MskPeriodo.ClipText)) < 6 Then
       MsgBox "Período Inválido.Período no formato MMAAAA.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       MskPeriodo = ""
       MskPeriodo.SetFocus
       Exit Sub
End If
If Val(Len(MskPeriodo.ClipText)) > 6 Then
       MsgBox "Período Inválido.Período no formato MMAAAA.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       MskPeriodo = ""
       MskPeriodo.SetFocus
       Exit Sub
End If


 If MskPeriodo.ClipText <> "" Then
    If Mid(MskPeriodo.ClipText, 1, 2) < "01" Or Mid(MskPeriodo.ClipText, 1, 2) > "12" Then
       MsgBox "Período Inválido.Período no formato MMAAAA.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       MskPeriodo = ""
       MskPeriodo.SetFocus
       Exit Sub
     End If
     If Val(Mid(MskPeriodo.ClipText, 3, 4)) < 1900 Then
       MsgBox "Período Inválido.Período no formato MMAAAA.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       MskPeriodo = ""
       MskPeriodo.SetFocus
       Exit Sub
     End If
 End If

End Sub






Private Sub SSFrame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub


Private Sub TxtCRC_GotFocus()

 TxtCRC.SelStart = 0
  TxtCRC.SelLength = Len(TxtCRC)
 

End Sub

Private Sub TxtCRC_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)

End Sub


Private Sub TxtInsc_GotFocus()

   TxtInsc.SelStart = 0
   TxtInsc.SelLength = Len(TxtInsc)

End Sub

Private Sub TxtInsc_KeyPress(KeyAscii As Integer)
Call PulaCampo(KeyAscii)
End Sub


Private Sub TxtInsc_LostFocus()
Dim MyInscricao As String
MyInscricao = TxtInsc.Text

If TypeOf Me.ActiveControl Is SSCommand Then
      If Me.ActiveControl.Tag = "sair" Or Me.ActiveControl.Tag = "cancelar" Then
         Exit Sub
      End If
End If

  If Not IsNumeric(MyInscricao) And Trim(MyInscricao) <> "" Then
      MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       TxtInsc.SetFocus
       Exit Sub
    End If

    If GenDV(CStr(Mid(MyInscricao, 1, 8)), BASE9) <> Val(Mid(MyInscricao, 9, 1)) Then
       MsgBox "Número de Inscrição Inválido.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
       TxtInsc.SetFocus
       Exit Sub
    End If
 ' TxtInsc.Text = Format(TxtInsc.Text, "@@@.@@@.@@@")
End Sub



