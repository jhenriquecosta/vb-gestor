VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TCIM403 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkImpresso 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      Caption         =   "Considerar impressos"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5730
      TabIndex        =   12
      Top             =   1830
      Width           =   1845
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   10
      Top             =   5550
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   979
      CorFundo        =   -2147483633
      Begin VB.ComboBox cboRelatorio 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         ItemData        =   "Tcim403.frx":0000
         Left            =   1200
         List            =   "Tcim403.frx":000A
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   135
         Width           =   2325
      End
      Begin VTOcx.cmdVISUAL cmdCancelar 
         Height          =   405
         Left            =   4860
         TabIndex        =   5
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "&Cancelar"
         Acao            =   9
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
         CorFoco         =   -2147483626
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   405
         Left            =   6090
         TabIndex        =   6
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   714
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   405
         Left            =   7140
         TabIndex        =   7
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   405
         Left            =   3630
         TabIndex        =   4
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
         CorFoco         =   -2147483626
      End
   End
   Begin VTOcx.grdVISUAL grdImoveis 
      Height          =   3645
      Left            =   180
      TabIndex        =   9
      Top             =   1800
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   4339
      CorBorda        =   16384
      CabecalhoEstado =   "Maranhão"
      CabecalhoCliente=   "Prefeitura Municipal de Imperatriz"
      CabecalhoSecretaria=   ""
      CabecalhoDepartamento=   ""
      CabecalhoTitulo =   "Imóveis"
      Caption         =   "Imóveis encontrados"
      CorTitulo       =   16384
      CorCaption      =   16777215
      CorDica         =   16384
      CheckBox        =   -1  'True
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   1138
      Icone           =   "Tcim403.frx":0025
   End
   Begin VTOcx.txtVISUAL txtIc 
      Height          =   300
      Left            =   150
      TabIndex        =   2
      Top             =   1350
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   529
      Caption         =   "IC - Imóvel"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   20
   End
   Begin VTOcx.txtVISUAL txtQuadra 
      Height          =   300
      Left            =   540
      TabIndex        =   1
      Top             =   1020
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
      Caption         =   "Quadra"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   5
   End
   Begin VTOcx.txtVISUAL txtSetor 
      Height          =   300
      Left            =   690
      TabIndex        =   0
      Top             =   690
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      Caption         =   "Setor"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   5
   End
   Begin VTOcx.cmdVISUAL cmdBuscar 
      Height          =   405
      Left            =   3210
      TabIndex        =   3
      Top             =   1260
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
      Caption         =   "&Buscar"
      Acao            =   5
      CorBorda        =   8421504
      CorFrente       =   16384
      CorFundo        =   -2147483633
      CorFoco         =   -2147483626
   End
End
Attribute VB_Name = "TCIM403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SelecaoRpt As String
Private Cancelar As Boolean
Public Enum enuTipoLoteImperatriz
    Condominio = 0
    Edificacao = 1
    Terreno = 2
End Enum
Private SetorImpressao As String, QuadraImpressao As String

Private Sub chkImpresso_Click()
    grdImoveis.MarcarTodos chkImpresso.Value
End Sub

Private Sub cmdBuscar_Click()
    Dim Where As String
    Dim sql As String
    
    sql = ""
    Where = ""
    If Trim(txtIc) <> "" Then
        'Where = Where & " and InscImob = '" & txtIc & "'"
        
        Dim Distrito As String, Setor As String, Quadra As String, Lote As String
        Distrito = Mid(txtIc, 1, 2)
        Setor = Mid(txtIc, 3, 2)
        Quadra = Mid(txtIc, 5, 3)
        Lote = Mid(txtIc, 8, 4)
        Where = Where & " and" & _
                " Substring(Cast(InscImob AS varchar),1,2)=" & Distrito & " AND " & _
                " Substring(Cast(InscImob AS varchar),3,2)=" & Setor & " AND " & _
                " Substring(Cast(InscImob AS varchar),5,3)=" & Quadra & " AND " & _
                " Substring(Cast(InscImob AS varchar),8,4)=" & Lote
    Else
        If Trim(txtSetor) <> "" Then Where = Where & " and " & Bdados.ParteTexto("InscImob", MidVs, 3, 2, True) & " ='" & txtSetor & "'"
        If Trim(txtQuadra) <> "" Then Where = Where & " and " & Bdados.ParteTexto("InscImob", MidVs, 5, 3, True) & " ='" & txtQuadra & "'"
    End If
    If Where <> "" Then Where = " WHERE " & Right(Where, Len(Where) - 4)
    sql = "SELECT InscImob, " & _
                Bdados.ParteTexto("InscImob", MidVs, 1, 2, True) & " as Dist, " & _
                Bdados.ParteTexto("InscImob", MidVs, 3, 2, True) & " as Setor, " & _
                Bdados.ParteTexto("InscImob", MidVs, 5, 3, True) & " as Quadra, " & _
                Bdados.ParteTexto("InscImob", MidVs, 8, 4, True) & " as Lote, " & _
                Bdados.ParteTexto("InscImob", MidVs, 12, 3, True) & " as Unid, " & _
                Bdados.ParteTexto("InscImob", MidVs, 15, 3, True) & " as Sub, " & _
                " AreConun, Compleme " & _
            " FROM IMOVEIS " & Where & _
            " ORDER BY " & _
                Bdados.ParteTexto("InscImob", MidVs, 1, 2, True) & ", " & _
                Bdados.ParteTexto("InscImob", MidVs, 3, 2, True) & ", " & _
                Bdados.ParteTexto("InscImob", MidVs, 5, 3, True) & ", " & _
                Bdados.ParteTexto("InscImob", MidVs, 8, 4, True) & ", " & _
                Bdados.ParteTexto("InscImob", MidVs, 12, 3, True)
    grdImoveis.CheckBox = True
    grdImoveis.Preencher Bdados, sql ', 0, 2500
    grdImoveis.MarcarTodos False
    chkImpresso.Value = vbUnchecked
End Sub

Private Sub cmdCancelar_Click()
    Cancelar = True
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo Trata
'ROTEIRO
'1. Identificar a situacao do lote
'2. Imprimir os boletins, de acordo com a situacao

    Dim Item As ListItem
    
    Cancelar = False
    Screen.MousePointer = vbArrowHourglass
    If cboRelatorio.ListIndex = 0 Then
        For Each Item In grdImoveis.ListItems
             If Not Item.Checked Then
                Item.Selected = True
                If Cancelar Then Exit For
                
                SetorImpressao = Mid(Item, 3, 2)
                QuadraImpressao = Mid(Item, 5, 3)
                 '1.
                 Select Case SituacaoLote(Item)
                     '2.
                     Case Terreno
                         Item.Checked = ImprimirBT(Item.Text)
                         
                     Case Edificacao
                         ImprimirEdificacoes Item
                         
                     Case Condominio
                         ImprimirCondominio Item
                 End Select
                Item.EnsureVisible
                DoEvents
             End If
         Next
    Else
        Dim Arquivo As String
        Dim Selecao As String
        
        Set Rpt = New VSRelatorio
        Arquivo = "TCIU403.rpt"
        If Not Rpt.DefinirArquivo(Bdados, App.Path & "\" & Arquivo) Then Exit Sub
        Selecao = ""
        If Trim$(txtIc) <> "" Then
            Selecao = "{BT.Inscricao} = '" & txtIc & "'"
        Else
            If Trim$(txtQuadra) <> "" Then Selecao = "{BT.Quadra} = '" & txtQuadra & "'"
            If Trim$(txtSetor) <> "" Then Selecao = Selecao & IIf(Selecao = "", "", " AND ") & " {BT.Setor} = '" & txtSetor & "'"
        End If
        If Selecao <> "" Then Rpt.Selecao = Selecao
        Rpt.Visualizar
        Set Rpt = Nothing
    End If
    txtSetor.SetFocus
Trata:
    Screen.MousePointer = vbNormal
End Sub

Private Function SituacaoLote(Item As ListItem) As enuTipoLoteImperatriz
'ROTEIRO
'1. Condominio : mais de uma unidade e complemento com LOJA, APT, ...
'2. Edificacao : pelo menos uma unidade com area construida
'3. Terreno : caso contrario
    Dim sql As String, Rs As VSRecordset
    Dim Where As String
    
    Dim Distrito As String
    Dim Setor As String
    Dim Quadra As String
    Dim Lote As String
    
    Distrito = Item.SubItems(1)
    Setor = Item.SubItems(2)
    Quadra = Item.SubItems(3)
    Lote = Item.SubItems(4)
    
    sql = "SELECT count(*) FROM IMOVEIS " & _
            " WHERE " & _
                " Substring(Cast(InscImob AS varchar),1,2)=" & Distrito & " AND " & _
                " Substring(Cast(InscImob AS varchar),3,2)=" & Setor & " AND " & _
                " Substring(Cast(InscImob AS varchar),5,3)=" & Quadra & " AND " & _
                " Substring(Cast(InscImob AS varchar),8,4)=" & Lote & " AND "

    '1.
    Where = " (Compleme LIKE '%LOJA%' OR Compleme LIKE '%AP%' OR Compleme LIKE '%SALA%' OR Compleme LIKE '%KIT%' OR Compleme LIKE '%SL%')"
    If Bdados.AbreTabela(sql & Where, Rs) Then
        If Rs(0) > 1 Then
            SituacaoLote = Condominio
            Bdados.FechaTabela Rs
            Exit Function
        End If
    End If
    Bdados.FechaTabela Rs
    
    '2.
    Where = " AreConun>0"
    If Bdados.AbreTabela(sql & Where, Rs) Then
        If Rs(0) >= 1 Then
            SituacaoLote = Edificacao
            Bdados.FechaTabela Rs
            Exit Function
        End If
    End If
    Bdados.FechaTabela Rs
    
    '3.
    SituacaoLote = Terreno
End Function

Private Function ImprimirBT(Inscricao As String) As Boolean
    Dim Arquivo As String
    
    '1.
    Set Rpt = New VSRelatorio
    Arquivo = "TBT_PGV.rpt"
    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\" & Arquivo) Then Exit Function
    Rpt.Selecao = "{Imoveis.InscImob} = '" & Inscricao & "'"
    ImprimirBT = Rpt.Imprimir()
    '2.
    If ImprimirBT Then GravarImpressaoBT Inscricao
    Set Rpt = Nothing
End Function

Private Sub GravarImpressaoBT(Inscricao As String)
    Dim Campos As String, Valores As String
    
    Campos = "Setor,Quadra,Inscricao,Data,Usuario"
    Valores = Bdados.PreparaValor(Bdados.Converte(SetorImpressao, tctexto), Bdados.Converte(QuadraImpressao, tctexto), Bdados.Converte(Inscricao, tctexto), Date, Aplicacoes.Usuario)
    Bdados.GravaDados "BT", Valores, Campos, "Inscricao='" & Inscricao & "'"
End Sub

Private Sub ImprimirEdificacoes(ByRef Item As ListItem)
'ROTEIRO
'1. Imprimir um BT para o lote
'2. Imprimir um BP para cada grupo de 08 unidades
  
    '1.
    Item.Checked = ImprimirBT(Item.Text)

    '2.
    Dim sql As String, Rs As VSRecordset
  
    Dim Distrito As String
    Dim Setor As String
    Dim Quadra As String
    Dim Lote As String
    Dim Terreno As String
    Dim InscReduzida As String
    
    Terreno = Item.Text
    
    Distrito = Item.SubItems(1)
    Setor = Item.SubItems(2)
    Quadra = Item.SubItems(3)
    Lote = Item.SubItems(4)
    
    sql = "SELECT * FROM IMOVEIS " & _
            " WHERE " & _
                " Substring(Cast(InscImob AS varchar),1,2)=" & Distrito & " AND " & _
                " Substring(Cast(InscImob AS varchar),3,2)=" & Setor & " AND " & _
                " Substring(Cast(InscImob AS varchar),5,3)=" & Quadra & " AND " & _
                " Substring(Cast(InscImob AS varchar),8,4)=" & Lote & " AND " & _
                " AreConun>0" & _
            " ORDER BY Substring(Cast(InscImob AS varchar),1,2), " & _
                " Substring(Cast(InscImob AS varchar),3,2)," & _
                " Substring(Cast(InscImob AS varchar),5,3)," & _
                " Substring(Cast(InscImob AS varchar),8,4)," & _
                " Substring(Cast(InscImob AS varchar),12,3)"

    Dim i As Integer, pag As Integer
    Dim Imovel As ListItem
    Dim Arquivo As String
    
    i = 0: pag = 1
    Arquivo = "TBP_PGV.rpt"
    Set Rpt = New VSRelatorio
    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\" & Arquivo) Then Exit Sub
    If Bdados.AbreTabela(sql, Rs) Then
        Do While Not Rs.EOF
            Set Imovel = grdImoveis.FindItem(Rs!InscImob)
            If Not Imovel Is Nothing Then
                Imovel.Selected = True
                Imovel.Checked = True
                If Cancelar Then Exit Do
                i = i + 1
                If i > 8 Then
                    Imovel.Checked = ImprimirBP(Item.Text)
                    Rpt.LimparFormulas True
                    i = 1
                    pag = pag + 1
                End If
                InscReduzida = Format(Imovel.SubItems(5), "000") '& Format(Imovel.SubItems(7), "0000000")
                Rpt.Formulas "VT_INSCIMOB" & Format(i, "00"), InscReduzida
                Rpt.Formulas "VTAREAEDIF" & Format(i, "00"), Imovel.SubItems(7)
                GravarImpressaoBP Terreno, pag, Imovel.Text, InscReduzida
                Imovel.EnsureVisible
                DoEvents
            End If
            Rs.MoveNext
        Loop
        If i > 0 Then
            Imovel.Checked = ImprimirBP(Item.Text)
            If Imovel.Checked Then
                GravarImpressaoBP Terreno, pag, Imovel.Text, InscReduzida
            End If
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Function ImprimirBP(Inscricao As String) As Boolean
    Rpt.Selecao = "{Imoveis.InscImob} = '" & Inscricao & "'"
    ImprimirBP = Rpt.Imprimir()
End Function

Private Sub GravarImpressaoBP(Terreno As String, Pagina As Integer, Inscricao As String, Reduzida As String)
    Dim Campos As String, Valores As String
    
    Campos = "Setor, Quadra, Terreno, Pagina, Inscricao, InscricaoReduzida"
    Valores = Bdados.PreparaValor(Bdados.Converte(SetorImpressao, tctexto), Bdados.Converte(QuadraImpressao, tctexto), Bdados.Converte(Terreno, tctexto), Pagina, Bdados.Converte(Inscricao, tctexto), Bdados.Converte(Reduzida, tctexto))
    Bdados.GravaDados "BP", Valores, Campos, "Terreno='" & Terreno & "' AND Inscricao='" & Inscricao & "'"
End Sub

Private Sub ImprimirCondominio(ByRef Item As ListItem)
'ROTEIRO
'1. Imprimir um BT para o lote
'2. Imprimir um BP e um BC para unidade

    
    '1.
    Item.Checked = ImprimirBT(Item.Text)

    '2.
    Dim sql As String, Rs As VSRecordset
    Dim Where As String
    
    Dim Distrito As String
    Dim Setor As String
    Dim Quadra As String
    Dim Lote As String
    Dim Imovel As ListItem
    Dim Terreno As String
    
    Terreno = Item
    
    Distrito = Item.SubItems(1)
    Setor = Item.SubItems(2)
    Quadra = Item.SubItems(3)
    Lote = Item.SubItems(4)
    
    sql = "SELECT * FROM IMOVEIS " & _
            " WHERE " & _
                " Substring(Cast(InscImob AS varchar),1,2)=" & Distrito & " AND " & _
                " Substring(Cast(InscImob AS varchar),3,2)=" & Setor & " AND " & _
                " Substring(Cast(InscImob AS varchar),5,3)=" & Quadra & " AND " & _
                " Substring(Cast(InscImob AS varchar),8,4)=" & Lote & "  " & _
            " ORDER BY Substring(Cast(InscImob AS varchar),1,2), " & _
                " Substring(Cast(InscImob AS varchar),3,2)," & _
                " Substring(Cast(InscImob AS varchar),5,3)," & _
                " Substring(Cast(InscImob AS varchar),8,4)," & _
                " Substring(Cast(InscImob AS varchar),12,3)"
    Dim Arquivo As String
    If Bdados.AbreTabela(sql, Rs) Then
        Do While Not Rs.EOF
            Set Imovel = grdImoveis.FindItem(Rs!InscImob)
            If Not Imovel Is Nothing Then
                Imovel.Selected = True
                If Cancelar Then Exit Do
                '2.1
                Set Rpt = New VSRelatorio
                Arquivo = "TBP_PGV.rpt"
                If Not Rpt.DefinirArquivo(Bdados, App.Path & "\" & Arquivo) Then Exit Sub
                Rpt.Selecao = "{Imoveis.InscImob} = '" & Imovel.Text & "'"
                Rpt.LimparFormulas True
                Rpt.Formulas "VT_INSCIMOB01", Format(Imovel.SubItems(5), "000") & Format(Imovel.SubItems(7), "0000000")
                Imovel.Checked = Rpt.Imprimir()
                If Imovel.Checked Then
                    GravarImpressaoBP Terreno, 1, Imovel.Text, Format(Imovel.SubItems(5), "000") & Format(Imovel.SubItems(7), "0000000")
                End If
                Set Rpt = Nothing
                
                '2.2
                Set Rpt = New VSRelatorio
                Arquivo = "TBC_PGV.rpt"
                If Not Rpt.DefinirArquivo(Bdados, App.Path & "\" & Arquivo) Then Exit Sub
                Rpt.Selecao = "{Imoveis.InscImob} = '" & Imovel.Text & "'"
                Imovel.Checked = Rpt.Imprimir()
                If Imovel.Checked Then
                    GravarImpressaoBC Terreno, Imovel.Text
                End If
                Set Rpt = Nothing
                
                Imovel.EnsureVisible
                DoEvents
            End If

            Rs.MoveNext
        Loop
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub GravarImpressaoBC(Terreno As String, Inscricao As String)
    Dim Campos As String, Valores As String
    
    Campos = "Setor, Quadra, Terreno, Inscricao"
    Valores = Bdados.PreparaValor(Bdados.Converte(SetorImpressao, tctexto), Bdados.Converte(QuadraImpressao, tctexto), Bdados.Converte(Terreno, tctexto), Bdados.Converte(Inscricao, tctexto))
    Bdados.GravaDados "BC", Valores, Campos, "Setor='" & SetorImpressao & "' AND Quadra='" & QuadraImpressao & "' AND Terreno='" & Terreno & "' AND Inscricao='" & Inscricao & "'"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    grdImoveis.Preencher Bdados, ""
    txtSetor.SetFocus
    cboRelatorio.ListIndex = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    Screen.MousePointer = 0
    cboRelatorio.ListIndex = 0
End Sub

Private Sub fraImpressao_MudancaStatus()

End Sub

