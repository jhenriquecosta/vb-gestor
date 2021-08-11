VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TARR402 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   60
      ScaleHeight     =   570
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TARR402.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2850
      Top             =   5820
   End
   Begin Threed.SSFrame fra 
      Height          =   3330
      Index           =   2
      Left            =   15
      TabIndex        =   11
      Top             =   660
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   5874
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Consulta"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VTOcx.txtVISUAL txtTop 
         Height          =   285
         Left            =   9330
         TabIndex        =   19
         Top             =   2265
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   503
         Caption         =   "Top"
         Text            =   ""
      End
      Begin VB.TextBox txtExercicio2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4680
         MaxLength       =   12
         TabIndex        =   4
         Top             =   1049
         Width           =   1485
      End
      Begin VB.TextBox txtExercicio1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3180
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1049
         Width           =   1485
      End
      Begin VB.ComboBox cboImposto 
         DataField       =   "ttl_nome"
         DataSource      =   "dtTipLogr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1830
         Width           =   7755
      End
      Begin VB.ComboBox cboRelatorio 
         DataField       =   "ttl_nome"
         DataSource      =   "dtTipLogr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo de Relatório"
         Top             =   1431
         Width           =   7755
      End
      Begin VB.TextBox txtPeriodo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3180
         MaxLength       =   12
         TabIndex        =   1
         Top             =   667
         Width           =   1485
      End
      Begin VB.TextBox txtPeriodo2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4680
         MaxLength       =   12
         TabIndex        =   2
         Top             =   667
         Width           =   1485
      End
      Begin VB.ComboBox cboAgente 
         DataField       =   "ttl_nome"
         DataSource      =   "dtTipLogr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   7755
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   11
         Left            =   1380
         TabIndex        =   12
         Top             =   300
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Agente Arrecadador:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   689
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Período de Pagamento(dd/mm/aaaa):"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   1380
         TabIndex        =   14
         Top             =   1461
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tipo de Relatório:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   1
         Left            =   1380
         TabIndex        =   15
         Top             =   1860
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tributo:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   90
         TabIndex        =   16
         Top             =   1071
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Exercício:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   9810
         TabIndex        =   9
         Top             =   2760
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Sair"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   7380
         TabIndex        =   7
         Top             =   2760
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   8610
         TabIndex        =   8
         Top             =   2760
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   900
      TabIndex        =   10
      Top             =   1050
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   1138
      Icone           =   "TARR402.frx":2123
   End
End
Attribute VB_Name = "TARR402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As New VSImposto
Dim Sql As String
Dim Filtro As String

Private Sub cboRelatorio_Click()
    If Trim(Right(Trim(cboRelatorio), 3)) = 99 Then
        txtTop.Enabled = True
        txtTop = 100
    Else
        txtTop = ""
        txtTop.Enabled = False
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    On Error Resume Next
    Dim rs As VSRecordset
    Dim Condicao As String
    Static Imposto As String
    Dim i As Byte
    Dim NomeImposto As String
    Dim Exercicio1 As Double
    Dim Exercicio2 As Double
    Dim Path As String
    Dim SubCondicao As String
    
    
    Select Case Index
        Case 0
            Unload Me
        Case 1
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Screen.MousePointer = 11
            
            Select Case Trim(Right(Trim(cboRelatorio), 3))
                Case 1
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TResumoArrecGeral.rpt") Then Exit Sub
                Case 2
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TResumoArrecBanco.rpt") Then Exit Sub
                Case 3
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TResumoArrecImpostoNovo.rpt") Then Exit Sub
                Case 4
                    'If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TDamPago.rpt") Then Exit Sub
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TResumoArrecContribuinte.rpt") Then Exit Sub
                Case 5
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TRPT4015.rpt") Then Exit Sub
                Case 6
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TComparaArrec.rpt") Then Exit Sub
                Case 7
'                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TResumoArrecImpostoAgente.rpt") Then Exit Sub
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TResumoArrecImpostoNovo.rpt") Then Exit Sub
                Case 8
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TGraficoArrecGeral2.rpt") Then Exit Sub
                Case 9
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TResumoContabil.rpt") Then Exit Sub
                Case 10
                    If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TResumoContabilBanco.rpt") Then Exit Sub
                
                Case Else
                    If Trim(Right(Trim(cboRelatorio), 3)) = 99 Then GoTo Top
                    Util.Informa "Tipo de relatorio inválido."
                    cboRelatorio.SetFocus
                    Screen.MousePointer = 0
                    Exit Sub
            End Select
            If Trim(Right(Trim(cboRelatorio), 3)) <> 99 Then
                    With Rpt
                        Condicao = "{Tab_Darm_Recebido.tdr_sit_pago} <> 2 " 'AND {Tab_Darm_Recebido.tdr_tgt_cod_pagamento_vinculado} ={Tab_Darm_Recebido.tdr_tgt_cod_pagamento} "
                        If Trim(CboImposto) <> "" Then
                            i = InStr(1, CboImposto, "#")
                            NomeImposto = Left(CboImposto.Text, i - 2)
                            Condicao = Condicao & " and {Tab_Imposto.tip_sigla_imposto} ='" & NomeImposto & "'"
                        End If
                        If Trim(Right(Trim(cboRelatorio), 3)) <> 4 Then
                            If Trim(cboAgente) <> "" Then Condicao = Condicao & " and {Tab_Agente_Arrecadador.tar_nome_agente} ='" & cboAgente & "'"
                        End If
                        If Trim(txtPeriodo1) <> "" And Trim(txtPeriodo2) <> "" Then
                            If Not IsDate(txtPeriodo1) Or Not IsDate(txtPeriodo2) Then
                                Avisa "Data inválida."
                                Screen.MousePointer = 0
                                Exit Sub
                            End If
                            
                            Condicao = Condicao & " and  {TAB_LOTE_PAGAMENTO.TLP_DATA_ARRECADACAO} in  " & _
                            "Date (" & Year(txtPeriodo1) & "," & Month(txtPeriodo1) & "," & Day(txtPeriodo1) & ") to Date " & _
                            "(" & Year(txtPeriodo2) & "," & Month(txtPeriodo2) & "," & Day(txtPeriodo2) & ")"
                        End If
                        
                        If Trim(txtExercicio1) <> "" And Trim(txtExercicio2) <> "" Then
                            Exercicio1 = IIf(Len(Trim(txtExercicio1)) = 4, txtExercicio1, Right(Trim(txtExercicio1), 4) & Left(Trim(txtExercicio1), 2))
                            Exercicio2 = IIf(Len(Trim(txtExercicio2)) = 4, txtExercicio2, Right(Trim(txtExercicio2), 4) & Left(Trim(txtExercicio2), 2))
                            Condicao = Condicao & " and  {TAB_DARM_RECEBIDO.tdr_periodo} >=" & Exercicio1 & _
                            " and  {TAB_DARM_RECEBIDO.tdr_periodo} <=" & Exercicio2
                        End If
                        If Trim(Right(Trim(cboRelatorio), 3)) <> 4 And Trim(Right(Trim(cboRelatorio), 3)) <> 7 Then
                            .Formulas "FILTRO ", IIf(Trim(cboAgente) <> "", "AGENTE : " & cboAgente & "    -     ", IIf(Trim(cboAgente) <> "", " - ", "")) & IIf(Trim(txtPeriodo1) <> "", "PERÍODO : " & txtPeriodo1 & " A " & txtPeriodo2, "")
                        End If
                        If Trim(Right(Trim(cboRelatorio), 3)) = 1 Then
                           .SubRelatorio = "TResumoContabilBanco.rpt"
                           SubCondicao = " {Tab_Darm_Recebido.tdr_sit_pago} <> 2 "
                           'If cboAgente.ListIndex >= 0 Then
                           ' SubCondicao = SubCondicao & " and {Tab_Agente_Arrecadador.tar_nome_agente} = '" & cboAgente.Text & "'"
                           'End If
                           .SELECAO = Condicao
                           .SubRelatorio = ""
                        End If
                        If Trim(Right(Trim(cboRelatorio), 3)) = 5 Then
                            If txtExercicio1 <> "" And txtExercicio2 <> "" Then
                                .Formulas "VTPERIODO", "PERIODO DE " & txtExercicio1 & " A " & txtExercicio2
                            ElseIf txtExercicio1 <> "" And txtExercicio2 = "" Then
                                .Formulas "VTPERIODO", "PERIODO DE " & txtExercicio1
                            End If
                        End If
                        
                        
                        
                        
                        .SELECAO = Condicao
                        'If CDbl(Trim(Right(Trim(cboRelatorio), 3))) = 1 Then
                        '    .SubRelatorio = "TResumoContabilBanco.rpt"
                        '    .SELECAO = Condicao
                        'End If
 '                       .SubRelatorio = ""
                        
                        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                        .Arvore = False
                        .Visualizar
                    End With
            Else
Top:
                    'Drop para Criar novamente...
                    Bdados.Executa "Drop view dbo.VIS_MAIORES_CONTRIBUINTES"
                    Sql = " CREATE VIEW dbo.VIS_MAIORES_CONTRIBUINTES"
                    Sql = Sql & " AS"
                    Sql = Sql & " SELECT     TOP " & txtTop & " dbo.VIS_ARREC_CONTRIBUINTE.TDR_INSCRICAO, dbo.TAB_CONTRIBUINTE.tci_nome, dbo.TAB_CONTRIBUINTE.tci_fantasia,"
                    Sql = Sql & " dbo.TAB_CONTRIBUINTE.tci_cgc_cpf, dbo.TAB_CONTRIBUINTE.TCI_FONE_FAX, SUM(dbo.VIS_ARREC_CONTRIBUINTE.TDR_TOTAL) AS TOTAL,"
                    Sql = Sql & " tci_logradouro ,"
                    Sql = Sql & " tci_nome_logradouro ,"
                    Sql = Sql & " tci_numero ,"
                    Sql = Sql & " tci_complemento ,"
                    Sql = Sql & " tci_bairro"
                    Sql = Sql & " FROM         dbo.VIS_ARREC_CONTRIBUINTE INNER JOIN"
                    Sql = Sql & " dbo.TAB_CONTRIBUINTE ON dbo.VIS_ARREC_CONTRIBUINTE.TDR_INSCRICAO = dbo.TAB_CONTRIBUINTE.tci_im"
                    Sql = Sql & " Where (dbo.VIS_ARREC_CONTRIBUINTE.tdr_tipo_inscricao = 2)"
                    Sql = Sql & " GROUP BY dbo.VIS_ARREC_CONTRIBUINTE.TDR_INSCRICAO, dbo.TAB_CONTRIBUINTE.tci_nome, dbo.TAB_CONTRIBUINTE.tci_cgc_cpf,"
                    Sql = Sql & " dbo.TAB_CONTRIBUINTE.TCI_FONE_FAX , dbo.TAB_CONTRIBUINTE.tci_fantasia,"
                    Sql = Sql & " tci_logradouro ,"
                    Sql = Sql & " tci_nome_logradouro ,"
                    Sql = Sql & " tci_numero ,"
                    Sql = Sql & " tci_complemento ,"
                    Sql = Sql & " tci_bairro"
                    Sql = Sql & " ORDER BY TOTAL desc"
                    If Bdados.Executa(Sql) Then
                        If Not Rpt.DefinirArquivo(Bdados, App.Path & "\CemMaioresContribuintes.rpt") Then Exit Sub
                        Rpt.Formulas "LABAL", "RELATÓRIO DOS " & txtTop & " MAIORES CONTRIBUINTES"
                        Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                        Rpt.Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                        Rpt.Visualizar
                    End If
            End If
            Set Rpt = Nothing
            Screen.MousePointer = 0
        Case 2
            LimpaCampos Me
            cboAgente.SetFocus
    End Select
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    AtualizaCombo Bdados, cboAgente, "Select tar_nome_agente from tab_agente_arrecadador where tar_ativo =0"
    cboAgente.AddItem " "
    'ederson, substituir o "Bdados.concatena & " SPACE(150) " & Bdados.concatena" pois não funciona no sql
    Call Edita.AtualizaCombo(Bdados, cboRelatorio, "SELECT TGE_NOME " & Bdados.Concatena & " SPACE(150) " & Bdados.Concatena & " CAST(TGE_CODIGO AS VARCHAR) FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =714 ORDER BY TGE_CODIGO ASC")
    cboRelatorio.AddItem " "
    Call Edita.AtualizaCombo(Bdados, CboImposto, "Select  TIP_sigla_IMPOSTO " & Bdados.Concatena & " ' # ' " & Bdados.Concatena & " tip_nome_imposto From TAB_IMPOSTO")
    CboImposto.AddItem " "
End Sub

Private Sub Timer1_Timer()
    FocalizaCaixa Me
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtExercicio1_KeyPress(KeyAscii As Integer)
     KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtExercicio2_KeyPress(KeyAscii As Integer)
     KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPeriodo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPeriodo1_LostFocus()
    If IsNumeric(txtPeriodo1) Then
        txtPeriodo1 = Edita.FormataTexto(txtPeriodo1, Data)
    End If
End Sub

Private Sub txtPeriodo2_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPeriodo2_LostFocus()
    If IsNumeric(txtPeriodo2) Then
        txtPeriodo2 = Edita.FormataTexto(txtPeriodo2, Data)
    End If
End Sub
