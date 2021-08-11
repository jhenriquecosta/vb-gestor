VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCTA102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   60
      ScaleHeight     =   570
      ScaleWidth      =   555
      TabIndex        =   25
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCTA102.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   210
      Top             =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fluxo..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2385
      Left            =   45
      TabIndex        =   12
      Top             =   1845
      Width           =   9675
      Begin Threed.SSPanel lblTotal 
         Height          =   270
         Left            =   2820
         TabIndex        =   19
         Top             =   1500
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483637
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
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lblConta 
         Height          =   270
         Left            =   2820
         TabIndex        =   20
         Top             =   1140
         Width           =   6750
         _ExtentX        =   11906
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
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin MSComctlLib.ProgressBar BarraProgresso 
         Height          =   195
         Left            =   2790
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin Threed.SSPanel lblStatus 
         Height          =   240
         Left            =   2790
         TabIndex        =   22
         Top             =   600
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   423
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
         Caption         =   "0%"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   6
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   8400
         TabIndex        =   23
         Top             =   1920
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   6840
         TabIndex        =   24
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Movimentar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VB.Label LblTempoAtual 
         AutoSize        =   -1  'True
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1980
         TabIndex        =   18
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tempo Atual....:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   1170
         Width           =   1380
      End
      Begin VB.Label LblFim 
         AutoSize        =   -1  'True
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1980
         TabIndex        =   16
         Top             =   1530
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fim....:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1260
         TabIndex        =   15
         Top             =   1500
         Width           =   600
      End
      Begin VB.Label LblInicio 
         AutoSize        =   -1  'True
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1980
         TabIndex        =   14
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Inicio....:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   810
         Width           =   780
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   1155
      Index           =   2
      Left            =   30
      TabIndex        =   3
      Top             =   690
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   2037
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
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtIC 
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "Inscrição Municipal"
         Top             =   98
         Width           =   1485
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   3660
         TabIndex        =   8
         Tag             =   "Inscrição Municipal"
         Top             =   98
         Width           =   5895
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   3660
         TabIndex        =   6
         Tag             =   "Inscrição Municipal"
         Top             =   450
         Width           =   5895
      End
      Begin VB.TextBox txtIM 
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
         Left            =   1800
         MaxLength       =   13
         TabIndex        =   1
         Tag             =   "Inscrição Municipal"
         Top             =   450
         Width           =   1485
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   11
         Left            =   150
         TabIndex        =   4
         Top             =   495
         Width           =   1650
         _ExtentX        =   2910
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
         Caption         =   "Inscrição Municipal:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1650
         _ExtentX        =   2910
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
         Caption         =   "Inscrição Cadastral:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaIC 
         Height          =   315
         Left            =   3300
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   90
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaIM 
         Height          =   315
         Left            =   3300
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   450
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   195
      Left            =   3600
      TabIndex        =   5
      Top             =   4020
      Width           =   315
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1138
      Icone           =   "TCTA102.frx":2123
   End
   Begin MSComctlLib.ListView lstTrans 
      Height          =   330
      Left            =   180
      TabIndex        =   2
      Top             =   750
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   582
      View            =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "TCTA102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnExcutando As Boolean
Dim blnParar As Boolean

Private Sub cmd_Click(Index As Integer)
    Dim Conta As New ContaCorrente
    Dim Sql As String
    Dim rs As VSRecordset
    Dim RsAux As VSRecordset
    
    Dim Contador As Double
    Dim Obrig As New Obrigacao
    Dim Inscricao As String
    On Error GoTo trata
    Select Case Index
        Case 0
            
            If cmd(0).Caption = "Parar" Then
                blnParar = True
                cmd(0).Caption = "&Movimentar"
                Exit Sub
            End If
            
            Screen.MousePointer = 11
            'cmd(0).Enabled = False
            cmd(0).Caption = "Parar"
            lblStatus = "Iniciando..."
            'Bdados.Executa "EXEC SP_APAGA_OBRIGACAO_DUPLICADA"
            BarraProgresso.Visible = True
            blnExcutando = True
            LblInicio = Format(Time, "hh:mm:ss")
            Sql = "SELECT TOC_COD_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE"
            If Trim$(txtIC) <> "" Then
                    Sql = Sql & " WHERE TOC_INSCRICAO='" & txtIC & "'"
                    Inscricao = txtIC
            ElseIf Trim(txtIm) <> "" Then
                Sql = Sql & " WHERE TOC_INSCRICAO ='" & txtIm & "' OR TOC_INSCRICAO IN (SELECT TIM_IC FROM TAB_IMOVEL WHERE TIM_TCI_IM = '" & txtIm & "' )"
                Inscricao = txtIm
            End If
            Contador = 0
            DoEvents
            If Bdados.AbreTabela(Sql, rs, Estatico, SomenteLeitura) Then
                rs.MoveFirst
                If rs.RecordCount > 0 Then BarraProgresso.Max = rs.RecordCount
                lblConta = "Apagando Contas Selecionadas..."
                If txtIm <> "" Then
                    Conta.ApagaContasAnteriores Inscricao, etiContribuinte
                Else
                    Conta.ApagaContasAnteriores Inscricao, etiImovel
                End If
                DoEvents
                lblConta = "Pendenciando Pagamentos..."
                If txtIm <> "" Then
                    Conta.PendenciarPagamentos Inscricao, etiContribuinte
                Else
                    Conta.PendenciarPagamentos Inscricao, etiImovel
                End If
                lblStatus = "Executando  0%"
                Do While Not rs.EOF
                    If blnParar Then
                        Exit Do
                    End If
                    If txtIm <> "" Then
                        Obrig.BuscaDetalheObrigacao rs!TOC_COD_OBRIGACAO, etiContribuinte
                    Else
                        Obrig.BuscaDetalheObrigacao rs!TOC_COD_OBRIGACAO, etiImovel
                    End If
                    Dim modo As TipoInscricaoObrigacao
                    If txtIm <> "" Then
                        modo = etiContribuinte
                    Else
                        modo = etiImovel
                    End If
                    lblConta = "Criando Conta: " & Nvl("" & rs!TOC_COD_OBRIGACAO, 0)
                    If Conta.CriaContaContribuinte(Nvl("" & rs!TOC_COD_OBRIGACAO, 0), , , modo) Then
                        Contador = Contador + 1
                        DoEvents
                        lblConta = "Movimentando Conta: " & Nvl("" & rs!TOC_COD_OBRIGACAO, 0)
                        Conta.MovimentaContaContribuinte Nvl("" & rs!TOC_COD_OBRIGACAO, 0), , Obrig
                        Conta.BaixaPagamentos Obrig.obContribuinte, Obrig.obPeriodo, Obrig.obCodImposto, 0, "", "" & rs!TOC_COD_OBRIGACAO, Obrig, modo
                    End If
                    DoEvents
                    lblTotal = "Total de Contas processadas em " & Pega_Tempo & " - " & Contador
                    DoEvents
                    rs.MoveNext
                    If rs.EOF = False Then
                        BarraProgresso.Value = Nvl(rs.AbsolutePosition, 0)
                        lblStatus = "Executando  " & CInt((BarraProgresso.Value * 100) / BarraProgresso.Max) & "%"
                    End If
                    DoEvents
                Loop
                If blnParar Then
                    GoTo Sair
                End If
                lblConta = "Fim de Processo!"
            End If
            'PAGAMENTOS SEM OBRIGACAO
            Dim TotalPago As Double
            Sql = "SELECT TDR_TGT_COD_PAGAMENTO,TDR_INSCRICAO,TDR_PERIODO,TDR_TIP_COD_IMPOSTO," & _
                "tdr_valor_real_pago,tdr_data_vencimento FROM TAB_DARM_RECEBIDO"
            
            If Trim$(txtIm.Text) <> "" Then
                Sql = Sql & " INNER JOIN VIS_INSCRICAO ON TAB_DARM_RECEBIDO.TDR_INSCRICAO = VIS_INSCRICAO.VIN_INSCRICAO"
            End If
            
            Sql = Sql & " WHERE " & _
                "TDR_SIT_PAGO = 0 " 'AND " & _
                "(TDR_TGT_COD_PAGAMENTO_VINCULADO = 0  OR TDR_TGT_COD_PAGAMENTO_VINCULADO IS NULL)"
                
            If Trim$(txtIm.Text) <> "" Then
                Sql = Sql & " and VIN_INSCRICAO='" & Inscricao & "'"
            ElseIf Trim$(txtIC.Text) <> "" Then
                Sql = Sql & " and TDR_INSCRICAO='" & Inscricao & "'"
            End If
            If Bdados.AbreTabela(Sql, rs, Estatico, SomenteLeitura) Then
                rs.MoveFirst
                BarraProgresso.Max = rs.RecordCount
                lblStatus = "Executando  0%"
                Do While Not rs.EOF
                    If blnParar Then
                        Exit Do
                    End If
                    Sql = "Select sum(tdr_valor_real_pago) as Total from tab_darm_recebido where " & _
                    "tdr_tgt_cod_pagamento =" & rs!TDR_TGT_COD_PAGAMENTO & " or TDR_TGT_COD_PAGAMENTO_VINCULADO =" & rs!TDR_TGT_COD_PAGAMENTO
                    If Bdados.AbreTabela(Sql, RsAux) Then
                        TotalPago = Nvl("" & RsAux!Total, 0)
                    End If
                    lblConta = "Criando Obrigacao: " & Nvl("" & rs!TDR_TGT_COD_PAGAMENTO, 0)
                    Obrig.CriaObrigacao "" & rs!tdr_tip_cod_imposto, "" & rs!tdr_periodo, rs!tdr_periodo, Trim("" & rs!tdr_inscricao), TotalPago, etsCreditoOriginalAberto, True, "" & rs!tdr_data_vencimento, , rs!TDR_TGT_COD_PAGAMENTO
                    lblConta = "Criando Conta: " & Nvl("" & rs!TDR_TGT_COD_PAGAMENTO, 0)
                    If Conta.CriaContaContribuinte(Nvl("" & rs!TDR_TGT_COD_PAGAMENTO, 0), , Obrig, modo) Then
                        Contador = Contador + 1
                        DoEvents
                        lblConta = "Movimentando Conta: " & Nvl("" & rs!TDR_TGT_COD_PAGAMENTO, 0)
                        Conta.MovimentaContaContribuinte Nvl("" & rs!TDR_TGT_COD_PAGAMENTO, 0), , Obrig
                        Conta.BaixaPagamentos Trim(rs!tdr_inscricao), rs!tdr_periodo, rs!tdr_tip_cod_imposto, 0, "", Nvl("" & rs!TDR_TGT_COD_PAGAMENTO, 0), Obrig
                    End If
                    DoEvents
                    lblTotal = "Total de Contas processadas em " & Pega_Tempo & " - " & Contador
                    DoEvents
                    rs.MoveNext
                    If rs.EOF = False Then
                        BarraProgresso.Value = rs.AbsolutePosition
                        lblStatus = "Executando  " & CInt((BarraProgresso.Value * 100) / BarraProgresso.Max) & "%"
                    End If
                    DoEvents
                Loop
                If blnParar Then
                    GoTo Sair
                End If
            End If
            '********
                        
Sair:
            Bdados.FechaTabela rs
            blnExcutando = False
            Screen.MousePointer = 0
            If blnParar = False Then
                lblStatus = "Concluído"
                lblConta.Caption = ""
                Informa "Movimentações de contas finalizadas às :" & Time & "."
                BarraProgresso.Visible = False
            Else
                'Util.Informa "Operação Cancelada"
                lblStatus = "Parado"
                blnParar = False
            End If
            cmd(0).Caption = "&Movimentar"
            'cmd(0).Enabled = True
            Exit Sub
        Case 1
            If blnExcutando Then
                If Util.Confirma("Deseja sair e cancelar o processo de movimentação") Then
                    blnParar = True
                    DoEvents
                Else
                    Exit Sub
                End If
            End If
            Unload Me
    End Select
    Exit Sub
trata:
    Avisa Err.Description
    Resume Next
    Exit Sub
    Resume
End Sub
Private Function Pega_Tempo()
    On Error Resume Next
    Dim hora
    Dim Minu
    Dim Segu
    
    hora = DateDiff("h", LblTempoAtual, LblInicio)
    Minu = DateDiff("m", LblTempoAtual, LblInicio)
    Segu = DateDiff("s", LblTempoAtual, LblInicio)
    
    Pega_Tempo = TimeSerial(Format(hora, "00"), Format(Minu, "00"), Format(Segu, "00"))
End Function
Private Sub cmdPesquisaIC_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtIC
End Sub

Private Sub cmdPesquisaIM_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub Command1_Click()
    SendKeys "{tab}"
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub lstTrans_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstTrans, ColumnHeader
End Sub

Private Sub Timer1_Timer()
    LblTempoAtual = Format(Time, "hh:mm:ss")
End Sub

Private Sub txtIC_Change()
    txtEndereco.Text = ""
End Sub

Private Sub txtIC_LostFocus()
    Dim Sql As String
    
    If Trim$(txtIC) <> "" Then
        'Sql = "SELECT * FROM VIS_ENDERECO_IMOVEL WHERE tim_ic='" & txtIC & "'"
        Sql = "SELECT * FROM VIS_IMOVEL WHERE tim_ic='" & txtIC & "'"
        If Bdados.AbreTabela(Sql) Then
            txtEndereco = Bdados.Tabela!TTL_NOME & " " & Bdados.Tabela!tlg_nome & ", " & Bdados.Tabela!tim_numero & " " & Bdados.Tabela!tim_complemento & " - " & Bdados.Tabela!TBA_NOME
       
            cmd(0).SetFocus
        End If
        Bdados.FechaTabela
    End If
End Sub

Private Sub txtIm_Change()
    txtNome.Text = ""
End Sub

Private Sub txtIm_LostFocus()
    Dim Imposto As New VSImposto
    Dim Sql As String
    Dim rs As VSRecordset
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIm = Imposto.FormataInscricao(txtIm, InscContrib)
    End If
    Sql = "select tci_nome from tab_contribuinte where tci_im = '" & txtIm & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtNome = "" & rs!tci_nome
    Else
        If Trim(txtIm) = "" Then Exit Sub
        Informa "Contribuinte não cadastrado."
        txtIm.SetFocus
        Exit Sub
    End If
End Sub
