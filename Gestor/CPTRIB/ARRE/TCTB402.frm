VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCTB402 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCTB402.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   705
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   690
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   1244
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
      Begin VB.TextBox txtNumLote 
         Alignment       =   1  'Right Justify
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
         Left            =   1980
         TabIndex        =   0
         Tag             =   "Lote"
         Top             =   240
         Width           =   1965
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   1
         Left            =   165
         TabIndex        =   19
         Top             =   270
         Width           =   1740
         _ExtentX        =   3069
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
         Caption         =   "Número do Lote"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   900
      Width           =   375
   End
   Begin Threed.SSFrame fra 
      Height          =   1095
      Index           =   2
      Left            =   0
      TabIndex        =   11
      Top             =   1470
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   1931
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
      Begin VB.ComboBox cboNumConta 
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
         Left            =   5310
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Conta"
         Top             =   630
         Width           =   1635
      End
      Begin VB.ComboBox cboCodSucursal 
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
         Left            =   1980
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Sucursal"
         Top             =   630
         Width           =   1965
      End
      Begin VB.ComboBox cboAgente 
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
         Left            =   1980
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Agente"
         Top             =   240
         Width           =   4965
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   4
         Left            =   480
         TabIndex        =   12
         Top             =   660
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   318
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
         Caption         =   "Código Sucursal"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   0
         Left            =   3630
         TabIndex        =   13
         Top             =   660
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   318
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
         Caption         =   "Num. Conta"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   1740
         _ExtentX        =   3069
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
         Caption         =   "Agente Arrecadador"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   510
      Index           =   1
      Left            =   30
      TabIndex        =   15
      Top             =   2595
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   900
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
      Begin VB.TextBox txtDtRecep 
         Alignment       =   1  'Right Justify
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
         Left            =   5310
         TabIndex        =   5
         Tag             =   "Data Recepção"
         Top             =   90
         Width           =   1605
      End
      Begin VB.TextBox txtDtArrecada 
         Alignment       =   1  'Right Justify
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
         Left            =   1950
         TabIndex        =   4
         Tag             =   "Data Arrecadação"
         Top             =   75
         Width           =   1605
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   105
         TabIndex        =   16
         Top             =   120
         Width           =   1770
         _ExtentX        =   3122
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
         Caption         =   "Data Arrecadação"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   255
         Index           =   6
         Left            =   3480
         TabIndex        =   17
         Top             =   120
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   450
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
         Caption         =   "Data Recepção"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin MSComctlLib.ListView lstLote 
      Height          =   2145
      Left            =   -15
      TabIndex        =   18
      Top             =   3630
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   3784
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
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
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2540
      EndProperty
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   1138
      Icone           =   "TCTB402.frx":2123
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   5940
      TabIndex        =   8
      Top             =   5850
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdBuscar 
      Height          =   375
      Left            =   5910
      TabIndex        =   6
      Top             =   3150
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Buscar"
      Acao            =   5
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdImprime 
      Height          =   375
      Left            =   4590
      TabIndex        =   7
      Top             =   5850
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      Caption         =   "&Imprimir"
      Acao            =   4
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCTB402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim NumAgente  As Double
Dim NumLote As Double
Private Sub cboAgente_Click()
    If cboAgente.ListIndex > 0 Then
        NumAgente = BuscaCodigo("Select tar_cod_agente from tab_agente_arrecadador where tar_nome_agente ='" & cboAgente & "'")
        AtualizaCombo Bdados, cboCodSucursal, "Select tcb_cod_sucursal from tab_conta_bancaria where tcb_tar_cod_agente =" & NumAgente
    Else
        cboCodSucursal.Clear
        cboNumConta.Clear
    End If
End Sub


Private Sub cboCodSucursal_Click()
    AtualizaCombo Bdados, cboNumConta, "Select tcb_num_conta from tab_conta_bancaria where tcb_tar_cod_agente =" & NumAgente & " and tcb_cod_sucursal ='" & cboCodSucursal & "'"
End Sub

Private Sub cmdBuscar_Click()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Condicao As String
    Sql = "Select TLP_COD_LOTE as Cod_Lote,tar_nome_agente as Agente,TLP_NUM_SUCURSAL as Sucursal,TLP_NUM_CONTA as Conta,TLP_DATA_ARRECADACAO as [Dt Arrecada]," & Bdados.Converte("TLP_VALOR_ARRECADADO", TCDuplo) & " as [Valor Lote(R$)],TLP_SITUACAO_LOTE AS Situacao from tab_lote_pagamento,tab_agente_arrecadador where tlp_tar_cod_agente = tar_cod_agente "
    
    If Trim(txtNumLote) <> "" Then
        Condicao = " and TLP_COD_LOTE=" & txtNumLote
    End If
    If Trim(cboAgente) <> "" Then
        Condicao = Condicao & " and tar_nome_agente='" & cboAgente & "'"
    End If
    If Trim(cboCodSucursal) <> "" Then
        Condicao = Condicao & " and TLP_NUM_SUCURSAL='" & cboCodSucursal & "'"
    End If
    If Trim(cboNumConta) <> "" Then
        Condicao = Condicao & " and TLP_NUM_CONTA='" & cboNumConta & "'"
    End If
    If Trim(txtDtArrecada) <> "" Then
        Condicao = Condicao & " and TLP_DATA_ARRECADACAO= " & Bdados.Converte(txtDtArrecada, TCDataHora)
    End If
    If Trim(txtDtRecep) <> "" Then
        Condicao = Condicao & " and TLP_DATA_RECEPCAO= " & Bdados.Converte(txtDtRecep, TCDataHora)
    End If
    Sql = Sql & Condicao
    MontaGrid Bdados, lstLote, Sql, 1400
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdNovo_Click()
    LimpaCampos Me
    txtNumLote.Enabled = True
    txtNumLote.SetFocus
End Sub

Private Sub cmdImprime_Click()
    On Error Resume Next
    Dim i As Integer
    If lstLote.ListItems.Count = 0 Then
        Avisa "Busca não retornou nenhum registro."
        Exit Sub
    End If
    Screen.MousePointer = 11
    If Rpt.DefinirArquivo(Bdados, App.Path + "\TCapaLote.rpt") Then
        For i = 1 To lstLote.ListItems.Count
            If lstLote.ListItems(i).ListSubItems(6).Text = 1 Then
                With Rpt
                    .SELECAO = "{TAB_LOTE_PAGAMENTO.TLP_COD_LOTE} = " & txtNumLote
                    .Formulas "municipio ", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
                    .Formulas "Lote ", lstLote.ListItems(i).Text
                    .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                    .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                    .Arvore = False
                    .Visualizar
                End With
            End If
        Next
        Avisa "Impressão concluída."
    End If
    Set Rpt = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Dim rs As VSRecordset
    cabVisual.Exibir Bdados, Me.Name, App.Path
    cboAgente.Clear
    cboCodSucursal.Clear
    cboNumConta.Clear
    AtualizaCombo Bdados, cboAgente, "Select tar_nome_agente from tab_agente_arrecadador where tar_ativo =0"
    cboAgente.AddItem ""
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub lstLote_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstLote, ColumnHeader
End Sub

Private Sub txtDtArrecada_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDtArrecada_LostFocus()
    txtDtArrecada = Edita.FormataTexto(txtDtArrecada, Data)
End Sub

Private Sub txtDtRecep_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDtRecep_LostFocus()
    txtDtRecep = Edita.FormataTexto(txtDtRecep, Data)
End Sub

Private Sub txtNumLote_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValorLote_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub
