VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCTB201 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   555
      TabIndex        =   22
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCTB201.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   705
      Index           =   0
      Left            =   45
      TabIndex        =   10
      Top             =   660
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
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "NÚMERO DO LOTE"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   960
      Width           =   375
   End
   Begin Threed.SSFrame fra 
      Height          =   1095
      Index           =   2
      Left            =   45
      TabIndex        =   13
      Top             =   1500
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
         TabIndex        =   3
         Tag             =   "Conta"
         Text            =   "cboNumConta"
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
         TabIndex        =   2
         Tag             =   "Sucursal"
         Text            =   "cboCodSucursal"
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
      Height          =   1125
      Index           =   1
      Left            =   45
      TabIndex        =   17
      Top             =   2745
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   1984
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
      Begin VB.TextBox txtValorLote 
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
         TabIndex        =   6
         Tag             =   "Valor Lote"
         Top             =   630
         Width           =   1605
      End
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
         Top             =   240
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
         Top             =   240
         Width           =   1605
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   240
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
         Height          =   270
         Index           =   6
         Left            =   3480
         TabIndex        =   19
         Top             =   240
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
         Caption         =   "Data Recepção"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   630
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
         Caption         =   "Valor do Lote(R$)"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   1138
      Icone           =   "TCTB201.frx":2123
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   5970
      TabIndex        =   9
      Top             =   3960
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   375
      Left            =   4740
      TabIndex        =   7
      Top             =   3960
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdNovo 
      Height          =   375
      Left            =   3510
      TabIndex        =   8
      Top             =   3960
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCTB201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim NumAgente  As Double
Dim NumLote As Double
Private Sub cboAgente_Click()
    If cboAgente.ListIndex > -1 Then
        NumAgente = BuscaCodigo("Select tar_cod_agente from tab_agente_arrecadador where tar_nome_agente ='" & cboAgente & "'")
        AtualizaCombo Bdados, cboCodSucursal, "Select DISTINCT(tcb_cod_sucursal) from tab_conta_bancaria where tcb_tar_cod_agente =" & NumAgente
    End If
End Sub


Private Sub cboCodSucursal_Click()
    AtualizaCombo Bdados, cboNumConta, "Select tcb_num_conta from tab_conta_bancaria where tcb_tar_cod_agente =" & NumAgente & " and tcb_cod_sucursal ='" & cboCodSucursal & "'"
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdNovo_Click()
    LimpaCampos Me
    txtNumLote.Enabled = True
    txtNumLote.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Condicao As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
    Campos = "TLP_TAR_COD_AGENTE,TLP_NUM_SUCURSAL,TLP_NUM_CONTA,TLP_VALOR_ARRECADADO,TLP_DATA_ABERTURA,TLP_DATA_ARRECADACAO,TLP_DATA_RECEPCAO,TLP_SITUACAO_LOTE,TLP_TUS_COD_USUARIO"
    Valores = Bdados.PreparaValor(NumAgente, cboCodSucursal, cboNumConta, Bdados.Converte(txtValorLote, TCDuplo), Bdados.Converte(Date, TCDataHora), Bdados.Converte(txtDtArrecada, TCDataHora), Bdados.Converte(txtDtRecep, TCDataHora), 0, Aplicacoes.Usuario)
    Bdados.AtualizaDados "TAB_LOTE_PAGAMENTO", Valores, Campos, "TLP_COD_LOTE=" & txtNumLote
    Bdados.AtualizaDados "TAB_DARM_RECEBIDO", Bdados.PreparaValor(Bdados.Converte(txtDtArrecada, TCDataHora)), "tdr_data_pagamento", "TDR_TLP_COD_LOTE=" & txtNumLote
    Util.Informa "Transação Realizada com Sucesso."
    Edita.LimpaCampos Me
    txtNumLote.Enabled = True
    txtNumLote.SetFocus
End Sub

Private Sub Form_Load()
    Dim rs As VSRecordset
    cabVisual.Exibir Bdados, Me.Name, App.Path
    cboAgente.Clear
    cboCodSucursal.Clear
    cboNumConta.Clear
    AtualizaCombo Bdados, cboAgente, "Select tar_nome_agente from tab_agente_arrecadador where tar_ativo =0"
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
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

Private Sub txtNumLote_LostFocus()
    Dim rs As VSRecordset
    On Error Resume Next
    If Trim(txtNumLote) = "" Then Exit Sub
    If Bdados.AbreTabela("Select   * from TAB_LOTE_PAGAMENTO where TLP_COD_LOTE=" & Trim(txtNumLote), rs) Then
        cboAgente.ListIndex = BuscaIndiceCombo(cboAgente, "Tab_Agente_Arrecadador", "tar_cod_agente", "tar_nome_agente", rs!TLP_TAR_COD_AGENTE)
        cboAgente_Click
        cboCodSucursal = "" & rs!TLP_NUM_SUCURSAL
        cboCodSucursal_Click
        cboNumConta = "" & rs!TLP_NUM_CONTA
        txtDtArrecada = "" & rs!TLP_DATA_ARRECADACAO
        txtDtRecep = "" & rs!TLP_DATA_RECEPCAO
        txtValorLote = Format("" & rs!TLP_VALOR_ARRECADADO, Const_Monetario)
    Else
        Avisa "Lote inexistente."
        txtNumLote.Enabled = True
        txtNumLote.SetFocus
        Exit Sub
    End If
    txtNumLote = Format(txtNumLote, "0000000000000")
    txtNumLote.Enabled = False
    Bdados.FechaTabela rs
End Sub

Private Sub txtValorLote_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValorLote_LostFocus()
    txtValorLote = Edita.FormataTexto(txtValorLote, Monetario, True)
End Sub
