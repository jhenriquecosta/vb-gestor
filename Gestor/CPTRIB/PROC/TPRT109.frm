VERSION 5.00
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TPRT109 
   Caption         =   "ATUALIZAÇÃO DE PROCESSO"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPrazo 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VTOcx.cboVISUAL cboTIPO 
      Height          =   510
      Left            =   120
      TabIndex        =   0
      Tag             =   "C"
      Top             =   1080
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   900
      Caption         =   "Etapa  "
      Text            =   ""
      AutoFocaliza    =   0   'False
      Requerido       =   0   'False
      Alinhamento     =   1
   End
   Begin VTOcx.cmdVISUAL cmdImprimir 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   4440
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Atualizar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdCancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Sair"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.txtVISUAL txtInicio 
      Height          =   480
      Left            =   120
      TabIndex        =   2
      Tag             =   "A"
      Top             =   3840
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   847
      Caption         =   "Data Ciência (Início)"
      Text            =   ""
      TipoLetras      =   0
      Formato         =   0
      AlinhamentoRotulo=   1
   End
   Begin VTOcx.txtVISUAL txtDias 
      Height          =   480
      Left            =   2160
      TabIndex        =   3
      Tag             =   "A"
      Top             =   3840
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   847
      Caption         =   "N. Dias (Prazo)"
      Text            =   ""
      TipoLetras      =   0
      AlinhamentoRotulo=   1
   End
   Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.txtVISUAL txtIm 
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   ""
      Text            =   ""
      Restricao       =   2
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtNOME 
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Tag             =   "C"
      Top             =   720
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   529
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
      Requerido       =   0   'False
   End
   Begin VTOcx.cboVISUAL cboStatus 
      Height          =   510
      Left            =   120
      TabIndex        =   11
      Tag             =   "C"
      Top             =   1680
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   900
      Caption         =   "Status"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Requerido       =   0   'False
      Alinhamento     =   1
   End
   Begin VTOcx.cboVISUAL cboFiscal 
      Height          =   510
      Left            =   120
      TabIndex        =   12
      Tag             =   "C"
      Top             =   2280
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   900
      Caption         =   "Fiscal 1"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Requerido       =   0   'False
      Alinhamento     =   1
   End
   Begin VTOcx.cboVISUAL cboFiscal2 
      Height          =   510
      Left            =   120
      TabIndex        =   13
      Tag             =   "C"
      Top             =   2880
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   900
      Caption         =   "Fiscal 2"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Requerido       =   0   'False
      Alinhamento     =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Inscrição"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Prazo para o processo!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
   End
End
Attribute VB_Name = "TPRT109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private r As VSRelatorio
Private Processo As Long
Dim os As New OrdemServico

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
     If Len(cboTIPO.Text) = 0 Then
        Mensagem "Informe a proxima etapa do processo"
        Exit Sub
     End If
     If Len(cboStatus.Text) = 0 Then
        Mensagem "Informe o status do processo"
        Exit Sub
     End If
     If Len(txtIm) > 0 Then
        If os.inserirInscricao(Processo, txtIm) Then
        End If
        
    End If
     If os.atualizaProcesso(Processo, CInt(cboTIPO.Coluna(1).valor), CDate(txtInicio), CInt(txtDias), chkPrazo.Value, cboStatus.Text, cboFiscal.Text, cboFiscal2.Text) Then
        'r.visualizar
     End If
     Unload Me
End Sub
Public Sub carregar(codOs As Long)
    Processo = codOs
    os.PreencheCombo cboTIPO
    cboStatus.AddItem "ABERTA"
    cboStatus.AddItem "FISCALIZAÇÃO"
    cboStatus.AddItem "COMPARECIMENTO"
    cboStatus.AddItem "ATENDIMENTO"
    cboStatus.AddItem "EXECUÇÃO"
    cboStatus.AddItem "FINALIZADA"
    cboStatus.AddItem "RE-ABERTA"
    Dim rs As VSRecordset
    If Bdados.AbreTabela("SELECT * FROM VIS_BCP_ORDEM_SERVICO WHERE SERVICO=" & Processo, rs) Then
            cboFiscal = rs("FISCAL")
            cboFiscal2 = rs("FISCAL2")
            cboTIPO = rs("ETAPA")
            cboStatus = rs("SITUACAO")
            If Not (IsNull(rs("INSCRICAO"))) And Len(Trim(rs("INSCRICAO"))) > 0 Then
                txtIm = rs("INSCRICAO")
                txtNOME = rs("RAZAO")
            End If
    End If
    txtInicio = Format(Now, "DD/MM/YYYY")
    txtDias = 0
    Me.Show vbModal
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub Form_Load()
    Dim rs As VSRecordset
    If Bdados.AbreTabela("SELECT TUS_COD_USUARIO FROM TAB_USUARIO ORDER BY TUS_COD_USUARIO", rs) Then
        Do While Not rs.EOF
            cboFiscal.AddItem rs(0)
            cboFiscal2.AddItem rs(0)
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub txtIm_LostFocus()
     Dim rs As VSRecordset
    Dim Sql As String
    'lblOrdem = ""
    If Len(txtIm) > 0 Or txtIm <> "-" Then
        Sql = " Select * from Tab_Contribuinte " _
            & " where tci_im='" & txtIm & "'"
        'If Not Conexao Is Nothing Then Set Bdados = Conexao
        If Bdados.AbreTabela(Sql, rs) Then
            'txtAtividade = Imposto.BuscaNomeCAE("" & rs("tci_tae_cae"))
            txtNOME = "" & rs("tci_nome")
            'txtEndereco = "" & rs("tci_logradouro") & " " & rs("tci_nome_logradouro") & "," & rs("tci_numero") & " " & rs("tci_complemento") & " " & rs("tci_bairro")
            
        Else
            'txtAtividade = ""
            txtNOME = ""
           'txtEndereco = ""
        End If
    Else
            'txtAtividade = ""
            txtNOME = ""
            'txtEndereco = ""
    End If
    
End Sub
