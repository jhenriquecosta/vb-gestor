VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCIS106 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIS106"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   ControlBox      =   0   'False
   Icon            =   "TCIS106.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   5400
      Left            =   30
      TabIndex        =   7
      Top             =   660
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   9525
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      TagVariant      =   ""
      Tabs            =   "TCIS106.frx":08CA
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5010
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8837
         _Version        =   131082
         TabGuid         =   "TCIS106.frx":0948
         Begin VB.Frame Frame1 
            Caption         =   "Valores"
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
            Height          =   1305
            Left            =   6930
            TabIndex        =   10
            Top             =   3600
            Width           =   2865
            Begin VTOcx.txtVISUAL txtFatorMutiplicador 
               Height          =   300
               Left            =   360
               TabIndex        =   11
               Top             =   570
               Width           =   2220
               _ExtentX        =   3916
               _ExtentY        =   529
               Caption         =   "Mutiplicador"
               Text            =   ""
               Restricao       =   2
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtValorApagar 
               Height          =   300
               Left            =   210
               TabIndex        =   12
               Top             =   930
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   529
               Caption         =   "Valor a Pagar"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtValor 
               Height          =   300
               Left            =   930
               TabIndex        =   13
               Top             =   210
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   529
               Caption         =   "Valor"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoTexto=   1
            End
         End
         Begin VTOcx.fraVISUAL fraAnu 
            CausesValidation=   0   'False
            Height          =   1860
            Left            =   135
            TabIndex        =   14
            Top             =   150
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   3281
            Altura          =   1905
            Caption         =   " Dados do Veículo de Divulgacão"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtDataInstalacao 
               Height          =   300
               Left            =   6690
               TabIndex        =   23
               Top             =   1095
               Width           =   2790
               _ExtentX        =   4921
               _ExtentY        =   529
               Caption         =   "Data Instalação"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoTexto=   1
            End
            Begin VTOcx.cboVISUAL cboMovimento 
               Height          =   315
               Left            =   780
               TabIndex        =   22
               Top             =   735
               Width           =   8730
               _ExtentX        =   15399
               _ExtentY        =   556
               Caption         =   "Movimento"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtArea 
               Height          =   300
               Left            =   840
               TabIndex        =   21
               Top             =   1440
               Width           =   2085
               _ExtentX        =   3678
               _ExtentY        =   529
               Caption         =   "Área Total"
               Text            =   ""
               Restricao       =   2
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtDimensao 
               Height          =   300
               Left            =   885
               TabIndex        =   20
               Top             =   1095
               Width           =   5775
               _ExtentX        =   10186
               _ExtentY        =   529
               Caption         =   "Descrição"
               Text            =   ""
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtPeriodo 
               Height          =   300
               Left            =   7410
               TabIndex        =   19
               Top             =   1440
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   529
               Caption         =   "Período"
               Text            =   ""
               Restricao       =   3
               AlinhamentoTexto=   1
            End
            Begin VTOcx.cboVISUAL CboStatus 
               Height          =   315
               Left            =   4410
               TabIndex        =   18
               Top             =   1440
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
               Height          =   285
               Left            =   3330
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   390
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               Caption         =   ""
               Acao            =   5
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.txtVISUAL txtContribuinte 
               Height          =   285
               Left            =   3690
               TabIndex        =   16
               Top             =   390
               Width           =   5805
               _ExtentX        =   10239
               _ExtentY        =   503
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.txtVISUAL txtInscMunicipal 
               Height          =   285
               Left            =   90
               TabIndex        =   15
               Top             =   390
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   503
               Caption         =   "Inscrição Municipal"
               Text            =   ""
               Restricao       =   2
               RetirarMascara  =   0   'False
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1545
            Left            =   120
            TabIndex        =   24
            Top             =   2040
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   2725
            Altura          =   1905
            Caption         =   " Localização do Imóvel"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboBairro 
               Height          =   315
               Left            =   990
               TabIndex        =   29
               Top             =   1050
               Width           =   8565
               _ExtentX        =   15108
               _ExtentY        =   556
               Caption         =   "Bairro"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNumero 
               Height          =   315
               Left            =   7950
               TabIndex        =   28
               Top             =   690
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               Caption         =   "Número"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtInscImob 
               Height          =   315
               Left            =   120
               TabIndex        =   27
               Top             =   330
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   556
               Caption         =   "Insc. Imobiliária"
               Text            =   ""
               Restricao       =   2
            End
            Begin VTOcx.cboVISUAL cboTipoLogr 
               Height          =   315
               Left            =   540
               TabIndex        =   26
               Top             =   690
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   556
               Caption         =   "Logradouro"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Editavel        =   -1  'True
            End
            Begin VTOcx.cboVISUAL cboLogr 
               Height          =   315
               Left            =   3750
               TabIndex        =   25
               Top             =   705
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5010
         Left            =   -99969
         TabIndex        =   8
         Top             =   30
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8837
         _Version        =   131082
         TabGuid         =   "TCIS106.frx":0970
         Begin VTOcx.grdVISUAL GrdDados 
            Height          =   3615
            Left            =   180
            TabIndex        =   30
            Top             =   1320
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   6376
         End
         Begin VTOcx.fraVISUAL fraVISUAL2 
            CausesValidation=   0   'False
            Height          =   1200
            Left            =   180
            TabIndex        =   32
            Top             =   90
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   2117
            Altura          =   1905
            Caption         =   " Dados do Veículo de Divulgacão"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtImConsulta 
               Height          =   285
               Left            =   90
               TabIndex        =   36
               Top             =   390
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   503
               Caption         =   "Inscrição Municipal"
               Text            =   ""
               Restricao       =   2
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtContribuinteConsulta 
               Height          =   285
               Left            =   3690
               TabIndex        =   35
               Top             =   390
               Width           =   5805
               _ExtentX        =   10239
               _ExtentY        =   503
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   16777215
            End
            Begin VTOcx.cmdVISUAL cmdVISUAL1 
               Height          =   285
               Left            =   3330
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   390
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               Caption         =   ""
               Acao            =   5
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cboVISUAL cboMovimentoConsulta 
               Height          =   315
               Left            =   780
               TabIndex        =   33
               Top             =   735
               Width           =   8730
               _ExtentX        =   15399
               _ExtentY        =   556
               Caption         =   "Movimento"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCIS106.frx":0998
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   1138
      Icone           =   "TCIS106.frx":2ABB
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   4
      Top             =   6075
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   6450
         TabIndex        =   31
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7620
         TabIndex        =   1
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   5280
         TabIndex        =   0
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8790
         TabIndex        =   2
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Threed.SSCheck chkCad 
      Height          =   195
      Index           =   4
      Left            =   75
      TabIndex        =   6
      Top             =   330
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   344
      _Version        =   196610
      Caption         =   "Cadastrar"
   End
End
Attribute VB_Name = "TCIS106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As VSImposto
'Dim Transportador As eTransportador
'Dim Contribuinte As eContribuinte
Dim Atividade As Atividade
Dim Imovel As eImovel
Dim GraveiContrib As Boolean
Dim VaiGravarSocio As Boolean
Dim InscricaoMunicipal As String
Dim InscricaoAuxiliar As String
Dim Anuncio As eAnuncio

Private Function VerificaCpfCgc() As Boolean
    Dim strCPF As String
    Dim strCGC As String
    Dim blnValido As Boolean
    
    blnValido = True
    
    If Trim(txtCgc) = "" Then Exit Function
    If Len(Edita.TiraTudo(txtCgc)) = 11 Then
        strCPF = Edita.TiraTudo(txtCgc)
        If Util.ValidaCpf(strCPF) = False Then
            blnValido = False
        End If
        Select Case strCPF
            Case String(11, "1")
                blnValido = False
            Case String(11, "2")
                blnValido = False
            Case String(11, "3")
                blnValido = False
            Case String(11, "4")
                blnValido = False
            Case String(11, "5")
                blnValido = False
            Case String(11, "6")
                blnValido = False
            Case String(11, "7")
                blnValido = False
            Case String(11, "8")
                blnValido = False
            Case String(11, "9")
                blnValido = False
            Case String(11, "0")
                blnValido = False
        End Select
    ElseIf Len(Edita.TiraTudo(txtCgc)) = 14 Then
        strCGC = Edita.TiraTudo(txtCgc)
'        If Util.ValidaCgc(strCGC) = False Then
'            blnValido = False
'            Exit Function
'        End If
    Else
        blnValido = False
    End If
    
    VerificaCpfCgc = blnValido
    
    If blnValido = False Then
        Util.Avisa "Cpf ou Cnpj Inválido"
    End If
    
End Function

Private Sub cboAtivServ_LostFocus()
    Dim RetFator As String
    If Trim(cboAtivServ) = "" Then Exit Sub
    If Atividade.BuscaFator(cboAtivServ, RetFator) Then
        txtFator.Visible = True
        txtFator.Caption = RetFator
        txtFator.Tag = "Fator"
        txtFator.SetFocus
    Else
        txtFator.Visible = False
        txtFator.Tag = ""
    End If
End Sub

Private Sub cboClassAtiv_Click()
    Atividade.PreencherCboAtiv cboAtivServ, CStr(cboClassAtiv.Coluna(1).Valor)
End Sub

Private Sub cboContador_Click()
    If cboContador.ListIndex = -1 Then Exit Sub
    If Contribuinte.Buscar(CStr(cboContador.Coluna(1).Valor), , False) = True Then
        txtCgcEscritorio = "": txtCrcContador = "": txtCpfContador = ""
        txtCrcContador = Contribuinte.Registro
        If Len(Trim(Contribuinte.CgcCpf)) = 14 And Not IsNumeric(Contribuinte.CgcCpf) Then
            txtCpfContador = Contribuinte.CgcCpf
            txtCgcEscritorio.SetFocus
        Else
            txtCgcEscritorio = Contribuinte.CgcCpf
            txtCpfContador.SetFocus
        End If
    End If
End Sub

Private Sub cboEstabelece_LostFocus()
    If cboEstabelece = "SIM" Then
        txtIc.Tag = "Insc. Cadastral"
        cboImovel.Tag = "Imovel"
    Else
        txtIc.Tag = ""
        cboImovel.Tag = ""
    End If
End Sub

Private Sub cboItem_Click()
 On Error Resume Next
    txtValor = cboItem.Coluna(2).Valor
    Calcula
End Sub

Private Sub cboTipoLogr_Click()
    If cboTipoLogr.ListIndex = -1 Then Exit Sub
    Endereco.PreencherCboRua cboLogr, cboTipoLogr
End Sub

Private Sub cboTipoLogrRepresentante_Click()
    If cboTipoLogrRepresentante.ListIndex = -1 Then Exit Sub
    Endereco.PreencherCboRua cboLogrRepresentante, cboTipoLogrRepresentante
End Sub

Private Sub cboTipoLogrSocio_Click()
    If cboTipoLogrSocio.ListIndex = -1 Then Exit Sub
    Endereco.PreencherCboRua cboLogrSocio, cboTipoLogrSocio
End Sub

Private Sub chkCad_Click(Index As Integer, Value As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            fraSocio1.Enabled = chkCad(Index).Value
            fraSocio2.Enabled = chkCad(Index).Value
            cmdAdEdif.Enabled = chkCad(Index).Value
            txtCpfSocio.SetFocus
        Case 1
            fraContador.Enabled = chkCad(Index).Value
            cboContador.SetFocus
        Case 2
            fraRepresen1.Enabled = chkCad(Index).Value
            fraRepresen2.Enabled = chkCad(Index).Value
            txtCpfRepresentante.SetFocus
        Case 3
            fraTrans.Enabled = chkCad(Index).Value
            cmdAdVeiculo.Enabled = chkCad(Index).Value
            cboAtividadeVeiculo.Enabled = chkCad(Index).Value
            cboAtividadeVeiculo.SetFocus
        Case 4
            fraAnu.Enabled = chkCad(Index).Value
            cmdAdAnuncio.Enabled = chkCad(Index).Value
            cboMovimento.SetFocus
    End Select
End Sub

Private Sub cmdAdAnuncio_Click()
   Dim RetIm                           As String
    Dim i                                   As Byte
    Dim Sql                               As String
    Dim rs                                As VSRecordset
    Dim Index                            As Integer
    If cboMovimento.ListIndex = -1 Then Exit Sub
    
'    grid.ColumnHeaders.Add , , "Item"
'    grid.ColumnHeaders.Add , , "Movimento"
'    grid.ColumnHeaders.Add , , "Descrição"
'    grid.ColumnHeaders.Add , , "Data Instalação"
'    grid.ColumnHeaders.Add , , "Área Total"
'    grid.ColumnHeaders.Add , , "Mutiplicador"
'    grid.ColumnHeaders.Add , , "Valor a pagar"
'    grid.ColumnHeaders.Add , , "Status"
'    grid.ColumnHeaders.Add , , "Período"
    
   Index = grdAnuncio.ListItems.Count + 1
   grdAnuncio.ListItems.Add Index, , Index
   grdAnuncio.ListItems.Item(Index).SubItems(1) = cboMovimento.Coluna(0).Valor & " - " & cboMovimento.Text
   grdAnuncio.ListItems.Item(Index).SubItems(2) = txtDimensao
   grdAnuncio.ListItems.Item(Index).SubItems(3) = txtDataInstalacao
   grdAnuncio.ListItems.Item(Index).SubItems(4) = txtArea
   grdAnuncio.ListItems.Item(Index).SubItems(5) = txtValor
   grdAnuncio.ListItems.Item(Index).SubItems(6) = txtFatorMutiplicador
   grdAnuncio.ListItems.Item(Index).SubItems(7) = txtValorApagar
   grdAnuncio.ListItems.Item(Index).SubItems(8) = CboStatus.Coluna(1).Valor & " - " & CboStatus.Text
   grdAnuncio.ListItems.Item(Index).SubItems(9) = txtPeriodo
   
   cboMovimento.ListIndex = -1
   txtDimensao = ""
   txtDataInstalacao = ""
   txtArea = ""
   txtValor = ""
   txtFatorMutiplicador = ""
   txtValorApagar = ""
   CboStatus.ListIndex = -1
   txtPeriodo = ""
   txtDimensao.SetFocus
    txtDimensao.SetFocus
   End Sub

Private Sub CmdConsultaContribuinte_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscMunicipal, txtContribuinte
End Sub

Private Sub cmdSalvar_Click()
    On Error GoTo trata
    Dim Item As Object
    'GRAVANDO ANUNCIOS
     Anuncio.Excluir txtInscMunicipal 'EXCLUI ANUNCIOS
'        cboMovimento.SetarLinha Trim(Left(.SelectedItem.SubItems(1), 9)), 0
'        txtDimensao = .SelectedItem.SubItems(2)
'        txtDataInstalacao = .SelectedItem.SubItems(3)
'        txtArea = .SelectedItem.SubItems(4)
'        CboStatus.SetarLinha Left(.SelectedItem.SubItems(8), 1), 1
'        txtPeriodo = .SelectedItem.SubItems(9)
'        txtValor = .SelectedItem.SubItems(5)
'        txtFatorMutiplicador = .SelectedItem.SubItems(6)
'        txtValorApagar = .SelectedItem.SubItems(7)

         If grdAnuncio.ListItems.Count <> "0" Then
             For Each Item In grdAnuncio.ListItems
                 With Anuncio
                     .Im = txtInscMunicipal
                     .ICAD = Bdados.Converte(Item.Text, tctexto)
                     .Movimento = Trim(Bdados.Converte(Trim(Left(Item.SubItems(1), 9)), tctexto))
                     .Dimensao = Item.SubItems(2)
                     .DataInstalacao = Item.SubItems(3)
                     .Area = Item.SubItems(4)
                     .Status = Left(Item.SubItems(8), 1)
                     .Periodo = Item.SubItems(9)
                     .Valor = Item.SubItems(5)
                     .Mutiplicador = Item.SubItems(6)
                     .Valor_Apagar = Item.SubItems(7)
                     .Salvar
                 End With
             Next
         End If
    Call Util.Informa("Registro gravado com sucesso." & InscricaoMunicipal & ".")
    cmdLimpar_Click
    Screen.MousePointer = 0
    Exit Sub
trata:
        Erro Err.Number & " - " & Err.Description
        Screen.MousePointer = 0
        Exit Sub
        Resume
End Sub

Private Sub cmdAdAtiv_Click()
    Dim CodGrupo As Double
    CodGrupo = cboClassAtiv.ListIndex
    TATV101.Tag = cboClassAtiv.Text
    TATV101.Show 1
    cboClassAtiv_Click
    cboClassAtiv.ListIndex = CodGrupo
    If TCIS101.Tag <> "" Then
        cboAtivServ.ListIndex = ListIndexDe(cboAtivServ, TCIS101.Tag)
    End If
    Unload TATV101
End Sub

Private Sub cmdAdEdif_Click()
    Dim ItmX As Object
    Dim i As Byte
    If Trim(txtCpfSocio) = "" Then Exit Sub
    Set ItmX = grdSocio.ListItems.Add(, , txtCpfSocio)
    With ItmX
        .SubItems(1) = txtNomeSocio
        .SubItems(2) = txtCargoSocio
        .SubItems(3) = cboTipoLogrSocio
        .SubItems(4) = cboLogrSocio
        .SubItems(5) = txtNumSocio
        .SubItems(6) = txtCompSocio
        .SubItems(7) = cboBairroSocio
        .SubItems(8) = txtTelSocio
        .SubItems(9) = txtCidadeSocio
        .SubItems(10) = cboUFSocio
    End With
    txtCpfSocio = ""
    txtNomeSocio = ""
    txtCargoSocio = ""
    cboTipoLogrSocio = ""
    cboLogrSocio = ""
    txtNumSocio = ""
    txtCompSocio = ""
    cboBairroSocio = ""
    txtTelSocio = ""
    txtCidadeSocio = ""
    cboUFSocio.ListIndex = -1
    txtCpfSocio.SetFocus
End Sub

Private Sub MontaCabGrid()
'grid veiculo
    'grdVeiculo.ColumnHeaders.Add , , "Veículo"
    
End Sub

Private Sub cmdAdVeiculo_Click()
    Dim RetIm       As String
    Dim i               As Byte
    Dim Index        As Integer
    
    Dim Sql           As String
    Dim rs            As VSRecordset
    
    If Trim(txtPlaca) = "" Then Exit Sub
    If Transportador.VerificaChassi(txtChassi, RetIm) Then
        Util.Informa "Chassi cadastrado para contribuinte IM = '" & RetIm & "'."
        Exit Sub
    End If
    
    Index = grdVeiculo.ListItems.Count + 1
    grdVeiculo.ListItems.Add Index, , txtVeiculo
    grdVeiculo.ListItems.Item(Index).SubItems(1) = txtMarca
    grdVeiculo.ListItems.Item(Index).SubItems(2) = txtModelo
    grdVeiculo.ListItems.Item(Index).SubItems(3) = txtAnoFabric
    grdVeiculo.ListItems.Item(Index).SubItems(4) = txtPlaca
    grdVeiculo.ListItems.Item(Index).SubItems(5) = txtChassi
    grdVeiculo.ListItems.Item(Index).SubItems(6) = txtMunicipio
    grdVeiculo.ListItems.Item(Index).SubItems(7) = cboUFTransp
    grdVeiculo.ListItems.Item(Index).SubItems(8) = txtLicenca
    grdVeiculo.ListItems.Item(Index).SubItems(9) = cboAtividadeVeiculo.Coluna(1).Valor & " - " & cboAtividadeVeiculo.Text
    grdVeiculo.ListItems.Item(Index).SubItems(10) = txtInicioAtividadeCarro
'    Set ItmX = grdVeiculo.ListItems.Add(, , txtVeiculo)
'    With ItmX
'        .SubItems(1) = txtMarca
'        .SubItems(2) = txtModelo
'        .SubItems(3) = txtAnoFabric
'        .SubItems(4) = txtPlaca
'        .SubItems(5) = txtChassi
'        .SubItems(6) = txtMunicipio
'        .SubItems(7) = cboUFTransp
'        .SubItems(8) = txtLicenca
'    End With
    txtVeiculo = ""
    txtMarca = ""
    txtModelo = ""
    txtInicioAtividadeCarro = ""
    txtAnoFabric = ""
    txtPlaca = ""
    txtChassi = ""
    txtMunicipio = ""
    txtLicenca = ""
    txtCidadeSocio = ""
    cboUFTransp.ListIndex = -1
    cboAtividadeVeiculo.ListIndex = -1
    cboAtividadeVeiculo.SetFocus
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{tab}"
End Sub

Private Sub cmdImprimir_Click()
        Screen.MousePointer = 11
        If Trim(InscricaoMunicipal) <> "" Then Imposto.ImprimeFC InscricaoMunicipal, Rpt
        Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
'    'tabCadastro.TabEnable'd(0) = True
 '   tabCadastro.Tabs(1).Selected = True
'    txtProtocolo.SetFocus
    InscricaoMunicipal = ""
    Edita.LimpaCampos Me
    GraveiContrib = False
    VaiGravarSocio = False
    'grdSocio.ListItems.Clear
'    grdVeiculo.ListItems.Clear
    grdAnuncio.ListItems.Clear
'    chkCad(0).Value = 0
'    chkCad(1).Value = 0
'    chkCad(2).Value = 0
'    chkCad(3).Value = 0
    txtCidade = Aplicacoes.Municipio
'    cboUF.Text = "MA"
    txtCep = CepCliente
End Sub

Private Sub cmdSair_Click()
        Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtImConsulta, txtContribuinteConsulta
End Sub

Private Sub Form_Activate()
    cboMovimento.Preencher Bdados, "Select tip_cod_imposto as [Código Receita] , tip_nome_imposto as Tributo FROM Tab_Imposto where tip_sigla_Imposto = '" & Temp.PegaParametro(Bdados, "NOME ANUNCIO") & "'", 1
    CboStatus.PreencherGeral Bdados, "STATUS ANUNCIO"
End Sub

Private Sub Form_Load()
    Set cadastro = New VSImposto
    Set Atividade = New Atividade
    Set Imovel = New eImovel
    Set Anuncio = New eAnuncio
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision

    Screen.MousePointer = 0
    GraveiContrib = False
    'txtCidade = Aplicacoes.Municipio
    'cboUF.Text = Temp.PegaParametro(Bdados, "ESTADO UF")
    'txtCep = CepCliente
'    MontaCabGrid
    Rem Anuncio.MontarGrid grdAnuncio
    
End Sub



Private Sub grdAnuncio_DblClick()
   Dim Contador As Integer
    
    With grdAnuncio
        If grdAnuncio.SelectedItem Is Nothing Then Exit Sub
        cboMovimento.SetarLinha Trim(Left(.SelectedItem.SubItems(1), 9)), 0
        txtDimensao = .SelectedItem.SubItems(2)
        txtDataInstalacao = .SelectedItem.SubItems(3)
        txtArea = .SelectedItem.SubItems(4)
        CboStatus.SetarLinha Left(.SelectedItem.SubItems(8), 1), 1
        txtPeriodo = .SelectedItem.SubItems(9)
        txtValor = .SelectedItem.SubItems(5)
        txtFatorMutiplicador = .SelectedItem.SubItems(6)
        txtValorApagar = .SelectedItem.SubItems(7)
        .ListItems.Remove .SelectedItem.Index
        txtArea_LostFocus
        For Contador = 1 To grdAnuncio.ListItems.Count
            grdAnuncio.ListItems(Contador) = Contador
            grdAnuncio.ListItems(Contador).SubItems(8) = Left(grdAnuncio.ListItems(Contador).SubItems(8), Len(grdAnuncio.ListItems(Contador).SubItems(8)) - 2) & Format(Contador, "00")
        Next
    End With
End Sub


Private Sub grdSocio_DblClick()
    If grdSocio.SelectedItem Is Nothing Then Exit Sub
    txtCpfSocio = grdSocio.SelectedItem
    txtNomeSocio = grdSocio.SelectedItem.SubItems(1)
    txtCargoSocio = grdSocio.SelectedItem.SubItems(2)
    cboTipoLogrSocio = grdSocio.SelectedItem.SubItems(3)
    cboLogrSocio = grdSocio.SelectedItem.SubItems(4)
    txtNumSocio = grdSocio.SelectedItem.SubItems(5)
    txtCompSocio = grdSocio.SelectedItem.SubItems(6)
    cboBairroSocio = grdSocio.SelectedItem.SubItems(7)
    txtTelSocio = grdSocio.SelectedItem.SubItems(8)
    txtCidadeSocio = grdSocio.SelectedItem.SubItems(9)
    cboUFSocio = grdSocio.SelectedItem.SubItems(10)
    grdSocio.ListItems.Remove (grdSocio.SelectedItem.Index)
End Sub

Private Sub grdVeiculo_DblClick()
    If grdVeiculo.SelectedItem Is Nothing Then Exit Sub
    txtVeiculo = grdVeiculo.SelectedItem
    txtMarca = grdVeiculo.SelectedItem.SubItems(1)
    txtModelo = grdVeiculo.SelectedItem.SubItems(2)
    txtAnoFabric = grdVeiculo.SelectedItem.SubItems(3)
    txtPlaca = grdVeiculo.SelectedItem.SubItems(4)
    txtMunicipio = grdVeiculo.SelectedItem.SubItems(6)
    cboUFTransp = grdVeiculo.SelectedItem.SubItems(7)
    txtLicenca = grdVeiculo.SelectedItem.SubItems(8)
    txtChassi = grdVeiculo.SelectedItem.SubItems(5)
    Dim Pos As Integer
    Pos = InStr(grdVeiculo.SelectedItem.SubItems(9), "-")
    cboAtividadeVeiculo.ListIndex = ListIndexDe(cboAtividadeVeiculo, CStr(Trim(Right(grdVeiculo.SelectedItem.SubItems(9), Len(grdVeiculo.SelectedItem.SubItems(9)) - Pos - 1))))
    txtInicioAtividadeCarro = grdVeiculo.SelectedItem.SubItems(10)
    grdVeiculo.ListItems.Remove (grdVeiculo.SelectedItem.Index)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set Transportador = Nothing
    'Set Contribuinte = Nothing
    'Set Endereco = Nothing
    'Set Contador = Nothing
    'Set Imovel = Nothing
    'Set Representante = Nothing
End Sub

Private Sub txtArea_LostFocus()
 Calcula
End Sub

Private Sub txtcgc_LostFocus()
    On Error GoTo TrataErro
    If Trim(txtCgc) = "" Then Exit Sub
    
    If txtCgc = "99999999999" Or txtCgc = "999.999.999-99" Or txtCgc = "00000000000" Or txtCgc = "000.000.000-00" Then
        Util.Avisa "Valor do CPF inválido."
        txtCgc.SetFocus
    End If
    
    'TEMPORARIAMENTE(13/07/2004)
    If Len(Edita.TiraTudo(txtCgc)) = 11 Then
        txtCgc.Formato = formCPF
    ElseIf Len(Edita.TiraTudo(txtCgc)) = 14 And IsNumeric(Edita.TiraTudo(txtCgc)) Then
        txtCgc.Formato = formCGC
    Else
        Util.Informa "Cpf ou Cnpj inválido."
        tabCadastro.Tabs(1).Selected = True
        txtCgc.SetFocus
        Exit Sub
    End If
    
    If VerificaCpfCgc = False Then
        txtCgc.SetFocus
        Exit Sub
    End If
    
    If txtCgc = "" Then Exit Sub
    If cadastro.VerificaEmpresaAntiga(txtCgc, txtRazao) = 1 Then
        If Not Util.Confirma("Já existe uma empresa cadastrada com o mesmo CNPJ/CPF. Confirma cadastro.") Then
            txtCgc.SetFocus
            Exit Sub
        End If
    End If
    txtCgc.Formato = formNenhum
    
    Exit Sub
TrataErro:
    Util.Erro Err.Description
End Sub

Private Sub txtCidade_LostFocus()
    If Trim(txtCidade) = "" Then
        txtCidade = Aplicacoes.Municipio
        cboUF.Text = "MA"
        txtCep = CepCliente
    End If
End Sub

Private Sub txtCpfSocio_LostFocus()
    If Trim(txtCpfSocio) = "" Then Exit Sub
    If Socio.Buscar(, txtCpfSocio) Then
        With Socio
            txtCpfSocio = .Cpf
            txtNomeSocio = .Nome
            txtCargoSocio = .Cargo
            cboTipoLogrSocio = .TipoLogr
            cboLogrSocio = .Logr
            txtNumSocio = .Numero
            txtCompSocio = .Complemento
            cboBairroSocio = Bairro
            txtTelSocio = .Telefone
            txtCidadeSocio = .Cidade
            cboUFSocio = .Uf
        End With
    Else
        If Contribuinte.Buscar(, txtCpfSocio, False) Then
            With Contribuinte
                txtCpfSocio = .CgcCpf
                txtNomeSocio = .Nome
                txtCargoSocio = ""
                cboTipoLogrSocio = .Logradouro
                cboLogrSocio = .NomeLogradouro
                txtNumSocio = .Numero
                txtCompSocio = .Complemento
                cboBairroSocio = Bairro
                txtTelSocio = .FoneFax
                txtCidadeSocio = .Cidade
                cboUFSocio = .Uf
            End With
        End If
    End If
End Sub

Private Sub txtFatorMutiplicador_Change()
    If txtFatorMutiplicador = "" Or txtFatorMutiplicador = "0" Then
        txtFatorMutiplicador = 1
    End If
    Calcula
End Sub

Private Sub txtic_LostFocus()
    If Trim(txtIc) = "" Then Exit Sub
    If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) <> 1 Then
        txtIc = cadastro.FormataInscricao(txtIc, InscImovel)
    End If
    If Imovel.BuscarImovel(txtIc, cboTipoLogr, cboLogr, txtNum, txtComplemento, cboBairro, txtCep, txtCidade, cboUF) = False Then
        Util.Informa ("Imóvel não cadastrado.")
        cboTipoLogr.ListIndex = -1
        txtNum = ""
    End If
End Sub

Private Sub txtImRepresentante_LostFocus()
    If Trim(txtImRepresentante) = "" Then Exit Sub
    With Contribuinte
        If .Buscar(txtImRepresentante, , False) Then
            txtCpfRepresentante = .CgcCpf
            txtNomeRepresentante = .Nome
            cboTipoLogrRepresentante.SetarLinha .Logradouro
            cboLogrRepresentante = .NomeLogradouro
            txtNumRepresentante = .Numero
            txtComplementoRepresentante = .Complemento
            cboBairroRepresentante.SetarLinha .Bairro
            txtTelefoneRepresentante = .FoneFax
            txtCidadeRepresentante = .Cidade
            cboUfRepresentante.SetarLinha .Uf
        Else
            Util.Avisa "Contribuinte não encontrado."
        End If
    End With
End Sub

Private Sub txtchassi_lostfocus()
    Dim RetIm As String
    If Trim(txtChassi) <> "" Then
        If Transportador.VerificaChassi(txtChassi, RetIm) Then
            Informa "Chassi já cadastrado para contribuinte IM = " & RetIm & "."
            txtPlaca = ""
        End If
    End If
End Sub

Private Sub txtrazao_LostFocus()
    If Trim(txtRazao) = "" Then Exit Sub
    If cadastro.VerificaEmpresaAntiga(txtCgc, txtRazao) = 2 Then
        If Not Util.Confirma("Já existe uma empresa cadastrada com a mesma razao social. Confirma cadastro.") Then
            txtRazao.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub Calcula()
    txtValorApagar = Nvl(txtFatorMutiplicador, 1) * Nvl(txtValor, 0)
End Sub

Private Sub txtImConsulta_LostFocus()
    BuscaContribuinte txtImConsulta, txtContribuinteConsulta
    If txtImConsulta <> "" Then
        Anuncio.PreencherGrd GrdDados, txtImConsulta
    End If

End Sub

Private Sub txtValor_Change()
    Calcula
End Sub

Private Sub txtVISUAL6_Change()

End Sub
