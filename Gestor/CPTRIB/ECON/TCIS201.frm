VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIS201 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   ControlBox      =   0   'False
   Icon            =   "TCIS201.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   24
      Top             =   6135
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   6660
         TabIndex        =   17
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   7830
         TabIndex        =   18
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9000
         TabIndex        =   19
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   5340
      Left            =   60
      TabIndex        =   20
      Top             =   735
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   9419
      Caption         =   "Contribuintes"
      Descricao       =   "Exclui contribuintes através de seu IM"
      corFaixa        =   16711680
      Icone           =   "TCIS201.frx":08CA
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.fraVISUAL fraVISUAL3 
         Height          =   1560
         Left            =   105
         TabIndex        =   23
         Top             =   3690
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   2752
         Altura          =   1905
         Caption         =   " Detalhes"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Enabled         =   0   'False
         Begin VTOcx.cboVISUAL cboNaturezaJuridica 
            Height          =   510
            Left            =   90
            TabIndex        =   12
            Top             =   315
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   900
            Caption         =   "Natureza Jurídica"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
            Enabled         =   0   'False
         End
         Begin VTOcx.cboVISUAL cboClassiAtivi 
            Height          =   510
            Left            =   3405
            TabIndex        =   13
            Top             =   315
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   900
            Caption         =   "Classificação da Atividade"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
            Enabled         =   0   'False
         End
         Begin VTOcx.cboVISUAL cboAtivExercPoder 
            Height          =   510
            Left            =   6720
            TabIndex        =   14
            Top             =   315
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   900
            Caption         =   "Atividade Exercida pelo Poder"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
            Enabled         =   0   'False
         End
         Begin VTOcx.cboVISUAL cboAtividade 
            Height          =   510
            Left            =   105
            TabIndex        =   15
            Top             =   840
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   900
            Caption         =   "Atividade ou Serviço"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
            Enabled         =   0   'False
         End
         Begin VTOcx.cboVISUAL cboEstabelecido 
            Height          =   510
            Left            =   5025
            TabIndex        =   16
            Top             =   840
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   900
            Caption         =   "Estabelecido"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
            Enabled         =   0   'False
         End
      End
      Begin VTOcx.fraVISUAL fraVISUAL2 
         Height          =   1440
         Left            =   105
         TabIndex        =   22
         Top             =   2205
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   2540
         Altura          =   1905
         Caption         =   " Localização"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Enabled         =   0   'False
         Begin VTOcx.txtVISUAL txtTipoLogr 
            Height          =   480
            Left            =   60
            TabIndex        =   5
            Top             =   345
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   847
            Caption         =   "Logradouro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNomeLogrContrib 
            Height          =   285
            Left            =   1155
            TabIndex        =   6
            Top             =   540
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   503
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
         End
         Begin VTOcx.txtVISUAL txtNumero 
            Height          =   480
            Left            =   4770
            TabIndex        =   7
            Top             =   345
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   847
            Caption         =   "Nº"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtComp 
            Height          =   480
            Left            =   5640
            TabIndex        =   8
            Top             =   345
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   847
            Caption         =   "Compl."
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtCep 
            Height          =   480
            Left            =   4785
            TabIndex        =   11
            Top             =   840
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   847
            Caption         =   "CEP"
            Text            =   ""
            Enabled         =   0   'False
            Formato         =   4
            Restricao       =   2
            AlinhamentoRotulo=   1
            AgruparValores  =   0   'False
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtZona 
            Height          =   480
            Left            =   3000
            TabIndex        =   10
            Top             =   840
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   847
            Caption         =   "Zona ou Reg. Adm."
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtBairro 
            Height          =   480
            Left            =   60
            TabIndex        =   9
            Top             =   840
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   847
            Caption         =   "Distrito ou Bairro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
      End
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   1440
         Left            =   105
         TabIndex        =   21
         Top             =   720
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   2540
         Altura          =   1905
         Caption         =   " Contribuinte"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtIm 
            Height          =   480
            Left            =   120
            TabIndex        =   0
            Top             =   345
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   847
            Caption         =   "Ins. Municipal"
            Text            =   ""
            Formato         =   8
            Restricao       =   2
            AlinhamentoRotulo=   1
            AgruparValores  =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdBuscar 
            Height          =   330
            Index           =   0
            Left            =   1545
            TabIndex        =   1
            Top             =   510
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   582
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtCgc 
            Height          =   480
            Left            =   1950
            TabIndex        =   2
            Top             =   345
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   847
            Caption         =   "CPF/CNPJ"
            Text            =   ""
            Restricao       =   2
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNomeContrib 
            Height          =   480
            Left            =   105
            TabIndex        =   3
            Top             =   855
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   847
            Caption         =   "Nome/Razão Social"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtFantasia 
            Height          =   480
            Left            =   5190
            TabIndex        =   4
            Top             =   855
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   847
            Caption         =   "Nome Fantasia"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1138
      Icone           =   "TCIS201.frx":11A4
   End
End
Attribute VB_Name = "TCIS201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cadastro As VSImposto
Dim eContribuinte As eContribuinte

'Sub PreencheTela(Criterio As String)
'    Dim i As Byte
'    Dim RsISS As VSRecordset
'    Dim sql As String
'    Dim RsAux As VSRecordset
'    sql = "Select Tab_Contribuinte.* from Tab_Contribuinte "
'    sql = sql & " where " & Criterio & " and tci_tsc_cod_sit_cad = 1"
'    If Bdados.AbreTabela(sql, RsISS) Then
'        If RsISS!tci_tipo_contribuinte = 0 Then
'            Avisa "Contribuinte Pessoa Física."
'            Bdados.FechaTabela RsISS
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'        txtcgc.MaxLength = 20
'        txtcgc = "" & RsISS!TCI_CGC_CPF
'        txtIm = RsISS!TCI_IM
'        txtcgc.MaxLength = 14
'        txtrazao = RsISS!tci_nome
'        txtFantasia = "" & RsISS!tci_fantasia
'        txtCep = RsISS!tci_cep
'        cboNatJur.ListIndex = RsISS!tci_tnj_cod_natureza - 1
'        cboClassAtiv.ListIndex = RsISS!tci_tga_cod_grupo - 1
'        cboAtivPoder.ListIndex = RsISS!tci_tap_cod_ativ_poder - 1
'        txtTipoLogr = RsISS!tci_logradouro
'        txtLogr = RsISS!tci_nome_logradouro
'        txtnum = RsISS!tci_NUMERO
'        txtComplemento = RsISS!tci_COMPLEMENTO
'        txtBairro = RsISS!tci_BAIRRO
'        sql = "SELECT tae_nome  from tab_atividade_economica where tae_cae = " & RsISS!tci_tae_cae
'        If Bdados.AbreTabela(sql, RsAux) Then
'            cboAtivServ.Text = RsAux(0)
'        End If
'        Bdados.FechaTabela RsAux
'        cboEstabelece.ListIndex = RsISS!tci_estab - 1
'    Else
'        Call Util.Informa("Contribuinte não encontrado.")
'        txtIm.SetFocus
'    End If
'    Bdados.FechaTabela RsISS
'    Bdados.FechaTabela RsAux
'End Sub

Private Sub cmdEnter_Click()
    SendKeys "{tab}"
End Sub

Private Sub cabVisual_GotFocus()

End Sub

Private Sub cmdBuscar_Click(Index As Integer)
    Select Case Index
        Case 0
            AplicacoesVTFuncoes.BuscaNoEconomico TcoJuridica, txtIm
            
    End Select
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo trata
    If Trim(txtIm) = "" Then Util.Informa "Informe o contribuinte.": Exit Sub
    If Util.Confirma("Deseja eliminar o contribuinte?") Then
        'crítaca pra verificar se existem imoveis
        If eContribuinte.VerificaTEMImovel(txtIm) = True Then
            Call Util.Informa("Existe(m) imóvel(is) cadastrado(s) para o contribuinte. Verifique o Cadastro Imobiliário.")
            Screen.MousePointer = 0
            Exit Sub
        End If
        Screen.MousePointer = 11
        'crítica pra verificar se existe debito
        If eContribuinte.VerificaTEMDebito(txtIm) = True Then
            Call Util.Informa("Contribuinte com débito em aberto.")
            Screen.MousePointer = 0
            Exit Sub
        End If
        If eContribuinte.Excluir(txtIm) Then
            Util.Informa ("Registro eliminado com sucesso.")
            Screen.MousePointer = 0
            cmdNovo_Click
        End If
    End If
    Exit Sub
trata:
    Erro "Erro ao excluir"
End Sub

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    txtIm.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Cadastro = New VSImposto
    Set eContribuinte = New eContribuinte
    
    Screen.MousePointer = 0
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    eContribuinte.PreencherCboAtividadeEcon CboAtividade
    eContribuinte.PreencherCboNaturezaJur cboNaturezaJuridica
    eContribuinte.PreencherCboClasseAtividade cboClassiAtivi
    eContribuinte.PreencherCboAtividadePoder cboAtivExercPoder
    
    cboEstabelecido.PreencherGeral Bdados, "ESTABELECIDO"
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
        txtIm.Formato = formNenhum
    Else
        txtIm.Formato = formDoisDigitos
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set eContribuinte = Nothing
End Sub

Private Sub txtcgc_LostFocus()
    If Mid(txtCgc, 4, 1) = "." Then Exit Sub
     If txtCgc = "99999999999" Or txtCgc = "999.999.999-99" Or txtCgc = "00000000000" Or txtCgc = "000.000.000-00" Then
        Util.Avisa "Valor do CPF inválido."
        txtCgc.SetFocus
    End If
    If Len(txtCgc) = 11 Then
        txtCgc.Formato = formCPF
    ElseIf Len(txtCgc) = 14 Then
        txtCgc.Formato = formCGC
    ElseIf Trim(txtCgc) <> "" And Len(txtCgc) <> 18 Then
        Call Util.Informa("Número de CGC ou CPF inválido.")
        txtCgc.SetFocus
        txtCgc.Formato = formNenhum
        Exit Sub
    End If
    If txtCgc = "" Or txtIm <> "" Then txtCgc.Formato = formNenhum: Exit Sub
    Call MostraDados(txtIm, txtCgc)
    txtCgc.Formato = formNenhum
End Sub

Private Sub txtIM_LostFocus()
    If Trim(txtIm) = "" Then Exit Sub
    Call MostraDados(txtIm, txtCgc)
End Sub

Private Sub MostraDados(Im As String, Cgc As String)
    With eContribuinte
        If .Buscar(Im, Cgc, True) Then
            If .CodSitCadastral = 1 Then
                If .TipoContribuinte = 0 Then
                    Avisa "Contribuinte Pessoa Física."
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                txtIm = .Im
                txtCgc = .CgcCpf
                txtcgc_LostFocus
                txtNomeContrib = .Nome
                txtFantasia = .Fantasia
                txtCep = .Cep
                cboNaturezaJuridica.SetarLinha .CodNatureza, 1
                cboClassiAtivi.SetarLinha .CodGrupo, 1
                cboAtivExercPoder.SetarLinha .CodAtivPoder, 1
                txtTipoLogr = .Logradouro
                txtNomeLogrContrib = .NomeLogradouro
                txtNumero = .Numero
                txtComp = .Complemento
                txtBairro = .Bairro
                CboAtividade.SetarLinha .CodAtividade, 1
                cboEstabelecido.SetarLinha .Estabelecido + 1, 1
            Else
                Call Util.Informa("Contribuinte não encontrado.")
                txtIm.SetFocus
            End If
        End If
    End With
End Sub
