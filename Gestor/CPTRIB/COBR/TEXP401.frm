VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TEXP401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   27
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TEXP401.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1320
      Left            =   45
      TabIndex        =   13
      Top             =   645
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   2328
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   570
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   9720
         TabIndex        =   3
         Top             =   900
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   330
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   423
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   16777215
         PictureMaskColor=   16777215
         BackStyle       =   1
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
         Caption         =   "Nº Extrato:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   450
         TabIndex        =   19
         Top             =   990
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   16777215
         PictureMaskColor=   16777215
         BackStyle       =   1
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
         Caption         =   "Endereço:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   15
         Left            =   1410
         TabIndex        =   18
         Top             =   330
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   423
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   16777215
         PictureMaskColor=   16777215
         BackStyle       =   1
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
         Caption         =   "Inscrição:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtContrib 
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
         Left            =   3525
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   570
         Width           =   6000
      End
      Begin VB.TextBox txtIm 
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
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   1
         Top             =   570
         Width           =   1755
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
         Left            =   1380
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   930
         Width           =   8145
      End
      Begin VB.TextBox txtExtrato 
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
         Left            =   90
         TabIndex        =   0
         Top             =   570
         Width           =   1185
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   12
      Top             =   6240
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdImprime 
         Height          =   375
         Left            =   6645
         TabIndex        =   6
         Top             =   120
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         Caption         =   "&Imprimir Extrato"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   8595
         TabIndex        =   7
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9825
         TabIndex        =   8
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   9
      Top             =   -420
      Width           =   375
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9660
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   10
      Top             =   1350
      Visible         =   0   'False
      Width           =   795
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1138
      Icone           =   "TEXP401.frx":2123
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1080
      Left            =   7635
      TabIndex        =   14
      Top             =   5130
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   1905
      Altura          =   1905
      Caption         =   " Resultados Parciais do Parcelamento"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   2
         Left            =   1725
         TabIndex        =   23
         Top             =   330
         Width           =   1530
         _ExtentX        =   2699
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
         Caption         =   "Data Vencimento:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   7
         Left            =   165
         TabIndex        =   22
         Top             =   330
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Valor a Pagar:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   5
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtValorPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   180
         TabIndex        =   21
         Tag             =   "Valor Pago"
         Top             =   630
         Width           =   1200
      End
      Begin VB.TextBox txtDtVence 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1785
         TabIndex        =   5
         Tag             =   "Data Vencimento"
         Top             =   615
         Width           =   1440
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL4 
      Height          =   1080
      Left            =   45
      TabIndex        =   15
      Top             =   5130
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   1905
      Altura          =   1905
      Caption         =   " Detalhes"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   19
         Left            =   90
         TabIndex        =   24
         Top             =   345
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   370
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
         Caption         =   "Observações :"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtObservacao 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   345
         Width           =   6165
      End
   End
   Begin VTOcx.grdVISUAL lstParcelas 
      Height          =   1770
      Left            =   45
      TabIndex        =   25
      Top             =   3570
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   3122
      Caption         =   "DAM"
      CorTitulo       =   32768
      CorCaption      =   -2147483634
      CorDica         =   255
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.grdVISUAL grdExtrato 
      Height          =   1770
      Left            =   30
      TabIndex        =   26
      Top             =   2010
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   3122
      Caption         =   "Extrato(s)"
      CorTitulo       =   32768
      CorCaption      =   -2147483634
      CorDica         =   255
      OcultarRodape   =   -1  'True
   End
End
Attribute VB_Name = "TEXP401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim MaxCotas As Byte
Dim CodImp As String
Dim Cgc As String
Dim EnderecoContrib As String
Dim CodPagamento As String
Dim SqlExtrato As String
Dim MantemData  As Boolean
Dim Valor As Double
Dim Multa As Double
Dim Juros As Double
Dim Saldo As Double
Dim Atualizacao As Double
Dim Desconto As Double

Dim Periodo As String
Dim Tributo As String

Private Sub Pega_Dados()
    Dim i As Integer
    Dim Base As String
    Dim Base2 As String
    Tributo = ""
    Periodo = ""
    
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked = True Then
            
            
            If Tributo = "" Then
                Tributo = "[ " & lstParcelas.ListItems(i).SubItems(1) & " ]"
            Else
                If lstParcelas.ListItems.Count - i = 1 Then
                    'Checo se é o fim da tabela, se for coloco a virgula...
                    If Base <> lstParcelas.ListItems(i).SubItems(1) Then
                        Tributo = Tributo & "e [ " & lstParcelas.ListItems(i).SubItems(1) & " ]"
                    End If
                Else
                   'If para não repetir a sigla outra vez...
                    If Base <> lstParcelas.ListItems(i).SubItems(1) Then
                        Tributo = Tributo & ", [ " & lstParcelas.ListItems(i).SubItems(1) & " ]"
                    End If
                End If
            End If
            
             If Periodo = "" Then
                    Periodo = "[ " & lstParcelas.ListItems(i).SubItems(2) & " ]"
                Else
                    If lstParcelas.ListItems.Count - i = 1 Then
                        'Checo se é o fim da tabela, se for coloco a virgula...
                        If Base2 <> lstParcelas.ListItems(i).SubItems(2) Then
                            Periodo = Periodo & "e [ " & lstParcelas.ListItems(i).SubItems(2) & " ]"
                        End If
                    Else
                        If Base2 <> lstParcelas.ListItems(i).SubItems(2) Then
                            Periodo = Periodo & ", [ " & lstParcelas.ListItems(i).SubItems(2) & " ]"
                        End If
                    End If
             End If
            Base = lstParcelas.ListItems(i).SubItems(1)
            Base2 = lstParcelas.ListItems(i).SubItems(2)
        End If
    Next
    
End Sub


Private Sub AtualizaValores()
    Dim i As Integer
    Valor = 0
    Juros = 0
    Multa = 0
    Saldo = 0
    Atualizacao = 0
    Desconto = 0
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked = True Then
            Valor = Valor + lstParcelas.ListItems(i).SubItems(4)
            Atualizacao = Atualizacao + lstParcelas.ListItems(i).SubItems(5)
            If Val(lstParcelas.ListItems(i).SubItems(4)) >= 1 Then
                Desconto = Desconto + lstParcelas.ListItems(i).SubItems(6)
            Else
                Desconto = Desconto + 0
            End If
            Juros = Juros + lstParcelas.ListItems(i).SubItems(7)
            Multa = Multa + lstParcelas.ListItems(i).SubItems(8)
            If lstParcelas.ListItems(i).SubItems(9) <> "" Then
                Saldo = Saldo + lstParcelas.ListItems(i).SubItems(9)
            Else
                Saldo = Saldo + 0
            End If
        End If
    Next
End Sub

Private Sub cmdBuscar_Click()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Conta As Double
    Dim i As Byte
    Dim Extrato As String
    Dim Condicao As String
    'If Trim(txtExtrato) = "" Then Exit Sub
    
    lstParcelas.ListItems.Clear
    SqlExtrato = "SELECT VEX_COD_EXTRATO AS Numero, VEX_INSCRICAO as Inscricao, VEX_RAZAO as Contribuinte, VEX_ENDERECO as Endereco  FROM VIS_EXTRATO"
    
    If Trim(txtExtrato) <> "" Then
        If Condicao <> "" Then Condicao = Condicao & " AND "
        Condicao = Condicao & "VEX_COD_EXTRATO =" & txtExtrato.Text
    End If
    
    If Trim(txtIM) <> "" Then
        If Condicao <> "" Then Condicao = Condicao & " AND "
        Condicao = Condicao & "VEX_INSCRICAO = '" & txtIM.Text + "'"
    End If
    
    If Condicao <> "" Then
        SqlExtrato = SqlExtrato & " WHERE " & Condicao
    End If
        
    grdExtrato.Preencher Bdados, SqlExtrato, 900, 1800, 3000, 4800, 0
    
    If grdExtrato.ListItems.Count = 0 Then
        Util.Avisa "Nenhum registro encontrado"
    Else
        If txtExtrato <> "" Then grdExtrato_DblClick
        grdExtrato.SetFocus
    End If
    
End Sub

Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    lstParcelas.ListItems.Clear
    grdExtrato.ListItems.Clear
    txtExtrato.SetFocus
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImprime_Click()
      Dim Barra As Boolean
    Dim Cobranca As New VSCobranca
    Dim i As Integer
    Dim ListaDocs As String
    Dim M As Boolean
    
    M = False
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked Then
            M = True
            Exit For
        End If
    Next

    Barra = False
   
    If grdExtrato.SelectedItem Is Nothing Then
        Util.Avisa "Não existe extrato selecionado para a impressão"
        txtExtrato.SetFocus
        Exit Sub
    Else
        grdExtrato_DblClick
    End If
    If lstParcelas.ListItems.Count = 0 Then
        Util.Avisa "Não existe parcelas para a impressão"
        lstParcelas.SetFocus
        Exit Sub
    End If
   
    Screen.MousePointer = 11
    For i = 1 To lstParcelas.ListItems.Count
        ListaDocs = ListaDocs & lstParcelas.ListItems(i).Text & " (" & lstParcelas.ListItems(i).SubItems(5) & ") - "
    Next
    ListaDocs = Left(ListaDocs, Len(ListaDocs) - 3)
'    If Confirma("O extrato será usado para pagamento do débito?") Then
        If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TDAMExtratoBarra.rpt") Then Exit Sub
'        Barra = True
'    Else
'        If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TDAMExtrato.rpt") Then Exit Sub
'    End If
    With Rpt
        'Formulas do Dam...
        .Formulas "DATAVENCIMENTO", txtDtVence
        .Formulas "PARCELA", "UNICA"
        .Formulas "NODOCUMENTO", CStr(grdExtrato.SelectedItem.Text)
        If UCase(AplicacoesVTFuncoes.Municipio) = "PETROLINA" Then
            .Formulas "TXDAM", TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXTDAM"), 0), ".", ",")
        Else
            .Formulas "TXDAM", " "
        End If
        .Formulas "VENCIMENTONORMAL", txtDtVence
        .Formulas "NOSSONUMERO", CStr(grdExtrato.SelectedItem.Text)
        .Formulas "CODIGOTRIBUTO", Const_Extrato
        .Formulas "PERIDO", Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
        .Formulas "EMISSAO", Format(Date, "DD/MM/YYYY")
        'DADOS DO SACADO...
        .Formulas "NOME", grdExtrato.SelectedItem.SubItems(1) & " - " & grdExtrato.SelectedItem.SubItems(2)
        .Formulas "ENDERECO", grdExtrato.SelectedItem.SubItems(3)
        'Atualizo os valores...
        Call AtualizaValores
        .Formulas "VALORTRIBUTO", CStr(Format(CDbl(txtValorPago), "###,###,###,##0.00"))
        
         .Formulas "DEDUCAO", CStr(Format(Desconto, "###,###,###,##0.00"))
         .Formulas "TAXAEXPEDIENTE", TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
         'COLOCO A DESCRIÇÃO DA TXDAM...
         If UCase(AplicacoesVTFuncoes.Municipio) = "PETROLINA" Then
            .Formulas "MENSAGEM1", "[ TXDAM - " & TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",") & " ]"
         End If
         .Formulas "VALORJUROS", CStr(Format(Juros, "###,###,###,##0.00"))
         .Formulas "VALORMULTA", CStr(Format(Multa, "###,###,###,##0.00"))
         .Formulas "CORRECAO", CStr(Format(Atualizacao, "###,###,###,##0.00"))
         .Formulas "VALORTOTAL", CStr(Format(CDbl(Nvl(Trim(txtValorPago), 0)) + CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0)) + Juros + Atualizacao + Multa, "###,###,###,##0.00"))       ' + Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0), "###,###,###,##0.00"))
         'FORMULAS DO EXTRATO
        '.Formulas "VT_EXTRATO ", CStr(grdExtrato.SelectedItem.Text)
        .Formulas "VT_PRAZO ", txtDtVence
        'Atualizo os Dados da Observação referente as Período e Tributo...
        Call Pega_Dados
        .Formulas "MENSAGEM2", "Tributo(s) - " & Tributo
        .Formulas "MENSAGEM3", "Período(s) - " & Periodo
        .Formulas "VT_CONTRIBUINTE", grdExtrato.SelectedItem.SubItems(1)
        .Formulas "VT_RAZAO", grdExtrato.SelectedItem.SubItems(2)
        .Formulas "VT_ENDERECO", grdExtrato.SelectedItem.SubItems(3)
        .Formulas "VT_OBS_GERAL ", txtObservacao
        .Selecao = "{TAB_PAGAMENTO_EXTRATO.TPE_COD_PAGAMENTO_EXTRATO} = " & grdExtrato.SelectedItem.Text
        
        'If Barra Then
        '    .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
            Dim CgcPref As String
            
        '    CgcPref = Temp.PegaParametro(Bdados, "CGC CLIENTE")
        '    CgcPref = TiraTudo(CgcPref)
        '    .Formulas "CgcPrefeitura", "CNPJ " & Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
      '      '.Formulas "LinhaDigitavel", Cobranca.GeraCodBarra(txtExtrato, 0, CDbl(txtValorPago), PicBarra, Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), txtDtVence)
      '  Else
      '      .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
      '  End If
      '  .Titulo = "Extrato de Lançamento"
      '  .Arvore = False
      '  .Visualizar
        'If Barra Then
            .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
            CgcPref = Temp.PegaParametro(Bdados, "CGC CLIENTE")
            .Formulas "CgcPrefeitura", "CNPJ " & Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
            '.Formulas "RAZAO", grdExtrato.SelectedItem.SubItems(2)
            '.Formulas "ENDERECO", grdExtrato.SelectedItem.SubItems(3)
            Cobranca.ImprimeDamBarra Rpt, txtIM, Const_Extrato, CDbl(Nvl(Trim(txtValorPago), 0)) + Juros + Atualizacao + Multa, Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), PicBarra, txtDtVence, 0, grdExtrato.SelectedItem
        'Else
        '    .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        'End If
        'Bdados.GravaDados "TAB_PARAMETRO_TEXTO", ListaDocs, "TPT_TEXTO", "TPT_PARAMETRO = 'DOCUMENTOS EXTRATO'"
        .Titulo = "Extrato de Lançamento"
        .Arvore = False
        .Visualizar
        'Bdados.GravaDados "TAB_PARAMETRO_TEXTO", "", "TPT_TEXTO", "TPT_PARAMETRO = 'DOCUMENTOS EXTRATO'"
    End With
    Screen.MousePointer = 0
    Set Rpt = Nothing
End Sub

Private Sub cmdParcela_Click()
    Dim Sql As String
    Dim CCorrente As New ContaCorrente
    Dim rs As VSRecordset
    Dim Conta As New ContaCorrente
    Dim i As Integer
    On Error Resume Next
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    CodPagamento = Conta.GeraCodPagamento(EtsExtratoPagamento)
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked Then
            Bdados.InsereDados "TAB_PAGAMENTO_EXTRATO", Bdados.PreparaValor(CodPagamento, lstParcelas.ListItems(i).Text, Bdados.Converte(lstParcelas.ListItems(i).SubItems(4), TCDuplo), lstParcelas.ListItems(i).SubItems(1))
        End If
    Next
    'Conta.GeraPagamento txtIm, txtic, Const_Extrato, Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), txtDtVence, CDbl(txtValorPago), 0, 0, CodPagamento, 0, 0, 0, , EtcCreditoTributario
    'Cobranca.ImprimeDam Rpt, CodPagamento, txtIm, txtContrib, txtCgc, txtEndereco, "", "", Const_Extrato, Const_Extrato, "EXTRATO RESUMO DE LANÇAMENTO DE TRIBUTOS", Mid(Format(Date, "DD/MM/YYYY"), 4, 2) & Right(Format(Date, "DD/MM/YYYY"), 4), 0, 1, txtDtVence, txtValorPago, txtValorPago, 0, 0, 0, 0, "", txtObservacao, PicBarra
    Screen.MousePointer = 0
    Informa "Extrato gerado com sucesso."
    Bdados.FechaTabela rs
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscGrupo, txtIM
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub txtCotas_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub


Private Sub grdExtrato_DblClick()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Conta As Double
    Dim CCorrente As New ContaCorrente
    Dim i As Byte
    Dim Obrig As New Obrigacao
    If grdExtrato.SelectedItem Is Nothing Then Exit Sub
    
    Sql = "select tgt_valor_tributo, tgt_data_vencimento,tgt_im,tgt_tim_ic"
    Sql = Sql & " from tab_geracao_tributo where tgt_cod_pagamento = " & grdExtrato.SelectedItem.Text
    Screen.MousePointer = 11
    If Bdados.AbreTabela(Sql, rs) Then
        If Not MantemData Then txtDtVence = rs!TGT_DATA_VENCIMENTO
        
        SqlExtrato = "SELECT TPE_TGT_COD_PAGAMENTO FROM TAB_PAGAMENTO_EXTRATO"
        SqlExtrato = SqlExtrato & " WHERE TPE_COD_PAGAMENTO_EXTRATO =" & grdExtrato.SelectedItem.Text
        
        'MOVIMENTA CONTAS
        
        Sql = "SELECT tcc_im,tcc_tim_ic,tcc_tip_cod_imposto,tcc_periodo,tcc_num_conta,tcc_parcela,tcc_codigo_conta,tcc_tipo_inscricao FROM tab_conta_contribuinte where (tcc_tgt_ativo =0 or tcc_tgt_ativo is null) "
        Sql = Sql & " and tcc_codigo_conta in (" & SqlExtrato & ")"
        Sql = Sql & " ORDER BY TCC_PERIODO ASC"
        If Bdados.AbreTabela(Sql, rs) Then
            rs.MoveFirst
            Do While Not rs.EOF
                Obrig.BuscaDetalheObrigacao rs!tcc_codigo_conta
                CCorrente.MovimentaContaContribuinte rs!tcc_codigo_conta, txtDtVence, Obrig
                rs.MoveNext
                DoEvents
            Loop
        End If
        'FIM MOVIMENTO
        '0- aberta ; 1 - fechada ; 2 - Em parcelamento;3 - Divida Ativa
        
        Sql = "select tcc_codigo_conta AS Documento, tip_sigla_imposto as Tributo , tcc_periodo as Periodo, " & _
        " tcc_data_vencimento as Vencimento, tcc_imposto_original as Imposto,tcc_correcao_monetaria as Atualização,tcc_desconto_concedido as Desconto,tcc_juros_atual as Juros,tcc_multa_atual as Multa,tcc_imposto_original + tcc_juros_atual + tcc_multa_atual + tcc_correcao_monetaria - tcc_desconto_concedido as Saldo from tab_conta_contribuinte,tab_imposto,TAB_PAGAMENTO_EXTRATO where " & _
        " tcc_codigo_conta = TPE_TGT_COD_PAGAMENTO and tcc_tip_cod_imposto = tip_cod_imposto "
        Sql = Sql & " AND TPE_COD_PAGAMENTO_EXTRATO = " & grdExtrato.SelectedItem.Text
        
        lstParcelas.Preencher Bdados, Sql, 1000, 900, 900, 1200, 900, 900, 900, 900
        Conta = 0
        For i = 1 To lstParcelas.ListItems.Count
            Conta = Conta + CDbl(lstParcelas.ListItems(i).SubItems(7))
        Next
        txtValorPago = Format(Conta, Const_Monetario)
        'ATUALIZO OS VALORES...
        Call AtualizaValores
        txtValorPago = CStr(Format(Valor + Juros + Multa + Atualizacao - Desconto, "###,###,###,##0.00"))
    Else
        txtExtrato = ""
        Avisa "Nº de extrato não encontrado."
        txtExtrato.SetFocus
    End If
    Screen.MousePointer = 0
    MantemData = False
    Bdados.FechaTabela rs
End Sub

Private Sub txtDtVence_LostFocus()
    If Trim(txtDtVence) = "" Then Exit Sub
    txtDtVence = Edita.FormataTexto(txtDtVence, Data)
    MantemData = True
    grdExtrato_DblClick
End Sub

Private Sub txtIm_Change()
    txtContrib.Text = ""
    txtEndereco.Text = ""
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim RsPag As VSRecordset
    Dim i As Double
    On Error Resume Next
    Dim ValorTotal As Double
    Dim Conta As New ContaCorrente
    If Trim(txtIM) = "" Then Exit Sub
    Screen.MousePointer = 11
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIM = Imposto.FormataInscricao(txtIM, InscContrib)
    End If
    
    Sql = "Select tci_nome,tci_logradouro,tci_nome_logradouro," & _
    " tci_numero,tci_complemento,tci_bairro,tci_cidade,tci_uf,tci_cgc_cpf FROM tab_Contribuinte where tci_im='" & txtIM & _
    "' and tci_tsc_cod_sit_cad=1"
        
    If Bdados.AbreTabela(Sql, rs) Then
        txtContrib = "" & rs!tci_nome
        txtEndereco = rs!tci_logradouro & " " & rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_BAIRRO & " " & rs!tci_cidade & " " & rs!tci_UF
        'TXTCGC = Rs!TCI_CGC_CPF
    Else
        
        Dim Ic As String
        If Trim(txtIM) <> "" Then
        txtIM = BuscaContribuinte(txtIM, txtContrib, txtEndereco, , etiImovel)
        If Trim(txtIM) = "" Then
            Avisa "Inscricão não encontrada"
            txtIM.SetFocus
        End If
    End If
    End If
    Bdados.FechaTabela rs
    Screen.MousePointer = 0
End Sub

Private Sub txtInicio_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValidade_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

