VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TAID501 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAID501"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7440
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TAID501.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.txtVISUAL txtBusca 
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   750
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   503
      Caption         =   "Nome"
      Text            =   ""
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   5265
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   360
         Left            =   5130
         TabIndex        =   3
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         Caption         =   "Buscar"
         Acao            =   5
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6300
         TabIndex        =   4
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.grdVISUAL lstObrig 
      Height          =   4140
      Left            =   45
      TabIndex        =   1
      Top             =   1095
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   7303
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   1138
      Formulario      =   "Visão de Rapida de Obrigações..."
      Descricao       =   "Realiza uma visão rapida das obrigações do contribuinte..."
      Icone           =   "TAID501.frx":2123
   End
End
Attribute VB_Name = "TAID501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NotaAidf As New cNotaAidf

Private Sub cmdBuscar_Click()
Dim Obrig As Obrigacao
    Dim Conta As New ContaCorrente
    'Dim Inscri As String
    Dim Sql As String
    Set Obrig = New Obrigacao
    
    If txtBusca = "" Then
        Util.Avisa "Informe contribuinte."
        txtBusca.SetFocus
        Exit Sub
    End If
    
    If Me.Tag = TAID201.Name Then
        
    
            Sql = " SELECT TOC_COD_OBRIGACAO as Documento,TOC_INSCRICAO as INSCRICAO,VIN_RAZAO as [Razão Social],"
            Sql = Sql & " TIP_SIGLA_IMPOSTO AS TRIBUTO, TOC_PERIODO AS PERIODO,TOC_DATA_VENCIMENTO AS VENCIMENTO,"
            Sql = Sql & " TOC_VALOR_OBRIGACAO AS VALOR ,TOC_VALOR_JUROS AS JUROS,"
            Sql = Sql & " TOC_VALOR_MULTA as MULTA,TGE_NOME as SIT,TOC_TOTAL_TAXA_INCLUSA as TAXA,"
            Sql = Sql & " TOC_TIP_COD_IMPOSTO as IMPOSTO,TOC_PARCELA AS PARCELA,TOC_OBSERVACAO AS OBSERVACAO,"
            Sql = Sql & " VIN_DOCUMENTO,TOC_NUM_DOC_ORIGEM AS [DOC ORIGEM]"
            Sql = Sql & " From TAB_OBRIGACAO_CONTRIBUINTE, TAB_IMPOSTO, VIS_STATUS_OBRIGACAO, VIS_INSCRICAO"
            Sql = Sql & " Where TOC_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
            Sql = Sql & " AND TOC_INSCRICAO = VIN_INSCRICAO"
            Sql = Sql & " AND  TOC_STATUS_OBRIGACAO = TGE_CODIGO"
            Sql = Sql & " AND TOC_STATUS_OBRIGACAO in (1,2)"
            Sql = Sql & " AND VIN_INSCRICAO = " & Bdados.Converte(Inscri, tctexto) & "  AND TIP_SIGLA_IMPOSTO LIKE '%" & txtBusca & "%'  ORDER BY TOC_PERIODO"
            
            If Trim(Inscri) <> "" Then Conta.ExecutaAtualizacao Inscri
            lstObrig.Preencher Bdados, Sql
    Else
            
            With NotaAidf
                
                If .PreencherGrid(lstObrig, , , , txtBusca, , , , CStr(TipoConsulta)) = False Then
                Avisa "Nenhuma nota encontrada."
                'txtContribuinte.SetFocus
        End If
    End With
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim Obrig As Obrigacao
    Dim Conta As New ContaCorrente
    'Dim Inscri As String
    Dim Sql As String
    Set Obrig = New Obrigacao
    
    If Me.Tag = TAID201.Name Then
            txtBusca.Caption = "Informe a sigla do tributo"
            Sql = " SELECT TOC_COD_OBRIGACAO as Documento,TOC_INSCRICAO as INSCRICAO,VIN_RAZAO as [Razão Social],"
            Sql = Sql & " TIP_SIGLA_IMPOSTO AS TRIBUTO, TOC_PERIODO AS PERIODO,TOC_DATA_VENCIMENTO AS VENCIMENTO,"
            Sql = Sql & " TOC_VALOR_OBRIGACAO AS VALOR ,TOC_VALOR_JUROS AS JUROS,"
            Sql = Sql & " TOC_VALOR_MULTA as MULTA,TGE_NOME as SIT,TOC_TOTAL_TAXA_INCLUSA as TAXA,"
            Sql = Sql & " TOC_TIP_COD_IMPOSTO as IMPOSTO,TOC_PARCELA AS PARCELA,TOC_OBSERVACAO AS OBSERVACAO,"
            Sql = Sql & " VIN_DOCUMENTO,TOC_NUM_DOC_ORIGEM AS [DOC ORIGEM]"
            Sql = Sql & " From TAB_OBRIGACAO_CONTRIBUINTE, TAB_IMPOSTO, VIS_STATUS_OBRIGACAO, VIS_INSCRICAO"
            Sql = Sql & " Where TOC_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
            Sql = Sql & " AND TOC_INSCRICAO = VIN_INSCRICAO"
            Sql = Sql & " AND  TOC_STATUS_OBRIGACAO = TGE_CODIGO"
            Sql = Sql & " AND TOC_STATUS_OBRIGACAO in (1,2)"
            Sql = Sql & " AND VIN_INSCRICAO = " & Bdados.Converte(Inscri, tctexto) & " ORDER BY TOC_PERIODO"
            
            If Trim(Inscri) <> "" Then Conta.ExecutaAtualizacao Inscri
            lstObrig.Preencher Bdados, Sql
    Else
            cabVisual.Formulario = "Visão rapida de AIDF"
            cabVisual.Descricao = "Visão rapida de AIDF"
            With NotaAidf
                
                'If .PreencherGrid(lstObrig) = False Then
                'Avisa "Nenhuma nota encontrada."
                'txtContribuinte.SetFocus
        'End If
    End With
    End If
End Sub

Private Sub lstObrig_DblClick()
    If Me.Tag = TAID202.Name Then
        Inscri = lstObrig.SelectedItem
        Unload Me
    End If
End Sub

