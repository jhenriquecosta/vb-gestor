VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TDEC107 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TDEC107"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   5940
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   847
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   3690
         TabIndex        =   1
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         Caption         =   "OK"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.grdVISUAL GrdTaxas 
      Height          =   5535
      Left            =   0
      TabIndex        =   2
      Top             =   660
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   9763
      CorBorda        =   16711680
      Caption         =   "Taxas"
      CorTitulo       =   16711680
      CheckBox        =   -1  'True
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   1138
      Icone           =   "TDEC107.frx":0000
   End
End
Attribute VB_Name = "TDEC107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Pega_taxas()
    Dim i As Integer
    Dim pos As Integer
    String_Taxas = ""
    Total_Taxas = 0
    For i = 1 To Grdtaxas.ListItems.Count
        If Grdtaxas.ListItems(i).Checked Then
            pos = InStr(Grdtaxas.ListItems(i).SubItems(1), "-") - 1
            If String_Taxas = "" Then
                String_Taxas = String_Taxas & " [ " & Left(Grdtaxas.ListItems(i).SubItems(1), pos) & " ]" & " - " & Format(Grdtaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            Else
                String_Taxas = String_Taxas & ", [ " & Left(Grdtaxas.ListItems(i).SubItems(1), pos) & " ]" & " - " & Format(Grdtaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            End If
            Total_Taxas = Total_Taxas + CCur(Grdtaxas.ListItems(i).SubItems(2))
        End If
    Next
End Sub

Private Sub cmdSair_Click()
    Pega_taxas
    Unload Me
End Sub

Private Sub Form_Load()
    Grdtaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
    'Pega_taxas
End Sub
