VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form VSVisualizador 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Titulo"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   Icon            =   "VSVisualizador.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6690
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2370
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CRVIEWERLibCtl.CRViewer Rpt 
      Height          =   4440
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   6405
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "VSVisualizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cam As String

Public Sub Imprimir(Optional Caminho As String)
    Cam = Caminho
    Rpt.ViewReport
    Me.Show
    DoEvents
End Sub

Private Sub Form_Resize()
    Rpt.Top = 0
    Rpt.Left = 0
    Rpt.Height = Me.ScaleHeight
    Rpt.Width = Me.ScaleWidth
End Sub

Private Sub Rpt_PrintButtonClicked(UseDefault As Boolean)
    On Error Resume Next
    
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlPDPrintSetup + cdlPDHidePrintToFile
    CommonDialog1.ShowPrinter
    
    If Err = cdlCancel Then UseDefault = False
End Sub
