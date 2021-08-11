VERSION 5.00
Object = "{6D63F73C-3688-3000-9C0F-00A0C90F29FC}#3.0#0"; "DCube3.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form formObrigAnalise 
   BackColor       =   &H00FBEDE8&
   Caption         =   "ANALISE MENSAL DE ARRECADAÇÃO"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18075
   Icon            =   "formObrigAnalise.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   18075
   StartUpPosition =   2  'CenterScreen
   Begin DynamiCubeLibCtl.DCube DCube1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   13996
      DataSource      =   ""
      BeginProperty HeadingsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FieldFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowAlignment    =   0
      ColAlignment    =   0
      RowStyle        =   1
      ColStyle        =   1
      OutlineIconAlignment=   1
      GridColor       =   12632256
      BackColor       =   16510440
      DCConnect       =   ""
      DCDatabaseName  =   ""
      CursorStyle     =   0
      FieldsBackColor =   8421504
      FieldsForeColor =   16777215
      HeadingsForeColor=   0
      HeadingsBackColor=   16777215
      DCRecordSource  =   ""
      TotalsBackColor =   16777215
      TotalsForeColor =   0
      BeginProperty TotalsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridStyle       =   1
      ForeColor       =   0
      AllowFiltering  =   -1  'True
      AllowUserPivotFields=   -1  'True
      LeftMargin      =   0,75
      RightMargin     =   0,75
      TopMargin       =   0,49
      BottomMargin    =   0,49
      HeaderMargin    =   0,49
      FooterMargin    =   0,49
      FooterCaption   =   "- Page &P -"
      HeaderCaption   =   "DynamiCube"
      HeaderJustification=   1
      FooterJustification=   1
      ColPageBreak    =   0
      RowPageBreak    =   0
      ColHeadingsOnEveryPage=   -1  'True
      RowHeadingsOnEveryPage=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DCOptions       =   0
      AutoDataRefresh =   -1  'True
      PrinterColumnSpacing=   0,01
      DCConnectType   =   5
      DCQueryTimeOut  =   0
      SQLYearPart     =   "datepart(""yyyy"",<field>)"
      SQLQuarterPart  =   "datepart(""q"",<field>)"
      SQLMonthPart    =   "datepart(""m"",<field>)"
      SQLWeekPart     =   "datepart(""ww"",<field>)"
      BorderStyle     =   1
      AllowSplitters  =   -1  'True
      QueryByPass     =   0   'False
      DataPath        =   ""
      DataNotAvailableCaption=   ""
      PageFieldsVisible=   -1  'True
      printerJobName  =   "DynamiCube Output"
      Fields          =   "formObrigAnalise.frx":08CA
      CubeBackColor   =   13160660
      BeginProperty FooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FooterBackColor =   -1
      FooterForeColor =   0
      HeaderBackColor =   -1
      HeaderForeColor =   0
      BeginProperty FilteredFieldFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FilteredFieldBackColor=   -1
      FilteredFieldForeColor=   16777215
      MousePointer    =   0
      LoadProgressNotifyDelay=   1000
      IncludeColorsInPrintout=   -1  'True
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   18075
      _ExtentX        =   31882
      _ExtentY        =   1138
      Formulario      =   "Arrecadação Mensal"
      Descricao       =   "Analise Mensal"
      Icone           =   "formObrigAnalise.frx":0926
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   8535
      Width           =   18075
      _ExtentX        =   31882
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   16920
         TabIndex        =   3
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777152
      End
   End
End
Attribute VB_Name = "formObrigAnalise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_IsProcessing As Boolean
Private m_numberFormat() As String
Dim cConfCubo As New cConfigCubo
Public sqlCubo As String

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    'allow user to view Print Preview
    DCube1.PrintPreview
End Sub

'*************************************************************************************
' Method: GetDataField()
' Author: Data Dynamics
' Parameters: N/A
' Returns: (DynamiCubeLibCtl.Field) returns datafield that user selected in cboFields
' Description: This method uses the index of the item from cboFields to retrieve the
'              DataField associated with that index
' Example: Dim df as DynamiCubeLibCtl.Field
'          set df = GetDataField

'*************************************************************************************



'*************************************************************************************
' Method: IndexOf()
' Author: Data Dynamics
' Parameters: (VB.ComboBox) theCombo is the combo box we are going to set the index for
'             (Long) itemData is the item we are looking for in theCombo
' Returns: N/A
' Description: This method looks through the list items in theCombo trying to find a
'              match with itemData.  If a match is found, itemData is selected in theCombo
' Example:  Call IndexOf(cboAggregateFuncs,  DCCount)
'*************************************************************************************

Private Sub DCube1_BeforeQuery(Sql As String)
    'allow the user to see the query we will send to the database
    ' -- you can use this to customize your query to your database engine --
   ' Debug.Print Sql
End Sub



'*************************************************************************************
' Method: InitializeNumberFormats()
' Author: Data Dynamics
' Parameters: N/A
' Returns: N/A
' Description: This method caches the number formatting for each field.
'*************************************************************************************
Private Sub InitializeNumberFormats()
Dim looper As Long

    ReDim m_numberFormat(DCube1.DataFields.Count - 1) As String
    
    'cache the number formatting for each field
    For looper = 0 To DCube1.DataFields.Count - 1
        m_numberFormat(looper) = DCube1.DataFields(looper).NumberFormat
    Next
    
End Sub


Private Sub Form_Load()
   
   cConfCubo.CuboDaArrecadacao DCube1, True, sqlCubo
   Call InitializeNumberFormats

End Sub

Private Sub Form_Resize()

    Dim topOfCube As Long
    Dim leftOfCube As Long
    Const spacer As Long = 90
    
    On Error Resume Next
    
    'resize DynamiCube and other controls elegantly
    'TopPicture.Move spacer, spacer, ScaleWidth - (spacer * 2)
    'topOfCube = TopPicture.Top + TopPicture.Height + spacer
   ' LeftPicture.Move spacer, topOfCube, LeftPicture.Width, ScaleHeight - (spacer + topOfCube)
   ' leftOfCube = LeftPicture.Left + LeftPicture.Width + spacer
   ' DCube1.Move leftOfCube, topOfCube, ScaleWidth - (leftOfCube + spacer), ScaleHeight - (topOfCube + spacer)
    
    
End Sub

