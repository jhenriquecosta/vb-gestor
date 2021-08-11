VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConfigCubo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' <winapi for getvbpath>
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpdirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpszValueName$, ByVal lpdwRes&, lpdwType&, ByVal lpDataBuff$, nSize&)
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
' </winapi for getvbpath>

Private m_cube As DCube

'*************************************************************************************
' Method: GetNorthwindProductSalesCube()
' Author: Data Dynamics
' Parameters: (DCube) dc is a reference to the DynamiCube object to which the user is viewing.
'             (Boolean) collapsedByDefault tells us whether or not to collapse the cube after it is populated.
' Returns: N/A
' Description: This method is used to set up cubes in the samples directory for generic analysis
' Example:
'       Dim nwindHelper As New DynamiCubeNorthwind
'       nwindHelper.GetNorthwindProductSalesCube(DCube1)
'       Set nwindHelper = Nothing
'*************************************************************************************

Public Sub IniciaCubo(ByVal dc As DynamiCubeLibCtl.DCube, Optional ByVal collapsedByDefault As Boolean = True)
Dim sqlString As String
Dim fieldObject As DynamiCubeLibCtl.Field
Dim looper As Integer

    ' we set an internal reference to the cube so that
    ' the helper methods can use the same cube object
    ' without requiring us to pass it around internally
    Set m_cube = dc
    
    ' delete any existing fields since we want to
    ' start a new cube from scratch. any data that
    ' might have been in the cube up to now will be lost
    m_cube.Fields.DeleteAll
    
    ' color the cube to ensure that it
    ' has all the default settings
    ' as defined by the data dynamics.
    Call FormatCube(dc)
    
    With m_cube
        
        ' this property is turned off (set to false) so that
        ' as the fields are added to the cube, the cube does
        ' not try to do any internal refreshes. leaving the
        ' AutoDataRefresh property to true causes problems in
        ' some cases. it is good practice to turn this off unil
        ' you are finished populating the cube
        .AutoDataRefresh = False
        
        ' tell the cube that we are going to be using an ado
        ' connection to the northwind database and provide a
        ' full connection string that points to a valid nwind.mdb file
        
        .DCConnectType = DCCT_ADO
       ' .DCConnect = BdCnn.ConnectionString
       
        '"Provider=SQLOLEDB.1;Data Source=LocalHost;Database=dbf_saturno;user id=sa;password=kabecao"
        
       '  ' create the sql string and do not include any
        ' order by statements or group by statements.
        ' DynamiCube does all the grouping, aggregating, and
        ' sorting internally unless QueryByPass is turned on
        sqlString = "SELECT * FROM VIEW_FINANCIAMENTOS"
        ' give the string to dynamicube
        .DCRecordSource = sqlString
         
        ' add the appropriate fields to the cube, there are a couple
        ' of things to note here also:
        ' 1.) FIELDS ARE CASE SENSITIVE
        ' 2.) Fields should be added in order of 'Row', 'Column', 'Page', then 'Data'
        ' 3.) Remember to call RefreshData
        ' 4.) Always remember to turn AutoDataRefresh back to True
        
        ' 1.) Rows
        Set fieldObject = .Fields.Add("NOME_CAIXA", "Financeira", DCRow)
        fieldObject.VarName = "NomeEmp"
        Set fieldObject = .Fields.Add("NOME_PLANO", "Tipo", DCRow)
        fieldObject.VarName = "NomeImo"
        
        ' 2.) Columns
        Set fieldObject = .Fields.Add("datepart(YY,DATA_VENCIMENTO)", "Ano", DCColumn)
        fieldObject.VarName = "OrdemAno"
        Set fieldObject = .Fields.Add("datepart(MM,DATA_VENCIMENTO)", "Mes", DCColumn)
        fieldObject.VarName = "OrdemMes"
        Set fieldObject = .Fields.Add("datepart(DD,DATA_VENCIMENTO)", "Dia", DCColumn)
        fieldObject.VarName = "OrdemDia"
        
        fieldObject.DataType = 2 ' numeric sort...
                    
        ' 3.) Page
            ' none
        
        ' 4.) Data
        ' the data field are very special due to the fact that they hold numeric
        ' values and are able to have calculations set throughout the cube at any
        ' given point. in DynamiCube 3.0 the VarName property is used extensively.
        ' ****************************************************************************
        ' ***** IT IS ESSENTIAL THAT YOU SET UNIQUE VARNAMEs FOR YOUR DATAFIELDS *****
        ' ****************************************************************************
        Set fieldObject = .Fields.Add("VALOR", "Saldo", DCData)
            fieldObject.NumberFormat = "##,##0.00"
            fieldObject.VarName = "SomaSaldo"
            
            
        ' call a method that will set the GroupFooterExpression property
        ' for all dimension field / datafield intersections. this will
        ' ensure that the average is always correct on every point in the cube
        'SetFooterExpressions fieldObject, "groupsum(sumProductSales)/groupsum(cntProductID)"
        
        ' this method is used to populate the cube
        ' with the values from the database. it will
        ' also call the FetchData event on unbound cubes
        .RefreshData
        
        ' be sure to set this property back to true
        ' this is needed for setting captions, colors,
        ' orientation changes, aggregate func, etc...
        ' it is used extensively on the internals of dynamicube
        .AutoDataRefresh = True
    End With

    ' fix dataitems
    'Call QuarterStrings("OrderQuarter")
    Call MonthStrings("OrdemMes")
    
    ' collapse all fields
    If collapsedByDefault Then
        For looper = 0 To m_cube.Fields.Count - 1
            m_cube.Fields(looper).ShowDetail False
         Next looper
    End If
    
End Sub

Public Sub CuboDeCustos(ByVal dc As DynamiCubeLibCtl.DCube, Optional ByVal collapsedByDefault As Boolean = True, Optional Sql As String)
Dim sqlString As String
Dim fieldObject As DynamiCubeLibCtl.Field
Dim looper As Integer

    ' we set an internal reference to the cube so that
    ' the helper methods can use the same cube object
    ' without requiring us to pass it around internally
    Set m_cube = dc
    
    ' delete any existing fields since we want to
    ' start a new cube from scratch. any data that
    ' might have been in the cube up to now will be lost
    m_cube.Fields.DeleteAll
    
    ' color the cube to ensure that it
    ' has all the default settings
    ' as defined by the data dynamics.
    Call FormatCube(dc)
    
    With m_cube
        
        ' this property is turned off (set to false) so that
        ' as the fields are added to the cube, the cube does
        ' not try to do any internal refreshes. leaving the
        ' AutoDataRefresh property to true causes problems in
        ' some cases. it is good practice to turn this off unil
        ' you are finished populating the cube
        .AutoDataRefresh = False
        
        ' tell the cube that we are going to be using an ado
        ' connection to the northwind database and provide a
        ' full connection string that points to a valid nwind.mdb file
        
        .DCConnectType = DCCT_ADO
    '    .DCConnect = BdCnn.ConnectionString
       
        '"Provider=SQLOLEDB.1;Data Source=LocalHost;Database=dbf_saturno;user id=sa;password=kabecao"
        
       '  ' create the sql string and do not include any
        ' order by statements or group by statements.
        ' DynamiCube does all the grouping, aggregating, and
        ' sorting internally unless QueryByPass is turned on
        
      '  sqlString = "SELECT NOME_CUSTO,NOME_PLANO,VALOR,DATA FROM VIEW_ANALISE_DE_CUSTOS WHERE " & IIf(Sql = "", " MONTH(DATA)=" & Month(strData_Atual) & " AND YEAR(Data)=" & Year(strData_Atual), Sql)
        
        ' give the string to dynamicube
        .DCRecordSource = sqlString
         
        ' add the appropriate fields to the cube, there are a couple
        ' of things to note here also:
        ' 1.) FIELDS ARE CASE SENSITIVE
        ' 2.) Fields should be added in order of 'Row', 'Column', 'Page', then 'Data'
        ' 3.) Remember to call RefreshData
        ' 4.) Always remember to turn AutoDataRefresh back to True
        
        ' 1.) Rows
        Set fieldObject = .Fields.Add("NOME_CUSTO", "Custo", DCRow)
        fieldObject.VarName = "NomeCusto"
        Set fieldObject = .Fields.Add("NOME_PLANO", "Plano de Conta", DCRow)
        fieldObject.VarName = "NomePlano"
        
        ' 2.) Columns
        Set fieldObject = .Fields.Add("datepart(YY,DATA)", "Ano", DCColumn)
        fieldObject.VarName = "OrdemAno"
        Set fieldObject = .Fields.Add("datepart(MM,DATA)", "Mes", DCColumn)
        fieldObject.VarName = "OrdemMes"
        Set fieldObject = .Fields.Add("datepart(DD,DATA)", "Dia", DCColumn)
        fieldObject.VarName = "OrdemDia"
        
        fieldObject.DataType = 1 ' numeric sort...
                    
        ' 3.) Page
            ' none
        
        ' 4.) Data
        ' the data field are very special due to the fact that they hold numeric
        ' values and are able to have calculations set throughout the cube at any
        ' given point. in DynamiCube 3.0 the VarName property is used extensively.
        ' ****************************************************************************
        ' ***** IT IS ESSENTIAL THAT YOU SET UNIQUE VARNAMEs FOR YOUR DATAFIELDS *****
        ' ****************************************************************************
        Set fieldObject = .Fields.Add("VALOR", "Pago", DCData)
            fieldObject.NumberFormat = "##,##0.00"
            fieldObject.VarName = "SomaSaldo"
            
            
        ' call a method that will set the GroupFooterExpression property
        ' for all dimension field / datafield intersections. this will
        ' ensure that the average is always correct on every point in the cube
        'SetFooterExpressions fieldObject, "groupsum(sumProductSales)/groupsum(cntProductID)"
        
        ' this method is used to populate the cube
        ' with the values from the database. it will
        ' also call the FetchData event on unbound cubes
        .RefreshData
        
        ' be sure to set this property back to true
        ' this is needed for setting captions, colors,
        ' orientation changes, aggregate func, etc...
        ' it is used extensively on the internals of dynamicube
        .AutoDataRefresh = True
    End With

    ' fix dataitems
    'Call QuarterStrings("OrderQuarter")
    Call MonthStrings("OrdemMes")
    
    ' collapse all fields
    If collapsedByDefault Then
        For looper = 0 To m_cube.Fields.Count - 1
            m_cube.Fields(looper).ShowDetail False
         Next looper
    End If
    
End Sub


Public Sub CuboResumoMensalSocio(ByVal dc As DynamiCubeLibCtl.DCube, Optional ByVal collapsedByDefault As Boolean = True, Optional Sql As String)
Dim sqlString As String
Dim fieldObject As DynamiCubeLibCtl.Field
Dim looper As Integer

    ' we set an internal reference to the cube so that
    ' the helper methods can use the same cube object
    ' without requiring us to pass it around internally
    Set m_cube = dc
    
    ' delete any existing fields since we want to
    ' start a new cube from scratch. any data that
    ' might have been in the cube up to now will be lost
    m_cube.Fields.DeleteAll
    
    ' color the cube to ensure that it
    ' has all the default settings
    ' as defined by the data dynamics.
    m_cube.HeaderCaption = "Resumo Mensal de Sócio"
 
    Call FormatCube(dc)
    
    With m_cube
        
        ' this property is turned off (set to false) so that
        ' as the fields are added to the cube, the cube does
        ' not try to do any internal refreshes. leaving the
        ' AutoDataRefresh property to true causes problems in
        ' some cases. it is good practice to turn this off unil
        ' you are finished populating the cube
        .AutoDataRefresh = False
        
        ' tell the cube that we are going to be using an ado
        ' connection to the northwind database and provide a
        ' full connection string that points to a valid nwind.mdb file
        
        .DCConnectType = DCCT_ADO
     '   .DCConnect = BdCnn.ConnectionString
       
        '"Provider=SQLOLEDB.1;Data Source=LocalHost;Database=dbf_saturno;user id=sa;password=kabecao"
        
       '  ' create the sql string and do not include any
        ' order by statements or group by statements.
        ' DynamiCube does all the grouping, aggregating, and
        ' sorting internally unless QueryByPass is turned on
        
       ' sqlString = "SELECT NOME_SOCIO,DATA,VALOR_SOCIO FROM VIEW_FN_SOCIOS_LUCROS WHERE " & IIf(Sql = "", " MONTH(DATA)=" & Month(strData_Atual) & " AND YEAR(Data)=" & Year(strData_Atual), Sql)
      '  sqlString = sqlString & " ORDER BY DATA"
        ' give the string to dynamicube
        .DCRecordSource = sqlString
         
        ' add the appropriate fields to the cube, there are a couple
        ' of things to note here also:
        ' 1.) FIELDS ARE CASE SENSITIVE
        ' 2.) Fields should be added in order of 'Row', 'Column', 'Page', then 'Data'
        ' 3.) Remember to call RefreshData
        ' 4.) Always remember to turn AutoDataRefresh back to True
        
        ' 1.) Rows
        Set fieldObject = .Fields.Add("NOME_SOCIO", "Socio", DCRow)
        fieldObject.VarName = "NomeSocio"
 
        ' 2.) Columns
        Set fieldObject = .Fields.Add("datepart(YY,DATA)", "Ano", DCColumn)
        fieldObject.VarName = "OrdemAno"
        fieldObject.DataType = 2
        
        Set fieldObject = .Fields.Add("datepart(MM,DATA)", "Mes", DCColumn)
        fieldObject.VarName = "OrdemMes"
        fieldObject.DataType = 2
        
        Set fieldObject = .Fields.Add("datepart(DD,DATA)", "Dia", DCColumn)
        fieldObject.VarName = "OrdemDia"
        
        fieldObject.DataType = 2 ' numeric sort...
        'fieldObject.DataType
                    
        ' 3.) Page
            ' none
        
        ' 4.) Data
        ' the data field are very special due to the fact that they hold numeric
        ' values and are able to have calculations set throughout the cube at any
        ' given point. in DynamiCube 3.0 the VarName property is used extensively.
        ' ****************************************************************************
        ' ***** IT IS ESSENTIAL THAT YOU SET UNIQUE VARNAMEs FOR YOUR DATAFIELDS *****
        ' ****************************************************************************
        Set fieldObject = .Fields.Add("VALOR_SOCIO", "Total", DCData)
            fieldObject.NumberFormat = "##,##0.00"
            fieldObject.VarName = "Total"
            
            
        ' call a method that will set the GroupFooterExpression property
        ' for all dimension field / datafield intersections. this will
        ' ensure that the average is always correct on every point in the cube
        'SetFooterExpressions fieldObject, "groupsum(sumProductSales)/groupsum(cntProductID)"
        
        ' this method is used to populate the cube
        ' with the values from the database. it will
        ' also call the FetchData event on unbound cubes
        .RefreshData
        
        ' be sure to set this property back to true
        ' this is needed for setting captions, colors,
        ' orientation changes, aggregate func, etc...
        ' it is used extensively on the internals of dynamicube
        .AutoDataRefresh = True
    End With

    ' fix dataitems
    'Call QuarterStrings("OrderQuarter")
    Call MonthStrings("OrdemMes")
    
    ' collapse all fields
    If collapsedByDefault Then
        For looper = 0 To m_cube.Fields.Count - 1
            m_cube.Fields(looper).ShowDetail False
         Next looper
    End If
    
End Sub
'*************************************************************************************
' Method: SetFooterExpressions()
' Author: Data Dynamics
' Parameters: (Field) dataField is a reference to the DataField containing the VarName
'       that is used to index all the GroupFooterType(n) and GroupFooterExpression(n) properties
'             (String) expression is the expression that gets set for the indexed properties
' Returns: N/A
' Description: This method will set all the GroupFooterExpressions for every dimension field.
'       This will ensure that the value displayed in the total area is always valid
' Example: SetFooterExpressions(calculatedDataField, "groupsum(sumField)/groupsum(countField)")
'*************************************************************************************
Private Sub SetFooterExpressions(ByVal dataField As DynamiCubeLibCtl.Field, ByVal expression As String)
Dim eachField As DynamiCubeLibCtl.Field

    ' loop through all dimension fields (page, column, row) and
    ' set the groupfootertype to calculated then set the associating
    ' groupfooterexpression to the parameter that was passed in
    For Each eachField In m_cube.Fields
        If eachField.Orientation <> DCData Then
            eachField.GroupFooterType(dataField.VarName) = DCFCalculated
            eachField.GroupFooterExpression(dataField.VarName) = expression
        End If
    Next eachField
    
End Sub

'*************************************************************************************
' Method: QuarterStrings()
' Author: Data Dynamics
' Parameters: (String) fieldVarName is the VarName of the field in which we wish to change
' Returns: N/A
' Description: This method is used to set the Quarter strings. This method allows you
'       to pass a Quarter field's varname and display pretty captions
'       1, 2, 3, and 4 become "1st Quarter", "2nd Quarter", "3rd Quarter", and "4th Quarter"
' Example: QuarterStrings("OrderQuarter")
'*************************************************************************************

'*************************************************************************************
' Method: MonthStrings()
' Author: Data Dynamics
' Parameters: (String) fieldVarName identifies the field by its unique varname property
' Returns: N/A
' Description: This method will change the dataitem captions to be localized month names
'       1, 2, 3 become "January", "February", "March" etc...
' Example: MonthStrings("OrderMonth")
'*************************************************************************************
Private Sub MonthStrings(fieldVarName As String)
Dim monthField As DynamiCubeLibCtl.Field
Dim eachMonth As DynamiCubeLibCtl.DataItem

    ' find the month field by using the varname
    ' this is more readable since the fieldname is huge
    ' due to the fact we have to include datepart,etc...
    Set monthField = m_cube.Fields.ItemByVarName(fieldVarName)
    For Each eachMonth In monthField.DataItems
        eachMonth.Caption = MonthName(eachMonth.Name)
    Next eachMonth
    
End Sub

'*************************************************************************************
' Method: NumString()
' Author: Data Dynamics
' Parameters: (Integer) numberValue holds the actual number that we are converting to a string
' Returns: (String) Returns a numeric representation of a string.
' Description: This method will turn a number into a superscript
'       1, 2, 3, 4 become '1st', '2nd', '3rd', and '4th', etc...
' Example: Debug.Print "This is the " & NumString(3) & " time I have told you"
'*************************************************************************************
Private Function NumString(numberValue As Integer) As String
Dim numberAsString As String
    
    ' turn the number into a string
    ' so that we can do string manipulation
    ' on it with left, right, mid, etc...
    numberAsString = CStr(numberValue)
    If Len(numberAsString) = 2 And Val(Left(numberAsString, 1)) = 1 Then
        
        ' if the string is a double digit number and number is
        ' between 10 and 19 then the number will always end with 'th'
        ' proof: 10th, 11th, 12th, 13th, 14th, 15th, 16th, 17th, 18th and 19th
        NumString = numberAsString + "th"
    
    Else
    
        ' if the number is a single digit number or
        ' the number is a double digit number outside
        ' of the 10's then it follows some different rules
        Select Case Val(Right(numberAsString, 1))
            Case 1
                NumString = numberAsString + "st"
            Case 2
                NumString = numberAsString + "nd"
            Case 3
                NumString = numberAsString + "rd"
            Case Else
                NumString = numberAsString + "th"
        End Select
        
    End If
    
End Function


'*************************************************************************************
' Method: FormatCube()
' Author: Data Dynamics
' Parameters: (DCube) theCube is the DynamiCube object we want to apply color to
' Returns: N/A
' Description: This method colors different areas of DynamiCube
'*************************************************************************************
Public Sub FormatCube(ByVal theCube As DCube)
Dim tempFont As New StdFont
    
    tempFont.Name = "MS Sans Serif"
    tempFont.Size = 8
    
    theCube.BorderStyle = DCFixedSingle
    theCube.RowStyle = DCRaisedLight
    theCube.ColStyle = DCRaisedLight
    theCube.FooterCaption = "- Page &P of  &PP -"
    
    theCube.FieldsBackColor = RGB(128, 128, 128)
    theCube.FieldsForeColor = vbWhite
    theCube.FieldFont = tempFont
    
    theCube.FilteredFieldBackColor = theCube.FieldsBackColor
    theCube.FilteredFieldForeColor = theCube.FieldsForeColor
    theCube.FilteredFieldFont = tempFont
    theCube.FilteredFieldFont.Italic = Not theCube.FieldFont.Italic
    
    theCube.CubeBackColor = SystemColorConstants.vbButtonFace
    theCube.GridStyle = DCGridFlat
    theCube.GridColor = RGB(128, 128, 128)
    theCube.Font = tempFont
    
    theCube.HeadingsBackColor = SystemColorConstants.vbButtonFace
    theCube.HeadingsForeColor = SystemColorConstants.vbWindowText
    theCube.HeadingsFont = tempFont
    
    theCube.TotalsBackColor = &H80000018
    theCube.TotalsForeColor = SystemColorConstants.vbWindowText
    theCube.TotalsFont = tempFont
    
End Sub


