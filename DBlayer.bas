Attribute VB_Name = "DBlayer"

'''''''''''''''''''''''''''''''''
'''''''''''''##########''''''''''
'''''''''''''#'''''''''''''''''''
'''''''''''''#'''''''''''''''''''
'''''''''##########''''''''''''''
'''''''''#'''#'''''''''''''''''''
'''''''''#'''#'''''''''''''''''''
'''''''''#'''#'''''''''''''''''''
'''''''''#'''#'''''''''''''''''''
'''''''''#'''''''''''''''''''''''
'''''''''#'''''''''''''''''''''''
'''''''''##########''''''''''''''
'''''''''''''''''''''''''''''''''
'Developed by Crins F.'''''''''''
'Contact: job@crins.eu'''''''''''
'Free for commercial use'''''''''
'just copy this header'''''''''''
'''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''
Option Explicit
Const passwd = "Pharaon00"
Const dbname = "Test2DBpw.mdb"
Const dbpath = "C:\temp\" & dbname
Const dbprovider = "Microsoft.Jet.OLEDB.4.0"
Const ConnectionString = "Provider=" & dbprovider & ";" & "Jet OLEDB:Database Password =" & passwd & ";" & "Data Source =" & dbpath & ";"

Sub CreateDB()
Dim Catalog As Object
Set Catalog = CreateObject("ADOX.Catalog")
Catalog.Create ConnectionString

End Sub
Sub Connection()
Dim connect As New ADODB.Connection
connect.Open ConnectionString

End Sub

Sub createtable(tablename As String)
Dim sqlstring As String
Dim tableid As String
Dim connect As New ADODB.Connection
On Error GoTo Handlerror:
    tableid = tablename & "_ID"
    connect.Open ConnectionString
    sqlstring = "CREATE TABLE " & tablename & "(" & tableid & " AUTOINCREMENT PRIMARY KEY)"
    connect.Execute sqlstring
    connect.Close
    
Exit Sub
Handlerror:
    errorhandling Err.number, Err.description, Err.source, "createtable"


End Sub
Sub Addcolumn(tablename As String, columnname As String, columntype As String)
Dim connect As New ADODB.Connection
Dim sqlstring As String
'On Error GoTo Handlerror:

    connect.Open ConnectionString
    sqlstring = "ALTER TABLE " & tablename & " ADD COLUMN " & columnname & " " & columntype
    connect.Execute sqlstring ' & ";"
    'connect.    TableDefs(tablename).Fields(columnname). = defaultvaalue
    connect.Close
    
Exit Sub
Handlerror:
    errorhandling Err.number, Err.description, Err.source, "Addcolumn"


End Sub
Sub modifycolumn(tablename As String, columnname As String, columntype As String)
Dim connect As New ADODB.Connection
Dim sqlstring As String
On Error GoTo Handlerror
    connect.Open ConnectionString
    sqlstring = "ALTER TABLE " & tablename & " ALTER COLUMN " & columnname & " " & columntype
    connect.Execute sqlstring
    connect.Close
    
Exit Sub

Handlerror:
    errorhandling Err.number, Err.description, Err.source, "modifycolumn"
End Sub
Sub droptable(tablename As String)
Dim connect As New ADODB.Connection
Dim sqlstring As String
On Error GoTo Handlerror:
    connect.Open ConnectionString
    sqlstring = "DROP TABLE " & tablename
    connect.Execute sqlstring
    connect.Close
    
Exit Sub

Handlerror:
    errorhandling Err.number, Err.description, Err.source, "droptable"
    Resume Next
End Sub
Sub Addvalues_in_one(tablename As String, columnname() As Variant, vaalue() As Variant)
'add multiple value  to table
'add multiple ROW to table (by single sql INSERT query) do not Work for access
' 1 single dimension array (0 to x) with columname x= column
' 2d array (0 to y, 0 to x) y= new roW FOR ACCESS y = 1 (0 to 0, 0 to x)
Dim connect As New ADODB.Connection
Dim sqlstring, columnamestring, valuestring As String
Dim a, b As Long
On Error GoTo Handlerror

 If UBound(columnname) <> UBound(vaalue, 2) Then
    MsgBox "Wrong Arrays size in addvalue :" & Chr(13) & _
        "  UBound(columnname) <> UBound(vaalue, 1)" & Chr(13) & _
        UBound(columnname) & " <> " & UBound(vaalue, 1)
    Exit Sub
End If
columnamestring = "("
For a = 0 To UBound(columnname)
    If a = UBound(columnname) Then
    columnamestring = columnamestring & columnname(a) & ")"
    Else
    columnamestring = columnamestring & columnname(a) & ", "
    End If
Next a
valuestring = "("
For a = 0 To UBound(vaalue, 1)
    
    For b = 0 To UBound(vaalue, 2)
        If b = UBound(vaalue, 2) And a = UBound(vaalue, 1) Then
            valuestring = valuestring & vaalue(a, b) & ")"
        ElseIf b = UBound(vaalue, 2) And a < UBound(vaalue, 1) Then
            valuestring = valuestring & vaalue(a, b) & "),("
        Else
            valuestring = valuestring & vaalue(a, b) & ", "
        End If
    Next b
Next a
    connect.Open ConnectionString
    
    sqlstring = "INSERT INTO " & tablename & " " & columnamestring & " VALUES " & valuestring & ";"
    connect.Execute sqlstring
    connect.Close
    
Exit Sub

Handlerror:
connect.Close
    errorhandling Err.number, Err.description, Err.source, "Addvalues_in_one"
End Sub
Sub Addvalues_in_mul(tablename As String, columnname() As Variant, vaalue() As Variant)
'add multiple value  to table
'add multiple ROW to table for access (by multiple sql INSERT query)
' 1 single dimension array (0 to x) with columname x= column
' 2d array (0 to y, 0 to x) y= new roW FOR ACCESS y = 1 (0 to 0, 0 to x)
Dim connect As New ADODB.Connection
Dim sqlstring, columnamestring, valuestring As String
Dim a, b As Long
'On Error GoTo Handlerror

 If UBound(columnname) <> UBound(vaalue, 2) Then
    MsgBox "Wrong Arrays size in addvalue :" & Chr(13) & _
        "  UBound(columnname) <> UBound(vaalue, 1)" & Chr(13) & _
        UBound(columnname) & " <> " & UBound(vaalue, 1)
    Exit Sub
End If
columnamestring = "("
For a = 0 To UBound(columnname)
    If a = UBound(columnname) Then
    columnamestring = columnamestring & columnname(a) & ")"
    Else
    columnamestring = columnamestring & columnname(a) & ", "
    End If
Next a
valuestring = "("
For a = 0 To UBound(vaalue, 1)
    
    For b = 0 To UBound(vaalue, 2)
        If b = UBound(vaalue, 2) Then
            valuestring = valuestring & vaalue(a, b) & ")"
            connect.Open ConnectionString
            sqlstring = "INSERT INTO " & tablename & " " & columnamestring & " VALUES " & valuestring & ";"
            connect.Execute sqlstring
            connect.Close
            valuestring = "("
        Else
            valuestring = valuestring & vaalue(a, b) & ", "
        End If
    Next b
Next a

    
    
Exit Sub

Handlerror:
    connect.Close
    errorhandling Err.number, Err.description, Err.source, "Addvalues_in_mul"
End Sub
Sub dbExecute(ByVal sqlstring As String)
Dim connect As New ADODB.Connection
connect.Open ConnectionString
connect.Execute sqlstring
connect.Close
End Sub
Sub viewfield(tablename As String)
'Dim connect As New ADODB.Connection
Dim connect As ADODB.Recordset
Dim sqlstring As String
Dim a As Long
Dim namtyp As String

    connect.Open ConnectionString
    Set record = New ADODB.Recordset
    record.Open tablename, connect
    For a = 0 To record.Fields.Count - 1
        Debug.Print record.Fields(a).Name & " " & record.Fields(a).Attributes & " " & record.Fields(a).Type,
      
        
    Next a

record = Nothing
Exit Sub
Handlerror:
    errorhandling Err.number, Err.description, Err.source, "vieww"
End Sub
Public Function errorhandling(number As Long, description As String, source As String, Optional subname As String)
Dim columnnam(0 To 4) As Variant
Dim valuee(0 To 0, 0 To 4) As Variant
If subname = "" Then
    subname = "Unknown"
End If
description = Replace(description, "'", "_")
 Select Case number
    Case Else
    'Comment this block until msgbox if no errorhandeling table
    columnnam(0) = "Errornumber"
    columnnam(1) = "Error_des"
    columnnam(2) = "Error_source"
    columnnam(3) = "Error_fct"
    columnnam(4) = "Who"

     valuee(0, 0) = "'" & number & "'"
     valuee(0, 1) = "'" & description & "'"
     valuee(0, 2) = "'" & source & "'"
     valuee(0, 3) = "'" & subname & "'"
     valuee(0, 4) = "'" & Application.UserName & "'"
     Addvalues_in_mul "errorhandeling", columnnam(), valuee()
     'Stop Comment this block
     MsgBox "ERROR: " & number & Chr(13) & description & Chr(13) & source & Chr(13) & subname, Title:="Error"
    End Select
    
End Function
Function dataview(tablename As String)


Dim connect As New ADODB.Connection
Dim dataa As ADODB.Recordset

Dim a, b As Long
Dim ab As Boolean
Dim array1() As Variant
    On Error GoTo Handlerror:
connect.Open ConnectionString
Set dataa = connect.Execute("SELECT * FROM  " & tablename & " ;")
 b = 0
 Do While Not dataa.EOF
 b = b + 1
dataa.MoveNext
Loop
ReDim array1(0 To b + 1, 0 To dataa.Fields.Count - 1)
ab = False
dataa.MoveFirst
b = 1
Do While Not dataa.EOF
    If ab = False Then
        For a = 0 To dataa.Fields.Count - 1
        array1(0, a) = dataa.Fields(a).Name
        Next a
        ab = True
    End If
    
    For a = 0 To dataa.Fields.Count - 1
        'For b = O To dataa.RecordCount
        Debug.Print dataa.Fields(a).Name & ", " & dataa.Fields(a).Value & "; ",
        array1(b, a) = dataa.Fields(a).Value
    Next a
    b = b + 1
    Debug.Print ""
    dataa.MoveNext
Loop
dataview = array1()
Exit Function
Handlerror:
    errorhandling Err.number, Err.description, Err.source, "dataview"
End Function
Function reverse2darray(ByRef array1() As Variant)
Dim array2() As Variant
Dim a, b, alenght, blenght As Long
alenght = UBound(array1, 1)
blenght = UBound(array1, 2)
ReDim array2(0 To blenght + 1, 0 To alenght)
For a = 0 To alenght
    For b = 0 To blenght
        array2(b, a) = array1(a, b)
    Next b
Next a
reverse2darray = array2
End Function
Function singlearrayto2darray(ByRef array1() As Variant)
Dim array2() As Variant
Dim a, alenght As Long
alenght = UBound(array1)
ReDim array2(0 To 0 + 1, 0 To alenght)
For a = 0 To alenght
    
        array2(0, a) = array1(a)
    
Next a
singlearrayto2darray = array2
End Function
Function dataview2(tablename As String) As Variant()
Dim connect As New ADODB.Connection
Dim dataa As ADODB.Recordset

Dim a, b As Long
Dim ab As Boolean
Dim array1(), array2(), returnarray() As Variant


   ' On Error GoTo Handlerror:
connect.Open ConnectionString
Set dataa = connect.Execute("SELECT * FROM  " & tablename & " ;")
array1 = dataa.GetRows(dataa.RecordCount)

ReDim array2(0 To dataa.Fields.Count - 1)

For a = 0 To dataa.Fields.Count - 1
    array2(a) = dataa.Fields(a).Name
Next a


connect.Close
ReDim returnarray(0 To 1)
returnarray(0) = array2
returnarray(1) = array1
dataview2 = returnarray
Debug.Print ""
End Function
Function fieldview2(ByVal tablename As String) As Variant()
Dim connect As New ADODB.Connection
Dim dataa As ADODB.Recordset
Dim a As Long
Dim array1() As Variant

   ' On Error GoTo Handlerror:
connect.Open ConnectionString
Set dataa = connect.Execute("SELECT * FROM  " & tablename & " ;")
ReDim array1(0 To dataa.Fields.Count - 1)
For a = 0 To dataa.Fields.Count - 1
    array1(a) = dataa.Fields(a).Name
Next a
connect.Close
fieldview2 = array1
Debug.Print ""
End Function
Sub adotablesfieldview()
  Dim connect As New ADODB.Connection
  Dim tableschema As ADODB.Recordset
  Dim columnschema As ADODB.Recordset
  Dim arr() As Variant
  
    'On Error GoTo Handlerror:
  connect.Open ConnectionString
    Set tableschema = connect.Execute("SELECT * FROM  errorhandeling ")
  
  Set tableschema = connect.OpenSchema(adSchemaTables)
  Do While Not tableschema.EOF
    'Get all table columns.
    
    Set columnschema = connect.OpenSchema(adSchemaColumns, Array(Empty, Empty, "" & tableschema("TABLE_NAME")))
    Do While Not columnschema.EOF
      Debug.Print tableschema("TABLE_NAME") & ", " & tableschema("DATE_MODIFIED"); ", " & _
        columnschema("COLUMN_NAME") & ", " & adotyp(columnschema("DATA_TYPE")) & ", " & columnschema("CHARACTER_MAXIMUM_LENGTH") & ", " & _
        columnschema("COLUMN_HASDEFAULT") & ", " & columnschema("COLUMN_DEFAULT")
        
      columnschema.MoveNext
    Loop
    tableschema.MoveNext
  Loop


  Exit Sub
Handlerror:
    errorhandling Err.number, Err.description, Err.source, "adotablesfieldview"
End Sub


Public Function adotyp(num As Long)
Select Case num
Case Is = 20
 adotyp = "adBigInt"
Case Is = 128
 adotyp = "adBinary"
Case Is = 11
 adotyp = "adBoolean"
Case Is = 8
 adotyp = "adBSTR"
Case Is = 136
 adotyp = "adChapter"
Case Is = 129
 adotyp = "adChar"
Case Is = 6
 adotyp = "adCurrency"
Case Is = 7
 adotyp = "adDate"
Case Is = 133
 adotyp = "adDBDate"
Case Is = 134
 adotyp = "adDBTime"
Case Is = 135
 adotyp = "adDBTimeStamp"
Case Is = 14
 adotyp = "adDecimal"
Case Is = 5
 adotyp = "adDouble"
Case Is = 0
 adotyp = "adEmpty"
Case Is = 10
 adotyp = "adError"
Case Is = 64
 adotyp = "adFileTime"
Case Is = 72
 adotyp = "adGUID"
Case Is = 9
 adotyp = "adIDispatch"

Case Is = 3
 adotyp = "adInteger"
Case Is = 13
 adotyp = "adIUnknown"

Case Is = 205
 adotyp = "adLongVarBinary"
Case Is = 201
 adotyp = "adLongVarChar"
Case Is = 203
 adotyp = "adLongVarWChar"
Case Is = 131
 adotyp = "adNumeric"
Case Is = 138
 adotyp = "adPropVariant"
Case Is = 4
 adotyp = "adSingle"
Case Is = 2
 adotyp = "adSmallInt"
Case Is = 16
 adotyp = "adTinyInt"
Case Is = 21
 adotyp = "adUnsignedBigInt"
Case Is = 19
 adotyp = "adUnsignedInt"
Case Is = 18
 adotyp = "adUnsignedSmallInt"
Case Is = 17
 adotyp = "adUnsignedTinyInt "
Case Is = 132
 adotyp = "adUserDefined"
Case Is = 204
 adotyp = "adVarBinary"
Case Is = 200
 adotyp = "adVarChar"
Case Is = 12
 adotyp = "adVariant"

Case Is = 139
 adotyp = "adVarNumeric"
Case Is = 202
 adotyp = "adVarWChar"
Case Is = 130
 adotyp = "adWChar"

 Case Else
 adotyp = "Error"
End Select
End Function
'Public Function daotyp(num As Long)
'
'Select Case num
'Case Is = 101
' daotyp = "dbAttachment"
'Case Is = 16
' daotyp = "dbBigInt"
'Case Is = 9
' daotyp = "dbBinary"
'Case Is = 1
' daotyp = "dbBoolean"
'Case Is = 2
' daotyp = "dbByte"
'Case Is = 18
' daotyp = "dbChar"
'Case Is = 102
' daotyp = "dbComplexByte"
'Case Is = 108
' daotyp = "dbComplexDecimal"
'Case Is = 106
' daotyp = "dbComplexDouble"
'Case Is = 107
' daotyp = "dbComplexGUID"
'Case Is = 103
' daotyp = "dbComplexInteger"
'Case Is = 104
' daotyp = "dbComplexLong"
'Case Is = 105
' daotyp = "dbComplexSingle"
'Case Is = 109
' daotyp = "dbComplexText"
'Case Is = 5
' daotyp = "dbCurrency"
'Case Is = 8
' daotyp = "dbDate"
'Case Is = 20
' daotyp = "dbDecimal"
'
'
'Case Is = 7
' daotyp = "dbDouble"
'Case Is = 21
' daotyp = "dbFloat"
'
'
'Case Is = 15
' daotyp = "dbGUID"
'Case Is = 3
' daotyp = "dbInteger"
'Case Is = 4
' daotyp = "dbLong"
'Case Is = 11
' daotyp = "dbLongBinary"
'Case Is = 12
' daotyp = "dbMemo"
'Case Is = 19
' daotyp = "dbNumeric"
'Case Is = 6
' daotyp = "dbSingle"
'Case Is = 10
' daotyp = "dbText"
'Case Is = 22
' daotyp = "dbTime"
'
'
'Case Is = 23
' daotyp = "dbTimeStamp"
'
'
'Case Is = 17
' daotyp = "dbVarBinary"
'Case Is = 130
' daotyp = "adWChar  "
' Case Else
' daotyp = "Error"
'End Select
'End Function
