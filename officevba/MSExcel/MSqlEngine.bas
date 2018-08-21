Public Sub SQLEngine(ByVal strExpression As String, _
Optional ByVal sWksName As String, Optional sWksTblName As String, Optional ByVal reportType As enuReportTemplate)
If sWksName = "" Then sWksName = "GENERIC_" & Int((Format(Now, "SS") + 1) * Rnd + 1000)
If sWksTblName = "" Then sWksTblName = "GENERIC_" & Int((Format(Now, "SS") + 1) * Rnd + 1000)
strExpression = QueryStringBuilder(strExpression)

Dim oConnection As Object
Dim oRecordSet As Object
Set oConnection = Nothing
Set oRecordSet = Nothing

Set oConnection = CreateObject("ADODB.Connection")
Set oRecordSet = CreateObject("ADODB.Recordset")

Dim sConnectionType As String
sConnectionType = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ActiveWorkbook.FullName _
                & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
On Err GoTo errhandler:
oConnection.Open sConnectionType
On Error GoTo errhandler:
oRecordSet.Open strExpression, oConnection

    If oRecordSet.EOF Then
    ufSqlParser.tbSQLLog.Text = "Unable to Query Data: EOF"
    Exit Sub
    
    Else
    
    Dim wks As Worksheet
    Set wks = ActiveWorkbook.Worksheets.Add
        wks.Name = sWksName & "_" & Int((Format(Now, "SS") + 1) * Rnd + 1000)
        wks.ListObjects.Add(xlSrcRange, Range("A1"), , xlYes).Name = sWksTblName
        wks.ListObjects(sWksTblName).TableStyle = ""
        
        Dim iColumnIndex As Integer
    
        For iColumnIndex = 0 To oRecordSet.Fields.Count - 1
        
        wks.Cells(1, iColumnIndex + 1) = _
        oRecordSet.Fields(iColumnIndex).Name
        
       FormatABColumn oRecordSet.Fields(iColumnIndex).Name, wks, iColumnIndex + 1
        
        Next
        wks.Range("A2").CopyFromRecordset oRecordSet
    End If
   ufSqlParser.tbSQLLog.Text = "OK!"
errhandler:
    If Err <> 0 Then
    ufSqlParser.tbSQLLog.Text = Err & Chr(13) & Error(Err)
    'oRecordSet.Close
    oConnection.Close
    Exit Sub
    End If
oRecordSet.Close
oConnection.Close
Set oConnection = Nothing
Set oRecordSet = Nothing
End Sub