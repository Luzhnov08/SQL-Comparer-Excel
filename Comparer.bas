Attribute VB_Name = "Comparer"
Sub SQLComparer()

Dim Conn As New ADODB.Connection
Dim mrs As New ADODB.Recordset
Dim DBPath As String, sconnect As String, Query As String

Dim t1!
    t1 = Timer
    
    Sheets("Result").Cells.Clear

DBPath = ThisWorkbook.FullName 'workbook path = database

sconnect = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & DBPath & ";HDR=Yes';" 'connection string using table headers

Conn.Open sconnect
    Query = "SELECT TOP 10 * FROM [Table1$] LEFT JOIN [Table2$] ON [Table1$].sum >= [Table2$].sum AND [Table1$].direction Like '%' & [Table2$].buyer & '%' WHERE [Table1$].direction IS NULL OR [Table2$].buyer IS NULL" 'query string
    
    mrs.Open Query, Conn 'open = execute Recordset
        
    Sheets("Result").Range("A1").CopyFromRecordset mrs 'paste data on Result sheet = display result
    mrs.Close 'close Recordset

Conn.Close 'close connection

    t1 = Timer - t1
    Debug.Print "Time SQL Comparer", Round(t1, 3) 'time of calculate

End Sub
