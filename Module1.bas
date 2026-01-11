Attribute VB_Name = "Module1"
Sub ImportPriceData()

    Dim wsTarget As Worksheet
    Dim wsControl As Worksheet
    Dim filePath As String
    
    Set wsTarget = ThisWorkbook.Sheets("Price_Data")
    Set wsControl = ThisWorkbook.Sheets("Portfolio_Input")
    
    filePath = wsControl.Range("B1").Value
    
    If filePath = "" Then
        MsgBox "Please enter the CSV file name or path in B1", vbExclamation
        Exit Sub
    End If
    
    If Dir(filePath) = "" Then
        filePath = ThisWorkbook.Path & "/" & filePath
    End If

    wsTarget.Cells.Clear
    
    With wsTarget.QueryTables.Add(Connection:="TEXT;" & filePath, _
                                  Destination:=wsTarget.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = Array(xlDMYFormat, xlGeneralFormat, xlDMYFormat, xlGeneralFormat, xlDMYFormat, xlGeneralFormat)
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    MsgBox "CSV imported into Price_Data successfully!", vbInformation

End Sub

