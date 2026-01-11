Attribute VB_Name = "Module2"
Sub CalculateDailyReturns()

    Dim wsPrice As Worksheet
    Dim wsReturns As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long

    Set wsPrice = ThisWorkbook.Sheets("Price_Data")
    Set wsReturns = ThisWorkbook.Sheets("Returns")

    wsReturns.Cells.Clear

    ' Get last row and last column
    lastRow = wsPrice.Cells(wsPrice.Rows.Count, 1).End(xlUp).Row
    lastCol = wsPrice.Cells(2, wsPrice.Columns.Count).End(xlToLeft).Column

    ' Copy ticker row (Row 1)
    wsPrice.Rows(1).Copy wsReturns.Rows(1)
    ' Copy second row (headers) - dates and close
    wsPrice.Rows(2).Copy wsReturns.Rows(2)

    ' Loop through columns
    For c = 1 To lastCol

        ' Identify Close columns by header in row 2
        If LCase(Trim(wsPrice.Cells(2, c).Value)) = "close" Then

            ' Rename header to "Returns"
            wsReturns.Cells(2, c).Value = "Returns"

            ' First return is blank
            wsReturns.Cells(3, c).Value = ""

            ' Find last row for this column
            Dim lastRowCol As Long
            lastRowCol = wsPrice.Cells(wsPrice.Rows.Count, c).End(xlUp).Row

            ' Calculate returns for this Close column
            For r = 4 To lastRowCol
                If IsNumeric(wsPrice.Cells(r, c).Value) _
                   And IsNumeric(wsPrice.Cells(r - 1, c).Value) _
                   And wsPrice.Cells(r - 1, c).Value <> 0 Then

                    wsReturns.Cells(r, c).Value = _
                        (CDbl(wsPrice.Cells(r, c).Value) / _
                         CDbl(wsPrice.Cells(r - 1, c).Value)) - 1
                Else
                    wsReturns.Cells(r, c).Value = ""
                End If
            Next r

        Else
            ' For Date columns, copy the data from Price_Data
            wsPrice.Range(wsPrice.Cells(3, c), wsPrice.Cells(lastRow, c)).Copy _
                wsReturns.Cells(3, c)
        End If

    Next c

    MsgBox "Returns calculated", vbInformation

End Sub






