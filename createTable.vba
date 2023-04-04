Sub CreateTable()
    Dim SourceSheet As Worksheet
    Dim TargetSheet As Worksheet
    Dim FoundPageFlow As Range
    Dim FoundWindows As Range
    Dim FoundMacOS As Range
    Dim FoundAndroid As Range
    Dim FoundIOS As Range
    Dim LastRow As Long
    Dim i As Long
    Dim NextColumn As Integer
    Dim HeaderRange As Range
    Dim tableRange As Range
    Dim ListObj As ListObject
    Dim colLetter As String
    Dim rng As Range
    Dim cell As Range
    Dim tablestartrow As Integer
    Dim tableendrow As Integer
    Dim tablestartcol As Integer
    Dim tableendcol As Integer
    Dim ws As Worksheet
    Dim startRow As Long
    startRow = 165 ' Starting row for the new table
    

    
    Set SourceSheet = ThisWorkbook.Worksheets("Execution Summary")
    Set TargetSheet = ThisWorkbook.Worksheets("Data & Chart") ' Replace with the name of the sheet where you want to create the table
    Set ws = TargetSheet
    
    ws.Range("K:AE").Delete xlShiftToLeft
    ws.Cells.EntireColumn.AutoFit
    ' ws.Rows("164:300").Delete Shift:=xlUp
    
    ' TargetSheet.Range("V:AH").Delete Shift:=xlToLeft
    
    ' Find the cells with the values "Pages / Flows", "Windows", "macOS", "Android", and "iOS"
    Set FoundPageFlow = SourceSheet.Cells.Find(What:="Pages / Flows", LookIn:=xlValues, LookAt:=xlWhole)
    Set FoundWindows = SourceSheet.Cells.Find(What:="Window", LookIn:=xlValues, LookAt:=xlWhole)
    Set FoundMacOS = SourceSheet.Cells.Find(What:="macOS", LookIn:=xlValues, LookAt:=xlWhole)
    Set FoundAndroid = SourceSheet.Cells.Find(What:="Android", LookIn:=xlValues, LookAt:=xlWhole)
    Set FoundIOS = SourceSheet.Cells.Find(What:="iOS", LookIn:=xlValues, LookAt:=xlWhole)
    
    
    
    ' Set the header for the Page/flow column
    
    NextColumn = 12 ' ASCII value for W column
    tablestartcol = NextColumn
    tablestartrow = startRow - 1
    
    
    colLetter = Cells(1, NextColumn).Address(False, False)
    colLetter = Left(colLetter, Len(colLetter) - 1)
    
    colLetter2nd = Cells(1, NextColumn - 1).Address(False, False)
    colLetter2nd = Left(colLetter2nd, Len(colLetter2nd) - 1)

    If Not FoundPageFlow Is Nothing Then
        TargetSheet.Columns(colLetter2nd & ":" & colLetter2nd).ColumnWidth = 36
        With TargetSheet.Cells(startRow - 1, NextColumn - 1)
            .Value = "Portal"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        TargetSheet.Columns(colLetter & ":" & colLetter).ColumnWidth = 56
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "Page/flow"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundPageFlow.Column).End(xlUp).Row
        SourceSheet.Range(FoundPageFlow.Offset(1, 0), SourceSheet.Cells(LastRow, FoundPageFlow.Column)).Copy TargetSheet.Cells(startRow, NextColumn)
        NextColumn = NextColumn + 1
    End If
    
    TargetSheet.Cells(startRow - 1, NextColumn).Value = "Execution Status"
    NextColumn = NextColumn + 1
    
    ' Copy the "Windows" data to the target table
    If Not FoundWindows Is Nothing Then
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "Windows"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundWindows.Column).End(xlUp).Row
        SourceSheet.Range(FoundWindows.Offset(1, 0), SourceSheet.Cells(LastRow, FoundWindows.Column)).Copy TargetSheet.Cells(startRow, NextColumn)
        NextColumn = NextColumn + 1
    End If
    
    ' Copy the "macOS" data to the target table
    If Not FoundMacOS Is Nothing Then
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "macOS"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundMacOS.Column).End(xlUp).Row
        SourceSheet.Range(FoundMacOS.Offset(1, 0), SourceSheet.Cells(LastRow, FoundMacOS.Column)).Copy TargetSheet.Cells(startRow, NextColumn)
        NextColumn = NextColumn + 1
    End If
    
    ' Copy the "Android" data to the target table
    If Not FoundAndroid Is Nothing Then
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "Android"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundAndroid.Column).End(xlUp).Row
        SourceSheet.Range(FoundAndroid.Offset(1, 0), SourceSheet.Cells(LastRow, FoundAndroid.Column)).Copy TargetSheet.Cells(startRow, NextColumn)
        NextColumn = NextColumn + 1
    End If
    
    ' Copy the "IOS" data to the target table
    If Not FoundIOS Is Nothing Then
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "iOS"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundIOS.Column).End(xlUp).Row
        SourceSheet.Range(FoundIOS.Offset(1, 0), SourceSheet.Cells(LastRow, FoundIOS.Column)).Copy TargetSheet.Cells(startRow, NextColumn)
        NextColumn = NextColumn + 1
    End If
    
    TargetSheet.Cells(startRow - 1, NextColumn).Value = "% Completed"
    NextColumn = NextColumn + 1
    
    
    Set HeaderRange = TargetSheet.Range(TargetSheet.Cells(startRow - 1, colLetter2nd), TargetSheet.Cells(startRow - 1, NextColumn - 1))
    
    With HeaderRange
        .Interior.Color = RGB(180, 198, 231) ' Background color #B4C6E7
        .Font.Color = RGB(0, 0, 0) ' Font color black
        .Font.Bold = True
    End With
    
    Dim FilledLastRow As Long
    FilledLastRow = TargetSheet.Cells(TargetSheet.Rows.Count, colLetter).End(xlUp).Row
    Dim wbName As String
    wbName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    tableendrow = FilledLastRow
    tableendcol = NextColumn - 1
    
    Dim tblRange As Range
    Set tblRange = ws.Range(ws.Cells(tablestartrow, tablestartcol), ws.Cells(tableendrow, tableendcol))
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes) ' Change range and table name as per your requirement
    tbl.Name = "Status_Logging_Table"
    tbl.TableStyle = ""
    
    Set tbl = ws.ListObjects("Status_Logging_Table")
    tbl.Range.Borders.LineStyle = xlContinuous
    
    With TargetSheet.Range(colLetter2nd & startRow & ":" & colLetter2nd & FilledLastRow)
        .Merge
        .Value = wbName ' Change as per your requirement
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .WrapText = True
    End With
    
    ' Adding the new table to the SHeet.....................................................................................................................................................................
    
    
    NextColumn = NextColumn + 2 ' Starting column for the new table
    
    tablestartcol = NextColumn
    
    Set FoundCritical = SourceSheet.Cells.Find(What:="Critical", LookIn:=xlValues, LookAt:=xlWhole)
    Set FoundHigh = SourceSheet.Cells.Find(What:="High", LookIn:=xlValues, LookAt:=xlWhole)
    Set FoundMedium = SourceSheet.Cells.Find(What:="Medium", LookIn:=xlValues, LookAt:=xlWhole)
    Set FoundLow = SourceSheet.Cells.Find(What:="Low", LookIn:=xlValues, LookAt:=xlWhole)
    
    colLetter = Cells(1, NextColumn).Address(False, False)
    colLetter = Left(colLetter, Len(colLetter) - 1)
    ' Copy the "Pages / Flows" data to the target table
    If Not FoundPageFlow Is Nothing Then
        TargetSheet.Columns(colLetter & ":" & colLetter).ColumnWidth = 56
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "Page/flow"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundPageFlow.Column).End(xlUp).Row
        SourceSheet.Range(FoundPageFlow.Offset(1, 0), SourceSheet.Cells(LastRow, FoundPageFlow.Column)).Copy TargetSheet.Cells(startRow, NextColumn)
        NextColumn = NextColumn + 1
    End If
    
    
     ' Add a new column for "Total Defect" and populate with sum
     
    TargetSheet.Cells(startRow - 1, NextColumn).Value = "Total Defect Logged"
     
    For i = startRow To FilledLastRow
        Dim totalFormula As String
        totalFormula = "=SUM(" & TargetSheet.Cells(i, NextColumn + 1).Address(False, False) & ":" & TargetSheet.Cells(i, NextColumn + 4).Address(False, False) & ")"
        TargetSheet.Cells(i, NextColumn).Formula = totalFormula
    Next i
    
    NextColumn = NextColumn + 1
    
    ' Copy the "Critical" data to the target table
    If Not FoundCritical Is Nothing Then
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "Critical Impact"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundCritical.Column).End(xlUp).Row
        TargetSheet.Range(TargetSheet.Cells(startRow, NextColumn), TargetSheet.Cells(LastRow + startRow - 2, NextColumn)).Value = SourceSheet.Range(FoundCritical.Offset(1, 0), SourceSheet.Cells(LastRow, FoundCritical.Column)).Value
        NextColumn = NextColumn + 1
    End If


    
    If Not FoundHigh Is Nothing Then
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "High Impact"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundHigh.Column).End(xlUp).Row
        TargetSheet.Range(TargetSheet.Cells(startRow, NextColumn), TargetSheet.Cells(LastRow + startRow - 2, NextColumn)).Value = SourceSheet.Range(FoundHigh.Offset(1, 0), SourceSheet.Cells(LastRow, FoundHigh.Column)).Value
        NextColumn = NextColumn + 1
    End If
    
    If Not FoundMedium Is Nothing Then
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "Medium Impact"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundMedium.Column).End(xlUp).Row
        TargetSheet.Range(TargetSheet.Cells(startRow, NextColumn), TargetSheet.Cells(LastRow + startRow - 2, NextColumn)).Value = SourceSheet.Range(FoundMedium.Offset(1, 0), SourceSheet.Cells(LastRow, FoundMedium.Column)).Value
        NextColumn = NextColumn + 1
    End If
    
    If Not FoundLow Is Nothing Then
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "Low Impact"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundLow.Column).End(xlUp).Row
        TargetSheet.Range(TargetSheet.Cells(startRow, NextColumn), TargetSheet.Cells(LastRow + startRow - 2, NextColumn)).Value = SourceSheet.Range(FoundLow.Offset(1, 0), SourceSheet.Cells(LastRow, FoundLow.Column)).Value
        NextColumn = NextColumn + 1
    End If
    
    
    Set HeaderRange = TargetSheet.Range(TargetSheet.Cells(startRow - 1, colLetter), TargetSheet.Cells(startRow - 1, NextColumn - 1))
    
    With HeaderRange
        .Interior.Color = RGB(180, 198, 231) ' Background color #B4C6E7
        .Font.Color = RGB(0, 0, 0) ' Font color black
        .Font.Bold = True
    End With
    
    tableendcol = NextColumn - 1
    Set tblRange = ws.Range(ws.Cells(tablestartrow, tablestartcol), ws.Cells(tableendrow, tableendcol))
    Set tbl1 = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes) ' Change range and table name as per your requirement
    tbl1.Name = "Defect_Logging_Table"
    tbl1.TableStyle = ""
    
    tbl1.ListColumns(2).Range.Font.Bold = True
' Center content of all columns from second column to last column
    tbl1.Range.Columns(2).HorizontalAlignment = xlCenter

' Center content of columns 3 to 6
    For i = 3 To 6
        tbl1.ListColumns(i).Range.HorizontalAlignment = xlCenter
    Next i
    
    ' Adding the new table to the SHeet....................................................................................................................................................................
    NextColumn = NextColumn + 2 ' Starting column for the new table
    tablestartcol = NextColumn
    
    Set FoundA = SourceSheet.Cells.Find(What:="Level A", LookIn:=xlValues, LookAt:=xlWhole)
    Set FoundAA = SourceSheet.Cells.Find(What:="Level AA", LookIn:=xlValues, LookAt:=xlWhole)
    
    colLetter = Cells(1, NextColumn).Address(False, False)
    colLetter = Left(colLetter, Len(colLetter) - 1)
    ' Copy the "Pages / Flows" data to the target table
    If Not FoundPageFlow Is Nothing Then
        TargetSheet.Columns(colLetter & ":" & colLetter).ColumnWidth = 56
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "Page/flow"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundPageFlow.Column).End(xlUp).Row
        SourceSheet.Range(FoundPageFlow.Offset(1, 0), SourceSheet.Cells(LastRow, FoundPageFlow.Column)).Copy TargetSheet.Cells(startRow, NextColumn)
        NextColumn = NextColumn + 1
    End If
    
    ' Copy the "Critical" data to the target table
    If Not FoundA Is Nothing Then
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "Level A"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundA.Column).End(xlUp).Row
        TargetSheet.Range(TargetSheet.Cells(startRow, NextColumn), TargetSheet.Cells(LastRow + startRow - 2, NextColumn)).Value = SourceSheet.Range(FoundA.Offset(1, 0), SourceSheet.Cells(LastRow, FoundA.Column)).Value
        NextColumn = NextColumn + 1
    End If
    
    If Not FoundAA Is Nothing Then
        TargetSheet.Cells(startRow - 1, NextColumn).Value = "Level AA"
        LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, FoundAA.Column).End(xlUp).Row
        TargetSheet.Range(TargetSheet.Cells(startRow, NextColumn), TargetSheet.Cells(LastRow + startRow - 2, NextColumn)).Value = SourceSheet.Range(FoundAA.Offset(1, 0), SourceSheet.Cells(LastRow, FoundAA.Column)).Value
        NextColumn = NextColumn + 1
    End If
    
    Set HeaderRange = TargetSheet.Range(TargetSheet.Cells(startRow - 1, colLetter), TargetSheet.Cells(startRow - 1, NextColumn - 1))
    
    With HeaderRange
        .Interior.Color = RGB(180, 198, 231) ' Background color #B4C6E7
        .Font.Color = RGB(0, 0, 0) ' Font color black
        .Font.Bold = True
    End With
    
    tableendcol = NextColumn - 1
    Set tblRange = ws.Range(ws.Cells(tablestartrow, tablestartcol), ws.Cells(tableendrow, tableendcol))
    Set tbl2 = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes) ' Change range and table name as per your requirement
    tbl2.Name = "Conf_Logging_Table"
    tbl2.TableStyle = ""
    
    tbl2.ListColumns(2).Range.Font.Bold = True
' Center content of all columns from second column to last column
    tbl2.Range.Columns(2).HorizontalAlignment = xlCenter
    
    tbl2.ListColumns(3).Range.Font.Bold = True
' Center content of all columns from third column to last column
    tbl2.Range.Columns(3).HorizontalAlignment = xlCenter
    
    ws.Rows(tableendrow + 1 & ":" & tableendrow + 2).Delete
    
    
    ' Adding boundary to the table....................................................................................................................................................................
    
    Set tbl = ws.ListObjects("Conf_Logging_Table")
    tbl.Range.Borders.LineStyle = xlContinuous
    
    ' Set the Defect_Logging_Table and add borders
    Set tbl = ws.ListObjects("Defect_Logging_Table")
    tbl.Range.Borders.LineStyle = xlContinuous
    
    ' Set the Status_Logging_Table and add borders
    
    

End Sub

