Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal HWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal HWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

Dim TimerID As LongPtr

Sub StartTimer()
    TimerID = SetTimer(0, 0, 300, AddressOf TimerProc)
End Sub

Sub StopTimer()
    On Error Resume Next
    KillTimer 0, TimerID
End Sub

Sub TimerProc(ByVal HWnd As LongPtr, ByVal uMsg As Long, ByVal nIDEvent As LongPtr, ByVal dwTimer As LongPtr)
    ' Stop the timer
    Call StopTimer
    
    ' Do the task
    Call AddRandomDataAndWait
    
    ' Restart the timer
    Call StartTimer
End Sub

Sub AddRandomDataAndWait()
    ' Set up variables
    Dim dataSheet As Worksheet
    Dim dataRange As Range
    Dim numRows As Integer
    Dim numCols As Integer
    Dim i As Integer
    Dim j As Integer
    
    ' Get reference to data sheet
    Set dataSheet = ThisWorkbook.Sheets("Sheet1")
    
    ' Determine number of rows and columns to add
    numRows = Int(Rnd() * 10) + 1
    numCols = Int(Rnd() * 10) + 1
    
    ' Add random data to sheet
    Set dataRange = dataSheet.Range("A1").Resize(numRows, numCols)
    For i = 1 To numRows
        For j = 1 To numCols
            dataRange.Cells(i, j).Value = Int(Rnd() * 100)
        Next j
    Next i
    
End Sub

