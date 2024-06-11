Attribute VB_Name = "TimeControl"
Sub UpdateTime() 'if open store then
    Dim ws As Worksheet
    Set ws = Worksheets("Interface")

    Dim currentTime As Integer
    currentTime = ws.Cells(2, 1).value
    
    If IsEmpty(currentTime) Then
        'currentTime = Now
        currentTime = 1
    Else
        
        'currentTime = DateAdd("s", 1, currentTime)
        currentTime = currentTime + 1 'DateAdd("s", 1, currentTime)
    End If
    
    ' renew
    ws.Cells(2, 1).value = currentTime
    If ws.Cells(2, 1).value = ws.Cells(2, 2).value Then
        'finish the timer
        StopTimer
    End If
    
    'customer control
    If currentTime Mod 10 = 0 And currentTime < 50 Then
        If IsRangeEmpty(Sheets("HidemarketQuantity").Range("A1:Z23")) Then
            MsgBox "we dont have anything to sell"
        Else
            RandomSelectCellWithNumbers
        End If
    End If
    
    nextTick = Now + TimeValue("00:00:01")
    If ws.Cells(2, 1).value = ws.Cells(2, 2).value Then
        'finish the timer
        StopTimer
    Else
        Application.OnTime nextTick, "UpdateTime"
    End If
End Sub

Sub StopTimer() 'use close shop button to stop the timer
    On Error Resume Next
    Application.OnTime nextTick, "UpdateTime", , False
    On Error GoTo 0
End Sub

Function IsRangeEmpty(rng As Range) As Boolean
    Dim cell As Range
    IsRangeEmpty = True
    
    For Each cell In rng
        If Not IsEmpty(cell.value) Then
            IsRangeEmpty = False
            Exit Function
        End If
    Next cell
End Function
