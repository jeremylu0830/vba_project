Attribute VB_Name = "Loading"
Sub downloadMap()
    ' 顯示加載消息
    LoadingUF.Show vbModeless
    ' 讓顯示 UserForm 的變化立即反映出來
    DoEvents
    
    ' 執行長時間運行的代碼
    Call insertMap
    
    ' 隱藏加載消息
    Unload LoadingUF
End Sub

Sub LongRunningTask()
    ' 這是模擬的長時間運行的代碼
    Dim i As Long
    For i = 1 To 1000
        ' 模擬一些計算或操作
        Debug.Print i
    Next i
End Sub

