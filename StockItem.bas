Attribute VB_Name = "StockItem"
Sub SockItemOnShelf()
 Set ws4 = Worksheets("Goods") 'each_order
 Set ws3 = Worksheets("HideWarehouse") 'hidewarehouse
 Set ws2 = Worksheets("Warehouse") 'warehouse
 ABSroute = ThisWorkbook.path & "\PictureInput\"
 For i = 1 To 32
    If ws4.Cells(i, "h") = 0 And ws4.Cells(i, "d") > 0 Then
        picturePath = ABSroute & ws4.Cells(i, "a") & ".png"
        ws3.Cells(21, "a") = picturePath
        'index=i
        ws3.Cells(i Mod 4 + 2, Int(i / 4 + 1) + 4) = ws4.Cells(i, "a")
        'put UI
        Set targetCell = ws2.Cells(i Mod 4 + 2, Int(i / 4 + 1) + 4) ' = ws4.Cells(i, "a")
        pictureName = ws4.Cells(i, "a")

        picturesetting picturePath, targetCell, pictureName, 0
    End If
        ws4.Cells(i, "h") = ws4.Cells(i, "d") + ws4.Cells(i, "h")
        
        'ws4.Cells(i, "c") = 0
        'ws3.Range(temp) = ws4.Cells(i, "a")
        
    'End If
    Next
    
End Sub
