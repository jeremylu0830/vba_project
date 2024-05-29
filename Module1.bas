Attribute VB_Name = "Module1"
Dim currentrow As Integer
Dim currentcol As Integer
Dim arrowrow As Integer
Dim arrowcol As Integer
Dim gamestatus As String
'BUG: 訂購完東西有時候腳印無法移動，懷疑是鎖定的問題
'待處理:select按鍵後減少庫存，可以放回去的功能、箱子帶著跑等等


Sub InitializeCellsToSquareInRange()
    
    'gamestatus = "move"
    Dim ws As Worksheet
    Dim rowHeight As Double
    Dim colWidth As Double
    Dim targetRange As Range
    Dim i As Long

    ' 設置工作表對象
    Set ws = ThisWorkbook.Sheets("Warehouse")
    Set targetRange = ws.Range("A1:Z100") ' 設置你希望調整的範圍

    ' 設定行高
    rowHeight = 20 ' 這裡設定行高，你可以根據需要調整
    colWidth = rowHeight * 0.1428 ' 估算列寬，使單元格接近正方形

    ' 設置目標範圍內的行高度
    targetRange.Rows.rowHeight = rowHeight
    ' 設置目標範圍內的列寬度
    For i = 1 To targetRange.Columns.Count
        targetRange.Columns(i).ColumnWidth = colWidth
    Next i
    
    
    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("HideWarehouse")
    'all Range
    ws3.Range("a1:t20") = 0
    'shelf
    ws3.Range("E3:L9") = 3
    'wall
    ws3.Range("A1:A20") = 1
    ws3.Range("B1:t1") = 1
    ws3.Range("t1:t20") = 1
    ws3.Range("b20:t20") = 1
    'leave
    ws3.Range("s19:t20") = 4
    'can use pickup area
    ws3.Range("e10:l10") = 2
    
    'cart and order
    ws3.Range("b18:c19") = 5
    
    '訂單的form 初始化
    Dim ws4 As Worksheet
    Set ws4 = ThisWorkbook.Sheets("each_order")
    For i = 2 To 10
        ws4.Cells(i, "c") = 0
        ws4.Cells(i, "d") = 0
        ws4.Cells(i, "e") = 0
    Next
    
    
    
End Sub
'use this to start
Sub InsertPictureInCell()
    gamestatus = "move"
    InitializeCellsToSquareInRange '初始化格子正方形 與背景倉庫
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Warehouse")
    'ws2.Activate
    ws2.Unprotect
    deleteallpicture
    
    'wall
    k = 1
    For i = 1 To 20
        For j = 1 To 20
            If 工作表3.Cells(i, j) = 1 Then
                picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\wall.png" ' 更改為你的圖片的完整路徑
                Set targetCell = ws2.Cells(i, j)
                pictureName = "wall" & k
                picturesetting picturePath, targetCell, pictureName, 0
                k = k + 1
            End If
        Next
    Next
     'shelf1
    picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\shelf2.png" ' 更改為你的圖片的完整路徑
    Set targetCell = ws2.Range("B3:O9")
    pictureName = "shelf"
    picturesetting picturePath, targetCell, pictureName, 0
    

    'milk
    'picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\milk.png"
    'Set targetCell = ws2.Range("E3:F4") ' 更改為你想插入圖片的單元格
    'pictureName = "milk"
    'picturesetting picturePath, targetCell, pictureName, 0
    
    'test
    'egg
    'picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\egg.png"
    'Set targetCell = ws2.Range("G3:H4") ' 更改為你想插入圖片的單元格
    'pictureName = "egg"
    'picturesetting picturePath, targetCell, pictureName, 0
    
    'candy
    'picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\candy.png"
    'Set targetCell = ws2.Range("i3:j4") ' 更改為你想插入圖片的單元格
    'pictureName = "candy"
    'picturesetting picturePath, targetCell, pictureName, 0
    
    'cola
    'picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\cola.png"
    'Set targetCell = ws2.Range("k3:l4") ' 更改為你想插入圖片的單元格
    'pictureName = "cola"
    'picturesetting picturePath, targetCell, pictureName, 0
   
    
    'where is me
    picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\me.png"
    Set targetCell = ws2.Cells(2, 2) ' 更改為你想插入圖片的單元格
    pictureName = "me"
    picturesetting picturePath, targetCell, pictureName, 0
    currentrow = 2
    currentcol = 2
   
    'leave sign
    picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\leave.png"
    Set targetCell = ws2.Range("S19:t20")
    pictureName = "leave"
    picturesetting picturePath, targetCell, pictureName, 0
    
    'cart sign
    picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\cart.png"
    Set targetCell = ws2.Range("b18:c19")
    pictureName = "cart"
    picturesetting picturePath, targetCell, pictureName, 0
    
    
   
End Sub
Sub AfterForm()
    gamestatus = "move"
    'InitializeCellsToSquareInRange '初始化格子正方形 與背景倉庫
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Warehouse")
    'ws2.Activate
    ws2.Unprotect
    'deleteallpicture
    
 
    For Each pic In ws2.Shapes
                If pic.Name = "me" Then
                    pic.Delete
                    Exit For
                End If
    Next pic
    'where is me
    picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\me.png"
    Set targetCell = ws2.Cells(17, "d") ' 更改為你想插入圖片的單元格
    pictureName = "me"
    picturesetting picturePath, targetCell, pictureName, 0
    currentrow = 17
    currentcol = 4
   
   
    
End Sub

Sub picturesetting(picturePath, targetCell, pictureName, angle)
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Warehouse")
    ws2.Activate
    'Dim pic As Shape
    Set pic = ws2.Pictures.Insert(picturePath)
    
    pic.Name = pictureName
    
    With pic
        'If pictureName = "me" Then
        '    Set pic = ws2.Shapes("me")
        '    .Rotation = angle
        'End If
        '.Rotation = angle
        .ShapeRange.LockAspectRatio = msoFalse
        .Top = targetCell.Top
        .Left = targetCell.Left
        .Width = targetCell.Width
        .Height = targetCell.Height
        .Locked = True ' 鎖定圖片
        .PrintObject = False ' 使圖片不可選中
        
    End With
    ws2.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    
    
    
End Sub
'刪除全部
Sub deleteallpicture()


    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Warehouse")
    'ws2.Activate
    Dim pic As Shape
    'test delete
    ' 刪除工作表中的所有圖片
    For Each pic In ws2.Shapes
            'pic.Delete
            'If pic.Type = msoPicture Then
            pic.Delete
            'End If
    Next pic
    
    ws2.Range("A1:U21") = " "
End Sub

Sub MovePictureByArrowKey(rowOffset As Integer, colOffset As Integer, angle As Integer)


    Dim ws2 As Worksheet
    Dim pic As Shape
    Set ws2 = ThisWorkbook.Sheets("Warehouse")
    Set ws3 = ThisWorkbook.Sheets("HideWarehouse")
    
    Dim targetCell As Range
    If gamestatus = "move" Then
        Set pic = ws2.Shapes("me")
        If ws3.Cells(currentrow + rowOffset, currentcol + colOffset).Value = 0 Or ws3.Cells(currentrow + rowOffset, currentcol + colOffset).Value = 2 Or ws3.Cells(currentrow + rowOffset, currentcol + colOffset).Value = 5 Then  ' not wall
            For Each pic In ws2.Shapes
                If pic.Name = "me" Then
                    pic.Delete
                    Exit For
                End If
            Next pic
            
            currentrow = currentrow + rowOffset
            currentcol = currentcol + colOffset
            picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\me.png"
            Set targetCell = ws2.Cells(currentrow, currentcol) ' 更改為你想插入圖片的單元格
            pictureName = "me"
            picturesetting picturePath, targetCell, pictureName, angle
            Set pic = ws2.Shapes("me")
            pic.Rotation = angle
        End If
    ElseIf gamestatus = "stop" Then
        
        Set pic = ws2.Shapes("point")
        If ws3.Cells(arrowrow + rowOffset, arrowcol + colOffset).Value <> 0 And ws3.Cells(arrowrow + rowOffset, arrowcol + colOffset).Value <> 2 Then
            For Each pic In ws2.Shapes
                If pic.Name = "point" Then
                    pic.Delete
                    Exit For
                End If
            Next pic
            
            arrowrow = arrowrow + rowOffset
            arrowcol = arrowcol + colOffset
            picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\point.png"
            Set targetCell = ws2.Cells(arrowrow, arrowcol) ' 更改為你想插入圖片的單元格
            pictureName = "point"
            picturesetting picturePath, targetCell, pictureName, angle
            Set pic = ws2.Shapes("point")
        End If
    End If
    
End Sub
Sub RotatePicture(angle) 'change the direction of footprint
    Dim ws2 As Worksheet
    Dim pic As Shape
    Set ws2 = ThisWorkbook.Sheets("Warehouse")
    Set pic = ws2.Shapes("me")
    'pic.ShapeRange.Rotation = 90
    'pic.Rotation = angle
    'End With
End Sub


Sub MovePictureUp()
    'If gamestatus = "move" Then
        MovePictureByArrowKey -1, 0, 0
    'End If
End Sub

Sub MovePictureDown()
    'If gamestatus = "move" Then
        MovePictureByArrowKey 1, 0, 180
    'End If
End Sub

Sub MovePictureLeft()
    'If gamestatus = "move" Then
        MovePictureByArrowKey 0, -1, 270
    'End If
End Sub

Sub MovePictureRight()
    'If gamestatus = "move" Then
        MovePictureByArrowKey 0, 1, 90
    'End If
End Sub
Sub PickUp2()
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Warehouse")
    ws2.Cells(21, "a") = "Type TAB to choose item to pick"
    'If 工作表3.Cells(currentrow, currentcol).Value = 2 Then
        
        picturePath = "C:\Users\NUTC\OneDrive\桌面\PictureInput\point.png"
        arrowrow = 4
        arrowcol = 6
        Set targetCell = ws2.Cells(arrowrow, arrowcol)
        pictureName = "point"
        picturesetting picturePath, targetCell, pictureName, 0
    'End If
End Sub

Sub PickUp()
    'gamestatus = "stop"
    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("HideWarehouse")
    If ws3.Cells(currentrow, currentcol).Value = 2 And gamestatus = "move" Then
        gamestatus = "stop"
        
        PickUp2 '為啥要呼叫兩次...
    'End If
    ' order mode and call
    ElseIf ws3.Cells(currentrow, currentcol).Value = 5 Then
        'gamestatus = "order"
        'Dim OrderItem As OrderLevel1
        Dim OrderItem As New OrderLevel1
            OrderItem.Show
            
    'End If
    ElseIf gamestatus = "select" Then
        Dim SelectItem As New SelectQuantity
            SelectItem.Show
    'End If
    
    ElseIf gamestatus = "stop" And ws3.Cells(arrowrow, arrowcol) <> 3 And ws3.Cells(currentrow, currentcol).Value = 2 Then
        'open form to select item to carry
        gamestatus = "select"
        'Dim SelectItem As New SelectQuantity
        '    SelectItem.Show
    End If
End Sub

