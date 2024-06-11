Attribute VB_Name = "WarehouseControl"
Dim currentrow As Integer
Dim currentcol As Integer
Dim arrowrow As Integer
Dim arrowcol As Integer
Dim gamestatus As String
Dim inventoryS As Integer



Sub InitializeCellsToSquareInRange()
    
    'gamestatus = "move"
    Dim ws As Worksheet
    Dim rowHeight As Double
    Dim colWidth As Double
    Dim targetRange As Range
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Warehouse")
    Set targetRange = ws.Range("A1:Z100")

    rowHeight = 20
    colWidth = rowHeight * 0.1428
    targetRange.rows.rowHeight = rowHeight
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
    
End Sub
'use this to start
Sub InsertPictureInCell()
    gamestatus = "move"
    'InitializeCellsToSquareInRange
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Warehouse")
    ws2.Activate
    ws2.Unprotect
    
    If inventoryS <> 1 Then
        deleteallpicture
        Sheets("Goods").Range("H1:H38").value = ""
    End If
    
    'wall
    k = 1
    For i = 1 To 20
        For j = 1 To 20
            If Sheets("HideWarehouse").Cells(i, j) = 1 Then
                picturePath = ThisWorkbook.path & "\PictureInput\wall.png"
                Set targetCell = ws2.Cells(i, j)
                pictureName = "wall" & k
                picturesetting picturePath, targetCell, pictureName, 0
                k = k + 1
            End If
        Next
    Next
    
    'where is me
    picturePath = ThisWorkbook.path & "\PictureInput\me.png"
    Set targetCell = ws2.Cells(2, 2)
    pictureName = "me"
    picturesetting picturePath, targetCell, pictureName, 0
    currentrow = 2
    currentcol = 2
   
    'leave sign
    picturePath = ThisWorkbook.path & "\PictureInput\leave.png"
    Set targetCell = ws2.Range("S19:t20")
    pictureName = "leave"
    picturesetting picturePath, targetCell, pictureName, 0
    
    'cart sign
    picturePath = ThisWorkbook.path & "\PictureInput\cart.png"
    Set targetCell = ws2.Range("b18:c19")
    pictureName = "cart"
    picturesetting picturePath, targetCell, pictureName, 0
    
    
   
End Sub
Sub AfterForm()
    gamestatus = "move"
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
    picturePath = ThisWorkbook.path & "\PictureInput\me.png"
    Set targetCell = ws2.Cells(17, "d")
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
        .width = targetCell.width
        .height = targetCell.height
        .Locked = True
        .PrintObject = False
        
    End With
    ws2.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    
    
    
End Sub
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

        If ws3.Cells(currentrow + rowOffset, currentcol + colOffset).value = 0 Or _
            ws3.Cells(currentrow + rowOffset, currentcol + colOffset).value = 2 Or _
            ws3.Cells(currentrow + rowOffset, currentcol + colOffset).value = 5 Then  ' not wall
            For Each pic In ws2.Shapes
                If pic.Name = "me" Then
                    pic.Delete
                    Exit For
                End If
            Next pic
            
            currentrow = currentrow + rowOffset
            currentcol = currentcol + colOffset
            picturePath = ThisWorkbook.path & "\PictureInput\me.png"
            Set targetCell = ws2.Cells(currentrow, currentcol)
            pictureName = "me"
            picturesetting picturePath, targetCell, pictureName, angle
            Set pic = ws2.Shapes("me")
            pic.Rotation = angle
            
        ElseIf ws3.Cells(currentrow + rowOffset, currentcol + colOffset).value = 4 Then
            'set the status, 1 = init, 2 = back from the warehouse
            
            For Each pic In ws2.Shapes
                If pic.Name = "me" Then
                    pic.Delete
                    Exit For
                End If
            Next pic
            inventoryS = 1
            MainFunction 2
        
        End If
    'select thing in warehouse
    ElseIf gamestatus = "selectQ" Then
        
        Set pic = ws2.Shapes("point")
        If ws3.Cells(arrowrow + rowOffset, arrowcol + colOffset).value <> 0 And ws3.Cells(arrowrow + rowOffset, arrowcol + colOffset).value <> 2 Then
            For Each pic In ws2.Shapes
                If pic.Name = "point" Then
                    pic.Delete
                    Exit For
                End If
            Next pic
            
            arrowrow = arrowrow + rowOffset
            arrowcol = arrowcol + colOffset
            picturePath = ThisWorkbook.path & "\PictureInput\point.png"
            Set targetCell = ws2.Cells(arrowrow, arrowcol)
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
End Sub


Sub MovePictureUp()
    MovePictureByArrowKey -1, 0, 0
End Sub

Sub MovePictureDown()
    MovePictureByArrowKey 1, 0, 180
End Sub

Sub MovePictureLeft()
    MovePictureByArrowKey 0, -1, 270
End Sub

Sub MovePictureRight()
    MovePictureByArrowKey 0, 1, 90
End Sub

'select -> selectQ
'selectSHOP
'tab key
Sub PickUp()
    'gamestatus = "stop"
    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("HideWarehouse")
    
    Dim ws As Worksheet
    Dim pic As Shape
    Set ws = ThisWorkbook.Sheets("Warehouse")
    
    
    If ws3.Cells(currentrow, currentcol).value = 2 And gamestatus <> "selectQ" Then 'selectquantity
        gamestatus = "selectQ"
        
        Dim ws2 As Worksheet
        Set ws2 = ThisWorkbook.Sheets("Warehouse")
        picturePath = ThisWorkbook.path & "\PictureInput\point.png"
        arrowrow = 4
        arrowcol = 6
        Set targetCell = ws2.Cells(arrowrow, arrowcol)
        pictureName = "point"
        picturesetting picturePath, targetCell, pictureName, 0
    ElseIf ws3.Cells(currentrow, currentcol).value = 5 And gamestatus <> "selectSHOP" Then 'selectSHOP
    
        gamestatus = "selectSHOP"
        RandomPrice
        SHOP.Show
        gamestatus = "move"
    
    ElseIf gamestatus = "selectQ" Then
        
        SelectQuantity.Label1.Caption = ws3.Cells(arrowrow, arrowcol).value
        SelectQuantity.Show
        gamestatus = "move"
        
        For Each pic In ws.Shapes
            If pic.Name = "point" Then
                pic.Delete
                Exit For
            End If
        Next pic
        
    End If
End Sub


Sub RandomPrice()
    Set ws4 = Worksheets("Goods")
    For i = 1 To 32
    lowerBound = ws4.Cells(i, "f").value
    upperBound = ws4.Cells(i, "g").value
    ws4.Cells(i, "b") = Int(lowerBound + Rnd * (upperBound - lowerBound))
    Next
End Sub
 
