Attribute VB_Name = "MarketControl"
'!!! control main market map

Dim currentrow As Integer
Dim currentcol As Integer
Dim arrowrow As Integer
Dim arrowcol As Integer
Dim gamestatus As String

Sub control_ws()
    Dim ws As Worksheet
    Dim wsH As Worksheet
    Set ws = ThisWorkbook.Sheets("market")
    Set wsH = ThisWorkbook.Sheets("Hidemarket")
    
    ws.Unprotect
    
End Sub

Sub test()
    
    Set ws = ThisWorkbook.Sheets("market")
    
    
End Sub

'use this to start

Sub startMain()

    'initial UI setting
    
    Dim ws As Worksheet
    Dim wsH As Worksheet
    Set ws = ThisWorkbook.Sheets("market")
    Set wsH = ThisWorkbook.Sheets("Hidemarket")
    
    Dim rowHeight As Double
    Dim colWidth As Double
    Dim targetRange As Range
    Dim i As Long
    
    
    'imgFloderPath = ThisWorkbook.path & "\PictureInput"
    
    ws.Unprotect

    Set targetRange = ws.Range("A1:Z100") ' setting the range u need

    rowHeight = 20
    colWidth = rowHeight * 0.1428

    targetRange.rows.rowHeight = rowHeight
    For i = 1 To targetRange.Columns.Count
        targetRange.Columns(i).ColumnWidth = colWidth
    Next i
    
            
    'cell setting

    delAll ws
    
    ws.Range("A1:W23").ClearContents
    
    'call loadind UF while downloading map.
    downloadMap
    
    
    reset
    
    ws.[aa1] = "customer list"
    ws.[aa1].ColumnWidth = 20
    ws.[aa1].rowHeight = 20
    ws.[aa1].Font.Size = 20
    ws.[ab1] = "TIME"
    ws.[ab1].ColumnWidth = 20
    ws.[ab1].rowHeight = 20
    ws.[ab1].Font.Size = 20
    With ws.Range("AB1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.[ab2] = Sheets("Interface").[a2]
     ws.[ab2].Font.Size = 20
     
     ws.Activate
    'start game!!
    MainFunction 1
    
End Sub

Sub MainFunction(status)
    'gamestatus = "move"
    'InitializeCellsToSquareInRange
    
    Dim ws As Worksheet
    Dim wsH As Worksheet
    Set ws = ThisWorkbook.Sheets("market")
    Set wsH = ThisWorkbook.Sheets("Hidemarket")
    ws.Activate
    ws.Unprotect
    
    If status = 1 Then
        currentrow = 2
        currentcol = 2
        'where is me
        picturePath = ThisWorkbook.path & "\PictureInput\me.png"
        Set targetCell = wsH.Cells(2, "B")
        pictureName = "me"
        picset picturePath, targetCell, pictureName, 0, ws
    ElseIf status = 2 Then
        
        currentrow = 20
        currentcol = 22
    
    End If
    
    'suppose start the game

End Sub

Sub picset(picturePath, targetCell, pictureName, angle, ws)
    ws.Activate
    'Dim pic As Shape
    Set pic = ws.Pictures.Insert(picturePath)
    
    pic.Name = pictureName
    
    With pic
        '.ShapeRange.LockAspectRatio = msoFalse
        .Top = targetCell.Top
        .Left = targetCell.Left
        .width = targetCell.width
        .height = targetCell.height
        .Locked = True
        .PrintObject = False
        .Placement = 1
        
    End With
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    
End Sub

Sub delAll(ws)

    Dim pic As Shape

    For Each pic In ws.Shapes

            pic.Delete
            'End If
    Next pic
    
    ws.Range("A1:AP29") = " "
End Sub

Sub MovePic(rowOffset As Integer, colOffset As Integer, angle As Integer)


    Dim ws As Worksheet
    Dim pic As Shape
    Set ws = ThisWorkbook.Sheets("Market")
    Set wsH = ThisWorkbook.Sheets("Hidemarket")
    
    Dim targetCell As Range
    Set pic = ws.Shapes("me")
    If wsH.Cells(currentrow + rowOffset, currentcol + colOffset).value = 0 Then 'move
        For Each pic In ws.Shapes
            If pic.Name = "me" Then
                pic.Delete
                Exit For
            End If
        Next pic
            
        currentrow = currentrow + rowOffset
        currentcol = currentcol + colOffset
        picturePath = ThisWorkbook.path & "\PictureInput\me.png"
        Set targetCell = ws.Cells(currentrow, currentcol)
        pictureName = "me"
        picset picturePath, targetCell, pictureName, angle, ws
        Set pic = ws.Shapes("me")
        pic.Rotation = angle
        
    ElseIf wsH.Cells(currentrow + rowOffset, currentcol + colOffset).value = 4 Then
        
        'Go to warehouse
        InsertPictureInCell
  
    ElseIf wsH.Cells(currentrow + rowOffset, currentcol + colOffset).value = 9 Then
        
        'Open the store or not
        OPEN_CLOSE.Show
        
        'after open, let the customer step into the shop
    ElseIf wsH.Cells(currentrow + rowOffset, currentcol + colOffset).value = 8 Then
        
        PoS_UF.Show
        
    End If
    
End Sub



Sub MovePicUp()
    MovePic -1, 0, 0
End Sub

Sub MovePicDown()
    MovePic 1, 0, 180
End Sub

Sub MovePicLeft()
    MovePic 0, -1, 270
End Sub

Sub MovePicRight()
    MovePic 0, 1, 90
End Sub

Sub Release()
    Dim wsH As Worksheet
    Set wsH = ThisWorkbook.Sheets("Hidemarket")
    
    Dim ReleaseFlag As Boolean
    ReleaseFlag = False
    'Infront of shelf!!!
    
    Dim direction
    direction = Array(Array(-1, 0), Array(0, 1), Array(1, 0), Array(0, -1))
    
    For i = 0 To 3
        currentrow = currentrow + direction(i)(0)
        currentcol = currentcol + direction(i)(1)

        If wsH.Cells(currentrow, currentcol).value = 3 Then  'When the shlef is infront of player, release the goods
            ReleaseFlag = True
            Exit For
        End If
        currentrow = currentrow - direction(i)(0)
        currentcol = currentcol - direction(i)(1)
    Next
    
    If ReleaseFlag = True Then   'selectquantity
        price.Label4.Caption = Sheets("Interface").Cells(2, "h")
        price.Label5.Caption = Sheets("Interface").Cells(2, "j") 'cost
        price.ScrollBar1.Min = Sheets("Interface").Cells(2, "j")
        price.ScrollBar1.Max = 3 * Sheets("Interface").Cells(2, "j")
        
        price.Show
    End If
End Sub


Sub place()
    Dim ws As Worksheet
    Dim wsH As Worksheet
    Dim wsQ As Worksheet
    Dim wsP As Worksheet
    Dim wsGC As Worksheet
    Dim wsPro As Worksheet
    
    Set ws = ThisWorkbook.Sheets("market")
    Set wsH = ThisWorkbook.Sheets("Hidemarket")
    Set wsQ = ThisWorkbook.Sheets("HidemarketQuantity")
    Set wsP = ThisWorkbook.Sheets("HidemarketPrice")
    Set wsGC = ThisWorkbook.Sheets("goodCust")
    Set wsPro = ThisWorkbook.Sheets("product_info")
    wsH.Cells(currentrow, currentcol) = Sheets("Interface").Cells(2, "h")
    wsGC.Cells(currentrow, currentcol) = Sheets("Interface").Cells(2, "h")
    wsQ.Cells(currentrow, currentcol) = Sheets("Interface").Cells(2, "i")
    wsP.Cells(currentrow, currentcol) = Sheets("Interface").Cells(2, "j")
    
    myrow = WorksheetFunction.Match(CStr(Sheets("Interface").Cells(2, "h")), wsPro.[c:c], 0)
    
    wsPro.Cells(myrow, "d") = Sheets("Interface").Cells(2, "j")
    
    
    
    
    picturePath = ThisWorkbook.path & "\PictureInput\" & Sheets("Interface").Cells(2, "h") & ".png"
    Set targetCell = ws.Cells(currentrow, currentcol)
    pictureName = Sheets("Interface").Cells(2, "h")
    
    picset picturePath, targetCell, pictureName, 0, ws
    
    Sheets("Interface").Cells(2, "h").value = 0
    Sheets("Interface").Cells(2, "i").value = 0
    Sheets("Interface").Cells(2, "j").value = 0
    
End Sub

Sub reset()
    

    
    With Sheets("goodCust")
    .Range("g3:l4") = ""
    .Range("o5:t6") = ""
    .Range("g8:l9") = ""
    .Range("g14:l15") = ""
    .Range("o11:t12") = ""
   End With
   
     With Sheets("HidemarketQuantity")
    .Range("g3:l4") = ""
    .Range("o5:t6") = ""
    .Range("g8:l9") = ""
    .Range("g14:l15") = ""
    .Range("o11:t12") = ""
   End With
   With Sheets("HidemarketPrice")
    .Range("g3:l4") = ""
    .Range("o5:t6") = ""
    .Range("g8:l9") = ""
    .Range("g14:l15") = ""
    .Range("o11:t12") = ""
   End With
   
   With Sheets("Hidemarket")
    .Range("g3:l4") = 3
    .Range("o5:t6") = 3
    .Range("g8:l9") = 3
    .Range("g14:l15") = 3
    .Range("o11:t12") = 3
   End With

    With Sheets("HideWarehouse")
    .Range("e2:l9") = 3
    
    End With
    
    
End Sub

Sub insertMap()

    Dim ws As Worksheet
    Dim wsH As Worksheet
    Set ws = ThisWorkbook.Sheets("market")
    Set wsH = ThisWorkbook.Sheets("Hidemarket")
    
    
     For Row = 1 To 23
        For Column = 1 To 23
            If wsH.Cells(Row, Column) = 1 Then
                'Sheets("photo").[b1].Copy Destination:=ws.Cells(Row, Column) 'copy paste the wall picture to the 1 place
                picturePath = ThisWorkbook.path & "\PictureInput\Wall_F.png"
                Set targetCell = ws.Cells(Row, Column)
                pictureName = "Wall_F"
                picset picturePath, targetCell, pictureName, 0, ws
                
            ElseIf wsH.Cells(Row, Column) = 8 Then
            
                'Sheets("photo").[a2].Copy Destination:=ws.Cells(Row, Column) 'copy paste the cashier picture to the 8 place
                 picturePath = ThisWorkbook.path & "\PictureInput\PoS.png"
                Set targetCell = ws.Cells(Row, Column)
                pictureName = "PoS"
                picset picturePath, targetCell, pictureName, 0, ws
            ElseIf wsH.Cells(Row, Column) = 0 Or wsH.Cells(Row, Column) = 6 Then
                'Sheets("photo").[a3].Copy Destination:=ws.Cells(Row, Column) 'copy paste the floor picture to the 0 ,6 place
                picturePath = ThisWorkbook.path & "\PictureInput\Ground.png"
                Set targetCell = ws.Cells(Row, Column)
                pictureName = "Ground"
                picset picturePath, targetCell, pictureName, 0, ws
            ElseIf wsH.Cells(Row, Column) = 4 Then
                'Sheets("photo").[a4].Copy Destination:=ws.Cells(Row, Column) 'copy paste the door picture to the 4 place
                picturePath = ThisWorkbook.path & "\PictureInput\Door.png"
                Set targetCell = ws.Cells(Row, Column)
                pictureName = "PoS"
                picset picturePath, targetCell, pictureName, 0, ws
            ElseIf wsH.Cells(Row, Column) = 3 Then
                'Sheets("photo").[a5].Copy Destination:=ws.Cells(Row, Column) 'copy paste the shelf picture to the 3 place
                picturePath = ThisWorkbook.path & "\PictureInput\Shelves.png"
                Set targetCell = ws.Cells(Row, Column)
                pictureName = "Shelves"
                picset picturePath, targetCell, pictureName, 0, ws
            ElseIf wsH.Cells(Row, Column) = 7 Then
                'Sheets("photo").[b4].Copy Destination:=ws.Cells(Row, Column) 'copy paste the entrance door picture to the 7 place
                picturePath = ThisWorkbook.path & "\PictureInput\Door.png"
                Set targetCell = ws.Cells(Row, Column)
                pictureName = "Door"
                picset picturePath, targetCell, pictureName, 0, ws
            ElseIf wsH.Cells(Row, Column) = 9 Then
                'Sheets("photo").[a6].Copy Destination:=ws.Cells(Row, Column) 'copy paste the open picture to the 9 place
                picturePath = ThisWorkbook.path & "\PictureInput\Opening.png"
                Set targetCell = ws.Cells(Row, Column)
                pictureName = "Opening"
                picset picturePath, targetCell, pictureName, 0, ws
           End If
        Next
    Next
    
End Sub
