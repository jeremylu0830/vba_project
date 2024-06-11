Attribute VB_Name = "Customer"
Type NodeType
    Row As Long
    Col As Long
End Type

Sub CreateRectangleShape()
    Dim ws As Worksheet
    Dim wsM As Worksheet
    Dim rowHeight As Double
    Dim colWidth As Double
    Dim targetRange As Range
    Dim i As Long
    Dim shp As Shape
    Dim topLeftCell As Range
    Dim bottomRightCell As Range
    Dim leftPosition As Single
    Dim topPosition As Single
    Dim width As Single
    Dim height As Single

    ' Set up the "Customermove" worksheet
    Set ws = ThisWorkbook.Sheets("Customermove")
    Set targetRange = ws.Range("A1:Z100")

    rowHeight = 20
    colWidth = rowHeight * 0.1428
    targetRange.rows.rowHeight = rowHeight
    For i = 1 To targetRange.Columns.Count
        targetRange.Columns(i).ColumnWidth = colWidth
    Next i

    ws.Cells(2, 2) = 1000
    ws.Cells(19, "s") = 8

    ' Create and position the rectangle shape in "Customermove"
    Set topLeftCell = ws.Range("B2")
    Set bottomRightCell = ws.Range("B2")
    leftPosition = topLeftCell.Left
    topPosition = topLeftCell.Top
    width = bottomRightCell.Left + bottomRightCell.width - leftPosition
    height = bottomRightCell.Top + bottomRightCell.height - topPosition
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPosition, topPosition, width, height)
    shp.Name = "triangle"

    ' Set up the "market" worksheet
    Set wsM = ThisWorkbook.Sheets("market")
    wsM.Unprotect
    
    For Each pic In wsM.Shapes
        If pic.Name = "customer" Then
            pic.Delete
            Exit For
        End If
    Next pic
    
    'customer
    picturePath = ThisWorkbook.path & "\PictureInput\cart.png"
    Set targetCell = wsM.Cells(2, "B")
    pictureName = "customer"
    picset picturePath, targetCell, pictureName, 0, wsM
        

    
End Sub

Sub Delay(seconds As Single)
    Dim endTime As Single
    endTime = Timer + seconds
    Do While Timer < endTime
        DoEvents
    Loop
End Sub


Sub FindShortestPath2(Max As Integer)
    Dim ws As Worksheet
    Dim maze() As Variant
    Dim rows As Long, cols As Long
    Dim i As Long, j As Long
    Dim startRow As Long, startCol As Long
    Dim endRow As Long, endCol As Long
    Dim queue As Collection
    Dim current As Variant
    Dim directions As Variant
    Dim newRow As Long, newCol As Long
    Dim visited() As Boolean
    Dim distance() As Long
    Dim previous() As NodeType
    Dim pathFound As Boolean
    Dim path As Collection
    Dim node As NodeType
    
    directions = Array(Array(-1, 0), Array(1, 0), Array(0, -1), Array(0, 1))
    

    Set ws = ThisWorkbook.Sheets("Customermove")
    maze = ws.Range("A1:W23").value
    rows = UBound(maze, 1)
    cols = UBound(maze, 2)
    
 
    ReDim visited(1 To rows, 1 To cols)
    ReDim distance(1 To rows, 1 To cols)
    ReDim previous(1 To rows, 1 To cols)
    
    m = 1001
    ' If cust finish to buy goods, then go to the cashier(destiantion=8)
    If Max = 1 Then
        m = 8
    End If

    For i = 1 To rows
        For j = 1 To cols
            If maze(i, j) = 1000 Then
                startRow = i
                startCol = j
            ElseIf maze(i, j) = m Then
                endRow = i
                endCol = j
            End If
            visited(i, j) = False
            distance(i, j) = -1
            previous(i, j).Row = -1
            previous(i, j).Col = -1
        Next j
    Next i
    

    Set queue = New Collection
    queue.Add Array(startRow, startCol)
    visited(startRow, startCol) = True
    distance(startRow, startCol) = 0

    pathFound = False
    Do While queue.Count > 0
        current = queue(1)
        queue.Remove 1
        If current(0) = endRow And current(1) = endCol Then
            pathFound = True
            Exit Do
        End If
        For Each direction In directions
            newRow = current(0) + direction(0)
            newCol = current(1) + direction(1)
            If newRow >= 1 And newRow <= rows And newCol >= 1 And newCol <= cols Then
                If Not visited(newRow, newCol) And maze(newRow, newCol) <> 1 Then
                    queue.Add Array(newRow, newCol)
                    visited(newRow, newCol) = True
                    distance(newRow, newCol) = distance(current(0), current(1)) + 1
                    previous(newRow, newCol).Row = current(0)
                    previous(newRow, newCol).Col = current(1)
                End If
            End If
        Next direction
    Loop
    

    If pathFound Then
        'MsgBox "the shortest route is: " & distance(endRow, endCol)
        

        Set path = New Collection
        node.Row = endRow
        node.Col = endCol
        Do
            path.Add Array(node.Row, node.Col)
            node = previous(node.Row, node.Col)
        Loop Until node.Row = -1 And node.Col = -1
        

        Dim pathRow As Long
        pathRow = 1
        For i = path.Count To 1 Step -1
            current = path(i)
            ws.Cells(pathRow, 24).value = "Next cells"
            ws.Cells(pathRow, 25).value = current(0)
            ws.Cells(pathRow, 26).value = current(1)
            pathRow = pathRow + 1
        Next i
        
    Else
        MsgBox "no searchable route"
    End If
End Sub


' customer movement
Sub RandomMoveToCellInTime(Max As Integer)
    Dim ws As Worksheet
    Dim picturePath As String
    Dim shp As Shape
    Dim startX As Integer, startY As Integer
    Dim endX As Integer, endY As Integer
    Dim X As Single, Y As Single
    Dim step_x As Integer, step_y As Integer
    
    ' Set the worksheet and shape
    Set ws = ThisWorkbook.Sheets("Customermove")
    Set shp = ws.Shapes("triangle")

    'declaration for the customer cell in market sheet
    Dim targetCell As Range
    Dim pic As Shape
    Set wsM = ThisWorkbook.Sheets("market")

    startX = 2
    startY = 2
    
    ' Get the top-left coordinates of the start cell
    X = ws.Cells(startX, startY).Left
    Y = ws.Cells(startX, startY).Top
    shp.Left = X
    shp.Top = Y

    ' Set the delay time for each step (in seconds)
    Dim delayTime As Single
    delayTime = 0.1 ' Adjust as needed

    Step = 1
    
    first_x = ws.Cells(Step, 25)
    first_y = ws.Cells(Step, 26)
    
    Do While ws.Cells(Step, 25) <> ""
        
        startX = ws.Cells(Step, 25)
        startY = ws.Cells(Step, 26)
        
        ' Update the position of the shape
        X = ws.Cells(startX, startY).Left
        Y = ws.Cells(startX, startY).Top
        
        shp.Left = X
        shp.Top = Y
        
        For Each pic In wsM.Shapes
            If pic.Name = "customer" Then
                pic.Delete
                Exit For
            End If
        Next pic
       
        picturePath = ThisWorkbook.path & "\PictureInput\cart.png"
        Set targetCell = wsM.Cells(startX, startY)
        pictureName = "customer"
        picset picturePath, targetCell, pictureName, angle, wsM

        
        ' Short delay between each step
        Delay 0.1
        Step = Step + 1
    Loop
    ws.Range("X:Z").ClearContents
    
    ws.Cells(first_x, first_y) = 0
    If Max <> 1 Then
        ws.Cells(startX, startY) = 1000
    End If
End Sub




'click here to activate
Sub RandomSelectCellWithNumbers()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim cellsWithNumbers As Collection
    Dim randomIndex As Long
    Dim randomCell As Range
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim itemrow As Long
    Dim itemcol As Long
    
    
    Item = 1
    r = 1

    Set ws = ThisWorkbook.Sheets("HidemarketPrice")
    Set ouo = ThisWorkbook.Sheets("Customermove")
    Set rng = ws.Range("A1:Z20")

    Set cellsWithNumbers = New Collection
    
    'how many goods the customer would buy
    'Max = Int(cellsWithNumbers.Count * Rnd) + 1
    'only can by max-1
    Max = 3
    CreateRectangleShape
    
    For Each cell In rng
        If IsNumeric(cell.value) And cell.value <> "" Then
            'if the shelf at the current cutomer point has the number(price), then the arry in cellsWithNumbers pick up random value
            'cellsWithNumbers function like an array
            cellsWithNumbers.Add cell
        End If
    Next cell
    
    Do While Max <> 0
        If cellsWithNumbers.Count > 0 Then
            Randomize 'random
            randomIndex = Int(cellsWithNumbers.Count * Rnd) + 1
            Set randomCell = cellsWithNumbers(randomIndex)
            
            rowIndex = randomCell.Row
            colIndex = randomCell.Column
         End If
    
        itemrow = rowIndex
        itemcol = colIndex
    
        If ouo.Cells(rowIndex + 1, colIndex) = 0 Then
            rowIndex = rowIndex + 1
        ElseIf ouo.Cells(rowIndex - 1, colIndex) = 0 Then
            rowIndex = rowIndex - 1
        ElseIf ouo.Cells(rowIndex, colIndex + 1) = 0 Then
            colIndex = colIndex + 1
        ElseIf ouo.Cells(rowIndex, colIndex - 1) = 0 Then
            colIndex = colIndex - 1
        End If
        
        'customer buying position
        ouo.Cells(rowIndex, colIndex) = 1001
        
        FindShortestPath2 (Max)
        
        'move
        RandomMoveToCellInTime (Max)
        
        'Randomize
        'If 4 < Int((7 * Rnd) + 1) Then
            Max = Max - 1
            ouo.Cells(29, Item) = Sheets("HidemarketQuantity").Cells(itemrow, itemcol)
            ouo.Cells(30, Item) = Sheets("HidemarketPrice").Cells(itemrow, itemcol)
            ouo.Cells(31, Item) = Sheets("goodCust").Cells(itemrow, itemcol)
            ouo.Cells(28, Item) = r
            Item = Item + 1
        'End If
        
       If randomIndex > 0 And randomIndex <= cellsWithNumbers.Count Then
            cellsWithNumbers.Remove randomIndex
        Else
            MsgBox "Error: Invalid index - " & randomIndex
        End If

        r = r + 1
    Loop
    
    For Each cell In rng
        If IsNumeric(cell.value) And cell.value = 1001 Then
            cell.value = 0
        End If
        If IsNumeric(cell.value) And cell.value = 1000 Then
            cell.value = 8
        End If
    Next cell
    
    ouo.Cells(2, 2) = 1000
    ouo.Cells(rowIndex, colIndex) = 0
    
    'delete shape(rectangle)
    ouo.Shapes("triangle").Delete
    
    display_order


'call another sub in here!!!!, may put the function after this like accounting or find a change(money funciton)
End Sub

Sub display_order()
    Dim order
    Dim product, quantity, price
    Set ws = Sheets("Customermove")
    
    maxProduct = ws.Cells(28, 1).End(xlToRight).Column
    ReDim order(1 To maxProduct)
    
    For i = 1 To ws.Cells(28, 1).End(xlToRight).Column
        
            
            product = ws.Cells(31, i)
            quantity = ws.Cells(29, i).value
            
            order(i) = product & "¦@" & quantity & "­Ó"
            
            
        
    Next
    
    
    
    msg = Join(order, vbCr)
    
        
    MsgBox prompt:=msg, Title:="Customer order "
    

End Sub
