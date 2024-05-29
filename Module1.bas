Attribute VB_Name = "Module1"
Dim currentrow As Integer
Dim currentcol As Integer
Dim arrowrow As Integer
Dim arrowcol As Integer
Dim gamestatus As String
'BUG: �q�ʧ��F�観�ɭԸ}�L�L�k���ʡA�h�ìO��w�����D
'�ݳB�z:select������֮w�s�A�i�H��^�h���\��B�c�l�a�۶]����


Sub InitializeCellsToSquareInRange()
    
    'gamestatus = "move"
    Dim ws As Worksheet
    Dim rowHeight As Double
    Dim colWidth As Double
    Dim targetRange As Range
    Dim i As Long

    ' �]�m�u�@���H
    Set ws = ThisWorkbook.Sheets("Warehouse")
    Set targetRange = ws.Range("A1:Z100") ' �]�m�A�Ʊ�վ㪺�d��

    ' �]�w�氪
    rowHeight = 20 ' �o�̳]�w�氪�A�A�i�H�ھڻݭn�վ�
    colWidth = rowHeight * 0.1428 ' ����C�e�A�ϳ椸�汵�񥿤��

    ' �]�m�ؼнd�򤺪��氪��
    targetRange.Rows.rowHeight = rowHeight
    ' �]�m�ؼнd�򤺪��C�e��
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
    
    '�q�檺form ��l��
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
    InitializeCellsToSquareInRange '��l�Ʈ�l����� �P�I���ܮw
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Warehouse")
    'ws2.Activate
    ws2.Unprotect
    deleteallpicture
    
    'wall
    k = 1
    For i = 1 To 20
        For j = 1 To 20
            If �u�@��3.Cells(i, j) = 1 Then
                picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\wall.png" ' ��אּ�A���Ϥ���������|
                Set targetCell = ws2.Cells(i, j)
                pictureName = "wall" & k
                picturesetting picturePath, targetCell, pictureName, 0
                k = k + 1
            End If
        Next
    Next
     'shelf1
    picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\shelf2.png" ' ��אּ�A���Ϥ���������|
    Set targetCell = ws2.Range("B3:O9")
    pictureName = "shelf"
    picturesetting picturePath, targetCell, pictureName, 0
    

    'milk
    'picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\milk.png"
    'Set targetCell = ws2.Range("E3:F4") ' ��אּ�A�Q���J�Ϥ����椸��
    'pictureName = "milk"
    'picturesetting picturePath, targetCell, pictureName, 0
    
    'test
    'egg
    'picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\egg.png"
    'Set targetCell = ws2.Range("G3:H4") ' ��אּ�A�Q���J�Ϥ����椸��
    'pictureName = "egg"
    'picturesetting picturePath, targetCell, pictureName, 0
    
    'candy
    'picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\candy.png"
    'Set targetCell = ws2.Range("i3:j4") ' ��אּ�A�Q���J�Ϥ����椸��
    'pictureName = "candy"
    'picturesetting picturePath, targetCell, pictureName, 0
    
    'cola
    'picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\cola.png"
    'Set targetCell = ws2.Range("k3:l4") ' ��אּ�A�Q���J�Ϥ����椸��
    'pictureName = "cola"
    'picturesetting picturePath, targetCell, pictureName, 0
   
    
    'where is me
    picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\me.png"
    Set targetCell = ws2.Cells(2, 2) ' ��אּ�A�Q���J�Ϥ����椸��
    pictureName = "me"
    picturesetting picturePath, targetCell, pictureName, 0
    currentrow = 2
    currentcol = 2
   
    'leave sign
    picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\leave.png"
    Set targetCell = ws2.Range("S19:t20")
    pictureName = "leave"
    picturesetting picturePath, targetCell, pictureName, 0
    
    'cart sign
    picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\cart.png"
    Set targetCell = ws2.Range("b18:c19")
    pictureName = "cart"
    picturesetting picturePath, targetCell, pictureName, 0
    
    
   
End Sub
Sub AfterForm()
    gamestatus = "move"
    'InitializeCellsToSquareInRange '��l�Ʈ�l����� �P�I���ܮw
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
    picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\me.png"
    Set targetCell = ws2.Cells(17, "d") ' ��אּ�A�Q���J�Ϥ����椸��
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
        .Locked = True ' ��w�Ϥ�
        .PrintObject = False ' �ϹϤ����i�襤
        
    End With
    ws2.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    
    
    
End Sub
'�R������
Sub deleteallpicture()


    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Warehouse")
    'ws2.Activate
    Dim pic As Shape
    'test delete
    ' �R���u�@�����Ҧ��Ϥ�
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
            picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\me.png"
            Set targetCell = ws2.Cells(currentrow, currentcol) ' ��אּ�A�Q���J�Ϥ����椸��
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
            picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\point.png"
            Set targetCell = ws2.Cells(arrowrow, arrowcol) ' ��אּ�A�Q���J�Ϥ����椸��
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
    'If �u�@��3.Cells(currentrow, currentcol).Value = 2 Then
        
        picturePath = "C:\Users\NUTC\OneDrive\�ୱ\PictureInput\point.png"
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
        
        PickUp2 '��ԣ�n�I�s�⦸...
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

