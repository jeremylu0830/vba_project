Attribute VB_Name = "PoS"
Sub show_category()

    'init
    
    Dim rng As Range
    Dim dataRange As Range
    Dim ws As Worksheet
    Dim uf As UserForm
    
    'plugin
    Set uf = PoS_UF
    Set ws = Sheets("product_info")
    
    
    'main
    
    uf.ListBox1.Clear
    
    product_lastrow = Sheets("product_info").[b1048576].End(xlUp).Row
    
    ws.Range("B1:B" & CStr(product_lastrow)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Columns( _
        "F:F"), Unique:=True
        
    
    lastrow = ws.[f1].End(xlDown).Row
    uf.ListBox1.List = ws.Range("F2:F" & CStr(lastrow)).Cells.value
    
End Sub

Sub show_product()
          
    Dim ws As Worksheet
    Dim uf As UserForm
    Dim product_info, category
    Dim products As Collection
    
   'plugin
     Set uf = PoS_UF
    Set ws = Sheets("product_info")
    Set products = New Collection
    
    'main
    
    uf.ListBox2.Clear
    
    lastrow = ws.[c1].End(xlDown).Row
    category_info = ws.Range("B2:B" & CStr(lastrow))
    category = uf.ListBox1.value
    
   
   
    For Each cell In ws.Range("B2:B" & CStr(lastrow))
        If cell.value = category Then
            products.Add cell.Offset(0, 1).value
        End If
    Next cell
    
    For Each Item In products
        uf.ListBox2.AddItem Item
    Next Item
    
End Sub

Sub show_product_price()

    Dim category, product, price
    Dim product_info
    Dim products As Collection
    
   'plugin
     Set uf = PoS_UF
    Set ws = Sheets("product_info")


    With uf
    category = .ListBox1.value
    'category = "電子產品"
    product = .ListBox2.value
    'product = "手機"
    lastrow = ws.[c1].End(xlDown).Row
    End With
    
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)
        If (cell.value = category Or category = "") And (cell.Offset(0, 1).value = product Or product = "") Then
            price = cell.Offset(0, 2).value
            GoTo labe11
        End If
    Next cell
    
labe11:
With uf
.Label6 = price
End With
    
    
End Sub


Sub add_order_item()
    
    Dim product, price, order_number, total
    Dim product_info, category
    Dim products As Collection
    
   'plugin
     Set uf = PoS_UF
    Set ws = Sheets("product_info")
    
    total = CDbl(uf.Label13.Caption)
    
    With uf
        product = CStr(.ListBox2.value)
        price = CStr(.Label6.Caption)
        order_number = CStr(.TextBox4.value)
    End With
    
    
    With uf.ListBox3
            .AddItem
            .List(.ListCount - 1, 0) = product
            .List(.ListCount - 1, 1) = price
            .List(.ListCount - 1, 2) = order_number
    End With
    
    uf.Label13.Caption = total + price * order_number
    
End Sub

Sub delete_order_item()
    
    Dim total, price, order_number
    Dim product_info, category
    Dim products As Collection
    
   'plugin
     Set uf = PoS_UF
    Set ws = Sheets("product_info")
   
    
    total = CDbl(uf.Label13.Caption)
    
    With uf.ListBox3
     price = .List(.ListIndex, 1)
     order_number = .List(.ListIndex, 2)
     UserForm1.Label13.Caption = total - price * order_number
      .RemoveItem .ListIndex
    End With
    
End Sub

Sub settle_accounts()
    Dim account, total, ID As Integer
    Dim product_info, category
    Dim products As Collection
    
   'plugin
    Set uf = PoS_UF
    Set ws = Sheets("sales_records")

    
    id_lastrow = ws.[a1048576].End(xlUp).Row
    
    lastrow = ws.[b1048576].End(xlUp).Row
    
    If Application.WorksheetFunction.IsText(ws.Cells(id_lastrow, 1)) Then
        ID = 1
    Else
        ID = ws.Cells(id_lastrow, 1).value + 1
    End If
    
    account = uf.ListBox3.List
    
    ws.Cells(lastrow + 1, 1) = ID
    
    ws.Cells(lastrow + 1, 2).Resize(UBound(account) + 1, 3) = account
    
    total = CDbl(uf.Label13.Caption)
    
    
    ws.Cells(lastrow + 1, 5) = total
    
    
    uf.ListBox3.Clear
    uf.Label13.Caption = "0"
    
    
    
End Sub


