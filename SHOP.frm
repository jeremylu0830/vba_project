VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SHOP 
   Caption         =   "ORDERPAGE"
   ClientHeight    =   5698
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9058.001
   OleObjectBlob   =   "SHOP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SHOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OrderButton_Click()
    Set ws = Worksheets("Goods")
    Set ws2 = Worksheets("Finance")
    Total = 0
    For i = 1 To 36 ' the type of all products
        tempstr = ws.Cells(i, "b")
        ws.Cells(i, "e") = CInt(Right(tempstr, Len(tempstr) - 1)) * ws.Cells(i, "d")
        
    Next
    ws2.Cells(3, "a") = WorksheetFunction.Sum(ws.Range("E1:E36"))
    ws2.Cells(2, "a") = ws2.Cells(2, "a") - ws2.Cells(3, "a")
    
    
    UserForm_Initialize
    
End Sub

Private Sub UserForm_Initialize()
    'clear
    Set ws2 = Worksheets("Goods")
    ws2.Columns("D:D").ClearContents
    Me.Controls("quantity").Caption = 0
    Set ws = Worksheets("Finance")
    Me.Controls("balance").Caption = ws.Cells(2, "a")
End Sub

Sub VisibleProcess(t As Integer)

    Set ws = Worksheets("Goods")
    With MultiPage1
        For i = 0 To MultiPage1.Count - 1
            .Pages(i).Visible = False
        Next
        .Pages(t).Visible = True
        .Value = t '
        .Pages(t).Caption = ""
    End With
        Dim ctl As Control
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                ctl.Caption = ""
            End If
        Next ctl
    For i = 1 To 4
            Me.Controls("name" & i + t * 4).Caption = ws.Cells(i + 4 * t, 1).Value
            Me.Controls("price" & i + t * 4).Caption = ws.Cells(i + 4 * t, 2).Value
            'level unlock
            If ws.Cells(i + 4 * t, 3).Value <> 1 Then
                Me.Controls("add" & i + t * 4).Visible = False
            End If
    Next i
    
    
    
End Sub
Sub AddtoCart(t As Integer) 'pay attention to the range of products
    Set ws = Worksheets("Goods")
    ws.Cells(t, 4).Value = ws.Cells(t, 4).Value + 1
    Total = WorksheetFunction.Sum(ws.Range("D1:D36"))
    Me.Controls("quantity").Caption = Total
End Sub

Private Sub add1_Click() ' add to cart
    AddtoCart 1
End Sub
Private Sub add2_Click() ' add to cart
    AddtoCart 2
End Sub
Private Sub add3_Click() ' add to cart
    AddtoCart 3
End Sub
Private Sub add4_Click() ' add to cart
    AddtoCart 4
End Sub
Private Sub add5_Click() ' add to cart
    AddtoCart 5
End Sub
Private Sub add6_Click() ' add to cart
    AddtoCart 6
End Sub
Private Sub add7_Click() ' add to cart
    AddtoCart 7
End Sub
Private Sub add8_Click() ' add to cart
    AddtoCart 8
End Sub
Private Sub add9_Click() ' add to cart
    AddtoCart 9
End Sub
Private Sub add10_Click() ' add to cart
    AddtoCart 10
End Sub
Private Sub add11_Click() ' add to cart
    AddtoCart 11
End Sub
Private Sub add12_Click() ' add to cart
    AddtoCart 12
End Sub
Private Sub add13_Click() ' add to cart
    AddtoCart 13
End Sub
Private Sub add14_Click() ' add to cart
    AddtoCart 14
End Sub
Private Sub add15_Click() ' add to cart
    AddtoCart 15
End Sub
Private Sub add16_Click() ' add to cart
    AddtoCart 16
End Sub
Private Sub add17_Click() ' add to cart
    AddtoCart 17
End Sub
Private Sub add18_Click() ' add to cart
    AddtoCart 18
End Sub
Private Sub add19_Click() ' add to cart
    AddtoCart 19
End Sub
Private Sub add20_Click() ' add to cart
    AddtoCart 20
End Sub
Private Sub add21_Click() ' add to cart
    AddtoCart 21
End Sub
Private Sub add22_Click() ' add to cart
    AddtoCart 22
End Sub
Private Sub add23_Click() ' add to cart
    AddtoCart 23
End Sub
Private Sub add24_Click() ' add to cart
    AddtoCart 24
End Sub
Private Sub add25_Click() ' add to cart
    AddtoCart 25
End Sub
Private Sub add26_Click() ' add to cart
    AddtoCart 26
End Sub
Private Sub add27_Click() ' add to cart
    AddtoCart 27
End Sub
Private Sub add28_Click() ' add to cart
    AddtoCart 28
End Sub
Private Sub add29_Click() ' add to cart
    AddtoCart 29
End Sub
Private Sub add30_Click() ' add to cart
    AddtoCart 30
End Sub
Private Sub add31_Click() ' add to cart
    AddtoCart 31
End Sub
Private Sub add32_Click() ' add to cart
    AddtoCart 32
End Sub
Private Sub add33_Click() ' add to cart
    AddtoCart 33
End Sub
Private Sub add34_Click() ' add to cart
    AddtoCart 34
End Sub
Private Sub add35_Click() ' add to cart
    AddtoCart 35
End Sub
Private Sub add36_Click() ' add to cart
    AddtoCart 36
End Sub
Private Sub CommandButton1_Click() 't=0
    
    VisibleProcess 0

End Sub

Private Sub CommandButton2_Click() 't=1
    
    VisibleProcess 1
         
End Sub

Private Sub CommandButton3_Click() 't=1
    
    VisibleProcess 2
         
End Sub
Private Sub CommandButton4_Click() 't=1
    
    VisibleProcess 3
         
End Sub
Private Sub CommandButton5_Click() 't=1
    
    VisibleProcess 4
         
End Sub
Private Sub CommandButton6_Click() 't=1
    
    VisibleProcess 5
         
End Sub
Private Sub CommandButton7_Click() 't=1
    
    VisibleProcess 6
         
End Sub
Private Sub CommandButton8_Click() 't=1
    
    VisibleProcess 7
         
End Sub
Private Sub CommandButton9_Click() 't=1
    
    VisibleProcess 8
         
End Sub



