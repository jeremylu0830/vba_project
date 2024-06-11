VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectQuantity 
   Caption         =   "SELECT QUANTITY"
   ClientHeight    =   1980
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   3744
   OleObjectBlob   =   "SelectQuantity.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "SelectQuantity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PickedPrice As Integer


Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub CommandButton1_Click()
    
    'out of selection
    'MsgBox TextBox1.Value
    Dim wsControl As Worksheet
    Set wsControl = Worksheets("Interface") 'later renew to Goods
    wsControl.Cells(2, "J") = PickedPrice
    wsControl.Cells(2, "I") = TextBox1.value
    wsControl.Cells(2, "H") = Label1.Caption ' would be better if use index to directly put the name
    
    SelectQuantity.Hide
    
End Sub

Private Sub ScrollBar1_Change()
    TextBox1.value = ScrollBar1.value
    
End Sub



Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Activate()
    Set ws4 = Worksheets("Goods")
    Set ws3 = Worksheets("HideWarehouse")
    ScrollBar1.Min = 1
    
        
    'use match index...
    Dim matchRow As Integer
    MsgBox Label1.Caption
    matchRow = WorksheetFunction.Match(Label1.Caption, ws4.[A1:A38], 0)
    
    ScrollBar1.Max = ws4.Cells(matchRow, "h")
    PickedPrice = ws4.Cells(matchRow, "b")
    'Label1.Caption = ws3.Cells(5, "a")
    'name to be curcel value in ws3.cells(curpos)

End Sub

