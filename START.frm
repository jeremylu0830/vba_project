VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} START 
   Caption         =   "START"
   ClientHeight    =   5664
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   7212
   OleObjectBlob   =   "START.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "START"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub MENU_Click()
    START.Hide
    RULES.Show
End Sub

Private Sub RESTART_Click()
    'set your shop name
    shopName = InputBox("Please enter the shop name:")
    MsgBox "Shop Name: " & shopName, vbInformation, "Shop Confirmation"
    'initialize the whole shop
    
End Sub


