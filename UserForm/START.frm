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





Private Sub MENU_Click()
    START.Hide
    RULES.Show
End Sub

Private Sub RESTART_Click()
    'set your shop name
    shopName = InputBox("Please enter the shop name:")
    MsgBox "Shop Name: " & shopName, vbInformation, "Shop Confirmation"
    'initialize the whole shop
    
   
    Dim marketws As Worksheet
    Set marketws = ThisWorkbook.Sheets("Goods")
    marketws.Range("h1:h38") = 0 'stock=0
    
    Dim financews As Worksheet
    Set financews = ThisWorkbook.Sheets("Finance")
    financews.Range("B:D").ClearContents 'stock=0
    financews.Cells(1, "b") = "before bal"
    financews.Cells(2, "b") = financews.Cells(2, "a")
    financews.Cells(1, "c") = "code of order"
    financews.Cells(1, "d") = "each price"
    financews.Cells(7, "a") = 0
    
    reset 'market reset
    'interface?
    Dim interfacews As Worksheet
    Set interfacews = ThisWorkbook.Sheets("Interface")
    interfacews.Cells(2, "a") = 0 'time
    interfacews.Cells(2, "c") = 0 'exp
    interfacews.Cells(2, "d") = 1   'level
    interfacews.Cells(2, "h") = 0
    interfacews.Cells(2, "i") = 0
    interfacews.Cells(2, "j") = 0
    interfacews.Cells(2, "m") = 1 'day
    InitializeCellsToSquareInRange
    InsertPictureInCell
    startMain
   
    
    START.Hide
    
    'put name on the right of the interface
End Sub


