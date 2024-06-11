VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OPEN_CLOSE 
   Caption         =   "UserForm1"
   ClientHeight    =   1560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "OPEN_CLOSE.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "OPEN_CLOSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()

End Sub

Private Sub NoButton_Click()
    OPEN_CLOSE.Hide
End Sub

Private Sub YseButton_Click()
    'statr timer and let the custmer in
    UpdateTime
    OPEN_CLOSE.Hide
    RandomSelectCellWithNumbers
End Sub
