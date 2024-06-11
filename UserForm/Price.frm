VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Price 
   Caption         =   "Enter price"
   ClientHeight    =   3696
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5916
   OleObjectBlob   =   "Price.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "Price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click() 'placestuff
    place
    Price.Hide
End Sub

Private Sub CommandButton2_Click() 'cancel
    Price.Hide
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub ScrollBar1_Change()
    TextBox1.value = ScrollBar1.value
End Sub
