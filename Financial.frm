VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Financial 
   Caption         =   "Financial"
   ClientHeight    =   5664
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10548
   OleObjectBlob   =   "Financial.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "Financial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
    
End Sub

Private Sub ComboBox1_Click()

End Sub

Private Sub CommandButton1_Click()
    Financial.TextBox3.Value = WorksheetFunction.SumIf(Sheet4.[a:a], Financial.ComboBox1.Text, Sheet4.[c:c])
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox3_Change()
    
End Sub

Private Sub UserForm_Click()
   
    
End Sub

Private Sub UserForm_Initialize()
    Financial.ComboBox1.Clear
    Columns("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("L1"), Unique:=True
    myend = Sheet4.[L1].End(xlDown).Row
    For i = 2 To myend
        Financial.ComboBox1.AddItem Sheet4.Cells(i, "l").Value
    Next
End Sub
