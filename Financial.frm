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
Private Sub CommandButton1_Click()
    'Financial.TextBox1.Clear
    'Financial.TextBox3.Clear
    'Financial.TextBox4.Clear
    Financial.TextBox3.Value = WorksheetFunction.SumIf(Sheet4.[a:a], Financial.ComboBox1.Text, Sheet4.[c:c])
    Financial.TextBox1.Value = WorksheetFunction.SumIf(Sheet4.[a:a], Financial.ComboBox1.Text, Sheet4.[d:d])
    Financial.TextBox4.Value = Financial.TextBox3.Value - Financial.TextBox1.Value
End Sub
Private Sub UserForm_Initialize()
    Financial.ComboBox1.Clear
    Columns("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("L1"), Unique:=True
    myend = Sheet4.[L1].End(xlDown).Row
    For i = 2 To myend
        Financial.ComboBox1.AddItem Sheet4.Cells(i, "l").Value
    Next
    Columns("l:l").Clear
End Sub
