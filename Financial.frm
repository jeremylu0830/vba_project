VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Financial 
   Caption         =   "Financial"
   ClientHeight    =   5664
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10548
   OleObjectBlob   =   "Financial.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Financial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Set st = Sheet4
    Financial.TextBox3.Value = WorksheetFunction.SumIf(st.[a:a], Financial.ComboBox1.Text, st.[c:c])
    Financial.TextBox1.Value = WorksheetFunction.SumIf(st.[a:a], Financial.ComboBox1.Text, st.[d:d])
    Financial.TextBox4.Value = Financial.TextBox3.Value - Financial.TextBox1.Value
End Sub
Private Sub CommandButton2_Click()
    Set st = Sheet4
    st.Cells(1, 9) = "成本"
    st.Cells(1, 10) = "收入"
    st.Cells(1, 11) = "利潤"
    st.Cells(2, 9) = Financial.TextBox1.Value
    st.Cells(2, 10) = Financial.TextBox3.Value
    st.Cells(2, 11) = Financial.TextBox4.Value
    
    Set DataRange = st.Range("I1:K2")
    
    ' 創建圖表
    Set chartObj = st.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
    chartObj.Chart.SetSourceData Source:=DataRange
    chartObj.Chart.ChartType = xlColumnClustered
    
    ' 設置圖表標題
    chartObj.Chart.HasTitle = True
    chartObj.Chart.ChartTitle.Text = "Financial Data Chart"
    
    chartObj.Chart.FullSeriesCollection(1).ApplyDataLabels
    chartObj.Chart.Legend.LegendEntries(1).Delete
    
    chartObj.Name = "MyChart"
    
    'DataRange.Clear
    
End Sub
Private Sub UserForm_Initialize()
    Set st = Sheet4
    Financial.ComboBox1.Clear
    st.Columns("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=st.Range("L1"), Unique:=True
    myend = st.[L1].End(xlDown).Row
    For i = 2 To myend
        Financial.ComboBox1.AddItem st.Cells(i, "l").Value
    Next
    Columns("l:l").Clear

End Sub
Private Sub UserForm_Terminate()
    Set st = Sheet4
    '嘗試刪除名為 "MyChart" 的圖表
    On Error Resume Next
    Set chartObj = st.ChartObjects("MyChart")
    If Not chartObj Is Nothing Then
        chartObj.Delete
    End If
    On Error GoTo 0
End Sub
