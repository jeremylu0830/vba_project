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

Private Sub CommandButton1_Click()
    Set st = Sheets("sales_records")
    Financial.TextBox3.value = WorksheetFunction.SumIf(st.[a:a], Financial.ComboBox1.Text, st.[f:f]) 'compute income
    'Financial.TextBox1.value = WorksheetFunction.SumIf(st.[a:a], Financial.ComboBox1.Text, st.[d:d]) 'compute cost
    'Financial.TextBox4.value = Financial.TextBox3.value - Financial.TextBox1.value 'compute revenue
    Financial.ListBox1.Clear
    myend = st.Cells(1, 1).End(xlDown).Row
    Financial.ListBox1.ColumnCount = 5
    k = 0
    For i = 2 To myend
        If st.Cells(i, 1) = Financial.ComboBox1.Text Then
            Financial.ListBox1.AddItem
            Financial.ListBox1.List(k, 0) = st.Cells(i, 2)
            Financial.ListBox1.List(k, 1) = st.Cells(i, 3)
            Financial.ListBox1.List(k, 2) = st.Cells(i, 4)
            Financial.ListBox1.List(k, 3) = st.Cells(i, 5)
            Financial.ListBox1.List(k, 4) = st.Cells(i, 6)
            k = k + 1
        End If
    Next
    
End Sub
Private Sub CommandButton2_Click()
    'export chart
    Set st = Sheets("sales_records")
    st.Cells(1, 9) = "cost"
    st.Cells(1, 10) = "income"
    st.Cells(1, 11) = "revenue"
    st.Cells(2, 9) = Financial.TextBox1.value
    st.Cells(2, 10) = Financial.TextBox3.value
    st.Cells(2, 11) = Financial.TextBox4.value
    
    Set dataRange = st.Range("I1:K2")
    
    ' create chart
    Set chartObj = st.ChartObjects.Add(Left:=100, width:=375, Top:=50, height:=225)
    chartObj.Chart.SetSourceData Source:=dataRange
    chartObj.Chart.ChartType = xlColumnClustered
    
    ' set chart title
    chartObj.Chart.HasTitle = True
    chartObj.Chart.ChartTitle.Text = "Financial Data Chart"
    
    chartObj.Chart.FullSeriesCollection(1).ApplyDataLabels
    chartObj.Chart.Legend.LegendEntries(1).Delete
    
    chartObj.Name = "MyChart"
    
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Initialize()
    Set st = Sheets("sales_records")
    Financial.ComboBox1.Clear
    st.Columns("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=st.Range("m1"), Unique:=True
    myend = st.[m1].End(xlDown).Row
    For i = 2 To myend
        Financial.ComboBox1.AddItem st.Cells(i, "m")
    Next
    Columns("m:m").Clear

End Sub
Private Sub UserForm_Terminate()
    Set st = Sheets("sales_records")
    'delete "MyChart"
    On Error Resume Next
    Set chartObj = st.ChartObjects("MyChart")
    If Not chartObj Is Nothing Then
        chartObj.Delete
    End If
    On Error GoTo 0
End Sub
