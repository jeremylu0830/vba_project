VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PoS_UF 
   Caption         =   "UserForm1"
   ClientHeight    =   6504
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13284
   OleObjectBlob   =   "PoS_UF.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "PoS_UF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton1_Click()
    add_order_item
End Sub

Private Sub CommandButton2_Click()
    delete_order_item
End Sub

Private Sub CommandButton3_Click()

    settle_accounts
    
End Sub



Private Sub Frame2_Click()

End Sub


Private Sub Label1_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label17_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub ListBox1_Click()
    show_product
    
End Sub

Private Sub ListBox2_Click()
    show_product_price
End Sub
    
Private Sub ListBox3_Click()

End Sub

Private Sub SpinButton1_Change()
    PoS_UF.TextBox4.value = PoS_UF.SpinButton1
End Sub

Private Sub TextBox4_Change()
    Dim value As Long
    ' 當 TextBox 中的值改變時，更新 SpinButton 的值
    ' 確保 TextBox 中的值是有效的數字
    With PoS_UF
    If IsNumeric(.TextBox4.value) Then
        value = CLng(.TextBox4.value)
        ' 確保值在 SpinButton 的範圍內
        If value >= .SpinButton1.Min And value <= .SpinButton1.Max Then
            .SpinButton1.value = value
        End If
    End If
    
    End With
    
End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub UserForm_Initialize()
    
    show_category

End Sub
