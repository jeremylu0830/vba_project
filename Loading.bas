Attribute VB_Name = "Loading"
Sub downloadMap()
    ' ��ܥ[������
    LoadingUF.Show vbModeless
    ' ����� UserForm ���ܤƥߧY�ϬM�X��
    DoEvents
    
    ' ������ɶ��B�檺�N�X
    Call insertMap
    
    ' ���å[������
    Unload LoadingUF
End Sub

Sub LongRunningTask()
    ' �o�O���������ɶ��B�檺�N�X
    Dim i As Long
    For i = 1 To 1000
        ' �����@�ǭp��ξާ@
        Debug.Print i
    Next i
End Sub

