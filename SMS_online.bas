Attribute VB_Name = "Airwar_online"
Public Local_State, TCPNum As Long
Public Function Server_GetData(ByVal GetData As String)

End Function
Public Function Server_SendData(ByRef StrData As String)
Dim i As Long
For i = 1 To TCPNum
    If Form3.Winsock1(i).State = 7 Then
        Form3.Winsock1(i).SendData StrData
    End If
Next
End Function
Public Function Client_GetData()

End Function
Public Function Client_SendData(ByRef StrData As String)
Form4.Winsock1.SendData StrData
End Function
Public Function ����״̬����(ByVal SID As Long) As String
Select Case SID
    Case 0
        ����״̬���� = "δ����"
    Case 1
        ����״̬���� = "��״̬"
    Case 2
        ����״̬���� = "�ȴ�����"
    Case 3
        ����״̬���� = "���ӹ���"
    Case 4
        ����״̬���� = "����������..."
    Case 5
        ����״̬���� = "���������ɹ�"
    Case 6
        ����״̬���� = "��������..."
    Case 7
        ����״̬���� = "������"
    Case 9
        ����״̬���� = "���Ӵ���"
End Select
End Function
