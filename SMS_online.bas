Attribute VB_Name = "SMS_online"
Public Local_State, TCPNum As Long
Public Function ��������Ϣ����(ByVal ��Ϣ As String)

End Function
Public Function ������������Ϣ����(ByVal ˭ As String, ByVal ���� As String)
Dim i As Long: Dim ��Ϣ As String
��Ϣ = "��ׯ#" & ���� & "|"
For i = 1 To 14
    If ����(i).��� = ˭ And ������.Winsock1(i).State = 7 Then ������.Winsock1(i).SendData ��Ϣ
Next
End Function
Public Function ��Ϸ��ʼ��_������״̬��Ϣ����()
Dim ��Ϣ As String
On Error GoTo Er:
For i = 1 To 14
    ��Ϣ = "��ʼ��#" & ��ǰ�غ� & "#" & ��ǰ�׶� & "#" & �������� & "#" & ����(i).��� & "#" & ����(i).��Ϣ���ȼ� & "#" & ����(i).�ʽ� & "#" & ������� & "|"
    If ������.Winsock1(i).State = 7 Then ������.Winsock1(i).SendData ��Ϣ
Next
Er:
MsgBox "��Ϸ��ʼ��_������״̬��Ϣ����>����"
End Function
Public Function ������״̬��Ϣ����()
Dim ��Ϣ As String
��Ϣ = "��̬#" & ��ǰ�غ� & "#" & ��ǰ�׶� & "#" & �������� & "|"
For i = 1 To 14
    If ������.Winsock1(i).State = 7 Then ������.Winsock1(i).SendData ��Ϣ
Next
End Function
Public Function �����Ϣ����(ByVal ��Ϣ As String)
Dim ��Ϣ����, ���ݻ���, ����: Dim i As Long
��Ϣ���� = Split(��Ϣ, "|")
For i = 0 To UBound(��Ϣ����) - 1
    ���ݻ��� = Split(��Ϣ����(i), "#")
    Select Case ��Ϣ����(0)
        Case "��ׯ"
            ���� = Split(��Ϣ����(1), "-")
            ������Ϣ��(1, Val(����(0))) = ����(1)
        Case "��ʼ��"
            ��ǰ�غ� = Val(��Ϣ����(1)): ��ǰ�׶� = ��Ϣ����(2): �������� = ��Ϣ����(3)
            ��.��� = ��Ϣ����(4): ��.��Ϣ���ȼ� = Val(��Ϣ����(5)): ��.�ʽ� = Val(��Ϣ����(6))
            ������� = ��Ϣ����(7)
            If ������� = ��.��� Then ����Ȩ�� = True
        Case "��̬"
            ��ǰ�غ� = Val(��Ϣ����(1)): ��ǰ�׶� = ��Ϣ����(2): �������� = ��Ϣ����(3)
    End Select
Next
End Function
Public Function Client_SendData(ByRef StrData As String)
�ͻ���.Winsock1.SendData StrData
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
