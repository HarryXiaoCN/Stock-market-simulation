Attribute VB_Name = "SMS_Var_Public"
Public ����(15) As ���
Public �� As ���
Public GitHub As ��˾
Public ������Ϣ��(4, 1000) As �鱨
Public ������Ϣ��(4, 1000) As �鱨
Public ����Ȩ�� As Boolean
Public ��������, ��ǰ�׶�, ��ǰ�غ�, ������� As String
Public Function ��Ϸ��ʼ��()
��ǰ�غ� = 1
��ǰ�׶� = "׼���׶�"
������� = "ׯ��"
��Ϸ��ʼ��_������״̬��Ϣ����
��ҳ�ʼ��
��˾��ʼ��
End Function
Public Function ��ҳ�ʼ��()
Dim ��ݷ�����(15) As Boolean
Dim ��ҺŻ���, i As Long
Randomize
��ҺŻ��� = Int(Rnd * (14)) + 1
��ݷ�����(��ҺŻ���) = True
ׯ�ҳ�ʼ�� (��ҺŻ���)
For i = 0 To 1
    ��ҺŻ��� = Int(Rnd * (14)) + 1
    Do Until ��ݷ�����(��ҺŻ���) = False
        ��ҺŻ��� = Int(Rnd * (14)) + 1
    Loop
    ���ʳ�ʼ�� (��ҺŻ���)
Next
For i = 1 To 14
    If ��ݷ�����(i) = False Then
        ɢ����ʼ�� (i)
    End If
Next
End Function
Public Function ׯ�ҳ�ʼ��(ByVal ��Һ� As Long)
With ����(��Һ�)
    .��� = "ׯ��"
    .�ʽ� = 10000
    .��Ϣ���ȼ� = 1
End With
End Function
Public Function ���ʳ�ʼ��(ByVal ��Һ� As Long)
With ����(��Һ�)
    .��� = "����"
    .�ʽ� = 1000
    .��Ϣ���ȼ� = 2
End With
End Function
Public Function ɢ����ʼ��(ByVal ��Һ� As Long)
With ����(��Һ�)
    .��� = "ɢ��"
    .�ʽ� = 10
    .��Ϣ���ȼ� = 3
End With
End Function
Public Function ��˾��ʼ��()
With GitHub
    .��ͨ���� = 20000
    .�ܹ��� = 30000
End With
End Function

