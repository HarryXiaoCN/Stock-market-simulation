Attribute VB_Name = "SMS_WorldRuler"
Public Function ������Ϣ����ѡȡ() As String
Dim Temp As Single
Randomize
Temp = Rnd - 0.5
If Temp >= 0 Then ������Ϣ����ѡȡ = "����" Else ������Ϣ����ѡȡ = "����"
End Function
