Attribute VB_Name = "SMS_WorldRuler"
Public Function 个股消息性质选取() As String
Dim Temp As Single
Randomize
Temp = Rnd - 0.5
If Temp >= 0 Then 个股消息性质选取 = "利好" Else 个股消息性质选取 = "利空"
End Function
