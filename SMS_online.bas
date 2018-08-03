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
Public Function 连接状态反馈(ByVal SID As Long) As String
Select Case SID
    Case 0
        连接状态反馈 = "未连接"
    Case 1
        连接状态反馈 = "打开状态"
    Case 2
        连接状态反馈 = "等待连接"
    Case 3
        连接状态反馈 = "连接挂起"
    Case 4
        连接状态反馈 = "域名解析中..."
    Case 5
        连接状态反馈 = "域名解析成功"
    Case 6
        连接状态反馈 = "正在连接..."
    Case 7
        连接状态反馈 = "已连接"
    Case 9
        连接状态反馈 = "连接错误"
End Select
End Function
