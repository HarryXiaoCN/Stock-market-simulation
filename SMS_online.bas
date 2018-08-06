Attribute VB_Name = "SMS_online"
Public Local_State, TCPNum As Long
Public Function 主持人消息接收(ByVal 消息 As String)

End Function
Public Function 主持人行情消息发送(ByVal 谁 As String, ByVal 内容 As String)
Dim i As Long: Dim 消息 As String
消息 = "主庄#" & 内容 & "|"
For i = 1 To 14
    If 股民(i).身份 = 谁 And 服务器.Winsock1(i).State = 7 Then 服务器.Winsock1(i).SendData 消息
Next
End Function
Public Function 游戏初始化_主持人状态消息发送()
Dim 消息 As String
On Error GoTo Er:
For i = 1 To 14
    消息 = "初始化#" & 当前回合 & "#" & 当前阶段 & "#" & 操作详情 & "#" & 股民(i).身份 & "#" & 股民(i).信息优先级 & "#" & 股民(i).资金 & "#" & 操作身份 & "|"
    If 服务器.Winsock1(i).State = 7 Then 服务器.Winsock1(i).SendData 消息
Next
Er:
MsgBox "游戏初始化_主持人状态消息发送>错误"
End Function
Public Function 主持人状态消息发送()
Dim 消息 As String
消息 = "主态#" & 当前回合 & "#" & 当前阶段 & "#" & 操作详情 & "|"
For i = 1 To 14
    If 服务器.Winsock1(i).State = 7 Then 服务器.Winsock1(i).SendData 消息
Next
End Function
Public Function 玩家消息接收(ByVal 消息 As String)
Dim 消息缓存, 内容缓存, 缓存: Dim i As Long
消息缓存 = Split(消息, "|")
For i = 0 To UBound(消息缓存) - 1
    内容缓存 = Split(消息缓存(i), "#")
    Select Case 消息缓存(0)
        Case "主庄"
            缓存 = Split(消息缓存(1), "-")
            个股消息池(1, Val(缓存(0))) = 缓存(1)
        Case "初始化"
            当前回合 = Val(消息缓存(1)): 当前阶段 = 消息缓存(2): 操作详情 = 消息缓存(3)
            我.身份 = 消息缓存(4): 我.信息优先级 = Val(消息缓存(5)): 我.资金 = Val(消息缓存(6))
            操作身份 = 信息缓存(7)
            If 操作身份 = 我.身份 Then 操作权限 = True
        Case "主态"
            当前回合 = Val(消息缓存(1)): 当前阶段 = 消息缓存(2): 操作详情 = 消息缓存(3)
    End Select
Next
End Function
Public Function Client_SendData(ByRef StrData As String)
客户端.Winsock1.SendData StrData
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
