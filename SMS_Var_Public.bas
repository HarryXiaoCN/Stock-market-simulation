Attribute VB_Name = "SMS_Var_Public"
Public 股民(15) As 玩家
Public 我 As 玩家
Public GitHub As 公司
Public 个股消息池(4, 1000) As 情报
Public 大盘消息池(4, 1000) As 情报
Public 操作权限 As Boolean
Public 操作详情, 当前阶段, 当前回合, 操作身份 As String
Public Function 游戏初始化()
当前回合 = 1
当前阶段 = "准备阶段"
操作身份 = "庄家"
游戏初始化_主持人状态消息发送
玩家初始化
公司初始化
End Function
Public Function 玩家初始化()
Dim 身份分配锁(15) As Boolean
Dim 玩家号缓存, i As Long
Randomize
玩家号缓存 = Int(Rnd * (14)) + 1
身份分配锁(玩家号缓存) = True
庄家初始化 (玩家号缓存)
For i = 0 To 1
    玩家号缓存 = Int(Rnd * (14)) + 1
    Do Until 身份分配锁(玩家号缓存) = False
        玩家号缓存 = Int(Rnd * (14)) + 1
    Loop
    游资初始化 (玩家号缓存)
Next
For i = 1 To 14
    If 身份分配锁(i) = False Then
        散户初始化 (i)
    End If
Next
End Function
Public Function 庄家初始化(ByVal 玩家号 As Long)
With 股民(玩家号)
    .身份 = "庄家"
    .资金 = 10000
    .信息优先级 = 1
End With
End Function
Public Function 游资初始化(ByVal 玩家号 As Long)
With 股民(玩家号)
    .身份 = "游资"
    .资金 = 1000
    .信息优先级 = 2
End With
End Function
Public Function 散户初始化(ByVal 玩家号 As Long)
With 股民(玩家号)
    .身份 = "散户"
    .资金 = 10
    .信息优先级 = 3
End With
End Function
Public Function 公司初始化()
With GitHub
    .流通股数 = 20000
    .总股数 = 30000
End With
End Function

