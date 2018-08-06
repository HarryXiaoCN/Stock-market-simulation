Attribute VB_Name = "SMS_Type"
Public Type 玩家
    身份 As String
    资金 As Currency
    信息优先级 As Long
End Type
Public Type 公司
    股价 As Currency
    总市值 As Currency
    流通市值 As Currency
    总股数 As Currency
    流通股数 As Currency
End Type
Public Type 情报
    存在 As Boolean
    好坏度 As String
End Type
Public Type 股价
    时间 As Long
    价位 As Currency
End Type
