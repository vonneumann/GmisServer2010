'定义条件判断接口
Public Interface ICondition
    '处理条件表达式
    Function GetResult(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal expFlag As String, ByVal transCondition As String) As Boolean
End Interface

