Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplTrialFee
    Implements ICondition

    '定义评审费支出金额、收入总额
    Private TrialFeePayout, TotalTrialFeeIncome, GuaranteeFee As Single

    '定义项目意见对象引用
    Private ProjectOpinion As ProjectOpinion

    '定义项目对象引用
    Private Project As Project

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化项目意见对象
        ProjectOpinion = New ProjectOpinion(conn, ts)

        '实例化项目对象
        Project = New Project(conn, ts)


    End Sub


    Public Function GetResult(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal expFlag As String, ByVal transCondition As String) As Boolean Implements ICondition.GetResult


        '计算评审费的收入总额
        GetTotalTrialFeeIncome(projectID)

        ''获取担保金额
        'GetGuaranteeFee(projectID)

        '判断评审费的收入总额是否小于需要收取的评审费金额并且初审结论是否为“通过免收”
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and item_type='51' and item_code='005'}"
        Dim dsTempConclusion As DataSet = ProjectOpinion.GetProjectOpinion(strSql)

        '异常处理  
        If dsTempConclusion.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempConclusion.Tables(0))
            Throw wfErr
        End If

        Dim tmpConculsion As String = Trim(IIf(IsDBNull(dsTempConclusion.Tables(0).Rows(0).Item("conclusion")), "", dsTempConclusion.Tables(0).Rows(0).Item("conclusion")))
        If TotalTrialFeeIncome < TrialFeePayout And tmpConculsion <> "通过免收" Then
            Return True
        Else
            '将补收评审费任务的状态置为“F”
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CashlossReview'}"
            Dim WfProjectTaskAttendee As New WfProjectTaskAttendee(conn, ts)
            Dim dsTempTaskStatus As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            Dim i As Integer
            For i = 0 To dsTempTaskStatus.Tables(0).Rows.Count - 1
                dsTempTaskStatus.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskStatus)

            If transCondition = ".T." Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Private Function GetTotalTrialFeeIncome(ByVal ProjectID As String)

        '获取该项目关于评审费的记录
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and item_type='31'" & " and item_code='001'" & "}"
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer

        '计算评审费的收入总额
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            TrialFeePayout = TrialFeePayout + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("payout")), 0, dsTemp.Tables(0).Rows(i).Item("payout"))
            TotalTrialFeeIncome = TotalTrialFeeIncome + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income"))
        Next
    End Function

    ''获取担保费
    'Private Function GetGuaranteeFee(ByVal projectID As String)
    '    Dim strSql As String = "{project_code=" & "'" & projectID & "'" & "}"
    '    Dim dsTempGuaranteeFee As DataSet = Project.GetProjectInfo(strSql)
    '    GuaranteeFee = IIf(IsDBNull(dsTempGuaranteeFee.Tables(0).Rows(0).Item("guarantee_sum")), 0, dsTempGuaranteeFee.Tables(0).Rows(0).Item("guarantee_sum"))
    'End Function

End Class
