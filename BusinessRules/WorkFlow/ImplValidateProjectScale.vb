Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplValidateProjectScale
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    Private project As Project

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

        project = New Project(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String

        '获取项目是否再次申请
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsTemp As DataSet = project.GetProjectInfo(strSql)

        '异常处理  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        Dim applySum As Single = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("apply_sum")), 0, dsTemp.Tables(0).Rows(0).Item("apply_sum"))

        'ValidateProjectScale接口实现当申请金额大于等于1000万元时，将ValidateProjectScale-ProjectPause置为假，ValidateProjectScale-ValidateProjectPause置为真。
        If applySum >= 1000 Then
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectScale' and next_task='ProjectPause'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectScale' and next_task='ValidateProjectPause'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)


            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        Else

            '否则,将ValidateProjectScale-ProjectPause置为真，ValidateProjectScale-ValidateProjectPause置为假。

            '2004-7-6 应担保中心秦勇的要求更改如下代码 (by qxd)
            '接口实现的是如果项目申请金额小于1000万元(, 将ValidateProjectScale - ProjectPause置为真, ValidateProjectScale - ValidateProjectPause置为假
            '现改为ValidateProjectScale-SubmitCancelProArchives置为真, ValidateProjectScale - ValidateProjectPause置为假)

            'qxd delete 2005-5-10
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectScale' and next_task='ProjectPause'}"
            'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectScale' and next_task='SubmitCancelProArchives_normal'}"

            'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectScale' and next_task='SubmitCancelProArchives'}"
            'qxd add start 2004-10-29
            '此为：项目在提交调研结论、登记银行回复、项目暂缓流程处理项目暂缓，请完善这几处的项目终止。现在的问题是由于项目流程中多次出现SubmitCancelProArchives任务，
            '导致审核项目暂缓提交错误.同时修改了中间层接口() : ImplValidateProjectScale。 qxd modify 2004-10-29
            'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectScale' and next_task='SubmitCancelProArchives_normal'}"
            'qxd add end
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectScale' and next_task='ValidateProjectPause'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        End If

    End Function

End Class
