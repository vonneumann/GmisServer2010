Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'判断反担保物是否需要分配评估师
Public Class ImplIsAssignValuator
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义转移任务对象引用
    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    '定义反担保物对象引用
    Private Guaranty As Guaranty

    '定义类目字典对象引用
    Private Item As Item

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化转移任务对象
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

        '实例化反担保物对象
        Guaranty = New Guaranty(conn, ts)

        '实例化类目数据字典对象
        Item = New Item(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim i, j As Integer
        Dim strSql As String
        Dim dsGuaranty, dsItem, dsTempTaskTrans As DataSet
        Dim tmpItemCode As String
        Dim isAssignValuator As Boolean

        '判断该项目的反担保物是否都不需要分配评估师
        strSql = "{project_code=" & "'" & projectID & "'" & " and evaluate_person is null}"
        dsGuaranty = Guaranty.GetGuarantyInfo(strSql, "null")

        For i = 0 To dsGuaranty.Tables(0).Rows.Count - 1
            tmpItemCode = dsGuaranty.Tables(0).Rows(i).Item("guaranty_type")
            dsItem = Item.GetItemType(tmpItemCode)

            '异常处理  
            If dsItem.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsItem.Tables(0))
                Throw wfErr
            End If

            If IIf(IsDBNull(dsItem.Tables(0).Rows(0).Item("additional_remark")), 0, dsItem.Tables(0).Rows(0).Item("additional_remark")) = 1 Then

                isAssignValuator = True

                Exit For

            End If
        Next


        If isAssignValuator = False Then


            '将判定是否资产评估虚任务（TID=CHKneedEvaluated）到分配资产评估师的转移条件置为假.F.
            '将判定是否资产评估虚任务（TID=CHKneedEvaluated）到登记反担保物的转移条件置为假.T.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKneedEvaluated' and next_task='AssignValuator'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKneedEvaluated' and next_task='ApplyCapitialEvaluated'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

        Else
            '否则

            '将判定是否资产评估虚任务（TID=CHKneedEvaluated）到分配资产评估师的转移条件置为假.T.
            '将判定是否资产评估虚任务（TID=CHKneedEvaluated）到登记反担保物的转移条件置为假.F.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKneedEvaluated' and next_task='AssignValuator'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKneedEvaluated' and next_task='ApplyCapitialEvaluated'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)
        End If

        ''判断是否有反担保措施为：保证金(item_type:2A)，有：启动确认保证金收费标准任务；

        'strSql = "{project_code=" & "'" & projectID & "'" & " and guaranty_type='2A'}"
        'dsGuaranty = Guaranty.GetGuarantyInfo(strSql, "null")

        'If dsGuaranty.Tables(0).Rows.Count > 0 Then
        '    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKneedEvaluated' and next_task='ValidateDepositFee'}"
        '    dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        '    If dsTempTaskTrans.Tables(0).Rows.Count > 0 Then
        '        dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
        '        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)
        '    End If
        'End If

    End Function
End Class
