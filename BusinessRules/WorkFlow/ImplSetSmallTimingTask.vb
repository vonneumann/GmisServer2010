
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'����С����Ϣ��ʱ��������
Public Class ImplSetSmallTimingTask
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans
    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        Dim i As Integer

        '��ɾ��ԭ�е���ʾ
        strSql = "{project_code='" & projectID & "' and task_id='ValidateServiceFeeEx'}"

        Dim WfProjectTimingTask As New WfProjectTimingTask(conn, ts)
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Delete()
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        strSql = "select terms,starttime,endtime from queryProjectInfo where ProjectCode='" & projectID & "'"
        Dim objCommonQuery As New CommonQuery(conn, ts)
        Dim dsProjectInfo As DataSet = objCommonQuery.GetCommonQueryInfo(strSql)
        Dim dStartTime As Date = dsProjectInfo.Tables(0).Rows(0).Item("starttime")
        Dim dEndTime As Date = dsProjectInfo.Tables(0).Rows(0).Item("endtime")

        Dim iTerms As Integer = dsProjectInfo.Tables(0).Rows(0).Item("terms")
        Dim dFormDate As Date = CDate(dStartTime.Year.ToString & "-" & dStartTime.Month.ToString() & "-" & 20)
        Dim refeeDate As String
        Dim newRow As DataRow

        If dStartTime.Day < 20 Then

            For i = 0 To iTerms - 1
                newRow = dsTempTimingTask.Tables(0).NewRow
                refeeDate = dFormDate.AddMonths(i).ToShortDateString()
                With newRow
                    .Item("workflow_id") = workFlowID
                    .Item("project_code") = projectID
                    .Item("task_id") = "ValidateServiceFeeEx"
                    .Item("role_id") = "43"
                    .Item("type") = "T"
                    .Item("start_time") = refeeDate
                    .Item("status") = "P"
                    .Item("time_limit") = 0
                    .Item("distance") = 0
                    .Item("message_id") = 30
                End With
                dsTempTimingTask.Tables(0).Rows.Add(newRow)
            Next

            newRow = dsTempTimingTask.Tables(0).NewRow
            refeeDate = dStartTime.AddMonths(iTerms).ToShortDateString()
            With newRow
                .Item("workflow_id") = workFlowID
                .Item("project_code") = projectID
                .Item("task_id") = "ValidateServiceFeeEx"
                .Item("role_id") = "43"
                .Item("type") = "T"
                .Item("start_time") = refeeDate
                .Item("status") = "P"
                .Item("time_limit") = 0
                .Item("distance") = 0
                .Item("message_id") = 30
            End With
            dsTempTimingTask.Tables(0).Rows.Add(newRow)

        End If

        If dStartTime.Day > 20 Then

            For i = 0 To iTerms - 1
                newRow = dsTempTimingTask.Tables(0).NewRow
                refeeDate = dFormDate.AddMonths(i + 1).ToShortDateString()
                With newRow
                    .Item("workflow_id") = workFlowID
                    .Item("project_code") = projectID
                    .Item("task_id") = "ValidateServiceFeeEx"
                    .Item("role_id") = "43"
                    .Item("type") = "T"
                    .Item("start_time") = refeeDate
                    .Item("status") = "P"
                    .Item("time_limit") = 0
                    .Item("distance") = 0
                    .Item("message_id") = 30
                End With
                dsTempTimingTask.Tables(0).Rows.Add(newRow)
            Next


            newRow = dsTempTimingTask.Tables(0).NewRow
            refeeDate = dStartTime.AddMonths(iTerms).ToShortDateString()
            With newRow
                .Item("workflow_id") = workFlowID
                .Item("project_code") = projectID
                .Item("task_id") = "ValidateServiceFeeEx"
                .Item("role_id") = "43"
                .Item("type") = "T"
                .Item("start_time") = refeeDate
                .Item("status") = "P"
                .Item("time_limit") = 0
                .Item("distance") = 0
                .Item("message_id") = 30
            End With
            dsTempTimingTask.Tables(0).Rows.Add(newRow)

        End If

        If dStartTime.Day = 20 Then

            For i = 0 To iTerms - 1
                newRow = dsTempTimingTask.Tables(0).NewRow
                refeeDate = dFormDate.AddMonths(i + 1).ToShortDateString()
                With newRow
                    .Item("workflow_id") = workFlowID
                    .Item("project_code") = projectID
                    .Item("task_id") = "ValidateServiceFeeEx"
                    .Item("role_id") = "43"
                    .Item("type") = "T"
                    .Item("start_time") = refeeDate
                    .Item("status") = "P"
                    .Item("time_limit") = 0
                    .Item("distance") = 0
                    .Item("message_id") = 30
                End With
                dsTempTimingTask.Tables(0).Rows.Add(newRow)
            Next

        End If

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)



    End Function

End Class
