Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplTrialFee
    Implements ICondition

    '���������֧���������ܶ�
    Private TrialFeePayout, TotalTrialFeeIncome, GuaranteeFee As Single

    '������Ŀ�����������
    Private ProjectOpinion As ProjectOpinion

    '������Ŀ��������
    Private Project As Project

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

        'ʵ������Ŀ�������
        ProjectOpinion = New ProjectOpinion(conn, ts)

        'ʵ������Ŀ����
        Project = New Project(conn, ts)


    End Sub


    Public Function GetResult(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal expFlag As String, ByVal transCondition As String) As Boolean Implements ICondition.GetResult


        '��������ѵ������ܶ�
        GetTotalTrialFeeIncome(projectID)

        ''��ȡ�������
        'GetGuaranteeFee(projectID)

        '�ж�����ѵ������ܶ��Ƿ�С����Ҫ��ȡ������ѽ��ҳ�������Ƿ�Ϊ��ͨ�����ա�
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and item_type='51' and item_code='005'}"
        Dim dsTempConclusion As DataSet = ProjectOpinion.GetProjectOpinion(strSql)

        '�쳣����  
        If dsTempConclusion.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempConclusion.Tables(0))
            Throw wfErr
        End If

        Dim tmpConculsion As String = Trim(IIf(IsDBNull(dsTempConclusion.Tables(0).Rows(0).Item("conclusion")), "", dsTempConclusion.Tables(0).Rows(0).Item("conclusion")))
        If TotalTrialFeeIncome < TrialFeePayout And tmpConculsion <> "ͨ������" Then
            Return True
        Else
            '����������������״̬��Ϊ��F��
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

        '��ȡ����Ŀ��������ѵļ�¼
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and item_type='31'" & " and item_code='001'" & "}"
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer

        '��������ѵ������ܶ�
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            TrialFeePayout = TrialFeePayout + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("payout")), 0, dsTemp.Tables(0).Rows(i).Item("payout"))
            TotalTrialFeeIncome = TotalTrialFeeIncome + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income"))
        Next
    End Function

    ''��ȡ������
    'Private Function GetGuaranteeFee(ByVal projectID As String)
    '    Dim strSql As String = "{project_code=" & "'" & projectID & "'" & "}"
    '    Dim dsTempGuaranteeFee As DataSet = Project.GetProjectInfo(strSql)
    '    GuaranteeFee = IIf(IsDBNull(dsTempGuaranteeFee.Tables(0).Rows(0).Item("guarantee_sum")), 0, dsTempGuaranteeFee.Tables(0).Rows(0).Item("guarantee_sum"))
    'End Function

End Class
