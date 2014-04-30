Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'�Ƿ���������ȷ����������
'��������ȷ��ValidateInsurance(44,008),�����ı�ȷ��ValidateLawText(44,014),
'Ԥ�䵱����ȷ��ValidatePrehock(44,009),��������������ȷ��IsValidateFraternity(44,011),�������ȷ��ValidateWard(44,013)
'��Ѻ����ȷ��IsValidateInGage(44,006)����Ѻ����ȷ��IsValidateImpawned(44,007)


Public Class ImplIsStartAffirmTaskExp
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '����ת�������������
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String

        Dim i, j, len As Integer
        Dim dsTempTaskTrans, dsAttend As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)

        Dim strItemType = "44"
        Dim strItemCode As String() = {"006", "007", "008", "009", "011", "013", "014"}
        Dim strTaskID As String() = {"IsValidateInGageExp", "IsValidateImpawnedExp", "IsValidateInsuranceExp", "ValidatePrehockExp", "IsValidateFraternityExp", "ValidateWardExp", "ValidateLawTextExp"}
        Dim strNextTaskID As String() = {"ValidateInGageExp", "ValidateImpawnedExp", "ValidateInsuranceExp", "ValidatePrehockExp", "ValidateFraternityExp", "ValidateWardExp", "ValidateLawTextExp"}

        len = strItemCode.Length
        If len > 0 Then
            For i = 0 To len - 1
                If isHaveRecord(projectID, strItemType, strItemCode(i)) Then
                    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckedSignatureExp' and next_task='" & strTaskID(i) & "'}" '���ǩԼ
                    dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                    If Not dsTempTaskTrans Is Nothing Then
                        If dsTempTaskTrans.Tables(0).Rows.Count > 0 Then
                            If dsTempTaskTrans.Tables(0).Rows(0).Item("next_task") = strTaskID(i) Then
                                dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
                            End If
                        End If
                    End If
                Else
                    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckedSignatureExp' and next_task='" & strTaskID(i) & "'}" '���ǩԼ
                    dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                    If Not dsTempTaskTrans Is Nothing Then
                        If dsTempTaskTrans.Tables(0).Rows.Count > 0 Then
                            If dsTempTaskTrans.Tables(0).Rows(0).Item("next_task") = strTaskID(i) Then
                                dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
                            End If
                        End If
                    End If

                    '������ȷ������,��Ѹ�����(�������񣬼���������¸�����)��Ϊ.F.
                    ' {"ValidateInGage", "ValidateImpawned", "ValidateInsurance", "ValidatePrehock", "ValidateFraternity", "ValidateWard", "ValidateLawText"}
                    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='" & strNextTaskID(i) & "'}"
                    dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                    If dsAttend.Tables(0).Rows.Count > 0 Then
                        For j = 0 To dsAttend.Tables(0).Rows.Count - 1
                            dsAttend.Tables(0).Rows(j).Item("task_status") = "F"
                            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)
                        Next
                    End If
                End If
                WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)
            Next

        End If

    End Function

    '�ж��Ƿ�����Ҫȷ�ϵļ�¼
    Private Function isHaveRecord(ByVal projectID As String, ByVal itemType As String, ByVal itemCode As String) As Boolean
        Dim projectGuaranteeFormAdditional As ProjectGuaranteeFormAdditional = New ProjectGuaranteeFormAdditional(conn, ts)
        Dim ds As DataSet
        Dim countA, countB As Integer

        ds = projectGuaranteeFormAdditional.GetProjectGuaranteeFormAdditional(projectID, itemType, itemCode)
        countA = ds.Tables(0).Rows.Count
        countB = ds.Tables(1).Rows.Count
        If countA > 0 Or countB > 0 Then
            Return True
        Else
            Return False
        End If

    End Function
End Class
