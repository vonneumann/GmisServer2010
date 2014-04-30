Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplApplyCount
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
        '��	�ڵ��������飨itent_letter����ȡ���������
        Dim applyCount As Integer = GetApplyCount(projectID)

        '��	����������С��3����ת������FROMID=65��TOID=24 ��ת��������Ϊ��.T.������FROMID=65��TOID=62 ��ת��������Ϊ��.F.��;
        '���򣬽�FROMID=65��TOID=24 ��ת��������Ϊ��.F.������FROMID=65��TOID=62 ��ת��������
        'Ϊ��.T.��;
        Dim strSql As String
        Dim i As Integer
        Dim dsTempTaskTrans As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckApplyTimes'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)
        If applyCount < _ApplyNumLimit Then
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ApplyLetterIntent" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        Else
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ApplyLetterIntent" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)


        '������Ŀ��Ϣ


    End Function

    '��ȡ�������
    Private Function GetApplyCount(ByVal ProjectID As String) As Integer
        Dim IntentLetter As New IntentLetter(conn, ts)
        'qxd modify 2004-9-24
        'Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & "}"
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and bank_reply='��ͬ��'}"
        Dim dsTemp As DataSet = IntentLetter.GetIntentLetterInfo(strSql)

        '��¼������������Ĵ���
        Return dsTemp.Tables(0).Rows.Count

    End Function
End Class
