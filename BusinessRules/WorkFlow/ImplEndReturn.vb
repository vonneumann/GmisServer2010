Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplEndReturn
    Implements ICondition


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


    Public Function GetResult(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal expFlag As String, ByVal transCondition As String) As Boolean Implements ICondition.GetResult

        '�ж��Ƿ��л���֤����
        If GetReturnReceipt(ProjectID) = True Then
            Return True
        Else
            Return False
        End If

    End Function

    '��ȡ����֤����
    Private Function GetReturnReceipt(ByVal ProjectID As String) As Boolean

        '��ȡ����Ŀ�Ļ���֤�����¼
        Dim RefundCertificate As New RefundCertificate(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & "}"
        Dim dsTemp As DataSet = RefundCertificate.GetRefundCertificateInfo(strSql)

        '�ж��Ƿ��л���֤�����¼
        If dsTemp.Tables(0).Rows.Count <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
