Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplEndReturn
    Implements ICondition


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

    End Sub


    Public Function GetResult(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal expFlag As String, ByVal transCondition As String) As Boolean Implements ICondition.GetResult

        '判断是否有还款证明书
        If GetReturnReceipt(ProjectID) = True Then
            Return True
        Else
            Return False
        End If

    End Function

    '获取还款证明书
    Private Function GetReturnReceipt(ByVal ProjectID As String) As Boolean

        '获取该项目的还款证明书记录
        Dim RefundCertificate As New RefundCertificate(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & "}"
        Dim dsTemp As DataSet = RefundCertificate.GetRefundCertificateInfo(strSql)

        '判断是否有还款证明书记录
        If dsTemp.Tables(0).Rows.Count <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
