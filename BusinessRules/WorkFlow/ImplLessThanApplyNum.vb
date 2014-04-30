Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplLessThanApplyNum
    Implements ICondition

    '定义申请次数
    Private applyNum As Integer

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

        '不同意且申请次数<申请次数限制
        If expFlag = "不同意" Then

            '获取申请次数
            Dim i As Integer = GetApplyNum(ProjectID)

            '判断申请次数是否<申请次数限制
            If i < _ApplyNumLimit Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function

    '获取申请次数
    Private Function GetApplyNum(ByVal ProjectID As String) As Integer
        Dim IntentLetter As New IntentLetter(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & "}"
        Dim dsTemp As DataSet = IntentLetter.GetIntentLetterInfo(strSql)

        '记录的条数即申请的次数
        applyNum = dsTemp.Tables(0).Rows.Count

    End Function
End Class
