
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfMessagesTemplate
    Public Const Table_Messages_Dict As String = "messages_dict"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfMessagesTemplate As SqlDataAdapter

    '定义查询命令
    Private GetWfMessagesTemplateInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfMessagesTemplate = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfMessagesTemplateInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfMessagesTemplateInfo(ByVal strSQL_Condition_WfMessagesTemplate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfMessagesTemplateInfoCommand Is Nothing Then

            GetWfMessagesTemplateInfoCommand = New SqlCommand("GetWfMessagesTemplateInfo", conn)
            GetWfMessagesTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfMessagesTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfMessagesTemplate
            .SelectCommand = GetWfMessagesTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfMessagesTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfMessagesTemplate
            .Fill(tempDs, Table_Messages_Dict)
        End With

        Return tempDs

    End Function

End Class
