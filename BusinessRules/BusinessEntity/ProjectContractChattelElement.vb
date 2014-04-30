Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectContractChattelElement

    Public Const Table_ProjectContractChattelElement As String = "project_contract_chattel_element"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ProjectContractChattelElement As SqlDataAdapter

    '定义查询命令
    Private GetProjectContractChattelElementInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ProjectContractChattelElement = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectContractChattelElementInfo("null")
    End Sub

    '获取项目基本信息
    Public Function GetProjectContractChattelElementInfo(ByVal strSQL_Condition_ProjectContractChattelElement As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectContractChattelElementInfoCommand Is Nothing Then

            GetProjectContractChattelElementInfoCommand = New SqlCommand("GetProjectContractChattelElementInfo", conn)
            GetProjectContractChattelElementInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectContractChattelElementInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectContractChattelElement
            .SelectCommand = GetProjectContractChattelElementInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectContractChattelElementInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectContractChattelElement
            .Fill(tempDs, Table_ProjectContractChattelElement)
        End With

        Return tempDs

    End Function

    '更新项目基本信息
    Public Function UpdateProjectContractChattelElement(ByVal ProjectContractChattelElementSet As DataSet)

        If ProjectContractChattelElementSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ProjectContractChattelElementSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectContractChattelElement)

        With dsCommand_ProjectContractChattelElement
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectContractChattelElementSet, Table_ProjectContractChattelElement)

        End With

        ProjectContractChattelElementSet.AcceptChanges()

    End Function
End Class
