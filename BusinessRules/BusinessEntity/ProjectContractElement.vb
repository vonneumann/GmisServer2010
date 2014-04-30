Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectContractElement

    Public Const Table_ProjectContractElement As String = "project_contract_element"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ProjectContractElement As SqlDataAdapter

    '定义查询命令
    Private GetProjectContractElementInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ProjectContractElement = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectContractElementInfo("null")
    End Sub

    '获取项目基本信息
    Public Function GetProjectContractElementInfo(ByVal strSQL_Condition_ProjectContractElement As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectContractElementInfoCommand Is Nothing Then

            GetProjectContractElementInfoCommand = New SqlCommand("GetProjectContractElementInfo", conn)
            GetProjectContractElementInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectContractElementInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectContractElement
            .SelectCommand = GetProjectContractElementInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectContractElementInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectContractElement
            .Fill(tempDs, Table_ProjectContractElement)
        End With

        Return tempDs

    End Function

    '更新项目基本信息
    Public Function UpdateProjectContractElement(ByVal ProjectContractElementSet As DataSet)

        If ProjectContractElementSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ProjectContractElementSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectContractElement)

        With dsCommand_ProjectContractElement
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectContractElementSet, Table_ProjectContractElement)

        End With

        ProjectContractElementSet.AcceptChanges()

    End Function
End Class
