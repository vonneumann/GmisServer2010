Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectContractEstateElement

    Public Const Table_ProjectContractEstateElement As String = "project_contract_estate_element"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ProjectContractEstateElement As SqlDataAdapter

    '定义查询命令
    Private GetProjectContractEstateElementInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ProjectContractEstateElement = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectContractEstateElementInfo("null")
    End Sub

    '获取项目基本信息
    Public Function GetProjectContractEstateElementInfo(ByVal strSQL_Condition_ProjectContractEstateElement As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectContractEstateElementInfoCommand Is Nothing Then

            GetProjectContractEstateElementInfoCommand = New SqlCommand("GetProjectContractEstateElementInfo", conn)
            GetProjectContractEstateElementInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectContractEstateElementInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectContractEstateElement
            .SelectCommand = GetProjectContractEstateElementInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectContractEstateElementInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectContractEstateElement
            .Fill(tempDs, Table_ProjectContractEstateElement)
        End With

        Return tempDs

    End Function

    '更新项目基本信息
    Public Function UpdateProjectContractEstateElement(ByVal ProjectContractEstateElementSet As DataSet)

        If ProjectContractEstateElementSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ProjectContractEstateElementSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectContractEstateElement)

        With dsCommand_ProjectContractEstateElement
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectContractEstateElementSet, Table_ProjectContractEstateElement)

        End With

        ProjectContractEstateElementSet.AcceptChanges()

    End Function
End Class
