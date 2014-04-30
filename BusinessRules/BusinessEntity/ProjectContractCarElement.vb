Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectContractCarElement

    Public Const Table_ProjectContractCarElement As String = "project_contract_car_element"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ProjectContractCarElement As SqlDataAdapter

    '定义查询命令
    Private GetProjectContractCarElementInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ProjectContractCarElement = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectContractCarElementInfo("null")
    End Sub

    '获取项目基本信息
    Public Function GetProjectContractCarElementInfo(ByVal strSQL_Condition_ProjectContractCarElement As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectContractCarElementInfoCommand Is Nothing Then

            GetProjectContractCarElementInfoCommand = New SqlCommand("GetProjectContractCarElementInfo", conn)
            GetProjectContractCarElementInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectContractCarElementInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectContractCarElement
            .SelectCommand = GetProjectContractCarElementInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectContractCarElementInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectContractCarElement
            .Fill(tempDs, Table_ProjectContractCarElement)
        End With

        Return tempDs

    End Function

    '更新项目基本信息
    Public Function UpdateProjectContractCarElement(ByVal ProjectContractCarElementSet As DataSet)

        If ProjectContractCarElementSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ProjectContractCarElementSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectContractCarElement)

        With dsCommand_ProjectContractCarElement
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectContractCarElementSet, Table_ProjectContractCarElement)

        End With

        ProjectContractCarElementSet.AcceptChanges()

    End Function
End Class
