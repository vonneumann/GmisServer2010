Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectCounterClaim
    Private Const Table_Project_Counter As String = "Project_counterclaim"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ProjectCounter As SqlDataAdapter

    '定义查询命令
    Private GetProjectCounterInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ProjectCounter = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectCounterClaimInfo("null")
    End Sub

    '获取索赔信息
    Public Function GetProjectCounterClaimInfo(ByVal strSQL_Condition_ProjectCounterClaim As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectCounterInfoCommand Is Nothing Then

            GetProjectCounterInfoCommand = New SqlCommand("GetProjectCounterClaim", conn)
            GetProjectCounterInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectCounterInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectCounter
            .SelectCommand = GetProjectCounterInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectCounterInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectCounterClaim
            .Fill(tempDs, Table_Project_Counter)
        End With

        Return tempDs

    End Function

    '更新索赔信息
    Public Function UpdateProjectCounterClaim(ByVal ProjectCounterClaimSet As DataSet)


        If ProjectCounterClaimSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ProjectCounterClaimSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectCounter)

        With dsCommand_ProjectCounter
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectCounterClaimSet, Table_Project_Counter)

            ProjectCounterClaimSet.AcceptChanges()
        End With


    End Function

End Class
