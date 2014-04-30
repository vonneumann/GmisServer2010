Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Appraisement

    Public Const Table_Project_Appraisement As String = "project_appraisement"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_Appraisement As SqlDataAdapter

    '定义查询命令
    Private GetAppraisementInfoCommand As SqlCommand
    Private GetMaxAppraisementNumCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_Appraisement = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetAppraisementInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetAppraisementInfo(ByVal strSQL_Condition_Appraisement As String) As DataSet

        Dim tempDs As New DataSet()

        If GetAppraisementInfoCommand Is Nothing Then

            GetAppraisementInfoCommand = New SqlCommand("GetAppraisementInfo", conn)
            GetAppraisementInfoCommand.CommandType = CommandType.StoredProcedure
            GetAppraisementInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Appraisement
            .SelectCommand = GetAppraisementInfoCommand
            .SelectCommand.Transaction = ts
            GetAppraisementInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Appraisement
            .Fill(tempDs, Table_Project_Appraisement)
        End With

        Return tempDs

    End Function

    '获取最大序列号
    Public Function GetMaxAppraisementNum(ByVal projectID As String) As Integer

        If GetMaxAppraisementNumCommand Is Nothing Then

            GetMaxAppraisementNumCommand = New SqlCommand("GetMaxAppraisementNum", conn)
            GetMaxAppraisementNumCommand.CommandType = CommandType.StoredProcedure
            GetMaxAppraisementNumCommand.Parameters.Add(New SqlParameter("@projectID", SqlDbType.NVarChar))
            GetMaxAppraisementNumCommand.Parameters.Add(New SqlParameter("@maxAppraisementNum", SqlDbType.Int))
            GetMaxAppraisementNumCommand.Parameters.Item("@maxAppraisementNum").Direction = ParameterDirection.Output
            GetMaxAppraisementNumCommand.Transaction = ts
        End If

        GetMaxAppraisementNumCommand.Parameters("@projectID").Value = projectID
        GetMaxAppraisementNumCommand.ExecuteNonQuery()
        GetMaxAppraisementNum = GetMaxAppraisementNumCommand.Parameters.Item("@maxAppraisementNum").Value
    End Function

    '更新项目评价信息
    Public Function UpdateAppraisement(ByVal AppraisementSet As DataSet)

        If AppraisementSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If AppraisementSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Appraisement)

        With dsCommand_Appraisement
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(AppraisementSet, Table_Project_Appraisement)

            AppraisementSet.AcceptChanges()
        End With


    End Function


End Class

