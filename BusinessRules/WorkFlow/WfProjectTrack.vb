Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectTrack

    Public Const Table_Project_Track As String = "project_track"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfProjectTrack As SqlDataAdapter

    '定义查询命令
    Private GetWfProjectTrackInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfProjectTrack = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfProjectTrackInfo("null")
    End Sub

    '获取工作流记录检查信息
    Public Function GetWfProjectTrackInfo(ByVal strSQL_Condition_WfProjectTrack As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfProjectTrackInfoCommand Is Nothing Then

            GetWfProjectTrackInfoCommand = New SqlCommand("GetWfProjectTrackInfo", conn)
            GetWfProjectTrackInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfProjectTrackInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectTrack
            .SelectCommand = GetWfProjectTrackInfoCommand
            .SelectCommand.Transaction = ts
            GetWfProjectTrackInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectTrack
            .Fill(tempDs, Table_Project_Track)
        End With

        Return tempDs
      
    End Function

    '更新工作流记录检查信息
    Public Function UpdateWfProjectTrack(ByVal WfProjectTrackSet As DataSet)

        If WfProjectTrackSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If WfProjectTrackSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectTrack)

        With dsCommand_WfProjectTrack
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectTrackSet, Table_Project_Track)

            WfProjectTrackSet.AcceptChanges()
        End With
       
    End Function

End Class
