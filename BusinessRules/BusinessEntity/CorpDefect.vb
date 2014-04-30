Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class CorpDefect

    Public Const Table_Corporation_Defect_Record As String = "corporation_defect_record"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_CorpDefect As SqlDataAdapter

    '定义查询命令
    Private GetCorpDefectInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_CorpDefect = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetCorpDefectInfo("null")
    End Sub

    '获取企业污点信息
    Public Function GetCorpDefectInfo(ByVal strSQL_Condition_CorpDefect As String) As DataSet

        Dim tempDs As New DataSet()

        If GetCorpDefectInfoCommand Is Nothing Then

            GetCorpDefectInfoCommand = New SqlCommand("GetCorpDefectInfo", conn)
            GetCorpDefectInfoCommand.CommandType = CommandType.StoredProcedure
            GetCorpDefectInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CorpDefect
            .SelectCommand = GetCorpDefectInfoCommand
            .SelectCommand.Transaction = ts
            GetCorpDefectInfoCommand.Parameters("@Condition").Value = strSQL_Condition_CorpDefect
            .Fill(tempDs, Table_Corporation_Defect_Record)
        End With

        GetCorpDefectInfo = tempDs

    End Function

    '更新企业污点信息
    Public Function UpdateCorpDefect(ByVal CorpDefectSet As DataSet)

        If CorpDefectSet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If CorpDefectSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_CorpDefect)

        With dsCommand_CorpDefect
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(CorpDefectSet, Table_Corporation_Defect_Record)

        End With

        CorpDefectSet.AcceptChanges()


    End Function

End Class
