Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class IntentLetter

    Public Const Table_IntentLetter As String = "intent_letter"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_IntentLetter As SqlDataAdapter

    '定义查询命令
    Private GetIntentLetterInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_IntentLetter = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetIntentLetterInfo("null")
    End Sub

    '获取意向书信息
    Public Function GetIntentLetterInfo(ByVal strSQL_Condition_IntentLetter As String) As DataSet

        Dim tempDs As New DataSet()

        If GetIntentLetterInfoCommand Is Nothing Then

            GetIntentLetterInfoCommand = New SqlCommand("GetIntentLetterInfo", conn)
            GetIntentLetterInfoCommand.CommandType = CommandType.StoredProcedure
            GetIntentLetterInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_IntentLetter
            .SelectCommand = GetIntentLetterInfoCommand
            .SelectCommand.Transaction = ts
            GetIntentLetterInfoCommand.Parameters("@Condition").Value = strSQL_Condition_IntentLetter
            .Fill(tempDs, Table_IntentLetter)
        End With

        Return tempDs

    End Function

    '更新意向书信息
    Public Function UpdateIntentLetter(ByVal IntentLetterSet As DataSet)

        If IntentLetterSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If IntentLetterSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_IntentLetter)

        With dsCommand_IntentLetter
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(IntentLetterSet, Table_IntentLetter)

        End With

        IntentLetterSet.AcceptChanges()

    End Function

End Class
