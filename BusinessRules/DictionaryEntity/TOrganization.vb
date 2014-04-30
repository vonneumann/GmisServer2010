Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class TOrganization
    Public Const Table_TOrganization As String = "TOrganization"


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_TOrganization As SqlDataAdapter


    '定义查询命令
    Private GetTOrganizationInfoCommand As SqlCommand


    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_TOrganization = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetTOrganizationInfo("null")
    End Sub


    Public Function GetTOrganizationInfo(ByVal strSQL_Condition_TOrganization) As DataSet

        Dim tempDs As New DataSet()

        If GetTOrganizationInfoCommand Is Nothing Then

            GetTOrganizationInfoCommand = New SqlCommand("GetTOrganizationInfo", conn)
            GetTOrganizationInfoCommand.CommandType = CommandType.StoredProcedure
            GetTOrganizationInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_TOrganization
            .SelectCommand = GetTOrganizationInfoCommand
            .SelectCommand.Transaction = ts
            GetTOrganizationInfoCommand.Parameters("@Condition").Value = strSQL_Condition_TOrganization
            .Fill(tempDs, Table_TOrganization)
        End With

        Return tempDs

    End Function


    Public Function UpdateTOrganization(ByVal TOrganization As DataSet)

        If TOrganization Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If TOrganization.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_TOrganization)

        With dsCommand_TOrganization
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(TOrganization, Table_TOrganization)

            TOrganization.AcceptChanges()
        End With


    End Function

End Class
