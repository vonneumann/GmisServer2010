Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectGuaranteeFormAdditional

    Public Const Table_form As String = "TProjectGuaranteeForm"
    Public Const Table_form_additional As String = "TProjectGuaranteeFormAdditional"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_form As SqlDataAdapter
    Private dsCommand_form_additional As SqlDataAdapter

    '定义查询命令
    Private GetFormInfoCommand As SqlCommand
    Private GetForm_additionalInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_form = New SqlDataAdapter()
        dsCommand_form_additional = New SqlDataAdapter()


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectGuaranteeFormAdditional("null", "null", "null")
    End Sub

    '获取信息
    Public Function GetProjectGuaranteeFormAdditional(ByVal projectCode As String, ByVal itemType As String, ByVal itemCode As String) As DataSet

        Dim tempDs As New DataSet()

        If GetFormInfoCommand Is Nothing Then

            GetFormInfoCommand = New SqlCommand("GetProjectOppositeForm", conn)
            GetFormInfoCommand.CommandType = CommandType.StoredProcedure
            GetFormInfoCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            GetFormInfoCommand.Parameters.Add(New SqlParameter("@ItemType", SqlDbType.NVarChar))
            GetFormInfoCommand.Parameters.Add(New SqlParameter("@ItemCode", SqlDbType.NVarChar))

        End If

        With dsCommand_form
            .SelectCommand = GetFormInfoCommand
            .SelectCommand.Transaction = ts
            GetFormInfoCommand.Parameters("@ProjectCode").Value = projectCode
            GetFormInfoCommand.Parameters("@ItemType").Value = itemType
            GetFormInfoCommand.Parameters("@ItemCode").Value = itemCode
            .Fill(tempDs, Table_form)
        End With

        If GetForm_additionalInfoCommand Is Nothing Then

            GetForm_additionalInfoCommand = New SqlCommand("GetProjectOppositeFormAdditional", conn)
            GetForm_additionalInfoCommand.CommandType = CommandType.StoredProcedure
            GetForm_additionalInfoCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            GetForm_additionalInfoCommand.Parameters.Add(New SqlParameter("@ItemType", SqlDbType.NVarChar))
            GetForm_additionalInfoCommand.Parameters.Add(New SqlParameter("@ItemCode", SqlDbType.NVarChar))

        End If

        With dsCommand_form_additional
            .SelectCommand = GetForm_additionalInfoCommand
            .SelectCommand.Transaction = ts
            GetForm_additionalInfoCommand.Parameters("@ProjectCode").Value = projectCode
            GetForm_additionalInfoCommand.Parameters("@ItemType").Value = itemType
            GetForm_additionalInfoCommand.Parameters("@ItemCode").Value = itemCode
            .Fill(tempDs, Table_form_additional)
        End With



        GetProjectGuaranteeFormAdditional = tempDs

    End Function

    '更新信息
    Public Function UpdateGuaranteeForm(ByVal formSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_form)

        With dsCommand_form

            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(formSet, Table_form)

        End With

    End Function

    '更新信息
    Public Function UpdateGuaranteeFormAdditional(ByVal form_additionalSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_form_additional)

        With dsCommand_form_additional

            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(form_additionalSet, Table_form_additional)

        End With

    End Function

End Class
