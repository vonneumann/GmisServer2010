Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectSignature

    Public Const Table_Project_Signature As String = "project_signature"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ProjectSignature As SqlDataAdapter

    '定义查询命令
    Private GetProjectSignatureInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ProjectSignature = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectSignatureInfo("null")
    End Sub

    '获取签约信息
    Public Function GetProjectSignatureInfo(ByVal strSQL_Condition_ProjectSignature As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectSignatureInfoCommand Is Nothing Then

            GetProjectSignatureInfoCommand = New SqlCommand("GetProjectSignatureInfo", conn)
            GetProjectSignatureInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectSignatureInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectSignature
            .SelectCommand = GetProjectSignatureInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectSignatureInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectSignature
            .Fill(tempDs, Table_Project_Signature)
        End With

        Return tempDs

    End Function


    '更新签约信息
    Public Function UpdateProjectSignature(ByVal ProjectSignatureSet As DataSet)


        If ProjectSignatureSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ProjectSignatureSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectSignature)

        With dsCommand_ProjectSignature
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectSignatureSet, Table_Project_Signature)

            ProjectSignatureSet.AcceptChanges()
        End With


    End Function

End Class
