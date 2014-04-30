Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class GuarantyStatus
    Private moGuarantyStatusCommand As SqlCommand
    Private moConnection As SqlConnection
    Private moTransaction As SqlTransaction

    Public Sub New(ByRef Connection As SqlConnection)
        moConnection = Connection
        moTransaction = Nothing
    End Sub

    Public Sub New(ByRef Connection As SqlConnection, ByRef Transaction As SqlTransaction)
        moConnection = Connection
        moTransaction = Transaction
    End Sub

    Public Function GetGuarantyStatus(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("GuarantyStatusDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchGuarantyStatus", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = Condition

        moGuarantyStatusCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TGuarantyStatus")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function GetGuarantyStatus(ByVal Status As String, ByVal ID As Int32) As DataSet
        Return GetGuarantyStatus("{dbo.dd_guaranty_status.ID = " + ID.ToString() + " AND " + _
                "dbo.dd_guaranty_status.guaranty_status LIKE '" + Status + "'}")
    End Function

    Public Function UpdateGuarantyStatus(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        If dstCommit Is Nothing Then
            Return -1
        End If
        If Not dstCommit.HasChanges() Then
            Return 0
        End If

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        If moGuarantyStatusCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchGuarantyStatus", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moGuarantyStatusCommand.Connection = moConnection
            da.SelectCommand = moGuarantyStatusCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TGuarantyStatus")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
