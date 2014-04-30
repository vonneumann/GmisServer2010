Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class ProjectAccount
    Private moProjectAccountCommand As SqlCommand
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

    Public Function GetProjectAccount(ByVal ProjectNo As String, Optional ByVal SerialID As Int32 = -1) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectAccountDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchProjectAccount", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        If SerialID < 0 Then
            da.SelectCommand.Parameters("@Condition").Value = ProjectNo
        Else
            da.SelectCommand.Parameters("@Condition").Value = _
                        "{dbo.project_account_detail.project_code LIKE '" + ProjectNo + "' AND " + _
                        "dbo.project_account_detail.serial_num = " + SerialID.ToString() + "}"
        End If

        moProjectAccountCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TProjectAccount")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateProjectAccount(ByVal dstCommit As DataSet) As Int32
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

        If moProjectAccountCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchProjectAccount", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moProjectAccountCommand.Connection = moConnection
            da.SelectCommand = moProjectAccountCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TProjectAccount")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
