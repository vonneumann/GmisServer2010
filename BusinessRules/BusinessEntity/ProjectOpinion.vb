Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class ProjectOpinion
    Private moProjectOpinionCommand As SqlCommand
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

    Public Function GetProjectOpinion(ByVal SerialID As Int64) As DataSet
        Return GetProjectOpinion("{serial_num = " + SerialID.ToString() + "}")
    End Function

    Public Function GetProjectOpinion(ByVal ProjectNo As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectOpinionDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        If ProjectNo Is Nothing Then
            ProjectNo = "NULL"
        End If

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchProjectOpinion", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = ProjectNo

        moProjectOpinionCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TProjectOpinion")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateProjectOpinion(ByVal dstCommit As DataSet) As Int32
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

        If moProjectOpinionCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchProjectOpinion", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moProjectOpinionCommand.Connection = moConnection
            da.SelectCommand = moProjectOpinionCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TProjectOpinion")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
