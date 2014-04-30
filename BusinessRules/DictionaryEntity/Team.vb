Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class Team
    Private moConnection As SqlConnection
    Private moTransaction As SqlTransaction

    Private moTeamCommand As SqlCommand
    Private moStaffTeamCommand As SqlCommand

    Public Sub New(ByRef Connection As SqlConnection)
        moConnection = Connection
        moTransaction = Nothing
    End Sub

    Public Sub New(ByRef Connection As SqlConnection, ByRef Transaction As SqlTransaction)
        moConnection = Connection
        moTransaction = Transaction
    End Sub

    Public Function FetchTeam(ByVal TeamID As String) As DataSet
        Dim dstResult As DataSet = New DataSet("TeamDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchTeam", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = TeamID

        moTeamCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TTeam")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateTeam(ByVal dstCommit As DataSet) As Int32
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

        If moTeamCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchTeam", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moTeamCommand.Connection = moConnection
            da.SelectCommand = moTeamCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TTeam")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function GetStaffTeam(ByVal TeamID As String, ByVal StaffID As String) As DataSet
        Return GetStaffTeam("{team_name LIKE '" + TeamID + "' AND staff_name LIKE '" + StaffID + "'}")
    End Function

    Public Function GetStaffTeam(ByVal TeamID As String) As DataSet
        Dim dstResult As DataSet = New DataSet("StaffTeamDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchStaffTeam", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = TeamID

        moStaffTeamCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TStaffTeam")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateStaffTeam(ByVal dstCommit As DataSet) As Int32
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

        If moStaffTeamCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchStaffTeam", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moStaffTeamCommand.Connection = moConnection
            da.SelectCommand = moStaffTeamCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TStaffTeam")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
