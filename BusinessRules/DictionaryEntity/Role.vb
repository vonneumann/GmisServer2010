Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class Role
    Private moConnection As SqlConnection
    Private moTransaction As SqlTransaction

    Private moRoleCommand As SqlCommand
    Private moStaffRoleCommand As SqlCommand

    Public Sub New(ByRef Connection As SqlConnection)
        moConnection = Connection
        moTransaction = Nothing
    End Sub

    Public Sub New(ByRef Connection As SqlConnection, ByRef Transaction As SqlTransaction)
        moConnection = Connection
        moTransaction = Transaction
    End Sub

    Public Function FetchRole(ByVal RoleID As String) As DataSet
        Dim dstResult As DataSet = New DataSet("RoleDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchRole", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = RoleID

        moRoleCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TRole")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateRole(ByVal dstCommit As DataSet) As Int32
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

        If moRoleCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchRole", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moRoleCommand.Connection = moConnection
            da.SelectCommand = moRoleCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TRole")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function GetStaffRole(ByVal RoleID As String, ByVal StaffID As String) As DataSet
        Return GetStaffRole("{role_id LIKE '" + RoleID + "' AND dbo.staff_role.staff_name LIKE '" + StaffID + "'}")
    End Function

    Public Function GetStaffRole(ByVal RoleID As String) As DataSet
        Dim dstResult As DataSet = New DataSet("StaffRoleDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchStaffRole", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = RoleID

        moStaffRoleCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TStaffRole")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateStaffRole(ByVal dstCommit As DataSet) As Int32
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

        If moStaffRoleCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchStaffRole", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moStaffRoleCommand.Connection = moConnection
            da.SelectCommand = moStaffRoleCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TStaffRole")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
