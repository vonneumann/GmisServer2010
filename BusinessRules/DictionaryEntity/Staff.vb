Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class Staff
    Private moConnection As SqlConnection
    Private moTransaction As SqlTransaction

    Private moStaffCommand As SqlCommand

    Public Sub New(ByRef Connection As SqlConnection)
        moConnection = Connection
        moTransaction = Nothing
    End Sub

    Public Sub New(ByRef Connection As SqlConnection, ByRef Transaction As SqlTransaction)
        moConnection = Connection
        moTransaction = Transaction
    End Sub

    Public Function FetchStaff(ByVal StaffID As String) As DataSet
        Dim dstResult As DataSet = New DataSet("StaffDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchStaff", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = StaffID

        moStaffCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TStaff")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function FetchStaffEx(ByVal TeamID As String) As DataSet
        Dim dstResult As DataSet = New DataSet("StaffExDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchStaffEx", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = TeamID

        moStaffCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TStaffEx")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateStaff(ByVal dstCommit As DataSet) As Int32
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

        If moStaffCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchStaff", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moStaffCommand.Connection = moConnection
            da.SelectCommand = moStaffCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TStaff")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
