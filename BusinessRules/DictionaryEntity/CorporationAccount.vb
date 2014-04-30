Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class CorporationAccount
    Private moCorporationAccountCommand As SqlCommand
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

    Public Function FetchCorporationAccountCredit(Optional ByVal ProjectNo As String = Nothing) As DataSet
        Dim dsResult As DataSet = New DataSet("CorporationAccountCreditDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationAccountCredit", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@ProjectNo", SqlDbType.VarChar, 25)
        da.SelectCommand.Parameters("@ProjectNo").Value = ProjectNo

        Try
            da.Fill(dsResult)

            If dsResult.Tables.Count > 1 Then
                dsResult.Tables(0).TableName = "TProjectCredit"
                dsResult.Tables(1).TableName = "TProjectCorporationCredit"

                dsResult.Relations.Add("TProjectCredit2TProjectCorporationCredit", dsResult.Tables(0).Columns("project_code"), dsResult.Tables(1).Columns("project_code"))
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dsResult
    End Function

    Public Function FetchCorporationAccountCreditEx(Optional ByVal ProjectNo As String = Nothing) As DataSet
        Dim dsResult As DataSet = New DataSet("CorporationAccountCreditDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationAccountCreditEx", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@ProjectNo", SqlDbType.VarChar, 25)
        da.SelectCommand.Parameters("@ProjectNo").Value = ProjectNo

        Try
            da.Fill(dsResult, "TCorporationAccount")

        Catch ex As System.Exception
            Throw ex
        End Try

        Return dsResult
    End Function


    Public Function FetchCorporationAccountMonth(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String) As DataSet
        Dim dsResult As DataSet = New DataSet("CorporationAccountMonthDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationAccountMonth", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = "{dbo.corporation_account.project_code LIKE '" + ProjectNo + "' AND " + _
          " dbo.corporation_account.corporation_code LIKE '" + CorporationNo + "' AND " + _
          " dbo.corporation_account.phase LIKE '" + Phase + "'}"

        Try
            da.Fill(dsResult, "TCorporationAccountMonth")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dsResult
    End Function

    Public Function GetCorporationAccount(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("CorporationAccountDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationAccount", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = Condition

        moCorporationAccountCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TCorporationAccount")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function GetCorporationAccount(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal ItemNo As String, ByVal ItemType As String) As DataSet
        Return Me.GetCorporationAccount("{dbo.corporation_account.project_code LIKE '" + ProjectNo + "' AND " + _
         " dbo.corporation_account.corporation_code LIKE '" + CorporationNo + "' AND " + _
         " dbo.corporation_account.phase LIKE '" + Phase + "' AND " + _
          " dbo.corporation_account.month LIKE '" + Month + "' AND " + _
          " dbo.corporation_account.item_code LIKE '" + ItemNo + "' AND " + _
          " dbo.corporation_account.item_type LIKE '" + ItemType + "'}")
    End Function

    Public Function UpdateCorporationAccount(ByVal dstCommit As DataSet) As Int32
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

        If moCorporationAccountCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCorporationAccount", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCorporationAccountCommand.Connection = moConnection
            da.SelectCommand = moCorporationAccountCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TCorporationAccount")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
