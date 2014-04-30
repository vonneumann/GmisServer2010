Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

'----------设计人：zhoufucai
Public Class LoanChargeManner

    Private moLoanChargeMannerCommand As SqlCommand
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

    '2006-4-21 By zhoufucai
    '获取项目收费方式
    Public Function GetLoanChargeManner(ByVal LoanChargeMannerNo As String) As DataSet
        Dim dstResult As DataSet = New DataSet("LoanChargeMannerDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchLoanChargeManner", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = LoanChargeMannerNo

        moLoanChargeMannerCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TLoanChargeManner")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    '2006-4-21 By zhoufucai
    '修改项目收费方式
    Public Function UpdateLoanChargeManner(ByVal dstCommit As DataSet) As Int32
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

        If moLoanChargeMannerCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchLoanChargeManner", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moLoanChargeMannerCommand.Connection = moConnection
            da.SelectCommand = moLoanChargeMannerCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TLoanChargeManner")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
