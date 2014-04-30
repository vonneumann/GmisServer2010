Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class ProjectGuaranteeForm
	Private moProjectGuaranteeFormCommand As SqlCommand
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

	Public Function GetProjectGuaranteeForm(ByVal Condition As String) As DataSet
		Dim dstResult As DataSet = New DataSet("ProjectGuaranteeFormDST")
		Dim da As SqlDataAdapter = New SqlDataAdapter()

		Try
			If moConnection.State = ConnectionState.Closed Then
				moConnection.Open()
			End If
		Catch ex As System.Exception
			Throw ex
		End Try

		da.SelectCommand = New SqlCommand("dbo.PFetchProjectGuaranteeForm", moConnection)
		da.SelectCommand.Transaction = moTransaction
		da.SelectCommand.CommandType = CommandType.StoredProcedure
		da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
		da.SelectCommand.Parameters("@Condition").Value = Condition

		moProjectGuaranteeFormCommand = da.SelectCommand

		Try
			da.Fill(dstResult, "TProjectGuaranteeForm")
		Catch ex As System.Exception
			Throw ex
		End Try

		Return dstResult
	End Function

	Public Function UpdateProjectGuaranteeForm(ByVal dstCommit As DataSet) As Int32
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

		If moProjectGuaranteeFormCommand Is Nothing Then
			da.SelectCommand = New SqlCommand("dbo.PFetchProjectGuaranteeForm", moConnection)
			da.SelectCommand.Transaction = moTransaction
			da.SelectCommand.CommandType = CommandType.StoredProcedure
			da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
			da.SelectCommand.Parameters("@Condition").Value = "NULL"
		Else
			moProjectGuaranteeFormCommand.Connection = moConnection
			da.SelectCommand = moProjectGuaranteeFormCommand
		End If

		Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

		da.InsertCommand = cmb.GetInsertCommand()
		da.InsertCommand.Transaction = moTransaction
		da.UpdateCommand = cmb.GetUpdateCommand()
		da.UpdateCommand.Transaction = moTransaction
		da.DeleteCommand = cmb.GetDeleteCommand()
		da.DeleteCommand.Transaction = moTransaction

		Try
			Return da.Update(dstCommit, "TProjectGuaranteeForm")
		Catch ex As System.Exception
			Throw ex
		End Try
    End Function

    '----------------------ProjectGuaranteeFormAdditional--------------------
    Public Function GetProjectGuaranteeFormAdd(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectGuaranteeFormAddDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchProjectGuaranteeFormAdd", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = Condition

        moProjectGuaranteeFormCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TProjectGuaranteeFormAdditional")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateProjectGuaranteeFormAdd(ByVal dstCommit As DataSet) As Int32
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

        If moProjectGuaranteeFormCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchProjectGuaranteeFormAdd", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moProjectGuaranteeFormCommand.Connection = moConnection
            da.SelectCommand = moProjectGuaranteeFormCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TProjectGuaranteeFormAdditional")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

End Class
