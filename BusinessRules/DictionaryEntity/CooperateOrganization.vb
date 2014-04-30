Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class CooperateOrganization
	Private moCooperateOrganizationCommand As SqlCommand
	Private moCooperateOrganizationOpinionCommand As SqlCommand
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

	Public Function GetCooperateOrganization(ByVal Condition As String) As DataSet
		Dim dstResult As DataSet = New DataSet("CooperateOrganizationDST")
		Dim da As SqlDataAdapter = New SqlDataAdapter()

		Try
			If moConnection.State = ConnectionState.Closed Then
				moConnection.Open()
			End If
		Catch ex As System.Exception
			Throw ex
		End Try

		da.SelectCommand = New SqlCommand("dbo.PFetchCooperateOrganization", moConnection)
		da.SelectCommand.Transaction = moTransaction
		da.SelectCommand.CommandType = CommandType.StoredProcedure
		da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
		da.SelectCommand.Parameters("@Condition").Value = Condition

		moCooperateOrganizationCommand = da.SelectCommand

		Try
			da.Fill(dstResult, "TCooperateOrganization")
		Catch ex As System.Exception
			Throw ex
		End Try

		Return dstResult
	End Function

	Public Function UpdateCooperateOrganization(ByVal dstCommit As DataSet) As Int32
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

		If moCooperateOrganizationCommand Is Nothing Then
			da.SelectCommand = New SqlCommand("dbo.PFetchCooperateOrganization", moConnection)
			da.SelectCommand.Transaction = moTransaction
			da.SelectCommand.CommandType = CommandType.StoredProcedure
			da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
			da.SelectCommand.Parameters("@Condition").Value = "NULL"
		Else
			moCooperateOrganizationCommand.Connection = moConnection
			da.SelectCommand = moCooperateOrganizationCommand
		End If

		Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

		da.InsertCommand = cmb.GetInsertCommand()
		da.InsertCommand.Transaction = moTransaction
		da.UpdateCommand = cmb.GetUpdateCommand()
		da.UpdateCommand.Transaction = moTransaction
		da.DeleteCommand = cmb.GetDeleteCommand()
		da.DeleteCommand.Transaction = moTransaction

		Try
			Return da.Update(dstCommit, "TCooperateOrganization")
		Catch ex As System.Exception
			Throw ex
		End Try
	End Function

	Public Function GetCooperateOrganizationOpinion(ByVal Condition As String) As DataSet
		Dim dstResult As DataSet = New DataSet("CooperateOrganizationOpinionDST")
		Dim da As SqlDataAdapter = New SqlDataAdapter()

		Try
			If moConnection.State = ConnectionState.Closed Then
				moConnection.Open()
			End If
		Catch ex As System.Exception
			Throw ex
		End Try

		da.SelectCommand = New SqlCommand("dbo.PFetchCooperateOrganizationOpinion", moConnection)
		da.SelectCommand.Transaction = moTransaction
		da.SelectCommand.CommandType = CommandType.StoredProcedure
		da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
		da.SelectCommand.Parameters("@Condition").Value = Condition

		moCooperateOrganizationOpinionCommand = da.SelectCommand

		Try
			da.Fill(dstResult, "TCooperateOrganizationOpinion")
		Catch ex As System.Exception
			Throw ex
		End Try

		Return dstResult
	End Function

	Public Function UpdateCooperateOrganizationOpinion(ByVal dstCommit As DataSet) As Int32
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

		If moCooperateOrganizationOpinionCommand Is Nothing Then
			da.SelectCommand = New SqlCommand("dbo.PFetchCooperateOrganizationOpinion", moConnection)
			da.SelectCommand.Transaction = moTransaction
			da.SelectCommand.CommandType = CommandType.StoredProcedure
			da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
			da.SelectCommand.Parameters("@Condition").Value = "NULL"
		Else
			moCooperateOrganizationOpinionCommand.Connection = moConnection
			da.SelectCommand = moCooperateOrganizationOpinionCommand
		End If

		Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

		da.InsertCommand = cmb.GetInsertCommand()
		da.InsertCommand.Transaction = moTransaction
		da.UpdateCommand = cmb.GetUpdateCommand()
		da.UpdateCommand.Transaction = moTransaction
		da.DeleteCommand = cmb.GetDeleteCommand()
		da.DeleteCommand.Transaction = moTransaction

		Try
			Return da.Update(dstCommit, "TCooperateOrganizationOpinion")
		Catch ex As System.Exception
			Throw ex
		End Try
	End Function
End Class
