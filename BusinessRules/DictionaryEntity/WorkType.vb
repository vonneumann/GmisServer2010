Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class WorkType
    Private moWorkTypeCommand As SqlCommand
    Private moSubWorkTypeCommand As SqlCommand
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

	Public Function GetWorkType(ByVal Condition As String) As DataSet
		Dim dstResult As DataSet = New DataSet("WorkTypeDST")
		Dim da As SqlDataAdapter = New SqlDataAdapter()

		Try
			If moConnection.State = ConnectionState.Closed Then
				moConnection.Open()
			End If
		Catch ex As System.Exception
			Throw ex
		End Try

		da.SelectCommand = New SqlCommand("dbo.PFetchWorkType", moConnection)
		da.SelectCommand.Transaction = moTransaction
		da.SelectCommand.CommandType = CommandType.StoredProcedure
		da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
		da.SelectCommand.Parameters("@Condition").Value = Condition

		moWorkTypeCommand = da.SelectCommand

		Try
			da.Fill(dstResult, "TWorkType")
		Catch ex As System.Exception
			Throw ex
		End Try

		Return dstResult
	End Function

	Public Function UpdateWorkType(ByVal dstCommit As DataSet) As Int32
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

		If moWorkTypeCommand Is Nothing Then
			da.SelectCommand = New SqlCommand("dbo.PFetchWorkType", moConnection)
			da.SelectCommand.Transaction = moTransaction
			da.SelectCommand.CommandType = CommandType.StoredProcedure
			da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
			da.SelectCommand.Parameters("@Condition").Value = "NULL"
		Else
			moWorkTypeCommand.Connection = moConnection
			da.SelectCommand = moWorkTypeCommand
		End If

		Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

		da.InsertCommand = cmb.GetInsertCommand()
		da.InsertCommand.Transaction = moTransaction
		da.UpdateCommand = cmb.GetUpdateCommand()
		da.UpdateCommand.Transaction = moTransaction
		da.DeleteCommand = cmb.GetDeleteCommand()
		da.DeleteCommand.Transaction = moTransaction

		Try
			Return da.Update(dstCommit, "TWorkType")
		Catch ex As System.Exception
			Throw ex
		End Try
    End Function

    Public Function GetWorkSubType(ByVal typeCode As String, ByVal subTypeCode As String) As DataSet
        If subTypeCode Is Nothing OrElse subTypeCode.Trim() = String.Empty Then
            Return Me.GetWorkSubType("{type_code LIKE '" + typeCode + "'}")
        Else
            Return Me.GetWorkSubType("{type_code LIKE '" + typeCode + "' AND subtype_code LIKE '" + subTypeCode + "'}")
        End If
    End Function

    Public Function GetWorkSubType(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("WorkSubTypeDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchWorkSubType", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = Condition

        moSubWorkTypeCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TWorkSubType")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateWorkSubType(ByVal dstCommit As DataSet) As Int32
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

        If moSubWorkTypeCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchWorkSubType", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moSubWorkTypeCommand.Connection = moConnection
            da.SelectCommand = moSubWorkTypeCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TWorkSubType")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
