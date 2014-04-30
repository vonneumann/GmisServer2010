Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class ProjectFile
    Private moProjectFileCommand As SqlCommand
    Private moProjectFileImageCommand As SqlCommand
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

    Public Function GetProjectFile(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectFileDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchProjectFile", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = Condition

        moProjectFileCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TProjectFile")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function GetProjectFileImage(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectFileImageDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchProjectFileImage", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = Condition

        moProjectFileImageCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TProjectFile")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function GetProjectFile(ByVal ProjectNo As String, ByVal ItemNo As String, ByVal ItemTypeNo As String) As DataSet
        Return GetProjectFile("{dbo.project_files.project_code LIKE '" + ProjectNo + "' AND " + _
                "dbo.project_files.item_code LIKE '" + ItemNo + "' AND " + _
                "dbo.project_files.item_type LIKE '" + ItemTypeNo + "'}")
    End Function

    Public Function GetProjectFileImage(ByVal ProjectNo As String, ByVal ItemNo As String, ByVal ItemTypeNo As String) As DataSet
        Return GetProjectFileImage("{dbo.project_files.project_code LIKE '" + ProjectNo + "' AND " + _
                "dbo.project_files.item_code LIKE '" + ItemNo + "' AND " + _
                "dbo.project_files.item_type LIKE '" + ItemTypeNo + "'}")
    End Function

    Public Function UpdateProjectFile(ByVal dstCommit As DataSet) As Int32
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

        If moProjectFileCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchProjectFile", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moProjectFileCommand.Connection = moConnection
            da.SelectCommand = moProjectFileCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TProjectFile")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function UpdateProjectFileImage(ByVal dstCommit As DataSet) As Int32
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

        If moProjectFileCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchProjectFileImage", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moProjectFileImageCommand.Connection = moConnection
            da.SelectCommand = moProjectFileImageCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TProjectFile")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function GetRelationID() As Int64
        Dim dr As SqlDataReader

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        Dim insertCommand As SqlCommand = New SqlCommand("INSERT INTO dbo.file_relative_num (relation_num) SELECT ISNULL(MAX(relation_num), 0) + 1 FROM dbo.file_relative_num", moConnection)
        Dim selectCommand As SqlCommand = New SqlCommand("SELECT TOP 1 relation_num FROM dbo.file_relative_num ORDER BY relation_num DESC", moConnection)

        If Not moTransaction.Connection Is Nothing Then
            insertCommand.Transaction = moTransaction
            selectCommand.Transaction = moTransaction
        End If

        Try
            insertCommand.ExecuteNonQuery()

            'If Not moTransaction Is Nothing Then
            '    If Not moTransaction.Connection Is Nothing Then
            '        moTransaction.Commit()
            '    End If
            'End If

            dr = selectCommand.ExecuteReader()

            If (dr.Read()) Then
                Return dr.GetInt64(0)
            Else
                Return -1
            End If
        Catch ex As System.Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Close()
            End If
        End Try
    End Function
End Class
