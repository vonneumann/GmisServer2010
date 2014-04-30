Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class ProjectDocument
    Private moProjectDocumentCommand As SqlCommand
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

    Public Function GetProjectDocument(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectDocumentDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchProjectDocument", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = Condition

        moProjectDocumentCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TProjectDocument")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function GetProjectDocument(ByVal ProjectNo As String, ByVal Phase As String, ByVal ItemNo As String, ByVal ItemTypeNo As String) As DataSet
        Return GetProjectDocument("{dbo.project_document.project_code LIKE '" + ProjectNo + "' AND " + _
                "dbo.project_document.phase LIKE '" + Phase + "' AND " + _
                "dbo.project_document.item_code LIKE '" + ItemNo + "' AND " + _
                "dbo.project_document.item_type LIKE '" + ItemTypeNo + "'}")
    End Function

    Public Function UpdateProjectDocument(ByVal dstCommit As DataSet) As Int32
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

        If moProjectDocumentCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchProjectDocument", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moProjectDocumentCommand.Connection = moConnection
            da.SelectCommand = moProjectDocumentCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TProjectDocument")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
