Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class Material
    Private moMaterialCommand As SqlCommand
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

    Public Sub DuplicateMaterial(ByVal sourceServiceType As String, ByVal destinationServiceType As String)
        Try
            Dim createCommand As SqlCommand = New SqlCommand("dbo.PDuplicateMaterial", moConnection, moTransaction)
            createCommand.CommandType = CommandType.StoredProcedure

            createCommand.Parameters.Add("@SourceServiceType", SqlDbType.NVarChar, 20)
            createCommand.Parameters.Add("@DestinationServiceType", SqlDbType.NVarChar, 20)

            createCommand.Parameters("@SourceServiceType").Value = sourceServiceType
            createCommand.Parameters("@DestinationServiceType").Value = destinationServiceType

            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            If createCommand.ExecuteNonQuery() = 0 Then
                Throw New System.Exception("复制文档材料出错。")
            End If
        Catch ex As System.Exception
            Throw ex
        End Try
    End Sub

    Public Function GetMaterial(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("MaterialDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchMaterial", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = Condition

        moMaterialCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TMaterial")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function GetMaterial(ByVal itemNo As String, ByVal itemTypeNo As String, ByVal serviceType As String) As DataSet
        Dim condition As String

        If Not itemNo Is Nothing Then
            condition = "dbo.material.item_code LIKE '" + itemNo + "'"
        End If

        If Not itemTypeNo Is Nothing Then
            If Not condition Is Nothing AndAlso condition.Trim() <> String.Empty Then
                condition += " AND "
            End If

            condition += "dbo.material.item_type LIKE '" + itemTypeNo + "'"
        End If

        If Not serviceType Is Nothing Then
            If Not condition Is Nothing AndAlso condition.Trim() <> String.Empty Then
                condition += " AND "
            End If

            condition += "dbo.material.service_type LIKE '" + serviceType + "'"
        End If

        If Not condition Is Nothing AndAlso condition.Trim() <> String.Empty Then
            condition = "{" + condition + "}"
        End If

        Return Me.GetMaterial(condition)
    End Function

    Public Function UpdateMaterial(ByVal dstCommit As DataSet) As Int32
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

        If moMaterialCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchMaterial", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moMaterialCommand.Connection = moConnection
            da.SelectCommand = moMaterialCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TMaterial")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
