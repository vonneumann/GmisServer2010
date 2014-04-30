Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class Item
    Private moItemCommand As SqlCommand
    Private moItemTypeCommand As SqlCommand
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

    Public Function GetItem(ByVal ItemNo As String, Optional ByVal ItemTypeNo As String = Nothing) As DataSet
        Dim dstResult As DataSet = New DataSet("ItemDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        If ItemNo Is Nothing Then
            ItemNo = "NULL"
        End If

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchItem", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        If ItemTypeNo Is Nothing Then
            da.SelectCommand.Parameters("@Condition").Value = ItemNo
        Else
            da.SelectCommand.Parameters("@Condition").Value = _
                            "{dbo.item.item_code LIKE '" + ItemNo + "' AND " + _
                            "dbo.item.item_type LIKE '" + ItemTypeNo + "'}"
        End If

        moItemCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TItem")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateItem(ByVal dstCommit As DataSet) As Int32
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

        If moItemCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchItem", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moItemCommand.Connection = moConnection
            da.SelectCommand = moItemCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TItem")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function GetItemType(ByVal ItemTypeNo As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ItemDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        If ItemTypeNo Is Nothing Then
            ItemTypeNo = "NULL"
        End If

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchItemType", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = ItemTypeNo

        moItemTypeCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TItemType")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateItemType(ByVal dstCommit As DataSet) As Int32
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

        If moItemTypeCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchItemType", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moItemTypeCommand.Connection = moConnection
            da.SelectCommand = moItemTypeCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TItemType")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function GetItemEx(ByVal ItemNo As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ItemExDST")
        Dim dstTemp As DataSet

        dstResult = Me.GetItem(ItemNo)
        dstTemp = Me.GetItemType("%")
        dstResult.Merge(dstTemp)

        Dim ParentCol As DataColumn
        Dim ChildCol As DataColumn

        '设置主表(Item_type)的关联字段集
        ParentCol = dstResult.Tables("TItemType").Columns("item_type")

        '设置从表(Item)的关联字段集
        ChildCol = dstResult.Tables("TItem").Columns("item_type")

        '设置数据集合中“项目企业信息”表与“企业财务报表”表的关联关系
        dstResult.Relations.Add("ItemType_Item", ParentCol, ChildCol)

        Return dstResult
    End Function
End Class
