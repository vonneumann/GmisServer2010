Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class ProjectCorporation
    Private moConnection As SqlConnection
    Private moTransaction As SqlTransaction

    Private moProjectCorporationCommand As SqlCommand
    Private moCorporationLoanCommand As SqlCommand
    Private moCorporationAccountCommand As SqlCommand
    Private moCorporationBusinessCommand As SqlCommand
    Private moCorporationBankSavingCommand As SqlCommand
    Private moCorporationStockStructureCommand As SqlCommand
    Private moCorporationExternalGuaranteeCommand As SqlCommand
    Private moCorporationLawsuitRecordCommand As SqlCommand
    Private moCorporationRatepayingRecordCommand As SqlCommand
	Private moCorporationPostalOrderCommand As SqlCommand

    Public Sub New(ByRef Connection As SqlConnection)
        moConnection = Connection
        moTransaction = Nothing
    End Sub

    Public Sub New(ByRef Connection As SqlConnection, ByRef Transaction As SqlTransaction)
        moConnection = Connection
        moTransaction = Transaction
    End Sub

    Public Function GetSchema(ByVal TableName As String) As DataSet
        Dim dstSchema As DataSet = New DataSet("SchemaDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("SELECT * FROM " + TableName + " WHERE 1=0", moConnection)
            da.FillSchema(dstSchema, SchemaType.Source, TableName)

        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstSchema
    End Function

	Public Function FetchProjectCorporation(ByVal Condition As String) As DataSet
		Dim dstResult As DataSet = New DataSet("ProjectCorporationDST")
		Dim da As SqlDataAdapter = New SqlDataAdapter()

		Try
			If moConnection.State = ConnectionState.Closed Then
				moConnection.Open()
			End If
		Catch ex As System.Exception
			Throw ex
		End Try

		da.SelectCommand = New SqlCommand("dbo.PFetchProjectCorporation", moConnection)
		da.SelectCommand.Transaction = moTransaction
		da.SelectCommand.CommandType = CommandType.StoredProcedure
		da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
		da.SelectCommand.Parameters("@Condition").Value = Condition

		moProjectCorporationCommand = da.SelectCommand

		Try
			da.Fill(dstResult, "TProjectCorporation")
		Catch ex As System.Exception
			Throw ex
		End Try

		Return dstResult
	End Function

    'Public Function FetchProjectCorporation(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationType As String, ByVal Phase As String) As DataSet
    '	Return Me.FetchProjectCorporation( _
    '	 "{dbo.project_corporation.project_code LIKE '" + ProjectNo + "' AND " + _
    '	 "dbo.project_corporation.corporation_code LIKE '" + CorporationNo + "' AND " + _
    '	 "dbo.project_corporation.corporation_type LIKE '" + CorporationType + "' AND " + _
    '	 "dbo.project_corporation.phase LIKE '" + Phase + "'}")
    '   End Function

    Public Function FetchProjectCorporation(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationType As String, ByVal Phase As String) As DataSet
        Return Me.FetchProjectCorporation( _
         "{dbo.project_corporation.project_code LIKE '" + ProjectNo + "' AND " + _
         "dbo.project_corporation.corporation_code LIKE '" + CorporationNo + "' AND " + _
         "dbo.project_corporation.corporation_type LIKE '" + CorporationType + "'}")
    End Function

    Public Function UpdateProjectCorporation(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter

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

        If moProjectCorporationCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchProjectCorporation", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moProjectCorporationCommand.Connection = moConnection
            da.SelectCommand = moProjectCorporationCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        'if(da.SelectCommand != null)
        '	System.Windows.Forms.MessageBox.Show(da.SelectCommand.CommandText, "SelectCommand");
        'if(da.UpdateCommand != null)
        '	System.Windows.Forms.MessageBox.Show(da.UpdateCommand.CommandText, "UpdateCommand");
        'if(da.InsertCommand != null)
        '	System.Windows.Forms.MessageBox.Show(da.InsertCommand.CommandText, "InsertCommand");
        'if(da.DeleteCommand != null)
        '	System.Windows.Forms.MessageBox.Show(da.DeleteCommand.CommandText, "DeleteCommand");

        Try
            Return da.Update(dstCommit, "TProjectCorporation")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchCorporationAccount(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("CorporationAccountDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

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

    Public Function FetchCorporationAccount(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal ItemType As String, ByVal ItemCode As String) As DataSet
        Return Me.FetchCorporationAccount("{dbo.corporation_account.project_code LIKE '" + ProjectNo + "' AND " + _
                  " dbo.corporation_account.corporation_code LIKE '" + CorporationNo + "' AND " + _
                  " dbo.corporation_account.phase LIKE '" + Phase + "' AND " + _
                  " dbo.corporation_account.month LIKE '" + Month + "' AND " + _
                  " dbo.corporation_account.item_type LIKE '" + ItemType + "' AND" + _
                  " dbo.corporation_account.item_code LIKE '" + ItemCode + "'}")
    End Function

    Public Function UpdateCorporationAccount(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter

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
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TCorporationAccount")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchCorporationBusiness(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String) As DataSet
        Dim dstResult As DataSet = New DataSet("CorporationBusinessDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationBusiness", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        da.SelectCommand.Parameters("@Condition").Value = _
          "{dbo.corporation_business.project_code LIKE '" + ProjectNo + "' AND " + _
          " dbo.corporation_business.corporation_code LIKE '" + CorporationNo + "' AND " + _
          " dbo.corporation_business.phase LIKE '" + Phase + "' AND " + _
          " dbo.corporation_business.month LIKE '" + Month + "'}"

        moCorporationBusinessCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TCorporationBusiness")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateCorporationBusiness(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter

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

        If moCorporationBusinessCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCorporationBusiness", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCorporationBusinessCommand.Connection = moConnection
            da.SelectCommand = moCorporationBusinessCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TCorporationBusiness")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchCorporationLoan(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, Optional ByVal SerialID As Int64 = -1) As DataSet
        Dim dstResult As DataSet = New DataSet("CorporationLoanDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationLoan", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        If SerialID < 0 Then
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_loan.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_loan.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_loan.phase LIKE '" + Phase + "'}"
        Else
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_loan.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_loan.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_loan.phase LIKE '" + Phase + "' AND " + _
               " dbo.corporation_loan.serial_num = " + SerialID.ToString() + "}"
        End If

        moCorporationLoanCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TCorporationLoan")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateCorporationLoan(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter

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

        If moCorporationLoanCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCorporationLoan", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCorporationLoanCommand.Connection = moConnection
            da.SelectCommand = moCorporationLoanCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TCorporationLoan")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchCorporationExternalGuarantee(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, Optional ByVal SerialID As Int64 = -1) As DataSet
        Dim dstResult As DataSet = New DataSet("CorporationExternalGuaranteeDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationExternalGuarantee", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        If SerialID < 0 Then
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_external_guarantee.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_external_guarantee.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_external_guarantee.phase LIKE '" + Phase + "'}"
        Else
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_external_guarantee.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_external_guarantee.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_external_guarantee.phase LIKE '" + Phase + "' AND " + _
               " dbo.corporation_external_guarantee.serial_num = " + SerialID.ToString() + "}"
        End If

        moCorporationExternalGuaranteeCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TCorporationExternalGuarantee")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateCorporationExternalGuarantee(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter

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

        If moCorporationExternalGuaranteeCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCorporationExternalGuarantee", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCorporationExternalGuaranteeCommand.Connection = moConnection
            da.SelectCommand = moCorporationExternalGuaranteeCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TCorporationExternalGuarantee")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchCorporationStockStructure(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, Optional ByVal SerialID As Int64 = -1) As DataSet
        Dim dstResult As DataSet = New DataSet("CorporationStockStructureDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationStockStructure", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        If SerialID < 0 Then
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_stock_structure.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_stock_structure.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_stock_structure.phase LIKE '" + Phase + "'}"
        Else
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_stock_structure.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_stock_structure.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_stock_structure.phase LIKE '" + Phase + "' AND " + _
               " dbo.corporation_stock_structure.serial_num = " + SerialID.ToString() + "}"
        End If

        moCorporationStockStructureCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TCorporationStockStructure")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateCorporationStockStructure(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter

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

        If moCorporationStockStructureCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCorporationStockStructure", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCorporationStockStructureCommand.Connection = moConnection
            da.SelectCommand = moCorporationStockStructureCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TCorporationStockStructure")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function GetMaxSerialID(ByVal FieldName As String, ByVal TableName As String) As Int64
        Dim dr As SqlDataReader

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            Dim getCommand As SqlCommand = New SqlCommand("SELECT MAX(" + FieldName + ") FROM " + TableName, moConnection)
            getCommand.Transaction = moTransaction
            dr = getCommand.ExecuteReader()

            If dr.Read() Then
                If dr.IsDBNull(0) Then
                    Return 1
                Else
                    Return dr.GetInt64(0)
                End If
            Else
                Return 1
            End If
        Catch ex As System.Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Close()
            End If

            'If Not moConnection Is Nothing Then
            '	If moConnection.State <> ConnectionState.Closed Then
            '		moConnection.Close()
            '	End If
            'End If
        End Try
    End Function

    Public Function FetchCorporationBankSaving(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, Optional ByVal SerialID As Int64 = -1) As DataSet
        Dim dstResult As DataSet = New DataSet("CorporationBankSavingDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationBankSaving", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        If SerialID < 0 Then
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_bank_saving.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_bank_saving.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_bank_saving.phase LIKE '" + Phase + "'}"
        Else
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_bank_saving.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_bank_saving.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_bank_saving.phase LIKE '" + Phase + "' AND " + _
               " dbo.corporation_bank_saving.serial_num = " + SerialID.ToString() + "}"
        End If

        moCorporationBankSavingCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TCorporationBankSaving")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateCorporationBankSaving(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter

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

        If moCorporationBankSavingCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCorporationBankSaving", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCorporationBankSavingCommand.Connection = moConnection
            da.SelectCommand = moCorporationBankSavingCommand
        End If
        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TCorporationBankSaving")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchCorporationLawsuitRecord(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, Optional ByVal SerialID As Int64 = -1) As DataSet
        Dim dstResult As DataSet = New DataSet("CorporationLawsuitRecordDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationLawsuitRecord", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        If SerialID < 0 Then
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_lawsuit_record.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_lawsuit_record.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_lawsuit_record.phase LIKE '" + Phase + "'}"
        Else
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_lawsuit_record.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_lawsuit_record.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_lawsuit_record.phase LIKE '" + Phase + "' AND " + _
               " dbo.corporation_lawsuit_record.serial_num = " + SerialID.ToString() + "}"
        End If

        moCorporationLawsuitRecordCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TCorporationLawsuitRecord")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateCorporationLawsuitRecord(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter

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

        If moCorporationLawsuitRecordCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCorporationLawsuitRecord", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCorporationLawsuitRecordCommand.Connection = moConnection
            da.SelectCommand = moCorporationLawsuitRecordCommand
        End If
        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TCorporationLawsuitRecord")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchCorporationRatepayingRecord(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, Optional ByVal SerialID As Int64 = -1) As DataSet
        Dim dstResult As DataSet = New DataSet("CorporationRatepayingRecordDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationRatepayingRecord", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        If SerialID < 0 Then
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_ratepaying_record.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_ratepaying_record.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_ratepaying_record.phase LIKE '" + Phase + "'}"
        Else
            da.SelectCommand.Parameters("@Condition").Value = _
              "{dbo.corporation_ratepaying_record.project_code LIKE '" + ProjectNo + "' AND " + _
             " dbo.corporation_ratepaying_record.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_ratepaying_record.phase LIKE '" + Phase + "' AND " + _
               " dbo.corporation_ratepaying_record.serial_num = " + SerialID.ToString() + "}"
        End If

        moCorporationRatepayingRecordCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TCorporationRatepayingRecord")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateCorporationRatepayingRecord(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter

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

        If moCorporationRatepayingRecordCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCorporationRatepayingRecord", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCorporationRatepayingRecordCommand.Connection = moConnection
            da.SelectCommand = moCorporationRatepayingRecordCommand
        End If
        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dstCommit, "TCorporationRatepayingRecord")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchCorporationPostalOrder(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, Optional ByVal SerialID As Int64 = -1) As DataSet
        Dim dstResult As DataSet = New DataSet("CorporationPostalOrderDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        da.SelectCommand = New SqlCommand("dbo.PFetchCorporationPostalOrder", moConnection)
        da.SelectCommand.Transaction = moTransaction
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
        If SerialID < 0 Then
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_postal_order.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_postal_order.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_postal_order.phase LIKE '" + Phase + "'}"
        Else
            da.SelectCommand.Parameters("@Condition").Value = _
              "{dbo.corporation_postal_order.project_code LIKE '" + ProjectNo + "' AND " + _
             " dbo.corporation_postal_order.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_postal_order.phase LIKE '" + Phase + "' AND " + _
               " dbo.corporation_postal_order.serial_num = " + SerialID.ToString() + "}"
        End If

        moCorporationPostalOrderCommand = da.SelectCommand

        Try
            da.Fill(dstResult, "TCorporationPostalOrder")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateCorporationPostalOrder(ByVal dstCommit As DataSet) As Int32
        Dim da As SqlDataAdapter = New SqlDataAdapter

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

        If moCorporationPostalOrderCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCorporationPostalOrder", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCorporationPostalOrderCommand.Connection = moConnection
            da.SelectCommand = moCorporationPostalOrderCommand
        End If
        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand.Transaction = moTransaction

        Try
            da.Update(dstCommit, "TCorporationPostalOrder")
            Return 1
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchProjectCorporationEx(ByVal ProjectNo As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectCorporationExDST")
        Dim dstTemp As DataSet

        Dim ParentCol(2) As DataColumn
        Dim ChildCol(2) As DataColumn

        dstResult = Me.FetchProjectCorporation(ProjectNo, "%", "%", "%")

        dstTemp = Me.FetchCorporationAccount(ProjectNo, "%", "%", "%", "%", "%")
        dstResult.Merge(dstTemp)

        dstTemp = Me.FetchCorporationBankSaving(ProjectNo, "%", "%")
        dstResult.Merge(dstTemp)

        dstTemp = Me.FetchCorporationBusiness(ProjectNo, "%", "%", "%")
        dstResult.Merge(dstTemp)

        dstTemp = Me.FetchCorporationExternalGuarantee(ProjectNo, "%", "%")
        dstResult.Merge(dstTemp)

        dstTemp = Me.FetchCorporationLoan(ProjectNo, "%", "%")
        dstResult.Merge(dstTemp)

        dstTemp = Me.FetchCorporationStockStructure(ProjectNo, "%", "%")
        dstResult.Merge(dstTemp)

        '设置主表(项目企业信息)的关联字段集
        ParentCol(0) = dstResult.Tables("TProjectCorporation").Columns("project_code")
        ParentCol(1) = dstResult.Tables("TProjectCorporation").Columns("corporation_code")
        ParentCol(2) = dstResult.Tables("TProjectCorporation").Columns("phase")

        '设置从表(企业财务报表)的关联字段集
        ChildCol(0) = dstResult.Tables("TCorporationAccount").Columns("project_code")
        ChildCol(1) = dstResult.Tables("TCorporationAccount").Columns("corporation_code")
        ChildCol(2) = dstResult.Tables("TCorporationAccount").Columns("phase")

        '设置数据集合中“项目企业信息”表与“企业财务报表”表的关联关系
        dstResult.Relations.Add("ProjectCorporation_CorporationAccount", ParentCol, ChildCol)

        '设置从表(企业借款记录)的关联字段集
        ChildCol(0) = dstResult.Tables("TCorporationLoan").Columns("project_code")
        ChildCol(1) = dstResult.Tables("TCorporationLoan").Columns("corporation_code")
        ChildCol(2) = dstResult.Tables("TCorporationLoan").Columns("phase")

        '设置数据集合中“项目企业信息”表与“企业借款记录”表的关联关系
        dstResult.Relations.Add("ProjectCorporation_CorporationLoan", ParentCol, ChildCol)

        '设置从表(企业经营情况)的关联字段集
        ChildCol(0) = dstResult.Tables("TCorporationBusiness").Columns("project_code")
        ChildCol(1) = dstResult.Tables("TCorporationBusiness").Columns("corporation_code")
        ChildCol(2) = dstResult.Tables("TCorporationBusiness").Columns("phase")

        '设置数据集合中“项目企业信息”表与“企业经营情况”表的关联关系
        dstResult.Relations.Add("ProjectCorporation_CorporationBusiness", ParentCol, ChildCol)

        '设置从表(企业银行存款记录)的关联字段集
        ChildCol(0) = dstResult.Tables("TCorporationBankSaving").Columns("project_code")
        ChildCol(1) = dstResult.Tables("TCorporationBankSaving").Columns("corporation_code")
        ChildCol(2) = dstResult.Tables("TCorporationBankSaving").Columns("phase")

        '设置数据集合中“项目企业信息”表与“企业银行存款记录”表的关联关系
        dstResult.Relations.Add("ProjectCorporation_CorporationBankSaving", ParentCol, ChildCol)

        '设置从表(对外担保记录)的关联字段集
        ChildCol(0) = dstResult.Tables("TCorporationExternalGuarantee").Columns("project_code")
        ChildCol(1) = dstResult.Tables("TCorporationExternalGuarantee").Columns("corporation_code")
        ChildCol(2) = dstResult.Tables("TCorporationExternalGuarantee").Columns("phase")

        '设置数据集合中“项目企业信息”表与“对外担保记录”表的关联关系
        dstResult.Relations.Add("ProjectCorporation_CorporationExternalGuarantee", ParentCol, ChildCol)

        '设置从表(企业股权结构)的关联字段集
        ChildCol(0) = dstResult.Tables("TCorporationStockStructure").Columns("project_code")
        ChildCol(1) = dstResult.Tables("TCorporationStockStructure").Columns("corporation_code")
        ChildCol(2) = dstResult.Tables("TCorporationStockStructure").Columns("phase")

        '设置数据集合中“项目企业信息”表与“企业股权结构”表的关联关系
        dstResult.Relations.Add("ProjectCorporation_CorporationStockStructure", ParentCol, ChildCol)

        Return dstResult
    End Function
End Class