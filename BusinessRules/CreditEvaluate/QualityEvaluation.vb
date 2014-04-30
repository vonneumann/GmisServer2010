Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class QualityEvaluation
    Private moConnection As SqlConnection
    Private moTransaction As SqlTransaction
	Private moCreditQualityIndexCommand As SqlCommand
	Private moCreditQualityStandardCommand As SqlCommand

    Public Sub New(ByRef Connection As SqlConnection)
        moConnection = Connection
        moTransaction = Nothing
    End Sub

    Public Sub New(ByRef Connection As SqlConnection, ByRef Transaction As SqlTransaction)
        moConnection = Connection
        moTransaction = Transaction
    End Sub

    Public Function GetSystemID() As Int32
        Dim dr As SqlDataReader

        Try
            Dim fetchCommand As SqlCommand = New SqlCommand()

            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            fetchCommand.CommandText = "SELECT TOP 1 system_id FROM dbo.credit_appraise_system WHERE used_flag = 1"
            fetchCommand.CommandType = CommandType.Text
            fetchCommand.Connection = moConnection
            fetchCommand.Transaction = moTransaction
            dr = fetchCommand.ExecuteReader()

            If dr.Read() Then
                Return dr.GetInt32(0)
            End If

            Return -99
        Catch
            Throw
        Finally
            If Not dr Is Nothing Then
                dr.Close()
            End If
        End Try
    End Function

    Public Function GetSystemID(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String) As Int32
        Try
            Dim fetchCommand As SqlCommand = New SqlCommand()

            fetchCommand.CommandText = "PGetCreditAppraiseSystemID"
            fetchCommand.CommandType = CommandType.StoredProcedure
            fetchCommand.Connection = moConnection
            fetchCommand.Transaction = moTransaction

            fetchCommand.Parameters.Add("@Result", SqlDbType.Int)
            fetchCommand.Parameters.Add("@ProjectNo", SqlDbType.VarChar, 25)
            fetchCommand.Parameters.Add("@CorporationNo", SqlDbType.VarChar, 25)
            fetchCommand.Parameters.Add("@Phase", SqlDbType.VarChar, 25)

            fetchCommand.Parameters("@Result").Direction = ParameterDirection.ReturnValue
            fetchCommand.Parameters("@ProjectNo").Value = ProjectNo
            fetchCommand.Parameters("@CorporationNo").Value = CorporationNo
            fetchCommand.Parameters("@Phase").Value = Phase

            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            If fetchCommand.ExecuteNonQuery() = 0 Then
                Throw New System.Exception("获取资信评分体系编号出错。")
            Else
                Return fetchCommand.Parameters("@Result").Value
            End If
        Catch ex As System.Exception
            Throw ex
            Return 0
        End Try

        Return 0
    End Function

    '查询 项目定性分析记录
    Public Function FetchProjectCreditQuality(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectCreditQualityDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchProjectCreditQuality", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = Condition

            da.Fill(dstResult, "TProjectCreditQuality")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    '查询 项目定性分析记录
    Public Function FetchProjectCreditQuality(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String) As DataSet
        Dim SystemID As Int32 = Me.GetSystemID(ProjectNo, CorporationNo, Phase)

        Return Me.FetchProjectCreditQuality( _
           "{dbo.project_credit_quality_score.project_code LIKE '" + ProjectNo + "' AND " + _
           " dbo.project_credit_quality_score.corporation_code LIKE '" + CorporationNo + "' AND " + _
           " dbo.project_credit_quality_score.phase LIKE '" + Phase + "' AND " + _
           " dbo.project_credit_quality_score.system_id = " + SystemID.ToString() + "}")
    End Function

    '创建 项目定性分析记录
    Public Function CreateProjectCreditQuality(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String) As Boolean
        Try
            Dim createCommand As SqlCommand = New SqlCommand("dbo.PCreateProjectCreditQuality", moConnection)
            createCommand.Transaction = moTransaction
            createCommand.CommandType = CommandType.StoredProcedure

            createCommand.Parameters.Add("@ProjectNo", SqlDbType.Char, 9)
            createCommand.Parameters.Add("@CorporationNo", SqlDbType.Char, 5)
            createCommand.Parameters.Add("@Phase", SqlDbType.NVarChar, 4)
            createCommand.Parameters.Add("@Month", SqlDbType.VarChar, 6)

            createCommand.Parameters("@ProjectNo").Value = ProjectNo
            createCommand.Parameters("@CorporationNo").Value = CorporationNo
            createCommand.Parameters("@Phase").Value = Phase
            createCommand.Parameters("@Month").Value = Month

            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            If createCommand.ExecuteNonQuery() = 0 Then
                Throw New System.Exception("创建项目定量分析表失败。")
            End If
        Catch ex As SystemException
            Throw ex
        End Try
    End Function

    Public Function UpdateProjectCreditQuality(ByVal dsCommit As DataSet) As Boolean
        Dim da As SqlDataAdapter = New SqlDataAdapter

        If dsCommit Is Nothing Then
            Return -1
        End If
        If Not dsCommit.HasChanges() Then
            Return 0
        End If

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.UpdateCommand = New SqlCommand("dbo.PUpdateProjectCreditQuality", moConnection)
            da.UpdateCommand.Transaction = moTransaction
            da.UpdateCommand.CommandType = CommandType.StoredProcedure
            da.UpdateCommand.Parameters.Add("@ProjectNo", SqlDbType.Char, 9, "project_code")
            da.UpdateCommand.Parameters.Add("@CorporationNo", SqlDbType.Char, 5, "corporation_code")
            da.UpdateCommand.Parameters.Add("@Phase", SqlDbType.NVarChar, 4, "phase")
            da.UpdateCommand.Parameters.Add("@SystemID", SqlDbType.Int, 4, "system_id")
            da.UpdateCommand.Parameters.Add("@IndexType", SqlDbType.Char, 2, "index_type")
            da.UpdateCommand.Parameters.Add("@IndexID", SqlDbType.Char, 3, "index_id")
            da.UpdateCommand.Parameters.Add("@IndexValue", SqlDbType.NVarChar, 10, "index_value")
            da.UpdateCommand.Parameters.Add("@Score", SqlDbType.Decimal, 9, "score")
            da.UpdateCommand.Parameters.Add("@Remark", SqlDbType.NVarChar, 50, "remark")

            Return da.Update(dsCommit, "TProjectCreditQuality")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchCreditQualityStandard(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("CreditQualityStandardDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchCreditQualityStandard", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = Condition

            moCreditQualityStandardCommand = da.SelectCommand

            da.Fill(dstResult, "TCreditQualityStandard")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    'Public Function FetchCreditQualityStandard(ByVal IndexType As String, ByVal IndexID As String) As DataSet
    '    Dim SystemID As Int32 = Me.GetSystemID()

    '    Return Me.FetchCreditQualityStandard(SystemID, IndexType, IndexID)
    'End Function

    Public Function FetchCreditQualityStandard(ByVal SystemID As Int32, ByVal IndexType As String, ByVal IndexID As String) As DataSet
        Return Me.FetchCreditQualityStandard( _
         "{dbo.credit_appraise_quality_standard.system_id = " + SystemID.ToString() + " AND " + _
         " dbo.credit_appraise_quality_standard.index_type LIKE '" + IndexType + "' AND " + _
         " dbo.credit_appraise_quality_standard.index_id LIKE '" + IndexID + "'}")
    End Function

    Public Function FetchCreditQualityIndex(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("CreditQualityIndexDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchCreditQualityIndex", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = Condition

            moCreditQualityIndexCommand = da.SelectCommand

            da.Fill(dstResult, "TCreditQualityIndex")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    'Public Function FetchCreditQualityIndex(ByVal IndexType As String, ByVal IndexID As String) As DataSet
    '    Dim SystemID As Int32 = Me.GetSystemID()

    '    Return Me.FetchCreditQualityIndex(SystemID, IndexType, IndexID)
    'End Function

    Public Function FetchCreditQualityIndex(ByVal SystemID As Int32, ByVal IndexType As String, ByVal IndexID As String) As DataSet
        Return Me.FetchCreditQualityIndex( _
         "{dbo.credit_appraise_quality_index.system_id = " + SystemID.ToString() + " AND " + _
         " dbo.credit_appraise_quality_index.index_type LIKE '" + IndexType + "' AND " + _
         " dbo.credit_appraise_quality_index.index_id LIKE '" + IndexID + "'}")
    End Function

    Public Function UpdateCreditQualityIndex(ByVal dsCommit As DataSet) As Boolean
        Dim da As SqlDataAdapter = New SqlDataAdapter

        If dsCommit Is Nothing Then
            Return -1
        End If
        If Not dsCommit.HasChanges() Then
            Return 0
        End If

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        If moCreditQualityIndexCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCreditQualityIndex", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCreditQualityIndexCommand.Connection = moConnection
            da.SelectCommand = moCreditQualityIndexCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dsCommit, "TCreditQualityIndex")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function UpdateCreditQualityStandard(ByVal dsCommit As DataSet) As Boolean
        Dim da As SqlDataAdapter = New SqlDataAdapter

        If dsCommit Is Nothing Then
            Return -1
        End If
        If Not dsCommit.HasChanges() Then
            Return 0
        End If

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If
        Catch ex As System.Exception
            Throw ex
        End Try

        If moCreditQualityStandardCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCreditQualityStandard", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCreditQualityStandardCommand.Connection = moConnection
            da.SelectCommand = moCreditQualityStandardCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dsCommit, "TCreditQualityStandard")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
