Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class Credit
    Private moConnection As SqlConnection
    Private moTransaction As SqlTransaction
	Private moCreditAppraiseSystemCommand As SqlCommand
    Private _creditIndexTypeCommand As SqlCommand

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

    Public Function DuplicateCreditAppraise(ByVal sourceID As Integer) As Integer
        Me.DuplicateCreditAppraise(sourceID, 0)
    End Function

    Public Function DuplicateCreditAppraise(ByVal sourceID As Integer, ByVal destinationID As Integer) As Integer
        Try
            Dim SystemID As Int32

            Dim createCommand As SqlCommand = New SqlCommand("dbo.PDuplicateCreditAppraise", moConnection, moTransaction)
            createCommand.CommandType = CommandType.StoredProcedure

            createCommand.Parameters.Add("@Result", SqlDbType.Int)
            createCommand.Parameters.Add("@SourceID", SqlDbType.Int)
            createCommand.Parameters.Add("@DestinationID", SqlDbType.Int)

            createCommand.Parameters("@Result").Direction = ParameterDirection.ReturnValue
            createCommand.Parameters("@SourceID").Value = sourceID
            createCommand.Parameters("@DestinationID").Value = destinationID

            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            If createCommand.ExecuteNonQuery() = 0 Then
                Throw New System.Exception("复制资信评分体系出错。")
            Else
                Return createCommand.Parameters("@Result").Value
            End If
        Catch ex As System.Exception
            Throw ex
            Return 0
        End Try

        Return 0
    End Function

    Public Function FetchCreditIndexType(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("CreditIndexTypeDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchCreditIndexType", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = Condition

            _creditIndexTypeCommand = da.SelectCommand

            da.Fill(dstResult, "TCreditIndexType")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateCreditIndexType(ByVal dsCommit As DataSet) As Boolean
        Dim da As SqlDataAdapter = New SqlDataAdapter()

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

        If _creditIndexTypeCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCreditIndexType", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            _creditIndexTypeCommand.Connection = moConnection
            da.SelectCommand = _creditIndexTypeCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dsCommit, "TCreditIndexType")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function FetchCreditAppraiseSystem(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("CreditAppraiseSystemDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchCreditAppraiseSystem", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = Condition

            moCreditAppraiseSystemCommand = da.SelectCommand

            da.Fill(dstResult, "TCreditAppraiseSystem")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function UpdateCreditAppraiseSystem(ByVal dsCommit As DataSet) As Boolean
        Dim da As SqlDataAdapter = New SqlDataAdapter()

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

        If moCreditAppraiseSystemCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCreditAppraiseSystem", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCreditAppraiseSystemCommand.Connection = moConnection
            da.SelectCommand = moCreditAppraiseSystemCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dsCommit, "TCreditAppraiseSystem")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    '查询 项目总量分析记录
    Public Function FetchProjectCredit(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectCreditDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchProjectCredit", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = Condition

            da.Fill(dstResult, "TProjectCredit")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    '查询 项目总量分析记录
    Public Function FetchProjectCredit(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As DataSet
        Dim SystemID As Int32 = Me.GetSystemID(ProjectNo, CorporationNo, Phase)

        Return Me.FetchProjectCredit( _
         "{dbo.project_credit_appraise.project_code LIKE '" + ProjectNo + "' AND " + _
         " dbo.project_credit_appraise.corporation_code LIKE '" + CorporationNo + "' AND " + _
         " dbo.project_credit_appraise.phase LIKE '" + Phase + "' AND " + _
         " dbo.project_credit_appraise.month LIKE '" + Month + "' AND " + _
         " dbo.project_credit_appraise.month_last LIKE '" + MonthLast + "' AND " + _
         " dbo.project_credit_appraise.system_id = " + SystemID.ToString() + "}")
    End Function

    '创建 项目总量分析记录
    Public Function CreateProjectCredit(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As Boolean
        Try
            Dim SystemID As Int32

            Dim createCommand As SqlCommand = New SqlCommand("dbo.PCreateProjectCredit", moConnection, moTransaction)
            createCommand.CommandType = CommandType.StoredProcedure

            createCommand.Parameters.Add("@Result", SqlDbType.Int)
            createCommand.Parameters.Add("@ProjectNo", SqlDbType.Char, 9)
            createCommand.Parameters.Add("@CorporationNo", SqlDbType.Char, 5)
            createCommand.Parameters.Add("@Phase", SqlDbType.NVarChar, 4)
            createCommand.Parameters.Add("@Month", SqlDbType.VarChar, 6)
            createCommand.Parameters.Add("@MonthLast", SqlDbType.VarChar, 6)

            createCommand.Parameters("@Result").Direction = ParameterDirection.ReturnValue
            createCommand.Parameters("@ProjectNo").Value = ProjectNo
            createCommand.Parameters("@CorporationNo").Value = CorporationNo
            createCommand.Parameters("@Phase").Value = Phase
            createCommand.Parameters("@Month").Value = Month
            createCommand.Parameters("@MonthLast").Value = MonthLast

            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            If createCommand.ExecuteNonQuery() = 0 Then
                Throw New System.Exception("创建项目总量分析表失败。")
            Else
                SystemID = createCommand.Parameters("@Result").Value
            End If
        Catch ex As System.Exception
            Throw ex
            Return False
        End Try

        Return True
    End Function
End Class
