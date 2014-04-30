Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class QuantityEvaluation
    Private moConnection As SqlConnection
    Private moTransaction As SqlTransaction
	Private moCreditQuantityIndexCommand As SqlCommand
	Private moCreditQuantityStandardCommand As SqlCommand

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

    Public Function GetItemValue(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal ItemType As String, ByVal ItemNo As String) As Decimal
        Dim dr As SqlDataReader

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            Dim fetchCommand As SqlCommand = New SqlCommand("dbo.PFetchCorporationAccount", moConnection)
            fetchCommand.CommandType = CommandType.StoredProcedure
            fetchCommand.Transaction = moTransaction
            fetchCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000)

            fetchCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_account.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_account.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_account.phase LIKE '" + Phase + "' AND " + _
               " dbo.corporation_account.month LIKE '" + Month + "' AND " + _
               " dbo.corporation_account.item_type LIKE '" + ItemType + "' AND " + _
               " dbo.corporation_account.item_code LIKE '" + ItemNo + "'}"

            dr = fetchCommand.ExecuteReader()

            If dr.Read() Then
                Dim Index As Int32 = dr.GetOrdinal("value")

                If Not dr.IsDBNull(Index) Then
                    Return dr.GetDecimal(Index)
                Else
                    Return 0
                End If
            Else
                Throw New System.Exception("无法读取数据！")
            End If
        Catch ex As System.Exception
            Throw ex
        Finally
            If Not dr Is Nothing Then
                dr.Close()
            End If
        End Try

        Return 0
    End Function

    '删除 项目定量分析记录
    Public Function DeleteProjectCreditQuantity(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As Integer
        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            Dim command As System.Data.SqlClient.SqlCommand = moConnection.CreateCommand()

            Dim parameter As SqlParameter

            command.CommandText = "PDeleteProjectCreditQuantity"
            command.CommandType = CommandType.StoredProcedure
            command.Transaction = moTransaction

            parameter = command.Parameters.Add("@ProjectNo", SqlDbType.VarChar, 25)
            parameter.Value = ProjectNo

            parameter = command.Parameters.Add("@CorporationNo", SqlDbType.VarChar, 25)
            parameter.Value = CorporationNo

            parameter = command.Parameters.Add("@Phase", SqlDbType.VarChar, 25)
            parameter.Value = Phase

            parameter = command.Parameters.Add("@Month", SqlDbType.VarChar, 25)
            parameter.Value = Month

            parameter = command.Parameters.Add("@MonthLast", SqlDbType.VarChar, 25)
            parameter.Value = MonthLast

            Return command.ExecuteNonQuery()
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    '删除 项目定量分析记录
    Public Function DeleteProjectCreditQuantity(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String) As Integer
        Dim monthLast As String
        Dim day As DateTime

        Try
            day = DateTime.ParseExact(Month, "yyyyMM", System.Globalization.DateTimeFormatInfo.CurrentInfo)
        Catch
            Throw
        End Try

        monthLast = day.AddYears(-1).ToString("yyyy12")

        Me.DeleteProjectCreditQuantity(ProjectNo, CorporationNo, Phase, Month, monthLast)
    End Function

    '查询 项目定量分析记录
    Public Function FetchProjectCreditQuantity(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectCreditQuantityDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchProjectCreditQuantity", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = Condition

            da.Fill(dstResult, "TProjectCreditQuantity")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    '查询 项目定量分析记录
    Public Function FetchProjectCreditQuantity(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As DataSet
        Dim SystemID As Int32 = Me.GetSystemID(ProjectNo, CorporationNo, Phase)

        Return Me.FetchProjectCreditQuantity( _
         "{dbo.project_credit_quantity_score.project_code LIKE '" + ProjectNo + "' AND " + _
         " dbo.project_credit_quantity_score.corporation_code LIKE '" + CorporationNo + "' AND " + _
         " dbo.project_credit_quantity_score.phase LIKE '" + Phase + "' AND " + _
         " dbo.project_credit_quantity_score.month LIKE '" + Month + "' AND " + _
         " dbo.project_credit_quantity_score.month_last LIKE '" + MonthLast + "' AND " + _
         " dbo.project_credit_quantity_score.system_id = " + SystemID.ToString() + "}")
    End Function

    '创建 项目定量分析记录
    Public Function CreateProjectCreditQuantity(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As Boolean
        Return Me.CreateProjectCreditQuantity(ProjectNo, CorporationNo, Phase, Month, MonthLast, Nothing)
    End Function

    '创建 项目定量分析记录
    Public Function CreateProjectCreditQuantity(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String, ByVal SystemID As Object) As Boolean
        Try
            'Dim SystemID As Int32

            Dim createCommand As SqlCommand = New SqlCommand("dbo.PCreateProjectCreditQuantity", moConnection, moTransaction)
            createCommand.CommandType = CommandType.StoredProcedure

            createCommand.Parameters.Add("@Result", SqlDbType.Int)
            createCommand.Parameters.Add("@ProjectNo", SqlDbType.Char, 9)
            createCommand.Parameters.Add("@CorporationNo", SqlDbType.Char, 5)
            createCommand.Parameters.Add("@Phase", SqlDbType.NVarChar, 4)
            createCommand.Parameters.Add("@Month", SqlDbType.VarChar, 6)
            createCommand.Parameters.Add("@MonthLast", SqlDbType.VarChar, 6)
            createCommand.Parameters.Add("@SystemID", SqlDbType.Int)

            createCommand.Parameters("@Result").Direction = ParameterDirection.ReturnValue
            createCommand.Parameters("@ProjectNo").Value = ProjectNo
            createCommand.Parameters("@CorporationNo").Value = CorporationNo
            createCommand.Parameters("@Phase").Value = Phase
            createCommand.Parameters("@Month").Value = Month
            createCommand.Parameters("@MonthLast").Value = MonthLast
            If SystemID Is Nothing Then
                createCommand.Parameters("@SystemID").Value = DBNull.Value
            Else
                createCommand.Parameters("@SystemID").Value = SystemID
            End If

            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            If createCommand.ExecuteNonQuery() = 0 Then
                Throw New System.Exception("创建项目定量分析表失败。")
            Else
                SystemID = createCommand.Parameters("@Result").Value
            End If

            Dim da As SqlDataAdapter = New SqlDataAdapter()
            Dim dsCreditQuantity As DataSet = New DataSet("ProjectCreditQuantityDST")

            da.SelectCommand = New SqlCommand("dbo.PFetchProjectCreditQuantity", moConnection, moTransaction)
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000)
            da.SelectCommand.Parameters("@Condition").Value = _
             "{dbo.project_credit_quantity_score.project_code LIKE '" + ProjectNo + "' AND " + _
             " dbo.project_credit_quantity_score.corporation_code LIKE '" + CorporationNo + "' AND " + _
             " dbo.project_credit_quantity_score.phase LIKE '" + Phase + "' AND " + _
             " dbo.project_credit_quantity_score.month LIKE '" + Month + "' AND " + _
             " dbo.project_credit_quantity_score.month_last LIKE '" + MonthLast + "' AND " + _
             " dbo.project_credit_quantity_score.system_id = " + SystemID.ToString + "}"

            da.Fill(dsCreditQuantity, "TProjectCreditQuantity")

            If dsCreditQuantity.Tables("TProjectCreditQuantity").Rows.Count > 0 Then
                Dim dsCorporationAccount As DataSet = New DataSet("CorporationAccountDST")
                Dim dsCreditQuantityStandard As DataSet = New DataSet("CreditQuantityStandardDST")

                '得到“企业帐户”数据集
                da.SelectCommand.CommandText = "dbo.PFetchCorporationAccount"
                da.SelectCommand.Parameters("@Condition").Value = _
                 "{dbo.corporation_account.project_code LIKE '" + ProjectNo + "' AND " + _
                 " dbo.corporation_account.corporation_code LIKE '" + CorporationNo + "' AND " + _
                 " dbo.corporation_account.phase LIKE '" + Phase + "' AND " + _
                 " (dbo.corporation_account.month LIKE '" + Month + "' OR " + _
                 " dbo.corporation_account.month LIKE '" + MonthLast + "')}"

                da.Fill(dsCorporationAccount, "TCorporationAccount")

                '得到“定量得分判断标准”数据集
                da.SelectCommand.CommandText = "dbo.PFetchCreditQuantityStandard"
                da.SelectCommand.Parameters("@Condition").Value = _
                   "{dbo.credit_appraise_quantity_standard.system_id = " + SystemID.ToString() + "}"
                da.Fill(dsCreditQuantityStandard, "TCreditQuantityStandard")

                Dim row As DataRow

                '循环更新“index_value”字段的值
                For Each row In dsCreditQuantity.Tables("TProjectCreditQuantity").Rows
                    Try
                        row("index_value") = Me.GetIndexValue(ProjectNo, CorporationNo, Phase, Month, MonthLast, row("index_type"), row("index_id"), dsCorporationAccount.Tables("TCorporationAccount"))

                        If Not row.IsNull("index_value") Then
                            row("score") = Me.GetIndexScore(SystemID, row("index_type"), row("index_id"), row("index_value"), row("compare"), dsCreditQuantityStandard.Tables("TCreditQuantityStandard"))
                        Else
                            If row("index_type") = "11" And row("index_id") = "010" Then
                                row("score") = 5
                                row("remark") = "无借款"
                            ElseIf row("index_type") = "13" And (row("index_id") = "001" Or row("index_id") = "002" Or row("index_id") = "003" Or row("index_id") = "004") Then
                                row("score") = 5
                                row("remark") = "无上年数据"
                            ElseIf row("index_type") = "12" And row("index_id") = "006" Then
                                row("score") = 10
                                row("remark") = "利息小于或等于0"
                            Else
                                row("score") = DBNull.Value
                            End If
                        End If

                        If (Not row("score") Is DBNull.Value) And (Not row("quotiety") Is DBNull.Value) Then
                            row("score_final") = row("score") * row("quotiety")
                        Else
                            row("score_final") = DBNull.Value
                        End If
                    Catch ex As System.Exception

                    End Try
                Next

                da.SelectCommand.CommandText = "dbo.PFetchProjectCreditQuantity"
                da.SelectCommand.Parameters("@Condition").Value = "NULL"

                da.UpdateCommand = New SqlCommand("dbo.PUpdateProjectCreditQuantity", moConnection, moTransaction)
                da.UpdateCommand.CommandType = CommandType.StoredProcedure
                da.UpdateCommand.Parameters.Add("@ProjectNo", SqlDbType.Char, 9, "project_code")
                da.UpdateCommand.Parameters.Add("@CorporationNo", SqlDbType.Char, 5, "corporation_code")
                da.UpdateCommand.Parameters.Add("@Phase", SqlDbType.NVarChar, 4, "phase")
                da.UpdateCommand.Parameters.Add("@Month", SqlDbType.VarChar, 6, "month")
                da.UpdateCommand.Parameters.Add("@MonthLast", SqlDbType.VarChar, 6, "month_last")
                da.UpdateCommand.Parameters.Add("@SystemID", SqlDbType.Int, 4, "system_id")
                da.UpdateCommand.Parameters.Add("@IndexType", SqlDbType.Char, 2, "index_type")
                da.UpdateCommand.Parameters.Add("@IndexID", SqlDbType.Char, 3, "index_id")
                da.UpdateCommand.Parameters.Add("@IndexValue", SqlDbType.Decimal, 9, "index_value")
                da.UpdateCommand.Parameters.Add("@Score", SqlDbType.Int, 4, "score")
                da.UpdateCommand.Parameters.Add("@ScoreFinal", SqlDbType.Decimal, 5, "score_final")
                da.UpdateCommand.Parameters.Add("@Remark", SqlDbType.NVarChar, 50, "remark")

                da.Update(dsCreditQuantity, "TProjectCreditQuantity")

                Return True
            End If
        Catch ex As System.Exception
            Throw ex
            Return False
        End Try

        Return False
    End Function

    Private Function GetCorporationAccountValue(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal ItemType As String, ByVal ItemCode As String, ByRef dtCorporationAccount As DataTable)
        Dim foundRows() As DataRow

        Try
            foundRows = dtCorporationAccount.Select("project_code = '" + ProjectNo + "' AND corporation_code = '" + CorporationNo + "' AND phase = '" + Phase + "' AND month = '" + Month + "' AND item_type = '" + ItemType + "' AND item_code = '" + ItemCode + "'")

            If foundRows.Length > 0 Then
                If foundRows(0).IsNull("value") Then
                    Return 0
                Else
                    Return CType(foundRows(0)("value"), Decimal)
                End If
            Else
                Return System.DBNull.Value
            End If
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function GetIndexValue(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String, ByVal IndexType As String, ByVal IndexID As String) As Object
        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            Dim da As SqlDataAdapter = New SqlDataAdapter
            Dim dsCorporationAccount As DataSet = New DataSet("CorporationAccountDST")

            da.SelectCommand.Connection = moConnection
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandText = "dbo.PFetchProjectCreditQuantity"
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.CommandText = "dbo.PFetchCorporationAccount"
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000)
            da.SelectCommand.Parameters("@Condition").Value = _
               "{dbo.corporation_account.project_code LIKE '" + ProjectNo + "' AND " + _
               " dbo.corporation_account.corporation_code LIKE '" + CorporationNo + "' AND " + _
               " dbo.corporation_account.phase LIKE '" + Phase + "' AND " + _
               " dbo.corporation_account.month LIKE '" + Month + "'}"
            da.Fill(dsCorporationAccount, "TCorporationAccount")

            Return Me.GetIndexValue(ProjectNo, CorporationNo, Phase, Month, MonthLast, IndexType, IndexID, dsCorporationAccount.Tables("TCorporationAccount"))
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function GetIndexValue(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String, ByVal IndexType As String, ByVal IndexID As String, ByVal dtCorporationAccount As DataTable) As Object
        If dtCorporationAccount Is Nothing Then
            Return System.DBNull.Value
        Else
            If dtCorporationAccount.Rows.Count <= 0 Then
                Return System.DBNull.Value
            End If
        End If

        Try
            Select Case IndexType + IndexID
                Case "11001"     '总资产：a34
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount)
                Case "11002"     '营运资金：a14-a49
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a14", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount)
                Case "11003"     '净资产：a34-a49-a55-a32
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount)
                Case "11004"     '资产负债率：a57/a34
                    Dim a34 As Decimal = Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount)

                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a57", dtCorporationAccount) / a34) * 100
                Case "11005"     '流动比率：a14/a49
                    Dim a49 As Decimal = Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount)

                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a14", dtCorporationAccount) / a49
                Case "11006"     '速动比率：(a14-a10)/a49
                    Dim a49 As Decimal = Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount)

                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a14", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a10", dtCorporationAccount)) / a49
                Case "11007"     '长期资产适宜率：(a55+a65)/(a28+a17)
                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) + _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a65", dtCorporationAccount)) / _
                     (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a28", dtCorporationAccount) + _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a17", dtCorporationAccount))
                Case "11008"     '齿轮比率：(a50+a35)/<003> = (a50+a35)/(a34-a49-a55-a32)
                    Return ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a50", dtCorporationAccount) + _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a35", dtCorporationAccount)) / _
                     (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount))) * 100
                Case "11009"     '或有负债比率：a91/<003> = a91/(a34-a49-a55-a32)
                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a91", dtCorporationAccount) / _
                    (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount))) * 100
                Case "11010"     '贷款按期偿还率：1-a92/a93
                    Dim a93 As Decimal = Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a93", dtCorporationAccount)

                    If a93 = 0 Then
                        Return DBNull.Value
                    Else
                        Return (1 - Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a92", dtCorporationAccount) / a93) * 100
                    End If
                Case "12001"     '年营业收入：b01
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount)
                Case "12002"     '销售利润率：b14/b01
                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b14", dtCorporationAccount) / _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount)) * 100
                Case "12003"     '应收帐款周转率：b01/((a06+a06<上年>)/2)
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount) / _
                     ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a06", dtCorporationAccount) + _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a06", dtCorporationAccount)) / 2)
                Case "12004"     '存货周转率：b02*2/(a10+a10<上年>)
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b02", dtCorporationAccount) * 2 / _
                     (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a10", dtCorporationAccount) + _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a10", dtCorporationAccount))
                Case "12005"     '净资产回报率：b18/((<03>+<03><上年>)/2) = b18/(((a34-a49-a55-a32)+(a34-a49-a55-a32)<上年>)/2)
                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b18", dtCorporationAccount) / _
                     (( _
                     ( _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount) _
                     ) + _
                     ( _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a34", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a49", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a55", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a32", dtCorporationAccount) _
                     )) / 2)) * 100
                Case "12006"     '利息保障倍数：(b14+b08)/b08
                    Dim b08 As Decimal = Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b08", dtCorporationAccount)

                    If b08 <= 0 Then
                        Return DBNull.Value
                    Else
                        Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b14", dtCorporationAccount) + _
                          Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b08", dtCorporationAccount)) / b08
                    End If
                Case "13001"     '净资产增长率：(<003>-<003><上年>)/ABS(<003><上年>) = ((a34-a49-a55-a32) - (a34-a49-a55-a32)<上年>) / ABS((a34-a49-a55-a32)<上年>)
                    Return (( _
                       ( _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount) _
                       ) - _
                       ( _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a34", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a49", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a55", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a32", dtCorporationAccount) _
                       ) _
                      ) / _
                       Math.Abs( _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a34", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a49", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a55", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a32", dtCorporationAccount) _
                       )) * 100
                Case "13002"     '销售收入增长率：(b01-b01<上年>)/ABS(b01<上年>)
                    Return ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b01", dtCorporationAccount)) / _
                      Math.Abs(Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b01", dtCorporationAccount))) * 100
                Case "13003"     '利润增长率：(b14-b14<上年>)/ABS(b14<上年>)
                    Return ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b14", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount)) / _
                      Math.Abs(Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount))) * 100
                Case "13004"     '利润增长额：b14-b14<上年>
                    Dim foundRows As DataRow() = dtCorporationAccount.Select("project_code = '" + ProjectNo + "' AND corporation_code = '" + CorporationNo + "' AND phase = '" + Phase + "' AND month = '" + MonthLast + "' AND item_type = '02' AND item_code = 'b14'")

                    If foundRows.Length = 0 Then
                        Return DBNull.Value
                    End If

                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b14", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount)
                Case "14001"     '经营产生的现金流入量/流动负债：c04/a49
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c04", dtCorporationAccount) / _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount)
                Case "14002"     '经营产生的现金流入量/负债总额：c04/a57
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c04", dtCorporationAccount) / _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a57", dtCorporationAccount)
                Case "14003"     '经营产生的现金流入量/主营业务收入：c04/b01
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c04", dtCorporationAccount) / _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount)
                Case "14004"     '经营产生的现金流量净额/流动负债：c10/a49
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c10", dtCorporationAccount) / _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount)
                Case "14005"     '经营产生的现金流量净额/负债总额：c10/a57
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c10", dtCorporationAccount) / _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a57", dtCorporationAccount)
                Case "14006"     '经营产生现金流量净额/主营业务收入：c10/b01
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c10", dtCorporationAccount) / _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount)
            End Select
        Catch ex As System.DivideByZeroException
            Return System.DBNull.Value
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function GetIndexScore(ByVal SystemID As Int32, ByVal IndexType As String, ByVal IndexID As String, ByVal Value As Decimal, ByVal Compare As Boolean, ByRef dtCreditQuantityStandard As DataTable) As Object
        Try
            Dim foundRows As DataRow()

            foundRows = dtCreditQuantityStandard.Select("system_id = " + SystemID.ToString() + " AND index_type = '" + IndexType + "' AND index_id = '" + IndexID + "'")

            Dim i As Int32

            For i = 0 To foundRows.Length - 1
                If Compare Then
                    If Not foundRows(i).IsNull("lower_bound") Then
                        If Not foundRows(i).IsNull("upper_bound") Then
                            If Value > foundRows(i)("lower_bound") And Value <= foundRows(i)("upper_bound") Then
                                Return foundRows(i)("score")
                            End If
                        Else
                            If Value > foundRows(i)("lower_bound") Then
                                Return foundRows(i)("score")
                            End If
                        End If
                    Else
                        If Not foundRows(i).IsNull("upper_bound") Then
                            If Value <= foundRows(i)("upper_bound") Then
                                Return foundRows(i)("score")
                            End If
                        Else
                            Return DBNull.Value
                        End If
                    End If
                Else
                    If Not foundRows(i).IsNull("lower_bound") Then
                        If Not foundRows(i).IsNull("upper_bound") Then
                            If Value >= foundRows(i)("lower_bound") And Value < foundRows(i)("upper_bound") Then
                                Return foundRows(i)("score")
                            End If
                        Else
                            If Value >= foundRows(i)("lower_bound") Then
                                Return foundRows(i)("score")
                            End If
                        End If
                    Else
                        If Not foundRows(i).IsNull("upper_bound") Then
                            If Value < foundRows(i)("upper_bound") Then
                                Return foundRows(i)("score")
                            End If
                        Else
                            Return DBNull.Value
                        End If
                    End If
                End If
            Next
        Catch ex As System.Exception
            Throw ex
        End Try

        Return 0
    End Function

    Public Function FetchCreditQuantityStandard(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("CreditQuantityStandardDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchCreditQuantityStandard", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = Condition

            moCreditQuantityStandardCommand = da.SelectCommand

            da.Fill(dstResult, "TCreditQuantityStandard")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function FetchCreditQuantityStandard(ByVal SystemID As Int32, ByVal IndexType As String, ByVal IndexID As String) As DataSet
        Return Me.FetchCreditQuantityStandard( _
         "{dbo.credit_appraise_quantity_standard.system_id = " + SystemID.ToString() + " AND " + _
         " dbo.credit_appraise_quantity_standard.index_type LIKE '" + IndexType + "' AND " + _
         " dbo.credit_appraise_quantity_standard.index_id LIKE '" + IndexID + "'}")
    End Function

    Public Function FetchCreditQuantityIndex(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("CreditQuantityIndexDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchCreditQuantityIndex", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = Condition

            moCreditQuantityIndexCommand = da.SelectCommand

            da.Fill(dstResult, "TCreditQuantityIndex")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function FetchCreditQuantityIndex(ByVal SystemID As Int32, ByVal IndexType As String, ByVal IndexID As String) As DataSet
        Return Me.FetchCreditQuantityIndex( _
         "{dbo.credit_appraise_quantity_index.system_id = " + SystemID.ToString() + " AND " + _
         " dbo.credit_appraise_quantity_index.index_type LIKE '" + IndexType + "' AND " + _
         " dbo.credit_appraise_quantity_index.index_id LIKE '" + IndexID + "'}")
    End Function

    Public Function UpdateCreditQuantityIndex(ByVal dsCommit As DataSet) As Boolean
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

        If moCreditQuantityIndexCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCreditQuantityIndex", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCreditQuantityIndexCommand.Connection = moConnection
            da.SelectCommand = moCreditQuantityIndexCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dsCommit, "TCreditQuantityIndex")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    Public Function UpdateCreditQuantityStandard(ByVal dsCommit As DataSet) As Boolean
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

        If moCreditQuantityStandardCommand Is Nothing Then
            da.SelectCommand = New SqlCommand("dbo.PFetchCreditQuantityStandard", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = "NULL"
        Else
            moCreditQuantityStandardCommand.Connection = moConnection
            da.SelectCommand = moCreditQuantityStandardCommand
        End If

        Dim cmb As SqlCommandBuilder = New SqlCommandBuilder(da)

        da.InsertCommand = cmb.GetInsertCommand()
        da.InsertCommand.Transaction = moTransaction
        da.UpdateCommand = cmb.GetUpdateCommand()
        da.UpdateCommand.Transaction = moTransaction
        da.DeleteCommand = cmb.GetDeleteCommand()
        da.DeleteCommand.Transaction = moTransaction

        Try
            Return da.Update(dsCommit, "TCreditQuantityStandard")
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function
End Class
