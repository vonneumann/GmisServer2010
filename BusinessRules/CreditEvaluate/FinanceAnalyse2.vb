Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

'新版的财务分析中间层类
Public Class FinanceAnalyse
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

    '删除 项目财务分析记录
    Public Function DeleteProjectFinanceAnalyse(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String) As Integer
        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            Dim command As System.Data.SqlClient.SqlCommand = moConnection.CreateCommand()

            Dim parameter As SqlParameter

            command.CommandText = "PDeleteProjectFinanceAnalyse"
            command.CommandType = CommandType.StoredProcedure
            command.Transaction = moTransaction

            parameter = command.Parameters.Add("@ProjectNo", SqlDbType.VarChar, 25)
            parameter.Value = ProjectNo

            parameter = command.Parameters.Add("@CorporationNo", SqlDbType.VarChar, 25)
            parameter.Value = CorporationNo

            parameter = command.Parameters.Add("@Phase", SqlDbType.VarChar, 25)
            parameter.Value = Phase

            'parameter = command.Parameters.Add("@Month", SqlDbType.VarChar, 25)
            'parameter.Value = Month

            'parameter = command.Parameters.Add("@MonthLast", SqlDbType.VarChar, 25)
            'parameter.Value = MonthLast

            Return command.ExecuteNonQuery()
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function


    '查询 项目财务分析记录
    Public Function FetchProjectFinanceAnalyse(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("ProjectFinanceAnalyseDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchProjectFinanceAnalyse", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = Condition

            da.Fill(dstResult, "TProjectFinanceAnalyse")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    '查询 项目财务分析记录
    Public Function FetchProjectFinanceAnalyse(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As DataSet
        Return Me.FetchProjectFinanceAnalyse( _
         "{dbo.Project_Finance_Analyse.project_code LIKE '" + ProjectNo + "' AND " + _
         " dbo.Project_Finance_Analyse.corporation_code LIKE '" + CorporationNo + "' AND " + _
         " dbo.Project_Finance_Analyse.phase LIKE '" + Phase + "' AND " + _
         " dbo.Project_Finance_Analyse.month LIKE '" + Month + "' AND " + _
         " dbo.Project_Finance_Analyse.month_last LIKE '" + MonthLast + "'}")
    End Function

    '创建 项目财务分析记录
    Public Function CreateProjectFinanceAnalyse(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As Boolean
        Try
            '定义创建项目财务分析的命令对象
            Dim createCommand As SqlCommand = New SqlCommand("dbo.PCreateProjectFinanceAnalyse", moConnection, moTransaction)
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

            '执行创建项目定量评分的命令
            If createCommand.ExecuteNonQuery() = 0 Then
                Throw New System.Exception("创建项目财务分析表失败。")
            End If

            '创建获取项目定量评分的命令
            Dim da As SqlDataAdapter = New SqlDataAdapter()
            Dim dsFinanceAnalyse As DataSet = New DataSet("ProjectFinanceAnalyseDST")

            da.SelectCommand = New SqlCommand("dbo.PFetchProjectFinanceAnalyse", moConnection, moTransaction)
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000)
            da.SelectCommand.Parameters("@Condition").Value = _
             "{dbo.Project_Finance_Analyse.project_code LIKE '" + ProjectNo + "' AND " + _
             " dbo.Project_Finance_Analyse.corporation_code LIKE '" + CorporationNo + "' AND " + _
             " dbo.Project_Finance_Analyse.phase LIKE '" + Phase + "' AND " + _
             " dbo.Project_Finance_Analyse.month LIKE '" + Month + "' AND " + _
             " dbo.Project_Finance_Analyse.month_last LIKE '" + MonthLast + "'}"

            '获得新创建的项目财务分析的记录集
            da.Fill(dsFinanceAnalyse, "TProjectFinanceAnalyse")

            If dsFinanceAnalyse.Tables("TProjectFinanceAnalyse").Rows.Count > 0 Then
                Dim dsCorporationAccount As DataSet = New DataSet("CorporationAccountDST")

                '得到“企业帐户”数据集
                da.SelectCommand.CommandText = "dbo.PFetchCorporationAccount"
                da.SelectCommand.Parameters("@Condition").Value = _
                 "{dbo.corporation_account.project_code LIKE '" + ProjectNo + "' AND " + _
                 " dbo.corporation_account.corporation_code LIKE '" + CorporationNo + "' AND " + _
                 " dbo.corporation_account.phase LIKE '" + Phase + "' AND " + _
                 " (dbo.corporation_account.month LIKE '" + Month + "' OR " + _
                 " dbo.corporation_account.month LIKE '" + MonthLast + "')}"

                da.Fill(dsCorporationAccount, "TCorporationAccount")

                Dim row As DataRow

                '循环更新“index_value”字段的值
                For Each row In dsFinanceAnalyse.Tables("TProjectFinanceAnalyse").Rows
                    row("index_value") = Me.GetIndexValue(ProjectNo, CorporationNo, Phase, Month, MonthLast, row("index_type"), row("index_id"), dsCorporationAccount.Tables("TCorporationAccount"))
                Next

                da.SelectCommand.CommandText = "dbo.PFetchProjectFinanceAnalyse"
                da.SelectCommand.Parameters("@Condition").Value = "NULL"

                da.UpdateCommand = New SqlCommand("dbo.PUpdateProjectFinanceAnalyse", moConnection, moTransaction)
                da.UpdateCommand.CommandType = CommandType.StoredProcedure
                da.UpdateCommand.Parameters.Add("@ProjectNo", SqlDbType.Char, 9, "project_code")
                da.UpdateCommand.Parameters.Add("@CorporationNo", SqlDbType.Char, 5, "corporation_code")
                da.UpdateCommand.Parameters.Add("@Phase", SqlDbType.NVarChar, 4, "phase")
                da.UpdateCommand.Parameters.Add("@Month", SqlDbType.VarChar, 6, "month")
                da.UpdateCommand.Parameters.Add("@MonthLast", SqlDbType.VarChar, 6, "month_last")
                da.UpdateCommand.Parameters.Add("@IndexType", SqlDbType.Char, 2, "index_type")
                da.UpdateCommand.Parameters.Add("@IndexID", SqlDbType.Char, 3, "index_id")
                da.UpdateCommand.Parameters.Add("@IndexValue", SqlDbType.Decimal, 9, "index_value")

                '更新项目财务分析记录集
                da.Update(dsFinanceAnalyse, "TProjectFinanceAnalyse")

                Return True
            End If
        Catch ex As System.Exception
            Throw ex
            Return False
        End Try

        Return False
    End Function

    '获取在企业财务表中指定的项目编号、企业编号、项目阶段、年月、财务科目的指标值。
    Private Function GetCorporationAccountValue(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal ItemType As String, ByVal ItemCode As String, ByRef dtCorporationAccount As DataTable) As Decimal
        Dim foundRows() As DataRow

        Try
            foundRows = dtCorporationAccount.Select("project_code = '" + ProjectNo + "' AND corporation_code = '" + CorporationNo + "' AND phase = '" + Phase + "' AND month = '" + Month + "' AND item_type = '" + ItemType + "' AND item_code = '" + ItemCode + "'")

            If foundRows.Length > 0 Then
                If foundRows(0).IsNull("value") Then
                    Return 0
                Else
                    Return CType(foundRows(0)("value"), Decimal)
                End If
            End If
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    '获取指定项目编号、企业编号、项目阶段、年月所对应会计科目的指标值。
    Public Function GetIndexValue(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String, ByVal IndexType As String, ByVal IndexID As String) As Object
        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            Dim da As SqlDataAdapter = New SqlDataAdapter()
            Dim dsCorporationAccount As DataSet = New DataSet("CorporationAccountDST")

            da.SelectCommand.Connection = moConnection
            da.SelectCommand.Transaction = moTransaction
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

    '在指定企业财务表中，获取指定财务科目的指标值。
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
                Case "11003"     '净资产：a34-a49-a55-a32 公式改为：a34-A57-a32 qxd 2004-7-27 担保中心秦工提出。
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                        Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a57", dtCorporationAccount) - _
                        Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount)
                    'Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                    'Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                    'Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount)
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

    Public Function FetchFinanceAnalyseIndex(ByVal Condition As String) As DataSet
        Dim dstResult As DataSet = New DataSet("FinanceAnalyseIndexDST")
        Dim da As SqlDataAdapter = New SqlDataAdapter()

        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            da.SelectCommand = New SqlCommand("dbo.PFetchFinanceAnalyseIndex", moConnection)
            da.SelectCommand.Transaction = moTransaction
            da.SelectCommand.CommandType = CommandType.StoredProcedure
            da.SelectCommand.Parameters.Add("@Condition", SqlDbType.NVarChar, 1000, "Condition")
            da.SelectCommand.Parameters("@Condition").Value = IIf(Condition Is Nothing, DBNull.Value, Condition)

            da.Fill(dstResult, "TFinanceAnalyseIndex")
        Catch ex As System.Exception
            Throw ex
        End Try

        Return dstResult
    End Function

    Public Function FetchFinanceAnalyseIndex(ByVal IndexType As String, ByVal IndexID As String) As DataSet
        Return Me.FetchFinanceAnalyseIndex( _
         "{dbo.Finance_Analyse_Index.index_type LIKE '" + IndexType + "' AND " + _
         " dbo.Finance_Analyse_Index.index_id LIKE '" + IndexID + "'}")
    End Function
End Class
