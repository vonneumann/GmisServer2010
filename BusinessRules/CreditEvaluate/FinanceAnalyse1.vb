Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient

'�°�Ĳ�������м����
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

    'ɾ�� ��Ŀ���������¼
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


    '��ѯ ��Ŀ���������¼
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

    '��ѯ ��Ŀ���������¼
    Public Function FetchProjectFinanceAnalyse(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As DataSet
        Return Me.FetchProjectFinanceAnalyse( _
         "{dbo.Project_Finance_Analyse.project_code LIKE '" + ProjectNo + "' AND " + _
         " dbo.Project_Finance_Analyse.corporation_code LIKE '" + CorporationNo + "' AND " + _
         " dbo.Project_Finance_Analyse.phase LIKE '" + Phase + "' AND " + _
         " dbo.Project_Finance_Analyse.month LIKE '" + Month + "' AND " + _
         " dbo.Project_Finance_Analyse.month_last LIKE '" + MonthLast + "'}")
    End Function

    '���� ��Ŀ���������¼
    Public Function CreateProjectFinanceAnalyse(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As Boolean
        Try
            '���崴����Ŀ����������������
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

            'ִ�д�����Ŀ�������ֵ�����
            If createCommand.ExecuteNonQuery() = 0 Then
                Throw New System.Exception("������Ŀ���������ʧ�ܡ�")
            End If

            '������ȡ��Ŀ�������ֵ�����
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

            '����´�������Ŀ��������ļ�¼��
            da.Fill(dsFinanceAnalyse, "TProjectFinanceAnalyse")

            If dsFinanceAnalyse.Tables("TProjectFinanceAnalyse").Rows.Count > 0 Then
                Dim dsCorporationAccount As DataSet = New DataSet("CorporationAccountDST")

                '�õ�����ҵ�ʻ������ݼ�
                da.SelectCommand.CommandText = "dbo.PFetchCorporationAccount"
                da.SelectCommand.Parameters("@Condition").Value = _
                 "{dbo.corporation_account.project_code LIKE '" + ProjectNo + "' AND " + _
                 " dbo.corporation_account.corporation_code LIKE '" + CorporationNo + "' AND " + _
                 " dbo.corporation_account.phase LIKE '" + Phase + "' AND " + _
                 " (dbo.corporation_account.month LIKE '" + Month + "' OR " + _
                 " dbo.corporation_account.month LIKE '" + MonthLast + "')}"

                da.Fill(dsCorporationAccount, "TCorporationAccount")

                Dim row As DataRow

                'ѭ�����¡�index_value���ֶε�ֵ
                For Each row In dsFinanceAnalyse.Tables("TProjectFinanceAnalyse").Rows
                    Try
                        row("index_value") = Me.GetIndexValue(ProjectNo, CorporationNo, Phase, Month, MonthLast, row("index_type"), row("index_id"), dsCorporationAccount.Tables("TCorporationAccount"))
                    Catch ex As System.Exception

                    End Try
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

                '������Ŀ���������¼��
                da.Update(dsFinanceAnalyse, "TProjectFinanceAnalyse")

                Return True
            End If
        Catch ex As System.Exception
            Throw ex
            Return False
        End Try

        Return False
    End Function

    '��ȡ����ҵ�������ָ������Ŀ��š���ҵ��š���Ŀ�׶Ρ����¡������Ŀ��ָ��ֵ��
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
                Return DBNull.Value
            End If
        Catch ex As System.Exception
            Throw ex
        End Try
    End Function

    '��ȡָ����Ŀ��š���ҵ��š���Ŀ�׶Ρ���������Ӧ��ƿ�Ŀ��ָ��ֵ��
    Public Function GetIndexValue(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String, ByVal IndexType As String, ByVal IndexID As String) As Object
        Try
            If moConnection.State = ConnectionState.Closed Then
                moConnection.Open()
            End If

            Dim da As SqlDataAdapter = New SqlDataAdapter
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

    '��ָ����ҵ������У���ȡָ�������Ŀ��ָ��ֵ��
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
                Case "11001"     '���ʲ���a34
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount)
                Case "11002"     'Ӫ���ʽ�a14-a49
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a14", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount)
                Case "11003"     '���ʲ���a34-a49-a55-a32 ��ʽ��Ϊ��a34-A57-a32 qxd 2004-7-27 ���������ع������
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                        Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a57", dtCorporationAccount) - _
                        Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount)
                    'Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                    'Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                    'Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount)
                Case "11004"     '�ʲ���ծ�ʣ�a57/a34
                    Dim a34 As Decimal = Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount)

                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a57", dtCorporationAccount) / a34) * 100
                Case "11005"     '�������ʣ�a14/a49
                    Dim a49 As Decimal = Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount)

                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a14", dtCorporationAccount) / a49
                Case "11006"     '�ٶ����ʣ�(a14-a10)/a49
                    Dim a49 As Decimal = Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount)

                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a14", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a10", dtCorporationAccount)) / a49
                Case "11007"     '�����ʲ������ʣ�(a55+a65)/(a28+a17)
                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) + _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a65", dtCorporationAccount)) / _
                     (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a28", dtCorporationAccount) + _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a17", dtCorporationAccount))
                Case "11008"     '���ֱ��ʣ�(a50+a35)/<003> = (a50+a35)/(a34-a49-a55-a32)
                    Return ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a50", dtCorporationAccount) + _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a35", dtCorporationAccount)) / _
                     (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount))) * 100
                Case "11009"     '���и�ծ���ʣ�a91/<003> = a91/(a34-a49-a55-a32)
                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a91", dtCorporationAccount) / _
                    (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount))) * 100
                Case "11010"     '����ڳ����ʣ�1-a92/a93
                    Dim a93 As Decimal = Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a93", dtCorporationAccount)

                    If a93 = 0 Then
                        Return DBNull.Value
                    Else
                        Return (1 - Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a92", dtCorporationAccount) / a93) * 100
                    End If
                Case "12001"     '��Ӫҵ���룺b01
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount)
                Case "12002"     '���������ʣ�b14/b01
                    Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b14", dtCorporationAccount) / _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount)) * 100
                Case "12003"     'Ӧ���ʿ���ת�ʣ�b01/((a06+a06<����>)/2)
                    If Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a06", dtCorporationAccount) Is DBNull.Value Then
                        Return DBNull.Value
                    Else
                        Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount) / _
                         ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a06", dtCorporationAccount) + _
                         Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a06", dtCorporationAccount)) / 2)
                    End If

                    'Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount) / _
                    ' ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a06", dtCorporationAccount) + _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a06", dtCorporationAccount)) / 2)
                Case "12004"     '�����ת�ʣ�b02*2/(a10+a10<����>)
                    If Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a10", dtCorporationAccount) Is DBNull.Value Then
                        Return DBNull.Value
                    Else
                        Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b02", dtCorporationAccount) * 2 / _
                         (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a10", dtCorporationAccount) + _
                         Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a10", dtCorporationAccount))
                    End If
                    'Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b02", dtCorporationAccount) * 2 / _
                    ' (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a10", dtCorporationAccount) + _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a10", dtCorporationAccount))
                Case "12005"     '���ʲ��ر��ʣ�b18/((<03>+<03><����>)/2) = b18/(((a34-a49-a55-a32)+(a34-a49-a55-a32)<����>)/2)
                    If Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a34", dtCorporationAccount) Is DBNull.Value Then
                        Return DBNull.Value
                    Else
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
                    End If
                    'Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b18", dtCorporationAccount) / _
                    ' (( _
                    ' ( _
                    '  Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                    '  Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                    '  Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                    '  Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount) _
                    ' ) + _
                    ' ( _
                    '  Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a34", dtCorporationAccount) - _
                    '  Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a49", dtCorporationAccount) - _
                    '  Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a55", dtCorporationAccount) - _
                    '  Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a32", dtCorporationAccount) _
                    ' )) / 2)) * 100
                Case "12006"     '��Ϣ���ϱ�����(b14+b08)/b08
                    Dim b08 As Decimal = Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b08", dtCorporationAccount)

                    If b08 <= 0 Then
                        Return DBNull.Value
                    Else
                        Return (Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b14", dtCorporationAccount) + _
                          Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b08", dtCorporationAccount)) / b08
                    End If
                Case "13001"     '���ʲ������ʣ�(<003>-<003><����>)/ABS(<003><����>) = ((a34-a49-a55-a32) - (a34-a49-a55-a32)<����>) / ABS((a34-a49-a55-a32)<����>)
                    If Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a34", dtCorporationAccount) Is DBNull.Value Then
                        Return DBNull.Value
                    Else
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
                    End If
                    'Return (( _
                    '   ( _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a34", dtCorporationAccount) - _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount) - _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a55", dtCorporationAccount) - _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a32", dtCorporationAccount) _
                    '   ) - _
                    '   ( _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a34", dtCorporationAccount) - _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a49", dtCorporationAccount) - _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a55", dtCorporationAccount) - _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a32", dtCorporationAccount) _
                    '   ) _
                    '  ) / _
                    '   Math.Abs( _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a34", dtCorporationAccount) - _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a49", dtCorporationAccount) - _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a55", dtCorporationAccount) - _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "01", "a32", dtCorporationAccount) _
                    '   )) * 100
                Case "13002"     '�������������ʣ�(b01-b01<����>)/ABS(b01<����>)
                    If Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b01", dtCorporationAccount) Is DBNull.Value Then
                        Return DBNull.Value
                    Else
                        Return ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount) - _
                          Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b01", dtCorporationAccount)) / _
                          Math.Abs(Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b01", dtCorporationAccount))) * 100
                    End If
                    'Return ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount) - _
                    '  Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b01", dtCorporationAccount)) / _
                    '  Math.Abs(Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b01", dtCorporationAccount))) * 100
                Case "13003"     '���������ʣ�(b14-b14<����>)/ABS(b14<����>)
                    If Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount) Is DBNull.Value Then
                        Return DBNull.Value
                    Else
                        Return ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b14", dtCorporationAccount) - _
                          Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount)) / _
                          Math.Abs(Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount))) * 100
                    End If
                    'Return ((Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b14", dtCorporationAccount) - _
                    '  Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount)) / _
                    '  Math.Abs(Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount))) * 100
                Case "13004"     '���������b14-b14<����>
                    'Dim foundRows As DataRow() = dtCorporationAccount.Select("project_code = '" + ProjectNo + "' AND corporation_code = '" + CorporationNo + "' AND phase = '" + Phase + "' AND month = '" + MonthLast + "' AND item_type = '02' AND item_code = 'b14'")

                    If Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount) Is DBNull.Value Then
                        Return DBNull.Value
                    Else

                        Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b14", dtCorporationAccount) - _
                         Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount)
                    End If
                    'Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b14", dtCorporationAccount) - _
                    ' Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, MonthLast, "02", "b14", dtCorporationAccount)
                Case "14001"     '��Ӫ�������ֽ�������/������ծ��c04/a49
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c04", dtCorporationAccount) / _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount)
                Case "14002"     '��Ӫ�������ֽ�������/��ծ�ܶc04/a57
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c04", dtCorporationAccount) / _
                      Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a57", dtCorporationAccount)
                Case "14003"     '��Ӫ�������ֽ�������/��Ӫҵ�����룺c04/b01
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c04", dtCorporationAccount) / _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "02", "b01", dtCorporationAccount)
                Case "14004"     '��Ӫ�������ֽ���������/������ծ��c10/a49
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c10", dtCorporationAccount) / _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a49", dtCorporationAccount)
                Case "14005"     '��Ӫ�������ֽ���������/��ծ�ܶc10/a57
                    Return Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "03", "c10", dtCorporationAccount) / _
                     Me.GetCorporationAccountValue(ProjectNo, CorporationNo, Phase, Month, "01", "a57", dtCorporationAccount)
                Case "14006"     '��Ӫ�����ֽ���������/��Ӫҵ�����룺c10/b01
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
        Dim da As SqlDataAdapter = New SqlDataAdapter

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
