Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'通用查询
Public Class CommonQuery

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_CommonQuery As SqlDataAdapter

    '定义查询命令
    Private CommonQueryCommand As SqlCommand
    Private GetProjectSearchInfoCommand As SqlCommand
    Private GetFinanceReviewDataCommand As SqlCommand
    Private GetQueryProjectInfoCommand As SqlCommand
    Private GetViewProjectInfoCommand As SqlCommand
    Private GetMeetProjectCommand As SqlCommand
    Private GetOverdueProjectListCommand As SqlCommand
    Private GetQueryFirstProjectCommand As SqlCommand
    Private GetQueryCorporationAttendeeCommand As SqlCommand
    Private GetQueryPauseProjectCommand As SqlCommand
    Private GetMaturityProjectReviewCommand As SqlCommand
    Private GetOnVouchProjectReviewCommand As SqlCommand
    Private GetProjectAssignReviewCommand As SqlCommand
    Private GetTerminateProjectReviewCommand As SqlCommand
    Private GetRefundDebtProjectListCommand As SqlCommand
    Private GetNeedMeetProjectInfoCommand As SqlCommand
    Private GetQueryStatisticsAssuranceInfoCommand As SqlCommand
    Private GetQueryStatisticsRegionInfoCommand As SqlCommand
    Private GetQueryStatisticsBankInfoCommand As SqlCommand
    Private GetQueryStatisticsIndustryInfoCommand As SqlCommand
    Private GetQueryStatisticsServiceTypeInfoCommand As SqlCommand
    Private ImportFirstFinanceDataCommand As SqlCommand
    Private FetchFinancialAnalysisInfoCommand As SqlCommand
    Private GetNeedSignatureProjectInfoCommand As SqlCommand
    Private FetchOppositeGuaranteeAssurerCommand As SqlCommand
    Private FetchProjectGuaranteeFormCommand As SqlCommand
    Private GetAcceptVouchDataCommand As SqlCommand
    Private GetTaskListInfoCommand As SqlCommand
    Private GetRefundProcessCommand As SqlCommand
    Private FQueryAcceptProjectCommand As SqlCommand
    Private FQueryAllocateProjectCommand As SqlCommand
    Private FQueryPresentingProjectCommand As SqlCommand
    Private FQueryLoanProjectCommand As SqlCommand
    Private FQueryProjectExpandDateCommand As SqlCommand
    Private FQuerySignProjectCommand As SqlCommand
    Private PQueryFirstTrialProjectCommand As SqlCommand
    Private FQueryRequiteProjectCommand As SqlCommand
    Private FQueryCreditProjectCommand As SqlCommand
    Private FQueryRecantProjectCommand As SqlCommand
    Private FQueryProcessingProjectCommand As SqlCommand
    Private FQueryRegionProjectCommand As SqlCommand
    Private FQueryChargeStatisticsCommand As SqlCommand
    Private DelProjectCommand As SqlCommand
    Private FQueryStatisticsCompensationCommand As SqlCommand
    Private FQueryStatisticsGEProprietorshipCommand As SqlCommand
    Private FQueryStatisticsRegionCommand As SqlCommand
    Private FQueryStatisticsCounterguaranteeByMonthCommand As SqlCommand
    Private FQueryStatisticsCounterguaranteeByYearCommand As SqlCommand
    Private FQueryStatisticsPMServiceCommand As SqlCommand
    Private PQueryWorkLogCommand As SqlCommand
    Private FQueryStatisticsGECraftCommand As SqlCommand
    Private PQueryStatisticsMarketingACommand As SqlCommand
    Private PQueryStatisticsMarketingBCommand As SqlCommand
    Private PQueryStatisticsMarketingCCommand As SqlCommand
    Private PStatisticsByTypeCommand As SqlCommand
    Private PStatisticsFeeCommand As SqlCommand
    Private PQueryStatisticsRecommendProjectByMonthCommand As SqlCommand
    Private PQueryStatisticsRecommendProjectByYearCommand As SqlCommand
    Private PQueryStatisticsRecommendProjectCommand As SqlCommand
    Private Usp_ListIsFirstLoanStatCommand As SqlCommand
    Private Usp_ListConsultationCommand As SqlCommand
    Private Usp_GetUnDealProjectCommand As SqlCommand
    Private Usp_GetGuaranteeProjectCommand As SqlCommand
    Private Usp_GetAfterGuaranteeRecordCommand As SqlCommand
    Private QryUnSignProjectCommand As SqlCommand
    Private Usp_ListGuaranteeFormCommand As SqlCommand
    Private QueryProjectRequiteCommand As SqlCommand
    Private IntentLetterCommand As SqlCommand

    Private PCopyOppGuaranteeCommand As SqlCommand
    Private PUpdateProcessCommand As SqlCommand
    Private GetGuarantyInfoExCommand As SqlCommand
    Private PQueryOppEvaluateCommand As SqlCommand

    Private FetchProjectFinanceAnalyseIntegrationCommand As SqlCommand
    Private GetTaskProjectListCommand As SqlCommand
    Private PGetGuarantingCorporationListCommand As SqlCommand
    Private DeleteAntiAssureCompanyCommand As SqlCommand
    Private LookUpWorkingCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_CommonQuery = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

    End Sub

    '通用查询信息
    Public Function GetCommonQueryInfo(ByVal strSql As String) As DataSet
        Dim tempDs As New DataSet()

        If CommonQueryCommand Is Nothing Then

            CommonQueryCommand = New SqlCommand("GetCommonQueryInfo", conn)
            CommonQueryCommand.CommandType = CommandType.StoredProcedure
            CommonQueryCommand.Parameters.Add(New SqlParameter("@strSql", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = CommonQueryCommand
            .SelectCommand.Transaction = ts
            .SelectCommand.CommandTimeout = 1200
            CommonQueryCommand.Parameters("@strSql").Value = strSql
            .Fill(tempDs)
        End With

        Return tempDs

    End Function

    '通用查询信息
    Public Function GetCommonQueryInfoForOA(ByVal strSql As String) As DataSet
        Dim tempDs As New DataSet()

        If CommonQueryCommand Is Nothing Then

            CommonQueryCommand = New SqlCommand(strSql, conn)
            CommonQueryCommand.CommandType = CommandType.Text
        End If

        With dsCommand_CommonQuery
            .SelectCommand = CommonQueryCommand
            .SelectCommand.Transaction = ts
            .SelectCommand.CommandTimeout = 1200
            .Fill(tempDs)
        End With

        Return tempDs

    End Function


    '通用查询信息
    Public Function GetCommonStatisticsInfo(ByVal condition As String, ByVal tableCondition As String, ByVal orderBy As String, ByVal cutOffDate As Date, ByVal feeStartDate As Date, ByVal feeEndDate As Date) As DataSet
        Dim tempDs As New DataSet()

        If CommonQueryCommand Is Nothing Then

            CommonQueryCommand = New SqlCommand("GetCommonStatisticsInfo", conn)
            CommonQueryCommand.CommandType = CommandType.StoredProcedure
            CommonQueryCommand.Parameters.Add(New SqlParameter("@Condion", SqlDbType.NVarChar))
            CommonQueryCommand.Parameters.Add(New SqlParameter("@tableConditon", SqlDbType.NVarChar))
            CommonQueryCommand.Parameters.Add(New SqlParameter("@orderBy", SqlDbType.NVarChar))
            CommonQueryCommand.Parameters.Add(New SqlParameter("@cutOffDate", SqlDbType.DateTime))
            CommonQueryCommand.Parameters.Add(New SqlParameter("@feeStartDate", SqlDbType.DateTime))
            CommonQueryCommand.Parameters.Add(New SqlParameter("@feeEndDate", SqlDbType.DateTime))


        End If

        With dsCommand_CommonQuery
            .SelectCommand = CommonQueryCommand
            .SelectCommand.Transaction = ts
            .SelectCommand.CommandTimeout = 1200
            CommonQueryCommand.Parameters("@Condion").Value = condition
            CommonQueryCommand.Parameters("@tableConditon").Value = tableCondition
            CommonQueryCommand.Parameters("@orderBy").Value = orderBy
            CommonQueryCommand.Parameters("@cutOffDate").Value = cutOffDate
            CommonQueryCommand.Parameters("@feeStartDate").Value = feeStartDate
            CommonQueryCommand.Parameters("@feeEndDate").Value = feeEndDate
            .Fill(tempDs)
        End With

        Return tempDs

    End Function

    '获取项目查询信息
    Public Function GetProjectSearchInfo(ByVal projectCode As String, ByVal enterpriseName As String, ByVal projectManager As String, ByVal phase As String, ByVal status As String) As DataSet
        Dim tempDs As New DataSet()

        If GetProjectSearchInfoCommand Is Nothing Then

            GetProjectSearchInfoCommand = New SqlCommand("dt_xjStatisticsProjectInfo", conn)
            GetProjectSearchInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectSearchInfoCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            GetProjectSearchInfoCommand.Parameters.Add(New SqlParameter("@EnterpriseName", SqlDbType.NVarChar))
            GetProjectSearchInfoCommand.Parameters.Add(New SqlParameter("@ProjectManager", SqlDbType.NVarChar))
            GetProjectSearchInfoCommand.Parameters.Add(New SqlParameter("@Phase", SqlDbType.NVarChar))
            GetProjectSearchInfoCommand.Parameters.Add(New SqlParameter("@Status", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetProjectSearchInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectSearchInfoCommand.Parameters("@ProjectCode").Value = projectCode
            GetProjectSearchInfoCommand.Parameters("@EnterpriseName").Value = enterpriseName
            GetProjectSearchInfoCommand.Parameters("@ProjectManager").Value = projectManager
            GetProjectSearchInfoCommand.Parameters("@Phase").Value = phase
            GetProjectSearchInfoCommand.Parameters("@Status").Value = status
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取财务信息
    Public Function GetFinanceReviewData(ByVal projectCode As String, ByVal CorporationCode As String, ByVal phase As String, ByVal ItemType As String) As DataSet
        Dim tempDs As New DataSet()

        If GetFinanceReviewDataCommand Is Nothing Then

            GetFinanceReviewDataCommand = New SqlCommand("dt_xjGetFinanceData", conn)
            GetFinanceReviewDataCommand.CommandType = CommandType.StoredProcedure
            GetFinanceReviewDataCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            GetFinanceReviewDataCommand.Parameters.Add(New SqlParameter("@CorporationCode", SqlDbType.VarChar))
            GetFinanceReviewDataCommand.Parameters.Add(New SqlParameter("@Phase", SqlDbType.NVarChar))
            GetFinanceReviewDataCommand.Parameters.Add(New SqlParameter("@ItemType", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetFinanceReviewDataCommand
            .SelectCommand.Transaction = ts
            GetFinanceReviewDataCommand.Parameters("@ProjectCode").Value = projectCode
            GetFinanceReviewDataCommand.Parameters("@CorporationCode").Value = CorporationCode
            GetFinanceReviewDataCommand.Parameters("@Phase").Value = phase
            GetFinanceReviewDataCommand.Parameters("@ItemType").Value = ItemType
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Private Function GetViewProjectInfo(ByVal strSQL_Condition As String) As DataSet

        Dim tempDs As New DataSet()

        If GetViewProjectInfoCommand Is Nothing Then

            GetViewProjectInfoCommand = New SqlCommand("GetViewProjectInfo", conn)
            GetViewProjectInfoCommand.CommandType = CommandType.StoredProcedure
            GetViewProjectInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetViewProjectInfoCommand
            .SelectCommand.Transaction = ts
            GetViewProjectInfoCommand.Parameters("@Condition").Value = strSQL_Condition
            .Fill(tempDs, "ViewProject")
        End With

        Return tempDs

    End Function

    Private Function GetViewReGuaranteeProjectInfo(ByVal strSQL_Condition As String) As DataSet

        Dim tempDs As New DataSet()

        If GetViewProjectInfoCommand Is Nothing Then

            GetViewProjectInfoCommand = New SqlCommand("GetReGuaranteeProjectInfo", conn)
            GetViewProjectInfoCommand.CommandType = CommandType.StoredProcedure
            GetViewProjectInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetViewProjectInfoCommand
            .SelectCommand.Transaction = ts
            GetViewProjectInfoCommand.Parameters("@Condition").Value = strSQL_Condition
            .Fill(tempDs, "ViewProject")
        End With

        Return tempDs

    End Function


    '获取项目查询信息
    Public Function GetQueryProjectInfo(ByVal projectCode As String, ByVal enterpriseName As String, ByVal projectManager As String, ByVal phase As String, ByVal status As String) As DataSet
        Dim tempDs As New DataSet()

        If GetProjectSearchInfoCommand Is Nothing Then

            GetProjectSearchInfoCommand = New SqlCommand("GetQueryProjectInfo", conn)
            GetProjectSearchInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectSearchInfoCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            GetProjectSearchInfoCommand.Parameters.Add(New SqlParameter("@EnterpriseName", SqlDbType.NVarChar))
            GetProjectSearchInfoCommand.Parameters.Add(New SqlParameter("@ProjectManager", SqlDbType.NVarChar))
            GetProjectSearchInfoCommand.Parameters.Add(New SqlParameter("@Phase", SqlDbType.NVarChar))
            GetProjectSearchInfoCommand.Parameters.Add(New SqlParameter("@Status", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetProjectSearchInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectSearchInfoCommand.Parameters("@ProjectCode").Value = projectCode
            GetProjectSearchInfoCommand.Parameters("@EnterpriseName").Value = enterpriseName
            GetProjectSearchInfoCommand.Parameters("@ProjectManager").Value = projectManager
            GetProjectSearchInfoCommand.Parameters("@Phase").Value = phase
            GetProjectSearchInfoCommand.Parameters("@Status").Value = status
            .Fill(tempDs, "ViewProject")
        End With

        Return tempDs
    End Function


    '获取项目的所有角色信息
    Public Function GetProjectInfoEx(ByVal strSql_Condition As String) As DataSet

        Dim i, j As Integer
        Dim strSql As String
        Dim tmpRoleName, tmpProjectCode, tmpRoleID, tmpAttend As String
        Dim WfProjectTaskAttendee As New WfProjectTaskAttendee(conn, ts)
        Dim Role As New Role(conn, ts)
        Dim dsTempAttend, dsRole, dsReturn As DataSet
        'strSql = "SELECT dbo.project_task_attendee.project_code," & _
        '         " dbo.project_task_attendee.role_id,dbo.project_task_attendee.attend_person" & _
        '         " FROM dbo.project_task INNER JOIN" & _
        '         " dbo.project_task_attendee ON " & _
        '         " dbo.project_task.project_code = dbo.project_task_attendee.project_code AND" & _
        '         " dbo.project_task.workflow_id = dbo.project_task_attendee.workflow_id AND" & _
        '         " dbo.project_task.task_id = dbo.project_task_attendee.task_id" & _
        '         " WHERE (dbo.project_task.start_time IS NOT NULL)" & _
        '         " order by dbo.project_task.start_time"
        strSql = " select distinct project_code,role_id,attend_person from project_task_attendee where attend_person<>''"
        dsTempAttend = GetCommonQueryInfo(strSql)

        '不取咨询人员
        strSql = "{role_id<>'11'}"
        dsRole = Role.FetchRole(strSql)
        dsReturn = GetViewProjectInfo(strSql_Condition)
        Dim iCount As Integer = dsReturn.Tables(0).Columns.Count

        '增加所有角色字段
        For i = 0 To dsRole.Tables(0).Rows.Count - 1
            tmpRoleName = Trim(dsRole.Tables(0).Rows(i).Item("role_id"))
            dsReturn.Tables(0).Columns.Add(tmpRoleName, GetType(System.String))
        Next

        For i = 0 To dsReturn.Tables(0).Rows.Count - 1
            tmpProjectCode = dsReturn.Tables(0).Rows(i).Item("projectcode")

            '填充该项目的所有角色参与人
            For j = 0 To dsRole.Tables(0).Rows.Count - 1
                tmpRoleID = Trim(dsRole.Tables(0).Rows(j).Item("role_id"))

                strSql = "project_code=" & "'" & tmpProjectCode & "'" & " and role_id=" & "'" & tmpRoleID & "'"
                If dsTempAttend.Tables(0).Select(strSql).Length() = 0 Then
                    tmpAttend = ""
                Else
                    tmpAttend = dsTempAttend.Tables(0).Select(strSql)(0).Item("attend_person")
                End If

                '填充项目参与人
                dsReturn.Tables(0).Rows(i).Item(j + iCount) = tmpAttend

            Next

        Next

        Return dsReturn
    End Function

    '获取项目的所有角色信息
    Public Function GetReGuaranteeProjectInfo(ByVal strSql_Condition As String) As DataSet

        Dim i, j As Integer
        Dim strSql As String
        Dim tmpRoleName, tmpProjectCode, tmpRoleID, tmpAttend As String
        Dim WfProjectTaskAttendee As New WfProjectTaskAttendee(conn, ts)
        Dim Role As New Role(conn, ts)
        Dim dsTempAttend, dsRole, dsReturn As DataSet

        strSql = " select distinct project_code,role_id,attend_person from project_task_attendee where attend_person<>''"
        dsTempAttend = GetCommonQueryInfo(strSql)

        '不取咨询人员
        strSql = "{role_id<>'11'}"
        dsRole = Role.FetchRole(strSql)
        dsReturn = GetViewReGuaranteeProjectInfo(strSql_Condition)
        Dim iCount As Integer = dsReturn.Tables(0).Columns.Count

        '增加所有角色字段
        For i = 0 To dsRole.Tables(0).Rows.Count - 1
            tmpRoleName = Trim(dsRole.Tables(0).Rows(i).Item("role_id"))
            dsReturn.Tables(0).Columns.Add(tmpRoleName, GetType(System.String))
        Next

        For i = 0 To dsReturn.Tables(0).Rows.Count - 1
            tmpProjectCode = dsReturn.Tables(0).Rows(i).Item("projectcode")

            '填充该项目的所有角色参与人
            For j = 0 To dsRole.Tables(0).Rows.Count - 1
                tmpRoleID = Trim(dsRole.Tables(0).Rows(j).Item("role_id"))

                strSql = "project_code=" & "'" & tmpProjectCode & "'" & " and role_id=" & "'" & tmpRoleID & "'"
                If dsTempAttend.Tables(0).Select(strSql).Length() = 0 Then
                    tmpAttend = ""
                Else
                    tmpAttend = dsTempAttend.Tables(0).Select(strSql)(0).Item("attend_person")
                End If

                '填充项目参与人
                dsReturn.Tables(0).Rows(i).Item(j + iCount) = tmpAttend

            Next

        Next

        Return dsReturn
    End Function

    Public Function GetQueryProjectInfo(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationName As String, ByVal DistrictName As String, ByVal Phase As String, ByVal ApplyDateFrom As DateTime, ByVal ApplyDateTo As DateTime) As DataSet
        Dim tempDs As New DataSet()

        If GetQueryProjectInfoCommand Is Nothing Then

            GetQueryProjectInfoCommand = New SqlCommand("PQueryProject", conn)
            GetQueryProjectInfoCommand.CommandType = CommandType.StoredProcedure
            GetQueryProjectInfoCommand.Parameters.Add(New SqlParameter("@ProjectNo", SqlDbType.NVarChar))
            GetQueryProjectInfoCommand.Parameters.Add(New SqlParameter("@CorporationNo", SqlDbType.NVarChar))
            GetQueryProjectInfoCommand.Parameters.Add(New SqlParameter("@CorporationName", SqlDbType.NVarChar))
            GetQueryProjectInfoCommand.Parameters.Add(New SqlParameter("@DistrictName", SqlDbType.NVarChar))
            GetQueryProjectInfoCommand.Parameters.Add(New SqlParameter("@Phase", SqlDbType.NVarChar))
            GetQueryProjectInfoCommand.Parameters.Add(New SqlParameter("@ApplyDateFrom", SqlDbType.DateTime))
            GetQueryProjectInfoCommand.Parameters.Add(New SqlParameter("@ApplyDateTo", SqlDbType.DateTime))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetQueryProjectInfoCommand
            .SelectCommand.Transaction = ts
            GetQueryProjectInfoCommand.Parameters("@ProjectNo").Value = ProjectNo
            GetQueryProjectInfoCommand.Parameters("@CorporationNo").Value = CorporationNo
            GetQueryProjectInfoCommand.Parameters("@CorporationName").Value = CorporationName
            GetQueryProjectInfoCommand.Parameters("@DistrictName").Value = DistrictName
            GetQueryProjectInfoCommand.Parameters("@Phase").Value = Phase
            GetQueryProjectInfoCommand.Parameters("@ApplyDateFrom").Value = ApplyDateFrom
            GetQueryProjectInfoCommand.Parameters("@ApplyDateTo").Value = ApplyDateTo
            .Fill(tempDs)
        End With

        Return tempDs

    End Function

    '获取会议项目信息
    Public Function GetMeetProject(ByVal startDate As DateTime, ByVal endDate As DateTime) As DataSet
        Dim tempDs As New DataSet()

        If GetMeetProjectCommand Is Nothing Then

            GetMeetProjectCommand = New SqlCommand("dt_lqfGetMeetProject", conn)
            GetMeetProjectCommand.CommandType = CommandType.StoredProcedure
            GetMeetProjectCommand.Parameters.Add(New SqlParameter("@startDate", SqlDbType.DateTime))
            GetMeetProjectCommand.Parameters.Add(New SqlParameter("@endDate", SqlDbType.DateTime))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetMeetProjectCommand
            .SelectCommand.Transaction = ts
            GetMeetProjectCommand.Parameters("@startDate").Value = startDate
            GetMeetProjectCommand.Parameters("@endDate").Value = endDate
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取项目评价信息
    Public Function GetProjectAppraisement(ByVal project_code As String, ByVal EnterpriseName As String, ByVal ServiceType As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet()

        If GetMeetProjectCommand Is Nothing Then

            GetMeetProjectCommand = New SqlCommand("FQueryProjectAppraisement", conn)
            GetMeetProjectCommand.CommandType = CommandType.StoredProcedure
            GetMeetProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.VarChar))
            GetMeetProjectCommand.Parameters.Add(New SqlParameter("@EnterpriseName", SqlDbType.VarChar))
            GetMeetProjectCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.VarChar))
            GetMeetProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.VarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetMeetProjectCommand
            .SelectCommand.Transaction = ts
            GetMeetProjectCommand.Parameters("@project_code").Value = project_code
            GetMeetProjectCommand.Parameters("@EnterpriseName").Value = EnterpriseName
            GetMeetProjectCommand.Parameters("@ServiceType").Value = ServiceType
            GetMeetProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取会议项目信息
    Public Function GetOverdueProjectList(ByVal ProjectCode As String, ByVal ServiceType As String, ByVal StartTime As String, ByVal EndTime As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet()

        If GetOverdueProjectListCommand Is Nothing Then
            GetOverdueProjectListCommand = New SqlCommand("dt_lqfGetOverdueProjectList", conn)
            GetOverdueProjectListCommand.CommandType = CommandType.StoredProcedure
            GetOverdueProjectListCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            GetOverdueProjectListCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.NVarChar))
            GetOverdueProjectListCommand.Parameters.Add(New SqlParameter("@StartTime", SqlDbType.DateTime))
            GetOverdueProjectListCommand.Parameters.Add(New SqlParameter("@EndTime", SqlDbType.DateTime))
            GetOverdueProjectListCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.VarChar))
            GetOverdueProjectListCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetOverdueProjectListCommand
            .SelectCommand.Transaction = ts
            GetOverdueProjectListCommand.Parameters("@ProjectCode").Value = ProjectCode
            GetOverdueProjectListCommand.Parameters("@ServiceType").Value = ServiceType
            GetOverdueProjectListCommand.Parameters("@StartTime").Value = StartTime
            GetOverdueProjectListCommand.Parameters("@EndTime").Value = EndTime
            GetOverdueProjectListCommand.Parameters("@vchPMA").Value = vchPMA
            GetOverdueProjectListCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取初审项目
    Public Function GetQueryFirstProject(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationName As String, ByVal Phase As String, ByVal ServiceType As String, ByVal FromDate As String, ByVal ToDate As String, ByVal vchAcceptBranch As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet()

        If GetQueryFirstProjectCommand Is Nothing Then
            GetQueryFirstProjectCommand = New SqlCommand("PQueryFirstProject", conn)
            GetQueryFirstProjectCommand.CommandType = CommandType.StoredProcedure
            GetQueryFirstProjectCommand.Parameters.Add(New SqlParameter("@ProjectNo", SqlDbType.NVarChar))
            GetQueryFirstProjectCommand.Parameters.Add(New SqlParameter("@CorporationNo", SqlDbType.NVarChar))
            GetQueryFirstProjectCommand.Parameters.Add(New SqlParameter("@CorporationName", SqlDbType.NVarChar))
            GetQueryFirstProjectCommand.Parameters.Add(New SqlParameter("@Phase", SqlDbType.NVarChar))
            GetQueryFirstProjectCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.NVarChar))
            GetQueryFirstProjectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime))
            GetQueryFirstProjectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime))
            GetQueryFirstProjectCommand.Parameters.Add(New SqlParameter("@vchAcceptBranch", SqlDbType.NVarChar))
            GetQueryFirstProjectCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.NVarChar))
            GetQueryFirstProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetQueryFirstProjectCommand
            .SelectCommand.Transaction = ts
            GetQueryFirstProjectCommand.Parameters("@ProjectNo").Value = ProjectNo
            GetQueryFirstProjectCommand.Parameters("@CorporationNo").Value = CorporationNo
            GetQueryFirstProjectCommand.Parameters("@CorporationName").Value = CorporationName
            GetQueryFirstProjectCommand.Parameters("@Phase").Value = Phase
            GetQueryFirstProjectCommand.Parameters("@ServiceType").Value = ServiceType
            GetQueryFirstProjectCommand.Parameters("@FromDate").Value = FromDate
            GetQueryFirstProjectCommand.Parameters("@ToDate").Value = ToDate
            GetQueryFirstProjectCommand.Parameters("@vchAcceptBranch").Value = vchAcceptBranch
            GetQueryFirstProjectCommand.Parameters("@vchPMA").Value = vchPMA
            GetQueryFirstProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function


    '获取企业参与人信息
    Public Function GetQueryCorporationAttendee(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationName As String, ByVal Phase As String, ByVal ServiceType As String, ByVal FromDate As DateTime, ByVal ToDate As DateTime) As DataSet
        Dim tempDs As New DataSet()

        If GetQueryCorporationAttendeeCommand Is Nothing Then
            GetQueryCorporationAttendeeCommand = New SqlCommand("PQueryCorporationAttendee", conn)
            GetQueryCorporationAttendeeCommand.CommandType = CommandType.StoredProcedure
            GetQueryCorporationAttendeeCommand.Parameters.Add(New SqlParameter("@ProjectNo", SqlDbType.NVarChar))
            GetQueryCorporationAttendeeCommand.Parameters.Add(New SqlParameter("@CorporationNo", SqlDbType.NVarChar))
            GetQueryCorporationAttendeeCommand.Parameters.Add(New SqlParameter("@CorporationName", SqlDbType.NVarChar))
            GetQueryCorporationAttendeeCommand.Parameters.Add(New SqlParameter("@Phase", SqlDbType.NVarChar))
            GetQueryCorporationAttendeeCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.NVarChar))
            GetQueryCorporationAttendeeCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime))
            GetQueryCorporationAttendeeCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetQueryCorporationAttendeeCommand
            .SelectCommand.Transaction = ts
            GetQueryCorporationAttendeeCommand.Parameters("@ProjectNo").Value = ProjectNo
            GetQueryCorporationAttendeeCommand.Parameters("@CorporationNo").Value = CorporationNo
            GetQueryCorporationAttendeeCommand.Parameters("@CorporationName").Value = CorporationName
            GetQueryCorporationAttendeeCommand.Parameters("@Phase").Value = Phase
            GetQueryCorporationAttendeeCommand.Parameters("@ServiceType").Value = ServiceType
            GetQueryCorporationAttendeeCommand.Parameters("@FromDate").Value = FromDate
            GetQueryCorporationAttendeeCommand.Parameters("@ToDate").Value = ToDate
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取暂停项目信息
    Public Function GetQueryPauseProject(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationName As String, ByVal Phase As String, ByVal ServiceType As String, ByVal FromDate As String, ByVal ToDate As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet()

        If GetQueryPauseProjectCommand Is Nothing Then
            GetQueryPauseProjectCommand = New SqlCommand("PQueryPauseProject", conn)
            GetQueryPauseProjectCommand.CommandType = CommandType.StoredProcedure
            GetQueryPauseProjectCommand.Parameters.Add(New SqlParameter("@ProjectNo", SqlDbType.NVarChar))
            GetQueryPauseProjectCommand.Parameters.Add(New SqlParameter("@CorporationNo", SqlDbType.NVarChar))
            GetQueryPauseProjectCommand.Parameters.Add(New SqlParameter("@CorporationName", SqlDbType.NVarChar))
            GetQueryPauseProjectCommand.Parameters.Add(New SqlParameter("@Phase", SqlDbType.NVarChar))
            GetQueryPauseProjectCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.NVarChar))
            GetQueryPauseProjectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime))
            GetQueryPauseProjectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime))
            GetQueryPauseProjectCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.NVarChar))
            GetQueryPauseProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetQueryPauseProjectCommand
            .SelectCommand.Transaction = ts
            GetQueryPauseProjectCommand.Parameters("@ProjectNo").Value = ProjectNo
            GetQueryPauseProjectCommand.Parameters("@CorporationNo").Value = CorporationNo
            GetQueryPauseProjectCommand.Parameters("@CorporationName").Value = CorporationName
            GetQueryPauseProjectCommand.Parameters("@Phase").Value = Phase
            GetQueryPauseProjectCommand.Parameters("@ServiceType").Value = ServiceType
            GetQueryPauseProjectCommand.Parameters("@FromDate").Value = FromDate
            GetQueryPauseProjectCommand.Parameters("@ToDate").Value = ToDate
            GetQueryPauseProjectCommand.Parameters("@vchPMA").Value = vchPMA
            GetQueryPauseProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取到期项目一览表
    Public Function GetMaturityProjectReview(ByVal ServiceType As String, ByVal StartDate As String, ByVal EndDate As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet()

        If GetMaturityProjectReviewCommand Is Nothing Then
            GetMaturityProjectReviewCommand = New SqlCommand("dt_xjMaturityProjectReview", conn)
            GetMaturityProjectReviewCommand.CommandType = CommandType.StoredProcedure
            GetMaturityProjectReviewCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.NVarChar))
            GetMaturityProjectReviewCommand.Parameters.Add(New SqlParameter("@StartDate", SqlDbType.DateTime))
            GetMaturityProjectReviewCommand.Parameters.Add(New SqlParameter("@EndDate", SqlDbType.DateTime))
            GetMaturityProjectReviewCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.VarChar))
            GetMaturityProjectReviewCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetMaturityProjectReviewCommand
            .SelectCommand.Transaction = ts
            GetMaturityProjectReviewCommand.Parameters("@ServiceType").Value = ServiceType
            GetMaturityProjectReviewCommand.Parameters("@StartDate").Value = StartDate
            GetMaturityProjectReviewCommand.Parameters("@EndDate").Value = EndDate
            GetMaturityProjectReviewCommand.Parameters("@vchPMA").Value = vchPMA
            GetMaturityProjectReviewCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取在保项目一览表
    Public Function GetOnVouchProjectReview(ByVal StartDate As DateTime, ByVal EndDate As DateTime) As DataSet
        Dim tempDs As New DataSet()

        If GetOnVouchProjectReviewCommand Is Nothing Then
            GetOnVouchProjectReviewCommand = New SqlCommand("dt_xjOnVouchProjectReview", conn)
            GetOnVouchProjectReviewCommand.CommandType = CommandType.StoredProcedure
            GetOnVouchProjectReviewCommand.Parameters.Add(New SqlParameter("@StartDate", SqlDbType.DateTime))
            GetOnVouchProjectReviewCommand.Parameters.Add(New SqlParameter("@EndDate", SqlDbType.DateTime))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetOnVouchProjectReviewCommand
            .SelectCommand.Transaction = ts
            GetOnVouchProjectReviewCommand.Parameters("@StartDate").Value = StartDate
            GetOnVouchProjectReviewCommand.Parameters("@EndDate").Value = EndDate
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取项目分配一览表
    Public Function GetProjectAssignReview(ByVal StartDate As DateTime, ByVal EndDate As DateTime) As DataSet
        Dim tempDs As New DataSet()

        If GetProjectAssignReviewCommand Is Nothing Then
            GetProjectAssignReviewCommand = New SqlCommand("dt_xjProjectAssignReview", conn)
            GetProjectAssignReviewCommand.CommandType = CommandType.StoredProcedure
            GetProjectAssignReviewCommand.Parameters.Add(New SqlParameter("@StartDate", SqlDbType.DateTime))
            GetProjectAssignReviewCommand.Parameters.Add(New SqlParameter("@EndDate", SqlDbType.DateTime))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetProjectAssignReviewCommand
            .SelectCommand.Transaction = ts
            GetProjectAssignReviewCommand.Parameters("@StartDate").Value = StartDate
            GetProjectAssignReviewCommand.Parameters("@EndDate").Value = EndDate
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取终止项目一览表
    Public Function GetTerminateProjectReview(ByVal ServiceType As String, ByVal StartDate As String, ByVal EndDate As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet()

        If GetTerminateProjectReviewCommand Is Nothing Then
            GetTerminateProjectReviewCommand = New SqlCommand("dt_xjTerminateProjectReview", conn)
            GetTerminateProjectReviewCommand.CommandType = CommandType.StoredProcedure
            GetTerminateProjectReviewCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.NVarChar))
            GetTerminateProjectReviewCommand.Parameters.Add(New SqlParameter("@StartDate", SqlDbType.DateTime))
            GetTerminateProjectReviewCommand.Parameters.Add(New SqlParameter("@EndDate", SqlDbType.DateTime))
            GetTerminateProjectReviewCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.VarChar))
            GetTerminateProjectReviewCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetTerminateProjectReviewCommand
            .SelectCommand.Transaction = ts
            GetTerminateProjectReviewCommand.Parameters("@ServiceType").Value = ServiceType
            GetTerminateProjectReviewCommand.Parameters("@StartDate").Value = StartDate
            GetTerminateProjectReviewCommand.Parameters("@EndDate").Value = EndDate
            GetTerminateProjectReviewCommand.Parameters("@vchPMA").Value = vchPMA
            GetTerminateProjectReviewCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取获取代偿项目列表
    Public Function GetRefundDebtProjectList(ByVal ProjectCode As String, ByVal ServiceType As String, ByVal StartTime As String, ByVal EndTime As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet()

        If GetRefundDebtProjectListCommand Is Nothing Then
            GetRefundDebtProjectListCommand = New SqlCommand("dt_lqfGetRefundDebtProjectList", conn)
            GetRefundDebtProjectListCommand.CommandType = CommandType.StoredProcedure
            GetRefundDebtProjectListCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            GetRefundDebtProjectListCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.NVarChar))
            GetRefundDebtProjectListCommand.Parameters.Add(New SqlParameter("@StartTime", SqlDbType.DateTime))
            GetRefundDebtProjectListCommand.Parameters.Add(New SqlParameter("@EndTime", SqlDbType.DateTime))
            GetRefundDebtProjectListCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.VarChar))
            GetRefundDebtProjectListCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetRefundDebtProjectListCommand
            .SelectCommand.Transaction = ts
            GetRefundDebtProjectListCommand.Parameters("@ProjectCode").Value = ProjectCode
            GetRefundDebtProjectListCommand.Parameters("@ServiceType").Value = ServiceType
            GetRefundDebtProjectListCommand.Parameters("@StartTime").Value = StartTime
            GetRefundDebtProjectListCommand.Parameters("@EndTime").Value = EndTime
            GetRefundDebtProjectListCommand.Parameters("@vchPMA").Value = vchPMA
            GetRefundDebtProjectListCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取需要上会的项目信息
    Public Function GetNeedMeetProjectInfo(ByVal ProjectList As String, ByVal ConferenceCode As String, ByVal Status As String) As DataSet
        Dim tempDs As New DataSet()

        If GetNeedMeetProjectInfoCommand Is Nothing Then
            GetNeedMeetProjectInfoCommand = New SqlCommand("dt_lqfGetNeedMeetProjectInfo", conn)
            GetNeedMeetProjectInfoCommand.CommandType = CommandType.StoredProcedure

            GetNeedMeetProjectInfoCommand.Parameters.Add(New SqlParameter("@ProjectList", SqlDbType.VarChar, 8000))
            GetNeedMeetProjectInfoCommand.Parameters.Add(New SqlParameter("@ConferenceCode", SqlDbType.Int))
            GetNeedMeetProjectInfoCommand.Parameters.Add(New SqlParameter("@Status", SqlDbType.VarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetNeedMeetProjectInfoCommand
            .SelectCommand.Transaction = ts
            GetNeedMeetProjectInfoCommand.Parameters("@ProjectList").Value = ProjectList
            GetNeedMeetProjectInfoCommand.Parameters("@ConferenceCode").Value = ConferenceCode
            GetNeedMeetProjectInfoCommand.Parameters("@Status").Value = Status
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取状态列表
    Public Function GetWfTaskStatus() As DataSet
        Dim strSql As String
        strSql = "SELECT DISTINCT ISNULL(status, '') AS status" & _
                 " FROM dbo.task_transfer_template" & _
                 " WHERE (ISNULL(status, '') <> '')"
        GetWfTaskStatus = GetCommonQueryInfo(strSql)
    End Function


    '获取担保查询统计信息
    Public Function GetQueryStatisticsAssuranceInfo(ByVal month_start As String, ByVal month_end As String, ByVal type As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet()

        If GetQueryStatisticsAssuranceInfoCommand Is Nothing Then
            GetQueryStatisticsAssuranceInfoCommand = New SqlCommand("Usp_GetReportStatByYM", conn)
            GetQueryStatisticsAssuranceInfoCommand.CommandType = CommandType.StoredProcedure
            GetQueryStatisticsAssuranceInfoCommand.Parameters.Add(New SqlParameter("@vchYMFrom", SqlDbType.NVarChar))
            GetQueryStatisticsAssuranceInfoCommand.Parameters.Add(New SqlParameter("@vchYMTo", SqlDbType.NVarChar))
            GetQueryStatisticsAssuranceInfoCommand.Parameters.Add(New SqlParameter("@vchType", SqlDbType.NVarChar))
            GetQueryStatisticsAssuranceInfoCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetQueryStatisticsAssuranceInfoCommand
            .SelectCommand.Transaction = ts
            GetQueryStatisticsAssuranceInfoCommand.Parameters("@vchYMFrom").Value = month_start
            GetQueryStatisticsAssuranceInfoCommand.Parameters("@vchYMTo").Value = month_end
            GetQueryStatisticsAssuranceInfoCommand.Parameters("@vchType").Value = type
            GetQueryStatisticsAssuranceInfoCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取区域统计信息
    Public Function GetQueryStatisticsRegionInfo(ByVal DateFrom As DateTime, ByVal DateTo As DateTime) As DataSet
        Dim tempDs As New DataSet()

        If GetQueryStatisticsRegionInfoCommand Is Nothing Then
            GetQueryStatisticsRegionInfoCommand = New SqlCommand("PQueryStatisticsRegion", conn)
            GetQueryStatisticsRegionInfoCommand.CommandType = CommandType.StoredProcedure
            GetQueryStatisticsRegionInfoCommand.Parameters.Add(New SqlParameter("@DateFrom", SqlDbType.DateTime))
            GetQueryStatisticsRegionInfoCommand.Parameters.Add(New SqlParameter("@DateTo", SqlDbType.DateTime))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetQueryStatisticsRegionInfoCommand
            .SelectCommand.Transaction = ts
            GetQueryStatisticsRegionInfoCommand.Parameters("@DateFrom").Value = DateFrom
            GetQueryStatisticsRegionInfoCommand.Parameters("@DateTo").Value = DateTo
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取银行统计信息
    Public Function GetQueryStatisticsBankInfo(ByVal DateFrom As DateTime, ByVal DateTo As DateTime, ByVal iType As Integer) As DataSet
        Dim tempDs As New DataSet()

        If GetQueryStatisticsBankInfoCommand Is Nothing Then
            GetQueryStatisticsBankInfoCommand = New SqlCommand("PQueryStatisticsBank", conn)
            GetQueryStatisticsBankInfoCommand.CommandType = CommandType.StoredProcedure
            GetQueryStatisticsBankInfoCommand.Parameters.Add(New SqlParameter("@DateFrom", SqlDbType.DateTime))
            GetQueryStatisticsBankInfoCommand.Parameters.Add(New SqlParameter("@DateTo", SqlDbType.DateTime))
            GetQueryStatisticsBankInfoCommand.Parameters.Add(New SqlParameter("@GroupType", SqlDbType.Int))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetQueryStatisticsBankInfoCommand
            .SelectCommand.Transaction = ts
            GetQueryStatisticsBankInfoCommand.Parameters("@DateFrom").Value = DateFrom
            GetQueryStatisticsBankInfoCommand.Parameters("@DateTo").Value = DateTo
            GetQueryStatisticsBankInfoCommand.Parameters("@GroupType").Value = iType
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取行业统计信息
    Public Function GetQueryStatisticsIndustryInfo(ByVal DateFrom As DateTime, ByVal DateTo As DateTime) As DataSet
        Dim tempDs As New DataSet()

        If GetQueryStatisticsIndustryInfoCommand Is Nothing Then
            GetQueryStatisticsIndustryInfoCommand = New SqlCommand("PQueryStatisticsIndustry", conn)
            GetQueryStatisticsIndustryInfoCommand.CommandType = CommandType.StoredProcedure
            GetQueryStatisticsIndustryInfoCommand.Parameters.Add(New SqlParameter("@DateFrom", SqlDbType.DateTime))
            GetQueryStatisticsIndustryInfoCommand.Parameters.Add(New SqlParameter("@DateTo", SqlDbType.DateTime))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetQueryStatisticsIndustryInfoCommand
            .SelectCommand.Transaction = ts
            GetQueryStatisticsIndustryInfoCommand.Parameters("@DateFrom").Value = DateFrom
            GetQueryStatisticsIndustryInfoCommand.Parameters("@DateTo").Value = DateTo
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取服务类型统计信息
    Public Function GetQueryStatisticsServiceTypeInfo(ByVal DateFrom As DateTime, ByVal DateTo As DateTime) As DataSet
        Dim tempDs As New DataSet()

        If GetQueryStatisticsServiceTypeInfoCommand Is Nothing Then
            GetQueryStatisticsServiceTypeInfoCommand = New SqlCommand("PQueryStatisticsServiceType", conn)
            GetQueryStatisticsServiceTypeInfoCommand.CommandType = CommandType.StoredProcedure
            GetQueryStatisticsServiceTypeInfoCommand.Parameters.Add(New SqlParameter("@DateFrom", SqlDbType.DateTime))
            GetQueryStatisticsServiceTypeInfoCommand.Parameters.Add(New SqlParameter("@DateTo", SqlDbType.DateTime))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetQueryStatisticsServiceTypeInfoCommand
            .SelectCommand.Transaction = ts
            GetQueryStatisticsServiceTypeInfoCommand.Parameters("@DateFrom").Value = DateFrom
            GetQueryStatisticsServiceTypeInfoCommand.Parameters("@DateTo").Value = DateTo
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function LookUpWorking(ByVal userID As String) As DataSet
        Dim tempDs As New DataSet

        If LookUpWorkingCommand Is Nothing Then
            LookUpWorkingCommand = New SqlCommand("LookUpWorking", conn)
            LookUpWorkingCommand.CommandType = CommandType.StoredProcedure
            LookUpWorkingCommand.Parameters.Add(New SqlParameter("@UserID", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = LookUpWorkingCommand
            .SelectCommand.Transaction = ts
            LookUpWorkingCommand.Parameters("@UserID").Value = userID
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取项目进度表
    Public Function GetProjectScheduleInfo(ByVal projectID As String) As DataSet
        Dim strSql As String
        Dim dsProjectSchedule As DataSet

        '获取完成任务的进度列表
        strSql = " SELECT project_code, task_id, task_name, attend_person," & _
                 " case task_status when 'F' then '完成' end as task_status," & _
                 " start_time,finish_time, project_phase, project_status,start_mode " & _
                 " FROM work_log" & _
                 " where auto=1 and project_code = " & " '" & projectID & "' order by serial_num"
        dsProjectSchedule = GetCommonQueryInfo(strSql)

        '获取正在进行任务的列表并合并列表
        strSql = " SELECT dbo.project_task_attendee.project_code, dbo.project_task_attendee.task_id, " & _
                 " dbo.project_task.task_name, dbo.project_task_attendee.attend_person, " & _
                 " case dbo.project_task_attendee.task_status when 'P' then '处理' end as task_status," & _
                 " dbo.project_task.start_time, " & _
                 " case dbo.project_task_attendee.task_status when 'P' then null else dbo.project_task_attendee.end_time end as finish_time," & _
                 "  '' AS project_phase," & _
                 "  '' AS project_status,start_mode " & _
                 " FROM dbo.project_task_attendee INNER JOIN" & _
                 " dbo.project_task ON " & _
                 " dbo.project_task_attendee.project_code = dbo.project_task.project_code AND " & _
                 " dbo.project_task_attendee.workflow_id = dbo.project_task.workflow_id AND " & _
                 " dbo.project_task_attendee.task_id = dbo.project_task.task_id" & _
                 " where dbo.project_task_attendee.task_status='P'" & _
                 " and dbo.project_task_attendee.project_code=" & "'" & projectID & "'"

        dsProjectSchedule.Merge(GetCommonQueryInfo(strSql))

        Return dsProjectSchedule
    End Function

    '导入财务数据
    Public Function ImportFinanceData(ByVal CorporationCode As String, ByVal FromProjectCode As String, ByVal FromPhase As String, ByVal FromMonth As String, ByVal ToCorporationCode As String, ByVal ToProjectCode As String, ByVal ToPhase As String, ByVal CreatePerson As String, ByVal CreateDate As DateTime, ByVal DeleteOriginalData As Boolean)
        'Modified by LQF 2003-12-24
        If ImportFirstFinanceDataCommand Is Nothing Then
            ImportFirstFinanceDataCommand = New SqlCommand("dbo.PImportFinanceData", conn)
            ImportFirstFinanceDataCommand.CommandType = CommandType.StoredProcedure
            ImportFirstFinanceDataCommand.Parameters.Add(New SqlParameter("@CorporationCode", SqlDbType.NVarChar))
            ImportFirstFinanceDataCommand.Parameters.Add(New SqlParameter("@FromProjectCode", SqlDbType.NVarChar))
            ImportFirstFinanceDataCommand.Parameters.Add(New SqlParameter("@FromPhase", SqlDbType.NVarChar))
            ImportFirstFinanceDataCommand.Parameters.Add(New SqlParameter("@FromMonth", SqlDbType.NVarChar))
            ImportFirstFinanceDataCommand.Parameters.Add(New SqlParameter("@ToCorporationCode", SqlDbType.NVarChar))
            ImportFirstFinanceDataCommand.Parameters.Add(New SqlParameter("@ToProjectCode", SqlDbType.NVarChar))
            ImportFirstFinanceDataCommand.Parameters.Add(New SqlParameter("@ToPhase", SqlDbType.NVarChar))
            ImportFirstFinanceDataCommand.Parameters.Add(New SqlParameter("@CreatePerson", SqlDbType.NVarChar))
            ImportFirstFinanceDataCommand.Parameters.Add(New SqlParameter("@CreateDate", SqlDbType.DateTime))
            ImportFirstFinanceDataCommand.Parameters.Add(New SqlParameter("@DeleteOriginalData", SqlDbType.Bit))
        End If
        'Modified by LQF 2005-8-24
        ImportFirstFinanceDataCommand.Transaction = ts
        If FromMonth.IndexOf(",") >= 0 Then
            Dim fromProjectCodes(), fromPhases(), fromMonths() As String
            fromProjectCodes = FromProjectCode.Split(New Char() {","})
            fromPhases = FromPhase.Split(New Char() {","})
            fromMonths = FromMonth.Split(New Char() {","})
            Dim i, length As Int16
            length = fromPhases.Length
            For i = 0 To length - 1
                With ImportFirstFinanceDataCommand
                    '.Transaction = ts
                    .Parameters("@CorporationCode").Value = CorporationCode
                    .Parameters("@FromProjectCode").Value = fromProjectCodes(i)
                    .Parameters("@FromPhase").Value = fromPhases(i)
                    .Parameters("@FromMonth").Value = fromMonths(i)
                    .Parameters("@ToCorporationCode").Value = ToCorporationCode
                    .Parameters("@ToProjectCode").Value = ToProjectCode
                    .Parameters("@ToPhase").Value = ToPhase
                    .Parameters("@CreatePerson").Value = CreatePerson
                    .Parameters("@CreateDate").Value = CreateDate
                    .Parameters("@DeleteOriginalData").Value = DeleteOriginalData
                    .ExecuteNonQuery()
                End With
            Next
        Else
            With ImportFirstFinanceDataCommand
                '.Transaction = ts
                .Parameters("@CorporationCode").Value = CorporationCode
                .Parameters("@FromProjectCode").Value = FromProjectCode
                .Parameters("@FromPhase").Value = FromPhase
                .Parameters("@FromMonth").Value = FromMonth
                .Parameters("@ToCorporationCode").Value = ToCorporationCode
                .Parameters("@ToProjectCode").Value = ToProjectCode
                .Parameters("@ToPhase").Value = ToPhase
                .Parameters("@CreatePerson").Value = CreatePerson
                .Parameters("@CreateDate").Value = CreateDate
                .Parameters("@DeleteOriginalData").Value = DeleteOriginalData
                .ExecuteNonQuery()
            End With
        End If

    End Function

    '删除反担保企业
    Public Function DeleteAntiAssureCompany(ByVal project_code As String, ByVal corporation_code As String)
        If DeleteAntiAssureCompanyCommand Is Nothing Then
            DeleteAntiAssureCompanyCommand = New SqlCommand("dbo.PDeleteAntiAssureCompany", conn)
            DeleteAntiAssureCompanyCommand.CommandType = CommandType.StoredProcedure
            DeleteAntiAssureCompanyCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.Char))
            DeleteAntiAssureCompanyCommand.Parameters.Add(New SqlParameter("@corporation_code", SqlDbType.Char))
        End If
        With DeleteAntiAssureCompanyCommand
            .Transaction = ts
            .Parameters("@project_code").Value = project_code
            .Parameters("@corporation_code").Value = corporation_code
            .ExecuteNonQuery()
        End With

    End Function

    '获取任务的项目列表信息
    Public Function GetTaskProjectList(ByVal taskID As String, ByVal userName As String) As DataSet
        Return GetTaskProjectList(taskID, userName, 0)
    End Function

    '获取任务的项目列表信息
    Public Function GetTaskProjectList(ByVal taskID As String, ByVal userName As String, ByVal flag As Integer) As DataSet
        Dim i, j As Integer
        Dim strSql As String
        Dim tmpProjectCode, tmpAttend, tmpRollBackTask, tmpRollBackTaskName As String
        Dim WfProjectTaskAttendee As New WfProjectTaskAttendee(conn, ts)
        Dim WfProjectTrack As New WfProjectTrack(conn, ts)
        Dim WfProjectTask As New WfProjectTask(conn, ts)
        Dim dsReturn As New DataSet
        Dim dsProjectTrack, dsTask As DataSet

        If GetTaskProjectListCommand Is Nothing Then

            GetTaskProjectListCommand = New SqlCommand("GetTaskProjectList", conn)
            GetTaskProjectListCommand.CommandType = CommandType.StoredProcedure
            GetTaskProjectListCommand.Parameters.Add(New SqlParameter("@TaskID", SqlDbType.NVarChar))
            GetTaskProjectListCommand.Parameters.Add(New SqlParameter("@UserName", SqlDbType.NVarChar))
            GetTaskProjectListCommand.Parameters.Add(New SqlParameter("@flag", SqlDbType.Int))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetTaskProjectListCommand
            .SelectCommand.Transaction = ts
            GetTaskProjectListCommand.Parameters("@TaskID").Value = taskID
            GetTaskProjectListCommand.Parameters("@UserName").Value = userName
            GetTaskProjectListCommand.Parameters("@flag").Value = flag
            Try
                .Fill(dsReturn)
            Catch e As Exception
                MsgBox(e.Message)
            End Try

        End With

        Return dsReturn

    End Function

    ''获取任务的项目列表信息
    'Public Function GetTaskProjectList(ByVal taskID As String, ByVal userName As String, ByVal flag As Integer) As DataSet
    '    Dim i, j As Integer
    '    Dim strSql As String
    '    Dim tmpRoleName, tmpProjectCode, tmpRoleID, tmpAttend, tmpRollBackTask, tmpRollBackTaskName As String
    '    Dim WfProjectTaskAttendee As New WfProjectTaskAttendee(conn, ts)
    '    Dim WfProjectTrack As New WfProjectTrack(conn, ts)
    '    Dim WfProjectTask As New WfProjectTask(conn, ts)
    '    Dim Role As New Role(conn, ts)
    '    Dim dsTempAttend, dsRole, dsReturn, dsProjectTrack, dsTask As DataSet
    '    strSql = "SELECT dbo.project_task_attendee.project_code," & _
    '             " dbo.project_task_attendee.role_id,dbo.project_task_attendee.attend_person" & _
    '             " FROM dbo.project_task INNER JOIN" & _
    '             " dbo.project_task_attendee ON " & _
    '             " dbo.project_task.project_code = dbo.project_task_attendee.project_code AND" & _
    '             " dbo.project_task.workflow_id = dbo.project_task_attendee.workflow_id AND" & _
    '             " dbo.project_task.task_id = dbo.project_task_attendee.task_id" & _
    '             " WHERE (dbo.project_task.start_time IS NOT NULL)" & _
    '             " order by dbo.project_task.start_time"
    '    dsTempAttend = GetCommonQueryInfo(strSql)
    '    If flag = 0 Then
    '        strSql = " SELECT distinct v.ProjectCode,v.EnterpriseName,v.phase,v.status,v.is_check_record,v.risk_grade,v.hasAppraised, v.serviceType,t.task_id, " & _
    '                 " dbo.project_task.task_name, dbo.project_task.start_time,dbo.project_task.apply_tool,dbo.project_task.start_mode, " & _
    '                 " ISNULL(dbo.project_timing_task.time_limit, 0) AS time_limit,v.ServiceType " & _
    '                 " FROM dbo.ViewProjectInfo v INNER JOIN" & _
    '                 " dbo.project_task_attendee t ON " & _
    '                 " v.ProjectCode = t.project_code INNER JOIN" & _
    '                 " dbo.project_task ON " & _
    '                 " t.project_code = dbo.project_task.project_code AND " & _
    '                 " t.task_id = dbo.project_task.task_id LEFT OUTER JOIN" & _
    '                 " dbo.project_timing_task ON " & _
    '                 " t.role_id = dbo.project_timing_task.role_id AND " & _
    '                 " t.task_id = dbo.project_timing_task.task_id AND " & _
    '                 " t.project_code = dbo.project_timing_task.project_code " & _
    '                 " WHERE t.task_status = 'P'" & _
    '                 " AND t.task_id =" & "'" & taskID & "' " & _
    '                 " AND t.attend_person=" & "'" & userName & "'  order by v.EnterpriseName"
    '    Else
    '        strSql = " SELECT distinct v.ProjectCode,v.EnterpriseName,v.phase,v.status,v.is_check_record,v.risk_grade,v.hasAppraised, v.serviceType, t.task_id, " & _
    '                 " dbo.project_task.task_name, dbo.project_task.start_time,dbo.project_task.apply_tool,dbo.project_task.start_mode, " & _
    '                 " ISNULL(dbo.project_timing_task.time_limit, 0) AS time_limit,v.ServiceType " & _
    '                 " FROM dbo.ViewProjectInfo v INNER JOIN" & _
    '                 " dbo.project_task_attendee t ON " & _
    '                 " v.ProjectCode = t.project_code INNER JOIN" & _
    '                 " dbo.project_task ON " & _
    '                 " t.project_code = dbo.project_task.project_code AND " & _
    '                 " t.task_id = dbo.project_task.task_id LEFT OUTER JOIN" & _
    '                 " dbo.project_timing_task ON " & _
    '                 " t.role_id = dbo.project_timing_task.role_id AND " & _
    '                 " t.task_id = dbo.project_timing_task.task_id AND " & _
    '                 " t.project_code = dbo.project_timing_task.project_code " & _
    '                 " WHERE t.task_status = 'P'" & _
    '                 " AND t.task_id =" & "'" & taskID & "' " & _
    '                 " AND t.attend_person=" & "'" & userName & "'  order by v.is_check_record desc ,v.risk_grade desc,v.hasAppraised asc"
    '        '" AND dbo.project_task_attendee.attend_person=" & "'" & userName & "' order by dbo.ViewProjectInfo.is_check_record,dbo.ViewProjectInfo.risk_grade"
    '    End If
    '    dsReturn = GetCommonQueryInfo(strSql)

    '    '不取咨询人员
    '    strSql = "{role_id<>'11'}"
    '    dsRole = Role.FetchRole(strSql)
    '    Dim iCount As Integer = dsReturn.Tables(0).Columns.Count

    '    '增加所有角色字段
    '    For i = 0 To dsRole.Tables(0).Rows.Count - 1
    '        tmpRoleName = Trim(dsRole.Tables(0).Rows(i).Item("role_id"))
    '        dsReturn.Tables(0).Columns.Add(tmpRoleName, GetType(System.String))
    '    Next


    '    '增加前置任务ID和名称,完成人字段
    '    dsReturn.Tables(0).Columns.Add("previous_task_id", GetType(System.String))
    '    dsReturn.Tables(0).Columns.Add("previous_task_name", GetType(System.String))
    '    dsReturn.Tables(0).Columns.Add("previous_task_attendee", GetType(System.String))


    '    For i = 0 To dsReturn.Tables(0).Rows.Count - 1
    '        tmpProjectCode = dsReturn.Tables(0).Rows(i).Item("projectcode")

    '        '填充该项目的所有角色参与人
    '        For j = 0 To dsRole.Tables(0).Rows.Count - 1
    '            tmpRoleID = dsRole.Tables(0).Rows(j).Item("role_id")

    '            strSql = "project_code=" & "'" & tmpProjectCode & "'" & " and role_id=" & "'" & tmpRoleID & "'"
    '            If dsTempAttend.Tables(0).Select(strSql).Length() = 0 Then
    '                tmpAttend = ""
    '            Else
    '                tmpAttend = dsTempAttend.Tables(0).Select(strSql)(0).Item("attend_person")
    '            End If

    '            '填充项目参与人
    '            dsReturn.Tables(0).Rows(i).Item(j + iCount) = tmpAttend

    '        Next

    '        '获取工作流ID、StartupTask= TaskID、Status=“P”的Project_Track对象；
    '        strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and StartupTask=" & "'" & taskID & "'" & " and isnull(status,'')='P'}"
    '        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)

    '        '获取前置任务ID；
    '        If dsProjectTrack.Tables(0).Rows.Count = 0 Then
    '            tmpRollBackTask = ""
    '            tmpRollBackTaskName = ""
    '            tmpAttend = ""
    '        Else
    '            tmpRollBackTask = Trim(dsProjectTrack.Tables(0).Rows(0).Item("FinishedTask"))

    '            '获取前置任务的名称
    '            Dim worklog As New WorkLog(conn, ts)
    '            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id=" & "'" & tmpRollBackTask & "'" & "}"
    '            dsTask = worklog.GetWorkLogInfo(strSql)
    '            'dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
    '            If dsTask.Tables(0).Rows.Count <> 0 Then
    '                tmpRollBackTaskName = Trim(dsTask.Tables(0).Rows(0).Item("task_name"))
    '                tmpAttend = Trim(dsTask.Tables(0).Rows(0).Item("attend_person"))
    '            Else
    '                tmpRollBackTaskName = ""
    '                tmpAttend = ""
    '            End If

    '        End If


    '        dsReturn.Tables(0).Rows(i).Item("previous_task_id") = tmpRollBackTask
    '        dsReturn.Tables(0).Rows(i).Item("previous_task_name") = tmpRollBackTaskName
    '        dsReturn.Tables(0).Rows(i).Item("previous_task_attendee") = tmpAttend

    '    Next

    '    Return dsReturn
    'End Function

    Public Function GetConferenceProjectList(ByVal userName As String) As DataSet
        Dim dsTemp As DataSet
        dsTemp = GetTaskProjectList("ReviewMeetingPlan", userName)
        dsTemp.Merge(GetTaskProjectList("ReviewMeetingPlanExp", userName))
        Return dsTemp
    End Function



    Public Function FetchFinancialAnalysisInfo(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal ThisYear As String, ByVal LastYear1 As String, ByVal LastYear2 As String, ByVal LastYear3 As String) As DataSet
        Dim tempDs As New DataSet

        If FetchFinancialAnalysisInfoCommand Is Nothing Then
            FetchFinancialAnalysisInfoCommand = New SqlCommand("PFetchFinancialAnalysis", conn)
            FetchFinancialAnalysisInfoCommand.CommandType = CommandType.StoredProcedure
            FetchFinancialAnalysisInfoCommand.Parameters.Add(New SqlParameter("@ProjectNo", SqlDbType.NVarChar))
            FetchFinancialAnalysisInfoCommand.Parameters.Add(New SqlParameter("@CorporationNo", SqlDbType.NVarChar))
            FetchFinancialAnalysisInfoCommand.Parameters.Add(New SqlParameter("@Phase", SqlDbType.NVarChar))
            FetchFinancialAnalysisInfoCommand.Parameters.Add(New SqlParameter("@ThisYear", SqlDbType.NVarChar))
            FetchFinancialAnalysisInfoCommand.Parameters.Add(New SqlParameter("@LastYear1", SqlDbType.NVarChar))
            FetchFinancialAnalysisInfoCommand.Parameters.Add(New SqlParameter("@LastYear2", SqlDbType.NVarChar))
            FetchFinancialAnalysisInfoCommand.Parameters.Add(New SqlParameter("@LastYear3", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = FetchFinancialAnalysisInfoCommand
            .SelectCommand.Transaction = ts
            FetchFinancialAnalysisInfoCommand.Parameters("@ProjectNo").Value = ProjectNo
            FetchFinancialAnalysisInfoCommand.Parameters("@CorporationNo").Value = CorporationNo
            FetchFinancialAnalysisInfoCommand.Parameters("@Phase").Value = Phase
            FetchFinancialAnalysisInfoCommand.Parameters("@ThisYear").Value = ThisYear
            FetchFinancialAnalysisInfoCommand.Parameters("@LastYear1").Value = LastYear1
            FetchFinancialAnalysisInfoCommand.Parameters("@LastYear2").Value = LastYear2
            FetchFinancialAnalysisInfoCommand.Parameters("@LastYear3").Value = LastYear3
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function GetNeedSignatureProjectInfo(ByVal Condition As String) As DataSet
        Dim tempDs As New DataSet

        If GetNeedSignatureProjectInfoCommand Is Nothing Then
            GetNeedSignatureProjectInfoCommand = New SqlCommand("dt_lqfGetNeedSignatureProjectInfo", conn)
            GetNeedSignatureProjectInfoCommand.CommandType = CommandType.StoredProcedure
            GetNeedSignatureProjectInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetNeedSignatureProjectInfoCommand
            .SelectCommand.Transaction = ts
            GetNeedSignatureProjectInfoCommand.Parameters("@Condition").Value = Condition

            .Fill(tempDs)
        End With

        Return tempDs
    End Function


    Public Function FetchOppositeGuaranteeAssurer(ByVal Condition As String) As DataSet
        Dim tempDs As New DataSet

        If FetchOppositeGuaranteeAssurerCommand Is Nothing Then
            FetchOppositeGuaranteeAssurerCommand = New SqlCommand("PFetchOppositeGuaranteeAssurer", conn)
            FetchOppositeGuaranteeAssurerCommand.CommandType = CommandType.StoredProcedure
            FetchOppositeGuaranteeAssurerCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = FetchOppositeGuaranteeAssurerCommand
            .SelectCommand.Transaction = ts
            FetchOppositeGuaranteeAssurerCommand.Parameters("@Condition").Value = Condition
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FetchProjectGuaranteeForm(ByVal Condition As String) As DataSet
        Dim tempDs As New DataSet

        If FetchProjectGuaranteeFormCommand Is Nothing Then
            FetchProjectGuaranteeFormCommand = New SqlCommand("PFetchProjectGuaranteeForm", conn)
            FetchProjectGuaranteeFormCommand.CommandType = CommandType.StoredProcedure
            FetchProjectGuaranteeFormCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = FetchProjectGuaranteeFormCommand
            .SelectCommand.Transaction = ts
            FetchProjectGuaranteeFormCommand.Parameters("@Condition").Value = Condition
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function GetAcceptVouchData(ByVal ProjectCode As String) As DataSet
        Dim tempDs As New DataSet

        If GetAcceptVouchDataCommand Is Nothing Then
            GetAcceptVouchDataCommand = New SqlCommand("dt_xjGetAcceptVouchData", conn)
            GetAcceptVouchDataCommand.CommandType = CommandType.StoredProcedure
            GetAcceptVouchDataCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetAcceptVouchDataCommand
            .SelectCommand.Transaction = ts
            GetAcceptVouchDataCommand.Parameters("@ProjectCode").Value = ProjectCode
            .Fill(tempDs)
        End With

        Return tempDs
    End Function


    Public Function GetTaskListInfo(ByVal Condition As String) As DataSet
        Dim tempDs As New DataSet

        If GetTaskListInfoCommand Is Nothing Then
            GetTaskListInfoCommand = New SqlCommand("GetTaskListInfo", conn)
            GetTaskListInfoCommand.CommandType = CommandType.StoredProcedure
            GetTaskListInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetTaskListInfoCommand
            .SelectCommand.Transaction = ts
            GetTaskListInfoCommand.Parameters("@Condition").Value = Condition
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function GetRefundProcess(ByVal projectcode As String) As DataSet
        Dim tempDs As New DataSet

        If GetRefundProcessCommand Is Nothing Then
            GetRefundProcessCommand = New SqlCommand("dt_xjGetRefundProcess", conn)
            GetRefundProcessCommand.CommandType = CommandType.StoredProcedure
            GetRefundProcessCommand.Parameters.Add(New SqlParameter("@projectcode", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetRefundProcessCommand
            .SelectCommand.Transaction = ts
            GetRefundProcessCommand.Parameters("@projectcode").Value = projectcode
            .Fill(tempDs)
        End With

        Return tempDs
    End Function


    '获取初审人员初审任务的任务量()
    Public Function GetReviewListInfo() As DataSet
        Return GetTaskListInfo("project_task.task_id='Review'")
    End Function

    '获取法物经理制作合同任务量
    Public Function GetDraftOutContractListInfo() As DataSet
        Return GetTaskListInfo("project_task.task_id='DraftOutContract'")
    End Function

    '获取资产评估师资产评估任务的任务量
    Public Function GetCapitialEvaluatedListInfo() As DataSet
        Return GetTaskListInfo("project_task.task_id='CapitialEvaluated'")
    End Function

    '获取项目经理项目评审任务的任务量
    Public Function GetManagerAppraiseListInfo() As DataSet
        Return GetTaskListInfo("project_task.task_id='ProjectAppraiseReport'")
    End Function

    '获取项目组所有项目经理项目评审任务的任务量
    Public Function GetTeamAppraiseListInfo() As DataSet
        Return GetAppraiseListInfo("project_task.task_id='ProjectAppraiseReport'")
    End Function

    '获取项目组所有项目经理的任务量
    Public Function GetAppraiseListInfo(ByVal Condition) As DataSet
        Dim tempDs As New DataSet

        If GetTaskListInfoCommand Is Nothing Then
            GetTaskListInfoCommand = New SqlCommand("GetAppraiseListInfo", conn)
            GetTaskListInfoCommand.CommandType = CommandType.StoredProcedure
            GetTaskListInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetTaskListInfoCommand
            .SelectCommand.Transaction = ts
            GetTaskListInfoCommand.Parameters("@Condition").Value = Condition
            .Fill(tempDs)
        End With

        Return tempDs
    End Function


    Public Function FQueryAcceptProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal apply_service_type As String, ByVal accept_date_start As String, ByVal accept_date_end As String, ByVal apply_bank As String, ByVal belong_area As String, ByVal vchAcceptBranch As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryAcceptProjectCommand Is Nothing Then
            FQueryAcceptProjectCommand = New SqlCommand("FQueryAcceptProject", conn)
            FQueryAcceptProjectCommand.CommandType = CommandType.StoredProcedure
            FQueryAcceptProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryAcceptProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryAcceptProjectCommand.Parameters.Add(New SqlParameter("@apply_service_type", SqlDbType.NVarChar))
            FQueryAcceptProjectCommand.Parameters.Add(New SqlParameter("@accept_date_start", SqlDbType.DateTime))
            FQueryAcceptProjectCommand.Parameters.Add(New SqlParameter("@accept_date_end", SqlDbType.DateTime))
            FQueryAcceptProjectCommand.Parameters.Add(New SqlParameter("@apply_bank", SqlDbType.NVarChar))
            FQueryAcceptProjectCommand.Parameters.Add(New SqlParameter("@belong_area", SqlDbType.NVarChar))
            FQueryAcceptProjectCommand.Parameters.Add(New SqlParameter("@vchAcceptBranch", SqlDbType.NVarChar))
            FQueryAcceptProjectCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.NVarChar))
            FQueryAcceptProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryAcceptProjectCommand
            .SelectCommand.Transaction = ts
            FQueryAcceptProjectCommand.Parameters("@project_code").Value = project_code
            FQueryAcceptProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryAcceptProjectCommand.Parameters("@apply_service_type").Value = apply_service_type
            FQueryAcceptProjectCommand.Parameters("@accept_date_start").Value = accept_date_start
            FQueryAcceptProjectCommand.Parameters("@accept_date_end").Value = accept_date_end
            FQueryAcceptProjectCommand.Parameters("@apply_bank").Value = apply_bank
            FQueryAcceptProjectCommand.Parameters("@belong_area").Value = belong_area
            FQueryAcceptProjectCommand.Parameters("@vchAcceptBranch").Value = vchAcceptBranch
            FQueryAcceptProjectCommand.Parameters("@vchPMA").Value = vchPMA
            FQueryAcceptProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryAllocateProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal apply_service_type As String, ByVal assign_date_start As String, ByVal assign_date_end As String, ByVal manager_a As String, ByVal manager_b As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryAllocateProjectCommand Is Nothing Then
            FQueryAllocateProjectCommand = New SqlCommand("FQueryAllocateProject", conn)
            FQueryAllocateProjectCommand.CommandType = CommandType.StoredProcedure
            FQueryAllocateProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryAllocateProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryAllocateProjectCommand.Parameters.Add(New SqlParameter("@apply_service_type", SqlDbType.NVarChar))
            FQueryAllocateProjectCommand.Parameters.Add(New SqlParameter("@assign_date_start", SqlDbType.DateTime))
            FQueryAllocateProjectCommand.Parameters.Add(New SqlParameter("@assign_date_end", SqlDbType.DateTime))
            FQueryAllocateProjectCommand.Parameters.Add(New SqlParameter("@manager_a", SqlDbType.NVarChar))
            FQueryAllocateProjectCommand.Parameters.Add(New SqlParameter("@manager_b", SqlDbType.NVarChar))
            FQueryAllocateProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryAllocateProjectCommand
            .SelectCommand.Transaction = ts
            FQueryAllocateProjectCommand.Parameters("@project_code").Value = project_code
            FQueryAllocateProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryAllocateProjectCommand.Parameters("@apply_service_type").Value = apply_service_type
            FQueryAllocateProjectCommand.Parameters("@assign_date_start").Value = assign_date_start
            FQueryAllocateProjectCommand.Parameters("@assign_date_end").Value = assign_date_end
            FQueryAllocateProjectCommand.Parameters("@manager_a").Value = manager_a
            FQueryAllocateProjectCommand.Parameters("@manager_b").Value = manager_b
            FQueryAllocateProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryPresentingProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal apply_service_type As String, ByVal evial_date_start As String, ByVal evial_date_end As String, ByVal belong_district As String, ByVal belong_trade As String, ByVal ownership_type As String, ByVal team_name As String, ByVal manager_a As String, ByVal evial_conclusion As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryPresentingProjectCommand Is Nothing Then
            FQueryPresentingProjectCommand = New SqlCommand("FQueryPresentingProject", conn)
            FQueryPresentingProjectCommand.CommandType = CommandType.StoredProcedure
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@apply_service_type", SqlDbType.NVarChar))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@evial_date_start", SqlDbType.DateTime))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@evial_date_end", SqlDbType.DateTime))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@belong_district", SqlDbType.NVarChar))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@belong_trade", SqlDbType.NVarChar))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@ownership_type", SqlDbType.NVarChar))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@team_name", SqlDbType.NVarChar))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@manager_a", SqlDbType.NVarChar))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@evial_conclusion", SqlDbType.NVarChar))
            FQueryPresentingProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryPresentingProjectCommand
            .SelectCommand.Transaction = ts
            FQueryPresentingProjectCommand.Parameters("@project_code").Value = project_code
            FQueryPresentingProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryPresentingProjectCommand.Parameters("@apply_service_type").Value = apply_service_type
            FQueryPresentingProjectCommand.Parameters("@evial_date_start").Value = evial_date_start
            FQueryPresentingProjectCommand.Parameters("@evial_date_end").Value = evial_date_end
            FQueryPresentingProjectCommand.Parameters("@belong_district").Value = belong_district
            FQueryPresentingProjectCommand.Parameters("@belong_trade").Value = belong_trade
            FQueryPresentingProjectCommand.Parameters("@ownership_type").Value = ownership_type
            FQueryPresentingProjectCommand.Parameters("@team_name").Value = team_name
            FQueryPresentingProjectCommand.Parameters("@manager_a").Value = manager_a
            FQueryPresentingProjectCommand.Parameters("@evial_conclusion").Value = evial_conclusion
            FQueryPresentingProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryLoanProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal loan_date_start As String, ByVal loan_date_end As String, ByVal manager_a As String, ByVal bank As String, ByVal branch_bank As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryLoanProjectCommand Is Nothing Then
            FQueryLoanProjectCommand = New SqlCommand("FQueryLoanProject", conn)
            FQueryLoanProjectCommand.CommandType = CommandType.StoredProcedure
            FQueryLoanProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryLoanProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryLoanProjectCommand.Parameters.Add(New SqlParameter("@service_type", SqlDbType.NVarChar))
            FQueryLoanProjectCommand.Parameters.Add(New SqlParameter("@loan_date_start", SqlDbType.DateTime))
            FQueryLoanProjectCommand.Parameters.Add(New SqlParameter("@loan_date_end", SqlDbType.DateTime))
            FQueryLoanProjectCommand.Parameters.Add(New SqlParameter("@manager_a", SqlDbType.NVarChar))
            FQueryLoanProjectCommand.Parameters.Add(New SqlParameter("@bank", SqlDbType.NVarChar))
            FQueryLoanProjectCommand.Parameters.Add(New SqlParameter("@branch_bank", SqlDbType.NVarChar))
            FQueryLoanProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryLoanProjectCommand
            .SelectCommand.Transaction = ts
            FQueryLoanProjectCommand.Parameters("@project_code").Value = project_code
            FQueryLoanProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryLoanProjectCommand.Parameters("@service_type").Value = service_type
            FQueryLoanProjectCommand.Parameters("@loan_date_start").Value = loan_date_start
            FQueryLoanProjectCommand.Parameters("@loan_date_end").Value = loan_date_end
            FQueryLoanProjectCommand.Parameters("@manager_a").Value = manager_a
            FQueryLoanProjectCommand.Parameters("@bank").Value = bank
            FQueryLoanProjectCommand.Parameters("@branch_bank").Value = branch_bank
            FQueryLoanProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryProjectExpandDate(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal loan_date_start As String, ByVal loan_date_end As String, ByVal manager_a As String, ByVal bank As String, ByVal branch_bank As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet
        If FQueryProjectExpandDateCommand Is Nothing Then
            FQueryProjectExpandDateCommand = New SqlCommand("FQueryProjectExpendDate", conn)
            FQueryProjectExpandDateCommand.CommandType = CommandType.StoredProcedure
            FQueryProjectExpandDateCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryProjectExpandDateCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryProjectExpandDateCommand.Parameters.Add(New SqlParameter("@service_type", SqlDbType.NVarChar))
            FQueryProjectExpandDateCommand.Parameters.Add(New SqlParameter("@loan_date_start", SqlDbType.DateTime))
            FQueryProjectExpandDateCommand.Parameters.Add(New SqlParameter("@loan_date_end", SqlDbType.DateTime))
            FQueryProjectExpandDateCommand.Parameters.Add(New SqlParameter("@manager_a", SqlDbType.NVarChar))
            FQueryProjectExpandDateCommand.Parameters.Add(New SqlParameter("@bank", SqlDbType.NVarChar))
            FQueryProjectExpandDateCommand.Parameters.Add(New SqlParameter("@branch_bank", SqlDbType.NVarChar))
            FQueryProjectExpandDateCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryProjectExpandDateCommand
            .SelectCommand.Transaction = ts
            FQueryProjectExpandDateCommand.Parameters("@project_code").Value = project_code
            FQueryProjectExpandDateCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryProjectExpandDateCommand.Parameters("@service_type").Value = service_type
            FQueryProjectExpandDateCommand.Parameters("@loan_date_start").Value = loan_date_start
            FQueryProjectExpandDateCommand.Parameters("@loan_date_end").Value = loan_date_end
            FQueryProjectExpandDateCommand.Parameters("@manager_a").Value = manager_a
            FQueryProjectExpandDateCommand.Parameters("@bank").Value = bank
            FQueryProjectExpandDateCommand.Parameters("@branch_bank").Value = branch_bank
            FQueryProjectExpandDateCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQuerySignProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal sign_date_start As String, ByVal sign_date_end As String, ByVal belong_district As String, ByVal belong_trade As String, ByVal ownership_type As String, ByVal manager_a As String, ByVal bank As String, ByVal branch_bank As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQuerySignProjectCommand Is Nothing Then
            FQuerySignProjectCommand = New SqlCommand("FQuerySignProject", conn)
            FQuerySignProjectCommand.CommandType = CommandType.StoredProcedure
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@service_type", SqlDbType.NVarChar))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@sign_date_start", SqlDbType.DateTime))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@sign_date_end", SqlDbType.DateTime))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@belong_district", SqlDbType.NVarChar))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@belong_trade", SqlDbType.NVarChar))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@ownership_type", SqlDbType.NVarChar))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@manager_a", SqlDbType.NVarChar))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@bank", SqlDbType.NVarChar))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@branch_bank", SqlDbType.NVarChar))
            FQuerySignProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQuerySignProjectCommand
            .SelectCommand.Transaction = ts
            FQuerySignProjectCommand.Parameters("@project_code").Value = project_code
            FQuerySignProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQuerySignProjectCommand.Parameters("@service_type").Value = service_type
            FQuerySignProjectCommand.Parameters("@sign_date_start").Value = sign_date_start
            FQuerySignProjectCommand.Parameters("@sign_date_end").Value = sign_date_end
            FQuerySignProjectCommand.Parameters("@belong_district").Value = belong_district
            FQuerySignProjectCommand.Parameters("@belong_trade").Value = belong_trade
            FQuerySignProjectCommand.Parameters("@ownership_type").Value = ownership_type
            FQuerySignProjectCommand.Parameters("@manager_a").Value = manager_a
            FQuerySignProjectCommand.Parameters("@bank").Value = bank
            FQuerySignProjectCommand.Parameters("@branch_bank").Value = branch_bank
            FQuerySignProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PQueryFirstTrialProject(ByVal project_code As String, ByVal enterprise_name As String, _
        ByVal apply_service_type As String, ByVal accept_date_start As String, ByVal accept_date_end As String, _
        ByVal apply_bank As String, ByVal belong_area As String, ByVal belong_trade As String, _
        ByVal ownership_type As String, ByVal vchAcceptBranch As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If PQueryFirstTrialProjectCommand Is Nothing Then
            PQueryFirstTrialProjectCommand = New SqlCommand("PQueryFirstTrialProject", conn)
            PQueryFirstTrialProjectCommand.CommandType = CommandType.StoredProcedure
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@apply_service_type", SqlDbType.NVarChar))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@accept_date_start", SqlDbType.DateTime))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@accept_date_end", SqlDbType.DateTime))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@apply_bank", SqlDbType.NVarChar))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@belong_area", SqlDbType.NVarChar))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@belong_trade", SqlDbType.NVarChar))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@ownership_type", SqlDbType.NVarChar))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@vchAcceptBranch", SqlDbType.NVarChar))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.NVarChar))
            PQueryFirstTrialProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = PQueryFirstTrialProjectCommand
            .SelectCommand.Transaction = ts
            PQueryFirstTrialProjectCommand.Parameters("@project_code").Value = project_code
            PQueryFirstTrialProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            PQueryFirstTrialProjectCommand.Parameters("@apply_service_type").Value = apply_service_type
            PQueryFirstTrialProjectCommand.Parameters("@accept_date_start").Value = accept_date_start
            PQueryFirstTrialProjectCommand.Parameters("@accept_date_end").Value = accept_date_end
            PQueryFirstTrialProjectCommand.Parameters("@apply_bank").Value = apply_bank
            PQueryFirstTrialProjectCommand.Parameters("@belong_area").Value = belong_area
            PQueryFirstTrialProjectCommand.Parameters("@belong_trade").Value = belong_trade
            PQueryFirstTrialProjectCommand.Parameters("@ownership_type").Value = ownership_type
            PQueryFirstTrialProjectCommand.Parameters("@vchAcceptBranch").Value = vchAcceptBranch
            PQueryFirstTrialProjectCommand.Parameters("@vchPMA").Value = vchPMA
            PQueryFirstTrialProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryRequiteProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal refund_date_start As String, ByVal refund_date_end As String, ByVal manager_a As String, ByVal bank As String, ByVal refund_type As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryRequiteProjectCommand Is Nothing Then
            FQueryRequiteProjectCommand = New SqlCommand("FQueryRequiteProject", conn)
            FQueryRequiteProjectCommand.CommandType = CommandType.StoredProcedure
            FQueryRequiteProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryRequiteProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryRequiteProjectCommand.Parameters.Add(New SqlParameter("@service_type", SqlDbType.NVarChar))
            FQueryRequiteProjectCommand.Parameters.Add(New SqlParameter("@refund_date_start", SqlDbType.DateTime))
            FQueryRequiteProjectCommand.Parameters.Add(New SqlParameter("@refund_date_end", SqlDbType.DateTime))
            FQueryRequiteProjectCommand.Parameters.Add(New SqlParameter("@manager_a", SqlDbType.NVarChar))
            FQueryRequiteProjectCommand.Parameters.Add(New SqlParameter("@bank", SqlDbType.NVarChar))
            FQueryRequiteProjectCommand.Parameters.Add(New SqlParameter("@refund_type", SqlDbType.NVarChar))
            FQueryRequiteProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryRequiteProjectCommand
            .SelectCommand.Transaction = ts
            FQueryRequiteProjectCommand.Parameters("@project_code").Value = project_code
            FQueryRequiteProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryRequiteProjectCommand.Parameters("@service_type").Value = service_type
            FQueryRequiteProjectCommand.Parameters("@refund_date_start").Value = refund_date_start
            FQueryRequiteProjectCommand.Parameters("@refund_date_end").Value = refund_date_end
            FQueryRequiteProjectCommand.Parameters("@manager_a").Value = manager_a
            FQueryRequiteProjectCommand.Parameters("@bank").Value = bank
            FQueryRequiteProjectCommand.Parameters("@refund_type").Value = refund_type
            FQueryRequiteProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryCreditProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal today_date As String, ByVal manager_a As String, ByVal bank As String, ByVal branch_bank As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryCreditProjectCommand Is Nothing Then
            FQueryCreditProjectCommand = New SqlCommand("FQueryCreditProject", conn)
            FQueryCreditProjectCommand.CommandType = CommandType.StoredProcedure
            FQueryCreditProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryCreditProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryCreditProjectCommand.Parameters.Add(New SqlParameter("@service_type", SqlDbType.NVarChar))
            FQueryCreditProjectCommand.Parameters.Add(New SqlParameter("@today_date", SqlDbType.DateTime))
            FQueryCreditProjectCommand.Parameters.Add(New SqlParameter("@manager_a", SqlDbType.NVarChar))
            FQueryCreditProjectCommand.Parameters.Add(New SqlParameter("@bank", SqlDbType.NVarChar))
            FQueryCreditProjectCommand.Parameters.Add(New SqlParameter("@branch_bank", SqlDbType.NVarChar))
            FQueryCreditProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryCreditProjectCommand
            .SelectCommand.Transaction = ts
            FQueryCreditProjectCommand.Parameters("@project_code").Value = project_code
            FQueryCreditProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryCreditProjectCommand.Parameters("@service_type").Value = service_type
            FQueryCreditProjectCommand.Parameters("@today_date").Value = today_date
            FQueryCreditProjectCommand.Parameters("@manager_a").Value = manager_a
            FQueryCreditProjectCommand.Parameters("@bank").Value = bank
            FQueryCreditProjectCommand.Parameters("@branch_bank").Value = branch_bank
            FQueryCreditProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryRecantProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal date_start As String, ByVal date_end As String, ByVal manager_a As String, ByVal bank As String, ByVal branch_bank As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryRecantProjectCommand Is Nothing Then
            FQueryRecantProjectCommand = New SqlCommand("FQueryRecantProject", conn)
            FQueryRecantProjectCommand.CommandType = CommandType.StoredProcedure
            FQueryRecantProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryRecantProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryRecantProjectCommand.Parameters.Add(New SqlParameter("@service_type", SqlDbType.NVarChar))
            FQueryRecantProjectCommand.Parameters.Add(New SqlParameter("@date_start", SqlDbType.DateTime))
            FQueryRecantProjectCommand.Parameters.Add(New SqlParameter("@date_end", SqlDbType.DateTime))
            FQueryRecantProjectCommand.Parameters.Add(New SqlParameter("@manager_a", SqlDbType.NVarChar))
            FQueryRecantProjectCommand.Parameters.Add(New SqlParameter("@bank", SqlDbType.NVarChar))
            FQueryRecantProjectCommand.Parameters.Add(New SqlParameter("@branch_bank", SqlDbType.NVarChar))
            FQueryRecantProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryRecantProjectCommand
            .SelectCommand.Transaction = ts
            FQueryRecantProjectCommand.Parameters("@project_code").Value = project_code
            FQueryRecantProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryRecantProjectCommand.Parameters("@service_type").Value = service_type
            FQueryRecantProjectCommand.Parameters("@date_start").Value = date_start
            FQueryRecantProjectCommand.Parameters("@date_end").Value = date_end
            FQueryRecantProjectCommand.Parameters("@manager_a").Value = manager_a
            FQueryRecantProjectCommand.Parameters("@bank").Value = bank
            FQueryRecantProjectCommand.Parameters("@branch_bank").Value = branch_bank
            FQueryRecantProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryProcessingProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal date_start As String, ByVal manager_a As String, ByVal manager_b As String, ByVal phase As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryProcessingProjectCommand Is Nothing Then
            FQueryProcessingProjectCommand = New SqlCommand("FQueryProcessingProject", conn)
            FQueryProcessingProjectCommand.CommandType = CommandType.StoredProcedure
            FQueryProcessingProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryProcessingProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryProcessingProjectCommand.Parameters.Add(New SqlParameter("@service_type", SqlDbType.NVarChar))
            FQueryProcessingProjectCommand.Parameters.Add(New SqlParameter("@date_start", SqlDbType.DateTime))
            FQueryProcessingProjectCommand.Parameters.Add(New SqlParameter("@manager_a", SqlDbType.NVarChar))
            FQueryProcessingProjectCommand.Parameters.Add(New SqlParameter("@manager_b", SqlDbType.NVarChar))
            FQueryProcessingProjectCommand.Parameters.Add(New SqlParameter("@phase", SqlDbType.NVarChar))
            FQueryProcessingProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryProcessingProjectCommand
            .SelectCommand.Transaction = ts
            FQueryProcessingProjectCommand.Parameters("@project_code").Value = project_code
            FQueryProcessingProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryProcessingProjectCommand.Parameters("@service_type").Value = service_type
            FQueryProcessingProjectCommand.Parameters("@date_start").Value = date_start
            FQueryProcessingProjectCommand.Parameters("@manager_a").Value = manager_a
            FQueryProcessingProjectCommand.Parameters("@manager_b").Value = manager_b
            FQueryProcessingProjectCommand.Parameters("@phase").Value = phase
            FQueryProcessingProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryRegionProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal date_start As String, ByVal date_end As String, ByVal cooperate_area As String, ByVal phase As String, ByVal vchPMA As String, ByVal userName As String, ByVal recommend_type As String, ByVal opinion As String, ByVal exempt As String, ByVal trial_conclusion As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryRegionProjectCommand Is Nothing Then
            FQueryRegionProjectCommand = New SqlCommand("FQueryRegionProject", conn)
            FQueryRegionProjectCommand.CommandType = CommandType.StoredProcedure
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@service_type", SqlDbType.NVarChar))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@date_start", SqlDbType.DateTime))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@date_end", SqlDbType.DateTime))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@cooperate_area", SqlDbType.NVarChar))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@phase", SqlDbType.NVarChar))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.NVarChar))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@recommend_type", SqlDbType.NVarChar))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@opinion", SqlDbType.NVarChar))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@exempt", SqlDbType.Int))
            FQueryRegionProjectCommand.Parameters.Add(New SqlParameter("@trial_conclusion", SqlDbType.NVarChar))


        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryRegionProjectCommand
            .SelectCommand.Transaction = ts
            FQueryRegionProjectCommand.Parameters("@project_code").Value = project_code
            FQueryRegionProjectCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryRegionProjectCommand.Parameters("@service_type").Value = service_type
            FQueryRegionProjectCommand.Parameters("@date_start").Value = date_start
            FQueryRegionProjectCommand.Parameters("@date_end").Value = date_end
            FQueryRegionProjectCommand.Parameters("@cooperate_area").Value = cooperate_area
            FQueryRegionProjectCommand.Parameters("@phase").Value = phase
            FQueryRegionProjectCommand.Parameters("@vchPMA").Value = vchPMA
            FQueryRegionProjectCommand.Parameters("@userName").Value = userName
            FQueryRegionProjectCommand.Parameters("@recommend_type").Value = recommend_type
            FQueryRegionProjectCommand.Parameters("@opinion").Value = opinion
            FQueryRegionProjectCommand.Parameters("@exempt").Value = exempt
            FQueryRegionProjectCommand.Parameters("@trial_conclusion").Value = trial_conclusion
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryChargeStatistics(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal date_start As String, ByVal date_end As String, ByVal manager_a As String, ByVal item_name As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryChargeStatisticsCommand Is Nothing Then
            FQueryChargeStatisticsCommand = New SqlCommand("FQueryChargeStatistics", conn)
            FQueryChargeStatisticsCommand.CommandType = CommandType.StoredProcedure
            FQueryChargeStatisticsCommand.Parameters.Add(New SqlParameter("@project_code", SqlDbType.NVarChar))
            FQueryChargeStatisticsCommand.Parameters.Add(New SqlParameter("@enterprise_name", SqlDbType.NVarChar))
            FQueryChargeStatisticsCommand.Parameters.Add(New SqlParameter("@service_type", SqlDbType.NVarChar))
            FQueryChargeStatisticsCommand.Parameters.Add(New SqlParameter("@date_start", SqlDbType.DateTime))
            FQueryChargeStatisticsCommand.Parameters.Add(New SqlParameter("@date_end", SqlDbType.DateTime))
            FQueryChargeStatisticsCommand.Parameters.Add(New SqlParameter("@manager_a", SqlDbType.NVarChar))
            FQueryChargeStatisticsCommand.Parameters.Add(New SqlParameter("@item_name", SqlDbType.NVarChar))
            FQueryChargeStatisticsCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))


        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryChargeStatisticsCommand
            .SelectCommand.Transaction = ts
            FQueryChargeStatisticsCommand.Parameters("@project_code").Value = project_code
            FQueryChargeStatisticsCommand.Parameters("@enterprise_name").Value = enterprise_name
            FQueryChargeStatisticsCommand.Parameters("@service_type").Value = service_type
            FQueryChargeStatisticsCommand.Parameters("@date_start").Value = date_start
            FQueryChargeStatisticsCommand.Parameters("@date_end").Value = date_end
            FQueryChargeStatisticsCommand.Parameters("@manager_a").Value = manager_a
            FQueryChargeStatisticsCommand.Parameters("@item_name").Value = item_name
            FQueryChargeStatisticsCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function DelProject(ByVal ProjectCode As String)

        If DelProjectCommand Is Nothing Then
            DelProjectCommand = New SqlCommand("DelProject", conn)
            DelProjectCommand.CommandType = CommandType.StoredProcedure
            DelProjectCommand.Parameters.Add(New SqlParameter("@sProjectCode", SqlDbType.NVarChar))

        End If

        With DelProjectCommand
            .Transaction = ts

            .Parameters("@sProjectCode").Value = ProjectCode

            .ExecuteNonQuery()
        End With

    End Function

    Public Function FQueryStatisticsCompensation(ByVal StartYear As String, ByVal EndYearMonth As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryStatisticsCompensationCommand Is Nothing Then
            FQueryStatisticsCompensationCommand = New SqlCommand("FQueryStatisticsCompensation", conn)
            FQueryStatisticsCompensationCommand.CommandType = CommandType.StoredProcedure
            FQueryStatisticsCompensationCommand.Parameters.Add(New SqlParameter("@StartYear", SqlDbType.NVarChar))
            FQueryStatisticsCompensationCommand.Parameters.Add(New SqlParameter("@EndYearMonth", SqlDbType.NVarChar))
            FQueryStatisticsCompensationCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryStatisticsCompensationCommand
            .SelectCommand.Transaction = ts
            FQueryStatisticsCompensationCommand.Parameters("@StartYear").Value = StartYear
            FQueryStatisticsCompensationCommand.Parameters("@EndYearMonth").Value = EndYearMonth
            FQueryStatisticsCompensationCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryStatisticsGEProprietorship(ByVal StartYear As String, ByVal EndYearMonth As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryStatisticsGEProprietorshipCommand Is Nothing Then
            FQueryStatisticsGEProprietorshipCommand = New SqlCommand("FQueryStatisticsGEProprietorship", conn)
            FQueryStatisticsGEProprietorshipCommand.CommandType = CommandType.StoredProcedure
            FQueryStatisticsGEProprietorshipCommand.Parameters.Add(New SqlParameter("@StartYear", SqlDbType.NVarChar))
            FQueryStatisticsGEProprietorshipCommand.Parameters.Add(New SqlParameter("@EndYearMonth", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryStatisticsGEProprietorshipCommand
            .SelectCommand.Transaction = ts
            FQueryStatisticsGEProprietorshipCommand.Parameters("@StartYear").Value = StartYear
            FQueryStatisticsGEProprietorshipCommand.Parameters("@EndYearMonth").Value = EndYearMonth
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryStatisticsRegion(ByVal StartYear As String, ByVal EndYearMonth As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryStatisticsRegionCommand Is Nothing Then
            FQueryStatisticsRegionCommand = New SqlCommand("FQueryStatisticsRegion", conn)
            FQueryStatisticsRegionCommand.CommandType = CommandType.StoredProcedure
            FQueryStatisticsRegionCommand.Parameters.Add(New SqlParameter("@StartYear", SqlDbType.NVarChar))
            FQueryStatisticsRegionCommand.Parameters.Add(New SqlParameter("@EndYearMonth", SqlDbType.NVarChar))
            FQueryStatisticsRegionCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryStatisticsRegionCommand
            .SelectCommand.Transaction = ts
            FQueryStatisticsRegionCommand.Parameters("@StartYear").Value = StartYear
            FQueryStatisticsRegionCommand.Parameters("@EndYearMonth").Value = EndYearMonth
            FQueryStatisticsRegionCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryStatisticsCounterguaranteeByMonth(ByVal StartYear As String, ByVal EndYearMonth As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryStatisticsCounterguaranteeByMonthCommand Is Nothing Then
            FQueryStatisticsCounterguaranteeByMonthCommand = New SqlCommand("FQueryStatisticsCounterguaranteeByMonth", conn)
            FQueryStatisticsCounterguaranteeByMonthCommand.CommandType = CommandType.StoredProcedure
            FQueryStatisticsCounterguaranteeByMonthCommand.Parameters.Add(New SqlParameter("@StartYear", SqlDbType.NVarChar))
            FQueryStatisticsCounterguaranteeByMonthCommand.Parameters.Add(New SqlParameter("@EndYearMonth", SqlDbType.NVarChar))
            FQueryStatisticsCounterguaranteeByMonthCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryStatisticsCounterguaranteeByMonthCommand
            .SelectCommand.Transaction = ts
            FQueryStatisticsCounterguaranteeByMonthCommand.Parameters("@StartYear").Value = StartYear
            FQueryStatisticsCounterguaranteeByMonthCommand.Parameters("@EndYearMonth").Value = EndYearMonth
            FQueryStatisticsCounterguaranteeByMonthCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryStatisticsCounterguaranteeByYear(ByVal StartYear As String, ByVal EndYearMonth As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryStatisticsCounterguaranteeByYearCommand Is Nothing Then
            FQueryStatisticsCounterguaranteeByYearCommand = New SqlCommand("FQueryStatisticsCounterguaranteeByYear", conn)
            FQueryStatisticsCounterguaranteeByYearCommand.CommandType = CommandType.StoredProcedure
            FQueryStatisticsCounterguaranteeByYearCommand.Parameters.Add(New SqlParameter("@StartYear", SqlDbType.NVarChar))
            FQueryStatisticsCounterguaranteeByYearCommand.Parameters.Add(New SqlParameter("@EndYearMonth", SqlDbType.NVarChar))
            FQueryStatisticsCounterguaranteeByYearCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryStatisticsCounterguaranteeByYearCommand
            .SelectCommand.Transaction = ts
            FQueryStatisticsCounterguaranteeByYearCommand.Parameters("@StartYear").Value = StartYear
            FQueryStatisticsCounterguaranteeByYearCommand.Parameters("@EndYearMonth").Value = EndYearMonth
            FQueryStatisticsCounterguaranteeByYearCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function


    Public Function FQueryStatisticsPMService(ByVal StartYear As String, ByVal EndYearMonth As String, ByVal ManagerA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryStatisticsPMServiceCommand Is Nothing Then
            FQueryStatisticsPMServiceCommand = New SqlCommand("FQueryStatisticsPMService", conn)
            FQueryStatisticsPMServiceCommand.CommandType = CommandType.StoredProcedure
            FQueryStatisticsPMServiceCommand.Parameters.Add(New SqlParameter("@StartYear", SqlDbType.NVarChar))
            FQueryStatisticsPMServiceCommand.Parameters.Add(New SqlParameter("@EndYearMonth", SqlDbType.NVarChar))
            FQueryStatisticsPMServiceCommand.Parameters.Add(New SqlParameter("@ManagerA", SqlDbType.NVarChar))
            FQueryStatisticsPMServiceCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryStatisticsPMServiceCommand
            .SelectCommand.Transaction = ts
            FQueryStatisticsPMServiceCommand.Parameters("@StartYear").Value = StartYear
            FQueryStatisticsPMServiceCommand.Parameters("@EndYearMonth").Value = EndYearMonth
            FQueryStatisticsPMServiceCommand.Parameters("@ManagerA").Value = ManagerA
            FQueryStatisticsPMServiceCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function


    Public Function PQueryWorkLog(ByVal QueryType As String, ByVal DateStart As DateTime, ByVal DateEnd As DateTime, ByVal AttendPerson As String, ByVal PostName As String, ByVal Responsibility As String, ByVal TaskName As String, ByVal Period As String) As DataSet
        Dim tempDs As New DataSet

        If PQueryWorkLogCommand Is Nothing Then
            PQueryWorkLogCommand = New SqlCommand("dbo.PQueryWorkLog", conn)
            PQueryWorkLogCommand.CommandType = CommandType.StoredProcedure
            PQueryWorkLogCommand.Parameters.Add(New SqlParameter("@Type", SqlDbType.Char))
            PQueryWorkLogCommand.Parameters.Add(New SqlParameter("@DateStart", SqlDbType.DateTime))
            PQueryWorkLogCommand.Parameters.Add(New SqlParameter("@DateEnd", SqlDbType.DateTime))
            PQueryWorkLogCommand.Parameters.Add(New SqlParameter("@AttendPerson", SqlDbType.NVarChar))
            PQueryWorkLogCommand.Parameters.Add(New SqlParameter("@PostName", SqlDbType.NVarChar))
            PQueryWorkLogCommand.Parameters.Add(New SqlParameter("@Responsibility", SqlDbType.NVarChar))
            PQueryWorkLogCommand.Parameters.Add(New SqlParameter("@TaskName", SqlDbType.NVarChar))
            PQueryWorkLogCommand.Parameters.Add(New SqlParameter("@Period", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = PQueryWorkLogCommand
            .SelectCommand.Transaction = ts
            PQueryWorkLogCommand.Parameters("@Type").Value = QueryType
            PQueryWorkLogCommand.Parameters("@DateStart").Value = DateStart
            PQueryWorkLogCommand.Parameters("@DateEnd").Value = DateEnd
            PQueryWorkLogCommand.Parameters("@AttendPerson").Value = AttendPerson
            PQueryWorkLogCommand.Parameters("@PostName").Value = PostName
            PQueryWorkLogCommand.Parameters("@Responsibility").Value = Responsibility
            PQueryWorkLogCommand.Parameters("@TaskName").Value = TaskName
            PQueryWorkLogCommand.Parameters("@Period").Value = Period
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQueryStatisticsGECraft(ByVal StartYear As String, ByVal EndYearMonth As String) As DataSet
        Dim tempDs As New DataSet

        If FQueryStatisticsGECraftCommand Is Nothing Then
            FQueryStatisticsGECraftCommand = New SqlCommand("FQueryStatisticsGECraft", conn)
            FQueryStatisticsGECraftCommand.CommandType = CommandType.StoredProcedure
            FQueryStatisticsGECraftCommand.Parameters.Add(New SqlParameter("@StartYear", SqlDbType.NVarChar))
            FQueryStatisticsGECraftCommand.Parameters.Add(New SqlParameter("@EndYearMonth", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = FQueryStatisticsGECraftCommand
            .SelectCommand.Transaction = ts
            FQueryStatisticsGECraftCommand.Parameters("@StartYear").Value = StartYear
            FQueryStatisticsGECraftCommand.Parameters("@EndYearMonth").Value = EndYearMonth
            .Fill(tempDs)
        End With

        Return tempDs
    End Function


    Public Function PQueryStatisticsMarketingA(ByVal DateStart As DateTime, ByVal DateEnd As DateTime, ByVal Branch As String, ByVal serviceType As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If PQueryStatisticsMarketingACommand Is Nothing Then
            PQueryStatisticsMarketingACommand = New SqlCommand("PQueryStatisticsMarketingA", conn)
            PQueryStatisticsMarketingACommand.CommandType = CommandType.StoredProcedure
            PQueryStatisticsMarketingACommand.Parameters.Add(New SqlParameter("@DateStart", SqlDbType.DateTime))
            PQueryStatisticsMarketingACommand.Parameters.Add(New SqlParameter("@DateEnd", SqlDbType.DateTime))
            PQueryStatisticsMarketingACommand.Parameters.Add(New SqlParameter("@Branch", SqlDbType.NVarChar))
            PQueryStatisticsMarketingACommand.Parameters.Add(New SqlParameter("@vchServiceType", SqlDbType.NVarChar))
            PQueryStatisticsMarketingACommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = PQueryStatisticsMarketingACommand
            .SelectCommand.Transaction = ts
            PQueryStatisticsMarketingACommand.Parameters("@DateStart").Value = DateStart
            PQueryStatisticsMarketingACommand.Parameters("@DateEnd").Value = DateEnd
            PQueryStatisticsMarketingACommand.Parameters("@Branch").Value = Branch
            PQueryStatisticsMarketingACommand.Parameters("@vchServiceType").Value = serviceType
            PQueryStatisticsMarketingACommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PQueryStatisticsMarketingB(ByVal DateStart As DateTime, ByVal DateEnd As DateTime, ByVal phase As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If PQueryStatisticsMarketingBCommand Is Nothing Then
            PQueryStatisticsMarketingBCommand = New SqlCommand("PQueryStatisticsMarketingB", conn)
            PQueryStatisticsMarketingBCommand.CommandType = CommandType.StoredProcedure
            PQueryStatisticsMarketingBCommand.Parameters.Add(New SqlParameter("@DateStart", SqlDbType.DateTime))
            PQueryStatisticsMarketingBCommand.Parameters.Add(New SqlParameter("@DateEnd", SqlDbType.DateTime))
            PQueryStatisticsMarketingBCommand.Parameters.Add(New SqlParameter("@vchPhase", SqlDbType.VarChar))
            PQueryStatisticsMarketingBCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = PQueryStatisticsMarketingBCommand
            .SelectCommand.Transaction = ts
            PQueryStatisticsMarketingBCommand.Parameters("@DateStart").Value = DateStart
            PQueryStatisticsMarketingBCommand.Parameters("@DateEnd").Value = DateEnd
            PQueryStatisticsMarketingBCommand.Parameters("@vchPhase").Value = phase
            PQueryStatisticsMarketingBCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PQueryStatisticsMarketingC(ByVal DateStart As DateTime, ByVal DateEnd As DateTime, ByVal phase As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If PQueryStatisticsMarketingCCommand Is Nothing Then
            PQueryStatisticsMarketingCCommand = New SqlCommand("PQueryStatisticsMarketingC", conn)
            PQueryStatisticsMarketingCCommand.CommandType = CommandType.StoredProcedure
            PQueryStatisticsMarketingCCommand.Parameters.Add(New SqlParameter("@DateStart", SqlDbType.DateTime))
            PQueryStatisticsMarketingCCommand.Parameters.Add(New SqlParameter("@DateEnd", SqlDbType.DateTime))
            PQueryStatisticsMarketingCCommand.Parameters.Add(New SqlParameter("@vchPhase", SqlDbType.VarChar))
            PQueryStatisticsMarketingCCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = PQueryStatisticsMarketingCCommand
            .SelectCommand.Transaction = ts
            PQueryStatisticsMarketingCCommand.Parameters("@DateStart").Value = DateStart
            PQueryStatisticsMarketingCCommand.Parameters("@DateEnd").Value = DateEnd
            PQueryStatisticsMarketingCCommand.Parameters("@vchPhase").Value = phase
            PQueryStatisticsMarketingCCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PStatisticsByType(ByVal month_start As String, ByVal month_end As String, ByVal sRange As String, ByVal sType As String) As DataSet
        Dim tempDs As New DataSet

        If PStatisticsByTypeCommand Is Nothing Then
            PStatisticsByTypeCommand = New SqlCommand("Usp_GetSumTimesOfAreaBank", conn)
            PStatisticsByTypeCommand.CommandType = CommandType.StoredProcedure
            PStatisticsByTypeCommand.Parameters.Add(New SqlParameter("@vchYMFrom", SqlDbType.Char))
            PStatisticsByTypeCommand.Parameters.Add(New SqlParameter("@vchYMTo", SqlDbType.Char))
            PStatisticsByTypeCommand.Parameters.Add(New SqlParameter("@vchRange", SqlDbType.Char))
            PStatisticsByTypeCommand.Parameters.Add(New SqlParameter("@vchType", SqlDbType.Char))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = PStatisticsByTypeCommand
            .SelectCommand.Transaction = ts
            .SelectCommand.CommandTimeout = 1200
            PStatisticsByTypeCommand.Parameters("@vchYMFrom").Value = month_start
            PStatisticsByTypeCommand.Parameters("@vchYMTo").Value = month_end
            PStatisticsByTypeCommand.Parameters("@vchRange").Value = sRange
            PStatisticsByTypeCommand.Parameters("@vchType").Value = sType
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PStatisticsByTypeEx(ByVal procedureName As String, ByVal month_start As String, ByVal month_end As String, ByVal sRange As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If PStatisticsByTypeCommand Is Nothing Then
            PStatisticsByTypeCommand = New SqlCommand(procedureName, conn)
            PStatisticsByTypeCommand.CommandType = CommandType.StoredProcedure
            PStatisticsByTypeCommand.Parameters.Add(New SqlParameter("@vchYMFrom", SqlDbType.Char))
            PStatisticsByTypeCommand.Parameters.Add(New SqlParameter("@vchYMTo", SqlDbType.Char))
            PStatisticsByTypeCommand.Parameters.Add(New SqlParameter("@vchRange", SqlDbType.Char))
            PStatisticsByTypeCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = PStatisticsByTypeCommand
            .SelectCommand.Transaction = ts
            .SelectCommand.CommandTimeout = 1200
            PStatisticsByTypeCommand.Parameters("@vchYMFrom").Value = month_start
            PStatisticsByTypeCommand.Parameters("@vchYMTo").Value = month_end
            PStatisticsByTypeCommand.Parameters("@vchRange").Value = sRange
            PStatisticsByTypeCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PStatisticsFee(ByVal month_start As String, ByVal month_end As String, ByVal sType As String, ByVal sSubType As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If PStatisticsFeeCommand Is Nothing Then
            PStatisticsFeeCommand = New SqlCommand("Usp_GetChargeSumByYM", conn)
            PStatisticsFeeCommand.CommandType = CommandType.StoredProcedure
            PStatisticsFeeCommand.Parameters.Add(New SqlParameter("@vchYMFrom", SqlDbType.Char))
            PStatisticsFeeCommand.Parameters.Add(New SqlParameter("@vchYMTo", SqlDbType.Char))
            PStatisticsFeeCommand.Parameters.Add(New SqlParameter("@vchType", SqlDbType.Char))
            PStatisticsFeeCommand.Parameters.Add(New SqlParameter("@vchSubType", SqlDbType.Char))
            PStatisticsFeeCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = PStatisticsFeeCommand
            .SelectCommand.Transaction = ts
            PStatisticsFeeCommand.Parameters("@vchYMFrom").Value = month_start
            PStatisticsFeeCommand.Parameters("@vchYMTo").Value = month_end
            PStatisticsFeeCommand.Parameters("@vchType").Value = sType
            PStatisticsFeeCommand.Parameters("@vchSubType").Value = sSubType
            PStatisticsFeeCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PQueryStatisticsRecommendProjectByMonth(ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal RecommendPerson As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If PQueryStatisticsRecommendProjectByMonthCommand Is Nothing Then
            PQueryStatisticsRecommendProjectByMonthCommand = New SqlCommand("PQueryStatisticsRecommendProjectByMonth", conn)
            PQueryStatisticsRecommendProjectByMonthCommand.CommandType = CommandType.StoredProcedure
            PQueryStatisticsRecommendProjectByMonthCommand.Parameters.Add(New SqlParameter("@StartDate", SqlDbType.DateTime))
            PQueryStatisticsRecommendProjectByMonthCommand.Parameters.Add(New SqlParameter("@EndDate", SqlDbType.DateTime))
            PQueryStatisticsRecommendProjectByMonthCommand.Parameters.Add(New SqlParameter("@RecommendPerson", SqlDbType.Char))
            PQueryStatisticsRecommendProjectByMonthCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))


        End If

        With dsCommand_CommonQuery
            .SelectCommand = PQueryStatisticsRecommendProjectByMonthCommand
            .SelectCommand.Transaction = ts
            PQueryStatisticsRecommendProjectByMonthCommand.Parameters("@StartDate").Value = StartDate
            PQueryStatisticsRecommendProjectByMonthCommand.Parameters("@EndDate").Value = EndDate
            PQueryStatisticsRecommendProjectByMonthCommand.Parameters("@RecommendPerson").Value = RecommendPerson
            PQueryStatisticsRecommendProjectByMonthCommand.Parameters("@userName").Value = userName

            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PQueryStatisticsRecommendProjectByYear(ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal RecommendPerson As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If PQueryStatisticsRecommendProjectByYearCommand Is Nothing Then
            PQueryStatisticsRecommendProjectByYearCommand = New SqlCommand("PQueryStatisticsRecommendProjectByYear", conn)
            PQueryStatisticsRecommendProjectByYearCommand.CommandType = CommandType.StoredProcedure
            PQueryStatisticsRecommendProjectByYearCommand.Parameters.Add(New SqlParameter("@StartDate", SqlDbType.DateTime))
            PQueryStatisticsRecommendProjectByYearCommand.Parameters.Add(New SqlParameter("@EndDate", SqlDbType.DateTime))
            PQueryStatisticsRecommendProjectByYearCommand.Parameters.Add(New SqlParameter("@RecommendPerson", SqlDbType.Char))
            PQueryStatisticsRecommendProjectByYearCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))


        End If

        With dsCommand_CommonQuery
            .SelectCommand = PQueryStatisticsRecommendProjectByYearCommand
            .SelectCommand.Transaction = ts
            PQueryStatisticsRecommendProjectByYearCommand.Parameters("@StartDate").Value = StartDate
            PQueryStatisticsRecommendProjectByYearCommand.Parameters("@EndDate").Value = EndDate
            PQueryStatisticsRecommendProjectByYearCommand.Parameters("@RecommendPerson").Value = RecommendPerson
            PQueryStatisticsRecommendProjectByYearCommand.Parameters("@userName").Value = userName

            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PQueryStatisticsRecommendProject(ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal RecommendPerson As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If PQueryStatisticsRecommendProjectCommand Is Nothing Then
            PQueryStatisticsRecommendProjectCommand = New SqlCommand("PQueryStatisticsRecommendProject", conn)
            PQueryStatisticsRecommendProjectCommand.CommandType = CommandType.StoredProcedure
            PQueryStatisticsRecommendProjectCommand.Parameters.Add(New SqlParameter("@StartDate", SqlDbType.DateTime))
            PQueryStatisticsRecommendProjectCommand.Parameters.Add(New SqlParameter("@EndDate", SqlDbType.DateTime))
            PQueryStatisticsRecommendProjectCommand.Parameters.Add(New SqlParameter("@RecommendPerson", SqlDbType.Char))
            PQueryStatisticsRecommendProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))


        End If

        With dsCommand_CommonQuery
            .SelectCommand = PQueryStatisticsRecommendProjectCommand
            .SelectCommand.Transaction = ts
            PQueryStatisticsRecommendProjectCommand.Parameters("@StartDate").Value = StartDate
            PQueryStatisticsRecommendProjectCommand.Parameters("@EndDate").Value = EndDate
            PQueryStatisticsRecommendProjectCommand.Parameters("@RecommendPerson").Value = RecommendPerson
            PQueryStatisticsRecommendProjectCommand.Parameters("@userName").Value = userName

            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function Usp_ListIsFirstLoanStat(ByVal dtFrom As DateTime, ByVal dtTo As DateTime, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If Usp_ListIsFirstLoanStatCommand Is Nothing Then
            Usp_ListIsFirstLoanStatCommand = New SqlCommand("Usp_ListIsFirstLoanStat", conn)
            Usp_ListIsFirstLoanStatCommand.CommandType = CommandType.StoredProcedure
            Usp_ListIsFirstLoanStatCommand.Parameters.Add(New SqlParameter("@dtFrom", SqlDbType.DateTime))
            Usp_ListIsFirstLoanStatCommand.Parameters.Add(New SqlParameter("@dtTo", SqlDbType.DateTime))
            Usp_ListIsFirstLoanStatCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = Usp_ListIsFirstLoanStatCommand
            .SelectCommand.Transaction = ts
            Usp_ListIsFirstLoanStatCommand.Parameters("@dtFrom").Value = dtFrom
            Usp_ListIsFirstLoanStatCommand.Parameters("@dtTo").Value = dtTo
            Usp_ListIsFirstLoanStatCommand.Parameters("@userName").Value = userName

            .Fill(tempDs)
        End With

        Return tempDs
    End Function


    Public Function Usp_ListConsultation(ByVal corporation_code As String, ByVal corporation_name As String, ByVal district_name As String, _
            ByVal recommend_person As String, ByVal consult_person As String, ByVal dtConsultFrom As String, ByVal dtConsultTo As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If Usp_ListConsultationCommand Is Nothing Then
            Usp_ListConsultationCommand = New SqlCommand("Usp_ListConsultation", conn)
            Usp_ListConsultationCommand.CommandType = CommandType.StoredProcedure
            Usp_ListConsultationCommand.Parameters.Add(New SqlParameter("@corporation_code", SqlDbType.Char))
            Usp_ListConsultationCommand.Parameters.Add(New SqlParameter("@corporation_name", SqlDbType.NVarChar))
            Usp_ListConsultationCommand.Parameters.Add(New SqlParameter("@district_name", SqlDbType.NVarChar))
            Usp_ListConsultationCommand.Parameters.Add(New SqlParameter("@recommend_person", SqlDbType.NVarChar))
            Usp_ListConsultationCommand.Parameters.Add(New SqlParameter("@consult_person", SqlDbType.NVarChar))
            Usp_ListConsultationCommand.Parameters.Add(New SqlParameter("@dtConsultFrom", SqlDbType.DateTime))
            Usp_ListConsultationCommand.Parameters.Add(New SqlParameter("@dtConsultTo", SqlDbType.DateTime))
            Usp_ListConsultationCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = Usp_ListConsultationCommand
            .SelectCommand.Transaction = ts
            Usp_ListConsultationCommand.Parameters("@corporation_code").Value = corporation_code
            Usp_ListConsultationCommand.Parameters("@corporation_name").Value = corporation_name
            Usp_ListConsultationCommand.Parameters("@district_name").Value = district_name
            Usp_ListConsultationCommand.Parameters("@recommend_person").Value = recommend_person
            Usp_ListConsultationCommand.Parameters("@consult_person").Value = consult_person
            Usp_ListConsultationCommand.Parameters("@dtConsultFrom").Value = dtConsultFrom
            Usp_ListConsultationCommand.Parameters("@dtConsultTo").Value = dtConsultTo
            Usp_ListConsultationCommand.Parameters("@userName").Value = userName

            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function Usp_GetUnDealProject(ByVal serviceType As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If Usp_GetUnDealProjectCommand Is Nothing Then
            Usp_GetUnDealProjectCommand = New SqlCommand("Usp_GetUnDealProject", conn)
            Usp_GetUnDealProjectCommand.CommandType = CommandType.StoredProcedure
            Usp_GetUnDealProjectCommand.Parameters.Add(New SqlParameter("@vchServiceType", SqlDbType.VarChar))
            Usp_GetUnDealProjectCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.VarChar))
            Usp_GetUnDealProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = Usp_GetUnDealProjectCommand
            .SelectCommand.Transaction = ts
            Usp_GetUnDealProjectCommand.Parameters("@vchServiceType").Value = serviceType
            Usp_GetUnDealProjectCommand.Parameters("@vchPMA").Value = vchPMA
            Usp_GetUnDealProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function Usp_GetAfterGuaranteeRecord(ByVal corporationName As String, ByVal serviceType As String, ByVal managerA As String, ByVal dtCheckFrom As String, ByVal dtCheckTo As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If Usp_GetAfterGuaranteeRecordCommand Is Nothing Then
            Usp_GetAfterGuaranteeRecordCommand = New SqlCommand("Usp_GetAfterGuaranteeRecord", conn)
            Usp_GetAfterGuaranteeRecordCommand.CommandType = CommandType.StoredProcedure
            Usp_GetAfterGuaranteeRecordCommand.Parameters.Add(New SqlParameter("@corporation_name", SqlDbType.NVarChar))
            Usp_GetAfterGuaranteeRecordCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.VarChar))
            Usp_GetAfterGuaranteeRecordCommand.Parameters.Add(New SqlParameter("@manager_A", SqlDbType.NVarChar))
            Usp_GetAfterGuaranteeRecordCommand.Parameters.Add(New SqlParameter("@dtCheckFrom", SqlDbType.DateTime))
            Usp_GetAfterGuaranteeRecordCommand.Parameters.Add(New SqlParameter("@dtCheckTo", SqlDbType.DateTime))
            Usp_GetAfterGuaranteeRecordCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = Usp_GetAfterGuaranteeRecordCommand
            .SelectCommand.Transaction = ts
            Usp_GetAfterGuaranteeRecordCommand.Parameters("@corporation_name").Value = corporationName
            Usp_GetAfterGuaranteeRecordCommand.Parameters("@ServiceType").Value = serviceType
            Usp_GetAfterGuaranteeRecordCommand.Parameters("@manager_A").Value = managerA
            Usp_GetAfterGuaranteeRecordCommand.Parameters("@dtCheckFrom").Value = dtCheckFrom
            Usp_GetAfterGuaranteeRecordCommand.Parameters("@dtCheckTo").Value = dtCheckTo
            Usp_GetAfterGuaranteeRecordCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function Usp_GetGuaranteeProject(ByVal LoanFrom As String, ByVal LoanTo As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If Usp_GetGuaranteeProjectCommand Is Nothing Then
            Usp_GetGuaranteeProjectCommand = New SqlCommand("Usp_GetGuaranteeProject", conn)
            Usp_GetGuaranteeProjectCommand.CommandType = CommandType.StoredProcedure
            Usp_GetGuaranteeProjectCommand.Parameters.Add(New SqlParameter("@dtLoanFrom", SqlDbType.DateTime))
            Usp_GetGuaranteeProjectCommand.Parameters.Add(New SqlParameter("@dtLoanTo", SqlDbType.DateTime))
            Usp_GetGuaranteeProjectCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.VarChar))
            Usp_GetGuaranteeProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = Usp_GetGuaranteeProjectCommand
            .SelectCommand.Transaction = ts
            Usp_GetGuaranteeProjectCommand.Parameters("@dtLoanFrom").Value = LoanFrom
            Usp_GetGuaranteeProjectCommand.Parameters("@dtLoanTo").Value = LoanTo
            Usp_GetGuaranteeProjectCommand.Parameters("@vchPMA").Value = vchPMA
            Usp_GetGuaranteeProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FQryUnsignProject(ByVal ProjectCode As String, ByVal CorpName As String, ByVal ServiceType As String, ByVal dtFrom As String, ByVal dtTo As String, ByVal phase As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If QryUnSignProjectCommand Is Nothing Then
            QryUnSignProjectCommand = New SqlCommand("Usp_ListUnviseProject", conn)
            QryUnSignProjectCommand.CommandType = CommandType.StoredProcedure
            QryUnSignProjectCommand.Parameters.Add(New SqlParameter("@vchProjectCode", SqlDbType.Char))
            QryUnSignProjectCommand.Parameters.Add(New SqlParameter("@vchCorpName", SqlDbType.Char))
            QryUnSignProjectCommand.Parameters.Add(New SqlParameter("@vchServiceType", SqlDbType.Char))
            QryUnSignProjectCommand.Parameters.Add(New SqlParameter("@dtTrialFrom", SqlDbType.DateTime))
            QryUnSignProjectCommand.Parameters.Add(New SqlParameter("@dtTrialTo", SqlDbType.DateTime))
            QryUnSignProjectCommand.Parameters.Add(New SqlParameter("@vchPhase", SqlDbType.VarChar))
            QryUnSignProjectCommand.Parameters.Add(New SqlParameter("@vchPMA", SqlDbType.VarChar))
            QryUnSignProjectCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = QryUnSignProjectCommand
            .SelectCommand.Transaction = ts
            QryUnSignProjectCommand.Parameters("@vchProjectCode").Value = ProjectCode
            QryUnSignProjectCommand.Parameters("@vchCorpName").Value = CorpName
            QryUnSignProjectCommand.Parameters("@vchServiceType").Value = ServiceType
            QryUnSignProjectCommand.Parameters("@dtTrialFrom").Value = dtFrom
            QryUnSignProjectCommand.Parameters("@dtTrialTo").Value = dtTo
            QryUnSignProjectCommand.Parameters("@vchPhase").Value = phase
            QryUnSignProjectCommand.Parameters("@vchPMA").Value = vchPMA
            QryUnSignProjectCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function Usp_ListGuaranteeForm(ByVal vchProjectCode As String, ByVal vchCorpName As String, _
                     ByVal dtSignFrom As String, ByVal dtSignTo As String, ByVal dtLoanFrom As String, ByVal dtLoanTo As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If Usp_ListGuaranteeFormCommand Is Nothing Then
            Usp_ListGuaranteeFormCommand = New SqlCommand("Usp_ListGuaranteeForm", conn)
            Usp_ListGuaranteeFormCommand.CommandType = CommandType.StoredProcedure
            Usp_ListGuaranteeFormCommand.Parameters.Add(New SqlParameter("@vchProjectCode", SqlDbType.VarChar))
            Usp_ListGuaranteeFormCommand.Parameters.Add(New SqlParameter("@vchCorpName", SqlDbType.VarChar))
            Usp_ListGuaranteeFormCommand.Parameters.Add(New SqlParameter("@dtSignFrom", SqlDbType.DateTime))
            Usp_ListGuaranteeFormCommand.Parameters.Add(New SqlParameter("@dtSignTo", SqlDbType.DateTime))
            Usp_ListGuaranteeFormCommand.Parameters.Add(New SqlParameter("@dtLoanFrom", SqlDbType.DateTime))
            Usp_ListGuaranteeFormCommand.Parameters.Add(New SqlParameter("@dtLoanTo", SqlDbType.DateTime))
            Usp_ListGuaranteeFormCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If
        With dsCommand_CommonQuery
            .SelectCommand = Usp_ListGuaranteeFormCommand
            .SelectCommand.Transaction = ts
            Usp_ListGuaranteeFormCommand.Parameters("@vchProjectCode").Value = vchProjectCode
            Usp_ListGuaranteeFormCommand.Parameters("@vchCorpName").Value = vchCorpName
            Usp_ListGuaranteeFormCommand.Parameters("@dtSignFrom").Value = dtSignFrom
            Usp_ListGuaranteeFormCommand.Parameters("@dtSignTo").Value = dtSignTo
            Usp_ListGuaranteeFormCommand.Parameters("@dtLoanFrom").Value = dtLoanFrom
            Usp_ListGuaranteeFormCommand.Parameters("@dtLoanTo").Value = dtLoanTo
            Usp_ListGuaranteeFormCommand.Parameters("@userName").Value = userName

            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function GetCorporationAttendeePerson(ByVal projectCode As String, ByVal serviceType As String, ByVal role_id As String, ByVal acceptPerson As String) As DataSet
        Dim tempDs As New DataSet

        If Usp_ListIsFirstLoanStatCommand Is Nothing Then
            'Usp_ListIsFirstLoanStatCommand = New SqlCommand("GetCorporationAttendeePersonEx", conn)
            Usp_ListIsFirstLoanStatCommand = New SqlCommand("GetCorporationAttendeePerson", conn)
            Usp_ListIsFirstLoanStatCommand.CommandType = CommandType.StoredProcedure
            Usp_ListIsFirstLoanStatCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            Usp_ListIsFirstLoanStatCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.NVarChar))
            Usp_ListIsFirstLoanStatCommand.Parameters.Add(New SqlParameter("@RoleID", SqlDbType.NVarChar))
            Usp_ListIsFirstLoanStatCommand.Parameters.Add(New SqlParameter("@Person", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = Usp_ListIsFirstLoanStatCommand
            .SelectCommand.Transaction = ts
            Usp_ListIsFirstLoanStatCommand.Parameters("@ProjectCode").Value = projectCode
            Usp_ListIsFirstLoanStatCommand.Parameters("@ServiceType").Value = serviceType
            Usp_ListIsFirstLoanStatCommand.Parameters("@RoleID").Value = role_id
            Usp_ListIsFirstLoanStatCommand.Parameters("@Person").Value = acceptPerson

            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function GetDefaultPerson(ByVal ProjectCode As String, ByVal role_id As String) As DataSet
        Dim tempDs As New DataSet

        If Usp_ListIsFirstLoanStatCommand Is Nothing Then
            'Usp_ListIsFirstLoanStatCommand = New SqlCommand("Usp_GetRolesByProject", conn)
            Usp_ListIsFirstLoanStatCommand = New SqlCommand("GetDefaultPerson", conn)
            Usp_ListIsFirstLoanStatCommand.CommandType = CommandType.StoredProcedure
            Usp_ListIsFirstLoanStatCommand.Parameters.Add(New SqlParameter("@vchProjectCode", SqlDbType.Char))
            Usp_ListIsFirstLoanStatCommand.Parameters.Add(New SqlParameter("@vchRoleID", SqlDbType.Char))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = Usp_ListIsFirstLoanStatCommand
            .SelectCommand.Transaction = ts
            Usp_ListIsFirstLoanStatCommand.Parameters("@vchProjectCode").Value = ProjectCode
            Usp_ListIsFirstLoanStatCommand.Parameters("@vchRoleID").Value = role_id

            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PQueryProjectRequite(ByVal PorjectCode As String, ByVal Corporation As String, ByVal ServiceType As String, ByVal ManangerA As String, ByVal RefundType As String, ByVal IsNormal As String, ByVal IsPartion As String, ByVal objDate As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If QueryProjectRequiteCommand Is Nothing Then
            QueryProjectRequiteCommand = New SqlCommand("PQueryProjectRequite", conn)
            QueryProjectRequiteCommand.CommandType = CommandType.StoredProcedure
            QueryProjectRequiteCommand.Parameters.Add(New SqlParameter("@PorjectCode", SqlDbType.NVarChar))
            QueryProjectRequiteCommand.Parameters.Add(New SqlParameter("@Corporation", SqlDbType.NVarChar))
            QueryProjectRequiteCommand.Parameters.Add(New SqlParameter("@ServiceType", SqlDbType.NVarChar))

            QueryProjectRequiteCommand.Parameters.Add(New SqlParameter("@ManangerA", SqlDbType.NVarChar))
            QueryProjectRequiteCommand.Parameters.Add(New SqlParameter("@RefundType", SqlDbType.NVarChar))
            QueryProjectRequiteCommand.Parameters.Add(New SqlParameter("@IsNormal", SqlDbType.Char))
            QueryProjectRequiteCommand.Parameters.Add(New SqlParameter("@IsPartion", SqlDbType.Char))
            QueryProjectRequiteCommand.Parameters.Add(New SqlParameter("@Date", SqlDbType.DateTime))
            QueryProjectRequiteCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = QueryProjectRequiteCommand
            .SelectCommand.Transaction = ts
            QueryProjectRequiteCommand.Parameters("@PorjectCode").Value = PorjectCode
            QueryProjectRequiteCommand.Parameters("@Corporation").Value = Corporation
            QueryProjectRequiteCommand.Parameters("@ServiceType").Value = ServiceType

            QueryProjectRequiteCommand.Parameters("@ManangerA").Value = ManangerA
            QueryProjectRequiteCommand.Parameters("@RefundType").Value = RefundType
            QueryProjectRequiteCommand.Parameters("@IsNormal").Value = IsNormal
            QueryProjectRequiteCommand.Parameters("@IsPartion").Value = IsPartion
            QueryProjectRequiteCommand.Parameters("@Date").Value = DateTime.Parse(objDate)
            QueryProjectRequiteCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function GetIntentLetter(ByVal strCondition As String) As DataSet
        Dim tempDs As New DataSet

        If IntentLetterCommand Is Nothing Then
            IntentLetterCommand = New SqlCommand("PQueryIntentLetterEx", conn)
            IntentLetterCommand.CommandType = CommandType.StoredProcedure
            IntentLetterCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.VarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = IntentLetterCommand
            .SelectCommand.Transaction = ts
            IntentLetterCommand.Parameters("@Condition").Value = strCondition
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PQueryIntentLetterInfo(ByVal PutOutType As String, ByVal signStartDate As String, ByVal signEndDate As String, _
            ByVal issueStartDate As String, ByVal issueEndDate As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If IntentLetterCommand Is Nothing Then
            IntentLetterCommand = New SqlCommand("PQueryIntentLetterInfo", conn)
            IntentLetterCommand.CommandType = CommandType.StoredProcedure
            IntentLetterCommand.Parameters.Add(New SqlParameter("@PutOutType", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@signStartDate", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@signEndDate", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@issueStartDate", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@issueEndDate", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))

        End If

        With dsCommand_CommonQuery
            .SelectCommand = IntentLetterCommand
            .SelectCommand.Transaction = ts
            IntentLetterCommand.Parameters("@PutOutType").Value = PutOutType
            IntentLetterCommand.Parameters("@signStartDate").Value = signStartDate
            IntentLetterCommand.Parameters("@signEndDate").Value = signEndDate
            IntentLetterCommand.Parameters("@issueStartDate").Value = issueStartDate
            IntentLetterCommand.Parameters("@issueEndDate").Value = issueEndDate
            IntentLetterCommand.Parameters("@userName").Value = userName

            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    '获取反担保物表opposite_guarantee的特殊办法

    Public Function GetGuarantyInfoEx(ByVal ProjectCode As String, ByVal CorporationName As String, ByVal ItemValue As String, ByVal OppGuaranteeForm As String, ByVal EvaluateDate As Object, ByVal Status As String, ByVal GuarantyType As String) As DataSet
        Dim tempDs As New DataSet
        Dim tableName As String = "opposite_guarantee"
        If GetGuarantyInfoExCommand Is Nothing Then
            GetGuarantyInfoExCommand = New SqlCommand("GetGuarantyInfoEX", conn)
            GetGuarantyInfoExCommand.CommandType = CommandType.StoredProcedure
            GetGuarantyInfoExCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            GetGuarantyInfoExCommand.Parameters.Add(New SqlParameter("@CorporationName", SqlDbType.NVarChar))
            GetGuarantyInfoExCommand.Parameters.Add(New SqlParameter("@ItemValue", SqlDbType.NVarChar))
            GetGuarantyInfoExCommand.Parameters.Add(New SqlParameter("@opposite_guarantee_form", SqlDbType.NVarChar))
            GetGuarantyInfoExCommand.Parameters.Add(New SqlParameter("@Evaluate_Date", SqlDbType.DateTime))
            GetGuarantyInfoExCommand.Parameters.Add(New SqlParameter("@Status", SqlDbType.NVarChar))
            GetGuarantyInfoExCommand.Parameters.Add(New SqlParameter("@GuarantyType", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = GetGuarantyInfoExCommand
            .SelectCommand.Transaction = ts
            GetGuarantyInfoExCommand.Parameters("@ProjectCode").Value = ProjectCode
            GetGuarantyInfoExCommand.Parameters("@CorporationName").Value = CorporationName
            GetGuarantyInfoExCommand.Parameters("@ItemValue").Value = ItemValue
            GetGuarantyInfoExCommand.Parameters("@opposite_guarantee_form").Value = OppGuaranteeForm
            GetGuarantyInfoExCommand.Parameters("@Evaluate_Date").Value = IIf(IsDBNull(EvaluateDate) OrElse EvaluateDate Is Nothing, DBNull.Value, EvaluateDate)
            GetGuarantyInfoExCommand.Parameters("@Status").Value = Status
            GetGuarantyInfoExCommand.Parameters("@GuarantyType").Value = GuarantyType
            .Fill(tempDs, tableName)
        End With

        Return tempDs
    End Function

    Public Function PCopyOppGuarantee(ByVal ProjectCode As String, ByVal SourceProjectCode As String, ByVal SourceSerialNum As String, ByVal CreatePerson As String, ByVal CreateDate As Date)


        If PCopyOppGuaranteeCommand Is Nothing Then
            PCopyOppGuaranteeCommand = New SqlCommand("PCopyOppGuarantee", conn)
            PCopyOppGuaranteeCommand.CommandType = CommandType.StoredProcedure
            PCopyOppGuaranteeCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            PCopyOppGuaranteeCommand.Parameters.Add(New SqlParameter("@SourceProjectCode", SqlDbType.NVarChar))
            PCopyOppGuaranteeCommand.Parameters.Add(New SqlParameter("@SourceSerialNum", SqlDbType.Int))
            PCopyOppGuaranteeCommand.Parameters.Add(New SqlParameter("@CreatePerson", SqlDbType.NVarChar))
            PCopyOppGuaranteeCommand.Parameters.Add(New SqlParameter("@CreateDate", SqlDbType.DateTime))

        End If

        'With dsCommand_CommonQuery
        '    .SelectCommand = PCopyOppGuaranteeCommand
        '    .SelectCommand.Transaction = ts
        PCopyOppGuaranteeCommand.Transaction = ts
        PCopyOppGuaranteeCommand.CommandTimeout = 1200
        PCopyOppGuaranteeCommand.Parameters("@ProjectCode").Value = ProjectCode
        PCopyOppGuaranteeCommand.Parameters("@SourceProjectCode").Value = SourceProjectCode
        PCopyOppGuaranteeCommand.Parameters("@SourceSerialNum").Value = SourceSerialNum
        PCopyOppGuaranteeCommand.Parameters("@CreatePerson").Value = CreatePerson
        PCopyOppGuaranteeCommand.Parameters("@CreateDate").Value = CreateDate
        PCopyOppGuaranteeCommand.ExecuteNonQuery()
        'End With

    End Function

    Public Function PQueryOppEvaluate(ByVal ProjectCode As String, ByVal CorporationName As String, _
        ByVal ManagerA As String, ByVal Evaluater As String, ByVal EvaluateStatus As String, _
        ByVal BookFrom As String, ByVal BookTo As String, ByVal AffirmFrom As String, _
        ByVal AffirmTo As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet

        If PQueryOppEvaluateCommand Is Nothing Then
            PQueryOppEvaluateCommand = New SqlCommand("Udp_FetchOpplist", conn)
            PQueryOppEvaluateCommand.CommandType = CommandType.StoredProcedure
            PQueryOppEvaluateCommand.Parameters.Add(New SqlParameter("@nvchProjectCode", SqlDbType.NVarChar))
            PQueryOppEvaluateCommand.Parameters.Add(New SqlParameter("@nvchCorpName", SqlDbType.NVarChar))
            PQueryOppEvaluateCommand.Parameters.Add(New SqlParameter("@nvchManagerA", SqlDbType.NVarChar))

            PQueryOppEvaluateCommand.Parameters.Add(New SqlParameter("@nvchEvaluater", SqlDbType.NVarChar))
            PQueryOppEvaluateCommand.Parameters.Add(New SqlParameter("@nvchEvaluateStatus", SqlDbType.NVarChar))
            PQueryOppEvaluateCommand.Parameters.Add(New SqlParameter("@dBookFrom", SqlDbType.DateTime))
            PQueryOppEvaluateCommand.Parameters.Add(New SqlParameter("@dBookTo", SqlDbType.DateTime))
            PQueryOppEvaluateCommand.Parameters.Add(New SqlParameter("@dAffirmFrom", SqlDbType.DateTime))
            PQueryOppEvaluateCommand.Parameters.Add(New SqlParameter("@dAffirmTo", SqlDbType.DateTime))
            PQueryOppEvaluateCommand.Parameters.Add(New SqlParameter("@userName", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = PQueryOppEvaluateCommand
            .SelectCommand.Transaction = ts
            PQueryOppEvaluateCommand.Parameters("@nvchProjectCode").Value = ProjectCode
            PQueryOppEvaluateCommand.Parameters("@nvchCorpName").Value = CorporationName
            PQueryOppEvaluateCommand.Parameters("@nvchManagerA").Value = ManagerA

            PQueryOppEvaluateCommand.Parameters("@nvchEvaluater").Value = Evaluater
            PQueryOppEvaluateCommand.Parameters("@nvchEvaluateStatus").Value = EvaluateStatus
            PQueryOppEvaluateCommand.Parameters("@dBookFrom").Value = BookFrom
            PQueryOppEvaluateCommand.Parameters("@dBookTo").Value = BookTo
            PQueryOppEvaluateCommand.Parameters("@dAffirmFrom").Value = AffirmFrom
            PQueryOppEvaluateCommand.Parameters("@dAffirmTo").Value = AffirmTo 'DateTime.Parse(AffirmTo)
            PQueryOppEvaluateCommand.Parameters("@userName").Value = userName
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function FetchProjectFinanceAnalyseIntegration(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal ThisYear As String, ByVal LastYear1 As String, ByVal LastYear2 As String, ByVal LastYear3 As String) As DataSet
        Dim tempDs As New DataSet

        If FetchProjectFinanceAnalyseIntegrationCommand Is Nothing Then
            FetchProjectFinanceAnalyseIntegrationCommand = New SqlCommand("PFetchProjectFinanceAnalyseIntegration", conn)
            FetchProjectFinanceAnalyseIntegrationCommand.CommandType = CommandType.StoredProcedure
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters.Add(New SqlParameter("@ProjectNo", SqlDbType.NVarChar))
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters.Add(New SqlParameter("@CorporationNo", SqlDbType.NVarChar))
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters.Add(New SqlParameter("@Phase", SqlDbType.NVarChar))
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters.Add(New SqlParameter("@ThisYear", SqlDbType.NVarChar))
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters.Add(New SqlParameter("@LastYear1", SqlDbType.NVarChar))
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters.Add(New SqlParameter("@LastYear2", SqlDbType.NVarChar))
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters.Add(New SqlParameter("@LastYear3", SqlDbType.NVarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = FetchProjectFinanceAnalyseIntegrationCommand
            .SelectCommand.Transaction = ts
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters("@ProjectNo").Value = ProjectNo
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters("@CorporationNo").Value = CorporationNo
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters("@Phase").Value = Phase
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters("@ThisYear").Value = ThisYear
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters("@LastYear1").Value = LastYear1
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters("@LastYear2").Value = LastYear2
            FetchProjectFinanceAnalyseIntegrationCommand.Parameters("@LastYear3").Value = LastYear3
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function PUpdateProcess(ByVal ProjectCode As String, ByVal WorkFlowID As String)

        If PUpdateProcessCommand Is Nothing Then
            PUpdateProcessCommand = New SqlCommand("InsertProjectTaskFromTemplate", conn)
            PUpdateProcessCommand.CommandType = CommandType.StoredProcedure
            PUpdateProcessCommand.Parameters.Add(New SqlParameter("@ProjectCode", SqlDbType.NVarChar))
            PUpdateProcessCommand.Parameters.Add(New SqlParameter("@WorkFlowID", SqlDbType.NVarChar))
        End If

        'With dsCommand_CommonQuery
        '    .SelectCommand = PCopyOppGuaranteeCommand
        '    .SelectCommand.Transaction = ts
        PUpdateProcessCommand.Transaction = ts
        '获取或设置在终止执行命令的尝试并生成错误之前的等待时间,默认为30 秒。
        PUpdateProcessCommand.CommandTimeout = 1200
        PUpdateProcessCommand.Parameters("@ProjectCode").Value = ProjectCode
        PUpdateProcessCommand.Parameters("@WorkFlowID").Value = WorkFlowID
        PUpdateProcessCommand.ExecuteNonQuery()
        'End With

    End Function

    '获得工作日志
    Public Function GetWorkingHours(ByVal staff_name As String, ByVal start_date As Object, ByVal end_date As Object, _
                        ByVal period As String, ByVal statisticsType As Integer) As DataSet
        Dim tempDs As New DataSet
        If IntentLetterCommand Is Nothing Then
            IntentLetterCommand = New SqlCommand("PStatisticsWorkingHours", conn)
            IntentLetterCommand.CommandType = CommandType.StoredProcedure
            IntentLetterCommand.Parameters.Add(New SqlParameter("@staff_name", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@start_date", SqlDbType.DateTime))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@end_date", SqlDbType.DateTime))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@period", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@statisticsType", SqlDbType.Int))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = IntentLetterCommand
            .SelectCommand.Transaction = ts
            IntentLetterCommand.Parameters("@staff_name").Value = staff_name
            IntentLetterCommand.Parameters("@start_date").Value = start_date
            IntentLetterCommand.Parameters("@end_date").Value = end_date
            IntentLetterCommand.Parameters("@period").Value = period
            IntentLetterCommand.Parameters("@statisticsType").Value = statisticsType

            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    'qxd add 2005-3-17
    '查询反担保物
    Public Function GetQueryOppGuarantInfo(ByVal projectCode As String, ByVal corporationName As String, ByVal oppForm As String, _
                        ByVal oppStatus As String, ByVal itemType As String, ByVal itemCodeFirst As String, _
                        ByVal itemValueFirst As String, ByVal itemCodeSecond As String, ByVal itemValueSecond As String, _
                        ByVal startDate As String, ByVal endDate As String) As DataSet
        Dim tempDs As New DataSet
        If IntentLetterCommand Is Nothing Then
            IntentLetterCommand = New SqlCommand("GetQueryOppGuarantInfo", conn)
            IntentLetterCommand.CommandType = CommandType.StoredProcedure
            IntentLetterCommand.Parameters.Add(New SqlParameter("@projectCode", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@corporationName", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@oppForm", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@oppStatus", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@itemType", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@itemCodeFirst", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@itemValueFirst", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@itemCodeSecond", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@itemValueSecond", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@startDate", SqlDbType.VarChar))
            IntentLetterCommand.Parameters.Add(New SqlParameter("@endDate", SqlDbType.VarChar))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = IntentLetterCommand
            .SelectCommand.Transaction = ts
            IntentLetterCommand.Parameters("@projectCode").Value = projectCode
            IntentLetterCommand.Parameters("@corporationName").Value = corporationName
            IntentLetterCommand.Parameters("@oppForm").Value = oppForm
            IntentLetterCommand.Parameters("@oppStatus").Value = oppStatus
            IntentLetterCommand.Parameters("@itemType").Value = itemType
            IntentLetterCommand.Parameters("@itemCodeFirst").Value = itemCodeFirst
            IntentLetterCommand.Parameters("@itemValueFirst").Value = itemValueFirst
            IntentLetterCommand.Parameters("@itemCodeSecond").Value = itemCodeSecond
            IntentLetterCommand.Parameters("@itemValueSecond").Value = itemValueSecond
            IntentLetterCommand.Parameters("@startDate").Value = startDate
            IntentLetterCommand.Parameters("@endDate").Value = endDate
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function GetGuarantingCorporationList(ByVal start_date As Object, ByVal end_date As Object) As DataSet
        Dim tempDs As New DataSet
        If PGetGuarantingCorporationListCommand Is Nothing Then
            PGetGuarantingCorporationListCommand = New SqlCommand("dbo.PGetGuarantingCorporationList", conn)
            PGetGuarantingCorporationListCommand.CommandType = CommandType.StoredProcedure
            PGetGuarantingCorporationListCommand.Parameters.Add(New SqlParameter("@start_date", SqlDbType.DateTime))
            PGetGuarantingCorporationListCommand.Parameters.Add(New SqlParameter("@end_date", SqlDbType.DateTime))
        End If

        With dsCommand_CommonQuery
            .SelectCommand = PGetGuarantingCorporationListCommand
            .SelectCommand.Transaction = ts
            PGetGuarantingCorporationListCommand.Parameters("@start_date").Value = start_date
            PGetGuarantingCorporationListCommand.Parameters("@end_date").Value = end_date
            .Fill(tempDs)
        End With

        Return tempDs
    End Function

    Public Function GetMaxContractNum(ByVal ProjectCode As String) As String

        '判断是否额度项下保函，如果不是，使用新合同号，否则，不显示合同号
        Dim MaxNum As String
        Dim tmpLen As Integer
        Dim i As Integer
        Dim dtProjectContract As DataTable
        Dim strSql As String
        Dim dsProject As DataSet = GetCommonQueryInfo("delete from project_contract_num where contract_num=''")
        dsProject = GetCommonQueryInfo("select workflow from project where project_code='" & ProjectCode & "'")
        If IIf(IsDBNull(dsProject.Tables(0).Rows(0).Item("workflow")), "", dsProject.Tables(0).Rows(0).Item("workflow")) <> "额度项下保函" Then
            strSql = "select top 1 contract_num from project_contract_num where contract_year='" & Year(Now).ToString & "' order by create_date desc"
            dtProjectContract = GetCommonQueryInfo(strSql).Tables(0)
            If dtProjectContract.Rows.Count <> 0 Then
                MaxNum = CInt(dtProjectContract.Rows(0).Item("contract_num")) + 1
            Else
                MaxNum = "1"
            End If

            tmpLen = 4 - Len(MaxNum)
            For i = 0 To tmpLen - 1
                MaxNum = "0" & MaxNum
            Next

            Return MaxNum
        Else
            Return ""
        End If


    End Function

End Class
