Imports System.Data.SqlClient
Imports System.Web.Services
Imports System.Configuration
Imports System.Web.Services.Protocols
Imports BusinessRules


Namespace TestWebService

<WebService(Namespace:="http://tempuri.org/")> _
Public Class Service1
    Inherits System.Web.Services.WebService

#Region " Web 服务设计器生成的代码 "

    Public Sub New()
        MyBase.New()

        '该调用是 Web 服务设计器所必需的。
        InitializeComponent()

        '在 InitializeComponent() 调用之后添加您自己的初始化代码

    End Sub

    'Web 服务设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意：以下过程是 Web 服务设计器所必需的
    '可以使用 Web 服务设计器修改此过程。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        'CODEGEN: 此过程是 Web 服务设计器所必需的
        '不要使用代码编辑器修改它。
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region

    ' Web 服务示例
    ' HelloWorld() 示例服务返回字符串 Hello World。
    ' 若要生成项目，请取消注释以下行，然后保存并生成项目。
    ' 若要测试此 Web 服务，请确保 .asmx 文件为起始页
    ' 并按 F5 键。
    '
    '获取Webconfig配置文件中的数据库连接设置
    Private strConn As String = ConfigurationSettings.AppSettings("DBConnection")
    'Private strStartTime As String = ConfigurationSettings.AppSettings("startTime")

    '定义预警消息扫描全局变量
    Private Ddone(30) As Boolean
    Private Hdone(23) As Boolean

    <WebMethod()> Public Function ScanTimingTask() As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim i As Integer
            Ddone = Application("Ddone")
            Hdone = Application("Hdone")
            Dim iDay As Integer = DateTime.Today.Day
            Dim iHour As Integer = DateTime.Now.Hour

            Dim tmpTimingServer As TimingServer

            '如果当天未扫描过，则扫描
            If Ddone(iDay - 1) = False Then
                tmpTimingServer = New TimingServer(conn, ts, Ddone(iDay - 1), Hdone(iHour - 1))
                tmpTimingServer.TimingServer()

                '标志当天已扫描完成
                For i = 0 To 30
                    Ddone(i) = False
                Next
                Ddone(iDay - 1) = True

                '标志本小时已扫描完成
                For i = 0 To 23
                    Hdone(i) = False
                Next

                If iHour = 0 Then
                    Hdone(23) = True
                Else
                    Hdone(iHour - 1) = True
                End If
            End If

            '如果当前小时未扫描过，则扫描
            If Hdone(iHour - 1) = False Then
                tmpTimingServer = New TimingServer(conn, ts, Ddone(iDay - 1), Hdone(iHour - 1))
                tmpTimingServer.TimingServer()

                '标志当天已扫描完成
                For i = 0 To 30
                    Ddone(i) = False
                Next
                Ddone(iDay - 1) = True

                '标志本小时已扫描完成
                For i = 0 To 23
                    Hdone(i) = False
                Next

                If iHour = 0 Then
                    Hdone(23) = True
                Else
                    Hdone(iHour - 1) = True
                End If
            End If

            ts.Commit()
            Return "1"

        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function GetMaxContractNum(ByVal ProjectCode As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetMaxContractNum = CommonQuery.GetMaxContractNum(ProjectCode)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function FQueryProjectExpandDate(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal loan_date_start As String, ByVal loan_date_end As String, ByVal manager_a As String, ByVal bank As String, ByVal branch_bank As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryProjectExpandDate = CommonQuery.FQueryProjectExpandDate(project_code, enterprise_name, service_type, loan_date_start, loan_date_end, manager_a, bank, branch_bank, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function GetProjectExpandDateInfo(ByVal strSQL_Condition_ProjectExpandDate As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectExpandDate As New BusinessRules.ProjectExpandDate(conn, ts)
            GetProjectExpandDateInfo = ProjectExpandDate.GetProjectExpandDateInfo(strSQL_Condition_ProjectExpandDate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function UpdateProjectExpandDate(ByVal ProjectExpandDateSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectExpandDate As New BusinessRules.ProjectExpandDate(conn, ts)
            ProjectExpandDate.UpdateProjectExpandDate(ProjectExpandDateSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectContractCarElementInfo(ByVal strSQL_Condition_ProjectContractCarElement As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectContractCarElement As New BusinessRules.ProjectContractCarElement(conn, ts)
            GetProjectContractCarElementInfo = ProjectContractCarElement.GetProjectContractCarElementInfo(strSQL_Condition_ProjectContractCarElement)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function UpdateProjectContractCarElement(ByVal ProjectContractCarElementSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectContractCarElement As New BusinessRules.ProjectContractCarElement(conn, ts)
            ProjectContractCarElement.UpdateProjectContractCarElement(ProjectContractCarElementSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectContractChattelElementInfo(ByVal strSQL_Condition_ProjectContractChattelElement As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectContractChattelElement As New BusinessRules.ProjectContractChattelElement(conn, ts)
            GetProjectContractChattelElementInfo = ProjectContractChattelElement.GetProjectContractChattelElementInfo(strSQL_Condition_ProjectContractChattelElement)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function UpdateProjectContractChattelElement(ByVal ProjectContractChattelElementSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectContractChattelElement As New BusinessRules.ProjectContractChattelElement(conn, ts)
            ProjectContractChattelElement.UpdateProjectContractChattelElement(ProjectContractChattelElementSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectContractEstateElementInfo(ByVal strSQL_Condition_ProjectContractEstateElement As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectContractEstateElement As New BusinessRules.ProjectContractEstateElement(conn, ts)
            GetProjectContractEstateElementInfo = ProjectContractEstateElement.GetProjectContractEstateElementInfo(strSQL_Condition_ProjectContractEstateElement)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function UpdateProjectContractEstateElement(ByVal ProjectContractEstateElementSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectContractEstateElement As New BusinessRules.ProjectContractEstateElement(conn, ts)
            ProjectContractEstateElement.UpdateProjectContractEstateElement(ProjectContractEstateElementSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectContractElementInfo(ByVal strSQL_Condition_ProjectContractElement As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectContractElement As New BusinessRules.ProjectContractElement(conn, ts)
            GetProjectContractElementInfo = ProjectContractElement.GetProjectContractElementInfo(strSQL_Condition_ProjectContractElement)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function UpdateProjectContractElement(ByVal ProjectContractElementSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectContractElement As New BusinessRules.ProjectContractElement(conn, ts)
            ProjectContractElement.UpdateProjectContractElement(ProjectContractElementSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetAppraisementInfo(ByVal strSQL_Condition_Appraisement As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Appraisement As New BusinessRules.Appraisement(conn, ts)
            GetAppraisementInfo = Appraisement.GetAppraisementInfo(strSQL_Condition_Appraisement)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function UpdateAppraisement(ByVal AppraisementSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Appraisement As New BusinessRules.Appraisement(conn, ts)
            Appraisement.UpdateAppraisement(AppraisementSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWorkflowTypeInfo(ByVal strSQL_Condition_WorkflowType As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkflowType As New BusinessRules.WorkflowType(conn, ts)
            GetWorkflowTypeInfo = WorkflowType.GetWorkflowTypeInfo(strSQL_Condition_WorkflowType)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function UpdateWorkflowType(ByVal WorkflowTypeSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkflowType As New BusinessRules.WorkflowType(conn, ts)
            WorkflowType.UpdateWorkflowType(WorkflowTypeSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetDdGuarantyStatusInfo(ByVal strSQL_Condition_DdGuarantyStatus As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim DdGuarantyStatus As New BusinessRules.DdGuarantyStatus(conn, ts)
            GetDdGuarantyStatusInfo = DdGuarantyStatus.GetDdGuarantyStatusInfo(strSQL_Condition_DdGuarantyStatus)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateDdGuarantyStatus(ByVal DdGuarantyStatusSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim DdGuarantyStatus As New BusinessRules.DdGuarantyStatus(conn, ts)
            DdGuarantyStatus.UpdateDdGuarantyStatus(DdGuarantyStatusSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod(MessageName:="GetAlarmCode2")> Public Function GetAlarmCode(ByVal alarmType As String, ByVal alarmNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim AlarmCode As New BusinessRules.AlarmCode(conn, ts)
            GetAlarmCode = AlarmCode.GetAlarmCode(alarmType, alarmNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetAlarmCode(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim AlarmCode As New BusinessRules.AlarmCode(conn, ts)
            GetAlarmCode = AlarmCode.GetAlarmCode(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateAlarmCode(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim AlarmCode As New BusinessRules.AlarmCode(conn, ts)
            AlarmCode.UpdateAlarmCode(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWorkType(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkType As New BusinessRules.WorkType(conn, ts)
            GetWorkType = WorkType.GetWorkType(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWorkType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkType As New BusinessRules.WorkType(conn, ts)
            WorkType.UpdateWorkType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWorkSubType(ByVal typeCode As String, ByVal subTypeCode As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkType As New BusinessRules.WorkType(conn, ts)
            GetWorkSubType = WorkType.GetWorkSubType(typeCode, subTypeCode)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(messagename:="GetWorkSubTypeEx")> Public Function GetWorkSubType(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkType As New BusinessRules.WorkType(conn, ts)
            GetWorkSubType = WorkType.GetWorkSubType(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWorkSubType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkType As New BusinessRules.WorkType(conn, ts)
            WorkType.UpdateWorkSubType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetAlarmType(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim AlarmType As New BusinessRules.AlarmType(conn, ts)
            GetAlarmType = AlarmType.GetAlarmType(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateAlarmType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim AlarmType As New BusinessRules.AlarmType(conn, ts)
            AlarmType.UpdateAlarmType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetCorporatioRelationType(ByVal TypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporatioRelationType As New BusinessRules.CorporatioRelationType(conn, ts)
            GetCorporatioRelationType = CorporatioRelationType.GetCorporatioRelationType(TypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporatioRelationType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporatioRelationType As New BusinessRules.CorporatioRelationType(conn, ts)
            CorporatioRelationType.UpdateCorporatioRelationType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetCooperateOrganization(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CooperateOrganization As New BusinessRules.CooperateOrganization(conn, ts)
            GetCooperateOrganization = CooperateOrganization.GetCooperateOrganization(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCooperateOrganization(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CooperateOrganization As New BusinessRules.CooperateOrganization(conn, ts)
            CooperateOrganization.UpdateCooperateOrganization(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetCooperateOrganizationOpinion(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CooperateOrganization As New BusinessRules.CooperateOrganization(conn, ts)
            GetCooperateOrganizationOpinion = CooperateOrganization.GetCooperateOrganizationOpinion(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCooperateOrganizationOpinion(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CooperateOrganization As New BusinessRules.CooperateOrganization(conn, ts)
            CooperateOrganization.UpdateCooperateOrganizationOpinion(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetHolidayInfo(ByVal strSQL_Condition_Holiday As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Holiday As New BusinessRules.Holiday(conn, ts)
            GetHolidayInfo = Holiday.GetHolidayInfo(strSQL_Condition_Holiday)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateHoliday(ByVal HolidaySet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Holiday As New BusinessRules.Holiday(conn, ts)
            Holiday.UpdateHoliday(HolidaySet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTracePlanInfo(ByVal strSQL_Condition_TracePlan As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TracePlan As New BusinessRules.TracePlan(conn, ts)
            GetTracePlanInfo = TracePlan.GetTracePlanInfo(strSQL_Condition_TracePlan)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateTracePlan(ByVal TracePlanSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TracePlan As New BusinessRules.TracePlan(conn, ts)
            TracePlan.UpdateTracePlan(TracePlanSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectTaskAttendeeInfo(ByVal strSQL_Condition_ProjectTaskAttendee As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectTaskAttendee As New BusinessRules.ProjectTaskAttendee(conn, ts)
            GetProjectTaskAttendeeInfo = ProjectTaskAttendee.GetProjectTaskAttendeeInfo(strSQL_Condition_ProjectTaskAttendee)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '<WebMethod()> Public Function GetProjectAttendeeInfo(ByVal projectID As String) As DataSet
    '    Dim conn As New SqlConnection(strConn)
    '    conn.Open()
    '    Dim ts As SqlTransaction = conn.BeginTransaction
    '    Dim ProjectTaskAttendee As New BusinessRules.ProjectTaskAttendee(conn, ts)
    '    GetProjectAttendeeInfo = ProjectTaskAttendee.GetProjectAttendeeInfo(projectID)
    '    ts.Commit()
    'End Function

    <WebMethod()> Public Function UpdateProjectTaskAttendee(ByVal ProjectTaskAttendeeSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectTaskAttendee As New BusinessRules.ProjectTaskAttendee(conn, ts)
            ProjectTaskAttendee.UpdateProjectTaskAttendee(ProjectTaskAttendeeSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetBankInfo(ByVal strSQL_Condition_Bank As String, ByVal strSQL_Condition_Branch As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Bank As New BusinessRules.Bank(conn, ts)
            GetBankInfo = Bank.GetBankInfo(strSQL_Condition_Bank, strSQL_Condition_Branch)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateBank(ByVal BankSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Bank As New BusinessRules.Bank(conn, ts)
            Bank.UpdateBank(BankSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateBankBranch(ByVal BranchSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Bank As New BusinessRules.Bank(conn, ts)
            Bank.UpdateBranch(BranchSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function UpdateBankAndBranch(ByVal BankAndBranchSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Bank As New BusinessRules.Bank(conn, ts)
            Bank.UpdateBankAndBranch(BankAndBranchSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectGuaranteeFormAdditional(ByVal projectCode As String, ByVal itemType As String, ByVal itemCode As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectGuaranteeFormAdditional As New BusinessRules.ProjectGuaranteeFormAdditional(conn, ts)
            GetProjectGuaranteeFormAdditional = ProjectGuaranteeFormAdditional.GetProjectGuaranteeFormAdditional(projectCode, itemType, itemCode)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateGuaranteeForm(ByVal FormSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectGuaranteeFormAdditional As New BusinessRules.ProjectGuaranteeFormAdditional(conn, ts)
            ProjectGuaranteeFormAdditional.UpdateGuaranteeForm(FormSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateGuaranteeFormAdditional(ByVal FormAdditionalSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectGuaranteeFormAdditional As New BusinessRules.ProjectGuaranteeFormAdditional(conn, ts)
            ProjectGuaranteeFormAdditional.UpdateGuaranteeFormAdditional(FormAdditionalSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetCheckRecordInfo(ByVal strSQL_Condition_CheckRecord As String, ByVal strSQL_Condition_CheckAlarm As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CheckRecord As New BusinessRules.CheckRecord(conn, ts)
            GetCheckRecordInfo = CheckRecord.GetCheckRecordInfo(strSQL_Condition_CheckRecord, strSQL_Condition_CheckAlarm)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCheckRecordAlarm(ByVal CheckRecordAlarmSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CheckRecord As New BusinessRules.CheckRecord(conn, ts)
            CheckRecord.UpdateCheckRecordAlarm(CheckRecordAlarmSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetConferenceInfo(ByVal strSQL_Condition_Conference As String, ByVal strSQL_Condition_Conference_Committeeman As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Conference As New BusinessRules.Conference(conn, ts)
            GetConferenceInfo = Conference.GetConferenceInfo(strSQL_Condition_Conference, strSQL_Condition_Conference_Committeeman)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateConferenceCommitteeman(ByVal ConferenceCommitteemanSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Conference As New BusinessRules.Conference(conn, ts)
            Conference.UpdateConferenceCommitteeman(ConferenceCommitteemanSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetConfTrialInfo(ByVal strSQL_Condition_ConferenceTrial As String, ByVal strSQL_Condition_CommitteemanOpinion As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ConfTrial As New BusinessRules.ConfTrial(conn, ts)
            GetConfTrialInfo = ConfTrial.GetConfTrialInfo(strSQL_Condition_ConferenceTrial, strSQL_Condition_CommitteemanOpinion)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateConfTrial(ByVal ConfTrialSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ConfTrial As New BusinessRules.ConfTrial(conn, ts)
            ConfTrial.UpdateConfTrial(ConfTrialSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetCooperateOpinionInfo(ByVal strSQL_Condition_CooperateOrganization As String, ByVal strSQL_Condition_CooperateOrganizationOpinion As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CooperateOpinion As New BusinessRules.CooperateOpinion(conn, ts)
            GetCooperateOpinionInfo = CooperateOpinion.GetCooperateOpinionInfo(strSQL_Condition_CooperateOrganization, strSQL_Condition_CooperateOrganizationOpinion)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCooperateOpinion(ByVal CooperateOpinionSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CooperateOpinion As New BusinessRules.CooperateOpinion(conn, ts)
            CooperateOpinion.UpdateCooperateOpinion(CooperateOpinionSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetCorpDefectInfo(ByVal strSQL_Condition_CorpDefect As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorpDefect As New BusinessRules.CorpDefect(conn, ts)
            GetCorpDefectInfo = CorpDefect.GetCorpDefectInfo(strSQL_Condition_CorpDefect)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorpDefect(ByVal CorpDefectSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorpDefect As New BusinessRules.CorpDefect(conn, ts)
            CorpDefect.UpdateCorpDefect(CorpDefectSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectCode(ByVal corporationCode As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim corporationAccess As New BusinessRules.corporationAccess(conn, ts)
            GetProjectCode = corporationAccess.GetProjectCode(corporationCode)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetCorporationMaxCode() As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim corporationAccess As New BusinessRules.corporationAccess(conn, ts)
            GetCorporationMaxCode = corporationAccess.GetCorporationMaxCode()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetCorporationMaxCode_Guarantee() As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim corporationAccess As New BusinessRules.corporationAccess(conn, ts)
            GetCorporationMaxCode_Guarantee = corporationAccess.GetCorporationMaxCode_Guarantee()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetcorporationInfo(ByVal strSQL_Condition_Corporation As String, ByVal strSQL_Condition_Consultation As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim corporationAccess As New BusinessRules.corporationAccess(conn, ts)
            GetcorporationInfo = corporationAccess.GetcorporationInfo(strSQL_Condition_Corporation, strSQL_Condition_Consultation)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorCon(ByVal dataSet_Corporation_Consultation As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim corporationAccess As New BusinessRules.corporationAccess(conn, ts)
            UpdateCorCon = corporationAccess.UpdateCorCon(dataSet_Corporation_Consultation)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationAndProjectCorporation(ByVal CorporationSet As DataSet, ByVal ProjectCorporationSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim corporationAccess As New BusinessRules.corporationAccess(conn, ts)
            Dim ProjectCorporation As New ProjectCorporation(conn, ts)
            corporationAccess.UpdateCorporation(CorporationSet)
            ProjectCorporation.UpdateProjectCorporation(ProjectCorporationSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetGuarantyInfo(ByVal strSQL_Condition_OppositeGuarantee As String, ByVal strSQL_Condition_OppositeGuaranteeDetail As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Guaranty As New BusinessRules.Guaranty(conn, ts)
            GetGuarantyInfo = Guaranty.GetGuarantyInfo(strSQL_Condition_OppositeGuarantee, strSQL_Condition_OppositeGuaranteeDetail)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetMaxGuarantyNum(ByVal projectID As String) As Integer
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Guaranty As New BusinessRules.Guaranty(conn, ts)
            GetMaxGuarantyNum = Guaranty.GetMaxGuarantyNum(projectID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetMaxSerialID(ByVal FieldName As String, ByVal TableName As String) As Int64
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            GetMaxSerialID = ProjectCorporation.GetMaxSerialID(FieldName, TableName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetMaxAppraisementNum(ByVal projectID As String) As Integer
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Appraisement As New BusinessRules.Appraisement(conn, ts)
            GetMaxAppraisementNum = Appraisement.GetMaxAppraisementNum(projectID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetMaxCheckRecordNum(ByVal projectID As String) As Integer
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CheckRecord As New BusinessRules.CheckRecord(conn, ts)
            GetMaxCheckRecordNum = CheckRecord.GetMaxCheckRecordNum(projectID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetMaxConferenceCodeNum() As Integer
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Conference As New BusinessRules.Conference(conn, ts)
            GetMaxConferenceCodeNum = Conference.GetMaxConferenceCodeNum()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateGuaranty(ByVal GuarantySet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Guaranty As New BusinessRules.Guaranty(conn, ts)
            UpdateGuaranty = Guaranty.UpdateGuaranty(GuarantySet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetIntentLetterInfo(ByVal strSQL_Condition_IntentLetter As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim IntentLetter As New BusinessRules.IntentLetter(conn, ts)
            GetIntentLetterInfo = IntentLetter.GetIntentLetterInfo(strSQL_Condition_IntentLetter)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '同时更新项目信息
    <WebMethod()> Public Function UpdateIntentLetter(ByVal IntentLetterSet As DataSet, ByVal ProjectSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim IntentLetter As New BusinessRules.IntentLetter(conn, ts)
            Dim Project As New BusinessRules.Project(conn, ts)
            If ProjectSet Is Nothing Then
                IntentLetter.UpdateIntentLetter(IntentLetterSet)
            Else
                IntentLetter.UpdateIntentLetter(IntentLetterSet)
                Project.UpdateProject(ProjectSet)
            End If
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetInvestigationInfo(ByVal strSQL_Condition_Investigation As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Investigation As New BusinessRules.Investigation(conn, ts)
            GetInvestigationInfo = Investigation.GetInvestigationInfo(strSQL_Condition_Investigation)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateInvestigation(ByVal InvestigationSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Investigation As New BusinessRules.Investigation(conn, ts)
            UpdateInvestigation = Investigation.UpdateInvestigation(InvestigationSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetLoanNoticeInfo(ByVal strSQL_Condition_LoanNotice As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim LoanNotice As New BusinessRules.LoanNotice(conn, ts)
            GetLoanNoticeInfo = LoanNotice.GetLoanNoticeInfo(strSQL_Condition_LoanNotice)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateLoanNotice(ByVal LoanNoticeSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim LoanNotice As New BusinessRules.LoanNotice(conn, ts)
            UpdateLoanNotice = LoanNotice.UpdateLoanNotice(LoanNoticeSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProcessInfo(ByVal strSQL_Condition_Process As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Process As New BusinessRules.Process(conn, ts)
            GetProcessInfo = Process.GetProcessInfo(strSQL_Condition_Process)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProcess(ByVal ProcessSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Process As New BusinessRules.Process(conn, ts)
            UpdateProcess = Process.UpdateProcess(ProcessSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectInfo(ByVal strSQL_Condition_Project As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Project As New BusinessRules.Project(conn, ts)
            GetProjectInfo = Project.GetProjectInfo(strSQL_Condition_Project)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProject(ByVal ProjectSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Project As New BusinessRules.Project(conn, ts)
            UpdateProject = Project.UpdateProject(ProjectSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetMaxProjectAccountDetailNum(ByVal projectID As String) As Integer
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectAccountDetail As New BusinessRules.ProjectAccountDetail(conn, ts)
            GetMaxProjectAccountDetailNum = ProjectAccountDetail.GetMaxProjectAccountDetailNum(projectID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectAccountDetailInfo(ByVal strSQL_Condition_ProjectAccountDetail As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectAccountDetail As New BusinessRules.ProjectAccountDetail(conn, ts)
            GetProjectAccountDetailInfo = ProjectAccountDetail.GetProjectAccountDetailInfo(strSQL_Condition_ProjectAccountDetail)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectAccountDetail(ByVal ProjectAccountDetailSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectAccountDetail As New BusinessRules.ProjectAccountDetail(conn, ts)
            UpdateProjectAccountDetail = ProjectAccountDetail.UpdateProjectAccountDetail(ProjectAccountDetailSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function



    <WebMethod()> Public Function GetRefundCertificateInfo(ByVal strSQL_Condition_RefundCertificate As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim RefundCertificate As New BusinessRules.RefundCertificate(conn, ts)
            GetRefundCertificateInfo = RefundCertificate.GetRefundCertificateInfo(strSQL_Condition_RefundCertificate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateRefundCertificate(ByVal RefundCertificateSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim RefundCertificate As New BusinessRules.RefundCertificate(conn, ts)
            UpdateRefundCertificate = RefundCertificate.UpdateRefundCertificate(RefundCertificateSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetReturnReceiptInfo(ByVal strSQL_Condition_ReturnReceipt As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ReturnReceipt As New BusinessRules.ReturnReceipt(conn, ts)
            GetReturnReceiptInfo = ReturnReceipt.GetReturnReceiptInfo(strSQL_Condition_ReturnReceipt)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateReturnReceipt(ByVal ReturnReceiptSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ReturnReceipt As New BusinessRules.ReturnReceipt(conn, ts)
            UpdateReturnReceipt = ReturnReceipt.UpdateReturnReceipt(ReturnReceiptSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetGuarantyStatus(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim GuarantyStatus As New BusinessRules.GuarantyStatus(conn, ts)
            GetGuarantyStatus = GuarantyStatus.GetGuarantyStatus(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetGuarantyStatusEx(ByVal Status As String, ByVal ID As Int32) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim GuarantyStatus As New BusinessRules.GuarantyStatus(conn, ts)
            GetGuarantyStatusEx = GuarantyStatus.GetGuarantyStatus(Status, ID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateGuarantyStatus(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim GuarantyStatus As New BusinessRules.GuarantyStatus(conn, ts)
            UpdateGuarantyStatus = GuarantyStatus.UpdateGuarantyStatus(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTerminateReportInfo(ByVal strSQL_Condition_TerminateReport As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TerminateReport As New BusinessRules.TerminateReport(conn, ts)
            GetTerminateReportInfo = TerminateReport.GetTerminateReportInfo(strSQL_Condition_TerminateReport)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateTerminateReport(ByVal TerminateReportSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TerminateReport As New BusinessRules.TerminateReport(conn, ts)
            UpdateTerminateReport = TerminateReport.UpdateTerminateReport(TerminateReportSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectGuaranteeForm(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectGuaranteeForm As New BusinessRules.ProjectGuaranteeForm(conn, ts)
            GetProjectGuaranteeForm = ProjectGuaranteeForm.GetProjectGuaranteeForm(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectGuaranteeForm(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectGuaranteeForm As New BusinessRules.ProjectGuaranteeForm(conn, ts)
            UpdateProjectGuaranteeForm = ProjectGuaranteeForm.UpdateProjectGuaranteeForm(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectGuaranteeFormAdd(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectGuaranteeForm As New BusinessRules.ProjectGuaranteeForm(conn, ts)
            GetProjectGuaranteeFormAdd = ProjectGuaranteeForm.GetProjectGuaranteeFormAdd(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectGuaranteeFormAdd(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectGuaranteeForm As New BusinessRules.ProjectGuaranteeForm(conn, ts)
            UpdateProjectGuaranteeFormAdd = ProjectGuaranteeForm.UpdateProjectGuaranteeFormAdd(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetSchema(ByVal TableName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            GetSchema = projectCorporation.GetSchema(TableName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationAccount(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchCorporationAccount = projectCorporation.FetchCorporationAccount(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(MessageName:="FetchCorporationAccount2")> Public Function FetchCorporationAccount(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal ItemType As String, ByVal ItemCode As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchCorporationAccount = projectCorporation.FetchCorporationAccount(ProjectNo, CorporationNo, Phase, Month, ItemType, ItemCode)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationLawsuitRecord(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal SerialID As Int32) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchCorporationLawsuitRecord = projectCorporation.FetchCorporationLawsuitRecord(ProjectNo, CorporationNo, Phase, SerialID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationRatepayingRecord(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal SerialID As Int32) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchCorporationRatepayingRecord = projectCorporation.FetchCorporationRatepayingRecord(ProjectNo, CorporationNo, Phase, SerialID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationBankSaving(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal SerialID As Int32) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchCorporationBankSaving = projectCorporation.FetchCorporationBankSaving(ProjectNo, CorporationNo, Phase, SerialID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationBusiness(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchCorporationBusiness = projectCorporation.FetchCorporationBusiness(ProjectNo, CorporationNo, Phase, Month)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationExternalGuarantee(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal SerialID As Int32) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchCorporationExternalGuarantee = projectCorporation.FetchCorporationExternalGuarantee(ProjectNo, CorporationNo, Phase, SerialID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationLoan(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal SerialID As Int32) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchCorporationLoan = projectCorporation.FetchCorporationLoan(ProjectNo, CorporationNo, Phase, SerialID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationStockStructure(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal SerialID As Int32) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchCorporationStockStructure = projectCorporation.FetchCorporationStockStructure(ProjectNo, CorporationNo, Phase, SerialID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchProjectCorporationEx(ByVal ProjectNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchProjectCorporationEx = projectCorporation.FetchProjectCorporationEx(ProjectNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchProjectCorporation(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationType As String, ByVal Phase As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchProjectCorporation = projectCorporation.FetchProjectCorporation(ProjectNo, CorporationNo, CorporationType, Phase)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(MessageName:="FetchProjectCorporation2")> Public Function FetchProjectCorporation(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchProjectCorporation = projectCorporation.FetchProjectCorporation(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationPostalOrder(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal SerialID As Int64) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            FetchCorporationPostalOrder = projectCorporation.FetchCorporationPostalOrder(ProjectNo, CorporationNo, Phase, SerialID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationPostalOrder(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            UpdateCorporationPostalOrder = projectCorporation.UpdateCorporationPostalOrder(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationAccount(ByVal rstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            UpdateCorporationAccount = projectCorporation.UpdateCorporationAccount(rstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationBankSaving(ByVal rstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            UpdateCorporationBankSaving = projectCorporation.UpdateCorporationBankSaving(rstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationLawsuitRecord(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            UpdateCorporationLawsuitRecord = projectCorporation.UpdateCorporationLawsuitRecord(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationRatepayingRecord(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            UpdateCorporationRatepayingRecord = projectCorporation.UpdateCorporationRatepayingRecord(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationBusiness(ByVal rstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            UpdateCorporationBusiness = projectCorporation.UpdateCorporationBusiness(rstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationExternalGuarantee(ByVal rstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            UpdateCorporationExternalGuarantee = projectCorporation.UpdateCorporationExternalGuarantee(rstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationLoan(ByVal rstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            UpdateCorporationLoan = projectCorporation.UpdateCorporationLoan(rstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationStockStructure(ByVal rstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            UpdateCorporationStockStructure = projectCorporation.UpdateCorporationStockStructure(rstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectCorporation(ByVal rstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim projectCorporation As New BusinessRules.ProjectCorporation(conn, ts)
            UpdateProjectCorporation = projectCorporation.UpdateProjectCorporation(rstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetBranch(ByVal BranchNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Branch As New BusinessRules.Branch(conn, ts)
            GetBranch = Branch.GetBranch(BranchNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateBranch(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Branch As New BusinessRules.Branch(conn, ts)
            UpdateBranch = Branch.UpdateBranch(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetAccount(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporationAccount As New BusinessRules.CorporationAccount(conn, ts)
            GetAccount = CorporationAccount.GetCorporationAccount(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationAccountCredit() As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporationAccount As New BusinessRules.CorporationAccount(conn, ts)
            FetchCorporationAccountCredit = CorporationAccount.FetchCorporationAccountCredit
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationAccountCreditEx(ByVal ProjectNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporationAccount As New BusinessRules.CorporationAccount(conn, ts)
            FetchCorporationAccountCreditEx = CorporationAccount.FetchCorporationAccountCreditEx(ProjectNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(MessageName:="FetchCorporationAccountCredit2")> Public Function FetchCorporationAccountCredit(ByVal ProjectNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporationAccount As New BusinessRules.CorporationAccount(conn, ts)
            FetchCorporationAccountCredit = CorporationAccount.FetchCorporationAccountCredit(ProjectNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetAccountEx(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal ItemNo As String, ByVal ItemType As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporationAccount As New BusinessRules.CorporationAccount(conn, ts)
            GetAccountEx = CorporationAccount.GetCorporationAccount(ProjectNo, CorporationNo, Phase, Month, ItemNo, ItemType)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCorporationAccountMonth(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporationAccount As New BusinessRules.CorporationAccount(conn, ts)
            FetchCorporationAccountMonth = CorporationAccount.FetchCorporationAccountMonth(ProjectNo, CorporationNo, Phase)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateAccount(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporationAccount As New BusinessRules.CorporationAccount(conn, ts)
            UpdateAccount = CorporationAccount.UpdateCorporationAccount(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetCorporationType(ByVal CorporationTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporationType As New BusinessRules.CorporationType(conn, ts)
            GetCorporationType = CorporationType.GetCorporationType(CorporationTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCorporationType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CorporationType As New BusinessRules.CorporationType(conn, ts)
            UpdateCorporationType = CorporationType.UpdateCorporationType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetCurrency(ByVal CurrencyNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Currency As New BusinessRules.Currency(conn, ts)
            GetCurrency = Currency.GetCurrency(CurrencyNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCurrency(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Currency As New BusinessRules.Currency(conn, ts)
            UpdateCurrency = Currency.UpdateCurrency(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetDistrict(ByVal DistrictNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim District As New BusinessRules.District(conn, ts)
            GetDistrict = District.GetDistrict(DistrictNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateDistrict(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim District As New BusinessRules.District(conn, ts)
            UpdateDistrict = District.UpdateDistrict(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetIndustryType(ByVal IndustryTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim IndustryType As New BusinessRules.IndustryType(conn, ts)
            GetIndustryType = IndustryType.GetIndustryType(IndustryTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateIndustryType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim IndustryType As New BusinessRules.IndustryType(conn, ts)
            UpdateIndustryType = IndustryType.UpdateIndustryType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetInvestForm(ByVal InvestFormNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim InvestForm As New BusinessRules.InvestForm(conn, ts)
            GetInvestForm = InvestForm.GetInvestForm(InvestFormNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateInvestForm(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim InvestForm As New BusinessRules.InvestForm(conn, ts)
            UpdateInvestForm = InvestForm.UpdateInvestForm(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetItem(ByVal ItemNo As String, ByVal ItemTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Item As New BusinessRules.Item(conn, ts)
            GetItem = Item.GetItem(ItemNo, ItemTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetItemType(ByVal ItemTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Item As New BusinessRules.Item(conn, ts)
            GetItemType = Item.GetItemType(ItemTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetItemEx(ByVal ItemNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Item As New BusinessRules.Item(conn, ts)
            GetItemEx = Item.GetItemEx(ItemNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateItem(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Item As New BusinessRules.Item(conn, ts)
            UpdateItem = Item.UpdateItem(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateItemType(ByVal dstCommit As DataSet) As Int32
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Item As New BusinessRules.Item(conn, ts)
            UpdateItemType = Item.UpdateItemType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetLoanType(ByVal LoanTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim LoanType As New BusinessRules.LoanType(conn, ts)
            GetLoanType = LoanType.GetLoanType(LoanTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateLoanType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim LoanType As New BusinessRules.LoanType(conn, ts)
            UpdateLoanType = LoanType.UpdateLoanType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetOppositeGuaranteeForm(ByVal FormNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim OppositeGuaranteeForm As New BusinessRules.OppositeGuaranteeForm(conn, ts)
            GetOppositeGuaranteeForm = OppositeGuaranteeForm.GetOppositeGuaranteeForm(FormNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateOppositeGuaranteeForm(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim OppositeGuaranteeForm As New BusinessRules.OppositeGuaranteeForm(conn, ts)
            UpdateOppositeGuaranteeForm = OppositeGuaranteeForm.UpdateOppositeGuaranteeForm(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectAccount(ByVal ProjectNo As String, ByVal SerialID As Int32) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectAccount As New BusinessRules.ProjectAccount(conn, ts)
            GetProjectAccount = ProjectAccount.GetProjectAccount(ProjectNo, SerialID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectAccount(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectAccount As New BusinessRules.ProjectAccount(conn, ts)
            UpdateProjectAccount = ProjectAccount.UpdateProjectAccount(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectDocumentByCondition(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectDocument As New BusinessRules.ProjectDocument(conn, ts)
            GetProjectDocumentByCondition = ProjectDocument.GetProjectDocument(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetProjectDocument(ByVal ProjectNo As String, ByVal Phase As String, ByVal ItemNo As String, ByVal ItemTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectDocument As New BusinessRules.ProjectDocument(conn, ts)
            GetProjectDocument = ProjectDocument.GetProjectDocument(ProjectNo, Phase, ItemNo, ItemTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectDocument(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectDocument As New BusinessRules.ProjectDocument(conn, ts)
            UpdateProjectDocument = ProjectDocument.UpdateProjectDocument(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectTask(ByVal ProjectNo As String, ByVal SerialID As Int32) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectTask As New BusinessRules.ProjectTask(conn, ts)
            GetProjectTask = ProjectTask.GetProjectTask(ProjectNo, SerialID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectTask(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectTask As New BusinessRules.ProjectTask(conn, ts)
            UpdateProjectTask = ProjectTask.UpdateProjectTask(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProprietorshipType(ByVal ProprietorshipTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProprietorshipType As New BusinessRules.ProprietorshipType(conn, ts)
            GetProprietorshipType = ProprietorshipType.GetProprietorshipType(ProprietorshipTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProprietorshipType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProprietorshipType As New BusinessRules.ProprietorshipType(conn, ts)
            UpdateProprietorshipType = ProprietorshipType.UpdateProprietorshipType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetRecommendType(ByVal RecommendTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim RecommendType As New BusinessRules.RecommendType(conn, ts)
            GetRecommendType = RecommendType.GetRecommendType(RecommendTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateRecommendType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim RecommendType As New BusinessRules.RecommendType(conn, ts)
            UpdateRecommendType = RecommendType.UpdateRecommendType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetRecordType(ByVal RecordTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim RecordType As New BusinessRules.RecordType(conn, ts)
            GetRecordType = RecordType.GetRecordType(RecordTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateRecordType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim RecordType As New BusinessRules.RecordType(conn, ts)
            UpdateRecordType = RecordType.UpdateRecordType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetRole(ByVal RoleID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Role As New BusinessRules.Role(conn, ts)
            GetRole = Role.FetchRole(RoleID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateRole(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Role As New BusinessRules.Role(conn, ts)
            UpdateRole = Role.UpdateRole(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateStaffRole(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Role As New BusinessRules.Role(conn, ts)
            UpdateStaffRole = Role.UpdateStaffRole(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetServiceType(ByVal ServiceTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ServiceType As New BusinessRules.ServiceType(conn, ts)
            GetServiceType = ServiceType.GetServiceType(ServiceTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateServiceType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ServiceType As New BusinessRules.ServiceType(conn, ts)
            UpdateServiceType = ServiceType.UpdateServiceType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTaskTemplate(ByVal TaskID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TaskTemplate As New BusinessRules.TaskTemplate(conn, ts)
            GetTaskTemplate = TaskTemplate.GetTaskTemplate(TaskID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateTaskTemplate(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TaskTemplate As New BusinessRules.TaskTemplate(conn, ts)
            UpdateTaskTemplate = TaskTemplate.UpdateTaskTemplate(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetOppositeGuaranteeAssurerInfo(ByVal strSQL_Condition_OppositeGuaranteeAssurer As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim OppositeGuaranteeAssurer As New BusinessRules.OppositeGuaranteeAssurer(conn, ts)
            GetOppositeGuaranteeAssurerInfo = OppositeGuaranteeAssurer.GetOppositeGuaranteeAssurerInfo(strSQL_Condition_OppositeGuaranteeAssurer)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateOppositeGuaranteeAssurer(ByVal OppositeGuaranteeAssurerSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim OppositeGuaranteeAssurer As New BusinessRules.OppositeGuaranteeAssurer(conn, ts)
            UpdateOppositeGuaranteeAssurer = OppositeGuaranteeAssurer.UpdateOppositeGuaranteeAssurer(OppositeGuaranteeAssurerSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTechnologyType(ByVal TechnologyTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TechnologyType As New BusinessRules.TechnologyType(conn, ts)
            GetTechnologyType = TechnologyType.GetTechnologyType(TechnologyTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateTechnologyType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TechnologyType As New BusinessRules.TechnologyType(conn, ts)
            UpdateTechnologyType = TechnologyType.UpdateTechnologyType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTerminateType(ByVal TerminateTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TerminateType As New BusinessRules.TerminateType(conn, ts)
            GetTerminateType = TerminateType.GetTerminateType(TerminateTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateTerminateType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TerminateType As New BusinessRules.TerminateType(conn, ts)
            UpdateTerminateType = TerminateType.UpdateTerminateType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetUser(ByVal UserID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim User As New BusinessRules.User(conn, ts)
            GetUser = User.GetUser(UserID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateUser(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim User As New BusinessRules.User(conn, ts)
            UpdateUser = User.UpdateUser(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWorkLogInfo(ByVal strSQL_Condition_WorkLog As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkLog As New BusinessRules.WorkLog(conn, ts)
            GetWorkLogInfo = WorkLog.GetWorkLogInfo(strSQL_Condition_WorkLog)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWorkLog(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkLog As New BusinessRules.WorkLog(conn, ts)
            UpdateWorkLog = WorkLog.UpdateWorkLog(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectEndCaseInfo(ByVal strSQL_Condition_ProjectEndCase As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectEndCase As New BusinessRules.ProjectEndCase(conn, ts)
            GetProjectEndCaseInfo = ProjectEndCase.GetProjectEndCaseInfo(strSQL_Condition_ProjectEndCase)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectEndCase(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectEndCase As New BusinessRules.ProjectEndCase(conn, ts)
            UpdateProjectEndCase = ProjectEndCase.UpdateProjectEndCase(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetRiskClass(ByVal RiskClassNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim RiskClass As New BusinessRules.RiskClass(conn, ts)
            GetRiskClass = RiskClass.GetRiskClass(RiskClassNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateRiskClass(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim RiskClass As New BusinessRules.RiskClass(conn, ts)
            UpdateRiskClass = RiskClass.UpdateRiskClass(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetPhase(ByVal PhaseNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Phase As New BusinessRules.Phase(conn, ts)
            GetPhase = Phase.GetPhase(PhaseNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdatePhase(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Phase As New BusinessRules.Phase(conn, ts)
            UpdatePhase = Phase.UpdatePhase(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTeam(ByVal TeamID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Team As New BusinessRules.Team(conn, ts)
            GetTeam = Team.FetchTeam(TeamID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetStaffTeam(ByVal TeamID As String, ByVal StaffID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Team As New BusinessRules.Team(conn, ts)
            GetStaffTeam = Team.GetStaffTeam(TeamID, StaffID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function UpdateTeam(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Team As New BusinessRules.Team(conn, ts)
            UpdateTeam = Team.UpdateTeam(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateStaffTeam(ByVal rstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Team As New BusinessRules.Team(conn, ts)
            UpdateStaffTeam = Team.UpdateStaffTeam(rstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetStaff(ByVal StaffID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Staff As New BusinessRules.Staff(conn, ts)
            GetStaff = Staff.FetchStaff(StaffID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetStaffRole(ByVal RoleID As String, ByVal StaffID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Role As New BusinessRules.Role(conn, ts)
            GetStaffRole = Role.GetStaffRole(RoleID, StaffID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetStaffByRoleID(ByVal RoleID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Role As New BusinessRules.Role(conn, ts)
            GetStaffByRoleID = Role.GetStaffRole(RoleID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetStaffEX(ByVal TeamID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Staff As New BusinessRules.Staff(conn, ts)
            GetStaffEX = Staff.FetchStaffEx(TeamID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateStaff(ByVal rstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Staff As New BusinessRules.Staff(conn, ts)
            UpdateStaff = Staff.UpdateStaff(rstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectOpinionBySerialID(ByVal SerialID As Int64) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectOpinion As New BusinessRules.ProjectOpinion(conn, ts)
            GetProjectOpinionBySerialID = ProjectOpinion.GetProjectOpinion(SerialID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectOpinionByProjectNo(ByVal ProjectNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectOpinion As New BusinessRules.ProjectOpinion(conn, ts)
            GetProjectOpinionByProjectNo = ProjectOpinion.GetProjectOpinion(ProjectNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectOpinion(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectOpinion As New BusinessRules.ProjectOpinion(conn, ts)
            UpdateProjectOpinion = ProjectOpinion.UpdateProjectOpinion(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectOpinionAndProjectAccountDetail(ByVal ProjectOpinionASet As DataSet, ByVal ProjectAccountDetailSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectOpinion As New BusinessRules.ProjectOpinion(conn, ts)
            Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
            ProjectOpinion.UpdateProjectOpinion(ProjectOpinionASet)
            ProjectAccountDetail.UpdateProjectAccountDetail(ProjectAccountDetailSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectFileByCondition(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectFile As New BusinessRules.ProjectFile(conn, ts)
            GetProjectFileByCondition = ProjectFile.GetProjectFile(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectFileImageByCondition(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectFile As New BusinessRules.ProjectFile(conn, ts)
            GetProjectFileImageByCondition = ProjectFile.GetProjectFileImage(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectFile(ByVal ProjectNo As String, ByVal ItemNo As String, ByVal ItemTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectFile As New BusinessRules.ProjectFile(conn, ts)
            GetProjectFile = ProjectFile.GetProjectFile(ProjectNo, ItemNo, ItemTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetRelationID() As Int64
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectFile As New BusinessRules.ProjectFile(conn, ts)
            GetRelationID = ProjectFile.GetRelationID()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetProjectFileImage(ByVal ProjectNo As String, ByVal ItemNo As String, ByVal ItemTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectFile As New BusinessRules.ProjectFile(conn, ts)
            GetProjectFileImage = ProjectFile.GetProjectFileImage(ProjectNo, ItemNo, ItemTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectFile(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectFile As New BusinessRules.ProjectFile(conn, ts)
            UpdateProjectFile = ProjectFile.UpdateProjectFile(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectFileImage(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectFile As New BusinessRules.ProjectFile(conn, ts)
            UpdateProjectFileImage = ProjectFile.UpdateProjectFileImage(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '<WebMethod()> Public Function GetConclusion(ByVal Condition As String) As DataSet
    '    Dim conn As New SqlConnection(strConn)
    '    conn.Open()
    '    Dim ts As SqlTransaction = conn.BeginTransaction
    '    Dim Conclusion As New BusinessRules.Conclusion(conn, ts)
    '    GetConclusion = Conclusion.GetConclusion(Condition)
    '    ts.Commit()
    'End Function

    '<WebMethod()> Public Function GetConclusionEx(ByVal WorkflowID As String, ByVal TaskID As String, ByVal strConclusion As String) As DataSet
    '    Dim conn As New SqlConnection(strConn)
    '    conn.Open()
    '    Dim ts As SqlTransaction = conn.BeginTransaction
    '    Dim Conclusion As New BusinessRules.Conclusion(conn, ts)
    '    GetConclusionEx = Conclusion.GetConclusion(WorkflowID, TaskID, strConclusion)
    '    ts.Commit()
    'End Function

    '<WebMethod()> Public Function UpdateConclusion(ByVal dstCommit As DataSet) As String
    '    Dim conn As New SqlConnection(strConn)
    '    conn.Open()
    '    Dim ts As SqlTransaction = conn.BeginTransaction
    '    Dim Conclusion As New BusinessRules.Conclusion(conn, ts)
    '    Try
    '        UpdateConclusion = Conclusion.UpdateConclusion(dstCommit)
    '        ts.Commit()
    '        Return "1"
    '    Catch e As Exception
    '        ts.Rollback()
    '        Return e.Message
    '    End Try
    'End Function

    <WebMethod()> Public Function GetFileTemplateByCondition(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim FileTemplate As New BusinessRules.FileTemplate(conn, ts)
            GetFileTemplateByCondition = FileTemplate.GetFileTemplate(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetFileTemplateEx(ByVal ItemType As String, ByVal ItemNo As String, ByVal Version As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim FileTemplate As New BusinessRules.FileTemplate(conn, ts)
            GetFileTemplateEx = FileTemplate.GetFileTemplate(ItemType, ItemNo, Version)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateFileTemplate(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim FileTemplate As New BusinessRules.FileTemplate(conn, ts)
            UpdateFileTemplate = FileTemplate.UpdateFileTemplate(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetLoanForm(ByVal LoanFormNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim LoanForm As New BusinessRules.LoanForm(conn, ts)
            GetLoanForm = LoanForm.GetLoanForm(LoanFormNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateLoanForm(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim LoanForm As New BusinessRules.LoanForm(conn, ts)
            UpdateLoanForm = LoanForm.UpdateLoanForm(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWfProjectTaskAttendeeInfo(ByVal strSQL_Condition_WfProjectTaskAttendee As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTaskAttendee As New BusinessRules.WfProjectTaskAttendee(conn, ts)
            GetWfProjectTaskAttendeeInfo = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSQL_Condition_WfProjectTaskAttendee)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWfProjectTaskAttendee(ByVal WfProjectTaskAttendeeSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTaskAttendee As New BusinessRules.WfProjectTaskAttendee(conn, ts)
            UpdateWfProjectTaskAttendee = WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(WfProjectTaskAttendeeSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTransCondition(ByVal WorkflowID As String, ByVal projectID As String, ByVal TaskID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTaskTransfer As New BusinessRules.WfProjectTaskTransfer(conn, ts)
            'qxd modify 2005-3-24 为了返回按isItem排序的结论值
            Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & TaskID & "'" & " and (isItem > 0 and not isItem is null ) order by isItem " & "}"
            GetTransCondition = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectTaskTransferInfo(ByVal strSQL_Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTaskTransfer As New BusinessRules.WfProjectTaskTransfer(conn, ts)
            GetProjectTaskTransferInfo = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSQL_Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function CreateProcess(ByVal workFlowID As String, ByVal projectID As String, ByVal userID As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.CreateProcess(workFlowID, projectID, userID)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(messagename:="CreateProcessEx")> Public Function CreateProcess(ByVal workFlowID As String, ByVal projectID As String, ByVal userID As String, ByVal phase As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.CreateProcess(workFlowID, projectID, userID, phase)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function suspendProcess(ByVal projectID As String, ByVal delayDay As Integer) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.suspendProcess(projectID, delayDay)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function resumeProcess(ByVal projectID As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.resumeProcess(projectID)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function isSuspendProcess(ByVal projectID As String) As Boolean
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            isSuspendProcess = WorkFlow.isSuspendProcess(projectID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function finishedTask(ByVal workFlowID As String, ByVal projectID As String, ByVal finishedTaskID As String, ByVal finishedFlag As String, ByVal userID As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.finishedTask(workFlowID, projectID, finishedTaskID, finishedFlag, userID)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(messagename:="finishedTaskEx")> Public Function finishedTask(ByVal workFlowID As String, ByVal projectID As String, ByVal finishedTaskID As String, ByVal finishedFlag As String, ByVal userID As String, ByVal flag As Integer) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.finishedTask(workFlowID, projectID, finishedTaskID, finishedFlag, userID, flag)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function RefreshConference() As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.RefreshConference()
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FinishedReviewConferencePlan(ByVal ConferenceCode As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.FinishedReviewConferencePlan(ConferenceCode)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function CancelReviewConferencePlan(ByVal ConferenceCode As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.CancelReviewConferencePlan(ConferenceCode)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function CancelReviewConferencePlanProject(ByVal projectID As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.CancelReviewConferencePlanProject(projectID)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function CancelSignaturePlan(ByVal SignaturePlanCode As Integer) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.CancelSignaturePlan(SignaturePlanCode)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function CancelSignaturePlanProject(ByVal projectID As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.CancelSignaturePlanProject(projectID)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function ReMeetingPlan(ByVal projectID As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.ReMeetingPlan(projectID)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function ReLoanApplication(ByVal projectID As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.ReLoanApplication(projectID)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function FinishedSignaturePlan(ByVal SignaturePlanCode As Integer) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.FinishedSignaturePlan(SignaturePlanCode)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function



    <WebMethod()> Public Function rollbackTask(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal userID As String, ByVal rollbackMsg As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.rollbackTask(workFlowID, projectID, taskID, userID, rollbackMsg)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function cancelProcess(ByVal projectID As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.cancelProcess(projectID)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function SplitPrjoect(ByVal fatherProjectID As String, ByVal sonProjectID As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.SplitPrjoect(fatherProjectID, sonProjectID)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '<WebMethod()> Public Function TimingServer() As String
    '    Dim conn As New SqlConnection(strConn)
    '    conn.Open()
    '    Dim ts As SqlTransaction = conn.BeginTransaction
    '    Try
    '        Dim tmpTimingServer As New BusinessRules.TimingServer(conn, ts)
    '        tmpTimingServer.TimingServer()
    '        ts.Commit()
    '        Return "1"
    '    Catch errWf As WorkFlowErr
    '        ts.Rollback()
    '        Return errWf.ErrMessage
    '    Catch e As Exception
    '        ts.Rollback()
    '        Return e.Message
    '    End Try
    'End Function

    <WebMethod()> Public Function deleteProcess(ByVal workFlowID As String)

    End Function

    <WebMethod()> Public Function modifiyProcess(ByVal workFlowID As String)

    End Function

    <WebMethod()> Public Function consignTask(ByVal staffID As String, ByVal roleID As String, ByVal consigner As String, ByVal isCurrent As Boolean) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.consignTask(staffID, roleID, consigner, isCurrent)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function CancelconsignTask(ByVal srcPerson As String, ByVal staffID As String, ByVal roleID As String, ByVal isCurrent As Boolean) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.CancelconsignTask(srcPerson, staffID, roleID, isCurrent)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function StartTaskByManual(ByVal workflowID As String, ByVal projectID As String, ByVal taskID As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.StartTaskByManual(workflowID, projectID, taskID)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function LookUpMessage(ByVal strCondition_ProjectMessage As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            LookUpMessage = WorkFlow.LookUpMessage(strCondition_ProjectMessage)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateMessage(ByVal MessageSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectMessages As New BusinessRules.WfProjectMessages(conn, ts)
            UpdateMessage = WfProjectMessages.UpdateWfProjectMessages(MessageSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '查询进行中的任务 
    <WebMethod()> Public Function LookUpWorking(ByVal UserID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            LookUpWorking = WorkFlow.LookUpWorking(UserID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '查询进行中的任务 
    <WebMethod()> Public Function LookUpWorkingEx(ByVal sql_Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            LookUpWorkingEx = WorkFlow.LookUpWorkingEx(sql_Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    '获取流程任务
    <WebMethod()> Public Function GetAllBusinessTasks(ByVal workflowID As String, ByVal projectID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            GetAllBusinessTasks = WorkFlow.GetAllBusinessTasks(workflowID, projectID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function LookUpStatus(ByVal workFlowID As String)

    End Function


    '2003-04-11 by yanxuekui
    '查询还款方式 
    <WebMethod()> Public Function GetRefundType(ByVal RefundTypeNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim RefundType As New BusinessRules.RefundType(conn, ts)
            GetRefundType = RefundType.GetRefundType(RefundTypeNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    '2003-04-11 by yanxuekui
    '修改还款方式 
    <WebMethod()> Public Function UpdateRefundType(ByVal dstCommit As DataSet) As Int32
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim RefundType As New BusinessRules.RefundType(conn, ts)
            UpdateRefundType = RefundType.UpdateRefundType(dstCommit)
            ts.Commit()
            conn.Close()
            conn.Dispose()
        Catch ex As Exception
            ts.Rollback()
            conn.Close()
            conn.Dispose()
            Throw ex
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '查询放款方式 
    <WebMethod()> Public Function GetLoanProvideForm(ByVal LoanProvideFormNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oGetLoanProvideForm As New BusinessRules.LoanProvideForm(conn, ts)
            GetLoanProvideForm = oGetLoanProvideForm.GetLoanProvideForm(LoanProvideFormNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    '2003-04-11 by yanxuekui
    '修改放款方式 
    <WebMethod()> Public Function UpdateLoanProvideForm(ByVal dstCommit As DataSet) As Int32
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oGetLoanProvideForm As New BusinessRules.LoanProvideForm(conn, ts)
            UpdateLoanProvideForm = oGetLoanProvideForm.UpdateLoanProvideForm(dstCommit)
            ts.Commit()
            conn.Close()
            conn.Dispose()
        Catch ex As Exception
            ts.Rollback()
            conn.Close()
            conn.Dispose()
            Throw ex
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    '2006-4-21 By zhoufucai
    '查询项目收费方式
    <WebMethod()> Public Function GetLoanChargeManner(ByVal LoanChargeMannerNo As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oGetLoanChargeManner As New BusinessRules.LoanChargeManner(conn, ts)
            GetLoanChargeManner = oGetLoanChargeManner.GetLoanChargeManner(LoanChargeMannerNo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2006-4-21 By zhoufucai
    '修改项目收费方式
    <WebMethod()> Public Function UpdateLoanChargeManner(ByVal dstCommit As DataSet) As String
      
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oGetLoanChargeManner As New BusinessRules.LoanChargeManner(conn, ts)
            UpdateLoanChargeManner = oGetLoanChargeManner.UpdateLoanChargeManner(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    '2003-04-11 by yanxuekui
    '定性评分--体系编号 
    <WebMethod()> Public Function GetSystemID() As Int32
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQualityEvaluation As New BusinessRules.Credit(conn, ts)
            GetSystemID = oQualityEvaluation.GetSystemID()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(messagename:="GetSystemIDEx")> Public Function GetSystemID(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String) As Int32
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            GetSystemID = oQualityEvaluation.GetSystemID(ProjectNo, CorporationNo, Phase)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    '2003-04-11 by yanxuekui
    '定性评分--查询 项目定性分析记录 
    <WebMethod()> Public Function FetchProjectCreditQuality(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction()
        Try
            Dim oQualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            FetchProjectCreditQuality = oQualityEvaluation.FetchProjectCreditQuality(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定性评分--查询 项目定性分析记录 
    <WebMethod(MessageName:="FetchProjectCreditQuality2")> Public Function FetchProjectCreditQuality(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction()
        Try
            Dim oQualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            FetchProjectCreditQuality = oQualityEvaluation.FetchProjectCreditQuality(ProjectNo, CorporationNo, Phase)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch ex As Exception
            ts.Rollback()
            Throw ex
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定性评分--创建 项目定性分析记录
    <WebMethod()> Public Function CreateProjectCreditQuality(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String) As Boolean
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            CreateProjectCreditQuality = oQualityEvaluation.CreateProjectCreditQuality(ProjectNo, CorporationNo, Phase, Month)
            ts.Commit()

        Catch
            ts.Rollback()
            Return False
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定性评分--更新 项目定性分析记录
    <WebMethod()> Public Function UpdateProjectCreditQuality(ByVal dsCommit As DataSet) As Boolean
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            UpdateProjectCreditQuality = oQualityEvaluation.UpdateProjectCreditQuality(dsCommit)
            ts.Commit()
        Catch
            ts.Rollback()
            Return False
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    '2003-04-11 by yanxuekui
    '定性评分--查询 项目定性分析标准
    <WebMethod()> Public Function FetchCreditQualityStandard(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction()
        Try
            Dim oQualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            FetchCreditQualityStandard = oQualityEvaluation.FetchCreditQualityStandard(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    '2003-04-11 by yanxuekui
    '定性评分--查询 项目定性分析标准
    <WebMethod(MessageName:="FetchCreditQualityStandard2")> Public Function FetchCreditQualityStandard(ByVal SystemID As Integer, ByVal IndexType As String, ByVal IndexID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction()
        Try
            Dim oQualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            FetchCreditQualityStandard = oQualityEvaluation.FetchCreditQualityStandard(SystemID, IndexType, IndexID)
            ts.Commit()
        Catch ex As Exception
            ts.Rollback()
            Throw ex
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定性评分--查询 项目定性分析指标
    <WebMethod()> Public Function FetchCreditQualityIndex(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            FetchCreditQualityIndex = oQualityEvaluation.FetchCreditQualityIndex(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定性评分--查询 项目定性分析指标
    <WebMethod(MessageName:="FetchCreditQualityIndex2")> Public Function FetchCreditQualityIndex(ByVal SystemID As Integer, ByVal IndexType As String, ByVal IndexID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            FetchCreditQualityIndex = oQualityEvaluation.FetchCreditQualityIndex(SystemID, IndexType, IndexID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    ''2003-04-11 by yanxuekui
    ''定性评分--查询 项目定性分析指标
    '<WebMethod(MessageName:="FetchCreditQualityIndex3")> Public Function FetchCreditQualityIndex(ByVal SystemID As Int32, ByVal IndexType As String, ByVal IndexID As String) As DataSet
    '    Dim conn As New SqlConnection(strConn)
    '    conn.Open()
    '    Dim ts As SqlTransaction = conn.BeginTransaction
    '    Dim oQualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
    '    FetchCreditQualityIndex = oQualityEvaluation.FetchCreditQualityIndex(SystemID, IndexType, IndexID)
    '    ts.Commit()
    'End Function


    '2003-04-11 by yanxuekui
    '定量评分--查询 项目定量分析记录
    <WebMethod()> Public Function FetchProjectCreditQuantity(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            FetchProjectCreditQuantity = oQuantityEvaluation.FetchProjectCreditQuantity(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定量评分--查询 项目定量分析记录
    <WebMethod(MessageName:="FetchProjectCreditQuantity2")> Public Function FetchProjectCreditQuantity(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            FetchProjectCreditQuantity = oQuantityEvaluation.FetchProjectCreditQuantity(ProjectNo, CorporationNo, Phase, Month, MonthLast)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function DuplicateCreditAppraise(ByVal sourceID As Integer) As Integer
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Credit As New BusinessRules.Credit(conn, ts)
            DuplicateCreditAppraise = Credit.DuplicateCreditAppraise(sourceID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(MessageName:="DuplicateCreditAppraise2")> Public Function DuplicateCreditAppraise(ByVal sourceID As Integer, ByVal destinationID As Integer) As Integer
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Credit As New BusinessRules.Credit(conn, ts)
            DuplicateCreditAppraise = Credit.DuplicateCreditAppraise(sourceID, destinationID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定量评分--创建 项目定量分析记录
    <WebMethod()> Public Function CreateProjectCreditQuantity(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As Boolean
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            CreateProjectCreditQuantity = oQuantityEvaluation.CreateProjectCreditQuantity(ProjectNo, CorporationNo, Phase, Month, MonthLast)
            ts.Commit()
        Catch
            ts.Rollback()
            Return False
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(messagename:="CreateProjectCreditQuantityEx")> Public Function CreateProjectCreditQuantity(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String, ByVal SystemID As Object) As Boolean
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            CreateProjectCreditQuantity = oQuantityEvaluation.CreateProjectCreditQuantity(ProjectNo, CorporationNo, Phase, Month, MonthLast, SystemID)
            ts.Commit()
        Catch
            ts.Rollback()
            Return False
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定量评分--查询 项目定量分析指标值
    '<WebMethod()> Public Function GetIndexValue(ByVal ProjectNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String, ByVal IndexType As String, ByVal IndexID As String) As Object
    '    Dim conn As New SqlConnection(strConn)
    '    conn.Open()
    '    Dim ts As SqlTransaction = conn.BeginTransaction
    '    Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
    '    GetIndexValue = oQuantityEvaluation.GetIndexValue(ProjectNo, Phase, Month, MonthLast, IndexType, IndexID)
    '    ts.Commit()
    'End Function

    '2003-04-11 by yanxuekui
    '定量评分--查询 项目定量分析指标值
    '<WebMethod(MessageName:="GetIndexValue2")> Public Function GetIndexValue(ByVal ProjectNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String, ByVal IndexType As String, ByVal IndexID As String, ByVal dtCorporationAccount As DataTable) As Object
    '    Dim conn As New SqlConnection(strConn)
    '    conn.Open()
    '    Dim ts As SqlTransaction = conn.BeginTransaction
    '    Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
    '    GetIndexValue = oQuantityEvaluation.GetIndexValue(ProjectNo, Phase, Month, MonthLast, IndexType, IndexID, dtCorporationAccount)
    '    ts.Commit()
    'End Function

    '2003-04-11 by yanxuekui
    '定量评分--查询 项目定量分析指标得分
    '<WebMethod()> Public Function GetIndexScore(ByVal SystemID As Int32, ByVal IndexType As String, ByVal IndexID As String, ByVal Value As Decimal, ByRef dtCreditQuantityStandard As DataTable) As Object
    '    Dim conn As New SqlConnection(strConn)
    '    conn.Open()
    '    Dim ts As SqlTransaction = conn.BeginTransaction
    '    Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
    '    GetIndexScore = oQuantityEvaluation.GetIndexScore(SystemID, IndexType, IndexID, Value, dtCreditQuantityStandard)
    '    ts.Commit()
    'End Function

    '2003-04-11 by yanxuekui
    '定量评分--查询 项目定量分析指标评分标准
    <WebMethod()> Public Function FetchCreditQuantityStandard(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            FetchCreditQuantityStandard = oQuantityEvaluation.FetchCreditQuantityStandard(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定量评分--查询 项目定量分析指标评分标准
    <WebMethod(MessageName:="FetchCreditQuantityStandard2")> Public Function FetchCreditQuantityStandard(ByVal SystemID As Int32, ByVal IndexType As String, ByVal IndexID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            FetchCreditQuantityStandard = oQuantityEvaluation.FetchCreditQuantityStandard(SystemID, IndexType, IndexID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定量评分--查询 项目定量分析指标
    <WebMethod()> Public Function FetchCreditQuantityIndex(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            FetchCreditQuantityIndex = oQuantityEvaluation.FetchCreditQuantityIndex(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2003-04-11 by yanxuekui
    '定量评分--查询 项目定量分析指标
    <WebMethod(MessageName:="FetchCreditQuantityIndex2")> Public Function FetchCreditQuantityIndex(ByVal SystemID As Int32, ByVal IndexType As String, ByVal IndexID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oQuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            FetchCreditQuantityIndex = oQuantityEvaluation.FetchCreditQuantityIndex(SystemID, IndexType, IndexID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCreditQuantityIndex(ByVal dsCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim QuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            UpdateCreditQuantityIndex = QuantityEvaluation.UpdateCreditQuantityIndex(dsCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCreditQuantityStandard(ByVal dsCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim QuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            UpdateCreditQuantityStandard = QuantityEvaluation.UpdateCreditQuantityStandard(dsCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCreditQualityIndex(ByVal dsCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim QualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            UpdateCreditQualityIndex = QualityEvaluation.UpdateCreditQualityIndex(dsCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function UpdateCreditQualityStandard(ByVal dsCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim QualityEvaluation As New BusinessRules.QualityEvaluation(conn, ts)
            UpdateCreditQualityStandard = QualityEvaluation.UpdateCreditQualityStandard(dsCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCreditAppraiseSystem(ByVal dsCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Credit As New BusinessRules.Credit(conn, ts)
            UpdateCreditAppraiseSystem = Credit.UpdateCreditAppraiseSystem(dsCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchCreditAppraiseSystem(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Credit As New BusinessRules.Credit(conn, ts)
            FetchCreditAppraiseSystem = Credit.FetchCreditAppraiseSystem(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function FetchCreditIndexType(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Credit As New BusinessRules.Credit(conn, ts)
            FetchCreditIndexType = Credit.FetchCreditIndexType(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateCreditIndexType(ByVal dsCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Credit As New BusinessRules.Credit(conn, ts)
            UpdateCreditIndexType = Credit.UpdateCreditIndexType(dsCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchProjectCredit(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Credit As New BusinessRules.Credit(conn, ts)
            FetchProjectCredit = Credit.FetchProjectCredit(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(MessageName:="FetchProjectCredit2")> Public Function FetchProjectCredit(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Credit As New BusinessRules.Credit(conn, ts)
            FetchProjectCredit = Credit.FetchProjectCredit(ProjectNo, CorporationNo, Phase, Month, MonthLast)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function CreateProjectCredit(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As Boolean
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction()
        Try
            Dim Credit As New BusinessRules.Credit(conn, ts)
            CreateProjectCredit = Credit.CreateProjectCredit(ProjectNo, CorporationNo, Phase, Month, MonthLast)
            ts.Commit()
            Return True
        Catch ex As Exception
            ts.Rollback()
            Return False
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '通用查询
    <WebMethod()> Public Function GetCommonQueryInfo(ByVal strSql As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetCommonQueryInfo = CommonQuery.GetCommonQueryInfo(strSql)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
        End Function

        '通用查询
        <WebMethod()> Public Function GetCommonStatisticsInfo(ByVal condition As String, ByVal tableCondition As String, ByVal orderBy As String, ByVal cutOffDate As Date, ByVal feeStartDate As Date, ByVal feeEndDate As Date) As DataSet
            Dim conn As New SqlConnection(strConn)
            conn.Open()
            Dim ts As SqlTransaction = conn.BeginTransaction
            Try
                Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
                GetCommonStatisticsInfo = CommonQuery.GetCommonStatisticsInfo(condition, tableCondition, orderBy, cutOffDate, feeStartDate, feeEndDate)
                ts.Commit()
            Catch dbEx As DBConcurrencyException
                ts.Rollback()
                Throw dbEx
            Catch oEx As Exception
                ts.Rollback()
                Throw oEx
            Finally
                conn.Close()
                conn.Dispose()
            End Try
        End Function

    <WebMethod()> Public Function GetProjectSearchInfo(ByVal projectCode As String, ByVal enterpriseName As String, ByVal projectManager As String, ByVal phase As String, ByVal status As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetProjectSearchInfo = CommonQuery.GetProjectSearchInfo(projectCode, enterpriseName, projectManager, phase, status)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetFinanceReviewData(ByVal projectCode As String, ByVal phase As String, ByVal ItemType As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetFinanceReviewData = CommonQuery.GetFinanceReviewData(projectCode, "%", phase, ItemType)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(MessageName:="GetFinanceReviewDataEx")> Public Function GetFinanceReviewData(ByVal projectCode As String, ByVal CorporationCode As String, ByVal phase As String, ByVal ItemType As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetFinanceReviewData = CommonQuery.GetFinanceReviewData(projectCode, CorporationCode, phase, ItemType)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectInfoEx(ByVal strSql_Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetProjectInfoEx = CommonQuery.GetProjectInfoEx(strSql_Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetReGuaranteeProjectInfo(ByVal strSql_Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetReGuaranteeProjectInfo = CommonQuery.GetReGuaranteeProjectInfo(strSql_Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(messageName:="GetQueryProjectInfoEx")> Public Function GetQueryProjectInfo(ByVal projectCode As String, ByVal enterpriseName As String, ByVal projectManager As String, ByVal phase As String, ByVal status As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetQueryProjectInfo = CommonQuery.GetQueryProjectInfo(projectCode, enterpriseName, projectManager, phase, status)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetQueryProjectInfo(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationName As String, ByVal DistrictName As String, ByVal Phase As String, ByVal ApplyDateFrom As DateTime, ByVal ApplyDateTo As DateTime) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetQueryProjectInfo = CommonQuery.GetQueryProjectInfo(ProjectNo, CorporationNo, CorporationName, DistrictName, Phase, ApplyDateFrom, ApplyDateTo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetMeetProject(ByVal startDate As DateTime, ByVal endDate As DateTime) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetMeetProject = CommonQuery.GetMeetProject(startDate, endDate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetOverdueProjectList(ByVal ProjectCode As String, ByVal ServiceType As String, ByVal StartTime As String, ByVal EndTime As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetOverdueProjectList = CommonQuery.GetOverdueProjectList(ProjectCode, ServiceType, StartTime, EndTime, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetQueryFirstProject(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationName As String, ByVal Phase As String, ByVal ServiceType As String, ByVal FromDate As String, ByVal ToDate As String, ByVal vchAcceptBranch As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetQueryFirstProject = CommonQuery.GetQueryFirstProject(ProjectNo, CorporationNo, CorporationName, Phase, ServiceType, FromDate, ToDate, vchAcceptBranch, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetQueryCorporationAttendee(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationName As String, ByVal Phase As String, ByVal ServiceType As String, ByVal FromDate As DateTime, ByVal ToDate As DateTime) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetQueryCorporationAttendee = CommonQuery.GetQueryCorporationAttendee(ProjectNo, CorporationNo, CorporationName, Phase, ServiceType, FromDate, ToDate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetQueryPauseProject(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal CorporationName As String, ByVal Phase As String, ByVal ServiceType As String, ByVal FromDate As String, ByVal ToDate As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetQueryPauseProject = CommonQuery.GetQueryPauseProject(ProjectNo, CorporationNo, CorporationName, Phase, ServiceType, FromDate, ToDate, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetMaturityProjectReview(ByVal ServiceType As String, ByVal StartDate As String, ByVal EndDate As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetMaturityProjectReview = CommonQuery.GetMaturityProjectReview(ServiceType, StartDate, EndDate, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetOnVouchProjectReview(ByVal StartDate As DateTime, ByVal EndDate As DateTime) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetOnVouchProjectReview = CommonQuery.GetOnVouchProjectReview(StartDate, EndDate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectAssignReview(ByVal StartDate As DateTime, ByVal EndDate As DateTime) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetProjectAssignReview = CommonQuery.GetProjectAssignReview(StartDate, EndDate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTerminateProjectReview(ByVal ServiceType As String, ByVal StartDate As String, ByVal EndDate As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetTerminateProjectReview = CommonQuery.GetTerminateProjectReview(ServiceType, StartDate, EndDate, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetRefundDebtProjectList(ByVal ProjectCode As String, ByVal ServiceType As String, ByVal StartTime As String, ByVal EndTime As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetRefundDebtProjectList = CommonQuery.GetRefundDebtProjectList(ProjectCode, ServiceType, StartTime, EndTime, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetNeedMeetProjectInfo(ByVal ProjectList As String, ByVal ConferenceCode As String, ByVal Status As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetNeedMeetProjectInfo = CommonQuery.GetNeedMeetProjectInfo(ProjectList, ConferenceCode, Status)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWfTaskStatus() As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetWfTaskStatus = CommonQuery.GetWfTaskStatus()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetQueryStatisticsAssuranceInfo(ByVal Month_start As String, ByVal Month_end As String, ByVal Type As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetQueryStatisticsAssuranceInfo = CommonQuery.GetQueryStatisticsAssuranceInfo(Month_start, Month_end, Type, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetQueryStatisticsRegionInfo(ByVal DateFrom As DateTime, ByVal DateTo As DateTime) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetQueryStatisticsRegionInfo = CommonQuery.GetQueryStatisticsRegionInfo(DateFrom, DateTo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetQueryStatisticsBankInfo(ByVal DateFrom As DateTime, ByVal DateTo As DateTime, ByVal iType As Integer) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetQueryStatisticsBankInfo = CommonQuery.GetQueryStatisticsBankInfo(DateFrom, DateTo, iType)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetQueryStatisticsIndustryInfo(ByVal DateFrom As DateTime, ByVal DateTo As DateTime) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetQueryStatisticsIndustryInfo = CommonQuery.GetQueryStatisticsIndustryInfo(DateFrom, DateTo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetQueryStatisticsServiceTypeInfo(ByVal DateFrom As DateTime, ByVal DateTo As DateTime) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetQueryStatisticsServiceTypeInfo = CommonQuery.GetQueryStatisticsServiceTypeInfo(DateFrom, DateTo)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectScheduleInfo(ByVal projectID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetProjectScheduleInfo = CommonQuery.GetProjectScheduleInfo(projectID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function ImportFinanceData(ByVal CorporationCode As String, ByVal FromProjectCode As String, ByVal FromPhase As String, ByVal FromMonth As String, ByVal ToCorporationCode As String, ByVal ToProjectCode As String, ByVal ToPhase As String, ByVal CreatePerson As String, ByVal CreateDate As Date, ByVal DeleteOriginalData As Boolean) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            ImportFinanceData = CommonQuery.ImportFinanceData(CorporationCode, FromProjectCode, FromPhase, FromMonth, ToCorporationCode, ToProjectCode, ToPhase, CreatePerson, CreateDate, DeleteOriginalData)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function DeleteAntiAssureCompany(ByVal project_code As String, ByVal corporation_code As String)
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            DeleteAntiAssureCompany = CommonQuery.DeleteAntiAssureCompany(project_code, corporation_code)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function DelProject(ByVal ProjectCode As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            DelProject = CommonQuery.DelProject(ProjectCode)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTaskProjectList(ByVal taskID As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetTaskProjectList = CommonQuery.GetTaskProjectList(taskID, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(MessageName:="GetTaskProjectListEx")> Public Function GetTaskProjectList(ByVal taskID As String, ByVal userName As String, ByVal flag As Integer) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetTaskProjectList = CommonQuery.GetTaskProjectList(taskID, userName, flag)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetConferenceProjectList(ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetConferenceProjectList = CommonQuery.GetConferenceProjectList(userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchFinancialAnalysisInfo(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal ThisYear As String, ByVal LastYear1 As String, ByVal LastYear2 As String, ByVal LastYear3 As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FetchFinancialAnalysisInfo = CommonQuery.FetchFinancialAnalysisInfo(ProjectNo, CorporationNo, Phase, ThisYear, LastYear1, LastYear2, LastYear3)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetNeedSignatureProjectInfo(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetNeedSignatureProjectInfo = CommonQuery.GetNeedSignatureProjectInfo(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function FetchOppositeGuaranteeAssurer(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FetchOppositeGuaranteeAssurer = CommonQuery.FetchOppositeGuaranteeAssurer(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchProjectGuaranteeForm(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FetchProjectGuaranteeForm = CommonQuery.FetchProjectGuaranteeForm(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetAcceptVouchData(ByVal ProjectCode As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetAcceptVouchData = CommonQuery.GetAcceptVouchData(ProjectCode)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTaskListInfo(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetTaskListInfo = CommonQuery.GetTaskListInfo(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetReviewListInfo() As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetReviewListInfo = CommonQuery.GetReviewListInfo()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetDraftOutContractListInfo() As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetDraftOutContractListInfo = CommonQuery.GetDraftOutContractListInfo()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetCapitialEvaluatedListInfo() As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetCapitialEvaluatedListInfo = CommonQuery.GetCapitialEvaluatedListInfo()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetManagerAppraiseListInfo() As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetManagerAppraiseListInfo = CommonQuery.GetManagerAppraiseListInfo()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTeamAppraiseListInfo() As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetTeamAppraiseListInfo = CommonQuery.GetTeamAppraiseListInfo()
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetRefundProcess(ByVal projectcode As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetRefundProcess = CommonQuery.GetRefundProcess(projectcode)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryAcceptProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal apply_service_type As String, ByVal accept_date_start As String, ByVal accept_date_end As String, ByVal apply_bank As String, ByVal belong_area As String, ByVal vchAcceptBranch As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryAcceptProject = CommonQuery.FQueryAcceptProject(project_code, enterprise_name, apply_service_type, accept_date_start, accept_date_end, apply_bank, belong_area, vchAcceptBranch, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryPresentingProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal apply_service_type As String, ByVal evial_date_start As String, ByVal evial_date_end As String, ByVal belong_district As String, ByVal belong_trade As String, ByVal ownership_type As String, ByVal team_name As String, ByVal manager_a As String, ByVal evial_conclusion As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryPresentingProject = CommonQuery.FQueryPresentingProject(project_code, enterprise_name, apply_service_type, evial_date_start, evial_date_end, belong_district, belong_trade, ownership_type, team_name, manager_a, evial_conclusion, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryAllocateProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal apply_service_type As String, ByVal assign_date_start As String, ByVal assign_date_end As String, ByVal manager_a As String, ByVal manager_b As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryAllocateProject = CommonQuery.FQueryAllocateProject(project_code, enterprise_name, apply_service_type, assign_date_start, assign_date_end, manager_a, manager_b, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryLoanProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal loan_date_start As String, ByVal loan_date_end As String, ByVal manager_a As String, ByVal bank As String, ByVal branch_bank As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryLoanProject = CommonQuery.FQueryLoanProject(project_code, enterprise_name, service_type, loan_date_start, loan_date_end, manager_a, bank, branch_bank, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQuerySignProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal sign_date_start As String, ByVal sign_date_end As String, ByVal belong_district As String, ByVal belong_trade As String, ByVal ownership_type As String, ByVal manager_a As String, ByVal bank As String, ByVal branch_bank As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQuerySignProject = CommonQuery.FQuerySignProject(project_code, enterprise_name, service_type, sign_date_start, sign_date_end, belong_district, belong_trade, ownership_type, manager_a, bank, branch_bank, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function PQueryFirstTrialProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal apply_service_type As String, ByVal accept_date_start As String, ByVal accept_date_end As String, ByVal apply_bank As String, ByVal belong_area As String, ByVal belong_trade As String, ByVal ownership_type As String, ByVal vchAcceptBranch As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PQueryFirstTrialProject = CommonQuery.PQueryFirstTrialProject(project_code, enterprise_name, apply_service_type, accept_date_start, accept_date_end, apply_bank, belong_area, belong_trade, ownership_type, vchAcceptBranch, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryCreditProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal today_date As String, ByVal manager_a As String, ByVal bank As String, ByVal branch_bank As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryCreditProject = CommonQuery.FQueryCreditProject(project_code, enterprise_name, service_type, today_date, manager_a, bank, branch_bank, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryRecantProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal date_start As String, ByVal date_end As String, ByVal manager_a As String, ByVal bank As String, ByVal branch_bank As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryRecantProject = CommonQuery.FQueryRecantProject(project_code, enterprise_name, service_type, date_start, date_end, manager_a, bank, branch_bank, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryProcessingProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal date_start As String, ByVal manager_a As String, ByVal manager_b As String, ByVal phase As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryProcessingProject = CommonQuery.FQueryProcessingProject(project_code, enterprise_name, service_type, date_start, manager_a, manager_b, phase, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryRegionProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal date_start As String, ByVal date_end As String, ByVal cooperate_area As String, ByVal phase As String, ByVal vchPMA As String, ByVal userName As String, ByVal recommend_type As String, ByVal opinion As String, ByVal exempt As String, ByVal trial_conclusion As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryRegionProject = CommonQuery.FQueryRegionProject(project_code, enterprise_name, service_type, date_start, date_end, cooperate_area, phase, vchPMA, userName, recommend_type, opinion, exempt, trial_conclusion)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function FQueryRequiteProject(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal refund_date_start As String, ByVal refund_date_end As String, ByVal manager_a As String, ByVal bank As String, ByVal refund_type As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryRequiteProject = CommonQuery.FQueryRequiteProject(project_code, enterprise_name, service_type, refund_date_start, refund_date_end, manager_a, bank, refund_type, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryChargeStatistics(ByVal project_code As String, ByVal enterprise_name As String, ByVal service_type As String, ByVal date_start As String, ByVal date_end As String, ByVal manager_a As String, ByVal item_name As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryChargeStatistics = CommonQuery.FQueryChargeStatistics(project_code, enterprise_name, service_type, date_start, date_end, manager_a, item_name, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryStatisticsCompensation(ByVal StartYear As String, ByVal EndYearMonth As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryStatisticsCompensation = CommonQuery.FQueryStatisticsCompensation(StartYear, EndYearMonth, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryStatisticsGEProprietorship(ByVal StartYear As String, ByVal EndYearMonth As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryStatisticsGEProprietorship = CommonQuery.FQueryStatisticsGEProprietorship(StartYear, EndYearMonth)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryStatisticsRegion(ByVal StartYear As String, ByVal EndYearMonth As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryStatisticsRegion = CommonQuery.FQueryStatisticsRegion(StartYear, EndYearMonth, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryStatisticsCounterguaranteeByMonth(ByVal StartYear As String, ByVal EndYearMonth As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryStatisticsCounterguaranteeByMonth = CommonQuery.FQueryStatisticsCounterguaranteeByMonth(StartYear, EndYearMonth, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryStatisticsCounterguaranteeByYear(ByVal StartYear As String, ByVal EndYearMonth As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryStatisticsCounterguaranteeByYear = CommonQuery.FQueryStatisticsCounterguaranteeByYear(StartYear, EndYearMonth, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryStatisticsPMService(ByVal StartYear As String, ByVal EndYearMonth As String, ByVal ManagerA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryStatisticsPMService = CommonQuery.FQueryStatisticsPMService(StartYear, EndYearMonth, ManagerA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(BufferResponse:=False)> Public Function PStatisticsByType(ByVal month_start As String, ByVal month_end As String, ByVal sRange As String, ByVal sType As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PStatisticsByType = CommonQuery.PStatisticsByType(month_start, month_end, sRange, sType)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(BufferResponse:=False)> Public Function PStatisticsByTypeEx(ByVal procedureName As String, ByVal month_start As String, ByVal month_end As String, ByVal sRange As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PStatisticsByTypeEx = CommonQuery.PStatisticsByTypeEx(procedureName, month_start, month_end, sRange, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function PQueryWorkLog(ByVal QueryType As String, ByVal DateStart As DateTime, ByVal DateEnd As DateTime, ByVal AttendPerson As String, ByVal PostName As String, ByVal Responsibility As String, ByVal TaskName As String, ByVal Period As String) As DataSet
        Dim tempDs As New DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PQueryWorkLog = CommonQuery.PQueryWorkLog(QueryType, DateStart, DateEnd, AttendPerson, PostName, Responsibility, TaskName, Period)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQueryStatisticsGECraft(ByVal StartYear As String, ByVal EndYearMonth As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQueryStatisticsGECraft = CommonQuery.FQueryStatisticsGECraft(StartYear, EndYearMonth)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function
    '****************
    'Public Delegate Function PQueryStatisticsMarketingAAsync(ByVal DateStart As DateTime, ByVal DateEnd As DateTime, ByVal Branch As String, ByVal serviceType As String) As DataSet

    <WebMethod(BufferResponse:=False)> Public Function PQueryStatisticsMarketingA(ByVal DateStart As DateTime, ByVal DateEnd As DateTime, ByVal Branch As String, ByVal serviceType As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PQueryStatisticsMarketingA = CommonQuery.PQueryStatisticsMarketingA(DateStart, DateEnd, Branch, serviceType, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '<WebMethod(BufferResponse:=False)> Public Function BeginPQueryStatisticsMarketingA(ByVal DateStart As DateTime, ByVal DateEnd As DateTime, ByVal Branch As String, ByVal serviceType As String, ByVal cb As System.AsyncCallback, ByVal obj As Object) As IAsyncResult
    '    Dim aysncStub As PQueryStatisticsMarketingAAsync = New PQueryStatisticsMarketingAAsync(AddressOf PQueryStatisticsMarketingA)
    '    Dim myState As myState = New myState()

    '    myState.obj = obj
    '    myState.asyncStub = aysncStub

    '    Return aysncStub.BeginInvoke(DateStart, DateEnd, Branch, serviceType, cb, obj)
    'End Function

    '<WebMethod(BufferResponse:=False)> Public Function EndPQueryStatisticsMarketingA(ByVal cal As System.IAsyncResult) As DataSet
    '    Dim myState As myState ' = New myState()

    '    myState = cal.AsyncState

    '    Return myState.asyncStub.EndInvoke(cal)
    'End Function

    'Public Class myState
    '    Public obj As Object
    '    Public asyncStub As PQueryStatisticsMarketingAAsync

    'End Class
    '*************************

    <WebMethod()> Public Function PQueryStatisticsMarketingB(ByVal DateStart As DateTime, ByVal DateEnd As DateTime, ByVal phase As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PQueryStatisticsMarketingB = CommonQuery.PQueryStatisticsMarketingB(DateStart, DateEnd, phase, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function PQueryStatisticsMarketingC(ByVal DateStart As DateTime, ByVal DateEnd As DateTime, ByVal phase As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PQueryStatisticsMarketingC = CommonQuery.PQueryStatisticsMarketingC(DateStart, DateEnd, phase, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function




    '取服务器当前系统日期
    <WebMethod()> Public Function GetSysTime() As DateTime
        Return Now
    End Function

    <WebMethod()> Public Function GetProjectSignatureInfo(ByVal strSQL_Condition_ProjectSignature As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectSignature As New BusinessRules.ProjectSignature(conn, ts)
            GetProjectSignatureInfo = ProjectSignature.GetProjectSignatureInfo(strSQL_Condition_ProjectSignature)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectSignature(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectSignature As New BusinessRules.ProjectSignature(conn, ts)
            UpdateProjectSignature = ProjectSignature.UpdateProjectSignature(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetSignaturePlanInfo(ByVal strSQL_Condition_SignaturePlan As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim SignaturePlan As New BusinessRules.SignaturePlan(conn, ts)
            GetSignaturePlanInfo = SignaturePlan.GetSignaturePlanInfo(strSQL_Condition_SignaturePlan)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateSignaturePlan(ByVal dsCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim SignaturePlan As New BusinessRules.SignaturePlan(conn, ts)
            UpdateSignaturePlan = SignaturePlan.UpdateSignaturePlan(dsCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWfTaskTemplateInfo(ByVal strSQL_Condition_WfTaskTemplate As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfTaskTemplate As New BusinessRules.WfTaskTemplate(conn, ts)
            GetWfTaskTemplateInfo = WfTaskTemplate.GetWfTaskTemplateInfo(strSQL_Condition_WfTaskTemplate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWfTaskTemplate(ByVal WfTaskTemplateSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfTaskTemplate As New BusinessRules.WfTaskTemplate(conn, ts)
            UpdateWfTaskTemplate = WfTaskTemplate.UpdateWfTaskTemplate(WfTaskTemplateSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWfTaskTransferTemplateInfo(ByVal strSQL_Condition_WfTaskTransferTemplate As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfTaskTransferTemplate As New BusinessRules.WfTaskTransferTemplate(conn, ts)
            GetWfTaskTransferTemplateInfo = WfTaskTransferTemplate.GetWfTaskTransferTemplateInfo(strSQL_Condition_WfTaskTransferTemplate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWfTaskTransferTemplate(ByVal WfTaskTransferTemplateSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfTaskTransferTemplate As New BusinessRules.WfTaskTransferTemplate(conn, ts)
            UpdateWfTaskTransferTemplate = WfTaskTransferTemplate.UpdateWfTaskTransferTemplate(WfTaskTransferTemplateSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWfTaskRoleTemplateInfo(ByVal strSQL_Condition_WfTaskRoleTemplate As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfTaskRoleTemplate As New BusinessRules.WfTaskRoleTemplate(conn, ts)
            GetWfTaskRoleTemplateInfo = WfTaskRoleTemplate.GetWfTaskRoleTemplateInfo(strSQL_Condition_WfTaskRoleTemplate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWfTaskRoleTemplate(ByVal WfTaskRoleTemplateSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfTaskRoleTemplate As New BusinessRules.WfTaskRoleTemplate(conn, ts)
            UpdateWfTaskRoleTemplate = WfTaskRoleTemplate.UpdateWfTaskRoleTemplate(WfTaskRoleTemplateSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWfTimingTaskTemplateInfo(ByVal strSQL_Condition_WfTimingTaskTemplate As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfTimingTaskTemplate As New BusinessRules.WfTimingTaskTemplate(conn, ts)
            GetWfTimingTaskTemplateInfo = WfTimingTaskTemplate.GetWfTimingTaskTemplateInfo(strSQL_Condition_WfTimingTaskTemplate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWfTimingTaskTemplate(ByVal WfTimingTaskTemplateSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfTimingTaskTemplate As New BusinessRules.WfTimingTaskTemplate(conn, ts)
            UpdateWfTimingTaskTemplate = WfTimingTaskTemplate.UpdateWfTimingTaskTemplate(WfTimingTaskTemplateSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function DeleteProjectCreditQuantity(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim QuantityEvaluation As New BusinessRules.QuantityEvaluation(conn, ts)
            DeleteProjectCreditQuantity = QuantityEvaluation.DeleteProjectCreditQuantity(ProjectNo, CorporationNo, Phase, Month)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FStatisticsFee(ByVal month_start As String, ByVal month_end As String, ByVal sType As String, ByVal sSubType As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FStatisticsFee = CommonQuery.PStatisticsFee(month_start, month_end, sType, sSubType, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetGuaranteeLetter(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim GuaranteeLetter As New BusinessRules.GuaranteeLetter(conn, ts)
            GetGuaranteeLetter = GuaranteeLetter.GetGuaranteeLetter(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(MessageName:="GetGuaranteeLetter2")> Public Function GetGuaranteeLetter(ByVal CorporationNo As String, ByVal applyDate As DateTime) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim GuaranteeLetter As New BusinessRules.GuaranteeLetter(conn, ts)
            GetGuaranteeLetter = GuaranteeLetter.GetGuaranteeLetter(CorporationNo, applyDate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateGuaranteeLetter(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim GuaranteeLetter As New BusinessRules.GuaranteeLetter(conn, ts)
            UpdateGuaranteeLetter = GuaranteeLetter.UpdateGuaranteeLetter(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetGuaranteeLetterType(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim GuaranteeLetterType As New BusinessRules.GuaranteeLetterType(conn, ts)
            GetGuaranteeLetterType = GuaranteeLetterType.GetGuaranteeLetterType(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateGuaranteeLetterType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim GuaranteeLetterType As New BusinessRules.GuaranteeLetterType(conn, ts)
            UpdateGuaranteeLetterType = GuaranteeLetterType.UpdateGuaranteeLetterType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetGuaranteeLetterUsage(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim GuaranteeLetterUsage As New BusinessRules.GuaranteeLetterUsage(conn, ts)
            GetGuaranteeLetterUsage = GuaranteeLetterUsage.GetGuaranteeLetterUsage(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateGuaranteeLetterUsage(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim GuaranteeLetterUsage As New BusinessRules.GuaranteeLetterUsage(conn, ts)
            UpdateGuaranteeLetterUsage = GuaranteeLetterUsage.UpdateGuaranteeLetterUsage(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function GetReimburseType(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ReimburseType As New BusinessRules.ReimburseType(conn, ts)
            GetReimburseType = ReimburseType.GetReimburseType(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateReimburseType(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ReimburseType As New BusinessRules.ReimburseType(conn, ts)
            UpdateReimburseType = ReimburseType.UpdateReimburseType(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetMaterial(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Material As New BusinessRules.Material(conn, ts)
            GetMaterial = Material.GetMaterial(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(messagename:="GetMaterialEx")> Public Function GetMaterial(ByVal itemNo As String, ByVal itemTypeNo As String, ByVal serviceType As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Material As New BusinessRules.Material(conn, ts)
            GetMaterial = Material.GetMaterial(itemNo, itemTypeNo, serviceType)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function DuplicateMaterial(ByVal sourceServiceType As String, ByVal destinationServiceType As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Material As New BusinessRules.Material(conn, ts)
            Material.DuplicateMaterial(sourceServiceType, destinationServiceType)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return dbEx.Message
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateMaterial(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Material As New BusinessRules.Material(conn, ts)
            UpdateMaterial = Material.UpdateMaterial(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function


    <WebMethod()> Public Function PQueryStatisticsRecommendProjectByMonth(ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal RecommendPerson As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PQueryStatisticsRecommendProjectByMonth = CommonQuery.PQueryStatisticsRecommendProjectByMonth(StartDate, EndDate, RecommendPerson, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function PQueryStatisticsRecommendProjectByYear(ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal RecommendPerson As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PQueryStatisticsRecommendProjectByYear = CommonQuery.PQueryStatisticsRecommendProjectByYear(StartDate, EndDate, RecommendPerson, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function PQueryStatisticsRecommendProject(ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal RecommendPerson As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PQueryStatisticsRecommendProject = CommonQuery.PQueryStatisticsRecommendProject(StartDate, EndDate, RecommendPerson, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function Usp_ListIsFirstLoanStat(ByVal dtFrom As DateTime, ByVal dtTo As DateTime, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            Usp_ListIsFirstLoanStat = CommonQuery.Usp_ListIsFirstLoanStat(dtFrom, dtTo, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function Usp_GetUnDealProject(ByVal serviceType As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            Usp_GetUnDealProject = CommonQuery.Usp_GetUnDealProject(serviceType, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function Usp_GetGuaranteeProject(ByVal LoanFrom As String, ByVal LoanTo As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            Usp_GetGuaranteeProject = CommonQuery.Usp_GetGuaranteeProject(LoanFrom, LoanTo, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function Usp_GetAfterGuaranteeRecord(ByVal corporationName As String, ByVal serviceType As String, ByVal managerA As String, ByVal dtCheckFrom As String, ByVal dtCheckTo As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            Usp_GetAfterGuaranteeRecord = CommonQuery.Usp_GetAfterGuaranteeRecord(corporationName, serviceType, managerA, dtCheckFrom, dtCheckTo, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function Usp_ListConsultation(ByVal corporation_code As String, ByVal corporation_name As String, ByVal district_name As String, _
            ByVal recommend_person As String, ByVal consult_person As String, ByVal dtConsultFrom As String, ByVal dtConsultTo As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            Usp_ListConsultation = CommonQuery.Usp_ListConsultation(corporation_code, corporation_name, district_name, recommend_person, consult_person, dtConsultFrom, dtConsultTo, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectCounterClaimInfo(ByVal strSQL_Condition_ProjectCounterClaim As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectCounterClaim As New ProjectCounterClaim(conn, ts)
            GetProjectCounterClaimInfo = ProjectCounterClaim.GetProjectCounterClaimInfo(strSQL_Condition_ProjectCounterClaim)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectCounterClaim(ByVal ProjectCounterClaimSet As DataSet)
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectCounterClaim As New BusinessRules.ProjectCounterClaim(conn, ts)
            UpdateProjectCounterClaim = ProjectCounterClaim.UpdateProjectCounterClaim(ProjectCounterClaimSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FQryUnsignProject(ByVal ProjectCode As String, ByVal CorpName As String, ByVal ServiceType As String, ByVal dtFrom As String, ByVal dtTo As String, ByVal phase As String, ByVal vchPMA As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FQryUnsignProject = CommonQuery.FQryUnsignProject(ProjectCode, CorpName, ServiceType, dtFrom, dtTo, phase, vchPMA, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function Usp_ListGuaranteeForm(ByVal vchProjectCode As String, _
            ByVal vchCorpName As String, ByVal dtSignFrom As String, ByVal dtSignTo As Object, ByVal dtLoanFrom As String, ByVal dtLoanTo As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            Usp_ListGuaranteeForm = CommonQuery.Usp_ListGuaranteeForm(vchProjectCode, vchCorpName, dtSignFrom, dtSignTo, dtLoanFrom, dtLoanTo, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetTOrganizationInfo(ByVal strSQL_Condition_TOrganization As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TOrganization As New BusinessRules.TOrganization(conn, ts)
            GetTOrganizationInfo = TOrganization.GetTOrganizationInfo(strSQL_Condition_TOrganization)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateTOrganization(ByVal dsCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TOrganization As New BusinessRules.TOrganization(conn, ts)
            UpdateTOrganization = TOrganization.UpdateTOrganization(dsCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectOrganization(ByVal strSQL_Condition_TOrganization As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectOrganization As New BusinessRules.ProjectOrganization(conn, ts)
            GetProjectOrganization = ProjectOrganization.GetProjectOrganization(strSQL_Condition_TOrganization)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectOrganization(ByVal dsCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectOrganization As New BusinessRules.ProjectOrganization(conn, ts)
            UpdateProjectOrganization = ProjectOrganization.UpdateProjectOrganization(dsCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectAppraisement(ByVal project_code As String, ByVal EnterpriseName As String, ByVal ServiceType As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectOrganization As New BusinessRules.CommonQuery(conn, ts)
            GetProjectAppraisement = ProjectOrganization.GetProjectAppraisement(project_code, EnterpriseName, ServiceType, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetCorporationAttendeePerson(ByVal projectCode As String, ByVal serviceType As String, ByVal role_id As String, ByVal applyPerson As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectOrganization As New BusinessRules.CommonQuery(conn, ts)
            GetCorporationAttendeePerson = ProjectOrganization.GetCorporationAttendeePerson(projectCode, serviceType, role_id, applyPerson)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetDefaultPerson(ByVal projectCode As String, ByVal role_id As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectOrganization As New BusinessRules.CommonQuery(conn, ts)
            GetDefaultPerson = ProjectOrganization.GetDefaultPerson(projectCode, role_id)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function
    <WebMethod()> Public Function PQueryProjectRequite(ByVal ProjectCode As String, ByVal Corporation As String, ByVal ServiceType As String, ByVal ManangerA As String, ByVal RefundType As String, ByVal IsNormal As String, ByVal IsPartion As String, ByVal objDate As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim QueryProjectRequite As New BusinessRules.CommonQuery(conn, ts)
            PQueryProjectRequite = QueryProjectRequite.PQueryProjectRequite(ProjectCode, Corporation, ServiceType, ManangerA, RefundType, IsNormal, IsPartion, objDate, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function PQueryIntentLetter(ByVal strCondition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim QueryProjectRequite As New BusinessRules.CommonQuery(conn, ts)
            PQueryIntentLetter = QueryProjectRequite.GetIntentLetter(strCondition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function PQueryIntentLetterInfo(ByVal PutOutType As String, ByVal signStartDate As String, ByVal signEndDate As String, _
            ByVal issueStartDate As String, ByVal issueEndDate As String, ByVal userName As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim QueryProjectRequite As New BusinessRules.CommonQuery(conn, ts)
            PQueryIntentLetterInfo = QueryProjectRequite.PQueryIntentLetterInfo(PutOutType, signStartDate, signEndDate, issueStartDate, issueEndDate, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function PCopyOppGuarantee(ByVal ProjectCode As String, ByVal SourceProjectCode As String, ByVal SourceSerialNum As String, ByVal CreatePerson As String, ByVal CreateDate As Date) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PCopyOppGuarantee = CommonQuery.PCopyOppGuarantee(ProjectCode, SourceProjectCode, SourceSerialNum, CreatePerson, CreateDate)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetGuarantyInfoEx(ByVal ProjectCode As String, ByVal CorporationName As String, ByVal ItemValue As String, ByVal OppGuaranteeForm As String, ByVal EvaluateDate As Object, ByVal Status As String, ByVal GuarantyType As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetGuarantyInfoEx = CommonQuery.GetGuarantyInfoEx(ProjectCode, CorporationName, ItemValue, OppGuaranteeForm, EvaluateDate, Status, GuarantyType)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetProjectResponsibleInfo(ByVal strSQL_Condition_ProjectResponsible As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectResponsible As New BusinessRules.ProjectResponsible(conn, ts)
            GetProjectResponsibleInfo = ProjectResponsible.GetProjectResponsibleInfo(strSQL_Condition_ProjectResponsible)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateProjectResponsible(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ProjectResponsible As New BusinessRules.ProjectResponsible(conn, ts)
            UpdateProjectResponsible = ProjectResponsible.UpdateProjectResponsible(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function PQueryOppEvaluate(ByVal ProjectCode As String, ByVal CorporationName As String, _
        ByVal ManagerA As String, ByVal Evaluater As String, ByVal EvaluateStatus As String, _
        ByVal BookFrom As String, ByVal BookTo As String, ByVal AffirmFrom As String, ByVal AffirmTo As String, ByVal userName As String) As DataSet
        Dim tempDs As New DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            PQueryOppEvaluate = CommonQuery.PQueryOppEvaluate(ProjectCode, CorporationName, ManagerA, Evaluater, EvaluateStatus, BookFrom, BookTo, AffirmFrom, AffirmTo, userName)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '-----------------------workflow-------------------------
    <WebMethod()> Public Function GetWfProjectMessagesInfo(ByVal strSQL As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectMessages As New BusinessRules.WfProjectMessages(conn, ts)
            GetWfProjectMessagesInfo = WfProjectMessages.GetWfProjectMessagesInfo(strSQL)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWfProjectMessages(ByVal ds As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectMessages As New BusinessRules.WfProjectMessages(conn, ts)
            WfProjectMessages.UpdateWfProjectMessages(ds)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWfProjectTaskTransferInfo(ByVal strSQL As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTaskTransfer As New BusinessRules.WfProjectTaskTransfer(conn, ts)
            GetWfProjectTaskTransferInfo = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSQL)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWfProjectTaskTransfer(ByVal ds As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTaskTransfer As New BusinessRules.WfProjectTaskTransfer(conn, ts)
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(ds)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWfProjectTimingTaskInfo(ByVal strSQL As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTimingTask As New BusinessRules.WfProjectTimingTask(conn, ts)
            GetWfProjectTimingTaskInfo = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSQL)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWfProjectTimingTask(ByVal ds As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTimingTask As New BusinessRules.WfProjectTimingTask(conn, ts)
            WfProjectTimingTask.UpdateWfProjectTimingTask(ds)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWfProjectTrackInfo(ByVal strSQL As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTrack As New BusinessRules.WfProjectTrack(conn, ts)
            GetWfProjectTrackInfo = WfProjectTrack.GetWfProjectTrackInfo(strSQL)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWfProjectTrack(ByVal ds As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTrack As New BusinessRules.WfProjectTrack(conn, ts)
            WfProjectTrack.UpdateWfProjectTrack(ds)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetWfProjectTaskInfo(ByVal strSQL As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTask As New BusinessRules.WfProjectTask(conn, ts)
            GetWfProjectTaskInfo = WfProjectTask.GetWfProjectTaskInfo(strSQL)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateWfProjectTask(ByVal ds As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WfProjectTask As New BusinessRules.WfProjectTask(conn, ts)
            WfProjectTask.UpdateWfProjectTask(ds)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function AddMsg(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal tmpStaffID As String, ByVal messageID As String, ByVal readFlag As String)
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim TimingServer As TimingServer = New TimingServer(conn, ts, True, True)

            AddMsg = TimingServer.AddMsg(workFlowID, projectID, taskID, tmpStaffID, messageID, readFlag)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function
    '----------------------workflow end 2004-6-15 qxd ------------------------------

    '2004-05-09 by Popeye Zhong
    '获取财务分析
    <WebMethod()> Public Function FetchProjectFinanceAnalyse(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim financeAnalyse As New BusinessRules.FinanceAnalyse(conn, ts)
            FetchProjectFinanceAnalyse = financeAnalyse.FetchProjectFinanceAnalyse(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(MessageName:="FetchProjectFinanceAnalyseEx")> Public Function FetchProjectFinanceAnalyse(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim financeAnalyse As New BusinessRules.FinanceAnalyse(conn, ts)
            FetchProjectFinanceAnalyse = financeAnalyse.FetchProjectFinanceAnalyse(ProjectNo, CorporationNo, Phase, Month, MonthLast)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2004-05-09 by Popeye Zhong
    '创建财务分析
    <WebMethod()> Public Function CreateProjectFinanceAnalyse(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal Month As String, ByVal MonthLast As String) As Boolean
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim financeAnalyse As New BusinessRules.FinanceAnalyse(conn, ts)
            CreateProjectFinanceAnalyse = financeAnalyse.CreateProjectFinanceAnalyse(ProjectNo, CorporationNo, Phase, Month, MonthLast)
            ts.Commit()
        Catch
            ts.Rollback()
            Return False
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchProjectFinanceAnalyseIntegration(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String, ByVal ThisYear As String, ByVal LastYear1 As String, ByVal LastYear2 As String, ByVal LastYear3 As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            FetchProjectFinanceAnalyseIntegration = CommonQuery.FetchProjectFinanceAnalyseIntegration(ProjectNo, CorporationNo, Phase, ThisYear, LastYear1, LastYear2, LastYear3)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '2004-05-11 by Popeye Zhong
    '获取财务分析的指标记录
    <WebMethod()> Public Function FetchFinanceAnalyseIndex(ByVal Condition As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim financeAnalyse As New BusinessRules.FinanceAnalyse(conn, ts)
            FetchFinanceAnalyseIndex = financeAnalyse.FetchFinanceAnalyseIndex(Condition)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(MessageName:="FetchFinanceAnalyseIndexEx")> Public Function FetchFinanceAnalyseIndex(ByVal IndexType As String, ByVal IndexID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim financeAnalyse As New BusinessRules.FinanceAnalyse(conn, ts)
            FetchFinanceAnalyseIndex = financeAnalyse.FetchFinanceAnalyseIndex(IndexType, IndexID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function DeleteProjectFinanceAnalyse(ByVal ProjectNo As String, ByVal CorporationNo As String, ByVal Phase As String) As Integer
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim financeAnalyse As New BusinessRules.FinanceAnalyse(conn, ts)
            DeleteProjectFinanceAnalyse = financeAnalyse.DeleteProjectFinanceAnalyse(ProjectNo, CorporationNo, Phase)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function updateProcessEx() As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.updateProcess()
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod(messagename:="updateProcessExByCode")> Public Function updateProcessEx(ByVal ProjectCode As String) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim WorkFlow As New BusinessRules.WorkFlow(conn, ts)
            WorkFlow.updateProcess(ProjectCode)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '获取岗位和工作职责
    <WebMethod()> Public Function GetPostAndJobResponsibilityInfo(ByVal strSQL_Condition_Post As String, ByVal strSQL_Condition_JobResponsibility As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oPostAndResponsibility As New BusinessRules.PostAndResponsibility(conn, ts)
            GetPostAndJobResponsibilityInfo = oPostAndResponsibility.GetPostAndJobResponsibilityInfo(strSQL_Condition_Post, strSQL_Condition_JobResponsibility)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function
    '保存岗位和工作职责
    <WebMethod()> Public Function UpdatePostAndJobResponsibility(ByVal commitSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oPostAndResponsibility As New BusinessRules.PostAndResponsibility(conn, ts)
            oPostAndResponsibility.UpdatePostAndJobResponsibility(commitSet)
            ts.Commit()
            Return "1"
        Catch errWf As WorkFlowErr
            ts.Rollback()
            Return errWf.ErrMessage
        Catch e As Exception
            ts.Rollback()
            Return e.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function
    <WebMethod()> Public Function GetUserPostInfo(ByVal strSQL_Condition_UserPost As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oUserPost As New BusinessRules.UserPost(conn, ts)
            GetUserPostInfo = oUserPost.GetUserPostInfo(strSQL_Condition_UserPost)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '更新假期信息
    <WebMethod()> Public Function UpdateUserPost(ByVal UserPostSet As DataSet)
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oUserPost As New BusinessRules.UserPost(conn, ts)
            oUserPost.UpdateUserPost(UserPostSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '获取工作时间表信息
    <WebMethod()> Public Function GetJobPeriodInfo(ByVal strSQL_Condition_JobPeriod As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oJobPeriod As New BusinessRules.JobPeriod(conn, ts)
            GetJobPeriodInfo = oJobPeriod.GetJobPeriodInfo(strSQL_Condition_JobPeriod)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function
    '保存工作时间
    <WebMethod()> Public Function UpdateJobPeriod(ByVal JobPeriodSet As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim oJobPeriod As New BusinessRules.JobPeriod(conn, ts)
            oJobPeriod.UpdateJobPeriod(JobPeriodSet)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    '获取工作日志
    <WebMethod()> Public Function GetWorkingHours(ByVal staff_name As String, ByVal start_date As Object, ByVal end_date As Object, _
                        ByVal period As String, ByVal statisticsType As Integer) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim comm As New BusinessRules.CommonQuery(conn, ts)
            GetWorkingHours = comm.GetWorkingHours(staff_name, start_date, end_date, period, statisticsType)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetMoneyInfo(ByVal moneyID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Money As New BusinessRules.Money(conn, ts)
            GetMoneyInfo = Money.GetMoneyInfo(moneyID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function UpdateMoney(ByVal dstCommit As DataSet) As String
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim Money As New BusinessRules.Money(conn, ts)
            UpdateMoney = Money.UpdateMoney(dstCommit)
            ts.Commit()
            Return "1"
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    'qxd add 2005-3-17 反担保物查询
    <WebMethod()> Public Function GetQueryOppGuarantInfo(ByVal projectCode As String, ByVal corporationName As String, ByVal oppForm As String, _
                        ByVal oppStatus As String, ByVal itemType As String, ByVal itemCodeFirst As String, _
                        ByVal itemValueFirst As String, ByVal itemCodeSecond As String, ByVal itemValueSecond As String, _
                        ByVal startDate As String, ByVal endDate As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim comm As New BusinessRules.CommonQuery(conn, ts)
            GetQueryOppGuarantInfo = comm.GetQueryOppGuarantInfo(projectCode, corporationName, oppForm, _
                         oppStatus, itemType, itemCodeFirst, _
                         itemValueFirst, itemCodeSecond, itemValueSecond, _
                         startDate, endDate)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function GetGuarantingCorporationList(ByVal start_date As Object, ByVal end_date As Object) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim comm As New BusinessRules.CommonQuery(conn, ts)
            GetGuarantingCorporationList = comm.GetGuarantingCorporationList(start_date, end_date)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

    <WebMethod()> Public Function FetchConfernceRoom(ByVal ConfernceRoomID As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ConfernceRoom As New BusinessRules.ConfernceRoom(conn, ts)
            FetchConfernceRoom = ConfernceRoom.FetchConfernceRoom(ConfernceRoomID)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try

    End Function

    <WebMethod()> Public Function UpdateConfernceRoom(ByVal dstCommit As DataSet) As Int32
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim ConfernceRoom As New BusinessRules.ConfernceRoom(conn, ts)
            ConfernceRoom.UpdateConfernceRoom(dstCommit)
            ts.Commit()
            Return 1
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Return DataBaseErr.UpdateCommandErr
        Catch oEx As Exception
            ts.Rollback()
            Return oEx.Message
        Finally
            conn.Close()
            conn.Dispose()
        End Try
        End Function

End Class

End Namespace
