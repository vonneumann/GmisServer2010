Option Explicit On 
Imports System.IO
Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports BusinessRules.OAWorkflowXYDB
Imports ICSharpCode.SharpZipLib

Public Class ImplOAReviewConclusion
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    Private project As Project

    Private OAWorkflowXYDB As OAWorkflowXYDB.WorkflowServiceForXYDB
    Private webserviceCgmisForOA As WebserviceCgmisForOA.ServiceOA

    Private commonquery As CommonQuery

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        project = New Project(conn, ts)

        OAWorkflowXYDB = New OAWorkflowXYDB.WorkflowServiceForXYDB()
        webserviceCgmisForOA = New WebserviceCgmisForOA.ServiceOA()
        commonquery = New CommonQuery(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        Dim dsTemp As DataSet

        '获取企业名称
        strSql = "select 企业名称 ,放款银行,担保金额,期限,担保费率,处理A角,处理B角,上会通过日期 from queryProjectInfoForStatistics_Chinese where 项目编码='" + projectID + "'"
        dsTemp = commonquery.GetCommonQueryInfo(strSql)
        Dim strCorporationName As String = dsTemp.Tables(0).Rows(0).Item("企业名称")

        '获取银行名称
        Dim strBankName As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("放款银行")), "", dsTemp.Tables(0).Rows(0).Item("放款银行"))

        '获取担保金额
        Dim strGuaranteeSum As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("担保金额")), "0", dsTemp.Tables(0).Rows(0).Item("担保金额"))

        '获取担保期限
        Dim strGuaranteeTerm As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("期限")), "", dsTemp.Tables(0).Rows(0).Item("期限"))

        '获取担保费率
        Dim strGuaranteeRate As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("担保费率")), 0, dsTemp.Tables(0).Rows(0).Item("担保费率"))
        strGuaranteeRate = String.Format("{0:F2}", Convert.ToDecimal(strGuaranteeRate))

        'A角
        Dim strManagerA As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("处理A角")), 0, dsTemp.Tables(0).Rows(0).Item("处理A角"))


        'B角
        Dim strManagerB As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("处理B角")), 0, dsTemp.Tables(0).Rows(0).Item("处理B角"))

        '评审会日期
        Dim dConferenceDate As Date = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("上会通过日期")), "", dsTemp.Tables(0).Rows(0).Item("上会通过日期"))
        Dim strConferenceDate As String = dConferenceDate.ToString("yyyy-MM-dd")

        '把A角作为申请人 ,获取A角的OA对应ID


        strSql = "select id,departmentid from HrmResource where lastname='" + strManagerA + "'"
        dsTemp = WebserviceCgmisForOA.GetCommonQueryInfoForOA(strSql)
        Dim iStaffID As Integer = dsTemp.Tables(0).Rows(0).Item("id")
        Dim iDepID As Integer = dsTemp.Tables(0).Rows(0).Item("departmentid")


        '获取B角OA对应ID
        strSql = "select id,departmentid from HrmResource where lastname='" + strManagerB + "'"
        dsTemp = webserviceCgmisForOA.GetCommonQueryInfoForOA(strSql)
        Dim iStaffBID As Integer = dsTemp.Tables(0).Rows(0).Item("id")
        Dim iDepBID As Integer = dsTemp.Tables(0).Rows(0).Item("departmentid")




        '创建主表字段
        Dim wrti(12) As WorkflowRequestTableField

        '申请人
        wrti(0) = New WorkflowRequestTableField()
        wrti(0).fieldName = "shenqr"
        wrti(0).fieldValue = iStaffID
        wrti(0).view = True
        wrti(0).edit = True
        wrti(0).viewSpecified = True
        wrti(0).editSpecified = True

        '申请部门
        wrti(1) = New WorkflowRequestTableField()
        wrti(1).fieldName = "shenqbm"
        wrti(1).fieldValue = iDepID
        wrti(1).view = True
        wrti(1).edit = True
        wrti(1).viewSpecified = True
        wrti(1).editSpecified = True


        '申请日期
        wrti(2) = New WorkflowRequestTableField()
        wrti(2).fieldName = "shenqrq"
        wrti(2).fieldValue = Now.ToString("yyyy-MM-dd")
        wrti(2).view = True
        wrti(2).edit = True
        wrti(2).viewSpecified = True
        wrti(2).editSpecified = True


        '业务品种
        wrti(3) = New WorkflowRequestTableField()
        wrti(3).fieldName = "yewpz"
        wrti(3).fieldValue = GetServiceType(projectID, taskID)
        wrti(3).view = True
        wrti(3).edit = True
        wrti(3).viewSpecified = True
        wrti(3).editSpecified = True


        '项目编号
        wrti(4) = New WorkflowRequestTableField()
        wrti(4).fieldName = "xiangmbh"
        wrti(4).fieldValue = projectID
        wrti(4).view = True
        wrti(4).edit = True
        wrti(4).viewSpecified = True
        wrti(4).editSpecified = True


        '企业名称
        wrti(5) = New WorkflowRequestTableField()
        wrti(5).fieldName = "qiymc"
        wrti(5).fieldValue = strCorporationName
        wrti(5).view = True
        wrti(5).edit = True
        wrti(5).viewSpecified = True
        wrti(5).editSpecified = True


        '担保金额
        wrti(6) = New WorkflowRequestTableField()
        wrti(6).fieldName = "jine"
        wrti(6).fieldValue = strGuaranteeSum
        wrti(6).view = True
        wrti(6).edit = True
        wrti(6).viewSpecified = True
        wrti(6).editSpecified = True

        '担保费率
        wrti(7) = New WorkflowRequestTableField()
        wrti(7).fieldName = "feilv"
        wrti(7).fieldValue = strGuaranteeRate
        wrti(7).view = True
        wrti(7).edit = True
        wrti(7).viewSpecified = True
        wrti(7).editSpecified = True

        '担保期限
        wrti(8) = New WorkflowRequestTableField()
        wrti(8).fieldName = "qixian"
        wrti(8).fieldValue = strGuaranteeTerm
        wrti(8).view = True
        wrti(8).edit = True
        wrti(8).viewSpecified = True
        wrti(8).editSpecified = True


        '评审意见表附件
        Dim strDocName As String = GetConfTrialDoc(projectID)
        wrti(9) = New WorkflowRequestTableField()
        wrti(9).fieldName = "fujsc"
        wrti(9).fieldType = "http:" + strDocName
        wrti(9).fieldValue = "http://192.168.80.84/webservice_cgmis/tempdoc/" + strDocName
        wrti(9).view = True
        wrti(9).edit = True
        wrti(9).viewSpecified = True
        wrti(9).editSpecified = True

        'A角
        wrti(10) = New WorkflowRequestTableField()
        wrti(10).fieldName = "xiangmaj"
        wrti(10).fieldValue = iStaffID
        wrti(10).view = True
        wrti(10).edit = True
        wrti(10).viewSpecified = True
        wrti(10).editSpecified = True

        'B角
        wrti(11) = New WorkflowRequestTableField()
        wrti(11).fieldName = "xiangmbj"
        wrti(11).fieldValue = iStaffBID
        wrti(11).view = True
        wrti(11).edit = True
        wrti(11).viewSpecified = True
        wrti(11).editSpecified = True


        '评审会日期
        wrti(12) = New WorkflowRequestTableField()
        wrti(12).fieldName = "pingshrq"
        wrti(12).fieldValue = strConferenceDate
        wrti(12).view = True
        wrti(12).edit = True
        wrti(12).viewSpecified = True
        wrti(12).editSpecified = True


        '主字段只有一行数据
        Dim wrtri(0) As WorkflowRequestTableRecord
        wrtri(0) = New WorkflowRequestTableRecord
        wrtri(0).workflowRequestTableFields = wrti

        Dim wmi As New WorkflowMainTableInfo()
        wmi.requestRecords = wrtri


        Dim wbi As New WorkflowBaseInfo()
        wbi.workflowId = "63"
        Dim wri As New WorkflowRequestInfo()
        wri.creatorId = "56"
        wri.requestName = strCorporationName + "评审意见表签字流程"
        wri.workflowMainTableInfo = wmi
        wri.workflowBaseInfo = wbi

        Dim result As String
        result = OAWorkflowXYDB.doCreateWorkflowRequest(wri, 56)

        'OA中创建评审意见表会签流程

        'Try
        '    '2008-5-5 yjf edit 修改项目信息中取文档不取DOCUMENT字段后，需要重新获取的问题

        '    Dim strPath As String
        '    strPath = AppDomain.CurrentDomain.BaseDirectory + "tempdoc\bbb.doc"
        '    Dim dsDocument As DataSet = commonquery.GetCommonQueryInfo("select top 1 document from project_files where project_code='" + projectID + "' and item_type='45' and item_code='011' order by [date] desc")
        '    Dim document As Object = dsDocument.Tables(0).Rows(0).Item("document")

        '    If IsDBNull(document) Then
        '        Exit Function
        '    End If
        '    Dim data As Byte() = CType(document, Byte())

        '    Dim fs = New System.IO.FileStream(strPath, IO.FileMode.Create, IO.FileAccess.ReadWrite, IO.FileShare.ReadWrite, 5)
        '    fs.Write(data, 0, data.Length)
        '    fs.Close()
        '    'System.Diagnostics.Process.Start(strPath)
        'Catch ex As Exception

        'Finally
        '    GC.Collect()
        'End Try


    End Function

    Private Function GetServiceType(ByVal projectID As String, ByVal taskID As String) As Integer

        Dim objWorkflowTask As New WfProjectTask(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTask As DataSet = objWorkflowTask.GetWfProjectTaskInfo(strSql)
        Dim workflowNum As String = dsTempTask.Tables(0).Rows(0).Item("workflow_id")
        Select Case workflowNum
            Case "02"
                Return 0
            Case "03"
                Return 1
            Case "18"
                Return 3
            Case "10"
                Return 2
            Case "08"
                Return 2
            Case Else
                Return 4
        End Select
    End Function


    '获取评审意见表附件名字
    Private Function GetConfTrialDoc(ByVal projectID As String) As String


        Try
            Dim strPath As String
            strPath = AppDomain.CurrentDomain.BaseDirectory + "tempdoc\"

            Dim dsDocument As DataSet = commonquery.GetCommonQueryInfo("select top 1 document,title from project_files where project_code='" + projectID + "' and item_type='45' and item_code='011' order by [date] desc")
            Dim document As Object = dsDocument.Tables(0).Rows(0).Item("document")

            Dim docName As String = dsDocument.Tables(0).Rows(0).Item("title") + ".doc"

            strPath = strPath + docName

            If IsDBNull(document) Then
                Exit Function
            End If
            Dim data As Byte() = CType(document, Byte())

            Dim fs = New System.IO.FileStream(strPath, IO.FileMode.Create, IO.FileAccess.ReadWrite, IO.FileShare.ReadWrite, 5)
            fs.Write(data, 0, data.Length)
            fs.Close()
            'System.Diagnostics.Process.Start(strPath)
            Return docName
        Catch ex As Exception

        Finally
            GC.Collect()
        End Try
    End Function

End Class
