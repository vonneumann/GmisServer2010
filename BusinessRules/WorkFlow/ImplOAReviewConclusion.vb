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

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    Private project As Project

    Private OAWorkflowXYDB As OAWorkflowXYDB.WorkflowServiceForXYDB
    Private webserviceCgmisForOA As WebserviceCgmisForOA.ServiceOA

    Private commonquery As CommonQuery

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        project = New Project(conn, ts)

        OAWorkflowXYDB = New OAWorkflowXYDB.WorkflowServiceForXYDB()
        webserviceCgmisForOA = New WebserviceCgmisForOA.ServiceOA()
        commonquery = New CommonQuery(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        Dim dsTemp As DataSet

        '��ȡ��ҵ����
        strSql = "select ��ҵ���� ,�ſ�����,�������,����,��������,����A��,����B��,�ϻ�ͨ������ from queryProjectInfoForStatistics_Chinese where ��Ŀ����='" + projectID + "'"
        dsTemp = commonquery.GetCommonQueryInfo(strSql)
        Dim strCorporationName As String = dsTemp.Tables(0).Rows(0).Item("��ҵ����")

        '��ȡ��������
        Dim strBankName As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("�ſ�����")), "", dsTemp.Tables(0).Rows(0).Item("�ſ�����"))

        '��ȡ�������
        Dim strGuaranteeSum As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("�������")), "0", dsTemp.Tables(0).Rows(0).Item("�������"))

        '��ȡ��������
        Dim strGuaranteeTerm As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("����")), "", dsTemp.Tables(0).Rows(0).Item("����"))

        '��ȡ��������
        Dim strGuaranteeRate As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("��������")), 0, dsTemp.Tables(0).Rows(0).Item("��������"))
        strGuaranteeRate = String.Format("{0:F2}", Convert.ToDecimal(strGuaranteeRate))

        'A��
        Dim strManagerA As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("����A��")), 0, dsTemp.Tables(0).Rows(0).Item("����A��"))


        'B��
        Dim strManagerB As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("����B��")), 0, dsTemp.Tables(0).Rows(0).Item("����B��"))

        '���������
        Dim dConferenceDate As Date = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("�ϻ�ͨ������")), "", dsTemp.Tables(0).Rows(0).Item("�ϻ�ͨ������"))
        Dim strConferenceDate As String = dConferenceDate.ToString("yyyy-MM-dd")

        '��A����Ϊ������ ,��ȡA�ǵ�OA��ӦID


        strSql = "select id,departmentid from HrmResource where lastname='" + strManagerA + "'"
        dsTemp = WebserviceCgmisForOA.GetCommonQueryInfoForOA(strSql)
        Dim iStaffID As Integer = dsTemp.Tables(0).Rows(0).Item("id")
        Dim iDepID As Integer = dsTemp.Tables(0).Rows(0).Item("departmentid")


        '��ȡB��OA��ӦID
        strSql = "select id,departmentid from HrmResource where lastname='" + strManagerB + "'"
        dsTemp = webserviceCgmisForOA.GetCommonQueryInfoForOA(strSql)
        Dim iStaffBID As Integer = dsTemp.Tables(0).Rows(0).Item("id")
        Dim iDepBID As Integer = dsTemp.Tables(0).Rows(0).Item("departmentid")




        '���������ֶ�
        Dim wrti(12) As WorkflowRequestTableField

        '������
        wrti(0) = New WorkflowRequestTableField()
        wrti(0).fieldName = "shenqr"
        wrti(0).fieldValue = iStaffID
        wrti(0).view = True
        wrti(0).edit = True
        wrti(0).viewSpecified = True
        wrti(0).editSpecified = True

        '���벿��
        wrti(1) = New WorkflowRequestTableField()
        wrti(1).fieldName = "shenqbm"
        wrti(1).fieldValue = iDepID
        wrti(1).view = True
        wrti(1).edit = True
        wrti(1).viewSpecified = True
        wrti(1).editSpecified = True


        '��������
        wrti(2) = New WorkflowRequestTableField()
        wrti(2).fieldName = "shenqrq"
        wrti(2).fieldValue = Now.ToString("yyyy-MM-dd")
        wrti(2).view = True
        wrti(2).edit = True
        wrti(2).viewSpecified = True
        wrti(2).editSpecified = True


        'ҵ��Ʒ��
        wrti(3) = New WorkflowRequestTableField()
        wrti(3).fieldName = "yewpz"
        wrti(3).fieldValue = GetServiceType(projectID, taskID)
        wrti(3).view = True
        wrti(3).edit = True
        wrti(3).viewSpecified = True
        wrti(3).editSpecified = True


        '��Ŀ���
        wrti(4) = New WorkflowRequestTableField()
        wrti(4).fieldName = "xiangmbh"
        wrti(4).fieldValue = projectID
        wrti(4).view = True
        wrti(4).edit = True
        wrti(4).viewSpecified = True
        wrti(4).editSpecified = True


        '��ҵ����
        wrti(5) = New WorkflowRequestTableField()
        wrti(5).fieldName = "qiymc"
        wrti(5).fieldValue = strCorporationName
        wrti(5).view = True
        wrti(5).edit = True
        wrti(5).viewSpecified = True
        wrti(5).editSpecified = True


        '�������
        wrti(6) = New WorkflowRequestTableField()
        wrti(6).fieldName = "jine"
        wrti(6).fieldValue = strGuaranteeSum
        wrti(6).view = True
        wrti(6).edit = True
        wrti(6).viewSpecified = True
        wrti(6).editSpecified = True

        '��������
        wrti(7) = New WorkflowRequestTableField()
        wrti(7).fieldName = "feilv"
        wrti(7).fieldValue = strGuaranteeRate
        wrti(7).view = True
        wrti(7).edit = True
        wrti(7).viewSpecified = True
        wrti(7).editSpecified = True

        '��������
        wrti(8) = New WorkflowRequestTableField()
        wrti(8).fieldName = "qixian"
        wrti(8).fieldValue = strGuaranteeTerm
        wrti(8).view = True
        wrti(8).edit = True
        wrti(8).viewSpecified = True
        wrti(8).editSpecified = True


        '�����������
        Dim strDocName As String = GetConfTrialDoc(projectID)
        wrti(9) = New WorkflowRequestTableField()
        wrti(9).fieldName = "fujsc"
        wrti(9).fieldType = "http:" + strDocName
        wrti(9).fieldValue = "http://192.168.80.84/webservice_cgmis/tempdoc/" + strDocName
        wrti(9).view = True
        wrti(9).edit = True
        wrti(9).viewSpecified = True
        wrti(9).editSpecified = True

        'A��
        wrti(10) = New WorkflowRequestTableField()
        wrti(10).fieldName = "xiangmaj"
        wrti(10).fieldValue = iStaffID
        wrti(10).view = True
        wrti(10).edit = True
        wrti(10).viewSpecified = True
        wrti(10).editSpecified = True

        'B��
        wrti(11) = New WorkflowRequestTableField()
        wrti(11).fieldName = "xiangmbj"
        wrti(11).fieldValue = iStaffBID
        wrti(11).view = True
        wrti(11).edit = True
        wrti(11).viewSpecified = True
        wrti(11).editSpecified = True


        '���������
        wrti(12) = New WorkflowRequestTableField()
        wrti(12).fieldName = "pingshrq"
        wrti(12).fieldValue = strConferenceDate
        wrti(12).view = True
        wrti(12).edit = True
        wrti(12).viewSpecified = True
        wrti(12).editSpecified = True


        '���ֶ�ֻ��һ������
        Dim wrtri(0) As WorkflowRequestTableRecord
        wrtri(0) = New WorkflowRequestTableRecord
        wrtri(0).workflowRequestTableFields = wrti

        Dim wmi As New WorkflowMainTableInfo()
        wmi.requestRecords = wrtri


        Dim wbi As New WorkflowBaseInfo()
        wbi.workflowId = "63"
        Dim wri As New WorkflowRequestInfo()
        wri.creatorId = "56"
        wri.requestName = strCorporationName + "���������ǩ������"
        wri.workflowMainTableInfo = wmi
        wri.workflowBaseInfo = wbi

        Dim result As String
        result = OAWorkflowXYDB.doCreateWorkflowRequest(wri, 56)

        'OA�д�������������ǩ����

        'Try
        '    '2008-5-5 yjf edit �޸���Ŀ��Ϣ��ȡ�ĵ���ȡDOCUMENT�ֶκ���Ҫ���»�ȡ������

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


    '��ȡ���������������
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
