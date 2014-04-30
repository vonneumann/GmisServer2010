Imports BusinessRules
Imports System.Data.SqlClient
Imports System.Xml
Imports ServerClient.WebReference
Imports ServerClient.WebReference1
Imports System.Collections.Generic
Imports BusinessRules.OAWorkflowXYDB
Public Class frmTest
    Inherits System.Windows.Forms.Form

#Region " Windows 窗体设计器生成的代码 "

    Public Sub New()
        MyBase.New()

        '该调用是 Windows 窗体设计器所必需的。
        InitializeComponent()

        '在 InitializeComponent() 调用之后添加任何初始化

    End Sub

    '窗体重写处置以清理组件列表。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意：以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改此过程。
    '不要使用代码编辑器修改它。
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents DataGrid2 As System.Windows.Forms.DataGrid
    Friend WithEvents txtFinishedTaskID As System.Windows.Forms.TextBox
    Friend WithEvents txtFinishedFlag As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents task_Timer As System.Timers.Timer
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtCreatePerson As System.Windows.Forms.TextBox
    Friend WithEvents txtProjectID As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.DataGrid2 = New System.Windows.Forms.DataGrid()
        Me.txtFinishedTaskID = New System.Windows.Forms.TextBox()
        Me.txtFinishedFlag = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtUserID = New System.Windows.Forms.TextBox()
        Me.task_Timer = New System.Timers.Timer()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtCreatePerson = New System.Windows.Forms.TextBox()
        Me.txtProjectID = New System.Windows.Forms.TextBox()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.task_Timer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(0, 120)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(544, 120)
        Me.DataGrid1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(80, 369)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Create"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(224, 369)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Finished"
        '
        'DataGrid2
        '
        Me.DataGrid2.DataMember = ""
        Me.DataGrid2.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid2.Location = New System.Drawing.Point(0, 240)
        Me.DataGrid2.Name = "DataGrid2"
        Me.DataGrid2.Size = New System.Drawing.Size(544, 120)
        Me.DataGrid2.TabIndex = 3
        '
        'txtFinishedTaskID
        '
        Me.txtFinishedTaskID.Location = New System.Drawing.Point(144, 56)
        Me.txtFinishedTaskID.Name = "txtFinishedTaskID"
        Me.txtFinishedTaskID.Size = New System.Drawing.Size(100, 21)
        Me.txtFinishedTaskID.TabIndex = 4
        '
        'txtFinishedFlag
        '
        Me.txtFinishedFlag.Location = New System.Drawing.Point(432, 56)
        Me.txtFinishedFlag.Name = "txtFinishedFlag"
        Me.txtFinishedFlag.Size = New System.Drawing.Size(100, 21)
        Me.txtFinishedFlag.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(0, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "FinishedTaskID:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(320, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 23)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "FinishedFlag:"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(0, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 23)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "AttendeePerson::"
        '
        'txtUserID
        '
        Me.txtUserID.Location = New System.Drawing.Point(144, 88)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(100, 21)
        Me.txtUserID.TabIndex = 8
        '
        'task_Timer
        '
        Me.task_Timer.Enabled = True
        Me.task_Timer.Interval = 5000.0R
        Me.task_Timer.SynchronizingObject = Me
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(352, 369)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(88, 23)
        Me.Button3.TabIndex = 10
        Me.Button3.Text = "TimingServer"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(320, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 23)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "CreatePerson:"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(0, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 23)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "ProjectID:"
        '
        'txtCreatePerson
        '
        Me.txtCreatePerson.Location = New System.Drawing.Point(432, 8)
        Me.txtCreatePerson.Name = "txtCreatePerson"
        Me.txtCreatePerson.Size = New System.Drawing.Size(100, 21)
        Me.txtCreatePerson.TabIndex = 12
        '
        'txtProjectID
        '
        Me.txtProjectID.Location = New System.Drawing.Point(144, 8)
        Me.txtProjectID.Name = "txtProjectID"
        Me.txtProjectID.Size = New System.Drawing.Size(100, 21)
        Me.txtProjectID.TabIndex = 11
        '
        'frmTest
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(560, 397)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtCreatePerson)
        Me.Controls.Add(Me.txtProjectID)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtUserID)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFinishedFlag)
        Me.Controls.Add(Me.txtFinishedTaskID)
        Me.Controls.Add(Me.DataGrid2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DataGrid1)
        Me.Name = "frmTest"
        Me.Text = "d"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.task_Timer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    'Private sConnStr As String = "UID=sa;PWD=Pete1979;Initial Catalog=ABC;Data Source=192.168.0.100"
    Private sConnStr As String = "UID=sa;PWD=Xydb123456;Initial Catalog=XGMIS;Data Source=192.168.80.84\SQL2008"
    'Private sConnStr As String = "UID=sa;PWD=;Initial Catalog=gmis_2006;Data Source=127.0.0.1"
    'Private sConnStr As String = "UID=sa;PWD=;Initial Catalog=cgmis;Data Source=192.168.0.57"
    Private sConnStr2 As String = "dsn=OldData;uid=123;pwd=123"
    ' Private webSrv As New localhost5.Service1()
    'Private webs As New localhost2.Service1()
    Private dsTemp, dsTemp2 As DataSet
    Private dsMerge As New DataSet()
    Private dscommand As New SqlDataAdapter()
    Private strSql As String
    'Dim buss As BusinessRules.corporationAccess

    'Dim buss2 As BusinessRules.consultation


    'create button
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim tmpProjectID As String = Trim(txtProjectID.Text)
        Dim tmpCreatePerson As String = Trim(txtCreatePerson.Text)

        Dim tmpFinishedTaskID As String = Trim(txtFinishedTaskID.Text)
        Dim tmpFinishedFlag As String = Trim(txtFinishedFlag.Text)
        Dim tmpUserID As String = Trim(txtUserID.Text)
        Dim dbConnection As SqlConnection = New SqlConnection(sConnStr)

        dbConnection.Open()
        Dim ts As SqlTransaction
        ts = dbConnection.BeginTransaction
        'Dim buss As New BusinessRules.WorkFlow(dbConnection, ts)
        'Dim buss1 As New BusinessRules.CommonQuery(dbConnection, ts)
        'Dim buss2 As New BusinessRules.Pawn(dbConnection, ts)
        ''buss2.GetPawnInfo("null", "null", "null")
        'Dim buss3 As New BusinessRules.CommonQuery(dbConnection, ts)
        'dsTemp = buss3.GetTaskProjectList("CheckContract", "叶剑锋")
        Try
            'Dim strTest As String = buss.CreateProcess(tmpCreatePerson, tmpProjectID, tmpUserID, tmpFinishedFlag)
            'Dim strTest As String = buss.ReMeetingPlan(tmpProjectID)
            'Dim strTest As String = buss.CancelReviewConferencePlan(tmpProjectID)
            'Dim strTest As String = buss.CancelReviewConferencePlanProject(tmpProjectID)
            'Dim strTest As String = buss.CancelReviewConferencePlan("24")
            'dsTemp = buss1.GetQueryFirstProject("", "", "", "", "", "1900-01-01", "2050-12-30", "")

            'Dim tmpManagerA As String = dsTemp.Tables(0).Rows(0).Item("24")
            'Dim tmpManagerB As String = dsTemp.Tables(0).Rows(0).Item("25")

            'Dim strTest As String = buss.FinishedSignaturePlan(tmpProjectID)

            '创建OA工作流WEBSERVICE引用
            Dim OAWorkflowXYDB As New OAWorkflowXYDB.WorkflowServiceForXYDB

            '创建主表字段
            Dim wrti(0) As OAWorkflowXYDB.WorkflowRequestTableField

            '字段test赋值
            wrti(0) = New OAWorkflowXYDB.WorkflowRequestTableField()
            wrti(0).fieldName = "pingwei"
            wrti(0).fieldValue = "47,48,49"
            wrti(0).view = True
            wrti(0).edit = True
            wrti(0).viewSpecified = True
            wrti(0).editSpecified = True

            '主字段只有一行数据
            Dim wrtri(0) As OAWorkflowXYDB.WorkflowRequestTableRecord
            wrtri(0) = New OAWorkflowXYDB.WorkflowRequestTableRecord
            wrtri(0).workflowRequestTableFields = wrti

            Dim wmi As New OAWorkflowXYDB.WorkflowMainTableInfo()
            wmi.requestRecords = wrtri


            Dim wbi As New OAWorkflowXYDB.WorkflowBaseInfo()
            wbi.workflowId = "71"
            Dim wri As New OAWorkflowXYDB.WorkflowRequestInfo()
            wri.creatorId = "55"
            wri.requestName = "测试流程123345"
            wri.workflowMainTableInfo = wmi
            wri.workflowBaseInfo = wbi

            Dim result As String
            result = OAWorkflowXYDB.doCreateWorkflowRequest(wri, 55)

            'result = OAWorkflowXYDB.submitWorkflowRequest(wri, 459, 54, "submit", "结束审批")



            'Dim ri As New RequestInfo()
            'ri.creatorid = "56"
            'ri.workflowid = "66"
            'ri.remindtype = "0"
            'ri.requestlevel = "1"
            'ri.description = "createRequest方法测试"


            'Dim mti As New MainTableInfo()

            'Dim prts As New List(Of [Property])
            'Dim prt0 As New [Property]

            'prt0.name = "test"
            'prt0.value = "hello world"
            'prts.Add(prt0)

            'mti.property = prts.ToArray()
            'ri.mainTableInfo = mti


            'Dim client As New RequestService()

            'Dim requestID As String = client.createRequest(ri)


            ts.Commit()
        Catch ex As BusinessRules.WorkFlowErr
            ts.Rollback()
            MsgBox(ex.ErrMessage)
        End Try

        'Dim strTime As String = FormatDateTime(Now, DateFormat.ShortDate)
        'Dim dbConnection As SqlConnection = New SqlConnection(sConnStr)

        'dbConnection.Open()
        'Dim ts As SqlTransaction
        'ts = dbConnection.BeginTransaction
        'Dim buss As New BusinessRules.CommonQuery(dbConnection, ts)
        'Dim workflow As New WorkFlow(dbConnection, ts)
        'Try

        '    dsTemp = buss.GetProjectInfoEx("%")
        '    'dsTemp = workflow.LookUpWorking("周新华")
        '    'dsTemp.Tables(0).Rows(0).Delete()
        '    'dsTemp.AcceptChanges()
        '    ts.Commit()
        'Catch
        '    ts.Rollback()
        '    MsgBox(Err.Description())
        'End Try

        'Dim dbConnection As SqlConnection = New SqlConnection(sConnStr)

        'dbConnection.Open()
        'Dim ts As SqlTransaction
        'ts = dbConnection.BeginTransaction
        'Dim buss As New BusinessRules.CheckRecord(dbConnection, ts)
        'Dim ProjectCode As String = "000010305"
        'Dim ds As DataSet = buss.GetCheckRecordInfo("{project_code LIKE '" & ProjectCode & "'}", "{project_code LIKE '" & ProjectCode & "'}")
        'Dim dr As DataRow = ds.Tables("guarantee_check_record").NewRow
        'With dr
        '    .Item("project_code") = ProjectCode
        '    .Item("serial_num") = 1
        '    .Item("status") = False
        '    .Item("create_person") = "fff"
        '    .Item("create_date") = Now.Date
        '    .Item("check_date") = Now.Date

        'End With
        'ds.Tables("guarantee_check_record").Rows.Add(dr)
        'Try
        '    MessageBox.Show(buss.UpdateCheckRecordAlarm(ds.GetChanges))

        '    ts.Commit()
        'Catch
        '    ts.Rollback()
        '    MsgBox(Err.Description())
        'End Try


        'Dim tmpProjectID As String = Trim(txtProjectID.Text)
        'Dim tmpCreatePerson As String = Trim(txtCreatePerson.Text)

        'Dim tmpFinishedTaskID As String = Trim(txtFinishedTaskID.Text)
        'Dim tmpFinishedFlag As String = Trim(txtFinishedFlag.Text)
        'Dim tmpUserID As String = Trim(txtUserID.Text)
        'Dim dbConnection As SqlConnection = New SqlConnection(sConnStr)

        'dbConnection.Open()
        'Dim ts As SqlTransaction
        'ts = dbConnection.BeginTransaction
        'Dim buss As New BusinessRules.WorkFlow(dbConnection, ts)
        'Try
        '    Dim test As String = buss.CreateProcess("01", tmpProjectID, tmpCreatePerson)

        '    ts.Commit()
        'Catch
        '    ts.Rollback()
        '    MsgBox(Err.Description())
        'End Try
        'Dim sConnStr As String = "data source=localhost;initial catalog=cgmis;persist security info=False;user id=ioryfeng;password=ioryfeng;workstation id=localhost;packet size=4096"
        'Dim sConnStr2 As String = "dsn=OldData;uid=123;pwd=123"
        ''Dim dbConnection As OdbcConnection = New OdbcConnection(sConnStr2)
        'Dim dbConnection2 As SqlConnection = New SqlConnection(sConnStr)
        'Dim dsImportCorporation As New DataSet()
        'Dim dsImportConsultation As New DataSet()
        'Dim dsImportProject As New DataSet()
        'Dim dsTemp As New DataSet()
        'Dim ts2 As SqlTransaction
        'Dim newRow As DataRow
        'Dim i As Integer

        'Try
        '    'dbConnection.Open()
        '    dbConnection2.Open()
        '    ts2 = dbConnection2.BeginTransaction
        '    Dim buss As New corporationAccess(dbConnection2, ts2)

        '    '获取新系统企业表和咨询表结构
        '    dsTemp = buss.GetcorporationInfo("null", "null")

        ''获取旧系统企业信息
        'Dim daCorporation As OdbcDataAdapter = New OdbcDataAdapter("select * from ImportCorporation", dbConnection)
        'daCorporation.Fill(dsImportCorporation)

        ''将企业信息导入到企业表中
        'For i = 0 To dsImportCorporation.Tables(0).Rows.Count - 1

        '    newRow = dsTemp.Tables(0).NewRow
        '    With newRow
        '        .Item("corporation_code") = dsImportCorporation.Tables(0).Rows(i).Item("corporation_code")
        '        .Item("corporation_name") = dsImportCorporation.Tables(0).Rows(i).Item("corporation_name")
        '        .Item("web_site") = dsImportCorporation.Tables(0).Rows(i).Item("web_site")
        '        .Item("district_name") = dsImportCorporation.Tables(0).Rows(i).Item("district_name")
        '        .Item("industry_type") = dsImportCorporation.Tables(0).Rows(i).Item("industry_type")
        '        .Item("proprietorship_type") = dsImportCorporation.Tables(0).Rows(i).Item("proprietorship_type")
        '        If Trim(IIf(IsDBNull(dsImportCorporation.Tables(0).Rows(i).Item("technology_type")), "", dsImportCorporation.Tables(0).Rows(i).Item("technology_type"))) = "" Then
        '            .Item("technology_type") = DBNull.Value
        '        Else
        '            .Item("technology_type") = transTechType(Trim(dsImportCorporation.Tables(0).Rows(i).Item("technology_type")))
        '        End If
        '        .Item("contact_person") = dsImportCorporation.Tables(0).Rows(i).Item("contact_person")
        '        .Item("phone_num") = dsImportCorporation.Tables(0).Rows(i).Item("phone_num")
        '        .Item("fax") = dsImportCorporation.Tables(0).Rows(i).Item("fax")
        '    End With

        '    dsTemp.Tables(0).Rows.Add(newRow)

        'Next
        'buss.UpdateCorporation(dsTemp)



        ''获取旧系统咨询信息
        'Dim daConsultation As OdbcDataAdapter = New OdbcDataAdapter("select * from ImportConsultation", dbConnection)
        'daConsultation.Fill(dsImportConsultation)

        ''将咨询信息导入到企业表中
        'For i = 0 To dsImportConsultation.Tables(0).Rows.Count - 1

        '    newRow = dsTemp.Tables(1).NewRow
        '    With newRow
        '        .Item("corporation_code") = dsImportConsultation.Tables(0).Rows(i).Item("corporation_code")
        '        .Item("corporation_name") = dsImportConsultation.Tables(0).Rows(i).Item("corporation_name")
        '        .Item("web_site") = dsImportConsultation.Tables(0).Rows(i).Item("web_site")
        '        .Item("serial_num") = transNum(dsImportConsultation.Tables(0).Rows(i).Item("serial_num"))
        '        .Item("district_name") = dsImportConsultation.Tables(0).Rows(i).Item("district_name")
        '        .Item("industry_type") = dsImportConsultation.Tables(0).Rows(i).Item("industry_type")
        '        .Item("proprietorship_type") = dsImportConsultation.Tables(0).Rows(i).Item("proprietorship_type")
        '        If Trim(IIf(IsDBNull(dsImportConsultation.Tables(0).Rows(i).Item("technology_type")), "", dsImportConsultation.Tables(0).Rows(i).Item("technology_type"))) = "" Then
        '            .Item("technology_type") = DBNull.Value
        '        Else
        '            .Item("technology_type") = transTechType(Trim(dsImportConsultation.Tables(0).Rows(i).Item("technology_type")))
        '        End If

        '        If Trim(IIf(IsDBNull(dsImportConsultation.Tables(0).Rows(i).Item("is_first_loan")), "", dsImportConsultation.Tables(0).Rows(i).Item("is_first_loan"))) = "" Then
        '            .Item("is_first_loan") = DBNull.Value
        '        Else
        '            .Item("is_first_loan") = transFirstLoan(Trim(dsImportConsultation.Tables(0).Rows(i).Item("is_first_loan")))
        '        End If
        '        .Item("apply_guarantee_sum") = dsImportConsultation.Tables(0).Rows(i).Item("apply_guarantee_sum")
        '        .Item("contact_person") = dsImportConsultation.Tables(0).Rows(i).Item("contact_person")
        '        .Item("phone_num") = dsImportConsultation.Tables(0).Rows(i).Item("phone_num")
        '        .Item("fax") = dsImportConsultation.Tables(0).Rows(i).Item("fax")
        '        .Item("recommend_person") = dsImportConsultation.Tables(0).Rows(i).Item("recommend_person")
        '        .Item("consult_person") = dsImportConsultation.Tables(0).Rows(i).Item("consult_person")
        '        .Item("consult_date") = dsImportConsultation.Tables(0).Rows(i).Item("consult_date")

        '    End With

        '    dsTemp.Tables(1).Rows.Add(newRow)

        'Next
        'buss.UpdateConsultation(dsTemp)




        '    '获取旧系统项目信息
        '    Dim daProject As OdbcDataAdapter = New OdbcDataAdapter("select * from ImportProject", dbConnection)
        '    daProject.Fill(dsImportProject)

        '    '获取项目表结构
        '    Dim bussProject As New Project(dbConnection2, ts2)
        '    Dim dsProject As DataSet
        '    dsProject = bussProject.GetProjectInfo("null")

        '    '将项目信息导入到项目表中
        '    For i = 0 To dsImportProject.Tables(0).Rows.Count - 1
        '        newRow = dsProject.Tables(0).NewRow
        '        With newRow
        '            .Item("corporation_code") = dsImportProject.Tables(0).Rows(i).Item("corporation_code")
        '            .Item("project_code") = dsImportProject.Tables(0).Rows(i).Item("project_code")
        '            If Trim(IIf(IsDBNull(dsImportProject.Tables(0).Rows(i).Item("is_first_loan")), "", dsImportProject.Tables(0).Rows(i).Item("is_first_loan"))) = "" Then
        '                .Item("is_first_loan") = DBNull.Value
        '            Else
        '                .Item("is_first_loan") = transFirstLoan(Trim(dsImportProject.Tables(0).Rows(i).Item("is_first_loan")))
        '            End If
        '            .Item("apply_sum") = dsImportProject.Tables(0).Rows(i).Item("apply_sum")
        '        End With

        '        dsProject.Tables(0).Rows.Add(newRow)

        '    Next

        '    bussProject.UpdateProject(dsProject)

        '    ts2.Commit()
        'Catch
        '    ts2.Rollback()
        '    MsgBox(Err.Description)
        'End Try


        DataGrid1.DataSource = dsTemp

        'DataGrid2.DataSource = dsTemp

    End Sub

    '转换大写数字
    Private Function transNum(ByVal CNum As String) As Integer
        Select Case CNum
            Case "一"
                Return 1
            Case "二"
                Return 2
            Case "三"
                Return 3
            Case "四"
                Return 4
            Case "五"
                Return 5
            Case "六"
                Return 6
            Case "七"
                Return 7
            Case "八"
                Return 8
            Case "九"
                Return 9
            Case "十"
                Return 10
        End Select
    End Function

    '转换科技类型
    Private Function transTechType(ByVal techType As String) As Integer
        Select Case techType
            Case "高科技企业"
                Return 1
            Case "高科技项目"
                Return 2
            Case "民科企业"
                Return 4
            Case "其他"
                Return 32

        End Select

    End Function

    '转换是否首次贷款
    Private Function transFirstLoan(ByVal firstLoan As String) As Integer
        Select Case firstLoan
            Case "是"
                Return 0
            Case "否"
                Return 1
        End Select
    End Function


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim tmpProjectID As String = Trim(txtProjectID.Text)
        Dim tmpCreatePerson As String = Trim(txtCreatePerson.Text)

        Dim tmpFinishedTaskID As String = Trim(txtFinishedTaskID.Text)
        Dim tmpFinishedFlag As String = Trim(txtFinishedFlag.Text)
        Dim tmpUserID As String = Trim(txtUserID.Text)
        Dim dbConnection As SqlConnection = New SqlConnection(sConnStr)

        dbConnection.Open()
        Dim ts As SqlTransaction
        ts = dbConnection.BeginTransaction
        Dim buss As New BusinessRules.WorkFlow(dbConnection, ts)
        Dim buss1 As New BusinessRules.CommonQuery(dbConnection, ts)
        Dim buss2 As New BusinessRules.CommonQuery(dbConnection, ts)
        Dim workflow As WorkFlow = New WorkFlow(dbConnection, ts)

        'Dim Item As New BusinessRules.Item(dbConnection, ts)

        Dim gws As New localhost.Service1()

        Try
            'Dim strTest As String = buss.CreateProcess(tmpCreatePerson, tmpProjectID, tmpUserID)
            ' Dim strTest As String = buss.rollbackTask(tmpCreatePerson, tmpProjectID, tmpFinishedTaskID, tmpUserID, tmpFinishedFlag)

            'Dim ds As DataSet = Item.GetItemEx("%")
            ''Dim s As String = buss2.GetMaxContractNum("051721006")
            ''buss1.GetMaxContractNum("000011101")
            'Dim a As String = "000011101"
            'Dim b As String = a.Substring(0, 5)

            Dim sss As String = Now.ToString("yyyy-MM-dd")


            'Dim strTest As String = buss.finishedTask("01", "000021301", "LetterGuaranteeApplication", "", "韩璐")
            'Dim strTest As String = workflow.updateProcess() '
            'Dim strTest As String = buss.CancelReviewConferencePlan(tmpProjectID)
            'Dim strTest As String = buss.CancelReviewConferencePlanProject(tmpProjectID)
            'Dim strTest As String = buss.CancelReviewConferencePlan("24")
            'dsTemp = buss1.GetProjectInfoEx("{ProjectCode='000330302'}")

            'Dim tmpManagerA As String = dsTemp.Tables(0).Rows(0).Item("24")
            'Dim tmpManagerB As String = dsTemp.Tables(0).Rows(0).Item("25")

            'Dim strTest As String = buss.FinishedSignaturePlan(tmpProjectID)



           

            ts.Commit()
        Catch ex As BusinessRules.WorkFlowErr
            ts.Rollback()
            MsgBox(ex.ErrMessage)
        End Try


        'Dim dbConnection As SqlConnection = New SqlConnection(sConnStr)
        'dbConnection.Open()
        'Dim ts As SqlTransaction
        'ts = dbConnection.BeginTransaction
        'Dim buss As New BusinessRules.Project(dbConnection, ts)
        'Try
        '    'Dim strErr2 As String = webSrv.UpdateBranch(dsTemp.GetChanges)
        '    ts.Commit()
        '    'dsTemp.AcceptChanges()
        '    'Catch ex As DBConcurrencyException
        '    '    MsgBox(ex.Message)
        '    '    ts.Rollback()
        'Catch oErr As Exception
        '    MsgBox(oErr.Message)
        '    ts.Rollback()
        'End Try


        DataGrid2.DataSource = dsTemp
    End Sub

    'Private Sub frmcorporation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    'Dim a As Boolean = True
    '    'If a = False Then
    '    '    Button1.Enabled = False
    '    'End If
    '    task_Timer.Interval = 60000
    '    task_Timer.Enabled = True
    '    task_Timer.Start()
    'End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
       

    End Sub

    'Private Sub task_Timer_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles task_Timer.Elapsed
    '    Dim sysTime As String = FormatDateTime(Now, DateFormat.ShortTime)
    '    If sysTime = "09:00" Then
    '        Dim dbConnection As SqlConnection = New SqlConnection(sConnStr)

    '        dbConnection.Open()
    '        Dim ts As SqlTransaction
    '        ts = dbConnection.BeginTransaction
    '        Dim tmpTimingServer As New TimingServer(dbConnection, ts)
    '        Try
    '            tmpTimingServer.TimingServer()

    '            ts.Commit()
    '        Catch ex As BusinessRules.WorkFlowErr
    '            ts.Rollback()
    '            MsgBox(ex.ErrMessage)
    '        End Try

    '    End If
    'End Sub

    '获取配置文件的配置信息
    Private Function GetConfigString(ByVal objNodeList As XmlNodeList, ByVal strKey As String) As String
        Dim objXmlNode As XmlNode
        Dim objXmlAttr As XmlAttribute
        Dim i As Integer = objNodeList.Count
        For Each objXmlNode In objNodeList
            Dim strn As String = objXmlNode.Name
            If objXmlNode.NodeType <> XmlNodeType.XmlDeclaration Then

                If objXmlNode.HasChildNodes Then
                    GetConfigString = GetConfigString(objXmlNode.ChildNodes, strKey)
                Else

                    If objXmlNode.Name = "add" Then
                        Dim k As Integer = objXmlNode.Attributes.Count
                        If objXmlNode.Attributes("key").Value = strKey Then

                            GetConfigString = objXmlNode.Attributes("value").Value

                        End If
                    End If
                End If
            End If
        Next

    End Function

    'Private Sub DataGrid1_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid1.CurrentCellChanged
    '    Static lastRow As Integer = -1
    '    If DataGrid1.CurrentRowIndex <> lastRow Then

    '        '获取ID
    '        Dim bankCode As String = DataGrid1.Item(DataGrid1.CurrentRowIndex, 0)
    '        strSql = "bank_code=" & "'" & bankCode & "'"
    '        dsTemp.Tables(1).DefaultView.RowFilter = strSql
    '        DataGrid2.DataSource = dsTemp.Tables(1).DefaultView
    '    End If
    '    lastRow = DataGrid1.CurrentRowIndex
    'End Sub

    'Private Sub task_Timer_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles task_Timer.Elapsed
    '    MsgBox("go")
    '    MsgBox("stop")
    'End Sub

    Private Sub frmTest_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim dbConnection As SqlConnection = New SqlConnection(sConnStr)

        'dbConnection.Open()
        'Dim ts As SqlTransaction
        'ts = dbConnection.BeginTransaction

        'Dim ds As New DataSet

        'ds.ReadXmlSchema(Application.StartupPath() + "\yewu_date.xsd")

        'Dim dt As DataTable = ds.Tables(0)

        ''Dim objCommonQuery As New CommonQuery(dbConnection, ts)

        ''Dim dt1 As DataTable = objCommonQuery.GetCommonQueryInfo("select * from bank").Tables(0).Copy()

        ''dt1.TableName = "dt1"

        ''Dim dt2 As DataTable = objCommonQuery.GetCommonQueryInfo("select * from branch").Tables(0).Copy()

        ''dt2.TableName = "dt2"


        ''ds.Tables.Add(dt1)

        ''ds.Tables.Add(dt2)

        'ds.WriteXml(Application.StartupPath() + "\XMLDATA.xml")


    End Sub
End Class
