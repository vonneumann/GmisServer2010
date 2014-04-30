Imports System.Data.SqlClient
Imports BusinessRules

Public Class frmConsole
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents task_Timer As System.Timers.Timer
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSrv As System.Windows.Forms.TextBox
    Friend WithEvents txtUsr As System.Windows.Forms.TextBox
    Friend WithEvents txtPwd As System.Windows.Forms.TextBox
    Friend WithEvents txtDb As System.Windows.Forms.TextBox
    Friend WithEvents lblStart As System.Windows.Forms.Label
    Friend WithEvents lblStop As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmConsole))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSrv = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtUsr = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtPwd = New System.Windows.Forms.TextBox()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.lblStart = New System.Windows.Forms.Label()
        Me.lblStop = New System.Windows.Forms.Label()
        Me.task_Timer = New System.Timers.Timer()
        Me.txtDb = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        CType(Me.task_Timer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "服务器:"
        '
        'txtSrv
        '
        Me.txtSrv.Location = New System.Drawing.Point(91, 13)
        Me.txtSrv.Name = "txtSrv"
        Me.txtSrv.Size = New System.Drawing.Size(125, 21)
        Me.txtSrv.TabIndex = 1
        Me.txtSrv.Text = "localhost"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "用户名:"
        '
        'txtUsr
        '
        Me.txtUsr.Location = New System.Drawing.Point(91, 73)
        Me.txtUsr.Name = "txtUsr"
        Me.txtUsr.Size = New System.Drawing.Size(125, 21)
        Me.txtUsr.TabIndex = 3
        Me.txtUsr.Text = "sa"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 107)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 23)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "口令:"
        '
        'txtPwd
        '
        Me.txtPwd.Location = New System.Drawing.Point(91, 103)
        Me.txtPwd.Name = "txtPwd"
        Me.txtPwd.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtPwd.Size = New System.Drawing.Size(125, 21)
        Me.txtPwd.TabIndex = 4
        Me.txtPwd.Text = "123"
        '
        'btnStart
        '
        Me.btnStart.Image = CType(resources.GetObject("btnStart.Image"), System.Drawing.Bitmap)
        Me.btnStart.Location = New System.Drawing.Point(48, 144)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(32, 23)
        Me.btnStart.TabIndex = 5
        '
        'btnStop
        '
        Me.btnStop.Image = CType(resources.GetObject("btnStop.Image"), System.Drawing.Bitmap)
        Me.btnStop.Location = New System.Drawing.Point(48, 184)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(32, 23)
        Me.btnStop.TabIndex = 6
        '
        'lblStart
        '
        Me.lblStart.Location = New System.Drawing.Point(96, 152)
        Me.lblStart.Name = "lblStart"
        Me.lblStart.TabIndex = 7
        Me.lblStart.Text = "开始/继续"
        '
        'lblStop
        '
        Me.lblStop.Location = New System.Drawing.Point(96, 192)
        Me.lblStop.Name = "lblStop"
        Me.lblStop.Size = New System.Drawing.Size(56, 23)
        Me.lblStop.TabIndex = 8
        Me.lblStop.Text = "停止"
        '
        'task_Timer
        '
        Me.task_Timer.Interval = 60000
        Me.task_Timer.SynchronizingObject = Me
        '
        'txtDb
        '
        Me.txtDb.Location = New System.Drawing.Point(91, 43)
        Me.txtDb.Name = "txtDb"
        Me.txtDb.Size = New System.Drawing.Size(125, 21)
        Me.txtDb.TabIndex = 2
        Me.txtDb.Text = "cgmis"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 47)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 23)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "数据库:"
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "定时服务"
        Me.NotifyIcon1.Visible = True
        '
        'frmConsole
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(232, 229)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtDb, Me.Label6, Me.lblStop, Me.lblStart, Me.btnStop, Me.btnStart, Me.txtPwd, Me.Label3, Me.txtUsr, Me.Label2, Me.txtSrv, Me.Label1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmConsole"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "定时服务管理器"
        CType(Me.task_Timer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private strConn As String
    Private Ddone(30) As Boolean  '表示还未扫描当天的定时服务
    Private Hdone(23) As Boolean  '表示还未扫描本时定时服务

    Private Sub frmConsole_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''获取数据库连接设置
        'Dim strSrv As String = Trim(txtSrv.Text)
        'Dim strDb As String = Trim(txtDb.Text)
        'Dim strUsr As String = Trim(txtUsr.Text)
        'Dim strPwd As String = Trim(txtPwd.Text)

        'strConn = "UID=" & strUsr & ";PWD=" & strPwd & ";Initial Catalog=" & strDb & ";Data Source=" & strSrv

        ''设置定时器时间间隔，并启动定时器
        'task_Timer.Interval = 60000
        'task_Timer.Enabled = True
        'task_Timer.Start()

        '屏蔽停止按钮
        btnStop.Enabled = False
        lblStop.Enabled = False

    End Sub



    Private Sub task_Timer_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles task_Timer.Elapsed


        Dim iDay As Integer = DateTime.Today.Day
        Dim iHour As Integer = DateTime.Now.Hour

        '如果当天未扫描过，则扫描
        If Ddone(iDay - 1) = False Then
            ScanTimingTask(iDay, iHour)
        End If

        '如果当前小时未扫描过，则扫描
        If Hdone(iHour - 1) = False Then
            ScanTimingTask(iDay, iHour)
        End If

    End Sub

    Private Function ScanTimingTask(ByVal iDay As Integer, ByVal iHour As Integer)
        '判断
        Dim i As Integer
        'Dim sysTime As String = FormatDateTime(Now, DateFormat.ShortTime)
        Dim dbConnection As SqlConnection = New SqlConnection(strConn)
        Dim ts As SqlTransaction
        Try
            dbConnection.Open()
            ts = dbConnection.BeginTransaction
            Try

                Dim tmpTimingServer As New TimingServer(dbConnection, ts, Ddone(iDay - 1), Hdone(iHour - 1))
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

                ts.Commit()
            Catch
                ts.Rollback()
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "定时服务")
            Throw ex
        End Try

        dbConnection.Close()
        dbConnection.Dispose()

    End Function


    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Me.Cursor = Cursors.WaitCursor

        Try

            '获取数据库连接设置
            Dim strSrv As String = Trim(txtSrv.Text)
            Dim strDb As String = Trim(txtDb.Text)
            Dim strUsr As String = Trim(txtUsr.Text)
            Dim strPwd As String = Trim(txtPwd.Text)

            strConn = "UID=" & strUsr & ";PWD=" & strPwd & ";Initial Catalog=" & strDb & ";Data Source=" & strSrv

            '启动时先扫描一次
            task_Timer_Elapsed(Nothing, Nothing)


            '设置定时器时间间隔，并启动定时器
            task_Timer.Interval = 60000
            task_Timer.Enabled = True
            task_Timer.Start()

            '屏蔽开始按钮
            btnStart.Enabled = False
            lblStart.Enabled = False

            '打开停止按钮
            btnStop.Enabled = True
            lblStop.Enabled = True
        Catch
            Me.Cursor = Cursors.Default
        End Try

        Me.Cursor = Cursors.Default
    End Sub


    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        '停止定时器
        task_Timer.Enabled = False
        task_Timer.Stop()

        '屏蔽停止按钮
        btnStop.Enabled = False
        lblStop.Enabled = False

        '打开开始按钮
        btnStart.Enabled = True
        lblStart.Enabled = True

    End Sub


    Private Sub NotifyIcon1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
    End Sub


    Private Sub frmConsole_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Dim response As MsgBoxResult = MsgBox("是否关闭定时服务?", MsgBoxStyle.YesNo, "定时服务")
        If response = MsgBoxResult.Yes Then
            NotifyIcon1.Dispose()
            Me.Dispose()
        Else
            e.Cancel = True
        End If


    End Sub


    Private Sub frmConsole_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
        End If
    End Sub


End Class
