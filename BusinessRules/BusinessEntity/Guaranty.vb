Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Guaranty
    Public Const Table_Opposite_Guarantee As String = "opposite_guarantee"
    Public Const Table_Opposite_Guarantee_Detail As String = "opposite_guarantee_detail"


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_Opposite_Guarantee As SqlDataAdapter
    Private dsCommand_Opposite_Guarantee_Detail As SqlDataAdapter

    '�����ѯ����
    Private GetOppositeGuaranteeInfoCommand As SqlCommand
    Private GetOppositeGuaranteeDetailInfoCommand As SqlCommand
    Private GetMaxGuarantyNumCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_Opposite_Guarantee = New SqlDataAdapter()
        dsCommand_Opposite_Guarantee_Detail = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetGuarantyInfo("null", "null")
    End Sub

    '��ȡ����������Ϣ
    Public Function GetGuarantyInfo(ByVal strSQL_Condition_OppositeGuarantee As String, ByVal strSQL_Condition_OppositeGuaranteeDetail As String) As DataSet

        Dim tempDs As New DataSet()

        If GetOppositeGuaranteeInfoCommand Is Nothing Then

            GetOppositeGuaranteeInfoCommand = New SqlCommand("GetOppositeGuaranteeInfo", conn)
            GetOppositeGuaranteeInfoCommand.CommandType = CommandType.StoredProcedure
            GetOppositeGuaranteeInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Opposite_Guarantee
            .SelectCommand = GetOppositeGuaranteeInfoCommand
            .SelectCommand.Transaction = ts
            GetOppositeGuaranteeInfoCommand.Parameters("@Condition").Value = strSQL_Condition_OppositeGuarantee
            .Fill(tempDs, Table_Opposite_Guarantee)
        End With


        If GetOppositeGuaranteeDetailInfoCommand Is Nothing Then

            GetOppositeGuaranteeDetailInfoCommand = New SqlCommand("GetOppositeGuaranteeDetailInfo", conn)
            GetOppositeGuaranteeDetailInfoCommand.CommandType = CommandType.StoredProcedure
            GetOppositeGuaranteeDetailInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Opposite_Guarantee_Detail
            .SelectCommand = GetOppositeGuaranteeDetailInfoCommand
            .SelectCommand.Transaction = ts
            GetOppositeGuaranteeDetailInfoCommand.Parameters("@Condition").Value = strSQL_Condition_OppositeGuaranteeDetail
            .Fill(tempDs, Table_Opposite_Guarantee_Detail)
        End With

        GetGuarantyInfo = tempDs

    End Function

    '��ȡ������к�
    Public Function GetMaxGuarantyNum(ByVal projectID As String) As Integer

        If GetMaxGuarantyNumCommand Is Nothing Then

            GetMaxGuarantyNumCommand = New SqlCommand("GetMaxGuarantyNum", conn)
            GetMaxGuarantyNumCommand.CommandType = CommandType.StoredProcedure
            GetMaxGuarantyNumCommand.Parameters.Add(New SqlParameter("@projectID", SqlDbType.NVarChar))
            GetMaxGuarantyNumCommand.Parameters.Add(New SqlParameter("@maxGuarantyNum", SqlDbType.Int))
            GetMaxGuarantyNumCommand.Parameters.Item("@maxGuarantyNum").Direction = ParameterDirection.Output
            GetMaxGuarantyNumCommand.Transaction = ts
        End If

        GetMaxGuarantyNumCommand.Parameters("@projectID").Value = projectID
        GetMaxGuarantyNumCommand.ExecuteNonQuery()
        GetMaxGuarantyNum = GetMaxGuarantyNumCommand.Parameters.Item("@maxGuarantyNum").Value
    End Function

    '���·���������Ϣ
    Private Function UpdatedsOppositeGuarantee(ByVal OppositeGuaranteeSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Opposite_Guarantee)

        With dsCommand_Opposite_Guarantee
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(OppositeGuaranteeSet, Table_Opposite_Guarantee)

        End With

    End Function

    '���·���������ϸ��Ϣ
    Private Function UpdatedsOppositeGuaranteeDetail(ByVal OppositeGuaranteeDetailSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Opposite_Guarantee_Detail)

        With dsCommand_Opposite_Guarantee_Detail
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(OppositeGuaranteeDetailSet, Table_Opposite_Guarantee_Detail)

        End With


    End Function

    '���·�������,��ϸ��Ϣ
    Public Function UpdateGuaranty(ByVal GuarantySet As DataSet)

        If GuarantySet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If GuarantySet.HasChanges = False Then
            Exit Function
        End If

        'ɾ������
        If IsNothing(GuarantySet.GetChanges(DataRowState.Deleted)) = False Then
            UpdatedsOppositeGuaranteeDetail(GuarantySet.GetChanges(DataRowState.Deleted))
            UpdatedsOppositeGuarantee(GuarantySet.GetChanges(DataRowState.Deleted))
        End If

        '��������
        If IsNothing(GuarantySet.GetChanges(DataRowState.Added)) = False Then

            ''ͬʱ�����������ϸ��
            'If GuarantySet.GetChanges(DataRowState.Added).Tables(0).Rows.Count <> 0 And GuarantySet.GetChanges(DataRowState.Added).Tables(1).Rows.Count <> 0 Then

            '    '��������
            '    Dim i, j As Integer
            '    Dim tmpRowPrimary, tmpRowDetail As DataRow
            '    For i = 0 To GuarantySet.Tables(0).Rows.Count - 1
            '        tmpRowPrimary = GuarantySet.Tables(0).Rows(i)
            '        If tmpRowPrimary.RowState = DataRowState.Added Then

            '            '�����ͻ��˵���Ŀ����
            '            Dim projectCode As String = Trim(GuarantySet.Tables(0).Rows(i).Item("project_code"))
            '            Dim strSql As String = "{project_code=" & "'" & projectCode & "'" & " order by serial_num}"

            '            '��ȡ���ݿ��и���Ŀ�ķ��������¼��Ϊ��ȡ������������룩
            '            Dim dsTemp As DataSet = GetGuarantyInfo(strSql, "null")
            '            Dim rowNum As Integer = dsTemp.Tables(0).Rows.Count

            '            '�������ݿ��������������
            '            Dim serialNum As Integer
            '            If rowNum = 0 Then
            '                serialNum = 1
            '            Else
            '                serialNum = dsTemp.Tables(0).Rows(rowNum - 1).Item("serial_num") + 1
            '            End If

            '            '�����ͻ��������������
            '            Dim serialNumTemp As Integer = GuarantySet.Tables(0).Rows(i).Item("serial_num")

            '            '�ѿͻ��������ڸ��������������������������Ϊ���ݿ��е������������
            '            GuarantySet.Tables(0).Rows(i).Item("serial_num") = serialNum

            '            '�ѿͻ��������ڸ���������������ķ���������ϸ������������Ϊ���ݿ��е������������
            '            For j = 0 To GuarantySet.GetChanges(DataRowState.Added).Tables(1).Rows.Count - 1
            '                tmpRowDetail = GuarantySet.Tables(1).Rows(j)
            '                If tmpRowDetail.RowState = DataRowState.Added Then
            '                    If GuarantySet.Tables(1).Rows(j).Item("serial_num") = serialNumTemp Then
            '                        GuarantySet.Tables(1).Rows(j).Item("serial_num") = serialNum
            '                    End If
            '                End If
            '            Next
            '        End If
            '    Next
            'End If

            UpdatedsOppositeGuarantee(GuarantySet.GetChanges(DataRowState.Added))
            UpdatedsOppositeGuaranteeDetail(GuarantySet.GetChanges(DataRowState.Added))

        End If

        '�޸Ĳ���
        If IsNothing(GuarantySet.GetChanges(DataRowState.Modified)) = False Then

            '����ǵ���������Ϣ��ֱ�ӵ����·���
            UpdatedsOppositeGuarantee(GuarantySet.GetChanges(DataRowState.Modified))
            UpdatedsOppositeGuaranteeDetail(GuarantySet.GetChanges(DataRowState.Modified))

        End If

        GuarantySet.AcceptChanges()

    End Function

End Class
