Public Class WorkFlowErr
    Inherits System.Exception

    Private Err As String
    Public Property ErrMessage()
        Get
            Return Err
        End Get
        Set(ByVal Value)

        End Set
    End Property

    '�׳������������ڴ���
    Public Function ThrowNotExistWorkFlowErr()
        Err = "������������!"
    End Function

    '�׳��������Ѵ��ڴ���
    Public Function ThrowExistWorkFlowErr()
        Err = "�������Ѿ�����!"
    End Function

    '�׳����������Ѵ��ڴ���
    Public Function ThrowNotExistTaskErr()
        Err = "���񲻴���!"
    End Function

    '�׳������ύ��ͣ���������
    Public Function ThrowWaitingTaskErr()
        Err = "�����ύ����ͣ������!"
    End Function

    '�׳������ύ��ֹ���������
    Public Function ThrowTerminateTaskErr()
        Err = "�����ύ����ֹ������!"
    End Function

    '�׳��ύ����Ľ����Ч����
    Public Function ThrowInvalidSubmit()
        Err = "�ύ���������ҵ�����!"
    End Function

    '�׳��ᱣǰ���м�¼����δ��ɴ���
    Public Function ThrowPreguaranteeActivityErr()
        Err = "��¼��ǰ���м�¼����δ���!"
    End Function

    '�׳�δ������ί����
    Public Function ThrowAddRegentErr()
        Err = "δ������ί!"
    End Function

    '�׳�δ�����ʲ�����ʦ����
    Public Function ThrowAddValuatorErr()
        Err = "δ�����ʲ�����ʦ!"
    End Function

    '�׳�δ�ҵ���Ŀ�鳤����
    Public Function ThrowNoTeamHeaderErr()
        Err = "δ�ҵ���Ŀ�鳤!"
    End Function

    '�׳�δ�ҵ���Ŀ���Ա
    Public Function ThrowNoTeamStaffErr()
        Err = "δ�������Ŀ���Ա!"
    End Function

    '�׳�û��ί�е�Ȩ��
    Public Function ThrowNoConsignRight()
        Err = "û��ί��Ȩ��!"
    End Function

    '������ί��������
    Public Function ThrowIsConsign()
        Err = "��Ա����ί��������!"
    End Function

    '��������
    Public Function ThrowNoConsigner()
        Err = "��������!"
    End Function

    '�׳�"���ܳ����ѿ����������"
    Public Function ThrowRecordReviewConclusionErr()
        Err = "���ܳ����ѿ����������!"
    End Function

    '�׳�"���ܳ�����ǩԼ�ļƻ�"
    Public Function ThrowRecordSignatureErr()
        Err = "���ܳ�����ǩԼ�ļƻ�!"
    End Function

    '�׳�"���ܻ����ֹ�����������"
    Public Function ThrowRollBackManualTaskErr()
        Err = "���ܻ����ֹ�����������!"
    End Function

    Public Function ThrowNoStaffRole()
        Err = "û�ж����Ӧ��Ա����ɫ!"
    End Function

    Public Function ThrowMustRecord()
        Err = "�����¼�������ۺ�ſ������ϻ�!"
    End Function

    Public Function ThrowMustValidateLoan()
        Err = "��Ӧ�Ĵ������Ŀ��ǩ���ſ�֪ͨ��!"
    End Function

    Public Function ThrowMustValidateLoanSmall()
        Err = "��Ӧ��С����Ŀ��ǩ���ſ�֪ͨ��!"
    End Function

    Public Function ThrowMustRecordReturnReceipt()
        Err = "��Ӧ��С����Ŀ��ǼǷſ��ִ!"
    End Function

    Public Function ThrowMustDraftOutContract()
        Err = "��Ӧ�Ĵ������Ŀ��������ͬ!"
    End Function

    Public Function ThrowNoSmallCredit()
        Err = "����ЧС�����Ŷ�Ȼ��Ȳ���!"
    End Function

    Public Function ThrowRelateProjectRefund()
        Err = "������ǰС����Ŀ�軹���ſɷſ�!"
    End Function

    Public Function ThrowMustGuaranteeFee()
        Err = "��Ӧ�Ĵ������Ŀ����ȡ������!"
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''�쳣����'''''''''''''''''''''''

    Public Function ThrowNoRecordkErr(ByVal dtErr As DataTable)
        Err = "δ�ҵ���Ӧ��¼!" & vbCrLf & "Դ��:" & dtErr.TableName
    End Function

    Public Function ThrowMustRecordSignatureSubmitor()
        Err = "�Ǽ�ǩԼ�ύ�˱����Ƿ��ղ�������!"
    End Function

End Class
