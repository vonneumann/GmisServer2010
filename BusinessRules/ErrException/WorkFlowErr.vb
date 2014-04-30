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

    '抛出工作流不存在错误
    Public Function ThrowNotExistWorkFlowErr()
        Err = "工作流不存在!"
    End Function

    '抛出工作流已存在错误
    Public Function ThrowExistWorkFlowErr()
        Err = "工作流已经存在!"
    End Function

    '抛出工作任务已存在错误
    Public Function ThrowNotExistTaskErr()
        Err = "任务不存在!"
    End Function

    '抛出不能提交暂停的任务错误
    Public Function ThrowWaitingTaskErr()
        Err = "不能提交已暂停的任务!"
    End Function

    '抛出不能提交终止的任务错误
    Public Function ThrowTerminateTaskErr()
        Err = "不能提交已终止的任务!"
    End Function

    '抛出提交任务的结果无效错误
    Public Function ThrowInvalidSubmit()
        Err = "提交结果不满足业务规则!"
    End Function

    '抛出提保前调研记录任务未完成错误
    Public Function ThrowPreguaranteeActivityErr()
        Err = "记录保前调研记录任务未完成!"
    End Function

    '抛出未分配评委错误
    Public Function ThrowAddRegentErr()
        Err = "未分配评委!"
    End Function

    '抛出未分配资产评估师错误
    Public Function ThrowAddValuatorErr()
        Err = "未分配资产评估师!"
    End Function

    '抛出未找到项目组长错误
    Public Function ThrowNoTeamHeaderErr()
        Err = "未找到项目组长!"
    End Function

    '抛出未找到项目组成员
    Public Function ThrowNoTeamStaffErr()
        Err = "未定义该项目组成员!"
    End Function

    '抛出没有委托的权限
    Public Function ThrowNoConsignRight()
        Err = "没有委托权限!"
    End Function

    '该人已委托了任务
    Public Function ThrowIsConsign()
        Err = "该员工已委托了任务!"
    End Function

    '无受托人
    Public Function ThrowNoConsigner()
        Err = "无受托人!"
    End Function

    '抛出"不能撤销已开过的评审会"
    Public Function ThrowRecordReviewConclusionErr()
        Err = "不能撤销已开过的评审会!"
    End Function

    '抛出"不能撤销已签约的计划"
    Public Function ThrowRecordSignatureErr()
        Err = "不能撤销已签约的计划!"
    End Function

    '抛出"不能回退手工启动的任务"
    Public Function ThrowRollBackManualTaskErr()
        Err = "不能回退手工启动的任务!"
    End Function

    Public Function ThrowNoStaffRole()
        Err = "没有定义对应的员工角色!"
    End Function

    Public Function ThrowMustRecord()
        Err = "必须记录评审会结论后才可重新上会!"
    End Function

    Public Function ThrowMustValidateLoan()
        Err = "对应的贷款担保项目需签发放款通知书!"
    End Function

    Public Function ThrowMustValidateLoanSmall()
        Err = "对应的小贷项目需签发放款通知书!"
    End Function

    Public Function ThrowMustRecordReturnReceipt()
        Err = "对应的小贷项目需登记放款回执!"
    End Function

    Public Function ThrowMustDraftOutContract()
        Err = "对应的贷款担保项目需制作合同!"
    End Function

    Public Function ThrowNoSmallCredit()
        Err = "无有效小贷授信额度或额度不足!"
    End Function

    Public Function ThrowRelateProjectRefund()
        Err = "关联贷前小贷项目需还款后才可放款!"
    End Function

    Public Function ThrowMustGuaranteeFee()
        Err = "对应的贷款担保项目需收取担保费!"
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''异常处理'''''''''''''''''''''''

    Public Function ThrowNoRecordkErr(ByVal dtErr As DataTable)
        Err = "未找到相应记录!" & vbCrLf & "源表:" & dtErr.TableName
    End Function

    Public Function ThrowMustRecordSignatureSubmitor()
        Err = "登记签约提交人必须是风险部负责人!"
    End Function

End Class
