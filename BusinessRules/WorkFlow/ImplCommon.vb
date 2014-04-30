Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplCommon
    Implements ICondition

    Public Function GetResult(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal expFlag As String, ByVal transCondition As String) As Boolean Implements ICondition.GetResult


        '客户端传来的表达式是否与转移条件相符
        If transCondition = ".T." Then
            Return True
        Else
            If transCondition = expFlag Then
                Return True
            Else
                Return False
            End If
        End If

    End Function

End Class
