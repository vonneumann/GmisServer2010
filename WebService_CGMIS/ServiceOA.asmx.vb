Imports System.Data.SqlClient
Imports System.Web.Services
Imports System.Configuration
Imports System.Web.Services.Protocols
Imports BusinessRules

<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
Public Class ServiceOA
    Inherits System.Web.Services.WebService

    Private strConn As String = ConfigurationSettings.AppSettings("DBConnectionOA")

    '通用查询
    <WebMethod()> Public Function GetCommonQueryInfoForOA(ByVal strSql As String) As DataSet
        Dim conn As New SqlConnection(strConn)
        conn.Open()
        Dim ts As SqlTransaction = conn.BeginTransaction
        Try
            Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
            GetCommonQueryInfoForOA = CommonQuery.GetCommonQueryInfoForOA(strSql)
            ts.Commit()
        Catch dbEx As DBConcurrencyException
            ts.Rollback()
            Throw dbEx
        Catch oEx As Exception
            ts.Rollback()
            Throw oEx
        Finally
            conn.Close()
            conn.Dispose()
        End Try
    End Function

End Class