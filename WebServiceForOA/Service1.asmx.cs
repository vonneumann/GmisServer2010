using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Web.Services.Description;
using System.Configuration;
using BusinessRules;
using System.Data.SqlClient;

namespace WebServiceForOA
{
    /// <summary>
    /// Service1 的摘要说明
    /// </summary>
    [WebService(Namespace = "http://action.xingye/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.None)]
    [System.ComponentModel.ToolboxItem(false)]
    public class Service1 : System.Web.Services.WebService
    {
        [SoapRpcMethod(Use = SoapBindingUse.Literal, Action = "http://action.xingye/HelloWorld", RequestNamespace = "http://action.xingye/", ResponseNamespace = "http://action.xingye/")]
        [WebMethod]
        public string HelloWorld(string test)
        {
            return "Hello World you like ME?";
        }

        string strConn = ConfigurationSettings.AppSettings["DBConnection"];


        [SoapRpcMethod(Use = SoapBindingUse.Literal, Action = "http://action.xingye/finishedTask", RequestNamespace = "http://action.xingye/", ResponseNamespace = "http://action.xingye/")]
        [WebMethod]
        public string finishedTask(string workflowID, string projectID, string finishedTaskID, string finishedFlag, string userID)
        {
            SqlConnection conn = new SqlConnection(strConn);
            conn.Open();
            SqlTransaction trans = conn.BeginTransaction();
            try
            {
                WorkFlow WorkFlow = new WorkFlow(conn, ref trans);
                WorkFlow.finishedTask(workflowID, projectID, finishedTaskID, finishedFlag, userID);
                trans.Commit();
                return "1";
            }
            catch (WorkFlowErr errWf)
            {
                trans.Rollback();
                return errWf.ErrMessage.ToString();
            }
            catch (Exception e)
            {
                trans.Rollback();
                return e.Message;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
       
        }
    }
}