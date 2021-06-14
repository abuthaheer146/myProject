using System;
using System.Configuration;
using System.Data; // Use ADO.NET namespace
using System.Data.SqlClient; // Use SQL Server data provider namespace
using System.IO;
using System.Text;


namespace CMS.eCMSEPESAdminBatch
{
    /// <summary>
    /// Summary description for Class1.
    /// </summary>
    class MainClass
    {
        /// <summary>
        /// The main entry point for the application.a
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("Please provide with only one paramter to run the application");
                return;
            }
            try
            {
                string rptName = args[0].ToLower();
                switch (rptName)
                {
                    case "suspendnoactivityuserreport":
                        GenerateSuspendedUsersList();
                        break;

                    case "suspendnoactivityusers":
                        GenerateSuspendedUsers();
                        break;

                    case "activeuserslist":
                        GenerateActiveUserList();
                        break;

                    case "ssnetstats":
                        GenerateSSNetStats();
                        break;
                }

                Console.WriteLine("Program finished.");
            }
            catch (System.Exception ex)
            {
                //ErrorHandler.GetInstance().LogException(ex); 
                string strMessage = ex.Message;
                if (ex.InnerException != null) strMessage += ex.InnerException.Message;

                ErrorHandler.log(strMessage, "Exception");
                //log(ex.InnerException.Message, "Excpetion");
            }
        }
        static void GenerateActiveUserList()
        {
            string strSQL;

            string[] strAttachment = new string[2];

            strSQL = "select users.email as [Email Address], users.firstname+', '+lastname as [Name], agencies.agencyname as [Agency], "
                                    + "case len(isnull(officetel,'')) when 0 then '' else officetel+', ' end "
                                    + " + case len(isnull(hometel,'')) when 0 then '' else hometel+', ' end "
                                    + " + case len(isnull(mobile,'')) when 0 then '' else mobile end as [Contact Number],"
                                    + "Convert(nvarchar(17), login.token_time_stamp, 113) as [Date and Time Last Login], "
                                    + "case isnull(users.roleID,0) when 0 then 'No' else 'Yes' end as [eCMS User],"
                                    + "isnull(CR.RoleName,'') as [eCMS Role],"
                                    + "case isnull(DDRroleID,0) when 0 then 'No' else 'Yes' end as [DDR User], "
                                    + "isnull(DR.role_name,'')  as [DDR Role] "
                                    + "from users inner join [EPESDB].[dbo].[TOKEN_VALIDATION] login "
                                    + " on login.nirc=users.nric and login.destination='ECMS' and datediff(month,login.token_time_stamp, getdate())<12 "
                                    + " inner join agencies "
                                    + " on agencies.agencyid=users.agencyid "
                                    + " left join Roles CR on isnull(users.roleID,0)=CR.roleID "
                                    + " left join [DDR_DB].[dbo].[ACM_ROLE] DR on isnull(DDRroleID,0)=DR.role_id "
                                    + " where isnull(disabled,0)<>1 and isnull(deleted,0)<>1 "
                                    + "and isnull(agencies.isProduction,0)=1 "
                                    + "order by agencyname";
            strAttachment[0] = GetUsersList("eCMSDDR", strSQL);

            strSQL = "USE [EPESDB] "
                + " SELECT  P.email as [Email Address], p.first_name + ', ' +p.last_name as [Name],"
                + " p.PRIMARY_CONTACT as [Contact Number], "                // v.VWO_NAME as [Programme/Agency],"
                + " case ORGANISATION_ID when 'VWO' then v.VWO_NAME "
                + " when 'NCSS' then 'NCSS' "
                + " when 'EXTADM' then 'External Administrator' end as [Programme/Agency], "
                + " Convert(nvarchar(17), max(L.token_time_stamp), 113) as [Date and Time Last Login], "
                + " case ORGANISATION_ID when 'VWO'then 'Yes' else 'No' end as [VWO User], "
                + " 'Yes' as [Administrator User] "
                + " FROM [EPESDB].[dbo].[TOKEN_VALIDATION] L "
                + " inner join T_IC_USERS U on L.NIRC=U.USER_NAME and isnull(U.IS_DELETED,0)<>1 "
                + " inner join EPES_USER_PROFILE P on P.user_id=U.user_id "
                + " left join EPES_AMS_M_VWO V on p.VWO_ID=V.VWO_ID "
                + " where datediff(month,L.token_time_stamp, getdate())<12 "  //and L.destination='EPMS' "
                + " and L.NIRC not in (select NRIC from CMS_PROD.dbo.users) "
                + " and U.user_ID in (select distinct user_ID from T_IC_USERS_IN_ROLES where role_ID in (27,42)) "
                + " group by L.Nirc, P.EMAIL,p.FIRST_NAME,p.LAST_NAME,p.PRIMARY_CONTACT, p.ORGANISATION_ID,	v.VWO_NAME "
                + " union all "
                + " SELECT  P.email, p.first_name + ', ' +p.last_name, "
                + " p.PRIMARY_CONTACT, " // v.VWO_NAME, "
                + " case ORGANISATION_ID when 'VWO' then v.VWO_NAME "
                + " when 'NCSS' then 'NCSS' "
                + " when 'EXTADM' then 'External Administrator' end as [Programme/Agency], "
                + " Convert(nvarchar(17), max(L.token_time_stamp), 113),"
                + " case ORGANISATION_ID when 'VWO'then 'Yes' else 'No' end,"
                + " 'No' as [Administrator User] "
                + " FROM [EPESDB].[dbo].[TOKEN_VALIDATION] L "
                + " inner join T_IC_USERS U on L.NIRC=U.USER_NAME and isnull(U.IS_DELETED,0)<>1 "
                + " inner join EPES_USER_PROFILE P on P.user_id=U.user_id "
                + " left join EPES_AMS_M_VWO V on p.VWO_ID=V.VWO_ID "
                + " where datediff(month,L.token_time_stamp, getdate())<12 " // and L.destination='EPMS' "
                + " and L.NIRC not in (select NRIC from CMS_PROD.dbo.users) "
                + " and U.user_ID in (select distinct user_ID from T_IC_USERS_IN_ROLES where role_ID not in (27,42)) "
                + " and U.user_ID not in (select distinct user_ID from T_IC_USERS_IN_ROLES where role_ID in (27,42)) "
                + " group by L.Nirc, P.EMAIL,p.FIRST_NAME,p.LAST_NAME,p.PRIMARY_CONTACT, p.ORGANISATION_ID,	v.VWO_NAME "
                + " order by [Programme/Agency]";

            EmailSender sender = new EmailSender();
            strAttachment[1] = GetUsersList("EPES", strSQL);

            var rptConfig = (System.Collections.Specialized.NameValueCollection)System.Configuration.ConfigurationManager.GetSection("ActiveUsersList");
            sender.To = (string)rptConfig["smtpTo"];
            sender.CC = (string)rptConfig["smtpCC"];
            sender.Subject = (string)rptConfig["subject"] + " " + DateTime.Today.ToString("dd MMMM yyyy");
            sender.Body = (string)rptConfig["mailBody"];

            sender.SendAttach(strAttachment);
        }
        static void GenerateSSNetStats()
        {
            string strSQL;

            string[] strAttachment = new string[2];

            var rptConfig = (System.Collections.Specialized.NameValueCollection)System.Configuration.ConfigurationManager.GetSection("SSNetStats");
            strSQL = (string)rptConfig["StoredProcedure"];
            string filename = (string)rptConfig["ReportName"];

            EmailSender sender = new EmailSender();
            strAttachment[0] = GetRepors(filename, strSQL);

            sender.To = (string)rptConfig["smtpTo"];
            sender.CC = (string)rptConfig["smtpCC"];
            sender.Subject = (string)rptConfig["subject"] + " " + DateTime.Today.ToString("dd MMMM yyyy");
            sender.Body = (string)rptConfig["mailBody"];

            sender.SendAttach(strAttachment);
        }
        static void GenerateSuspendedUsers()
        {
            try
            {
                var rptConfig = (System.Collections.Specialized.NameValueCollection)ConfigurationManager.GetSection("SuspendedUsersList");
                string connectionString = ConfigurationManager.ConnectionStrings[1].ToString();
                ErrorHandler.log(connectionString, "connectionString");
                string strSQLGetList = (string)rptConfig["StoredProcedureGetList"];  // SQL stored procedure name pr_Get_InActivity_Users_Within_Days
                string strSQLSuspend = (string)rptConfig["StoredProcedureSuspend"];  // SQL stored procedure name pr_Suspend_InActivity_User_SelectByID
                string daysforwarning = (string)rptConfig["daysforwarning"];
                string daysforSuspend = (string)rptConfig["daysforSuspend"];
                DataSet ds90 = new DataSet();
                DataSet ds85 = new DataSet();
                ///90 days suspension and email notification 
                using (SqlConnection con = new SqlConnection(connectionString))
            
                {
                    using (SqlCommand cmd = new SqlCommand(strSQLGetList, con))
                    {
                        SqlDataAdapter da = new SqlDataAdapter();
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@within_days", Convert.ToInt32(daysforSuspend));
                        da.SelectCommand = cmd;
                        da.Fill(ds90);
                    }
                }
                using (SqlConnection con = new SqlConnection(connectionString))

                    if (ds90 != null && ds90.Tables != null && ds90.Tables.Count > 0)
                    {
                        //suspension of the inactive user 
                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            using (SqlCommand cmd = new SqlCommand(strSQLSuspend, connection))
                            {

                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.Add("@userID", SqlDbType.Int, 8);

                                connection.Open();
                                foreach (DataRow InActiveUser in ds90.Tables[0].Rows)
                                {
                                    int userid = 0;
                                    if (InActiveUser["UserID"] != null && int.TryParse(InActiveUser["UserId"].ToString(), out userid))
                                    {
                                        cmd.Parameters["@userID"].Value = userid;
                                        cmd.ExecuteNonQuery();
                                        //Send notification email to notify suspension for inactivity to user
                                        Console.WriteLine("Sending email notification for suspened user.......");
                                        SendNotificationToSuspended(InActiveUser, "SUSPEND");
                                        
                                    }
                                }
                            }

                        }
                    }
                ///85 days warning email notification 
                using (SqlConnection newcon = new SqlConnection(connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(strSQLGetList, newcon))
                    {
                        SqlDataAdapter da = new SqlDataAdapter();
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@within_days", Convert.ToInt32(daysforwarning));
                        da.SelectCommand = cmd;
                        da.Fill(ds85);
                    }

                    if (ds85 != null && ds85.Tables != null && ds85.Tables.Count > 0)
                    {
                        foreach (DataRow InActiveUser in ds85.Tables[0].Rows)
                        {
                            int userid = 0;
                            if (InActiveUser["UserID"] != null && int.TryParse(InActiveUser["UserId"].ToString(), out userid))
                            {
                                //Send notification email to warn inactivity to user
                                Console.WriteLine("Sending email notification for warning email.......");
                                SendNotificationToSuspended(InActiveUser, "WARNING");
                                
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string strMessage = ex.Message;
                if (ex.InnerException != null) strMessage += ex.InnerException.Message;
                ErrorHandler.log(strMessage, "Excpetion");
            }

        } // Method to generate users greater than 90 days of inactivity 
        static void SendNotificationToSuspended(DataRow tmprow, string action)
        {
            try
            {
                var rptConfig = (System.Collections.Specialized.NameValueCollection)ConfigurationManager.GetSection("SuspendedUsersList");
                EmailSender sender = new EmailSender();
                string[] strAttachment = new string[0];
                string getFullName = string.Empty;
                if (!tmprow.IsNull("UserName"))
                {
                    getFullName = tmprow["UserName"].ToString().Trim();
                }
                sender.To = tmprow["Email"].ToString().Trim();
                if (action == "SUSPEND")
                {
                    sender.Subject = (string)rptConfig["subject"] + " " + DateTime.Today.ToString("dd MMMM yyyy");
                    sender.Body = SetBody(getFullName, (string)rptConfig["suspendmailText"], (string)rptConfig["suspendmailText"]);
                }
                else if (action == "WARNING")
                {
                    sender.Subject = (string)rptConfig["warsubject"] + " " + DateTime.Today.ToString("dd MMMM yyyy");
                    sender.Body = SetBody(getFullName, (string)rptConfig["warnmailText"], (string)rptConfig["warnmailText"]);
                }
                sender.CC = (string)rptConfig["smtpCC"];
                sender.SendAttach(strAttachment);

            }
            catch (Exception ex)
            {
                string strMessage = ex.Message;
                if (ex.InnerException != null) strMessage += ex.InnerException.Message;
                ErrorHandler.log(strMessage, "Excpetion");
            }

        } // // Method to send "UR Account is suspended" email to users greater than 90 days o inactivity 
        static string SetBody(string userName, string system, string bodyText)
        {
            string getBody = bodyText.Replace("SYSTEMNAME", system).Replace("BR", "<BR>");
            StringBuilder sb = new StringBuilder();
            sb.Append("Dear ");
            sb.Append(userName + " ,");
            sb.Append("<br>");
            sb.Append(getBody);
            return sb.ToString();
        }
        static void GenerateSuspendedUsersList()
        {
            try
            {
                var rptConfig = (System.Collections.Specialized.NameValueCollection)System.Configuration.ConfigurationManager.GetSection("SuspendedUsersList");
                string connectionString = ConfigurationManager.ConnectionStrings["DBConn"].ToString();
                string strSQLGetList = (string)rptConfig["StoredProcedureGetList"];  // SQL stored procedure name pr_Get_InActivity_Users_Within_Days
                DataSet ds = new DataSet();
                int getSuspendDays = Convert.ToInt16((string.IsNullOrEmpty((string)rptConfig["SuspendDays"]) ? "0" : (string)rptConfig["SuspendDays"]));

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(strSQLGetList, con))
                    {
                        SqlDataAdapter da = new SqlDataAdapter();
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@within_days", getSuspendDays);
                        da.SelectCommand = cmd;
                        da.Fill(ds);
                    }
                }

                if (ds != null && ds.Tables != null && ds.Tables.Count > 0)
                {
                    //Send suspended user list email to Administrator
                    SendSuspendedListReport(ds.Tables[0]);
                }
            }
            catch (Exception ex)
            {

                string strMessage = ex.Message;
                if (ex.InnerException != null) strMessage += ex.InnerException.Message;
                ErrorHandler.log(strMessage, "Excpetion");

            }

        }
        static void SendSuspendedListReport(DataTable dt)
        {
            try
            {
                var rptConfig = (System.Collections.Specialized.NameValueCollection)System.Configuration.ConfigurationManager.GetSection("SuspendedUsersList");
                EmailSender sender = new EmailSender();
                string filename = (string)rptConfig["ReportName"];
                string[] strAttachment = new string[1];
                strAttachment[0] = GetReports(filename, dt);
                sender.To = (string)rptConfig["smtpTo"];
                sender.CC = (string)rptConfig["smtpCC"];
                sender.Subject = (string)rptConfig["subject"] + " " + DateTime.Today.ToString("dd MMMM yyyy");
                sender.Body = (string)rptConfig["mailBody"];
                sender.SendAttach(strAttachment);
            }
            catch (Exception ex)
            {
                string strMessage = ex.Message;
                if (ex.InnerException != null) strMessage += ex.InnerException.Message;
                ErrorHandler.log(strMessage, "Excpetion");
            }

        }
        static string GetReports(string filename, DataTable dt)
        {
            string strFilename = Environment.CurrentDirectory + "\\report\\" + filename + "_" + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            CreateXLSFile(dt, strFilename);
            return strFilename;
        }
        static void CreateXLSFile(DataTable dt, string strFileName)
        {

            StreamWriter sw = new StreamWriter(strFileName, false);

            int iColCount = dt.Columns.Count;

            // Write Columns to excel file
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sw.AutoFlush = true;
                sw.Write(dt.Columns[i].ToString().ToUpper() + "\t");
            }

            sw.WriteLine();
            //write rows to excel file
            for (int i = 0; i < (dt.Rows.Count); i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sw.AutoFlush = true;
                    if (dt.Rows[i][j] != null)
                    {
                        sw.Write(Convert.ToString(dt.Rows[i][j]) + "\t");
                    }
                    else
                    {
                        sw.Write("\t");
                    }
                }
                sw.WriteLine();
            }
            sw.Close();

        }
        static string GetUsersList(string strType, string strSQL)
        {
            string strFilename = Environment.CurrentDirectory + "\\report\\" + strType + "_ActiveUsers_" + DateTime.Now.ToString("yyyyMMdd") + ".xls";

            string connectionString = ConfigurationManager.ConnectionStrings["DBConn"].ToString();
            SqlConnection con = new SqlConnection(connectionString);
            using (con)
            {
                SqlCommand command = new SqlCommand(strSQL, con);
                con.Open();
                SqlDataReader reader = command.ExecuteReader();
                CreateXLSFile(reader, strFilename);
            }

            return strFilename;
        }
        static string GetRepors(string filename, string strSQL)
        {
            string strFilename = Environment.CurrentDirectory + "\\report\\" + filename + "_" + DateTime.Now.ToString("yyyyMMdd") + ".xls";

            string connectionString = ConfigurationManager.ConnectionStrings["DBConn"].ToString();
            SqlConnection con = new SqlConnection(connectionString);
            using (con)
            {
                SqlCommand command = new SqlCommand();
                command.Connection = con;
                command.CommandText = strSQL;
                command.CommandTimeout = 300;
                con.Open();
                SqlDataReader reader = command.ExecuteReader();
                CreateXLSFile(reader, strFilename);
            }

            return strFilename;
        }
        static void CreateXLSFile(SqlDataReader dt, string strFileName)
        {

            StreamWriter sw = new StreamWriter(strFileName, false);

            using (sw)
            {
                for (int i = 0; i < dt.FieldCount; i++)
                {
                    sw.AutoFlush = true;
                    sw.Write(dt.GetName(i) + "\t");
                }

                sw.Write("\n");

                while (dt.Read())
                {
                    for (int i = 0; i < dt.FieldCount; i++)
                    {
                        sw.AutoFlush = true;
                        sw.Write(dt[i].ToString() + "\t");
                    }
                    sw.Write("\n");
                }
            }
        }
        //static void log(string strInfo, string strType)
        //{
        //    try
        //    {
        //        string logFile = Environment.CurrentDirectory + "\\log\\log_" + System.DateTime.Today.ToString("yyyyMMdd") + ".log";
        //        StreamWriter sw = new StreamWriter(logFile, true);
        //        using (sw)
        //        {
        //            sw.WriteLine(DateTime.Now.ToString() + ":" + strType + ":" + strInfo);
        //            sw.Close();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        string strMessage = ex.Message;
        //        if (ex.InnerException != null) strMessage += ex.InnerException.Message;
        //        log(strMessage, "Excpetion");
        //    }
        //}
    }
}



