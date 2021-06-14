using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CMS.eCMSEPESAdminBatch
{
    public static class ErrorHandler
    {
        public static void log(string strInfo, string strType)
        {
            try
            {
                string logFile = Environment.CurrentDirectory + "\\log\\log_" + System.DateTime.Today.ToString("yyyyMMdd") + ".log";
                StreamWriter sw = new StreamWriter(logFile, true);
                using (sw)
                {
                    sw.WriteLine(DateTime.Now.ToString() + ":" + strType + ":" + strInfo);
                    sw.Close();
                }
            }
            catch (Exception ex)
            {
                string strMessage = ex.Message;
                if (ex.InnerException != null) strMessage += ex.InnerException.Message;
                log(strMessage, "Excpetion");
            }
        }
    }
}
