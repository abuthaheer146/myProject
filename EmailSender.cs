using System;
using System.Configuration;
using System.Reflection;
using Microsoft.Build.Utilities;




namespace CMS.eCMSEPESAdminBatch
{
	/// <summary>
	/// Summary description for MailSender.
	/// </summary>
	public class EmailSender
	{
		public string From;
		public string To;
		public string CC;
		public string smtpServer;
		public string port;
		public string Subject;
		public string Body;
		public string Username;
		public string Password;

		public EmailSender()
		{
            From = ConfigurationManager.AppSettings["smtpFrom"];
            smtpServer = ConfigurationManager.AppSettings["smtpServer"];
            Username = ConfigurationManager.AppSettings["smtpServer"];
            port = ConfigurationManager.AppSettings["smtpPort"];
        }

		//public void SendAttach(string sTo, string sSubject, string sBody, string sAttach)
		public void SendAttach(string[] sAttach)
		{
            System.Net.Mail.MailMessage mailMessage = new System.Net.Mail.MailMessage();
            using (mailMessage)
            {
                try
                {
                    mailMessage.IsBodyHtml = true;
                    mailMessage.To.Add(To);
                    if (CC.Length>0)
                        mailMessage.CC.Add(CC);
                    mailMessage.Subject = Subject;
                    mailMessage.Body = Body;
                    mailMessage.From = new System.Net.Mail.MailAddress(From);

                    for (int i=0; i<sAttach.Length; i++)
                    {
                        if (!string.IsNullOrEmpty(sAttach[i]))
                            mailMessage.Attachments.Add(new System.Net.Mail.Attachment(sAttach[i]));
                    }

                    System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient(smtpServer);
                    client.Port =int.Parse(port);

                    client.Credentials = new System.Net.NetworkCredential(Username, Password);
                    client.UseDefaultCredentials = false; 
                    client.Send(mailMessage);
                }
                catch (System.Exception ex)
                {
                    string strMessage = ex.Message;
                    if (ex.InnerException != null) strMessage += ex.InnerException.Message;
                    ErrorHandler.log(strMessage, "Excpetion");
                }
            }
		}	
	}
}
