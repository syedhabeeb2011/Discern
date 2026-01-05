/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 10/30/2019
 * Time: 8:40 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Threading;
using System.Net;
using System.Net.Mail;
using System.Configuration;
using System.Collections.Generic;

namespace Discern_Notification
{
    /// <summary>
    /// Description of EmailHelper.
    /// </summary>
    public class EmailHelper
    {
        public string _senderEmail { get; set; }
        public string _senderPasword { get; set; }
        public string _host { get; set; }
        public int _port { get; set; }

        public EmailHelper(string senderEmail, string senderPassword, string host, int port)
        {
            _senderEmail = senderEmail;
            _senderPasword = _senderPasword;
            _host = host;
            _port = port;
        }

        public void SendEMail(string receiverEmail, string subject, string messageBody, string attachmentFilePath)
        {
            Configuration _configuration = ConfigurationManager.OpenExeConfiguration(this.GetType().Assembly.Location);

            string senderEmail = _configuration.AppSettings.Settings["Email.SenderEmail"] != null ? _configuration.AppSettings.Settings["Email.SenderEmail"].Value : string.Empty; //"osman.alikhan@atos.net"; //ConfigurationManager.AppSettings["Email.SenderEmail"]; //
            string senderPasword = _configuration.AppSettings.Settings["Email.SenderPassword"] != null ? _configuration.AppSettings.Settings["Email.SenderPassword"].Value : string.Empty; //ConfigurationManager.AppSettings["Email.Password"];//
            string host = _configuration.AppSettings.Settings["Email.Host"] != null ? _configuration.AppSettings.Settings["Email.Host"].Value : string.Empty;//"10.14.32.146";// ConfigurationManager.AppSettings["Email.Host"]; //"10.14.32.146";
            int port = _configuration.AppSettings.Settings["Email.Port"] != null ? Convert.ToInt32(_configuration.AppSettings.Settings["Email.Port"].Value) : 0;//25;//Convert.ToInt32(ConfigurationManager.AppSettings["Email.Port"]); //25;


            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = _host;
            smtpClient.Port = _port;

            //Server Credentials
            NetworkCredential NC = new NetworkCredential();
            NC.UserName = _senderEmail;
            NC.Password = _senderPasword;
            //assigned credetial details to server
            smtpClient.Credentials = NC;

            MailMessage Mymessage = new MailMessage(_senderEmail, receiverEmail);
            Mymessage.Subject = subject;
            Mymessage.Body = messageBody;
            Mymessage.IsBodyHtml = true;
            if (!string.IsNullOrEmpty(attachmentFilePath))
            {
                Mymessage.Attachments.Add(new Attachment(attachmentFilePath));
            }
            //sends the email
            smtpClient.Send(Mymessage);


        }

        public void SendEMail(string receiverEmail, string subject, string messageBody, List<string> attachments, bool isBodyHTML = true)
        {


            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = _host;
            smtpClient.Port = _port;

            //Server Credentials
            NetworkCredential NC = new NetworkCredential();
            NC.UserName = _senderEmail;
            NC.Password = _senderPasword;
            //assigned credetial details to server
            smtpClient.Credentials = NC;

            MailMessage Mymessage = new MailMessage(_senderEmail, receiverEmail);
            Mymessage.Subject = subject;
            Mymessage.Body = messageBody;
            Mymessage.IsBodyHtml = isBodyHTML;

            if (attachments.Count > 0)
            {
                foreach (var attachment in attachments)
                {
                    Mymessage.Attachments.Add(new Attachment(attachment));
                }
            }
            //sends the email
            smtpClient.Send(Mymessage);


        }
    }
}
