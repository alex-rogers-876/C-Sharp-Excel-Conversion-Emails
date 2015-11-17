using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace gemTest
{
    class SendMail
    {

        public string subject { get; set; }
        public string body { get; set; }
        public string to { get; set; }
        public string smtpAddres { get; set; }
        public Outlook.Application application { get; set; }
        
       public SendMail(Outlook.Application application, string subject, string body, string to, string smtpAddres)
        {
            this.application = application;
            this.subject = subject;
            this.body = body;
            this.to = to;
            this.smtpAddres = smtpAddres;

        }
        //Outlook.Application application, string subject, string body, string to, string smtpAddress
        public void SendEmailFromAccount()
        {

            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential(MaillConst.Username, MaillConst.Password, "tjhpayroll");
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(MaillConst.From);

            smtpClient.Host = MaillConst.SmtpServer;
            smtpClient.Port = 587;
            smtpClient.UseDefaultCredentials = false;

            System.Net.NetworkCredential credentials =
          new System.Net.NetworkCredential(MaillConst.Username, MaillConst.Password);

            smtpClient.EnableSsl = true;
            smtpClient.Credentials = credentials;
            smtpClient.Timeout = (60 * 5 * 1000);

            message.From = fromAddress;
            message.Subject = subject + " - " + DateTime.Now.Date.ToString().Split(' ')[0];
            message.Bcc.Add(fromAddress);
            message.IsBodyHtml = true;
            message.Body = body;
            message.To.Add(to);

            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            try
            {
                smtpClient.Send(message);
                Console.WriteLine("Email sent");
            }
            catch(Exception ex)
            {
                throw (ex);
            }
        }
    }
}
