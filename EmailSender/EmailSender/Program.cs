using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net.Mime;
using System.IO;

namespace MyEmailTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string server = "";
            int port = 0;
            string sender = "";
            string[] receivers = null;
            string[] attachments = null;
            string subject = "";

            // read config.txt
            StreamReader sr = new StreamReader("config.txt");
            string line = sr.ReadLine();
            while (line != null)
            {
                if (line == "" || line[0] == '#')
                {
                    line = sr.ReadLine();
                    continue;
                }
                string[] line_eles = line.Split(new char[]{'='});
                if (line_eles[0] == "server")
                    server = line_eles[1];
                else if (line_eles[0] == "port")
                    port = Convert.ToInt32(line_eles[1]);
                else if (line_eles[0] == "subject")
                    subject = line_eles[1];
                else if (line_eles[0] == "sender")
                    sender = line_eles[1];
                else if (line_eles[0] == "receivers")
                    receivers = line_eles[1].Split(new char[] { ',' });
                else if (line_eles[0] == "attachments")
                    attachments = line_eles[1].Split(new char[] { ',' });
                line = sr.ReadLine();
            }
            
            string body = File.ReadAllText("report.html");

            MailMessage mail = new MailMessage();
            mail.Subject = subject;
            mail.From = new MailAddress(sender);

            foreach (string s in receivers)
                mail.To.Add(s);

            foreach (string s in attachments)
            {
                Attachment data = new Attachment("csv_report\\" + s, MediaTypeNames.Application.Octet);
                mail.Attachments.Add(data);
            }

            AlternateView htmlView = AlternateView.CreateAlternateViewFromString(body, null,
                  MediaTypeNames.Text.Html);
            mail.AlternateViews.Add(htmlView);
            mail.IsBodyHtml = true;
            
            SmtpClient client = new SmtpClient();
            //client.Credentials = new System.Net.NetworkCredential("alerter.add", "Welcome1");
            client.Port = port;
            client.Host = server;
            //client.EnableSsl = true;
            try
            {
                client.Send(mail);
                System.Console.WriteLine("mail send ok");
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("error:" + ex.Message);
            }
        }

    }
}