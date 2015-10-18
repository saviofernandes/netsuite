using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.ComponentModel;

namespace ConsoleApplication2
{
    class Email
    {

        public void test() {
            MailMessage mail = new MailMessage("savio@starlinkme.net", "savio@starlinkme.net");
            mail.IsBodyHtml = true;
            mail.Subject = "Test email from .NET";
            mail.Body = @"<html><body><h1>Hello world</h1></body></html>";
            SmtpClient client = new SmtpClient("smtp.office365.com");
            client.Port=587;
            client.EnableSsl = true;
            client.UseDefaultCredentials = false; /*Important: This line of code must be executed 
            before setting the NetworkCredentials object, otherwise the setting will be reset (a bug in .NET)*/











            System.Net.NetworkCredential cred = new System.Net.NetworkCredential("savio@starlinkme.net","Cl@nd3st1n3");
            client.Credentials = cred;
            client.Send(mail);
        }
        

        
    }
}
