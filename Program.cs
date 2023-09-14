using ClosedXML.Excel;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace EmailSender_V4
{
    internal class Program
    {
        const string path = @"C:\Users\tarso\Desktop\Solutions\EmailSender_v4\Resource\Contacts.xlsx";

        static void Main(string[] args)
        {
            // Obter Mailing  
            Console.WriteLine("Programa de Envio de E-mails");
            Console.WriteLine("------------------------------------------------------");
            Console.WriteLine("1. Obtendo Mailing");
            var xls = new XLWorkbook(path);
            var page = xls.Worksheets.First(x => x.Name == "Planilha1");
            var totalRows = page.Rows().Count();
            List<string> mailing = new List<string>();

            for (int i = 2; i <= totalRows; i++)
            {
                var endereco = page.Cell($"E{i}").Value.ToString().Trim();

                if (endereco.Contains("conhec") || string.IsNullOrEmpty(endereco) || endereco.Contains("ulo"))
                    continue;

                mailing.Add(endereco);
            }

            // Teste
            //mailing = new List<string>() { "tarsocoelho98@gmail.com", "asjdhfjkshfjks", "clgamesoftware@gmail.com" };

            // Montagem do e-mail.
            Console.WriteLine("2. Montagem de Email");
            HtmlDocument html = new HtmlDocument();
            html.Load(@"C:\Users\tarso\Desktop\Solutions\EmailSender_v4\Resource\Email.html");
            string final = html.Text;
            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            message.From = new MailAddress("contact@clgamesoft.com");
            message.Subject = "Mobile Game Publishing";
            message.IsBodyHtml = true;
            message.Body = final;

            SmtpClient client = new SmtpClient("smtp.gmail.com", Convert.ToInt32("587"));
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential("#email#", "#key#");
            client.EnableSsl = true;
            Console.WriteLine("3. Iniciando envio.");
            Console.WriteLine("------------------------------------------------------");

            int count = 1;
            // Envio dos E-mails
            foreach (var element in mailing)
            {
                try
                {
                    message.To.Clear();
                    message.To.Add(element);
                    client.Send(message);
                    Console.WriteLine(string.Concat(count, ". Sucesso de envio: ", element));
                }
                catch (Exception e)
                {
                    Console.WriteLine(string.Concat(count, ". Falha de envio: ", element));
                }
                count++;
            }

            Console.ReadKey();
        }
    }
}
