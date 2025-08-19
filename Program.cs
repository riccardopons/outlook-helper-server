using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookHelperServer
{
    class MailRequest
    {
        public string Email { get; set; }
        public string[] Parcels { get; set; }
    }

    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            // Lancement du serveur HTTP local
            HttpListener listener = new HttpListener();
            listener.Prefixes.Add("http://localhost:5001/");
            listener.Start();
            Console.WriteLine("✅ OutlookHelperServer démarré sur http://localhost:5001/");

            while (true)
            {
                var context = listener.GetContext();
                ThreadPool.QueueUserWorkItem(_ => HandleRequest(context));
            }
        }

        static void HandleRequest(HttpListenerContext context)
        {
            try
            {
                if (context.Request.HttpMethod == "POST" &&
                    context.Request.Url.AbsolutePath == "/send-mails")
                {
                    using var reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding);
                    string body = reader.ReadToEnd();
                    var mails = JsonSerializer.Deserialize<MailRequest[]>(body);

                    if (mails != null)
                    {
                        GenerateMails(mails);
                    }

                    byte[] buffer = Encoding.UTF8.GetBytes("{\"status\":\"ok\"}");
                    context.Response.ContentType = "application/json";
                    context.Response.OutputStream.Write(buffer, 0, buffer.Length);
                }
                else
                {
                    context.Response.StatusCode = 404;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erreur : " + ex.Message);
                context.Response.StatusCode = 500;
            }
            finally
            {
                context.Response.OutputStream.Close();
            }
        }

        static void GenerateMails(MailRequest[] mails)
        {
            Outlook.Application outlookApp;
            try
            {
                outlookApp = new Outlook.Application();
            }
            catch
            {
                Console.WriteLine("⚠️ Outlook non disponible");
                return;
            }

            foreach (var req in mails)
            {
                if (string.IsNullOrWhiteSpace(req.Email)) continue;

                string colisList = string.Join(", ", req.Parcels);
                string subject = "Colis : " + colisList;

                Outlook.MailItem mail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                mail.To = req.Email;
                mail.Subject = subject;
                mail.Body = $"Bonjour,\n\nVos colis suivants sont en attente : {colisList}\n\nMerci de confirmer vos informations.\n\nCordialement.";
                mail.Display(); // → .Send() pour envoyer directement
            }
        }
    }
}