using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookHelperServer
{
    public class MailRequest
    {
        public string Email { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string[] Parcels { get; set; } = Array.Empty<string>();
    }

    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            HttpListener listener = new HttpListener();
            listener.Prefixes.Add("http://localhost:5000/");
            listener.Start();
            Console.WriteLine("✅ OutlookHelperServer démarré sur http://localhost:5000/");

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
                // --- CORS ---
                context.Response.AddHeader("Access-Control-Allow-Origin", "*");
                context.Response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
                context.Response.AddHeader("Access-Control-Allow-Headers", "Content-Type");

                if (context.Request.HttpMethod == "OPTIONS")
                {
                    context.Response.StatusCode = 200;
                    context.Response.OutputStream.Close();
                    return;
                }

                // --- Health endpoint ---
                if (context.Request.HttpMethod == "GET" && context.Request.Url.AbsolutePath == "/health")
                {
                    byte[] buffer = Encoding.UTF8.GetBytes("{\"status\":\"ok\"}");
                    context.Response.ContentType = "application/json";
                    context.Response.OutputStream.Write(buffer, 0, buffer.Length);
                    context.Response.OutputStream.Close();
                    return;
                }

                // --- Send endpoint ---
                if (context.Request.HttpMethod == "POST" && context.Request.Url.AbsolutePath == "/send")
                {
                    Console.WriteLine($"📨 Requête reçue: {context.Request.HttpMethod} {context.Request.Url.AbsolutePath}");

                    using var reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding);
                    string body = reader.ReadToEnd();
                    Console.WriteLine("Corps reçu : " + body);

                    var payload = JsonSerializer.Deserialize<MailRequest[]>(body);
                    if (payload == null)
                    {
                        Console.WriteLine("⚠️ Payload JSON invalide");
                        context.Response.StatusCode = 400;
                    }
                    else
                    {
                        Console.WriteLine($"⚡ {payload.Length} emails reçus");
                        GenerateMails(payload);
                        context.Response.StatusCode = 200;
                    }

                    byte[] respBuffer = Encoding.UTF8.GetBytes("{\"status\":\"ok\"}");
                    context.Response.ContentType = "application/json";
                    context.Response.OutputStream.Write(respBuffer, 0, respBuffer.Length);
                    context.Response.OutputStream.Close();
                    return;
                }

                context.Response.StatusCode = 404;
                context.Response.OutputStream.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erreur HandleRequest : " + ex);
                try
                {
                    context.Response.StatusCode = 500;
                    context.Response.OutputStream.Close();
                }
                catch { }
            }
        }

        static void GenerateMails(MailRequest[] mails)
        {
            Outlook.Application outlookApp = null;
            bool outlookDisponible = true;

            try
            {
                outlookApp = new Outlook.Application();
            }
            catch
            {
                outlookDisponible = false;
                Console.WriteLine("⚠️ Outlook non disponible, les emails seront affichés dans la console uniquement.");
            }

            foreach (var req in mails)
            {
                if (string.IsNullOrWhiteSpace(req.Email)) continue;

                string colisList = string.Join(", ", req.Parcels);
                string subject = $"Colis : {colisList}";
                string phraseColis = req.Parcels.Length == 1 ?
                    "Votre colis suivant est actuellement en attente de livraison dans notre agence :" :
                    "Vos colis suivants sont actuellement en attente de livraison dans notre agence :";
                string phraseAction = req.Parcels.Length == 1 ? "le remettre en livraison" : "les remettre en livraison";

                string body = $@"
Bonjour {req.Name},

{phraseColis}
{colisList}

Afin que nous puissions {phraseAction}, pourriez-vous nous transmettre un numéro de téléphone valide, ainsi que le jour de la livraison souhaité (du lundi au vendredi).

Merci de contacter le 09 74 910 910 (numéro gratuit).

(Nous ne pouvons pas vous donner un horaire fixe de livraison, cela dépend de la tournée de notre chauffeur)

Merci de votre retour.

Tout e-mail reçu après 17h00 sera pris en charge le lendemain matin. (Hors samedi et dimanche)
";

                // Affiche dans la console
                Console.WriteLine("======================================");
                Console.WriteLine($"📧 Destinataire : {req.Email}");
                Console.WriteLine($"Objet : {subject}");
                Console.WriteLine($"Corps :\n{body}");
                Console.WriteLine("======================================");

                if (outlookDisponible)
                {
                    try
                    {
                        Outlook.MailItem mail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                        mail.To = req.Email;
                        mail.Subject = subject;
                        mail.Body = body;
                        mail.Display(); // .Send() si tu veux envoyer directement
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("❌ Erreur création mail Outlook : " + ex);
                    }
                }
            }
        }
    }
}
