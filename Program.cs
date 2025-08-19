using System.Text.Json;
using System.Threading;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Hosting;
using Outlook = Microsoft.Office.Interop.Outlook;

record MailRequest(string Email, string? Name, string[] Parcels);

internal static class Program
{
    [STAThread] // point d’entrée STA, mais on assurera STA explicite sur l’envoi
    public static async Task Main(string[] args)
    {
        var builder = Host.CreateApplicationBuilder(args);

        // Kestrel en localhost:5001 (aucun droit admin requis)
        builder.WebHost.ConfigureKestrel(k =>
        {
            k.ListenLocalhost(5001);
        });

        var app = builder.Build();

        // CORS minimal + préflight
        app.Use(async (ctx, next) =>
        {
            ctx.Response.Headers["Access-Control-Allow-Origin"] = "*";
            ctx.Response.Headers["Access-Control-Allow-Headers"] = "content-type";
            ctx.Response.Headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS";
            if (ctx.Request.Method == "OPTIONS") { ctx.Response.StatusCode = 200; return; }
            await next();
        });

        app.MapGet("/ping", () => Results.Ok(new { ok = true }));

        app.MapPost("/send-mails", async (HttpContext ctx) =>
        {
            try
            {
                var mails = await JsonSerializer.DeserializeAsync<List<MailRequest>>(
                    ctx.Request.Body,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true }
                );

                if (mails is null || mails.Count == 0)
                    return Results.BadRequest(new { error = "Payload vide ou invalide." });

                // Outlook Interop doit s’exécuter sur un thread STA
                await RunSTAAsync(() => GenerateMails(mails));

                return Results.Ok(new { status = "ok", count = mails.Count });
            }
            catch (Exception ex)
            {
                return Results.Problem(ex.Message);
            }
        });

        await app.RunAsync(); // process Windows sans fenêtre
    }

    private static Task RunSTAAsync(Action action)
    {
        var tcs = new TaskCompletionSource<object?>();
        var th = new Thread(() =>
        {
            try { action(); tcs.SetResult(null); }
            catch (Exception ex) { tcs.SetException(ex); }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
        return tcs.Task;
    }

    private static void GenerateMails(IEnumerable<MailRequest> mails)
    {
        Outlook.Application outlookApp;
        try
        {
            outlookApp = new Outlook.Application();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Outlook n’est pas disponible sur cette machine.", ex);
        }

        foreach (var req in mails)
        {
            if (string.IsNullOrWhiteSpace(req.Email) || req.Parcels is null || req.Parcels.Length == 0)
                continue;

            string colisList = string.Join(", ", req.Parcels);
            int nbColis = req.Parcels.Length;

            string phraseColis = nbColis == 1
                ? "Votre colis suivant est actuellement en attente de livraison dans notre agence :"
                : "Vos colis suivants sont actuellement en attente de livraison dans notre agence :";

            string phraseAction = nbColis == 1 ? "le remettre en livraison" : "les remettre en livraison";
            string subject = "Colis : " + colisList;

            string messageHTML =
                "<p>Bonjour,</p>" +
                $"<p>{phraseColis}</p>" +
                $"<p><b>{colisList}</b></p>" +
                $"<p>Afin que nous puissions {phraseAction}, pourriez-vous nous transmettre un numéro de téléphone valide, " +
                "ainsi que le jour de la livraison souhaité (du <b>lundi au vendredi</b>).</p>" +
                "<p><b>Merci de contacter le 09 74 910 910</b> (numéro gratuit).</p>" +
                "<p><i><u>(Nous ne pouvons pas vous donner un horaire fixe de livraison, cela dépend de la tournée de notre chauffeur)</u></i></p>" +
                "<p>Merci de votre retour.</p>" +
                "<p><i>Tout e-mail reçu après <b>17h00</b> sera pris en charge le lendemain matin. (Hors samedi et dimanche)</i></p>";

            string signatureHTML =
                "<div style='font-family:Calibri,Arial,Helvetica,sans-serif;font-size:12pt;color:#000'>" +
                "<br>Cordialement,<br><br>" +
                "<b style='font-size:14pt; color:#242424;'>Service Livraison Pôle réclamation</b><br>" +
                "<small style='font-family:Verdana;'>GLS Roquemaure - " +
                "<a href='https://www.google.com/maps/search/11+avenue+de+l''Aspre+-+30150+Roquemaure'>11 avenue de l'Aspre - 30150 Roquemaure</a></small><br><br>" +
                "<table><tr>" +
                "<td><a href='https://fr.linkedin.com/company/gls-france'><img src='https://www.gls-france.com/signature/in.png' width='18' height='18'></a> LinkedIn</td>" +
                "<td><a href='https://www.youtube.com/channel/UCp-IPDFX5NGaLgxwdwAjcrA/featured'><img src='https://www.gls-france.com/signature/yt.png' width='18' height='18'></a> YouTube</td>" +
                "<td><a href='https://www.facebook.com/GLSFrance'><img src='https://www.gls-france.com/signature/fb.png' width='18' height='18'></a> Facebook</td>" +
                "<td><a href='https://www.instagram.com/GLSFrance/'><img src='https://www.gls-france.com/signature/ig.png' width='18' height='18'></a> Instagram</td>" +
                "<td><a href='https://www.tiktok.com/@glsfrance'><img src='https://www.gls-france.com/signature/tiktok.png' width='18' height='18'></a> TikTok</td>" +
                "</tr></table><br>" +
                "<a href='https://gls-group.eu/FR/fr/home'><img src='https://www.gls-france.com/signature/logogls.png' width='173' height='60'></a><br>" +
                "<b>@</b> <a href='mailto:destinataire.fr0084@gls-france.com'>destinataire.fr0084@gls-france.com</a><br>" +
                "<b>W</b> <a href='http://www.gls-group.com/FR'>www.gls-group.com/FR</a><br><br>" +
                "<a href='https://gls-group.com/FR/fr/envoyer-colis/solutions-pros/flexdeliveryservice/'>" +
                "<img src='https://www.gls-france.com/signature/banner.png' width='600' height='150'></a>" +
                "</div>";

            var mail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            mail.To = req.Email;
            mail.Subject = subject;
            mail.HTMLBody = messageHTML + signatureHTML;
            mail.SentOnBehalfOfName = "destinataire.fr0084@gls-france.com";
            mail.Display(); // .Send() pour envoyer direct
        }
    }
}