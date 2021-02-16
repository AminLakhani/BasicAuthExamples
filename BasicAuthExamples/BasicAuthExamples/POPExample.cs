using MailKit;
using MailKit.Net.Pop3;
using MailKit.Security;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace BasicAuthExamples
{
    class POPExample
    {
        public static void FetchAllMessagesBasicAuth()
        {
            using (var client = new Pop3Client(new ProtocolLogger("pop3.log")))
            {
                client.Connect("outlook.office365.com", 993, SecureSocketOptions.SslOnConnect);

                client.Authenticate("username", "password");

                for (int i = 0; i < client.Count; i++)
                {
                    var message = client.GetMessage(i);

                    // write the message to a file
                    message.WriteTo(string.Format("{0}.msg", i));

                    // mark the message for deletion
                    client.DeleteMessage(i);
                }

                client.Disconnect(true);
            }
        }

        public static async Task FetchAllMessagesOAuthAsync()
        {
            // The permission scope required for EWS access
            var ewsScopes = new string[] { "https://outlook.office365.com/EWS.AccessAsUser.All" };
            var ClientId = "";
            var TenantId = "";


            // Configure the MSAL client to get tokens
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = ClientId,
                TenantId = TenantId
            };


            var pca = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions).Build();


            // Make the interactive token request
            var authResult = await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync();


            using (var client = new Pop3Client(new ProtocolLogger("pop3.log")))
            {
                client.Connect("outlook.office365.com", 993, SecureSocketOptions.SslOnConnect);

                client.Authenticate(new SaslMechanismOAuth2("username", authResult.AccessToken));

                for (int i = 0; i < client.Count; i++)
                {
                    var message = client.GetMessage(i);

                    // write the message to a file
                    message.WriteTo(string.Format("{0}.msg", i));

                    // mark the message for deletion
                    client.DeleteMessage(i);
                }

                client.Disconnect(true);
            }

        }
    }
}
