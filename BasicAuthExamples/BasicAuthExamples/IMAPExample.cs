using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;


namespace BasicAuthExamples
{
    class IMAPExample
    {
        public static void FetchAllMessagesBasicAuth()
        {
            using (var client = new ImapClient(new ProtocolLogger("imap.log")))
            {
                client.Connect("outlook.office365.com", 993, SecureSocketOptions.SslOnConnect);

                client.Authenticate("username", "password");

                client.Inbox.Open(FolderAccess.ReadOnly);

                var uids = client.Inbox.Search(SearchQuery.All);

                foreach (var uid in uids)
                {
                    var message = client.Inbox.GetMessage(uid);

                    // write the message to a file
                    message.WriteTo(string.Format("{0}.eml", uid));
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


            using (var client = new ImapClient(new ProtocolLogger("imap.log")))
            {
                client.Connect("outlook.office365.com", 993, SecureSocketOptions.SslOnConnect);

                client.Authenticate(new SaslMechanismOAuth2("username", authResult.AccessToken));

                client.Inbox.Open(FolderAccess.ReadOnly);

                var uids = client.Inbox.Search(SearchQuery.All);

                foreach (var uid in uids)
                {
                    var message = client.Inbox.GetMessage(uid);

                    // write the message to a file
                    message.WriteTo(string.Format("{0}.eml", uid));
                }

                client.Disconnect(true);
            }

        }
    }
}
