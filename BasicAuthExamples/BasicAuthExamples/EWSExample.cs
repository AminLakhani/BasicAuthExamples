using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System.Threading.Tasks;
using System.Security;
using System.Net;

//https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols

namespace BasicAuthExamples
{
    class EWSExample
    {
        public static void SendEmailBasicAuth()
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            service.Credentials = new WebCredentials("user1@contoso.com", "password");
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            service.AutodiscoverUrl("user1@contoso.com", RedirectionUrlValidationCallback);

            // Make an EWS call
            EmailMessage email = new EmailMessage(service);
            email.ToRecipients.Add("user1@contoso.com");
            email.Subject = "HelloWorld";
            email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
            email.Send();
        }

        public static async Task<string> SendEmailMSALAsync()
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



            SecureString passwordSecure = new NetworkCredential("", "myPass").SecurePassword;


            // Make the interactive token request
            var authResult = await pca.AcquireTokenByUsernamePassword(ewsScopes, "USERNAME", passwordSecure).ExecuteAsync();

            // Configure the ExchangeService with the access token
            var ewsClient = new ExchangeService();
            ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);

            // Make an EWS call
            EmailMessage email = new EmailMessage(ewsClient);
            email.ToRecipients.Add("user1@contoso.com");
            email.Subject = "HelloWorld";
            email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
            email.Send();

            return "sucess";
        }


        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);

            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
