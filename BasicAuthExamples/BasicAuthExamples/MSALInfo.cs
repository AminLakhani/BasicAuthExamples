using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using Microsoft.Identity.Client;
using System.Threading.Tasks;
using System.Security;
using System.Net;

namespace BasicAuthExamples
{
    class MSALInfo
    {
        /*
         *  Documenation for different auth flows can be found here -> https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-authentication-flows
         *  
         *  Below are just code snippets for how to implement a few different auth scenarios along with what arguments each method takes.
         */

        public static async Task Snippets() 
        {

            // The permission scope required for EWS access
            var scopes = new string[] { "https://outlook.office365.com/EWS.AccessAsUser.All" };
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


            /*
             * https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/Acquiring-tokens-interactively
             * 
             * Interactive request to acquire a token for the specified scopes. The interactive window will be parented to the specified window. The user will be required to select an account.
             * 
             * AcquireTokenInteractive has only one mandatory parameter scopes, which contains an enumeration of strings which define the scopes for which a token is required. 
             * If the token is for the Microsoft Graph, the required scopes can be found in api reference of each Microsoft graph API in the section named "Permissions"
             * 
             */
            var authResult = await pca.AcquireTokenInteractive(scopes).ExecuteAsync();

            /*
             * https://aka.ms/msal-net-iwa
             * 
             * Non-interactive request to acquire a security token for the signed-in user in Windows, via Integrated Windows Authentication. 
             * The account used in this overrides is pulled from the operating system as the current user principal name.
             * 
             * Federated users only, i.e. those created in an Active Directory and backed by Azure Active Directory. 
             * Users created directly in AAD, without AD backing - managed users - cannot use this auth flow. This limitation does not affect the Username/Password flow.
             * 
             * Takes 1 argument which is the scopes you're requesting. 
             */
            var authResult1 = await pca.AcquireTokenByIntegratedWindowsAuth(scopes).ExecuteAsync();

            /*
             * https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/Username-Password-Authentication
             * 
             * In your desktop application, you can use the Username/Password flow to acquire a token silently. No UI is required when using the application.
             * 
             * This flow is not recommended because your application asking a user for their password is not secure.
             * The preferred flow for acquiring a token silently on Windows domain joined machines is Integrated Windows Authentication.
             * 
             * 
             */
            var authResult2 = await pca.AcquireTokenByUsernamePassword(scopes, "USERNAME", passwordSecure).ExecuteAsync();

            /* 
             * https://aka.ms/msal-device-code-flow
             * 
             * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code
             * https://docs.microsoft.com/en-us/dotnet/api/microsoft.identity.client.publicclientapplication.acquiretokenwithdevicecode?view=azure-dotnet
             * 
             * Acquires a security token on a device without a web browser, by letting the user authenticate on another device.
             * 
             * Takes 1 argument which is the scopes you're requesting. 
            */
            var authResult3 = await pca.AcquireTokenWithDeviceCode(scopes, deviceCodeResult =>
            {
                Console.WriteLine(deviceCodeResult.Message);
                return Task.FromResult(0);
            }).ExecuteAsync();

        }
    }
}
