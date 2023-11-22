using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using GraphEmailClient;
using System.Security;
using System;
using System.Windows.Interop;
using System.Windows;
using static System.Formats.Asn1.AsnWriter;
using Azure.Identity;


namespace EmailCalendarsClient.MailSender
{
    public class AadGraphApiDelegatedClient
    {
        private readonly HttpClient _httpClient = new HttpClient();
        private IPublicClientApplication _app;

        private static readonly string AadInstance = ConfigurationManager.AppSettings["AADInstance"];
        private static readonly string Tenant = ConfigurationManager.AppSettings["Tenant"];
        private static readonly string ClientId = ConfigurationManager.AppSettings["ClientId"];
        private static readonly string Scope = ConfigurationManager.AppSettings["Scope"];
        private static readonly string Username = ConfigurationManager.AppSettings["Username"];
        private static readonly string Password = ConfigurationManager.AppSettings["Password"];
        private static readonly string ClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

        private static readonly string Authority = string.Format(CultureInfo.InvariantCulture, AadInstance, Tenant);
        private static readonly string[] Scopes = { Scope };

        public void InitClient()
        {
            //_app = PublicClientApplicationBuilder.Create(ClientId)
            //    .WithAuthority(Authority)
            //    .WithRedirectUri("http://localhost:65419") // needed only for the system browser
            //    .Build();

            _app = PublicClientApplicationBuilder.Create(ClientId)
                  .WithAuthority(Authority)
                  .Build();

            TokenCacheHelper.EnableSerialization(_app.UserTokenCache);
        }

        public async Task<IAccount> SignIn()
        {
            try
            {
                var result = await AcquireTokenSilent();
                return result.Account;
            }
            catch (MsalUiRequiredException)
            {
                //var result = await GetATokenForGraph();
                //return await AcquireTokenInteractive().ConfigureAwait(false);
                return await AcquireTokenUsernamePassword().ConfigureAwait(false);
            }
        }

        private async Task<IAccount> AcquireTokenInteractive()
        {
            var accounts = (await _app.GetAccountsAsync()).ToList();

            var builder = _app.AcquireTokenInteractive(Scopes)
                .WithAccount(accounts.FirstOrDefault())
                .WithUseEmbeddedWebView(false)
                .WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount);

            var result = await builder.ExecuteAsync().ConfigureAwait(false);

            return result.Account;
        }

        private async Task<IAccount> AcquireTokenUsernamePassword()
        {
            var accounts = await _app.GetAccountsAsync();

            AuthenticationResult result = null;
            if (accounts.Any())
            {
                result = await _app.AcquireTokenSilent(Scopes, accounts.FirstOrDefault())
                                  .ExecuteAsync();
            }
            else
            {
                try
                {
                    var securePassword = new SecureString();
                    foreach (char c in Password)        // you should fetch the password
                        securePassword.AppendChar(c);  // keystroke by keystroke

                    result = await _app.AcquireTokenByUsernamePassword(Scopes,
                                                                     Username,
                                                                      securePassword).ExecuteAsync();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return null;
                    // See details below
                }
                //catch (MsalException)
                //{
                //    // See details below
                //}
            }

            return result.Account;
        }

        public async Task<AuthenticationResult> AcquireTokenSilent()
        {
            var accounts = await GetAccountsAsync();
            var result = await _app.AcquireTokenSilent(Scopes, accounts.FirstOrDefault())
                    .ExecuteAsync()
                    .ConfigureAwait(false);

            return result;
        }

        public async Task<IList<IAccount>> GetAccountsAsync()
        {
            var accounts = await _app.GetAccountsAsync();
            return accounts.ToList();
        }

        public async Task RemoveAccountsAsync()
        {
            IList<IAccount> accounts = await GetAccountsAsync();

            // Clears the library cache. Does not affect the browser cookies.
            while (accounts.Any())
            {
                await _app.RemoveAsync(accounts.First());
                accounts = await GetAccountsAsync();
            }
        }

        public async Task SendEmailAsync(Message message)
        {
            var result = await AcquireTokenSilent();

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var graphClient = new GraphServiceClient(_httpClient)
            {
                AuthenticationProvider = new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    await Task.FromResult<object>(null);
                })
            };

            var saveToSentItems = true;

            await graphClient.Me
                .SendMail(message, saveToSentItems)
                .Request()
                .PostAsync();
        }

        public async Task SendEmailWithSecretAsync(Message message)
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            // using Azure.Identity;
            var options = new ClientSecretCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                Tenant, ClientId, ClientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            await graphClient.Users[Username]
                .SendMail(message, true)
                .Request()
                .PostAsync();

            MessageBox.Show("Message sent successfully!", "Message", MessageBoxButton.OK);
        }

        public async Task GetInboxMessages()
        {
            List<Microsoft.Graph.QueryOption> options = new List<Microsoft.Graph.QueryOption>
            {
                    new Microsoft.Graph.QueryOption("$count","true")
            };

            var result = await AcquireTokenSilent();

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var graphClient = new GraphServiceClient(_httpClient)
            {
                AuthenticationProvider = new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    await Task.FromResult<object>(null);
                })
            };

            var subjectText = "PSS DCP DICE Certificate Signing";
            var messages = await graphClient.Me.MailFolders.Inbox.Messages.Request(options).Filter($"hasAttachments eq true and startsWith(subject,'{subjectText}')").Expand("attachments").GetAsync();

            List<Message> allMessages = new List<Message>();
            allMessages.AddRange(messages.CurrentPage);
            while (messages.NextPageRequest != null)
            {
                messages = await messages.NextPageRequest.GetAsync();
                allMessages.AddRange(messages.CurrentPage);
            }

            MessageBox.Show(string.Format("Got {0} messages with subject 'PSS DCP DICE Certificate Signing'", allMessages.Count));

        }

        public async Task GetInboxMessagesWithSecret()
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            List<Microsoft.Graph.QueryOption> qOptions = new List<Microsoft.Graph.QueryOption>
            {
                    new Microsoft.Graph.QueryOption("$count","true")
            };

            // using Azure.Identity;
            var options = new ClientSecretCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                Tenant, ClientId, ClientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            var subjectText = "PSS DCP DICE Certificate Signing";
            var messages = await graphClient.Users[Username].MailFolders.Inbox.Messages.Request(qOptions).Filter($"hasAttachments eq true and startsWith(subject,'{subjectText}')").Expand("attachments").GetAsync();

            List<Message> allMessages = new List<Message>();
            allMessages.AddRange(messages.CurrentPage);
            while (messages.NextPageRequest != null)
            {
                messages = await messages.NextPageRequest.GetAsync();
                allMessages.AddRange(messages.CurrentPage);
            }

            MessageBox.Show(string.Format("Got {0} messages with subject 'PSS DCP DICE Certificate Signing'", allMessages.Count));
        }
    }
}
