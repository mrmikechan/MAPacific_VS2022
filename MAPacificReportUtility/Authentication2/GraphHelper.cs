using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Azure.Identity;



namespace GraphAuthentication
{
    //TODO: Add another method for TokenAsync

    /// <summary>
    /// GraphHelper is the class used to perform authentication and various other Graph related action with the Graph API.
    /// Microsoft.Graph package is obtained through the Nuget Package Solution manager.  Here are the list of packages used. Note that
    /// a lot of these files are dependency files needed for the origional Microsoft.Graph package.
    /// 
    /// Azure.Core.1.24.0
    /// Azure.Identity.1.6.0
    /// Microsoft.Bcl.AsyncInterfaces.6.0.0
    /// Microsoft.Extensions.Configuration.6.0.1
    /// Microsoft.Extensions.Configuration.Abstractions.6.0.0
    /// Microsoft.Extensions.Configuration.Binder.6.0.0
    /// Microsoft.Extensions.Configuration.FileExtensions.6.0.0
    /// Microsoft.Extensions.Configuration.Json.6.0.0
    /// Microsoft.Extensions.FileProviders.Abstractions.6.0.0
    /// Microsoft.Extensions.FileProviders.Physical.6.0.0
    /// Microsoft.Extensions.FileSystemGlobbing.6.0.0
    /// Microsoft.Extensions.Primitives.6.0.0
    /// Microsoft.Graph.4.31.0
    /// Microsoft.Graph.Core.2.0.9
    /// Microsoft.Identity.Client.4.44.0
    /// Microsoft.Identity.Client.Extensions.Msal.2.19.3
    /// Microsoft.IdentityModel.Abstractions.6.19.0
    /// Microsoft.IdentityModel.JsonWebTokens.6.19.0
    /// Microsoft.IdentityModel.Logging.6.19.0
    /// Microsoft.IdentityModel.Protocols.6.19.0
    /// Microsoft.IdentityModel.Protocols.OpenIdConnect.6.19.0
    /// Microsoft.IdentityModel.Tokens.6.19.0
    /// System.Buffers.4.5.1
    /// System.Diagnostics.DiagnosticSource.4.7.1
    /// System.IdentityModel.Tokens.Jwt.6.19.0
    /// System.IO.4.3.0
    /// System.Memory.4.5.4
    /// System.Memory.Data.1.0.2
    /// System.Net.Http.4.3.4
    /// System.Net.Http.WinHttpHandler.6.0.1
    /// System.Numerics.Vectors.4.5.0
    /// System.Runtime.4.3.0
    /// System.Runtime.CompilerServices.Unsafe.6.0.0
    /// System.Runtime.InteropServices.RuntimeInformation.4.3.0
    /// System.Security.Cryptography.Algorithms.4.3.0
    /// System.Security.Cryptography.Algorithms.4.3.1
    /// System.Security.Cryptography.Encoding.4.3.0
    /// System.Security.Cryptography.Primitives.4.3.0
    /// System.Security.Cryptography.ProtectedData.4.7.0
    /// System.Security.Cryptography.X509Certificates.4.3.0
    /// System.Text.Encodings.Web.6.0.0
    /// System.Text.Json.6.0.4
    /// System.Threading.Tasks.Extensions.4.5.4
    /// System.ValueTuple.4.5.0
    /// </summary>
    public class GraphHelper
    {
        private static Microsoft.Graph.GraphServiceClient _userClient;
        private Settings _settings;
        private string _displayName, _email;

        //user auth token credentials
        private static DeviceCodeCredential _deviceCodeCredential;
        public GraphHelper()
        {
            _settings = Settings.LoadSettings();
            _displayName = null;
            _email = null;
            _userClient = null;
        }

        public static void InitializeGraphForUserAuth(Settings settings,
            Func<DeviceCodeInfo, System.Threading.CancellationToken, Task> deviceCodePrompt)
        {
            _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
            settings.AuthTenant, settings.ClientId);

            _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
        }

        public static Task<User> GetUserAsync()
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            return _userClient.Me
                .Request()
                .Select(u => new
                {
                    // Only request specific properties
                    u.DisplayName,
                    u.Mail,
                    u.UserPrincipalName
                })
                .GetAsync();
        }

        public static async Task<IUserContactsCollectionPage> GetContactListAsync()
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            //   return await _userClient.Me.ContactFolders["{contactFolder-id}"].Contacts
            //       .Request()
            //       .GetAsync();

            return await _userClient.Me.Contacts.Request().GetAsync();

        }

        public static async Task<IGraphServiceGroupsCollectionPage> GetContactGroupsAsync()
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            return await _userClient.Groups.Request().GetAsync();
        }

        public static async Task<IGroupMembersCollectionWithReferencesPage> GetMembersofGroup(string id)
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            return await _userClient.Groups[id].Members.Request().GetAsync();
        }

        //------------------------------
        public void InitializeGraph()
        {
            InitializeGraphForUserAuth(_settings, (info, cancel) =>
            {
                if (info.Message != null)
                {
                    System.Threading.Thread t = new System.Threading.Thread(() =>
                    {
                        AuthenticateForm myForm = new AuthenticateForm();
                        myForm.NavigatePage(info.Message);
                        myForm.ShowDialog();

                    });
                    t.SetApartmentState(System.Threading.ApartmentState.STA);
                    t.Start();
                }
                //      System.Diagnostics.Process.Start("https://microsoft.com/devicelogin");
                //   System.Windows.Forms.MessageBox.Show(info.Message, "Open Browser for Authentication", System.Windows.Forms.MessageBoxButtons.OK);
                return Task.FromResult(0);
            });
        }

        public async Task GreetUserAsync()
        {
            try
            {
                var user = await GetUserAsync();
                Console.WriteLine($"Hello, {user?.DisplayName}!");
                _displayName = user?.DisplayName;
                // For Work/school accounts, email is in Mail property
                // Personal accounts, email is in UserPrincipalName
                _email = user?.Mail ?? user?.UserPrincipalName ?? "";
                Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting user: {ex.Message}");
            }
        }

        public IUserContactsCollectionPage GetContactsAsync()
        {
            var contactPage = GraphHelper.GetContactListAsync();
            return (IUserContactsCollectionPage)contactPage;
        }

        public IGraphServiceGroupsCollectionPage GetContactGroups()
        {
            var groupsPage = GraphHelper.GetContactGroupsAsync();
            return (IGraphServiceGroupsCollectionPage)groupsPage;
        }

        public IGroupMembersCollectionWithReferencesPage GetGroupMembers(string id)
        {
            var members = GraphHelper.GetMembersofGroup(id);
            return (IGroupMembersCollectionWithReferencesPage)members;
        }

        public GraphServiceClient getClient()
        {
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            return _userClient;
        }

        public string getDisplayName()
        {
            return _displayName;
        }

        public string getEmail()
        {
            return _email;
        }
    }
}
