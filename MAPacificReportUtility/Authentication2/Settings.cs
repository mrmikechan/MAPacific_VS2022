using Microsoft.Extensions.Configuration;
using System;

namespace GraphAuthentication
{
    public class Settings
    {
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string TenantId { get; set; }
        public string AuthTenant { get; set; }
        public string[] GraphUserScopes { get; set; }

        public static Settings LoadSettings()
        {
            IConfiguration config = null;
            try
            {
                // Load settings
                config = new ConfigurationBuilder()
                    // appsettings.json is required
                    .AddJsonFile("appsettings.json", optional: false)
                    // appsettings.Development.json" is optional, values override appsettings.json
                    .AddJsonFile($"appsettings.Development.json", optional: true)
                    // User secrets are optional, values override both JSON files
                    //  .AddUserSecrets<Program>()
                    .Build();

                //return config.GetRequiredSection("Settings").Get<Settings>();
            }catch(Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Error");
            }
            return config.GetRequiredSection("Settings").Get<Settings>();
        }
    }
}
