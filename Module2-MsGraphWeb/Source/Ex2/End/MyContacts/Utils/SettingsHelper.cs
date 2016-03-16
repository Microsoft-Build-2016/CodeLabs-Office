using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MyContacts.Utils
{
    public static class SettingsHelper
    {
        private static IConfigurationRoot settings;
        static SettingsHelper()
        {
            var builder = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json");
            settings = builder.Build();

            // Read from settings file
            AppId = settings["AzureAd:AppId"];
            AppPassword = settings["AzureAd:AppPassword"];
            GraphResourceId = settings["AzureAd:GraphResourceId"];
            AadInstance = settings["AzureAd:AadInstance"];
            Tenant = settings["AzureAd:Tenant"];
            Authority = AadInstance + Tenant + "/v2.0/";
        }

        public static string AppId;
        public static string AppPassword;
        public static string GraphResourceId;
        private static string AadInstance;
        public static string Tenant;
        public static string Authority;
        public static string ObjectIdentifierKey = "http://schemas.microsoft.com/identity/claims/objectidentifier";
    }
}
