using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Net;
using System.ServiceModel.Description;

namespace ActionsTodayEmailNotifications
{
    public class AuthService
    {
        private readonly string _url;

        public AuthService(string url)
        {
            if (string.IsNullOrEmpty(url)) throw new ArgumentException("Crm url is empty");

            _url = url;
        }

        public IOrganizationService Connect(string userName, string password)
        {
            ClientCredentials clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = userName;
            clientCredentials.UserName.Password = password;

            // For Dynamics 365 Customer Engagement V9.X, set Security Protocol as TLS12
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            // Get the URL from CRM, Navigate to Settings -> Customizations -> Developer Resources
            // Copy and Paste Organization Service Endpoint Address URL
            var proxy = new OrganizationServiceProxy(new Uri(_url),
             null, clientCredentials, null);

            try
            {
                proxy.Authenticate();
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null && ex.InnerException.Message == "Authentication Failure")
                {
                    throw new Exception($"Failed to connect Url: {_url}, UserName: ${userName}");
                }
                throw;
            }
            return (IOrganizationService)proxy;
        }
    }
}
