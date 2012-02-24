using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Configuration;
using log4net;

using System.Net.Security;
using System.Security.Cryptography.X509Certificates;


namespace ConsoleApplication1
{
    class ExchangeConnection
    {
        protected static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        ExchangeService service;


        public ExchangeConnection()
        {
            log.Debug("Attempting to connect to exchange");
            SetCertificatePolicy();
        //    connectToExchangeUsingImpersonation();
            connectToExchangeOnPremise();
        }

        public ExchangeService GetExchangeService()
        {
            return service;
        }

        private static void SetCertificatePolicy()
        {
            log.Debug("Setting certificate policy");
            ServicePointManager.ServerCertificateValidationCallback
                       += RemoteCertificateValidate;
        }
        private static bool RemoteCertificateValidate(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors error)
        {
            // trust any certificate!!!
            log.Debug("Warning, trust any certificate");
            return true;
        }
           
        private void connectToExchangeOnPremise()
        {
            log.Debug("Setting credentials and EWS url");
            service = new ExchangeService(ExchangeVersion.Exchange2010);
            service.Credentials = new WebCredentials("dtest1", "Password1", "mcdev.local");
            service.Url = service.Url = new Uri(ConfigurationManager.AppSettings["EWSurl"]);
       //     service.TraceEnabled = true;
       //     service.ImpersonatedUserId = ""; TODO
        //    ImpersonatedUserId id = new ImpersonatedUserId();
        }

        private void connectToExchangeUsingImpersonation()
        {
            log.Debug("Setting credentials (using impersonation) and EWS url");
            service = new ExchangeService(ExchangeVersion.Exchange2010);
            service.Credentials = new WebCredentials("daviturndc", "Password1", "mcdev.local");
            service.Url = service.Url = new Uri(ConfigurationManager.AppSettings["EWSurl"]);

            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "dtest1@mandc-test.com");
            //     service.TraceEnabled = true;
        }
    }
}


