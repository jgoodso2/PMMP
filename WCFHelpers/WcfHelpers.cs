using System;
using System.Globalization;
using System.ServiceModel;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel.Channels;
using System.ServiceModel.Web;
//using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Xml;
using PSLib = Microsoft.Office.Project.Server.Library;
using WCFHelpers.WsfSvcAuthentication;

namespace WCFHelpers
{
    /// <summary>
    /// 
    /// </summary>
    public class WcfHelpers
    {
        // Use the Authentication class in the Microsoft SharePoint Foundation web services, not in the PSI.
        private static Authentication authentication = null;
        private static AuthenticationMode mode = AuthenticationMode.Windows;
        private static CookieContainer cookieContainer = null;
        private static String impersonationContextString = String.Empty; 

        public static CookieContainer CookieContainer
        {
            get
            { return cookieContainer; }
        }

        public static AuthenticationMode AuthenticationMode
        {
            get
            { return mode; }
        }

        public static string HeaderXformsKey
        {
            get
            { return "X-FORMS_BASED_AUTH_ACCEPTED"; }
        }

        public static string HeaderXformsValue
        {
            get
            { return "f"; }
        }

        public static string ImpersonationContextString
        {
            get
            { return impersonationContextString; }
        }

        private static bool customCertificateValidation(object sender, X509Certificate cert, 
                                                        X509Chain chain, SslPolicyErrors error)
        {
            return true;
        }

        // Log on by using the Authentication web service in Microsoft SharePoint Foundation.
        public static bool LogonWithMsf(string username, string password, Uri pwaURI)
        {
            if (authentication == null)
            {
                string rootUrl = pwaURI.Scheme + Uri.SchemeDelimiter + pwaURI.Host + ":" + pwaURI.Port;

                authentication = new Authentication();
                authentication.Url = rootUrl + "/_vti_bin/Authentication.asmx";
            }

            authentication.CookieContainer = new System.Net.CookieContainer();
            LoginResult result = authentication.Login(username, password);

            mode = authentication.Mode();
            cookieContainer = authentication.CookieContainer;

            return (result.ErrorCode == LoginErrorCode.NoError);
        }

        // Get the Project Web App site id.
        public Guid GetPwaSiteId(Uri pwaUri)
        {
            Guid siteId = Guid.Empty;
            string pwaUrl = pwaUri.OriginalString;

            if (!pwaUrl.Trim().EndsWith("/"))
                pwaUrl += "/";

            WebSvcSiteData.SiteData siteData = new WebSvcSiteData.SiteData();
            siteData.Url = pwaUrl + "_vti_bin/SiteData.asmx";
            siteData.Credentials = CredentialCache.DefaultCredentials;

            string pwaIdString = string.Empty;
            siteData.GetSiteUrl(pwaUrl, out pwaUrl, out pwaIdString);

            siteId = new Guid(pwaIdString);
            return siteId;
        }

        public static void UseWindowsAuthOnMultiAuthHeader()
        {
            WebOperationContext.Current.OutgoingRequest.Headers.Remove(HeaderXformsKey);
            WebOperationContext.Current.OutgoingRequest.Headers.Add(HeaderXformsKey, HeaderXformsValue);
        }

        public static void UseCookieHeader()
        {
            if (cookieContainer != null)
            {
                var cookieString = cookieContainer.GetCookieHeader(new Uri(authentication.Url));

                WebOperationContext.Current.OutgoingRequest.Headers.Remove("Cookie");
                WebOperationContext.Current.OutgoingRequest.Headers.Add("Cookie", cookieString);
            }
        }

        public static void UseCorrectHeaders(bool isImpersonated)
        {
            if (isImpersonated)
            {
                // Use WebOperationContext in the HTTP channel, not the OperationContext.
                WebOperationContext.Current.OutgoingRequest.Headers.Remove("PjAuth");
                WebOperationContext.Current.OutgoingRequest.Headers.Add("PjAuth", impersonationContextString);
            }

            if (mode == AuthenticationMode.Windows)
            {
                UseWindowsAuthOnMultiAuthHeader();
            }

            UseCookieHeader();
        }

        // Set the impersonation context for calls to the PSI on behalf of the impersonated user.
        public static void SetImpersonationContext(bool isWindowsUser, String userNTAccount,
                                                   Guid userGuid, Guid trackingGuid, Guid siteId,
                                                   CultureInfo languageCulture, CultureInfo localeCulture)
        {
            impersonationContextString = GetImpersonationContext(isWindowsUser, userNTAccount, userGuid,
                                                                  trackingGuid, siteId, 
                                                                  languageCulture, localeCulture);
        }

        public static void SetImpersonation(bool isWindowsUser, string impersonatedUser, Guid resourceGuid)
        {
            Guid trackingGuid = Guid.NewGuid();
            Guid siteId = Guid.Empty;           // Project Web App site ID.
            CultureInfo languageCulture = null; // The language culture is not used.
            CultureInfo localeCulture = null;   // The locale culture is not used.

            WcfHelpers.SetImpersonationContext(isWindowsUser, impersonatedUser, resourceGuid, trackingGuid, siteId,
                                               languageCulture, localeCulture);
            
        }

        // Get the impersonation context.
        private static String GetImpersonationContext(bool isWindowsUser, String userNTAccount,
                                                      Guid userGuid, Guid trackingGuid, Guid siteId,
                                                      CultureInfo languageCulture, CultureInfo localeCulture)
        {
            PSLib.PSContextInfo contextInfo = new PSLib.PSContextInfo(isWindowsUser, userNTAccount, userGuid, 
                                                                      trackingGuid, siteId, 
                                                                      languageCulture, localeCulture);
            String contextInfoString = PSLib.PSContextInfo.SerializeToString(contextInfo);
            return contextInfoString;
        }

        // Clear the impersonation context.
        public static void ClearImpersonationContext()
        {
            impersonationContextString = string.Empty;
        }

        /// <summary>
        /// Extract a PSClientError object from the ServiceModel.FaultException,
        /// for use in output of the GetPSClientError stack of errors.
        /// </summary>
        /// <param name="e"></param>
        /// <param name="errOut">Shows that FaultException has more information 
        /// about the errors than PSClientError has. FaultException can also contain 
        /// other types of errors, such as failure to connect to the server.</param>
        /// <returns>PSClientError object, for enumerating errors.</returns>
        public static PSLib.PSClientError GetPSClientError(FaultException e, out string errOut)
        {
            const string PREFIX = "GetPSClientError() returns null: ";
            errOut = string.Empty;
            PSLib.PSClientError psClientError = null;

            if (e == null)
            {
                errOut = PREFIX + "Null parameter (FaultException e) passed in.";
                psClientError = null;
            }
            else
            {
                // Get a ServiceModel.MessageFault object.
                var messageFault = e.CreateMessageFault();

                if (messageFault.HasDetail)
                {
                    using (var xmlReader = messageFault.GetReaderAtDetailContents())
                    {
                        var xml = new XmlDocument();
                        xml.Load(xmlReader);

                        var serverExecutionFault = xml["ServerExecutionFault"];
                        if (serverExecutionFault != null)
                        {
                            var exceptionDetails = serverExecutionFault["ExceptionDetails"];
                            if (exceptionDetails != null)
                            {
                                try
                                {
                                    errOut = exceptionDetails.InnerXml + "\r\n";
                                    psClientError =
                                        new PSLib.PSClientError(exceptionDetails.InnerXml);
                                }
                                catch (InvalidOperationException ex)
                                {
                                    errOut = PREFIX + "Unable to convert fault exception info ";
                                    errOut += "a valid Project Server error message. Message: \n\t";
                                    errOut += ex.Message;
                                    psClientError = null;
                                }
                            }
                            else
                            {
                                errOut = PREFIX + "The FaultException e is a ServerExecutionFault, "
                                    + "but does not have ExceptionDetails.";
                            }
                        }
                        else
                        {
                            errOut = PREFIX + "The FaultException e is not a ServerExecutionFault.";
                        }
                    }
                }
                else // No detail in the MessageFault.
                {
                    errOut = PREFIX + "The FaultException e does not have any detail.";
                }
            }
            errOut += "\r\n" + e.ToString() + "\r\n";
            return psClientError;
        }
    }

}
