using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using SvcProject;
using SvcLookupTable;
using SvcCustomFields;
using WCFHelpers;
using System.Collections;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Configuration;
using System.ServiceModel;
using PSLib = Microsoft.Office.Project.Server.Library;
using System.Security.Principal;
using System.Xml;
using System.Data.SqlClient;
using System.Web.Services.Protocols;

namespace Repository
{
    public struct FiscalMonth
    {
        public DateTime To { get; set; }
        public DateTime From { get; set; }
    }

    public class DataRepository
    {
        // WCF endpoint names in app.config.
        private const string ENDPOINT_ADMIN = "basicHttp_Admin";
        private const string ENDPOINT_Q = "basicHttp_QueueSystem";
        private const string ENDPOINT_RES = "basicHttp_Resource";
        private const string ENDPOINT_PROJ = "basicHttp_Project";
        private const string ENDPOINT_LUT = "basicHttp_LookupTable";
        private const string ENDPOINT_CF = "basicHttp_CustomFields";
        private const string ENDPOINT_CAL = "basicHttp_Calendar";
        private const string ENDPOINT_AR = "basicHttp_Archive";
        private const string ENDPOINT_PWA = "basicHttp_PWA";
        private const int NO_QUEUE_MESSAGE = -1;

        public static SvcAdmin.AdminClient adminClient;
        public static SvcQueueSystem.QueueSystemClient queueSystemClient;
        public static SvcResource.ResourceClient resourceClient;
        public static SvcProject.ProjectClient projectClient;
        public static SvcLookupTable.LookupTableClient lookupTableClient;
        public static SvcCustomFields.CustomFieldsClient customFieldsClient;
        public static SvcCalendar.CalendarClient calendarClient;
        public static SvcArchive.ArchiveClient archiveClient;
        public static SvcStatusing.StatusingClient pwaClient;
        public static bool isImpersonated = false;
        private static SvcLoginWindows.LoginWindows loginWindows;    // Use for logon of different Windows user.

        //public static SvcLoginForms.LoginForms loginForms = new SvcLoginForms.LoginForms();
        //public static SvcLoginWindows.LoginWindows loginWindows = new SvcLoginWindows.LoginWindows();

        public static string projectServerUrl = "";
        public static string userName = "";
        public static string password = "";
        public static bool isWindowsAuth = true;

        public static bool useDefaultWindowsCredentials = true; // Currently must be true for Windows authentication in ProjTool.
        public static int windowsPort = 80;
        public static int formsPort = 81;
        public static bool waitForQueue = true;
        public static bool waitForIndividualQueue = false;
        public static bool autoLogin = false;

        public static Guid pwaSiteId = Guid.Empty;
        public static Guid jobGuid;
        public static Guid projectGuid = new Guid();

        public static int loginStatus = 0;
        public static string impersonatedUserName = "";

        public enum queryType
        {
            GroupAndName,
            GroupAndDisplayName,
            UserNameAndName,
            UserNameandDisplayName
        }

        public static ArrayList adUsers = new ArrayList();
        public static MySettings mySettings = new MySettings();


        public void LoadProjects(string url)
        {
            ClearImpersonation();

            if (P14Login(url))
            {
                ReadProjectsList();
            }
        }

        public static SvcProject.ProjectDataSet ReadProjectsList()  //called by configuration screen
        {
            SvcProject.ProjectDataSet projectList = new SvcProject.ProjectDataSet();
            try
            {


                using (OperationContextScope scope = new OperationContextScope(projectClient.InnerChannel))
                {
                    WcfHelpers.UseCorrectHeaders(isImpersonated);
                    SvcStatusing.StatusingDataSet dataSet = pwaClient.ReadStatus(Guid.Empty, DateTime.MinValue, DateTime.MaxValue);
                    // Get projects of type normal, templates, proposals, master, and inserted.
                    string projectName = string.Empty;

                    projectList.Merge(projectClient.ReadProjectStatus(Guid.Empty, SvcProject.DataStoreEnum.PublishedStore,
                        projectName, (int)PSLib.Project.ProjectType.Project));

                    projectList.Merge(projectClient.ReadProjectStatus(Guid.Empty, SvcProject.DataStoreEnum.PublishedStore,
                        projectName, (int)PSLib.Project.ProjectType.InsertedProject));
                    projectList.Merge(projectClient.ReadProjectStatus(Guid.Empty, SvcProject.DataStoreEnum.PublishedStore,
                        projectName, (int)PSLib.Project.ProjectType.MasterProject));
                }

            }
            catch (Exception ex)
            {

            }
            finally
            {

            }
            return projectList;
        }

        public static bool P14Login(string projectserverURL)
        {
            bool endPointError = false;
            bool result = false;

            try
            {
                projectServerUrl = projectserverURL.Trim();

                if (!projectServerUrl.EndsWith("/"))
                {
                    projectServerUrl = projectServerUrl + "/";
                }
                String baseUrl = projectServerUrl;

                // Configure the WCF endpoints of PSI services used in ProjTool, before logging on.
                if (mySettings.UseAppConfig)
                {
                    endPointError = !ConfigClientEndpoints();
                }
                else
                {
                    endPointError = !SetClientEndpointsProg(baseUrl);
                }

                if (endPointError) return false;

                // NOTE: Windows logon with the default Windows credentials, Forms logon, and impersonation work in ProjTool. 
                // Windows logon without the default Windows credentials does not currently work.
                if (!isImpersonated)
                {
                    if (isWindowsAuth)
                    {
                        if (useDefaultWindowsCredentials)
                        {
                            result = true;
                        }
                        else
                        {
                            String[] splits = userName.Split('\\');

                            if (splits.Length != 2)
                            {
                                String errorMessage = "User name must be in the format domainname\\accountname";
                                result = false;
                            }
                            else
                            {
                                // Order of strings returned by String.Split is not deterministic
                                // Hence we cannot use splits[0] and splits[1] to obtain domain name and user name

                                int positionOfBackslash = userName.IndexOf('\\');
                                String windowsDomainName = userName.Substring(0, positionOfBackslash);
                                String windowsUserName = userName.Substring(positionOfBackslash + 1);

                                loginWindows = new SvcLoginWindows.LoginWindows();
                                loginWindows.Url = baseUrl + "_vti_bin/PSI/LoginWindows.asmx";
                                loginWindows.Credentials = new NetworkCredential(windowsUserName, password, windowsDomainName);

                                result = loginWindows.Login();
                            }
                        }
                    }
                    else
                    {
                        // Forms authentication requires the Authentication web service in Microsoft SharePoint Foundation.
                        result = WcfHelpers.LogonWithMsf(userName, password, new Uri(baseUrl));
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static void ClearImpersonation()
        {
            WcfHelpers.ClearImpersonationContext();
            isImpersonated = false;
        }

        public static string CatchFaultException(FaultException faultEx)
        {
            string errAttributeName;
            string errAttribute;
            string errOut;
            string errMess = "".PadRight(30, '=') + "\r\n"
                + "Error details: " + "\r\n";

            PSLib.PSClientError error = WcfHelpers.GetPSClientError(faultEx, out errOut);
            errMess += errOut;

            PSLib.PSErrorInfo[] errors = error.GetAllErrors();
            PSLib.PSErrorInfo thisError;

            for (int i = 0; i < errors.Length; i++)
            {
                thisError = errors[i];
                errMess += "\r\n".PadRight(30, '=') + "\r\nPSClientError output:\r\n";
                errMess += thisError.ErrId.ToString() + "\n";

                for (int j = 0; j < thisError.ErrorAttributes.Length; j++)
                {
                    errAttributeName = thisError.ErrorAttributeNames()[j];
                    errAttribute = thisError.ErrorAttributes[j];
                    errMess += "\r\n\t" + errAttributeName
                        + ": " + errAttribute;
                }
            }
            return errMess;
        }

        #region Configure WCF client and impersonation

        // Configure the PSI client endpoints by using the settings in app.config.
        public static bool ConfigClientEndpoints()
        {
            bool result = true;

            string[] endpoints = { ENDPOINT_ADMIN, ENDPOINT_Q, ENDPOINT_RES, ENDPOINT_PROJ, 
                                   ENDPOINT_LUT, ENDPOINT_CF, ENDPOINT_CAL, ENDPOINT_AR, 
                                   ENDPOINT_PWA };
            try
            {
                foreach (string endPt in endpoints)
                {
                    switch (endPt)
                    {
                        case ENDPOINT_ADMIN:
                            adminClient = new SvcAdmin.AdminClient(endPt);
                            break;
                        case ENDPOINT_PROJ:
                            projectClient = new SvcProject.ProjectClient(endPt);
                            break;
                        case ENDPOINT_Q:
                            queueSystemClient = new SvcQueueSystem.QueueSystemClient(endPt);
                            break;
                        case ENDPOINT_RES:
                            resourceClient = new SvcResource.ResourceClient(endPt);
                            break;
                        case ENDPOINT_LUT:
                            lookupTableClient = new SvcLookupTable.LookupTableClient(endPt);
                            break;
                        case ENDPOINT_CF:
                            customFieldsClient = new SvcCustomFields.CustomFieldsClient(endPt);
                            break;
                        case ENDPOINT_CAL:
                            calendarClient = new SvcCalendar.CalendarClient(endPt);
                            break;
                        case ENDPOINT_AR:
                            archiveClient = new SvcArchive.ArchiveClient(endPt);
                            break;
                        case ENDPOINT_PWA:
                            pwaClient = new SvcStatusing.StatusingClient(endPt);
                            break;
                        default:
                            result = false;
                            Console.WriteLine("Invalid endpoint: {0}", endPt);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                result = false;
            }
            return result;
        }

        // Set the PSI client endpoints programmatically; don't use app.config.
        private static bool SetClientEndpointsProg(string pwaUrl)
        {
            const int MAXSIZE = int.MaxValue;
            const string SVC_ROUTER = "_vti_bin/PSI/ProjectServer.svc";

            bool isHttps = pwaUrl.ToLower().StartsWith("https");
            bool result = true;
            BasicHttpBinding binding = null;

            try
            {
                if (isHttps)
                {
                    // Create a binding for HTTPS.
                    binding = new BasicHttpBinding(BasicHttpSecurityMode.Transport);
                }
                else
                {
                    // Create a binding for HTTP.
                    binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportCredentialOnly);
                }

                binding.Name = "basicHttpConf";
                binding.MessageEncoding = WSMessageEncoding.Text;

                binding.CloseTimeout = new TimeSpan(00, 05, 00);
                binding.OpenTimeout = new TimeSpan(00, 05, 00);
                binding.ReceiveTimeout = new TimeSpan(00, 05, 00);
                binding.SendTimeout = new TimeSpan(00, 05, 00);
                binding.TextEncoding = System.Text.Encoding.UTF8;

                // If the TransferMode is buffered, the MaxBufferSize and 
                // MaxReceived MessageSize must be the same value.
                binding.TransferMode = TransferMode.Buffered;
                binding.MaxBufferSize = MAXSIZE;
                binding.MaxReceivedMessageSize = MAXSIZE;
                binding.MaxBufferPoolSize = MAXSIZE;


                binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Ntlm;
                binding.GetType().GetProperty("ReaderQuotas").SetValue(binding, XmlDictionaryReaderQuotas.Max, null);
                // The endpoint address is the ProjectServer.svc router for all public PSI calls.
                EndpointAddress address = new EndpointAddress(pwaUrl + SVC_ROUTER);

                adminClient = new SvcAdmin.AdminClient(binding, address);
                adminClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                adminClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                projectClient = new SvcProject.ProjectClient(binding, address);
                projectClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                projectClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                queueSystemClient = new SvcQueueSystem.QueueSystemClient(binding, address);
                queueSystemClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                queueSystemClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                resourceClient = new SvcResource.ResourceClient(binding, address);
                resourceClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                resourceClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                lookupTableClient = new SvcLookupTable.LookupTableClient(binding, address);
                lookupTableClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                lookupTableClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;


                customFieldsClient = new SvcCustomFields.CustomFieldsClient(binding, address);
                customFieldsClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                customFieldsClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                calendarClient = new SvcCalendar.CalendarClient(binding, address);
                calendarClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                calendarClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                archiveClient = new SvcArchive.ArchiveClient(binding, address);
                archiveClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                archiveClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;

                pwaClient = new SvcStatusing.StatusingClient(binding, address);
                pwaClient.ChannelFactory.Credentials.Windows.AllowedImpersonationLevel
                    = TokenImpersonationLevel.Impersonation;
                pwaClient.ChannelFactory.Credentials.Windows.AllowNtlm = true;
            }
            catch (Exception ex)
            {
                result = false;
            }
            return result;
        }

        #endregion

        public static FiscalMonth GetCurrentFiscalMonth()
        {
            SvcAdmin.FiscalPeriodDataSet fiscalPeriods = adminClient.ReadFiscalPeriods(DateTime.Now.Year);
            if (fiscalPeriods.FiscalPeriods.Rows.Count > 0)
            {
                foreach (DataRow row in fiscalPeriods.FiscalPeriods.Rows)
                {
                    SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow fiscalRow = (SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow)row;
                    if (DateTime.Now >= fiscalRow.WFISCAL_PERIOD_START_DATE && DateTime.Now <= fiscalRow.WFISCAL_PERIOD_FINISH_DATE)
                    {
                        return new FiscalMonth() { From = fiscalRow.WFISCAL_PERIOD_START_DATE, To = fiscalRow.WFISCAL_PERIOD_FINISH_DATE };
                    }
                }
            }
            fiscalPeriods = adminClient.ReadFiscalPeriods(DateTime.Now.Year - 1);
            if (fiscalPeriods.FiscalPeriods.Rows.Count > 0)
            {
                foreach (DataRow row in fiscalPeriods.FiscalPeriods.Rows)
                {
                    SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow fiscalRow = (SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow)row;
                    if (DateTime.Now >= fiscalRow.WFISCAL_PERIOD_START_DATE && DateTime.Now <= fiscalRow.WFISCAL_PERIOD_FINISH_DATE)
                    {
                        return new FiscalMonth() { From = fiscalRow.WFISCAL_PERIOD_START_DATE, To = fiscalRow.WFISCAL_PERIOD_FINISH_DATE };
                    }
                }
            }

            fiscalPeriods = adminClient.ReadFiscalPeriods(DateTime.Now.Year + 1);
            if (fiscalPeriods.FiscalPeriods.Rows.Count > 0)
            {
                foreach (DataRow row in fiscalPeriods.FiscalPeriods.Rows)
                {
                    SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow fiscalRow = (SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow)row;
                    if (DateTime.Now >= fiscalRow.WFISCAL_PERIOD_START_DATE && DateTime.Now <= fiscalRow.WFISCAL_PERIOD_FINISH_DATE)
                    {
                        return new FiscalMonth() { From = fiscalRow.WFISCAL_PERIOD_START_DATE, To = fiscalRow.WFISCAL_PERIOD_FINISH_DATE };
                    }
                }
            }
            return new FiscalMonth() { From = DateTime.MinValue, To = DateTime.Now };
        }
        public static CustomFieldDataSet ReadCustomFields()
        {
            return customFieldsClient.ReadCustomFields(string.Empty, false);
        }

        public static LookupTableDataSet ReadLookupTables()
        {
            return lookupTableClient.ReadLookupTables(string.Empty, false, 1);
        }

        public static ProjectDataSet ReadProject(Guid projectUID)
        {
            return projectClient.ReadProject(projectUID, SvcProject.DataStoreEnum.PublishedStore);
        }

        public static string ReadTaskEntityUID()
        {
            return Constants.LOOKUP_ENTITY_ID;
        }



        internal static void UpdateLookupTables(LookupTableDataSet lookupTableDataSet)
        {
            using (OperationContextScope scope = new OperationContextScope(lookupTableClient.InnerChannel))
            {
                try
                {
                    WcfHelpers.UseCorrectHeaders(isImpersonated);
                    lookupTableClient.CheckOutLookupTables(new Guid[] { new Guid(Constants.LOOKUP_ENTITY_ID) });
                    lookupTableClient.UpdateLookupTables(lookupTableDataSet, false, true, 1033);
                    //lookupTableClient.CheckInLookupTables(new Guid[] { new Guid(Constants.LOOKUP_ENTITY_ID) },true);
                }
                catch (SoapException ex)
                {
                    string errMess = "";
                    // Pass the exception to the PSClientError constructor to get 
                    // all error information.
                    PSLib.PSClientError psiError = new PSLib.PSClientError(ex);
                    PSLib.PSErrorInfo[] psiErrors = psiError.GetAllErrors();

                    for (int j = 0; j < psiErrors.Length; j++)
                    {
                        errMess += psiErrors[j].ErrId.ToString() + "\n";
                    }
                    errMess += "\n" + ex.Message.ToString();
                    // Send error string to console or message box.
                }
            }
        }
    }


    public sealed class MySettings : ApplicationSettingsBase
    {
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("http://LocalHost/PWA/")]
        public string ProjectServerURL
        {
            get { return (string)this["ProjectServerURL"]; }
            set { this["ProjectServerURL"] = value; }
        }
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("FormsAdmin")]
        public string UserName
        {
            get { return (string)this["UserName"]; }
            set { this["UserName"] = value; }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("pass@word1")]
        public string PassWord
        {
            get { return (string)this["PassWord"]; }
            set { this["PassWord"] = value; }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("true")]
        public bool IsWindowsAuth
        {
            get { return (bool)this["IsWindowsAuth"]; }
            set { this["IsWindowsAuth"] = value; }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("true")]
        public bool UseDefaultWindowsCredentials
        {
            get { return (bool)this["UseDefaultWindowsCredentials"]; }
            set { this["UseDefaultWindowsCredentials"] = value; }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("80")]
        public int WindowsPort
        {
            get { return (int)this["WindowsPort"]; }
            set { this["WindowsPort"] = value; }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("81")]
        public int FormsPort
        {
            get { return (int)this["FormsPort"]; }
            set { this["FormsPort"] = value; }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("false")]
        public bool WaitForQueue
        {
            get { return (bool)this["WaitForQueue"]; }
            set { this["WaitForQueue"] = value; }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("false")]
        public bool WaitForIndividualQueue
        {
            get { return (bool)this["WaitForIndividualQueue"]; }
            set { this["WaitForIndividualQueue"] = value; }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("false")]
        public bool AutoLogin
        {
            get { return (bool)this["AutoLogin"]; }
            set { this["AutoLogin"] = value; }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("false")]
        public bool UseAppConfig
        {
            get { return (bool)this["UseAppConfig"]; }
            set { this["UseAppConfig"] = value; }
        }
    }

    public class LangItem
    {
        int lcid; string langName;
        public int LCID
        {
            get { return lcid; }
            set { lcid = value; }
        }
        public string LangName
        {
            get { return langName; }
            set { langName = value; }
        }
        public LangItem(int Lcid, string name)
        { LCID = Lcid; LangName = name; }
    }

    #region SSL Certificate Handling Class
    // The MyCertificateValidation class is needed for handling SSL certificates.
    // It checks the validity of a certificate and shows a 
    // message to allow the user to choose whether to continue
    // in case of a potentially invalid certificate.

    public class MyCertificateValidation : ICertificatePolicy
    {
        // Default policy for certificate validation.
        public static bool CheckDefaultValidate;

        public enum CertificateProblem : long
        {
            CertEXPIRED = 0x800B0101,
            CertVALIDITYPERIODNESTING = 0x800B0102,
            CertROLE = 0x800B0103,
            CertPATHLENCONST = 0x800B0104,
            CertCRITICAL = 0x800B0105,
            CertPURPOSE = 0x800B0106,
            CertISSUERCHAINING = 0x800B0107,
            CertMALFORMED = 0x800B0108,
            CertUNTRUSTEDROOT = 0x800B0109,
            CertCHAINING = 0x800B010A,
            CertREVOKED = 0x800B010C,
            CertUNTRUSTEDTESTROOT = 0x800B010D,
            CertREVOCATION_FAILURE = 0x800B010E,
            CertCN_NO_MATCH = 0x800B010F,
            CertWRONG_USAGE = 0x800B0110,
            CertUNTRUSTEDCA = 0x800B0112
        }

        public bool CheckValidationResult(ServicePoint sp, X509Certificate cert,
            WebRequest request, int problem)
        {
            if (problem == 0) return true;
            return CheckDefaultValidate;
        }

        private String GetProblemMessage(CertificateProblem Problem)
        {
            String ProblemMessage = "";
            CertificateProblem problemList = new CertificateProblem();
            String ProblemCodeName = Enum.GetName(problemList.GetType(), Problem);
            if (ProblemCodeName != null)
                ProblemMessage = ProblemMessage + ProblemCodeName;
            else
                ProblemMessage = "Unknown Certificate Problem";
            return ProblemMessage;
        }
    }
    #endregion
}
