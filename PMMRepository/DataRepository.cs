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
using SvcResource;
using System.Globalization;

namespace Repository
{
    public class FiscalUnit
    {
        public FiscalUnit()
        {
            From = DateTime.MinValue;
            To = DateTime.MaxValue;
        }
        public FiscalUnit(DateTime fromDate, DateTime toDate, int month, int year, bool isWeekly, int weekNoStart)
        {
            From = fromDate;
            To = toDate;
            Month = month;
            Year = year;
            IsWeekly = isWeekly;
            WeekNoStart = weekNoStart;
        }

        public FiscalUnit(DateTime fromDate, DateTime toDate, int month, int year, bool nonFiscal)
        {
            From = fromDate;
            To = toDate;
            Month = month;
            Year = year;
            IsWeekly = false;
            NonFiscal = nonFiscal;
        }

        public bool IsWeekly { get; set; }
        public int WeekNoStart { get; set; }
        public int Month { get; set; }
        public int Year { get; set; }
        public DateTime To { get; set; }
        public DateTime From { get; set; }
        public bool NonFiscal { get; set; }
        public string GetTitle()
        {
            if (NonFiscal)
            {
                return To.ToString("dd/MM");
            }
            if (IsWeekly)
            {
                DateTime date = new DateTime(Year, Month, 1);
                return date.ToString("MMM y " + string.Format("WK{0}", WeekNoStart));
            }
            else
            {
                DateTime date = new DateTime(Year, Month, 1);
                return date.ToString("MMM y");
            }

        }

        public int GetNoOfWeeks()
        {
            return (To - From).Days / 7;
        }
    }

    /// <summary>
    /// 
    /// </summary>
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


        public static void SetImpersonation(bool isWindowsUser)
        {
            string impersonatedUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            Guid resourceGuid = GetResourceUid(impersonatedUser);
            Guid trackingGuid = Guid.NewGuid();
            Guid siteId = Guid.Empty;           // Project Web App site ID.
            CultureInfo languageCulture = null; // The language culture is not used.
            CultureInfo localeCulture = null;   // The locale culture is not used.

            WcfHelpers.SetImpersonationContext(isWindowsUser, impersonatedUser, resourceGuid, trackingGuid, siteId,
                                               languageCulture, localeCulture);

        }

        // Get the GUID for a Project Server account name. 
        public static Guid GetResourceUid(String accountName)
        {
            Guid resourceUid = Guid.Empty;
            ResourceDataSet resourceDs = new ResourceDataSet();

            // Filter for the account name, which can be a 
            // Windows account or Project Server account.
            PSLib.Filter filter = new PSLib.Filter();
            filter.FilterTableName = resourceDs.Resources.TableName;

            PSLib.Filter.Field accountField = new PSLib.Filter.Field(
                    resourceDs.Resources.TableName,
                    resourceDs.Resources.WRES_ACCOUNTColumn.ColumnName);
            filter.Fields.Add(accountField);

            PSLib.Filter.FieldOperator op = new PSLib.Filter.FieldOperator(
                    PSLib.Filter.FieldOperationType.Equal,
                    resourceDs.Resources.WRES_ACCOUNTColumn.ColumnName, accountName);
            filter.Criteria = op;

            string filterXml = filter.GetXml();

            resourceDs = resourceClient.ReadResources(filterXml, false);

            // Return the account GUID.
            if (resourceDs.Resources.Rows.Count > 0)
                resourceUid = (Guid)resourceDs.Resources.Rows[0]["RES_UID"];

            return resourceUid;
        }
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

                    Utility.WriteLog(string.Format("Calling ReadStatus"), System.Diagnostics.EventLogEntryType.Information);
                    SvcStatusing.StatusingDataSet dataSet = pwaClient.ReadStatus(Guid.Empty, DateTime.MinValue, DateTime.MaxValue);
                    Utility.WriteLog(string.Format("ReadStatus Successful"), System.Diagnostics.EventLogEntryType.Information);
                    // Get projects of type normal, templates, proposals, master, and inserted.
                    string projectName = string.Empty;
                    Utility.WriteLog(string.Format("Calling ReadStatus on Project Store"), System.Diagnostics.EventLogEntryType.Information);
                    projectList.Merge(projectClient.ReadProjectStatus(Guid.Empty, SvcProject.DataStoreEnum.PublishedStore,
                        projectName, (int)PSLib.Project.ProjectType.Project));
                    Utility.WriteLog(string.Format("ReadStatus on Project Store Successful"), System.Diagnostics.EventLogEntryType.Information);
                    Utility.WriteLog(string.Format("Calling ReadStatus on Inserted Store"), System.Diagnostics.EventLogEntryType.Information);
                    projectList.Merge(projectClient.ReadProjectStatus(Guid.Empty, SvcProject.DataStoreEnum.PublishedStore,
                        projectName, (int)PSLib.Project.ProjectType.InsertedProject));
                    Utility.WriteLog(string.Format("ReadStatus on Inserted Store Successful"), System.Diagnostics.EventLogEntryType.Information);
                    Utility.WriteLog(string.Format("Calling ReadStatus on Published Store"), System.Diagnostics.EventLogEntryType.Information);
                    projectList.Merge(projectClient.ReadProjectStatus(Guid.Empty, SvcProject.DataStoreEnum.PublishedStore,
                        projectName, (int)PSLib.Project.ProjectType.MasterProject));
                    Utility.WriteLog(string.Format("ReadStatus on Inserted Published Store Successful"), System.Diagnostics.EventLogEntryType.Information);


                    //SvcProject.ProjectDataSet pds = projectClient.ReadProjectList(); // this fails if no permission... Conversely...ReadProjectStatus returns 0 if no permission.  
                }

            }
            catch (Exception ex)
            {
                Utility.WriteLog(string.Format("An error occured in ReadProjectsList and the error ={0}", ex.Message), System.Diagnostics.EventLogEntryType.Information);
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
                                Utility.WriteLog(string.Format("Logging in with username={0} password={1} Domain={1} ", windowsUserName, password, windowsDomainName), System.Diagnostics.EventLogEntryType.Information);
                                result = loginWindows.Login();
                                Utility.WriteLog(string.Format("Logging in result ={0} ", result.ToString()), System.Diagnostics.EventLogEntryType.Information);
                            }
                        }
                    }
                    else
                    {
                        // Forms authentication requires the Authentication web service in Microsoft SharePoint Foundation.
                        Utility.WriteLog(string.Format("Logging in with username={0} password={1}  ", userName, password), System.Diagnostics.EventLogEntryType.Information);
                        result = WcfHelpers.LogonWithMsf(userName, password, new Uri(baseUrl));
                        Utility.WriteLog(string.Format("Logging in result ={0} ", result.ToString()), System.Diagnostics.EventLogEntryType.Information);
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

        public static FiscalUnit GetFiscalMonth(DateTime? date)
        {
            if (!date.HasValue)
            {
                return new FiscalUnit() { From = DateTime.MinValue, To = DateTime.MaxValue };
            }
            Utility.WriteLog(string.Format("Calling GetCurrentFiscalMonth"), System.Diagnostics.EventLogEntryType.Information);
            using (OperationContextScope scope = new OperationContextScope(adminClient.InnerChannel))
            {
                WcfHelpers.UseCorrectHeaders(isImpersonated);
                SvcAdmin.FiscalPeriodDataSet fiscalPeriods = adminClient.ReadFiscalPeriods(date.Value.Year);
                if (fiscalPeriods.FiscalPeriods.Rows.Count > 0)
                {
                    foreach (DataRow row in fiscalPeriods.FiscalPeriods.Rows)
                    {
                        SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow fiscalRow = (SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow)row;
                        if (date >= fiscalRow.WFISCAL_PERIOD_START_DATE && date <= fiscalRow.WFISCAL_PERIOD_FINISH_DATE)
                        {
                            return new FiscalUnit() { From = fiscalRow.WFISCAL_PERIOD_START_DATE, To = fiscalRow.WFISCAL_PERIOD_FINISH_DATE };
                        }
                    }
                }
                fiscalPeriods = adminClient.ReadFiscalPeriods(date.Value.Year - 1);
                if (fiscalPeriods.FiscalPeriods.Rows.Count > 0)
                {
                    foreach (DataRow row in fiscalPeriods.FiscalPeriods.Rows)
                    {
                        SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow fiscalRow = (SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow)row;
                        if (date >= fiscalRow.WFISCAL_PERIOD_START_DATE && date <= fiscalRow.WFISCAL_PERIOD_FINISH_DATE)
                        {
                            return new FiscalUnit() { From = fiscalRow.WFISCAL_PERIOD_START_DATE, To = fiscalRow.WFISCAL_PERIOD_FINISH_DATE };
                        }
                    }
                }

                fiscalPeriods = adminClient.ReadFiscalPeriods(date.Value.Year + 1);
                if (fiscalPeriods.FiscalPeriods.Rows.Count > 0)
                {
                    foreach (DataRow row in fiscalPeriods.FiscalPeriods.Rows)
                    {
                        SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow fiscalRow = (SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow)row;
                        if (date >= fiscalRow.WFISCAL_PERIOD_START_DATE && date <= fiscalRow.WFISCAL_PERIOD_FINISH_DATE)
                        {
                            Utility.WriteLog(string.Format("GetCurrentFiscalMonth completed successfully"), System.Diagnostics.EventLogEntryType.Information);
                            return new FiscalUnit() { From = fiscalRow.WFISCAL_PERIOD_START_DATE, To = fiscalRow.WFISCAL_PERIOD_FINISH_DATE };
                        }
                    }
                }
            }
            Utility.WriteLog(string.Format("GetCurrentFiscalMonth completed successfully"), System.Diagnostics.EventLogEntryType.Information);
            return new FiscalUnit() { From = DateTime.MinValue, To = DateTime.MaxValue };
        }
        public static CustomFieldDataSet ReadCustomFields()
        {
            Utility.WriteLog(string.Format("Calling ReadCustomFields"), System.Diagnostics.EventLogEntryType.Information);
            using (OperationContextScope scope = new OperationContextScope(customFieldsClient.InnerChannel))
            {
                WcfHelpers.UseCorrectHeaders(isImpersonated);
            }
            var obj = customFieldsClient.ReadCustomFields(string.Empty, false);
            Utility.WriteLog(string.Format("ReadCustomFields completed successfully"), System.Diagnostics.EventLogEntryType.Information);
            return obj;
        }

        public static LookupTableDataSet ReadLookupTables()
        {
            using (OperationContextScope scope = new OperationContextScope(lookupTableClient.InnerChannel))
            {
                WcfHelpers.UseCorrectHeaders(isImpersonated);
            }
            return lookupTableClient.ReadLookupTables(string.Empty, false, 1);
        }


        public static ProjectDataSet ReadProject(Guid projectUID)
        {
            Utility.WriteLog(string.Format("Calling ReadProject"), System.Diagnostics.EventLogEntryType.Information);
            using (OperationContextScope scope = new OperationContextScope(projectClient.InnerChannel))
            {
                WcfHelpers.UseCorrectHeaders(isImpersonated);
            }
            var obj = projectClient.ReadProject(projectUID, SvcProject.DataStoreEnum.PublishedStore);
            Utility.WriteLog(string.Format("ReadProject completed successfully"), System.Diagnostics.EventLogEntryType.Information);
            return obj;
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
                    WcfHelpers.UseCorrectHeaders(true);
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

        internal static List<FiscalUnit> GetProjectStatusPeriods(DateTime? date)
        {
            List<FiscalUnit> months = new List<FiscalUnit>();
            if (!date.HasValue)
            {
                return months;
            }

            using (OperationContextScope scope = new OperationContextScope(adminClient.InnerChannel))
            {
                WcfHelpers.UseCorrectHeaders(isImpersonated);
                SvcAdmin.FiscalPeriodDataSet fiscalPeriods = adminClient.ReadFiscalPeriods(date.Value.Year);
                if (fiscalPeriods.FiscalPeriods.Rows.Count > 0)
                {

                    for (int count = 0; count < fiscalPeriods.FiscalPeriods.Rows.Count; count++)
                    {
                        DataRow row = fiscalPeriods.FiscalPeriods.Rows[count];
                        SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow fiscalRow = (SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow)row;
                        int noOfWeeks = 0;
                        if (date >= fiscalRow.WFISCAL_PERIOD_START_DATE && date <= fiscalRow.WFISCAL_PERIOD_FINISH_DATE)
                        {
                            for (int i = 3; i > 0; i--)
                            {

                                if (count >= 0)
                                {
                                    DataRow row1 = fiscalPeriods.FiscalPeriods.Rows[count - i];
                                    SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow fiscalRow1 = (SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow)row1;
                                    FiscalUnit fiscalMonth = new FiscalUnit(fiscalRow1.WFISCAL_PERIOD_START_DATE, fiscalRow1.WFISCAL_PERIOD_FINISH_DATE, fiscalRow1.WFISCAL_MONTH, fiscalRow1.WFISCAL_YEAR, false, 0);
                                    months.Add(fiscalMonth);
                                }
                            }
                            //count += 3;

                            FiscalUnit fiscalMonth1 = new FiscalUnit(fiscalRow.WFISCAL_PERIOD_START_DATE, fiscalRow.WFISCAL_PERIOD_START_DATE.AddDays(7), fiscalRow.WFISCAL_MONTH, fiscalRow.WFISCAL_YEAR, true, (1));
                            FiscalUnit fiscalMonth2 = new FiscalUnit(fiscalRow.WFISCAL_PERIOD_START_DATE.AddDays(7), fiscalRow.WFISCAL_PERIOD_START_DATE.AddDays(14), fiscalRow.WFISCAL_MONTH, fiscalRow.WFISCAL_YEAR, true, (2));
                            FiscalUnit fiscalMonth3 = new FiscalUnit(fiscalRow.WFISCAL_PERIOD_START_DATE.AddDays(14), fiscalRow.WFISCAL_PERIOD_START_DATE.AddDays(21), fiscalRow.WFISCAL_MONTH, fiscalRow.WFISCAL_YEAR, true, (3));
                            FiscalUnit fiscalMonth4 = new FiscalUnit(fiscalRow.WFISCAL_PERIOD_START_DATE.AddDays(21), fiscalRow.WFISCAL_PERIOD_START_DATE.AddDays(28), fiscalRow.WFISCAL_MONTH, fiscalRow.WFISCAL_YEAR, true, (4));
                            months.Add(fiscalMonth1);
                            months.Add(fiscalMonth2);
                            months.Add(fiscalMonth3);
                            months.Add(fiscalMonth4);
                            if (new FiscalUnit(fiscalRow.WFISCAL_PERIOD_START_DATE, fiscalRow.WFISCAL_PERIOD_FINISH_DATE, fiscalRow.WFISCAL_MONTH, fiscalRow.WFISCAL_YEAR, false, 0).GetNoOfWeeks() > 4)
                            {
                                FiscalUnit fiscalMonth5 = new FiscalUnit(fiscalRow.WFISCAL_PERIOD_START_DATE.AddDays(28), fiscalRow.WFISCAL_PERIOD_START_DATE.AddDays(35), fiscalRow.WFISCAL_MONTH, fiscalRow.WFISCAL_YEAR, true, (5));
                                months.Add(fiscalMonth5);
                            }

                            for (int i = 0; i < 3; i++)
                            {
                                count++;
                                if (count < fiscalPeriods.FiscalPeriods.Rows.Count)
                                {
                                    DataRow row1 = fiscalPeriods.FiscalPeriods.Rows[count];
                                    SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow fiscalRow1 = (SvcAdmin.FiscalPeriodDataSet.FiscalPeriodsRow)row1;
                                    FiscalUnit fiscalMonth = new FiscalUnit(fiscalRow1.WFISCAL_PERIOD_START_DATE, fiscalRow1.WFISCAL_PERIOD_FINISH_DATE, fiscalRow1.WFISCAL_MONTH, fiscalRow1.WFISCAL_YEAR, false, 0);
                                    months.Add(fiscalMonth);
                                }
                            }
                            break;
                        }
                        noOfWeeks += new FiscalUnit(fiscalRow.WFISCAL_PERIOD_START_DATE, fiscalRow.WFISCAL_PERIOD_FINISH_DATE, fiscalRow.WFISCAL_MONTH, fiscalRow.WFISCAL_YEAR, false, 0).GetNoOfWeeks();
                    }
                }
            }
            return months;
        }

        internal static DateTime? GetProjectStatusDate(ProjectDataSet projectDataSet, Guid projUID)
        {
            try
            {
                DateTime date = projectDataSet.Project.First(t => t.PROJ_UID == projUID).PROJ_INFO_STATUS_DATE;
                return date;
            }
            catch
            {
                return null;
            }
        }

        internal static List<FiscalUnit> GetProjectStatusWeekPeriods(DateTime? date)
        {
            List<FiscalUnit> weekly = new List<FiscalUnit>();
            if (!date.HasValue)
            {
                return weekly;
            }

            FiscalUnit fiscalMonth1 = new FiscalUnit(date.Value.AddDays(-35), date.Value.AddDays(-28),date.Value.Month, date.Value.Year, true);
            FiscalUnit fiscalMonth2 = new FiscalUnit(date.Value.AddDays(-28), date.Value.AddDays(-21), date.Value.Month, date.Value.Year, true);
            FiscalUnit fiscalMonth3 = new FiscalUnit(date.Value.AddDays(-21), date.Value.AddDays(-14), date.Value.Month, date.Value.Year, true);
            FiscalUnit fiscalMonth4 = new FiscalUnit(date.Value.AddDays(-14), date.Value.AddDays(-7), date.Value.Month, date.Value.Year, true);
            FiscalUnit fiscalMonth5 = new FiscalUnit(date.Value.AddDays(-7), date.Value, date.Value.Month, date.Value.Year, true);
            FiscalUnit fiscalMonth6 = new FiscalUnit(date.Value, date.Value.AddDays(7), date.Value.Month, date.Value.Year, true);
            FiscalUnit fiscalMonth7 = new FiscalUnit(date.Value.AddDays(7), date.Value.AddDays(14), date.Value.Month, date.Value.Year, true);
            FiscalUnit fiscalMonth8 = new FiscalUnit(date.Value.AddDays(14), date.Value.AddDays(21), date.Value.Month, date.Value.Year, true);
            FiscalUnit fiscalMonth9 = new FiscalUnit(date.Value.AddDays(21), date.Value.AddDays(28), date.Value.Month, date.Value.Year, true);
            FiscalUnit fiscalMonth10 = new FiscalUnit(date.Value.AddDays(28), date.Value.AddDays(25), date.Value.Month, date.Value.Year, true);
            weekly.Add(fiscalMonth1);
            weekly.Add(fiscalMonth2);
            weekly.Add(fiscalMonth3);
            weekly.Add(fiscalMonth4);
            weekly.Add(fiscalMonth5);
            weekly.Add(fiscalMonth6);
            weekly.Add(fiscalMonth7);
            weekly.Add(fiscalMonth8);
            weekly.Add(fiscalMonth9);
            weekly.Add(fiscalMonth10);
            return weekly;
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
