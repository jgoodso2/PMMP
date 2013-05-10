using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SvcQueueSystem;
using SvcResource;
using SvcCustomFields;
using SvcLookupTable;
using SvcProject;
using Repository;
using SvcAdmin;
using System.Collections;
using SvcArchive;
using SvcCalendar;

namespace PMMP
{
    public class DataRepository
    {
        public static AdminClient adminClient;
        public static ArrayList adUsers;
        public static ArchiveClient archiveClient;
        public static bool autoLogin;
        public static CalendarClient calendarClient;
        public static CustomFieldsClient customFieldsClient;
        public static int formsPort;
        public static string impersonatedUserName;
        public static bool isImpersonated;
        public static bool isWindowsAuth;
        public static Guid jobGuid;
        public static int loginStatus;
        public static LookupTableClient lookupTableClient;
        public static MySettings mySettings;
        public static string password;
        public static ProjectClient projectClient;
        public static Guid projectGuid;
        public static string projectServerUrl;
        public static Guid pwaSiteId;
        public static QueueSystemClient queueSystemClient;
        public static ResourceClient resourceClient;
        public static bool useDefaultWindowsCredentials;
        public static string userName;
        public static bool waitForIndividualQueue;
        public static bool waitForQueue;
        public static int windowsPort;

        public DataRepository();

        public static string CatchFaultException(System.ServiceModel.FaultException faultEx);
        public static void ClearImpersonation();
        public static bool ConfigClientEndpoints();
        public void LoadProjects(string url);
        public static bool P14Login(string projectserverURL);
        public static CustomFieldDataSet ReadCustomFields();
        public static LookupTableDataSet ReadLookupTables();
        public static ProjectDataSet ReadProject(Guid projectUID);
        public static ProjectDataSet ReadProjectsList();
        public static string ReadTaskEntityUID();

        public enum queryType
        {
            GroupAndName = 0,
            GroupAndDisplayName = 1,
            UserNameAndName = 2,
            UserNameandDisplayName = 3,
        }
    }
}
