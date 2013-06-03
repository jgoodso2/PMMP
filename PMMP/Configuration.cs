using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using PMMP;

namespace PMMPPresentation
{
    /// <summary>
    /// 
    /// </summary>
    public class Configuration
    {
        public static string ServiceURL
        {
            get { return GetConfigValue(Constants.PROPERTY_NAME_DB_SERVICE_URL); }
            set { SetConfigValue(Constants.PROPERTY_NAME_DB_SERVICE_URL, value); }
        }

        public static string ProjectUID
        {
            get { return GetConfigValue(Constants.PROPERTY_NAME_DB_PROJECT_UID); }
            set { SetConfigValue(Constants.PROPERTY_NAME_DB_PROJECT_UID, value); }
        }

        /// <summary>
        /// Gets a config parameter stored in the web properties bag
        /// </summary>
        /// <param name="key">The parameter key</param>
        /// <returns></returns>
        private static string GetConfigValue(string key)
        {
            string retVal = string.Empty;

            var web = SPContext.Current.Web;

            if (web.Properties.ContainsKey(key))
                retVal = web.Properties[key];

            return retVal;
        }

        private static void SetConfigValue(string key, string value)
        {
            var web = SPContext.Current.Web;

            if (web.Properties.ContainsKey(key))
                web.Properties[key] = value;
            else
                web.Properties.Add(key, value);


            web.AllowUnsafeUpdates = true;
            web.Properties.Update();
            web.AllowUnsafeUpdates = false;
        }
    }
}
