using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace PMMPresentation.Console
{
    public class Configuration
    {
        public static string SampleDataFile
        {
            get { return GetConfigurationValue("sampleDataFile"); }
        }

        public static string TemplateFile
        {
            get { return GetConfigurationValue("templateFile"); }
        }

        private static string GetConfigurationValue(string key)
        {
            string value = null;
            if (ConfigurationManager.AppSettings.Count > 0)
            {
                if (ConfigurationManager.AppSettings.AllKeys.Contains(key))
                {
                    value = ConfigurationManager.AppSettings[key];
                }
            }

            return value;
        }
    }
}
