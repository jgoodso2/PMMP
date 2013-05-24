using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Collections.Specialized;


namespace PMMP
{
    public class DataHelper
    {
        public static string GetValue(object value)
        {
            if (value != null)
                return value.ToString();

            return string.Empty;
        }

        public static int GetValueAsInteger(object oValue)
        {
            int value = 0;

            if (oValue != null)
                value = Convert.ToInt32(oValue.ToString());

            return value;
        }

        public static string[] GetValueFromMultiChoice(object oValue,CustomFieldType type)
        {
            StringCollection value = new StringCollection();

            if (oValue != null)
            {
                string[] fieldValue =oValue.ToString().Split(",".ToCharArray());
                for (int i = 0; i < fieldValue.Length; i++)
                {
                    foreach (string fieldval in fieldValue[i].Split(",".ToCharArray()))
                    {
                        if (fieldval.StartsWith(type.GetString()))
                        {
                            if (!string.IsNullOrEmpty(fieldval))
                            {
                                value.Add(fieldval);
                            }
                        }
                    }
                }
            }

            if (value.Count == 0)
            {
                return new string[0];
            }
            string[] array = new string[value.Count];
            value.CopyTo(array, 0);
            return array;
        }

        public static DateTime? GetValueAsDateTime(object oValue)
        {
            DateTime? value = null;

            if (oValue != null && !string.IsNullOrEmpty(oValue.ToString()))
                value = Convert.ToDateTime(oValue);

            return value;
        }
    }
}
