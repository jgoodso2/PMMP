using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using Microsoft.SharePoint;

namespace PMMPresentation.Support
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

        public static string[] GetValueFromMultiChoice(object oValue)
        {
            string[] value = null;

            if (oValue != null)
            {
                SPFieldMultiChoiceValue fieldValue = new SPFieldMultiChoiceValue(oValue.ToString());
                value = new string[fieldValue.Count];
                for (int i = 0; i < fieldValue.Count; i++)                 
                    value[i] = fieldValue[i];                
            }

            return value;
        }

        public static DateTime? GetValueAsDateTime(object oValue)
        {
            DateTime? value = null;

            if (oValue != null)
                value = Convert.ToDateTime(oValue);

            return value;
        }
    }
}
