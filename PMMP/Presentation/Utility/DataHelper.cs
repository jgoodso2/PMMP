using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Collections.Specialized;
using System.Data;
using SvcProject;
using SvcCustomFields;


namespace PMMP
{
    public class DataHelper
    {
        public static string GetValue(object value)
        {
            Repository.Utility.WriteLog("GetValue started", System.Diagnostics.EventLogEntryType.Information);
            if (value != null)
            {
                Repository.Utility.WriteLog("GetValue completed successfully", System.Diagnostics.EventLogEntryType.Information);
                return value.ToString();
            }
            Repository.Utility.WriteLog("GetValue completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return string.Empty;
        }

        public static int GetValueAsInteger(object oValue)
        {
            Repository.Utility.WriteLog("GetValueAsInteger started", System.Diagnostics.EventLogEntryType.Information);
            int value = 0;

            if (oValue != null)
            {
                value = Convert.ToInt32(oValue.ToString());
                Repository.Utility.WriteLog("GetValueAsInteger completed successfully", System.Diagnostics.EventLogEntryType.Information);
            }

            return value;
        }


        public static object GetValueFromCustomFieldTextOrDate(DataRow dataRow, CustomFieldType type, CustomFieldDataSet dataSet)
        {
            try
            {
                StringCollection value = new StringCollection();
                Guid mdPropID = Guid.Empty;
                if ((dataSet as CustomFieldDataSet).CustomFields.Any(t => t.MD_PROP_NAME == type.GetString()))
                {
                    mdPropID = (dataSet as CustomFieldDataSet).CustomFields.First(t => t.MD_PROP_NAME == type.GetString()).MD_PROP_UID;
                }
                if (mdPropID == Guid.Empty)
                {
                    return null;
                }
                if (type == CustomFieldType.EstStart || type == CustomFieldType.EstFinish)
                {
                    if ((dataRow.Table.DataSet as ProjectDataSet).TaskCustomFields.Any(t => t.TASK_UID == (dataRow as ProjectDataSet.TaskRow).TASK_UID && t.MD_PROP_UID == mdPropID))
                        return (dataRow.Table.DataSet as ProjectDataSet).TaskCustomFields.First(t => t.TASK_UID == (dataRow as ProjectDataSet.TaskRow).TASK_UID && t.MD_PROP_UID == mdPropID && !t.IsDATE_VALUENull()).DATE_VALUE;
                    else
                        return null;
                }
                else if (type == CustomFieldType.PMT || type == CustomFieldType.ReasonRecovery)
                {
                    if ((dataRow.Table.DataSet as ProjectDataSet).TaskCustomFields.Any(t => t.TASK_UID == (dataRow as ProjectDataSet.TaskRow).TASK_UID && t.MD_PROP_UID == mdPropID && !t.IsTEXT_VALUENull()))
                        return (dataRow.Table.DataSet as ProjectDataSet).TaskCustomFields.First(t => t.TASK_UID == (dataRow as ProjectDataSet.TaskRow).TASK_UID && t.MD_PROP_UID == mdPropID && !t.IsTEXT_VALUENull()).TEXT_VALUE;
                    else
                        return null;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        public static string[] GetValueFromMultiChoice(object oValue, CustomFieldType type)
        {
            Repository.Utility.WriteLog("GetValueFromMultiChoice started", System.Diagnostics.EventLogEntryType.Information);
            StringCollection value = new StringCollection();

            if (oValue != null)
            {
                string[] fieldValue = oValue.ToString().Split(",".ToCharArray());
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
            Repository.Utility.WriteLog("GetValueFromMultiChoice completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return array;
        }

        public static DateTime? GetValueAsDateTime(object oValue)
        {
            Repository.Utility.WriteLog("GetValueAsDateTime started", System.Diagnostics.EventLogEntryType.Information);
            DateTime? value = null;

            if (oValue != null && !string.IsNullOrEmpty(oValue.ToString()))
                value = Convert.ToDateTime(oValue);
            Repository.Utility.WriteLog("GetValueAsDateTime completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return value;
        }
    }
}
