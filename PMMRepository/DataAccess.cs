using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using SvcProject;

namespace Repository
{
    /// <summary>
    /// 
    /// </summary>
    public class DataAccess
    {
        private DataSet projectDataSet;
        private DataSet customFieldsDataSet;
        private DataSet lookUpDataSet;
        private Dictionary<string, CustomFieldDTO> customFieldsTaskDictionary = new Dictionary<string, CustomFieldDTO>();
        private string entity_uid;
        private Guid project_id;

        public DataAccess(Guid projectid)
        {
            project_id = projectid;
        }



        public DataSet ReadProject(DataSet taskDataSetToCompare)
        {
            entity_uid = DataRepository.ReadTaskEntityUID();
            customFieldsDataSet = DataRepository.ReadCustomFields();
            projectDataSet = DataRepository.ReadProject(project_id);
            lookUpDataSet = DataRepository.ReadLookupTables();
            DataSet outputDataSet = TransformDataSet();

            return CompareTaskData(outputDataSet, taskDataSetToCompare);
        }

        private DataSet CompareTaskData(DataSet outputDataSet, DataSet inputTaskData)
        {
            DataSet deltaDataSet = new DataSet();
            if (inputTaskData == null || inputTaskData.Tables["Task"].Rows.Count < 1)
            {
                return outputDataSet;
            }
            deltaDataSet.Tables.Add(outputDataSet.Tables["Task"].Clone());
            foreach (DataRow row in outputDataSet.Tables["Task"].Rows)
            {
                try
                {
                    if (inputTaskData.Tables["Task"].AsEnumerable().Any<DataRow>(t => t.Field<Guid>("TASK_UID") == new Guid(row["TASK_UID"].ToString())))
                    {
                        DataRow inputRow = inputTaskData.Tables["Task"].AsEnumerable().First<DataRow>(t => t.Field<Guid>("TASK_UID") == new Guid(row["TASK_UID"].ToString()));
                        if (inputRow != null)
                        {
                            if (!CompareRows(row, inputRow))
                            {
                                deltaDataSet.Tables["Task"].ImportRow(row);
                            }
                            else
                            {
                                deltaDataSet.Tables["Task"].ImportRow(row);
                            }

                        }
                    }
                    //IF row not found in destination table add the row as a delta
                    else
                    {
                        deltaDataSet.Tables["Task"].ImportRow(row);
                    }
                }
                catch (Exception ex)
                {
                    // Any exception occured while comparing skip and contniue with next row to compare by catching exception
                    continue;
                }
            }

            foreach (DataRow row in inputTaskData.Tables["Task"].Rows)
            {
                try
                {
                    if (outputDataSet.Tables["Task"].AsEnumerable().Any<DataRow>(t => t.Field<Guid>("TASK_UID") == new Guid(row["TASK_UID"].ToString())))
                    {
                        DataRow inputRow = outputDataSet.Tables["Task"].AsEnumerable().First<DataRow>(t => t.Field<Guid>("TASK_UID") == new Guid(row["TASK_UID"].ToString()));
                        //If a row from input not found in destination mark it deleted
                        if (inputRow == null)
                        {
                            deltaDataSet.Tables["Task"].ImportRow(row);
                            deltaDataSet.Tables["Task"].AsEnumerable().First<DataRow>(t => t.Field<Guid>("TASK_UID") == new Guid(row["TASK_UID"].ToString())).AcceptChanges();
                            deltaDataSet.Tables["Task"].AsEnumerable().First<DataRow>(t => t.Field<Guid>("TASK_UID") == new Guid(row["TASK_UID"].ToString())).Delete();
                        }
                    }
                    else
                    {
                        deltaDataSet.Tables["Task"].ImportRow(row);
                        deltaDataSet.Tables["Task"].AsEnumerable().First<DataRow>(t => t.Field<Guid>("TASK_UID") == new Guid(row["TASK_UID"].ToString())).AcceptChanges();
                        deltaDataSet.Tables["Task"].AsEnumerable().First<DataRow>(t => t.Field<Guid>("TASK_UID") == new Guid(row["TASK_UID"].ToString())).Delete();
                    }
                }
                catch (Exception ex)
                {
                    // Any exception occured while comparing skip and contniue with next row to compare by catching exception
                    continue;
                }
            }
            return deltaDataSet;
        }

        private bool CompareRows(DataRow row, DataRow inputRow)
        {
            string[] fieldsToCompare = new string[]{"TASK_PREDECESSORS","TASK_PCT_COMP","TASK_START_DATE","TASK_FINISH_DATE","TASK_DEADLINE","CUSTOMFIELD_DESC",
                                                                                                                    "TASK_DRIVINGPATH_ID"};
            foreach (string column in fieldsToCompare)
            {
                if (column == "CUSTOMFIELD_DESC")
                {
                    if (!string.IsNullOrEmpty(row["CUSTOMFIELD_TEXT"].ToString()))
                    {
                        if (row["CUSTOMFIELD_TEXT"].ToString().Contains("Show On"))
                        {
                            if (row[column].ToString() != inputRow[column].ToString())
                                return false;
                        }
                    }
                }
                else
                {
                    if (row[column].ToString() != inputRow[column].ToString())
                        return false;
                }
            }
            return true;
        }


        private DataSet TransformDataSet()
        {

            Utility.WriteLog(string.Format("Calling TransformDataSet"), System.Diagnostics.EventLogEntryType.Information);
            DataSet outputDataSet = projectDataSet.Copy();
            //add new table DrivingPath
            outputDataSet.Tables.Add("DrivingPath");
            outputDataSet.Tables["DrivingPath"].Columns.Add("PROJ_UID", typeof(String));
            outputDataSet.Tables["DrivingPath"].Columns.Add("TASK_DRIVINGPATH_ID", typeof(String));
            outputDataSet.Tables["DrivingPath"].Columns.Add("DrivingPathName", typeof(String));


            // Add new columns for output DataSet
            outputDataSet.Tables["Task"].Columns.Add("TASK_PREDECESSORS", typeof(String));
            outputDataSet.Tables["Task"].Columns.Add("TASK_DRIVINGPATH_ID", typeof(String));
            outputDataSet.Tables["Task"].Columns.Add("TASK_MODIFIED_ON", typeof(DateTime));

            //Add new Columns for Custom Fields
            outputDataSet.Tables["Task"].Columns.Add("CUSTOMFIELD_TEXT", typeof(String));
            outputDataSet.Tables["Task"].Columns.Add("CUSTOMFIELD_DESC", typeof(String));

            #region Build Predecessors Dcitionary

            // the query
            var taskResult =
                from x in outputDataSet.Tables["Dependency"].AsEnumerable()
                join y in outputDataSet.Tables["Task"].AsEnumerable()
                        on (Guid)x["LINK_SUCC_UID"] equals (Guid)y["TASK_UID"]
                join z in outputDataSet.Tables["Task"].AsEnumerable()
                        on (Guid)x["LINK_PRED_UID"] equals (Guid)z["TASK_UID"]
                select new { LINK_SUCC_UID = x["LINK_SUCC_UID"], LINK_PRED_UID = x["LINK_PRED_UID"], TASK_ID = z["TASK_ID"], DATA_ROW = z };
            Dictionary<string, List<TASKDTO>> PredecessorsDictionary = new Dictionary<string, List<TASKDTO>>();
            foreach (var dataRow in taskResult)
            {

                if (PredecessorsDictionary.ContainsKey(dataRow.LINK_SUCC_UID.ToString()))
                {
                    TASKDTO task = new TASKDTO();
                    List<TASKDTO> tasks = PredecessorsDictionary[dataRow.LINK_SUCC_UID.ToString()];
                    PredecessorsDictionary.Remove(dataRow.LINK_SUCC_UID.ToString());
                    task.LINK_PRED_UID = dataRow.LINK_PRED_UID.ToString();
                    task.TASK_ID = dataRow.TASK_ID.ToString();
                    task.Row = dataRow.DATA_ROW;
                    tasks.Add(task);
                    PredecessorsDictionary.Add(dataRow.LINK_SUCC_UID.ToString(), tasks);
                }
                else
                {
                    List<TASKDTO> tasks = new List<TASKDTO>();
                    TASKDTO task = new TASKDTO();
                    task.LINK_PRED_UID = dataRow.LINK_PRED_UID.ToString();
                    task.TASK_ID = dataRow.TASK_ID.ToString();
                    task.Row = dataRow.DATA_ROW;
                    tasks.Add(task);
                    PredecessorsDictionary.Add(dataRow.LINK_SUCC_UID.ToString(), tasks);
                }
            }
            #endregion

            #region Build CustomFields
            customFieldsTaskDictionary = new Dictionary<string, CustomFieldDTO>();

            // create the default row to be used when no value found
            var defaultRow = customFieldsDataSet.Tables["CustomFields"].NewRow();

            // the query
            var result = from x in outputDataSet.Tables["TaskCustomFields"].AsEnumerable()
                         join y in customFieldsDataSet.Tables["CustomFields"].AsEnumerable() on (Guid)x["MD_PROP_UID"] equals (Guid)y["MD_PROP_UID"]
                         select new { MD_PROP_NAME = y["MD_PROP_NAME"], MD_ENT_TYPE_UID = y["MD_ENT_TYPE_UID"], CODE_VALUE = x["CODE_VALUE"], TASK_UID = x["TASK_UID"] };

            foreach (var customFieldRow in result)
            {
                string propertyname = "";
                propertyname = customFieldRow.MD_PROP_NAME.ToString();


                string codevalue = customFieldRow.CODE_VALUE.ToString();
                string task_uid = customFieldRow.TASK_UID.ToString();
                if (!string.IsNullOrEmpty(codevalue))
                {
                    EnumerableRowCollection<DataRow> rows = lookUpDataSet.Tables["LookupTableTrees"].AsEnumerable().Where(t => t.Field<Guid>("LT_STRUCT_UID") == new Guid(codevalue));

                    foreach (DataRow row in rows)
                    {
                        CustomFieldDTO customFieldDTO = new CustomFieldDTO();
                        customFieldDTO.Text = propertyname + "_" + row["LT_VALUE_TEXT"].ToString();
                        customFieldDTO.Description = propertyname + "_" + row["LT_VALUE_DESC"].ToString();
                        if (customFieldsTaskDictionary.ContainsKey(task_uid))
                        {
                            CustomFieldDTO existingCustomFieldDTO = customFieldsTaskDictionary[task_uid];
                            customFieldsTaskDictionary.Remove(task_uid);
                            existingCustomFieldDTO.AppendText(customFieldDTO.Text);
                            existingCustomFieldDTO.AppendDescription(customFieldDTO.Description);
                            customFieldsTaskDictionary.Add(task_uid, existingCustomFieldDTO);
                        }
                        else
                        {
                            customFieldsTaskDictionary.Add(task_uid, customFieldDTO);
                        }
                    }
                }
            }
            #endregion

            int count = 0;

            List<DataRow> successorTasks = new List<DataRow>();
            #region Update Tasks Table with Driving Path and SPredecessors and Custom Fields


            //For each Task
            foreach (DataRow dataRow in projectDataSet.Tables["Task"].Rows)
            {
                if (customFieldsTaskDictionary.ContainsKey(dataRow["TASK_UID"].ToString()))
                {
                    CustomFieldDTO customDTO = customFieldsTaskDictionary[dataRow["TASK_UID"].ToString()];
                    DataRow outputRow = outputDataSet.Tables["Task"].AsEnumerable().First<DataRow>(t=>t.Field<Guid>("TASK_UID").ToString() == dataRow["TASK_UID"].ToString());
                    outputRow["CUSTOMFIELD_TEXT"] = customDTO.Text;
                    outputRow["CUSTOMFIELD_DESC"] = customDTO.Description;
                }
                DataRow SuccessorRow = outputDataSet.Tables["Task"].Rows[count];
                //From the Dependency Table get all Successors for this Task
                List<TASKDTO> rows = new List<TASKDTO>();
                if (PredecessorsDictionary.ContainsKey(dataRow["TASK_UID"].ToString()))
                {
                    rows = PredecessorsDictionary[dataRow["TASK_UID"].ToString()];
                }
                DataRow deadLineTaskRow = null;
                if (rows.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (TASKDTO row in rows)
                    {

                        DataRow taskRow = row.Row;
                        sb.Append("," + row.TASK_ID.ToString());
                        //If it as a deadline task stor it so that you can iterate through all tasks in the task chain list to update them later
                        if (SuccessorRow["TASK_DEADLINE"] != System.DBNull.Value)
                        {
                            deadLineTaskRow = SuccessorRow;
                            //Update Driving Path Table 
                            DataRow drivingPathRow = outputDataSet.Tables["DrivingPath"].Rows.Add(taskRow.Field<Guid>("PROJ_UID"), deadLineTaskRow.Field<int>("TASK_ID").ToString(), deadLineTaskRow.Field<string>("TASK_NAME"));
                        }
                        //mantian the list of tasks in the task chain list so as to update deadline task later
                        successorTasks.Add(taskRow);
                        //// add custom fields
                        //if (customFieldsTaskDictionary.ContainsKey(taskRow["TASK_UID"].ToString()))
                        //{
                        //    CustomFieldDTO customDTO = customFieldsTaskDictionary[taskRow["TASK_UID"].ToString()];
                        //    taskRow["CUSTOMFIELD_TEXT"] = customDTO.Text;
                        //    taskRow["CUSTOMFIELD_DESC"] = customDTO.Description;
                        //}
                        taskRow["TASK_MODIFIED_ON"] = DateTime.Now;

                    }
                    //update deadline tasks for each task in the task chain list
                    if (deadLineTaskRow != null)
                    {
                        foreach (DataRow row in successorTasks)
                        {
                            if (BelongsToDrivingPath(outputDataSet, deadLineTaskRow, row, PredecessorsDictionary))
                            {
                                row["TASK_DRIVINGPATH_ID"] = BuildCommaSeperatedValue(row["TASK_DRIVINGPATH_ID"], deadLineTaskRow.Field<int>("TASK_ID").ToString());
                            }
                        }
                        deadLineTaskRow["TASK_DRIVINGPATH_ID"] = BuildCommaSeperatedValue(deadLineTaskRow["TASK_DRIVINGPATH_ID"], deadLineTaskRow.Field<int>("TASK_ID").ToString());
                    }
                    string predescessor = sb.ToString();
                    predescessor = string.Join(",", predescessor.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                    SuccessorRow["TASK_PREDECESSORS"] = predescessor;


                }
                count++;
            }
            #endregion
            Utility.WriteLog(string.Format("TransformDataSet completed successfully"), System.Diagnostics.EventLogEntryType.Information);

            return outputDataSet;
        }

        private bool BelongsToDrivingPath(DataSet outputDataSet, DataRow deadLineTaskRow, DataRow row, Dictionary<string, List<TASKDTO>> PredecessorsDictionary)
        {
            if (deadLineTaskRow["TASK_ID"].ToString() == row["TASK_ID"].ToString())
            {
                return true;
            }
            if (!PredecessorsDictionary.ContainsKey(deadLineTaskRow["TASK_UID"].ToString()))
            {
                return false;
            }
            List<string> Pred = PredecessorsDictionary[deadLineTaskRow["TASK_UID"].ToString()].Select(t => t.TASK_ID).ToList();
            foreach (string pred in Pred)
            {
                DataRow predTask = outputDataSet.Tables["Task"].AsEnumerable().First(t => t.Field<int>("TASK_ID").ToString() == pred);
                if (BelongsToDrivingPath(outputDataSet, predTask, row, PredecessorsDictionary) == true)
                {
                    return true;
                }
            }

            return false;

        }

        private string BuildCommaSeperatedValue(object taskRow, string value)
        {
            if (taskRow != null && !taskRow.ToString().Split(",".ToCharArray()).Contains(value))
            {
                taskRow += "," + value;
                if (taskRow != null)
                {
                    string sTaskRow = taskRow.ToString();
                    sTaskRow = string.Join(",", sTaskRow.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                    return sTaskRow;
                }
            }
            return "";
        }

        public DateTime? GetProjectStatusDate(ProjectDataSet projectDataSet,Guid projectGuid)
        {
            return DataRepository.GetProjectStatusDate(projectDataSet, projectGuid);
        }

        public List<FiscalUnit> GetProjectStatusPeriods(DateTime? date)
        {
            return DataRepository.GetProjectStatusPeriods(date);
        }

        public List<FiscalUnit> GetProjectStatusWeekPeriods(DateTime? projectStatusDate)
        {
            return DataRepository.GetProjectStatusWeekPeriods(projectStatusDate);
        }
    }
}
