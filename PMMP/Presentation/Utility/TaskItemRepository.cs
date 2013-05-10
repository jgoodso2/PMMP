using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Repository;
using System.Data;
using PMMP;
using Constants = PMMP.Constants;
using System.Collections.Specialized;
using SvcProject;

namespace PMMP
{
    public class TaskItemRepository
    {
        public static void DeleteAllFromList()
        {
            SPWeb web=null;
            if (SPContext.Current != null)
                web = SPContext.Current.Web;
           

            var list = web.Lists.TryGetList(Constants.LIST_NAME_PROJECT_TASKS);
            if (list != null)
            {
                for (int i = list.ItemCount - 1; i >= 0; i--)
                    list.Items[i].Delete();

                list.Update();
            }

        }
        public static DateTime? GetLastUpdateDate()
        {
            DateTime? maxValue = null;

            try
            {
                SPWeb web;
                if (SPContext.Current != null)
                    web = SPContext.Current.Web;
                else
                    web = new SPSite("http://finweb.contoso.com/sites/PMM").OpenWeb();

                SPList objList = web.Lists.TryGetList(Constants.LIST_NAME_PROJECT_TASKS);

                SPQuery objQuery = new SPQuery();
                objQuery.Query = "<OrderBy><FieldRef Name='ModifiedOn' Ascending='False' /></OrderBy><RowLimit>1</RowLimit>";
                objQuery.Folder = objList.RootFolder;

                // Execute the query against the list
                SPListItemCollection colItems = objList.GetItems(objQuery);

                if (colItems.Count > 0)
                {
                    maxValue = Convert.ToDateTime(colItems[0]["ModifiedOn"]);
                }
            }
            catch (Exception ex)
            {

            }

            return maxValue;

        }

        public static TaskGroupData GetTaskGroups()
        {
            TaskGroupData taskData = new TaskGroupData();
            
            IList<TaskItemGroup> retVal = new List<TaskItemGroup>();
            Dictionary<string, IList<TaskItem>> ChartsData = GetChartsData();
            taskData.TaskItemGroups = retVal;
            taskData.ChartsData = ChartsData;
            SPWeb web=null;
            if (SPContext.Current != null)
                web = SPContext.Current.Web;

            var list = web.Lists.TryGetList(Constants.LIST_NAME_PROJECT_TASKS);
            if (list != null)
            {

                var dPathsField = list.Fields[Constants.FieldId_DrivingPath] as SPFieldMultiChoice;
                var dPaths = dPathsField.Choices;
                var chartTypesField = list.Fields[Constants.FieldId_ShowOn] as SPFieldMultiChoice;


                var chartTypes = chartTypesField.Choices;
                dPaths = dPathsField.Choices;
                foreach (string dPath in dPaths)
                {

                    var q = new SPQuery();
                    q.Query += "<Where>" +
                                    "<Eq>" +
                                        "<FieldRef ID='" + dPathsField.Id + "' />" +
                                        "<Value Type='MultiChoice'>" + dPath + "</Value>" +
                                    "</Eq>" +
                        //"<OrderBy>" +
                        //    "<FieldRef ID='" + startField.Id + "' Ascending='True' />" + 
                        //"</OrderBy>" +
                                "</Where>";

                    int taskCount = -1;
                    var taskItemGroup = new TaskItemGroup { DrivingPath = dPath, TaskItems = new List<TaskItem>() };
                    string previousTitle = string.Empty;
                    Dictionary<string, string> dictTitle = new Dictionary<string, string>();
                    int totalUnCompletedtaskCount = 0, totalCompletedTaskCount = 0;

                    List<TaskItem> chartItems = new List<TaskItem>();
                    List<TaskItemGroup> completedTasks = new List<TaskItemGroup>();
                    SPListItemCollection collection = list.GetItems(q);
                    int completedTaskCount = -1;
                    //DateTime? lastUpdate = GetLastUpdateDate();
                    TaskItemGroup completedTaskItemGroup = new TaskItemGroup { DrivingPath = dPath, TaskItems = new List<TaskItem>() };
                    foreach (SPListItem item in collection)
                    {
                        if (item[Constants.FieldId_Deadline] != null)
                        {
                            if (!dictTitle.ContainsKey(dPath.Split(",".ToCharArray())[0]))
                            {
                                dictTitle.Add(dPath.Split(",".ToCharArray())[0], item[Constants.FieldId_Task].ToString());
                            }
                        }

                        if (item[Constants.FieldId_ShowOn] != null)
                        {
                            chartItems.Add(BuildTaskItem(dPath, item));
                        }

                        if (item[Constants.FieldId_PercentComplete] != null && (Convert.ToInt32(item[Constants.FieldId_PercentComplete].ToString().Trim().Trim("%".ToCharArray()).Trim()) < 100))
                        {
                                totalUnCompletedtaskCount++;
                                taskCount++;
                                if (taskCount == 10)
                                {
                                    retVal.Add(taskItemGroup);
                                    taskItemGroup = new TaskItemGroup { DrivingPath = dPath, TaskItems = new List<TaskItem>() };

                                    taskItemGroup.Title = previousTitle;
                                    taskCount = 0;
                                    taskItemGroup.TaskItems.Add(BuildTaskItem(dPath, item));

                                }
                                else
                                {
                                    taskItemGroup.TaskItems.Add(BuildTaskItem(dPath, item));
                                }
                        }
                        else
                        {
                                totalCompletedTaskCount++;
                                completedTaskCount++;
                                if (completedTaskCount == 10)
                                {
                                    completedTasks.Add(completedTaskItemGroup);
                                    completedTaskItemGroup = new TaskItemGroup { DrivingPath = dPath, TaskItems = new List<TaskItem>() };
                                    completedTaskCount = 0;
                                    completedTaskItemGroup.TaskItems.Add(BuildTaskItem(dPath, item));
                                }
                                else
                                {
                                    completedTaskItemGroup.TaskItems.Add(BuildTaskItem(dPath, item));
                                }
                            
                        }
                    }

                    if (totalUnCompletedtaskCount % 10 != 0)
                    {
                        retVal.Add(taskItemGroup);

                    }

                    if (totalCompletedTaskCount % 10 != 0)
                    {
                        completedTasks.Add(completedTaskItemGroup);
                        if (totalUnCompletedtaskCount == 0)
                        {
                            retVal.Add(taskItemGroup);
                        }
                    }


                    if (taskItemGroup.TaskItems.Count > 0 || (completedTasks.Count > 0 && completedTasks[0].TaskItems != null && completedTasks[0].TaskItems.Count > 0))
                    {
                        taskItemGroup.CompletedTaskgroups = completedTasks;
                        taskItemGroup.ChartTaskItems = chartItems;
                        taskItemGroup.Charts = new string[chartTypes.Count];
                        chartTypes.CopyTo(taskItemGroup.Charts, 0);
                        taskItemGroup.Title = dictTitle.ContainsKey(dPath) ? dictTitle[dPath] : "Driving Path template";
                    }




                    if (dPath != null && dictTitle.ContainsKey(dPath.Split(",".ToCharArray())[0]))
                    {
                        foreach (TaskItemGroup group in retVal)
                        {
                            if (group.DrivingPath == dPath)
                            {
                                group.Title = dictTitle[dPath.Split(",".ToCharArray())[0]];
                            }

                            if (group.CompletedTaskgroups != null)
                            {
                                foreach (TaskItemGroup completedGroup in group.CompletedTaskgroups)
                                {
                                    completedGroup.Title = dictTitle[dPath.Split(",".ToCharArray())[0]];
                                }
                            }

                        }
                    }


                }
            }

            if (SPContext.Current == null)
                web.Dispose();

            return taskData;
        }

        private static Dictionary<string, IList<TaskItem>> GetChartsData()
        {
            Dictionary<string, IList<TaskItem>> chartsData = new Dictionary<string, IList<TaskItem>>() ;
            SPWeb web = null;
            if (SPContext.Current != null)
                web = SPContext.Current.Web;

            var list = web.Lists.TryGetList(Constants.LIST_NAME_PROJECT_TASKS);

            if (list != null)
            {
                var chartTypesField = list.Fields[Constants.FieldId_ShowOn] as SPFieldMultiChoice;
                var chartTypes = chartTypesField.Choices;
                foreach (string chartType in chartTypes)
                {

                    var q = new SPQuery();
                    q.Query += "<Where>" +
                                    "<Eq>" +
                                        "<FieldRef ID='" + chartTypesField.Id + "' />" +
                                        "<Value Type='MultiChoice'>" + chartType + "</Value>" +
                                    "</Eq>" +
                                "</Where>";
                    List<TaskItem> items = new List<TaskItem>();
                    SPListItemCollection collection = list.GetItems(q);
                    foreach (SPListItem item in collection)
                    {
                        TaskItem taskItem =  BuildTaskItem("", item);
                        items.Add(taskItem);
                    }
                    if (items.Count > 0)
                    {
                        chartsData.Add(chartType, items);
                    }
                }
            }
            return chartsData;
        }

        public static DataSet GetProjectDataSetFromList()
        {
            ProjectDataSet projectDataSet = new ProjectDataSet();
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
            SPWeb web;
            if (SPContext.Current != null)
                web = SPContext.Current.Web;
            else
                web = new SPSite("http://finweb.contoso.com/sites/PMM").OpenWeb();

            var list = web.Lists.TryGetList(Constants.LIST_NAME_PROJECT_TASKS);
            var dPathsField = list.Fields[Constants.FieldId_DrivingPath] as SPFieldMultiChoice;
            var dShownField = list.Fields[Constants.FieldId_ShowOn] as SPFieldMultiChoice;
            Guid guid = Guid.NewGuid();
            if (list != null)
            {
                for (int i = 0; i < list.ItemCount; i++)
                {
                    DataRow row = outputDataSet.Tables["Task"].NewRow();
                    BuildTaskRowFromListItem(row, list.Items[i], guid);
                    outputDataSet.Tables["Task"].Rows.Add(row);
                }
            }
            return outputDataSet;
        }

        private static void BuildTaskRowFromListItem(DataRow row, SPListItem sPListItem, Guid guid)
        {
            var dPathsField = sPListItem.ParentList.Fields[Constants.FieldId_DrivingPath] as SPFieldMultiChoice;
            var dShownField = sPListItem.ParentList.Fields[Constants.FieldId_ShowOn] as SPFieldMultiChoice;
            string dPath = dPathsField.GetFieldValueAsText(sPListItem[Constants.FieldId_DrivingPath]);
            TaskItem item = BuildTaskItem(dPath, sPListItem);
            if (item.Deadline.HasValue)
            {
                row["TASK_DEADLINE"] = item.Deadline.Value;
            }
            else
            {
                row["TASK_DEADLINE"] = System.DBNull.Value;
            }

            row["TASK_DRIVINGPATH_ID"] = item.DrivingPath;
            row["TASK_FINISH_DATE"] = item.Finish;
            row["TASK_ID"] = item.ID;
            if (item.ModifiedOn.HasValue)
            {
                row["TASK_MODIFIED_ON"] = item.ModifiedOn.Value;
            }
            row["TASK_PREDECESSORS"] = item.Predecessor;
            string[] showOnValues = dShownField.GetFieldValueAsText(sPListItem[Constants.FieldId_ShowOn]).Split(",".ToCharArray());
            for (int i = 0; i < showOnValues.Count(); i++)
            {
                if (!string.IsNullOrEmpty(showOnValues[i]))
                {
                    showOnValues[i] = "Show On_" + showOnValues[i].Trim();
                }
            }
            row["CUSTOMFIELD_DESC"] = string.Join(",", showOnValues);
            row["TASK_START_DATE"] = item.Start;
            row["TASK_UID"] = item.UniqueID;
            row["TASK_PCT_COMP"] = item.WorkCompletePercentage;
            row["PROJ_UID"] = guid.ToString();
        }
        private void AddItemToList()
        {

        }

        private static TaskItem BuildTaskItem(string dPath, SPListItem item)
        {
            return new TaskItem
            {
                ID = DataHelper.GetValueAsInteger(item.Title),
                UniqueID = DataHelper.GetValue(item[Constants.FieldId_UniqueID]),
                DrivingPath = dPath,
                Task = DataHelper.GetValue(item[Constants.FieldId_Task]),
                Duration = DataHelper.GetValue(item[Constants.FieldId_Duration]),
                Predecessor = DataHelper.GetValue(item[Constants.FieldId_Predecessor]),
                Start = DataHelper.GetValueAsDateTime(item[Constants.FieldId_Start]),
                Finish = DataHelper.GetValueAsDateTime(item[Constants.FieldId_Finish]),
                Deadline = DataHelper.GetValueAsDateTime(item[Constants.FieldId_Deadline]),
                ShowOn = DataHelper.GetValueFromMultiChoice(item[Constants.FieldId_ShowOn]),
                ModifiedOn = DataHelper.GetValueAsDateTime(item[Constants.FieldId_ModifiedOn]),
                WorkCompletePercentage = DataHelper.GetValueAsInteger(item[Constants.FieldId_PercentComplete])
            };
        }

        public static void UpdateTasksList(string serviceUrl, Guid projectUID)
        {
            DataRepository.ClearImpersonation();
            if (DataRepository.P14Login(serviceUrl))
            {
                DataAccess dataAccess = new Repository.DataAccess(projectUID);
                DataSet dataset = dataAccess.ReadProject(TaskItemRepository.GetProjectDataSetFromList());
                DataTable tasksDataTable = dataset.Tables["Task"];

                DataTable tDeletedRows = tasksDataTable.GetChanges(DataRowState.Deleted);

                var queryTasks = (from m in tasksDataTable.GetChanges(DataRowState.Added | DataRowState.Modified | DataRowState.Unchanged).AsEnumerable()
                                  select new TaskItem
                                  {
                                      ID = m.Field<int>("TASK_ID"),
                                      UniqueID = m.Field<Guid>("TASK_UID").ToString(),
                                      DrivingPath = m.Field<String>("TASK_DRIVINGPATH_ID"),
                                      Task = m.Field<String>("TASK_NAME"),
                                      Predecessor = m.Field<String>("TASK_PREDECESSORS"),
                                      Start = m.Field<DateTime?>("TASK_START_DATE"),
                                      Finish = m.Field<DateTime?>("TASK_FINISH_DATE"),
                                      Deadline = m.Field<DateTime?>("TASK_DEADLINE"),
                                      ShowOn = GetShownOnColumnValue(m.Field<String>("CUSTOMFIELD_DESC")),
                                      ModifiedOn = m.Field<DateTime?>("TASK_MODIFIED_ON"),
                                      WorkCompletePercentage = m.Field<int>("TASK_PCT_COMP")
                                  });

                var items = queryTasks.ToList();

                SPWeb web;
                if (SPContext.Current != null)
                    web = SPContext.Current.Web;
                else
                    web = new SPSite("http://finweb.contoso.com/sites/PMM").OpenWeb();

                var list = web.Lists.TryGetList(Constants.LIST_NAME_PROJECT_TASKS);
                var dPathsField = list.Fields[Constants.FieldId_DrivingPath] as SPFieldMultiChoice;
                var dShownField = list.Fields[Constants.FieldId_ShowOn] as SPFieldMultiChoice;
                var dIdField = list.Fields[Constants.FieldId_UniqueID] as SPFieldGuid;
                if (list != null)
                {
                    if (tDeletedRows != null && tDeletedRows.Rows.Count > 0)
                    {
                        foreach (DataRow row in tDeletedRows.Rows)
                        {
                            row.RejectChanges();
                            var q = new SPQuery();
                            q.Query += "<Where>" +
                                            "<Eq>" +
                                                "<FieldRef ID='" + dIdField.Id + "' />" +
                                                "<Value Type='Guid'>" + row["TASK_UID"].ToString() + "</Value>" +
                                            "</Eq>" +
                                //"<OrderBy>" +
                                //    "<FieldRef ID='" + startField.Id + "' Ascending='True' />" + 
                                //"</OrderBy>" +
                                        "</Where>";
                            list.GetItems(q)[0].Delete();
                            row.Delete();
                        }
                        list.Update();
                    }
                    
                   
                    foreach (TaskItem task in items)
                    {


                        var q = new SPQuery();
                        q.Query += "<Where>" +
                                        "<Eq>" +
                                            "<FieldRef ID='" + dIdField.Id + "' />" +
                                            "<Value Type='Text'>" + task.UniqueID + "</Value>" +
                                        "</Eq>" +
                            //"<OrderBy>" +
                            //    "<FieldRef ID='" + startField.Id + "' Ascending='True' />" + 
                            //"</OrderBy>" +
                                    "</Where>";
                        SPListItemCollection listItems = list.GetItems(q);
                        if (listItems.Count > 0)
                        {
                            var existingItem = listItems[0];
                            BuildListItem(existingItem, task);
                            existingItem.Update();
                        }
                        else
                        {
                            var newItem = list.Items.Add();
                            BuildListItem(newItem, task);
                            newItem.Update();
                        }

                        //web.AllowUnsafeUpdates = true;

                        //web.AllowUnsafeUpdates = false;
                    }
                }
                dPathsField.Update();
                dShownField.Update();
            }
        }

        private static void BuildListItem(SPListItem newItem, TaskItem task)
        {
            var dPathsField = newItem.ParentList.Fields[Constants.FieldId_DrivingPath] as SPFieldMultiChoice;
            var dShownField = newItem.ParentList.Fields[Constants.FieldId_ShowOn] as SPFieldMultiChoice;
            newItem[SPBuiltInFieldId.Title] = task.ID;
            newItem[Constants.FieldId_UniqueID] = task.UniqueID;
            SPFieldMultiChoiceValue value = GetDrivingPaths(task.DrivingPath) as SPFieldMultiChoiceValue;
            if (value != null)
            {
                for (int i = 0; i < value.Count; i++)
                {
                    if (!dPathsField.Choices.Contains(value[i]))
                    {
                        dPathsField.Choices.Add(value[i]);
                    }
                }
            }


            newItem[Constants.FieldId_DrivingPath] = GetDrivingPaths(task.DrivingPath);
            newItem[Constants.FieldId_Task] = task.Task;
            newItem[Constants.FieldId_Duration] = (task.Finish.HasValue && task.Start.HasValue) ? task.Finish.Value.Subtract(task.Start.Value).Days.ToString() + " days" : String.Empty;
            newItem[Constants.FieldId_Predecessor] = task.Predecessor;
            newItem[Constants.FieldId_Start] = task.Start;
            newItem[Constants.FieldId_Finish] = task.Finish;
            newItem[Constants.FieldId_Deadline] = task.Deadline;
            value = ConvertToMultiChoiceValue(task.ShowOn) as SPFieldMultiChoiceValue;
            if (value != null)
            {
                for (int i = 0; i < value.Count; i++)
                {
                    if (!dShownField.Choices.Contains(value[i]))
                    {
                        dShownField.Choices.Add(value[i]);
                    }
                }
            }
            newItem[Constants.FieldId_ShowOn] = ConvertToMultiChoiceValue(task.ShowOn);
            newItem[Constants.FieldId_ModifiedOn] = task.ModifiedOn;
            newItem[Constants.FieldId_PercentComplete] = task.WorkCompletePercentage;
        }

        private static object GetDrivingPaths(string dPaths)
        {
            var value = new SPFieldMultiChoiceValue();

            if (!String.IsNullOrEmpty(dPaths))
                foreach (string dPath in dPaths.Split(','))
                    value.Add(dPath);

            return value;
        }

        private static string[] GetShownOnColumnValue(string value)
        {
            string[] retVal = null;

            if (value != null)
            {
                var vList = new List<string>();
                var values = value.Split(',');
                foreach (string val in values)
                {
                    if (val.StartsWith("Show On"))
                        vList.Add(val.Split('_')[1]);
                }
                if (vList.Count > 0)
                    retVal = vList.ToArray();
            }

            return retVal;
        }

        private static object ConvertToMultiChoiceValue(string[] values)
        {
            var value = new SPFieldMultiChoiceValue();
            if (values != null)
                foreach (string val in values)
                    value.Add(val);

            return value;

        }
    }
}
