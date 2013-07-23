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
using SvcCustomFields;
using System.Linq;


namespace PMMP
{
    public class TaskItemRepository
    {

        public static TaskGroupData GetTaskGroups(string projectUID)
        {
            Repository.Utility.WriteLog("GetTaskGroups started", System.Diagnostics.EventLogEntryType.Information);

            IList<TaskItemGroup> retVal = new List<TaskItemGroup>();
            CustomFieldDataSet customFieldDataSet = DataRepository.ReadCustomFields();

            DataAccess dataAccess = new Repository.DataAccess(new Guid(projectUID));
            DataSet dataset = dataAccess.ReadProject(null);
            ProjectDataSet ds = DataRepository.ReadProject(new Guid(projectUID));
            DataTable tasksDataTable = dataset.Tables["Task"];
            Dictionary<string, IList<TaskItem>> ChartsData = GetChartsData(tasksDataTable, customFieldDataSet);
            TaskGroupData taskData = new TaskGroupData();
            DateTime? projectStatusDate = GetProjectCurrentDate(new Guid(projectUID), ds);
            FiscalUnit fiscalPeriod = DataRepository.GetFiscalMonth(projectStatusDate);
            taskData.FiscalPeriod = fiscalPeriod;
            IList<TaskItemGroup> LateTasksData = GetLateTasksData(tasksDataTable, fiscalPeriod, customFieldDataSet);
            IList<TaskItemGroup> UpComingTasksData = GetupComingTasksData(tasksDataTable, fiscalPeriod, customFieldDataSet);
            taskData.TaskItemGroups = retVal;
            taskData.ChartsData = ChartsData;
            taskData.LateTaskGroups = LateTasksData;
            taskData.UpComingTaskGroups = UpComingTasksData;

            taskData.SPDLSTartToBL = GetSPDLSTartToBLData(new Guid(projectUID), ds);
            taskData.SPDLFinishToBL = GetSPDLFinishToBLData(new Guid(projectUID), ds);
            taskData.BEIData = GetBEIData(new Guid(projectUID), ds);
            if (tasksDataTable != null)
            {
                var dPaths = tasksDataTable.AsEnumerable().Where(t => !string.IsNullOrEmpty(t.Field<string>("TASK_DRIVINGPATH_ID"))).Select(t => t.Field<string>("TASK_DRIVINGPATH_ID")).Distinct();
                var chartTypes = tasksDataTable.AsEnumerable().Select(t => t.Field<string>("CUSTOMFIELD_DESC")).Distinct(); ;

                foreach (string dPath in dPaths)
                {
                    int taskCount = -1;
                    var taskItemGroup = new TaskItemGroup { DrivingPath = dPath, TaskItems = new List<TaskItem>() };
                    string previousTitle = string.Empty;
                    Dictionary<string, string> dictTitle = new Dictionary<string, string>();
                    int totalUnCompletedtaskCount = 0, totalCompletedTaskCount = 0;

                    List<TaskItem> chartItems = new List<TaskItem>();
                    List<TaskItemGroup> completedTasks = new List<TaskItemGroup>();
                    EnumerableRowCollection<DataRow> collection = tasksDataTable.AsEnumerable().Where(t => t.Field<string>("TASK_DRIVINGPATH_ID") != null && t.Field<string>("TASK_DRIVINGPATH_ID").Split(",".ToCharArray()).Contains(dPath));
                    int completedTaskCount = -1;
                    //DateTime? lastUpdate = GetLastUpdateDate();
                    TaskItemGroup completedTaskItemGroup = new TaskItemGroup { DrivingPath = dPath, TaskItems = new List<TaskItem>() };
                    foreach (DataRow item in collection)
                    {
                        if (item["TASK_DEADLINE"] != System.DBNull.Value && !string.IsNullOrEmpty(item["TASK_DEADLINE"].ToString()))
                        {
                            if (!dictTitle.ContainsKey(dPath.Split(",".ToCharArray())[0]))
                            {
                                dictTitle.Add(dPath.Split(",".ToCharArray())[0], item["TASK_NAME"].ToString());
                            }
                        }

                        if (item["CUSTOMFIELD_DESC"] != null)
                        {
                            chartItems.Add(BuildTaskItem(dPath, item, customFieldDataSet));
                        }

                        if (!string.IsNullOrEmpty(item["TASK_ACT_FINISH"].ToString()) && (Convert.ToDateTime(item["TASK_ACT_FINISH"].ToString())).InCurrentFiscalMonth(fiscalPeriod))
                        {
                            totalCompletedTaskCount++;
                            completedTaskCount++;
                            if (completedTaskCount == 10)
                            {
                                completedTasks.Add(completedTaskItemGroup);
                                completedTaskItemGroup = new TaskItemGroup { DrivingPath = dPath, TaskItems = new List<TaskItem>() };
                                completedTaskCount = 0;
                                completedTaskItemGroup.TaskItems.Add(BuildTaskItem(dPath, item, customFieldDataSet));
                            }
                            else
                            {
                                completedTaskItemGroup.TaskItems.Add(BuildTaskItem(dPath, item, customFieldDataSet));
                            }

                        }
                        else
                        {

                            if (item["TASK_PCT_COMP"] != null && (Convert.ToInt32(item["TASK_PCT_COMP"].ToString().Trim().Trim("%".ToCharArray()).Trim()) < 100))
                            {
                                totalUnCompletedtaskCount++;
                                taskCount++;
                                if (taskCount == 10)
                                {
                                    retVal.Add(taskItemGroup);
                                    taskItemGroup = new TaskItemGroup { DrivingPath = dPath, TaskItems = new List<TaskItem>() };

                                    taskItemGroup.Title = previousTitle;
                                    taskCount = 0;
                                    taskItemGroup.TaskItems.Add(BuildTaskItem(dPath, item, customFieldDataSet));

                                }
                                else
                                {
                                    taskItemGroup.TaskItems.Add(BuildTaskItem(dPath, item, customFieldDataSet));
                                }
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
                        taskItemGroup.Charts = new string[chartTypes.Count()];
                        chartTypes.ToList().CopyTo(taskItemGroup.Charts, 0);
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
            Repository.Utility.WriteLog("GetTaskGroups completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return taskData;
        }



        private static List<GraphDataGroup> GetSPDLSTartToBLData(Guid projectUID, ProjectDataSet projectDataSet)
        {
            List<GraphDataGroup> group = new List<GraphDataGroup>();
            DateTime? projectStatusDate = GetProjectCurrentDate(projectUID, projectDataSet);
            if (!projectStatusDate.HasValue)
                return new List<GraphDataGroup>();
            List<FiscalUnit> projectStatusPeriods = GetProjectStatusPeriods(projectStatusDate.Value);
            IEnumerable<ProjectDataSet.TaskRow> tasks = projectDataSet.Task.Where(t => t.TASK_IS_SUMMARY == false && !t.IsTASK_DURNull() && t.TASK_DUR > 0);


            //Get CS Data
            List<GraphData> graphDataCS = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTASK_ACT_STARTNull() && t.TASK_ACT_START >= unit.From && t.TASK_ACT_START <= unit.To);
                graphDataCS.Add(new GraphData() { Count = count, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "CS", Data = graphDataCS });

            //Get FCS Data
            List<GraphData> graphDataFCS = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTASK_START_DATENull() && t.TASK_START_DATE >= unit.From && t.TASK_START_DATE <= unit.To && t.IsTASK_ACT_STARTNull() && t.TASK_PCT_COMP == 0);
                graphDataFCS.Add(new GraphData() { Count = count, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "FCS", Data = graphDataFCS });

            //Get DQ Data
            List<GraphData> graphDataDQ = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTASK_START_DATENull() && !t.IsTB_STARTNull() && t.TB_START >= unit.From && t.TB_START <= unit.To && t.TASK_START_DATE > t.TB_START);
                graphDataDQ.Add(new GraphData() { Count = count, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "DQ", Data = graphDataDQ });

            //Get FDQ Data
            List<GraphData> graphDataFDQ = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTASK_START_DATENull() && !t.IsTB_STARTNull() && t.TB_START >= unit.From && t.TB_START <= unit.To && t.TASK_START_DATE > t.TB_START && t.TASK_PCT_COMP == 0);
                graphDataFDQ.Add(new GraphData() { Count = count, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "FDQ", Data = graphDataFDQ });


            //Get CDQ Data
            List<GraphData> graphDataCDQ = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTB_STARTNull() && t.TB_START <= projectStatusDate && t.TB_START >= unit.From && t.TB_START <= unit.To && t.IsTASK_ACT_STARTNull());
                graphDataCDQ.Add(new GraphData() { Count = 0, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "CDQ", Data = graphDataCDQ });

            //Get FCDQ Data
            List<GraphData> graphDataFCDQ = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTASK_START_DATENull() && !t.IsTB_STARTNull() && t.TASK_START_DATE > projectStatusDate && t.TASK_START_DATE >= unit.From && t.TASK_START_DATE <= unit.To && t.TASK_START_DATE > t.TB_START);
                graphDataFCDQ.Add(new GraphData() { Count = 0, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "FCDQ", Data = graphDataFCDQ });

            return group;
        }

        private static List<GraphDataGroup> GetSPDLFinishToBLData(Guid projectUID, ProjectDataSet projectDataSet)
        {
            List<GraphDataGroup> group = new List<GraphDataGroup>();
            DateTime? projectStatusDate = GetProjectCurrentDate(projectUID, projectDataSet);
            List<FiscalUnit> projectStatusPeriods = GetProjectStatusPeriods(projectStatusDate);
            IEnumerable<ProjectDataSet.TaskRow> tasks = projectDataSet.Task.Where(t => t.TASK_IS_SUMMARY == false && !t.IsTASK_DURNull() && t.TASK_DUR > 0);


            //Get CS Data
            List<GraphData> graphDataCS = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTASK_ACT_FINISHNull() && t.TASK_ACT_FINISH >= unit.From && t.TASK_ACT_FINISH <= unit.To);
                graphDataCS.Add(new GraphData() { Count = count, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "CF", Data = graphDataCS });

            //Get FCS Data
            List<GraphData> graphDataFCS = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTASK_FINISH_DATENull() && t.TASK_FINISH_DATE >= unit.From && t.TASK_FINISH_DATE <= unit.To && t.IsTASK_ACT_FINISHNull() && t.TASK_PCT_COMP < 100);
                graphDataFCS.Add(new GraphData() { Count = count, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "FCF", Data = graphDataFCS });

            //Get DQ Data
            List<GraphData> graphDataDQ = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTASK_FINISH_DATENull() && !t.IsTB_FINISHNull() && t.TB_FINISH >= unit.From && t.TB_FINISH <= unit.To && t.TASK_FINISH_DATE > t.TB_FINISH);
                graphDataDQ.Add(new GraphData() { Count = count, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "DQF", Data = graphDataDQ });

            //Get FDQ Data
            List<GraphData> graphDataFDQ = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTASK_FINISH_DATENull() && t.TASK_FINISH_DATE > projectStatusDate && !t.IsTB_FINISHNull() && t.TB_FINISH >= unit.From && t.TB_FINISH <= unit.To && t.TASK_FINISH_DATE > t.TB_FINISH && t.TASK_PCT_COMP < 100);
                graphDataFDQ.Add(new GraphData() { Count = count, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "FDQF", Data = graphDataFDQ });


            //Get CDQ Data
            List<GraphData> graphDataCDQ = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTB_FINISHNull() && t.TB_FINISH <= projectStatusDate && t.TB_FINISH >= unit.From && t.TB_FINISH <= unit.To && t.IsTASK_ACT_FINISHNull());
                graphDataCDQ.Add(new GraphData() { Count = 0, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "CDQF", Data = graphDataCDQ });

            //Get FCDQ Data
            List<GraphData> graphDataFCDQ = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int count = tasks.Count(t => !t.IsTASK_FINISH_DATENull() && !t.IsTB_FINISHNull() && t.TASK_FINISH_DATE > projectStatusDate && t.TASK_FINISH_DATE >= unit.From && t.TASK_FINISH_DATE <= unit.To && t.TASK_FINISH_DATE > t.TB_FINISH);
                graphDataFCDQ.Add(new GraphData() { Count = 0, Title = unit.GetTitle() });
            }
            group.Add(new GraphDataGroup() { Type = "FCDQF", Data = graphDataFCDQ });

            return group;
        }

        private static List<GraphDataGroup> GetBEIData(Guid projectUID, ProjectDataSet projectDataSet)
        {
            List<GraphDataGroup> group = new List<GraphDataGroup>();
            DateTime? projectStatusDate = GetProjectCurrentDate(projectUID, projectDataSet);
            List<FiscalUnit> projectStatusPeriods = GetProjectStatusWeekPeriods(projectStatusDate);
            IEnumerable<ProjectDataSet.TaskRow> tasks = projectDataSet.Task.Where(t => t.TASK_IS_SUMMARY == false && !t.IsTASK_DURNull() && t.TASK_DUR > 0);


            //Get BEIStart Data
            List<GraphData> graphDataBES = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {

                int totalStart = tasks.Count(t => !t.IsTASK_ACT_STARTNull() && t.TASK_ACT_START >= unit.From && t.TASK_ACT_START <= unit.To);
                int totalTBStart = tasks.Count(t => !t.IsTB_STARTNull() && t.TB_START >= unit.From && t.TB_START <= unit.To);
                if (totalTBStart != 0)
                {
                    graphDataBES.Add(new GraphData() { Count = totalStart / totalTBStart, Title = unit.GetTitle() });
                }
                else
                {
                    graphDataBES.Add(new GraphData() { Count = 0, Title = unit.GetTitle() });
                }
            }
            group.Add(new GraphDataGroup() { Type = "BES", Data = graphDataBES });

            //Get BEIFinish Data
            List<GraphData> graphDataBEF = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {

                int totalFinish = tasks.Count(t => !t.IsTASK_ACT_FINISHNull() && t.TASK_ACT_FINISH >= unit.From && t.TASK_ACT_FINISH <= unit.To);
                int totalTBFinish = tasks.Count(t => !t.IsTB_FINISHNull() && t.TB_FINISH >= unit.From && t.TB_FINISH <= unit.To);
                if (totalTBFinish != 0)
                {
                    graphDataBEF.Add(new GraphData() { Count = totalFinish / totalTBFinish, Title = unit.GetTitle() });
                }
                else
                {
                    graphDataBEF.Add(new GraphData() { Count = 0, Title = unit.GetTitle() });
                }
            }
            group.Add(new GraphDataGroup() { Type = "BEF", Data = graphDataBEF });

            //Get BEI Forecast Start Data
            List<GraphData> graphDataBEFS = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {
                int totalStart = tasks.Count(t => !t.IsTASK_START_DATENull() && !t.IsTASK_PCT_COMPNull() && t.TASK_START_DATE >= unit.From && t.TASK_START_DATE <= unit.To && t.TASK_PCT_COMP == 0);
                int totalTBStart = tasks.Count(t => !t.IsTB_STARTNull() && t.TB_START >= unit.From && t.TB_START <= unit.To);
                if (totalTBStart != 0)
                {
                    graphDataBEFS.Add(new GraphData() { Count = totalStart / totalTBStart, Title = unit.GetTitle() });
                }
                else
                {
                    graphDataBEFS.Add(new GraphData() { Count = 0, Title = unit.GetTitle() });
                }
            }
            group.Add(new GraphDataGroup() { Type = "BEFS", Data = graphDataBEFS });

            //Get BEI Forecast Finish Data
            List<GraphData> graphDataBEFF = new List<GraphData>();
            foreach (FiscalUnit unit in projectStatusPeriods)
            {

                int totalFinish = tasks.Count(t => !t.IsTASK_FINISH_DATENull() && !t.IsTASK_PCT_COMPNull() && t.TASK_FINISH_DATE >= unit.From && t.TASK_FINISH_DATE <= unit.To);
                int totalTBFinish = tasks.Count(t => !t.IsTB_FINISHNull() && t.TB_FINISH >= unit.From && t.TB_FINISH <= unit.To);
                if (totalTBFinish != 0)
                {
                    graphDataBEFF.Add(new GraphData() { Count = totalFinish / totalTBFinish, Title = unit.GetTitle() });
                }
                else
                {
                    graphDataBEFF.Add(new GraphData() { Count = 0, Title = unit.GetTitle() });
                }
            }
            group.Add(new GraphDataGroup() { Type = "BEFF", Data = graphDataBEFF });


            return group;
        }

        private static List<FiscalUnit> GetProjectStatusWeekPeriods(DateTime? projectStatusDate)
        {
            DataAccess da = new DataAccess(Guid.Empty);
            return da.GetProjectStatusWeekPeriods(projectStatusDate);
        }
        private static List<FiscalUnit> GetProjectStatusPeriods(DateTime? date)
        {
            DataAccess da = new DataAccess(Guid.Empty);
            return da.GetProjectStatusPeriods(date);
        }

        private static DateTime? GetProjectCurrentDate(Guid projectUID, ProjectDataSet projectDataSet)
        {
            DataAccess da = new DataAccess(projectUID);
            return da.GetProjectCurrentDate(projectDataSet, projectUID);
        }
        private static IList<TaskItemGroup> GetLateTasksData(DataTable tasksDataTable, FiscalUnit month, CustomFieldDataSet dataSet)
        {
            if (month.From == DateTime.MinValue && month.To == DateTime.MaxValue)
                return new List<TaskItemGroup>();
            Repository.Utility.WriteLog("GetLateTasksData started", System.Diagnostics.EventLogEntryType.Information);
            int count = -1;
            int lateTaskCount = 0;
            IList<TaskItemGroup> retVal = new List<TaskItemGroup>();


            TaskItemGroup taskData = new TaskItemGroup() { TaskItems = new List<TaskItem>() };
            IList<TaskItem> items = new List<TaskItem>();
            EnumerableRowCollection<DataRow> collection =

                tasksDataTable.AsEnumerable()
                .Where((t => t.Field<bool>("TASK_IS_SUMMARY") == false && t.Field<int>("TASK_PCT_COMP") == 0 && t.Field<DateTime?>("TASK_START_DATE").HasValue && t.Field<DateTime?>("TB_START").HasValue && t.Field<DateTime?>("TB_START").Value.InCurrentFiscalMonth(month) &&
                       t.Field<int>("TASK_PCT_COMP") < 100 &&
                       t.Field<DateTime?>("TASK_START_DATE").Value.Date > t.Field<DateTime?>("TB_START").Value.Date));

            List<DataRow> mergedCollection = collection.Union(tasksDataTable.AsEnumerable()
           .Where(t => t.Field<bool>("TASK_IS_SUMMARY") == false && t.Field<DateTime?>("TASK_FINISH_DATE").HasValue && t.Field<DateTime?>("TB_FINISH").HasValue && t.Field<DateTime?>("TB_FINISH").Value.InCurrentFiscalMonth(month) &&
                  t.Field<int>("TASK_PCT_COMP") < 100 &&
                  t.Field<DateTime?>("TASK_FINISH_DATE").Value.Date > t.Field<DateTime?>("TB_FINISH").Value.Date)
                  ).ToList();
            foreach (DataRow item in mergedCollection)
            {
                count++;
                lateTaskCount++;
                TaskItem taskItem = BuildTaskItem("", item, dataSet);
                if (count == 10)
                {
                    retVal.Add(taskData);
                    taskData = new TaskItemGroup { TaskItems = new List<TaskItem>() };
                    count = -1;
                    taskData.TaskItems.Add(BuildTaskItem("", item, dataSet));
                }
                else
                {
                    taskData.TaskItems.Add(BuildTaskItem("", item, dataSet));
                }
            }

            if (count % 10 != 0)
            {
                retVal.Add(taskData);

            }
            Repository.Utility.WriteLog("GetLateTasksData completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return retVal;
        }

        private static IList<TaskItemGroup> GetupComingTasksData(DataTable tasksDataTable, FiscalUnit month, CustomFieldDataSet dataSet)
        {
            if (month.From == DateTime.MinValue && month.To == DateTime.MaxValue)
                return new List<TaskItemGroup>();

            FiscalUnit fiscalUnit = new FiscalUnit() { From = month.From, To = month.To.AddMonths(1) };

            Repository.Utility.WriteLog("GetupComingTasksData started", System.Diagnostics.EventLogEntryType.Information);
            int count = -1;
            int upComingTaskCount = 0;
            IList<TaskItemGroup> retVal = new List<TaskItemGroup>();


            TaskItemGroup taskData = new TaskItemGroup() { TaskItems = new List<TaskItem>() };
            IList<TaskItem> items = new List<TaskItem>();
            EnumerableRowCollection<DataRow> collection =

                tasksDataTable.AsEnumerable()
                .Where((t => t.Field<bool>("TASK_IS_SUMMARY") == false && t.Field<int>("TASK_PCT_COMP") < 100
                    && t.Field<DateTime?>("TASK_FINISH_DATE").HasValue && t.Field<DateTime?>("TASK_FINISH_DATE").Value.InCurrentFiscalMonth(fiscalUnit)
                       )).OrderBy(t => t.Field<int>("TASK_ID"));


            foreach (DataRow item in collection)
            {
                count++;
                upComingTaskCount++;
                TaskItem taskItem = BuildTaskItem("", item, dataSet);
                if (count == 10)
                {
                    retVal.Add(taskData);
                    taskData = new TaskItemGroup { TaskItems = new List<TaskItem>() };
                    count = -1;
                    taskData.TaskItems.Add(BuildTaskItem("", item, dataSet));
                }
                else
                {
                    taskData.TaskItems.Add(BuildTaskItem("", item, dataSet));
                }
            }

            if (count % 10 != 0)
            {
                retVal.Add(taskData);

            }
            Repository.Utility.WriteLog("GetupComingTasksData completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return retVal;
        }

        private static Dictionary<string, IList<TaskItem>> GetChartsData(DataTable tasksDataTable, CustomFieldDataSet dataSet)
        {
            Repository.Utility.WriteLog("GetLateTasksData started", System.Diagnostics.EventLogEntryType.Information);
            Dictionary<string, IList<TaskItem>> chartsData = new Dictionary<string, IList<TaskItem>>();


            var chartTypes = tasksDataTable.AsEnumerable().Select(t => t.Field<string>("CUSTOMFIELD_DESC")).Distinct();
            foreach (string chartType in chartTypes)
            {
                if (!string.IsNullOrEmpty(chartType))
                {
                    foreach (string chartTypeItem in chartType.Split(",".ToCharArray()))
                    {
                        IList<TaskItem> items = new List<TaskItem>();
                        EnumerableRowCollection<DataRow> collection = tasksDataTable.AsEnumerable().Where(t => t.Field<string>("CUSTOMFIELD_DESC") != null && t.Field<string>("CUSTOMFIELD_DESC").Split(",".ToCharArray()).Contains(chartTypeItem)).OrderBy(t => t.Field<int>("TASK_ID")).OrderByDescending(t=>t.Field<int>("TASK_ID"));
                        foreach (DataRow item in collection)
                        {
                            TaskItem taskItem = BuildTaskItem("", item, dataSet);
                            items.Add(taskItem);
                        }
                        if (items.Count > 0)
                        {
                            if (!chartsData.ContainsKey(chartTypeItem))
                            {
                                chartsData.Add(chartTypeItem, items);
                            }
                        }
                    }
                }
            }
            Repository.Utility.WriteLog("GetChartsData completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return chartsData;
        }

        private static TaskItem BuildTaskItem(string dPath, DataRow item, CustomFieldDataSet dataSet)
        {
            DateTime? estFinish = (DateTime?)DataHelper.GetValueFromCustomFieldTextOrDate(item, CustomFieldType.EstFinish, dataSet);
            DateTime? estStart = (DateTime?)DataHelper.GetValueFromCustomFieldTextOrDate(item, CustomFieldType.EstStart, dataSet);
            object objreason = DataHelper.GetValueFromCustomFieldTextOrDate(item, CustomFieldType.ReasonRecovery, dataSet);
            string reasonrecovery = objreason != null ? objreason.ToString() : "";
            return new TaskItem
            {
                ID = DataHelper.GetValueAsInteger(item["TASK_ID"].ToString()),
                UniqueID = DataHelper.GetValue(item["TASK_UID".ToString()]),
                DrivingPath = dPath,
                Task = DataHelper.GetValue(item["TASK_NAME"].ToString()),
                Duration = DataHelper.GetValue(item["TASK_DUR"].ToString()),
                Predecessor = DataHelper.GetValue(item["TASK_PREDECESSORS"].ToString()),
                Start = DataHelper.GetValueAsDateTime(item["TASK_START_DATE"].ToString()),
                Finish = DataHelper.GetValueAsDateTime(item["TASK_FINISH_DATE"].ToString()),
                Deadline = DataHelper.GetValueAsDateTime(item["TASK_DEADLINE"].ToString()),
                ShowOn = DataHelper.GetValueFromMultiChoice(item["CUSTOMFIELD_DESC"].ToString(), CustomFieldType.ShowOn),
                CA = string.Join(",", DataHelper.GetValueFromMultiChoice(item["CUSTOMFIELD_DESC"].ToString(), CustomFieldType.CA)),
                EstFinish = estFinish,
                EstStart = estStart,
                PMT = string.Join(",", DataHelper.GetValueFromMultiChoice(item["CUSTOMFIELD_DESC"].ToString(), CustomFieldType.PMT)),
                ReasonRecovery = reasonrecovery,
                ModifiedOn = DataHelper.GetValueAsDateTime(item["TASK_MODIFIED_ON"].ToString()),
                WorkCompletePercentage = DataHelper.GetValueAsInteger(item["TASK_PCT_COMP"].ToString()),
                TotalSlack = DataHelper.GetValue(item["TASK_TOTAL_SLACK"].ToString()),
                BaseLineStart = DataHelper.GetValueAsDateTime(item["TB_START"].ToString()),
                BaseLineFinish = DataHelper.GetValueAsDateTime(item["TB_FINISH"].ToString()),
                Hours = DataHelper.GetValue(item["TASK_WORK"].ToString()),
                BLDuration = DataHelper.GetValue(item["TB_DUR"].ToString())
            };
        }
    }
}
