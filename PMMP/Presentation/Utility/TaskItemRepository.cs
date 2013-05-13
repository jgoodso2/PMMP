﻿using System;
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
       
        public static TaskGroupData GetTaskGroups(string projectUID)
        {
            TaskGroupData taskData = new TaskGroupData();
            
            IList<TaskItemGroup> retVal = new List<TaskItemGroup>();
          
            
            DataAccess dataAccess = new Repository.DataAccess(new Guid(projectUID));
            DataSet dataset = dataAccess.ReadProject(null);
            DataTable tasksDataTable = dataset.Tables["Task"];
            Dictionary<string, IList<TaskItem>> ChartsData = GetChartsData(tasksDataTable);
            taskData.TaskItemGroups = retVal;
            taskData.ChartsData = ChartsData;
            if (tasksDataTable != null)
            {
                var dPaths = tasksDataTable.AsEnumerable().Select(t => t.Field<string>("TASK_DRIVINGPATH_ID")).Distinct();
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
                    EnumerableRowCollection<DataRow> collection = tasksDataTable.AsEnumerable().Where(t=>t.Field<string>("TASK_DRIVINGPATH_ID") != null && t.Field<string>("TASK_DRIVINGPATH_ID").Split(",".ToCharArray()).Contains(dPath));
                    int completedTaskCount = -1;
                    //DateTime? lastUpdate = GetLastUpdateDate();
                    TaskItemGroup completedTaskItemGroup = new TaskItemGroup { DrivingPath = dPath, TaskItems = new List<TaskItem>() };
                    foreach (DataRow item in collection)
                    {
                        if (item["TASK_DEADLINE"] != null)
                        {
                            if (!dictTitle.ContainsKey(dPath.Split(",".ToCharArray())[0]))
                            {
                                dictTitle.Add(dPath.Split(",".ToCharArray())[0], item["TASK_NAME"].ToString());
                            }
                        }

                        if (item["CUSTOMFIELD_DESC"] != null)
                        {
                            chartItems.Add(BuildTaskItem(dPath, item));
                        }

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
            return taskData;
        }

        private static Dictionary<string, IList<TaskItem>> GetChartsData(DataTable tasksDataTable)
        {
            Dictionary<string, IList<TaskItem>> chartsData = new Dictionary<string, IList<TaskItem>>() ;


            var chartTypes = tasksDataTable.AsEnumerable().Select(t => t.Field<string>("CUSTOMFIELD_DESC")).Distinct(); 
                foreach (string chartType in chartTypes)
                {
                    IList<TaskItem> items = new List<TaskItem>();
                    EnumerableRowCollection<DataRow> collection = tasksDataTable.AsEnumerable().Where(t => t.Field<string>("CUSTOMFIELD_DESC") != null && t.Field<string>("CUSTOMFIELD_DESC").Split(",".ToCharArray()).Contains(chartType));
                    foreach (DataRow item in collection)
                    {
                        TaskItem taskItem =  BuildTaskItem("", item);
                        items.Add(taskItem);
                    }
                    if (items.Count > 0)
                    {
                        chartsData.Add(chartType, items);
                    }
            }
            return chartsData;
        }

        private static TaskItem BuildTaskItem(string dPath, DataRow item)
        {
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
                ShowOn = DataHelper.GetValueFromMultiChoice(item["CUSTOMFIELD_DESC"].ToString()),
                ModifiedOn = DataHelper.GetValueAsDateTime(item["TASK_MODIFIED_ON"].ToString()),
                WorkCompletePercentage = DataHelper.GetValueAsInteger(item["TASK_PCT_COMP"].ToString())
            };
        }
    }
}