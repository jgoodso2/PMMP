﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using Microsoft.SharePoint;

namespace PMMP
{
    public class TaskItemGroup
    {
        string _title;
        public string Title
        {
            get
            {
                var retVal = this._title;
                if (this._title.EndsWith("Complete"))
                    retVal = String.Format("Driving Path: {0}", this._title);

                return retVal;
            }
            set { this._title = value; }
        }
        public string DrivingPath { get; set; }
        public string[] Charts { get; set; }
        public IList<TaskItem> TaskItems { get; set; }
        public IList<TaskItem> ChartTaskItems { get; set; }
        public IList<TaskItemGroup> CompletedTaskgroups { get; set; }
        public DataTable TaskItemsDataTable
        {
            get { return this.ToDataTable(this.TaskItems, SlideType.Grid); }
        }

        public DataTable GetChartDataTable(string chartName)
        {
            IList<TaskItem> tasks = new List<TaskItem>();
            string[] values = chartName.Split(",".ToCharArray());
            for (int i = 0; i < values.Count(); i++)
            {
                var items = this.ChartTaskItems.Where(x => x.ShowOn != null && x.ShowOn.Contains(values[i])).ToList();
                tasks =  tasks.Union(items.AsEnumerable()).ToList();
            }

            if (tasks.Count > 0)
                return this.ToDataTable(tasks, SlideType.Chart);
            else
                return null;
        }

        private DataTable ToDataTable(IList<TaskItem> data, SlideType type)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(TaskItem));
            DataTable table = new DataTable();
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                if (IsValidColumn(prop.Name, type))
                    table.Columns.Add(this.GetColumnName(prop.Name), GetColumnType(prop.PropertyType));
            }

            object[] values = new object[table.Columns.Count];

            foreach (TaskItem item in data)
            {
                var index = 0;
                for (int i = 0; i < props.Count; i++)
                {
                    if (IsValidColumn(props[i].Name, type))
                    {
                        values[index] = this.GetValue(props[i], item, type);
                        index++;
                    }
                }

                table.Rows.Add(values);
            }
            return table;
        }

        private bool IsValidColumn(string name, SlideType type)
        {
            bool retVal = true;

            if (type == SlideType.Grid && name == "Deadline" || name == "ShowOn")
                retVal = false;
            else if (name != "Task" && name != "Finish")
            {
                retVal = false;
            }

            return retVal;
        }

        private string GetColumnName(string propName)
        {
            string retVal = propName;

            switch (propName)
            {
                case "ID":
                    retVal = "ID_";
                    break;
                case "UniqueID":
                    retVal = "UniqueID";
                    break;
                case "DrivingPath":
                    retVal = "Driving Path";
                    break;
                case "Task":
                    retVal = "Task";
                    break;
                case "Duration":
                    retVal = "Duration";
                    break;
                case "Predecessor":
                    retVal = "Predecessor";
                    break;
                case "Start":
                    retVal = "Start";
                    break;
                case "Finish":
                    retVal = "Finish";
                    break;
                case "Deadline":
                    retVal = "Deadline";
                    break;
                case "ShowOn":
                    retVal = "Show On";
                    break;
                default:
                    break;
            }

            return retVal;
        }

        private Type GetColumnType(Type type)
        {
            Type retVal = type;

            if (type == typeof(DateTime?))
                retVal = typeof(DateTime);

            return retVal;
        }

        private object GetValue(PropertyDescriptor prop, TaskItem item, SlideType type)
        {
            object retVal = prop.GetValue(item);

            if (type == SlideType.Chart && prop.Name == "Task")
                retVal = String.Format("{0}: {1}", item.Task, item.Finish.HasValue ? item.Finish.Value.ToString("MM/dd") : String.Empty);

            return retVal;
        }
    }
}
