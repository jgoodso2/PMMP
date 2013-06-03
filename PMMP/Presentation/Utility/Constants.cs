using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PMMP
{
    /// <summary>
    /// 
    /// </summary>
    public class Constants
    {
        public const string DOCLIB_NAME_PMM_PRESENTATIONS = "PMM Presentations";

        public static readonly Guid FieldId_Comments = new Guid("{3836b42f-8de1-4e40-9a2b-8797cb18e102}");

        public static readonly Guid FieldId_Task = new Guid("{23042158-9c72-47ba-a676-0a9b59dd4a81}");
        public static readonly Guid FieldId_Duration = new Guid("{2a6ba979-b3f0-439b-987a-28e8e826c6af}");
        public static readonly Guid FieldId_Predecessor = new Guid("{7da94127-1926-42b4-8e7d-f37af75adb71}");
        public static readonly Guid FieldId_Start = new Guid("{00216713-71cf-41ba-b25d-c257a1815fb9}");
        public static readonly Guid FieldId_Finish = new Guid("{6bc6478d-af16-4026-babd-5c9eadc6f0e9}");
        public static readonly Guid FieldId_ModifiedOn = new Guid("{754c19ee-2a06-4185-b99e-df950aa43e4b}");
        public static readonly Guid FieldId_UniqueID = new Guid("{60797583-89d3-4f7a-9e74-b115c10d4b1b}");
        public static readonly Guid FieldId_PercentComplete = new Guid("{e33417bf-5996-40f5-9a30-e529a46c9f62}");
        public static readonly Guid FieldId_Deadline = new Guid("{73b8cc55-d228-43f6-9802-384e08fea888}");
        public static readonly Guid FieldId_DrivingPath = new Guid("{180b71da-55c1-45ee-9895-db22fd359391}");
        public static readonly Guid FieldId_ShowOn = new Guid("{64660977-9686-4f0b-a1da-ac9f6d591fbf}");

        public const string PROPERTY_NAME_DB_SERVICE_URL = "PMM-ServiceURL";
        public const string PROPERTY_NAME_DB_PROJECT_UID = "PMM-ProjectUID";

        public const string TEMPLATE_FILE_LOCATION = "_cts/PMM Presentation/PMM Template.pptx";

        public const string CT_PMM_NAME = "PMM Presentation";

        public const string LIST_NAME_PROJECT_TASKS = "Project Tasks";
    }
}
