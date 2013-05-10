using System;
using System.Linq;
using Microsoft.SharePoint;
using System.IO;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;


namespace PMMPresentation.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //TaskItemRepository.UpdateTasksList("http://intranet.contoso.com/projectserver5", new Guid("{4c0680a4-5074-4c7e-b198-c5d30aae83bf}"));
                //Run();
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(String.Format("An exception has ocurred. Message: {0}. Stack Trace:{1}", ex.Message, ex.StackTrace));
                System.Console.WriteLine("Press any key to exit");
                System.Console.Read();
            }
        }

        private static void Run()
        {
            var filePath = String.Format(@"{0}\Templates\{1}", Directory.GetCurrentDirectory().Replace(@"bin\Debug", String.Empty), Configuration.TemplateFile);

            var fs = File.OpenRead(filePath);
            var bytes = new byte[fs.Length];
            fs.Read(bytes, 0, bytes.Length);
            fs.Close();

            var ms = PresentationManager.CreateStatusReport(bytes);

            var newFilePath = String.Format(@"{0}\Output\Result.pptx", Directory.GetCurrentDirectory().Replace(@"bin\Debug", String.Empty));

            var file = new FileStream(newFilePath, FileMode.Create, System.IO.FileAccess.Write);
            ms.WriteTo(file);
            file.Close();
            ms.Close();
        }
    }
}
