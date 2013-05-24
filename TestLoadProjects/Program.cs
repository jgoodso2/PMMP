using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;


namespace TestLoadProjects
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                if (args.Length < 1)
                {
                    Console.WriteLine("Please specify project server url as command line argument");
                    return;
                }
                Console.WriteLine("Started----");
                
                    if (!String.IsNullOrEmpty(args[0]))
                    {
                        
                        if (Repository.DataRepository.P14Login(args[0]))
                        {
                            var projectList = Repository.DataRepository.ReadProjectsList();
                            if (projectList.Tables["Project"] != null)
                            {
                                Console.WriteLine("No of projects on the server={0}", projectList.Tables["Project"].Rows.Count.ToString());
                               // Console.WriteLine("Name of project on the server={0},{1}", projectList.Tables["Project"].Rows[0][0].ToString(), projectList.Tables["Project"].Rows[0][1].ToString());
                                Console.WriteLine("No of tables in returned data set = {0}: ", projectList.Tables.Count.ToString());
                            }
                            else
                            {
                                Console.Write(" No Projects returned in the Dataset");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Unable to login to the server");
                        }
                    }
              

            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occured and the exception message ={0}", ex.Message);
            }
            Console.ReadKey();
        }
    }

}
