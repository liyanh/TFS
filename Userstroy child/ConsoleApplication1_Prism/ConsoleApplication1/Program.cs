using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.VersionControl.Client;
using System.Net;
using System.Collections;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.Framework.Client;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.Data;

namespace ConsoleApplication1
{
    class Program
    {
        static string fileName = "";
        static VersionControlServer version = null;
       
        static void Main(string[] args)
        {                       
            Console.WriteLine("123!");
            System.Console.Write("Input FileName: ");
            fileName = Console.ReadLine();

            //read excel file-------------------------------
            Console.WriteLine("Reading File...");
            string path = Environment.CurrentDirectory + "\\";
            path = path + fileName + ".xlsx";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Macro;HDR=YES;IMEX=1'";   
            OleDbConnection conn = new OleDbConnection(strConn); 
            conn.Open();   
            string strExcel = "";    
            OleDbDataAdapter myCommand = null; 
            DataTable ds = null; 
            strExcel="select Id from [sheet1$]"; 
            myCommand = new OleDbDataAdapter(strExcel, strConn); 
            ds = new DataTable(); 
            myCommand.Fill(ds);           
            conn.Close();            

            //创建文件 
            //StreamWriter sw = File.AppendText(System.Environment.CurrentDirectory + "\\AnalysisResult.csv");           
            FileStream fs = new FileStream(Environment.CurrentDirectory + "\\" + fileName + "_Child.csv", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
            fs.SetLength(0);
            StreamWriter sw = new StreamWriter(fs);
            
            //服务器连接，相关对象的获取
            TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri("https://tfs.slb.com/tfs/SLB1/"));
            //版本控制
            version = tpc.GetService(typeof(VersionControlServer)) as VersionControlServer;
            WorkItemStore workItemStore = tpc.GetService<WorkItemStore>();

            // get child---------------------------
            sw.WriteLine(" ID,Description,total completed effort,Children");
            foreach (DataRow dr in ds.Rows)
            {
                if (dr[0] != DBNull.Value)
                {
                    WorkItem wi = workItemStore.GetWorkItem(Convert.ToInt32(dr[0]));
                    Console.WriteLine("Id: " + wi.Id);
                    string description = wi.Description;
                    description = description.Replace(",", ".");
                    int comEffort = 0;
                    StringBuilder childId = new StringBuilder();
                    if ("Task".Equals(wi.Type.Name))
                    {
                        int wiRevs = wi.Revision;
                        Revision wiRev = wi.Revisions[wiRevs - 1];
                        childId = childId.Append("");
                        comEffort = comEffort + Convert.ToInt32(wiRev.Fields["Completed work"].Value);
                    }
                    else
                    {
                        WorkItemLinkCollection linkWorkItem = wi.WorkItemLinks;
                        foreach (WorkItemLink wil in linkWorkItem)
                        {

                            if (wil.LinkTypeEnd.Name.Equals("Child"))
                            {
                                WorkItem childWI = workItemStore.GetWorkItem(wil.TargetId);
                                int childRevs = childWI.Revision;
                                Revision childRev = childWI.Revisions[childRevs - 1];
                                childId = childId.Append((childWI.Id).ToString()).Append(",");
                                if ("Task".Equals(childWI.Type.Name))
                                {
                                    comEffort = comEffort + Convert.ToInt32(childRev.Fields["Completed work"].Value);
                                }
                                else
                                {
                                    comEffort = comEffort + getChildEffort(childWI, workItemStore);
                                }
                            }
                        }
                    }
                    sw.WriteLine("" + wi.Id + ',' + description + ',' + comEffort + ',' + childId);
                }
            }                  
            Console.WriteLine("***********************try end**********************");
            sw.Close();
            fs.Close();
            Thread.Sleep(5000);
        }

        public static int getChildEffort(WorkItem wi, WorkItemStore workItemStore)
        {
            WorkItemLinkCollection linkWorkItem = wi.WorkItemLinks;
            int comEffort = 0;
            StringBuilder childId = new StringBuilder();
            foreach (WorkItemLink wil in linkWorkItem)
            {

                if (wil.LinkTypeEnd.Name.Equals("Child"))
                {
                    WorkItem childWI = workItemStore.GetWorkItem(wil.TargetId);
                    childId = childId.Append((childWI.Id).ToString()).Append("*");
                    if ("Task".Equals(childWI.Type.Name))
                    {
                        int childRevs = childWI.Revision;
                        Revision childRev = childWI.Revisions[childRevs - 1];
                        comEffort = comEffort + Convert.ToInt32(childRev.Fields["Completed work"].Value);
                    }
                    else
                    {
                        comEffort = comEffort + getChildEffort(childWI,workItemStore);
                    }
                }
            }
            return comEffort;
        }

    }

}
