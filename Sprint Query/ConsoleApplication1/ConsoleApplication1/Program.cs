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
namespace ConsoleApplication1
{
    class Program
    {
        static string sprintNum = "";
        static string tpName = "";
        static string tpURI = "";
        static string queryStr = "";
        static VersionControlServer version = null;
       
        static void Main(string[] args)
        {                       
            Console.WriteLine("123!");
            while (true)
            {
                System.Console.Write("Input IA/Prism: ");
                tpName = Console.ReadLine();
                System.Console.Write("Input SprintNum: ");
                sprintNum = Console.ReadLine();
                if ("IA".Equals(tpName, StringComparison.OrdinalIgnoreCase))
                {
                    tpURI = "https://tfs.slb.com/tfs/Real-Time_Collection/";
                    queryStr = "SELECT * FROM WorkItems WHERE [System.TeamProject] = 'InterACT Core' AND  [System.IterationPath] = 'InterACT Core\\Sprint " + sprintNum + "'" +
                                " AND [Slb.Interact.Team] = 'Real Time' AND [System.WorkItemType] = 'Task' AND  [System.State] != 'Removed'";
                    break;
                }
                else if ("Prism".Equals(tpName, StringComparison.OrdinalIgnoreCase))
                {
                    tpURI = "https://tfs.slb.com/tfs/SLB1/";
                    queryStr = "SELECT * FROM WorkItems WHERE [System.TeamProject] = 'Prism' AND  [System.IterationPath] = 'Prism\\Sprint " + sprintNum + "'" +
                                "AND [System.WorkItemType] = 'Task' AND  [System.State] != 'Removed'";
                    break;
                }
            }
            //创建文件 
            //StreamWriter sw = File.AppendText(System.Environment.CurrentDirectory + "\\AnalysisResult.csv");           
            FileStream fs = new FileStream(Environment.CurrentDirectory + "\\" + tpName + "_Sprint_" + sprintNum + "_Result.csv", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
            fs.SetLength(0);
            StreamWriter sw = new StreamWriter(fs);
            
            //服务器连接，相关对象的获取
            TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri(tpURI));
            //版本控制
            version = tpc.GetService(typeof(VersionControlServer)) as VersionControlServer;

            WorkItemStore workItemStore = tpc.GetService<WorkItemStore>();
            
            WorkItemCollection wc = workItemStore.Query(queryStr);
            int sumCount = wc.Count;            
            Console.WriteLine("num: {0}", sumCount);

            Dictionary<string, int[]> status = new Dictionary<string, int[]>();
            Dictionary<string, int[]> effort = new Dictionary<string, int[]>();
            Dictionary<string, int[]> actiEffort = new Dictionary<string, int[]>();
            actiEffort.Add("Grand Total", new int[2]);

            foreach(WorkItem wi in wc)
            {   
                string creator = wi.CreatedBy;                
                int revs = wi.Revision;
                Revision rev = wi.Revisions[revs-1];

                /*
                foreach (Field fi in wi.Fields)
                {
                    Console.WriteLine(fi.Name);
                    Console.WriteLine(rev.Fields[fi.Name].Value);
                }
                Thread.Sleep(500000);
                 * 
                 */
                //Activity VS effort----------------------------------------
                string[] effortType = new string[] { "Test automation", "Testing", "Development", "Meeting", "defect", "Training", "Review", "Knowledge transfer", "Commercial support", "Requirements" };
                foreach(string ty in effortType)
                {
                    //string activity = (string)rev.Fields["Activity"].Value;
                }
                int[] activity;
                string actiName;
                if ("".Equals((string)rev.Fields["Activity"].Value) == false)
                {
                    actiName = (string)rev.Fields["Activity"].Value;
                }
                else
                {
                    actiName = "(blank)";
                }

                int[] totalEffort = actiEffort["Grand Total"];
                totalEffort[0] = totalEffort[0] + Convert.ToInt32(rev.Fields["Original Estimate"].Value);
                totalEffort[1] = totalEffort[1] + Convert.ToInt32(rev.Fields["Completed work"].Value);
                if (actiEffort.ContainsKey(actiName))
                {
                    activity = actiEffort[actiName];
                }
                else
                {
                    activity = new int[2];
                    actiEffort.Add(actiName,activity);
                }
                activity[0] = activity[0] + Convert.ToInt32(rev.Fields["Original Estimate"].Value);
                activity[1] = activity[1] + Convert.ToInt32(rev.Fields["Completed work"].Value);

                //Sprint original estimation of real time team------------------------
                int[] effortCount;
                string associated = (string)rev.Fields["Assigned to"].Value;
                //Console.WriteLine(associated);
                if (effort.ContainsKey(associated))
                {
                    effortCount = effort[associated];
                }
                else
                {
                    effortCount = new int[2];
                    effort.Add(associated, effortCount);
                }
                
                effortCount[0] = effortCount[0] + Convert.ToInt32(rev.Fields["Original Estimate"].Value);
                effortCount[1] = effortCount[1] + Convert.ToInt32(rev.Fields["Completed work"].Value);
                
                //Sprint planned committed tasks VS status-------------                
                if (wi.Title.StartsWith("[C"))
                {
                    int[] staCount;
                    if (status.ContainsKey(creator))
                    {
                        staCount = status[creator];
                    }
                    else
                    {
                        staCount = new int[3];
                        status.Add(creator, staCount);
                    }
                    if ("New".Equals(wi.State))
                    {
                        staCount[0]++;
                    }
                    else if ("Active".Equals(wi.State))
                    {
                        staCount[1]++;
                    }
                    else if ("Closed".Equals(wi.State))
                    {
                        staCount[2]++;
                    }
                }
            }

            //写入文件-----------------------------------------------
            sw.WriteLine("Sprint original estimation of real time team");
            sw.WriteLine("Name,Sum of Original Estimate,Sum of Completed Work");
            foreach (KeyValuePair<string, int[]> item in effort)
            {
                sw.WriteLine("" + item.Key + ',' + item.Value[0] + ',' + item.Value[1]);
            }
            sw.WriteLine();
            sw.WriteLine();
            sw.WriteLine("Sprint planned committed tasks VS status");
            sw.WriteLine("Name,New,Active,Closed");
            foreach (KeyValuePair<string, int[]> item in status)
            {
                sw.WriteLine("" + item.Key + ',' + item.Value[0] + ',' + item.Value[1] + ',' + item.Value[2]);
            }
            sw.WriteLine();
            sw.WriteLine();            
            sw.WriteLine("Activity VS effort");
            sw.WriteLine("Activity,Sum of Original Estimate,% Original,Sum of Completed Work,% Real");
            double totalOrig = actiEffort["Grand Total"][0];
            double totalReal = actiEffort["Grand Total"][1];
            foreach (KeyValuePair<string, int[]> item in actiEffort)
            {
                if ("Grand Total".Equals(item.Key) == false)
                {
                    sw.WriteLine("" + item.Key + ',' + item.Value[0] + ',' + (item.Value[0] / totalOrig * 100).ToString("f2") + "%" + ',' + item.Value[1] + ',' + (item.Value[1] / totalReal * 100).ToString("f2") + "%");
                }
            }
            sw.WriteLine("" + "Grand Total" + ',' + totalOrig + ',' + 1 + ',' + totalReal + ',' + 1);
            Console.WriteLine("*************end*************");
            sw.Close();
            fs.Close();
            Thread.Sleep(5000);

        }
        
    }

}
