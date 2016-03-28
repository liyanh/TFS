using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
namespace SprintQuery
{
    class Program
    {
        static VersionControlServer version = null;
        static double closedStoryPoint = 0;
        static double usTaskComp = 0;
        static FileStream fs = null;
        static StreamWriter sw = null;
        static TfsTeamProjectCollection tpc = null;
        static WorkItemStore workItemStore = null;
        static Dictionary<string, double[]> actiEffort = null;
        static Dictionary<string, double[]> memEffort = null;

        static void Main(string[] args)
        {                       
            Console.WriteLine("123!");
            System.Console.Write("Input SprintNum: ");
            string sprintNum = Console.ReadLine();
            //创建文件           
            fs = new FileStream(Environment.CurrentDirectory + "\\" + "Sprint_" + sprintNum + "_Result.csv", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
            fs.SetLength(0);
            sw = new StreamWriter(fs);
            
            //服务器连接，相关对象的获取
            tpc = new TfsTeamProjectCollection(new Uri("https://tfs.slb.com/tfs/SLB1/"));
            version = tpc.GetService(typeof(VersionControlServer)) as VersionControlServer;
            workItemStore = tpc.GetService<WorkItemStore>();
          
            CountClosedStoryEffort(sprintNum);

            Dictionary<string, List<string>> user_story = storyQuery(sprintNum);

            taskQuery(sprintNum);

            //写入文件-----------------------------------------------
            sw.WriteLine("Sprint member effort");
            sw.WriteLine("Name,Sum of Original Estimate,Sum of Completed Work");
            foreach (KeyValuePair<string, double[]> item in memEffort)
            {
                sw.WriteLine("" + item.Key + ',' + item.Value[0] + ',' + item.Value[1]);
            }
            sw.WriteLine();
            sw.WriteLine();
           
            sw.WriteLine("Activity VS effort");
            sw.WriteLine("Activity,Sum of Original Estimate,% Original,Sum of Completed Work,% Real");
            double totalOrig = actiEffort["Grand Total"][0];
            double totalReal = actiEffort["Grand Total"][1];
            foreach (KeyValuePair<string, double[]> item in actiEffort)
            {
                if ("Grand Total".Equals(item.Key) == false)
                {
                    sw.WriteLine("" + item.Key + ',' + item.Value[0] + ',' + (item.Value[0] / totalOrig * 100).ToString("f2") + "%" + ',' + item.Value[1] + ',' + (item.Value[1] / totalReal * 100).ToString("f2") + "%");
                }
            }
            sw.WriteLine("" + "Grand Total" + ',' + totalOrig + ',' + 1 + ',' + totalReal + ',' + 1);

            //stroy points write
            sw.WriteLine();
            sw.WriteLine();
            sw.WriteLine("User Story");
            sw.WriteLine("StoryID, Assigned To, Status,Story Point,Completed effort in This Sprint,Title,Start Time,Lead Time");
            double usTaskCompPoint = usTaskComp / 8;
            foreach (KeyValuePair<string, List<string>> item in user_story)
            {
                sw.Write("" + item.Key);
                foreach(string res in item.Value)
                {
                    sw.Write(',' + res);                    
                }
                sw.WriteLine();
            }
            sw.WriteLine("" + "Sum");
            sw.WriteLine("" + "Closed Story Points" + ',' + closedStoryPoint);
            sw.WriteLine("" + "Completed effort in Closed Story" + ',' + usTaskCompPoint);
            Console.WriteLine("*************end*************");
            sw.Close();
            fs.Close();
            Thread.Sleep(5000);

        }

        private static void taskQuery(string sprintNum)
        {
            WorkItemCollection wc = workItemStore.Query(
                "SELECT * FROM WorkItems WHERE [System.TeamProject] = 'Prism' AND [System.AreaPath] UNDER 'Prism\\Core' AND  [System.IterationPath] = 'Prism\\Sprint " + sprintNum + "'" +
                "AND [System.WorkItemType] = 'Task' AND  [System.State] != 'Removed'");
            memEffort = new Dictionary<string, double[]>();
            actiEffort = new Dictionary<string, double[]>();
            actiEffort.Add("Grand Total", new double[2]);

            foreach (WorkItem wi in wc)
            {
                Revision rev = getWorkItemRevision(wi);
                activityEffort(wi,rev);
                memberEffort(wi,rev);
            }
        }

        private static void activityEffort(WorkItem wi,Revision rev)
        {
            actiEffort["Grand Total"][0] = actiEffort["Grand Total"][0] + Convert.ToDouble(rev.Fields["Original Estimate"].Value);
            actiEffort["Grand Total"][1] = actiEffort["Grand Total"][1] + Convert.ToDouble(rev.Fields["Completed work"].Value);
            double[] activity;
            string actiName = string.Empty;
            if ("".Equals((string)rev.Fields["Activity"].Value) == false)
            {
                actiName = (string)rev.Fields["Activity"].Value;
            }
            else
            {
                actiName = "(blank)";
            }            
            if (actiEffort.ContainsKey(actiName))
            {
                activity = actiEffort[actiName];
            }
            else
            {
                activity = new double[2];
                actiEffort.Add(actiName, activity);
            }
            activity[0] = activity[0] + Convert.ToInt32(rev.Fields["Original Estimate"].Value);
            activity[1] = activity[1] + Convert.ToInt32(rev.Fields["Completed work"].Value);
        }

        private static void memberEffort(WorkItem wi, Revision rev)
        {
            double[] effortCount;
            string associated = (string)rev.Fields["Assigned to"].Value;
            //Console.WriteLine(associated);
            if (memEffort.ContainsKey(associated))
            {
                effortCount = memEffort[associated];
            }
            else
            {
                effortCount = new double[2];
                memEffort.Add(associated, effortCount);
            }

            effortCount[0] = effortCount[0] + Convert.ToInt32(rev.Fields["Original Estimate"].Value);
            effortCount[1] = effortCount[1] + Convert.ToInt32(rev.Fields["Completed work"].Value);
        }

        private static Dictionary<string, List<string>> storyQuery(string sprintNum)
        {
            WorkItemCollection wcStory = workItemStore.Query(
                "SELECT * FROM WorkItems WHERE [System.TeamProject] = 'Prism' AND [System.AreaPath] UNDER 'Prism\\Core' AND [System.IterationPath] = 'Prism\\Sprint " + sprintNum + "'" +
                "AND [System.WorkItemType] = 'User Story' AND  [System.State] != 'Removed'");
            Dictionary<string, List<string>> user_story = new Dictionary<string, List<string>>();
            foreach (WorkItem wi in wcStory)
            {
                List<string> story_Info = new List<string>();
                Revision rev = getWorkItemRevision(wi);
                story_Info.Add(rev.Fields["Assigned To"].Value.ToString());
                story_Info.Add(rev.Fields["State"].Value.ToString());
                if (rev.Fields["Story Points"].Value != null)
                {
                    story_Info.Add(rev.Fields["Story Points"].Value.ToString());
                }
                else
                {
                    story_Info.Add("null");
                }
                double usCompletedEffort = 0;
                List<WorkItem> childItems = getWorkItemChild(wi);
                foreach (WorkItem child in childItems)
                {
                    Revision revChild = getWorkItemRevision(child);
                    if (revChild.Fields["Iteration Path"].Value.ToString().EndsWith("\\Sprint " + sprintNum))
                    {
                        double comp = Convert.ToDouble(revChild.Fields["Completed work"].Value);
                        usCompletedEffort = usCompletedEffort + comp;
                    }
                }
                story_Info.Add((usCompletedEffort / 8).ToString());
                story_Info.Add(rev.Fields["Title"].Value.ToString().Replace(",", "."));
                story_Info.Add(startTime(wi));
                story_Info.Add(leadTime(wi));
                string storyId = rev.Fields["Id"].Value.ToString();
                user_story.Add(storyId, story_Info);
            }
            return user_story;
        }

        //user story leadtime
        private static string leadTime(WorkItem wi)
        {
            string leadTime = string.Empty;
            Revision rev = getWorkItemRevision(wi);           
            string usState = rev.Fields["State"].Value.ToString();
            if ("Closed".Equals(usState) | "Resolved".Equals(usState))
            {
                DateTime earliestTask = DateTime.Parse(startTime(wi), CultureInfo.InvariantCulture);
                DateTime usEnd = new DateTime();                                
                if (rev.Fields["Closed Date"].Value != null)
                {
                    usEnd = DateTime.Parse(rev.Fields["Closed Date"].Value.ToString(), CultureInfo.InvariantCulture);
                }
                else
                {
                    usEnd = DateTime.Parse(rev.Fields["Resolved Date"].Value.ToString(), CultureInfo.InvariantCulture);
                }
                TimeSpan dura = usEnd - earliestTask;
                leadTime = (dura.Days + Convert.ToDouble(dura.Hours)/24).ToString("0.00");
            }           
            return leadTime;            
        }

        //user story start time
        private static string startTime(WorkItem wi)
        {
            string stTime = null;
            DateTime earliestTask = DateTime.Now;
            List<WorkItem> childItems = getWorkItemChild(wi);
            if (childItems.Count() == 0)
            {
                return stTime;
            }
            else
            {
                foreach (WorkItem child in childItems)
                {
                    Revision revChild = getWorkItemRevision(child);
                    DateTime taskCreate = DateTime.Parse(revChild.Fields["Created Date"].Value.ToString(), CultureInfo.InvariantCulture);
                    earliestTask = earliestTask < taskCreate ? earliestTask : taskCreate;
                }
                return earliestTask.ToString();
            }
        }

        //closed user story story point vs acturl copmleted work effort
        private static void CountClosedStoryEffort(string sprintNum)
        {
            WorkItemCollection wcClosedStory = workItemStore.Query(
                "SELECT * FROM WorkItems WHERE [System.TeamProject] = 'Prism' AND [System.AreaPath] UNDER 'Prism\\Core' AND [System.IterationPath] = 'Prism\\Sprint " + sprintNum + "'" +
                "AND [System.WorkItemType] = 'User Story' AND  [System.State] = 'Closed'");

            foreach (WorkItem wi in wcClosedStory)
            {
                Revision rev = getWorkItemRevision(wi);               
                if (rev.Fields["Story Points"].Value != null)
                {
                    closedStoryPoint = closedStoryPoint + double.Parse(rev.Fields["Story Points"].Value.ToString());
                }
                List<WorkItem> closedUSChild = getWorkItemChild(wi);
                foreach (WorkItem childItem in closedUSChild)
                {
                    Revision revChild = getWorkItemRevision(childItem);
                    double comp = Convert.ToDouble(revChild.Fields["Completed work"].Value);
                    usTaskComp = usTaskComp + comp;
                }
            }
        }

        private static List<WorkItem> getWorkItemChild(WorkItem wi)
        {
            List<WorkItem> linkedItem = new List<WorkItem>();
            WorkItemLinkCollection links = wi.WorkItemLinks;
            foreach (WorkItemLink wil in links)
            {
                //search for the correct link
                if (wil.LinkTypeEnd.Name.Equals("Child"))
                {
                    WorkItem childItem = workItemStore.GetWorkItem(wil.TargetId);                    
                    linkedItem.Add(childItem);
                }
            }
            return linkedItem;
        }

        private static Revision getWorkItemRevision(WorkItem item)
        {
            int revNum = item.Revision;
            Revision revis = item.Revisions[revNum - 1];
            return revis;
        }
        /*
        foreach (Field fi in childItem.Fields)
        {
            Console.WriteLine(fi.Name);
            Console.WriteLine(revChild.Fields[fi.Name].Value);
       }
       */
    }
}
