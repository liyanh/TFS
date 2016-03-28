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

namespace AllUserStory
{
    class UserStorySummary
    {
        static TfsTeamProjectCollection tpc = null;
        static VersionControlServer version = null;
        static WorkItemStore workItemStore = null;
        static FileStream fs = null;
        static StreamWriter sw = null;
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
            sw.WriteLine("User Story");
            sw.WriteLine("StoryID, Assigned To, Status,Iteration Path,Area Path,Story Point,Total Effort of User Story,Title,User story Created Date,User story Resolved Date,User story Closed Date,Earliest Task Create Time,User Stroy Resolved/Closed Time,Latest Task Closed Time,Lead Time");
            Dictionary<string, List<string>> user_story = new UserStorySummary().allUserStory();
            foreach (KeyValuePair<string, List<string>> item in user_story)
            {
                sw.Write("" + item.Key);
                foreach (string res in item.Value)
                {
                    sw.Write(',' + res);
                }
                sw.WriteLine();
            }
        }
        
        public Dictionary<string, List<string>> allUserStory()
        {
            WorkItemCollection wcStory = workItemStore.Query(
                "SELECT * FROM WorkItems WHERE [System.WorkItemType] = 'User Story' AND  [System.State] != 'Removed'");
            Dictionary<string, List<string>> user_story = new Dictionary<string, List<string>>();
            foreach (WorkItem wi in wcStory)
            {
                List<string> story_Info = new List<string>();
                Revision rev = WorkItemInfo.getWorkItemRevision(wi);
                story_Info.Add(rev.Fields["Assigned To"].Value.ToString());
                story_Info.Add(rev.Fields["State"].Value.ToString());
                story_Info.Add(rev.Fields["Iteration Path"].Value.ToString());
                story_Info.Add(rev.Fields["Area Path"].Value.ToString());
                if (rev.Fields["Story Points"].Value != null)
                {
                    story_Info.Add(rev.Fields["Story Points"].Value.ToString());
                }
                else
                {
                    story_Info.Add("null");
                }
                double totalCompletedEffort = 0;
                List<WorkItem> childItems = WorkItemInfo.getWorkItemChild(workItemStore, wi);
                foreach (WorkItem child in childItems)
                {
                    Revision revChild = WorkItemInfo.getWorkItemRevision(child);
                    double comp = Convert.ToDouble(revChild.Fields["Completed work"].Value);
                    totalCompletedEffort = totalCompletedEffort + comp;                    
                }
                story_Info.Add((totalCompletedEffort / 8).ToString());
                story_Info.Add(rev.Fields["Title"].Value.ToString().Replace(",", "."));
                story_Info.Add(rev.Fields["Created Date"].Value.ToString());
                if (rev.Fields["Resolved Date"].Value != null)
                {
                    story_Info.Add(rev.Fields["Resolved Date"].Value.ToString());
                }
                else
                {
                    story_Info.Add(string.Empty);
                }
                if (rev.Fields["Closed Date"].Value != null)
                {
                    story_Info.Add(rev.Fields["Closed Date"].Value.ToString());
                }
                else
                {
                    story_Info.Add(string.Empty);
                }
                story_Info.Add(startTime(wi));
                story_Info.Add(usClosedTime(wi));
                story_Info.Add(latestTaskClosedTime(wi));
                story_Info.Add(leadTime(wi));
                string storyId = rev.Fields["Id"].Value.ToString();
                user_story.Add(storyId, story_Info);
            }
            return user_story;
        }

        //user story leadtime
        public string leadTime(WorkItem wi)
        {
            string leadTime = string.Empty;
            Revision rev = WorkItemInfo.getWorkItemRevision(wi);
            string usState = rev.Fields["State"].Value.ToString();
            if (("Closed".Equals(usState) | "Resolved".Equals(usState)) && startTime(wi) != null)
            {
                DateTime earliestTask = DateTime.Parse(startTime(wi), CultureInfo.InvariantCulture);
                DateTime usEnd = DateTime.Parse(usClosedTime(wi), CultureInfo.InvariantCulture);
                TimeSpan dura = usEnd - earliestTask;
                leadTime = (dura.Days + Convert.ToDouble(dura.Hours) / 24).ToString("0.00");
            }
            return leadTime;
        }

        //user story closed time
        public string usClosedTime(WorkItem wi)
        {
            string closedTime = string.Empty;
            Revision rev = WorkItemInfo.getWorkItemRevision(wi);
            string usState = rev.Fields["State"].Value.ToString();
            if ("Closed".Equals(usState) | "Resolved".Equals(usState))
            {
                if (rev.Fields["Closed Date"].Value != null)
                {
                    closedTime = rev.Fields["Closed Date"].Value.ToString();
                }
                else
                {
                    closedTime = rev.Fields["Resolved Date"].Value.ToString();
                }
            }
            return closedTime;
        }

        //user story start time
        public string startTime(WorkItem wi)
        {
            string stTime = null;
            DateTime earliestTask = DateTime.Now;
            List<WorkItem> childItems = WorkItemInfo.getWorkItemChild(workItemStore, wi);
            if (childItems.Count() == 0)
            {
                return stTime;
            }
            else
            {
                foreach (WorkItem child in childItems)
                {
                    Revision revChild = WorkItemInfo.getWorkItemRevision(child);
                    DateTime taskCreate = DateTime.Parse(revChild.Fields["Created Date"].Value.ToString(), CultureInfo.InvariantCulture);
                    earliestTask = earliestTask < taskCreate ? earliestTask : taskCreate;
                }
                return earliestTask.ToString();
            }
        }

        //latest task closed time
        public string latestTaskClosedTime(WorkItem wi)
        {
            string ltcTime = string.Empty;
            Revision rev = WorkItemInfo.getWorkItemRevision(wi);
            DateTime latestTask = DateTime.Parse(rev.Fields["Created Date"].Value.ToString(), CultureInfo.InvariantCulture);
            string usState = rev.Fields["State"].Value.ToString();
            if ("Closed".Equals(usState) | "Resolved".Equals(usState))
            {
                List<WorkItem> childItems = WorkItemInfo.getWorkItemChild(workItemStore, wi);
                if (childItems.Count() != 0)
                {
                    foreach (WorkItem child in childItems)
                    {
                        Revision revChild = WorkItemInfo.getWorkItemRevision(child);
                        if ("Closed".Equals(revChild.Fields["State"].Value))
                        {
                            DateTime taskClosed = DateTime.Parse(revChild.Fields["Closed Date"].Value.ToString(), CultureInfo.InvariantCulture);
                            latestTask = latestTask > taskClosed ? latestTask : taskClosed;
                        }
                    }
                    ltcTime = latestTask.ToString();
                }
            }
            return ltcTime;
        }
    }
}
