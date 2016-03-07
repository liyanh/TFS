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
namespace SprintQuery
{
    class Program
    {
        static string sprintNum = "";
        static VersionControlServer version = null;
        static double storyPoint = 0;
        static double closedStoryPoint = 0;

        static void Main(string[] args)
        {                       
            Console.WriteLine("123!");
            System.Console.Write("Input SprintNum: ");
            sprintNum = Console.ReadLine();
            //创建文件 
            //StreamWriter sw = File.AppendText(System.Environment.CurrentDirectory + "\\AnalysisResult.csv");           
            FileStream fs = new FileStream(Environment.CurrentDirectory + "\\" + "Sprint_" + sprintNum + "_Result.csv", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
            fs.SetLength(0);
            StreamWriter sw = new StreamWriter(fs);
            
            //string status System.Configuration.ConfigurationSettings.AppSettings["Key"] 
            //服务器连接，相关对象的获取
            //System.Net.ICredentials credentials = new System.Net.NetworkCredential("YLi102","Lifor#900516","DIR.slb.com");
            TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri("https://tfs.slb.com/tfs/SLB1/"));
            //版本控制
            version = tpc.GetService(typeof(VersionControlServer)) as VersionControlServer;

            WorkItemStore workItemStore = tpc.GetService<WorkItemStore>();
            
            //story point

            WorkItemCollection wcClosedStory = workItemStore.Query(
                "SELECT * FROM WorkItems WHERE [System.TeamProject] = 'Prism' AND [System.AreaPath] UNDER 'Prism\\Core' AND [System.IterationPath] = 'Prism\\Sprint " + sprintNum + "'" +
                "AND [System.WorkItemType] = 'User Story' AND  [System.State] = 'Closed'");
            
            foreach (WorkItem wi in wcClosedStory)
            {
               
                int revs = wi.Revision;
                Revision rev = wi.Revisions[revs - 1];
                /*
                foreach (Field fi in wi.Fields)
                {
                    Console.WriteLine(fi.Name);
                    Console.WriteLine(rev.Fields[fi.Name].Value);
                }
                */                             

                if (rev.Fields["Story Points"].Value != null)
                {
                    closedStoryPoint = closedStoryPoint + double.Parse(rev.Fields["Story Points"].Value.ToString());
                }
            }


            WorkItemCollection wcStory = workItemStore.Query(
                "SELECT * FROM WorkItems WHERE [System.TeamProject] = 'Prism' AND [System.AreaPath] UNDER 'Prism\\Core' AND [System.IterationPath] = 'Prism\\Sprint " + sprintNum + "'" +
                "AND [System.WorkItemType] = 'User Story' AND  [System.State] != 'Removed'");

            Dictionary<string, List<string>> user_story = new Dictionary<string, List<string>>();
            foreach (WorkItem wi in wcStory)
            {
                List<string> story_Info = new List<string>();
                int revs = wi.Revision;
                Revision rev = wi.Revisions[revs - 1];
                /*
                foreach (Field fi in wi.Fields)
                {
                    Console.WriteLine(fi.Name);
                    Console.WriteLine(rev.Fields[fi.Name].Value);
                }
                */
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
                
                //Console.WriteLine(rev.Fields["Title"].Value.ToString());
                if (rev.Fields["Story Points"].Value != null)
                {
                    storyPoint = storyPoint + double.Parse(rev.Fields["Story Points"].Value.ToString());
                }

                //get workItem parent link
                WorkItemLinkCollection links = wi.WorkItemLinks;

                double taskPoint = 0;
                foreach (WorkItemLink wil in links)
                {
                    //search for the correct link
                    if (wil.LinkTypeEnd.Name.Equals("Child"))
                    {
                        WorkItem childItem = workItemStore.GetWorkItem(wil.TargetId);
                        if (childItem.State.Equals("Closed"))
                        {
                            int revsChild = childItem.Revision;
                            Revision revChild = childItem.Revisions[revsChild-1];
                            /*
                            foreach (Field fi in childItem.Fields)
                            {
                                Console.WriteLine(fi.Name);
                                Console.WriteLine(revChild.Fields[fi.Name].Value);
                            }
                            */
                            if (revChild.Fields["Iteration Path"].Value.ToString().EndsWith("\\Sprint " + sprintNum))
                            {                                
                                taskPoint = taskPoint + Convert.ToInt32(revChild.Fields["Completed work"].Value);
                            }                            
                        }
                    }
                }
                //Console.WriteLine(taskPoint);
                story_Info.Add((taskPoint/8).ToString());
                story_Info.Add(rev.Fields["Title"].Value.ToString());
                string storyId = rev.Fields["Id"].Value.ToString();
                //story_Info.Add(rev.Fields["Story Points"].Value.ToString());
                //story_Info.Add(rev.Fields["Story Points"].Value.ToString());

                user_story.Add(storyId, story_Info);
            }
            WorkItemCollection wcTask = workItemStore.Query(
                "SELECT * FROM WorkItems WHERE [System.TeamProject] = 'Prism' AND [System.AreaPath] UNDER 'Prism\\Core' AND  [System.IterationPath] = 'Prism\\Sprint " + sprintNum + "'" +
                "AND [System.WorkItemType] = 'Task' AND  [System.State] = 'Closed'");
            double complTaskEff = 0;

            foreach (WorkItem wi in wcTask)
            {
                string creator = wi.CreatedBy;
                int revs = wi.Revision;
                Revision rev = wi.Revisions[revs - 1];
                /*
                var parent_link = wi.WorkItemLinks.Cast<WorkItemLink>().FirstOrDefault(x => x.LinkTypeEnd.Name == "Parent");
                WorkItem parent = null;
                if (parent_link != null)
                {
                    parent = workItemStore.GetWorkItem(parent_link.TargetId);
                }
                //WorkItem par = (WorkItem)wico.GetEnumerator().;
                Console.WriteLine(parent.Title);
                foreach (Field fi in wi.Fields)
                {
                    Console.WriteLine(fi.Name);
                    Console.WriteLine(rev.Fields[fi.Name].Value);
                }
                 * 
                 * 
                 *   
                //get workItem parent link
                WorkItemLinkCollection links = workItem.WorkItemLinks;
                
                foreach (WorkItemLink wil in links)
                {
                    //search for the correct link
                    if (wil.LinkTypeEnd.Name.Equals("Parent"))
                    {
                        //retrieve the parent workItem
                        workItem = workItemStore.GetWorkItem(wil.TargetId);
                    }
                } 
                 */
                complTaskEff = complTaskEff + Convert.ToInt32(rev.Fields["Completed work"].Value);
            }
            WorkItemCollection wc = workItemStore.Query(
                "SELECT * FROM WorkItems WHERE [System.TeamProject] = 'Prism' AND [System.AreaPath] UNDER 'Prism\\Core' AND  [System.IterationPath] = 'Prism\\Sprint " + sprintNum + "'" +
                "AND [System.WorkItemType] = 'Task' AND  [System.State] != 'Removed'");
            int sumCount = wc.Count;            
            //Console.WriteLine("num: {0}", sumCount);

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
                 */

                //Activity VS effort----------------------------------------
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
                
                /*
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
                 */
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
            /*
            sw.WriteLine("Sprint planned committed tasks VS status");
            sw.WriteLine("Name,New,Active,Closed");
            foreach (KeyValuePair<string, int[]> item in status)
            {
                sw.WriteLine("" + item.Key + ',' + item.Value[0] + ',' + item.Value[1] + ',' + item.Value[2]);
            }
             */
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

            //stroy points write
            sw.WriteLine();
            sw.WriteLine();
            sw.WriteLine("StoryID, Assigned To, Status,Story Point,Task Complted,Title");
            double completedPoint = complTaskEff / 8;
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
            sw.WriteLine("" + "All Story Points" + ',' + storyPoint);
            sw.WriteLine("" + "Closed Story Points" + ',' + closedStoryPoint);
            sw.WriteLine("" + "Completed Effort" + ',' + completedPoint);
            Console.WriteLine("*************end*************");
            sw.Close();
            fs.Close();
            Thread.Sleep(5000);

        }
        
    }

}
