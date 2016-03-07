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

        static DateTime DateStart = DateTime.Now;
        static DateTime DateEnd = DateTime.Now;
        static VersionControlServer version = null;
        static string startdate = "";
        static string enddate = "";
       
        static void Main(string[] args)
        {                       
            Console.WriteLine("123!");
                    
            //查询时间
            Boolean startDateBoo = true;
            while (startDateBoo)
            {
                try
                {
                    //DateTime.TryParse(string.Format("{yyyy-MM-DD}:00", starttime),out dt);                   
                    // /*
                    //* 暂时注释，输入时间 datestart and dateend,try timeformate catch

                    System.Console.Write("Input the start time (yyyy-mm-dd): ");
                    startdate = Console.ReadLine();
                    string starttime = startdate + " 00:00:00";
                    DateStart = DateTime.ParseExact(starttime, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    startDateBoo = false;
                }
                catch (FormatException)
                {
                    Console.WriteLine("Unable to convert to date");
                    startDateBoo = true;
                }
            }

            Boolean endDateBoo = true;
            while (endDateBoo)
            {
                System.Console.Write("Input the end time (yyyy-mm-dd): ");
                enddate = Console.ReadLine();
                string endtime = enddate + " 23:59:59";
                try
                {
                    DateEnd = DateTime.ParseExact(endtime, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    endDateBoo = false;
                }
                catch (FormatException)
                {
                    Console.WriteLine("Unable to convert to date");
                    endDateBoo = true;

                }
            }

            //创建文件 
            //StreamWriter sw = File.AppendText(System.Environment.CurrentDirectory + "\\AnalysisResult.csv");           
            FileStream fs = new FileStream(Environment.CurrentDirectory + "\\" + startdate + "---" + enddate  + "_Result.csv", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
            fs.SetLength(0);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine("ChangesetID,Owner,Group,CreateDate,Branch,LinkID,LinkType,Comment");

            //服务器连接，相关对象的获取
            TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri("https://tfs.slb.com/tfs/Real-Time_Collection/"));
            //版本控制
            version = tpc.GetService(typeof(VersionControlServer)) as VersionControlServer;
            
            //--------------------------team group-------------------------------------------
            System.Console.WriteLine("Getting Team Information......");
            TeamProject[] tepro = version.GetAllTeamProjects(true);
            TeamProject mypro = tepro[0];
            string rootFolder = mypro.ArtifactUri.AbsoluteUri;

            IGroupSecurityService gss = (IGroupSecurityService)tpc.GetService(typeof(IGroupSecurityService));
            Identity[] appGroups = gss.ListApplicationGroups(rootFolder);

            Dictionary<string, ArrayList> groupMap = new Dictionary<string, ArrayList>();

            ArrayList RTgroup =  new ArrayList();
            RTgroup.Add("Yizhou WANG");
            groupMap.Add("Real time", RTgroup);
            groupMap.Add("Data delivery", new ArrayList());
            groupMap.Add("Photon", new ArrayList());
            groupMap.Add("ProXimity", new ArrayList());
            groupMap.Add("Visualization", new ArrayList());
            groupMap.Add("QA", new ArrayList());

            for (int k = 0; k < appGroups.Length; k++)
            {                
                Identity group = appGroups[k];
                if (groupMap.ContainsKey(group.DisplayName))
                {
                    Identity[] groupMembers = gss.ReadIdentities(SearchFactor.Sid, new string[] { group.Sid }, QueryMembership.Expanded);
                    foreach (Identity member in groupMembers)
                    {
                        if (member.Members != null)
                        {
                            foreach (string memberSid in member.Members)
                            {
                                Identity memberInfo = gss.ReadIdentity(SearchFactor.Sid, memberSid, QueryMembership.Expanded);
                                groupMap[group.DisplayName].Add(memberInfo.DisplayName);
                            }
                        }
                    }
                    
                }
                
            }
            //------------------------------------------------------------------
           
            VersionSpec fromDateVersion = new DateVersionSpec(DateStart);
            VersionSpec toDateVersion = new DateVersionSpec(DateEnd);
            // using versionControl API you can query histroy for changes set 
            IEnumerable changesets = version.QueryHistory(string.Concat("$/", mypro.Name), VersionSpec.Latest, 0, RecursionType.Full, null, fromDateVersion, toDateVersion, int.MaxValue, true , true);

            //查询 changeset 相关信息的获取--------------------------------------
            Dictionary<string, ArrayList> countUnlinked = new Dictionary<string, ArrayList>();
            Dictionary<string, ArrayList> countTasklinked = new Dictionary<string, ArrayList>();
            System.Console.WriteLine("Query Changesets by Date......");
            string csTime = "";
            foreach (Changeset changeset in changesets)
            {
                Changeset cs = changeset;

                if (csTime.Equals(cs.CreationDate.ToShortDateString()) == false)
                {
                    if ("".Equals(csTime) == false)
                    {
                        System.Console.WriteLine("      done!");
                    }
                    csTime = cs.CreationDate.ToShortDateString().ToString();
                    System.Console.Write(csTime);
                }

                //System.Console.WriteLine(cs.ChangesetId + " *** " + cs.CommitterDisplayName);
                Change[] thchan = cs.Changes;
                string branch = thchan[0].Item.ServerItem;
                String[] braPath = branch.Split('/');
                if (braPath.Length > 4)
                {
                    branch = "";
                    for (int k = 1; k < 4; k++)
                    {
                        branch = branch + braPath[k] + "/";
                    }
                    if ("Release".Equals(braPath[3]))
                    {
                        branch = branch + braPath[4];
                    }
                }
                //System.Console.WriteLine(branch);

                //teamInfo-----------------------------------
                string teamName = "others";
                foreach (KeyValuePair<string, ArrayList> groupInfo in groupMap)
                {
                    if (groupInfo.Value.Contains(cs.CommitterDisplayName))
                    {
                        teamName = groupInfo.Key;
                    }
                }
                
                //-----------------------------------

                //associatedWork-----------------------------------
                AssociatedWorkItemInfo[] work = cs.AssociatedWorkItems;
                string linkid = "";
                string linktype = "";
                if (work == null || work.Length < 1)
                {
                    if (countUnlinked.ContainsKey(cs.CommitterDisplayName))
                    {
                        countUnlinked[cs.CommitterDisplayName].Add(cs.ChangesetId);
                    }
                    else
                    {                        
                        countUnlinked.Add(cs.CommitterDisplayName, new ArrayList());
                        countUnlinked[cs.CommitterDisplayName].Add(cs.ChangesetId);
                    }
                    sw.WriteLine("" + cs.ChangesetId + ',' + cs.CommitterDisplayName + ',' + teamName + ',' + cs.CreationDate + ',' + branch + ',' + linkid + ',' + linktype + ',' + Regex.Replace(cs.Comment, @"[\n\r]", ""));

                }
                else if (work.Length == 1)
                {
                    if ("Task".Equals(work[0].WorkItemType))
                    {
                        if (countTasklinked.ContainsKey(cs.CommitterDisplayName))
                        {
                            countTasklinked[cs.CommitterDisplayName].Add(cs.ChangesetId);
                        }
                        else
                        {
                            countTasklinked.Add(cs.CommitterDisplayName, new ArrayList());
                            countTasklinked[cs.CommitterDisplayName].Add(cs.ChangesetId);
                        }
                    }
                    linkid = (work[0].Id).ToString();
                    linktype = work[0].WorkItemType;
                    sw.WriteLine("" + cs.ChangesetId + ',' + cs.CommitterDisplayName + ',' + teamName + ',' + cs.CreationDate + ',' + branch + ',' + linkid + ',' + linktype + ',' + Regex.Replace(cs.Comment, @"[\n\r]", ""));
                }
                else
                {
                    for (int k = 0; k < work.Length; k++)
                    {
                        if (!("Task".Equals(work[k].WorkItemType)))
                        {
                            linkid = (work[k].Id).ToString();
                            linktype =work[k].WorkItemType;
                            sw.WriteLine("" + cs.ChangesetId + ',' + cs.CommitterDisplayName + ',' + teamName + ',' + cs.CreationDate + ',' + branch + ',' + linkid + ',' + linktype + ',' + Regex.Replace(cs.Comment, @"[\n\r]", ""));

                        }
                        if (k == work.Length - 1 && "".Equals(linkid))
                        {
                            if (countTasklinked.ContainsKey(cs.CommitterDisplayName))
                            {
                                countTasklinked[cs.CommitterDisplayName].Add(cs.ChangesetId);
                            }
                            else
                            {
                                countTasklinked.Add(cs.CommitterDisplayName, new ArrayList());
                                countTasklinked[cs.CommitterDisplayName].Add(cs.ChangesetId);
                            }
                            linkid = (work[k].Id).ToString();
                            linktype = work[k].WorkItemType;
                            sw.WriteLine("" + cs.ChangesetId + ',' + cs.CommitterDisplayName + ',' + teamName + ',' + cs.CreationDate + ',' + branch + ',' + linkid + ',' + linktype + ',' + Regex.Replace(cs.Comment, @"[\n\r]", ""));

                        }

                    }
                }
                
                //sw.WriteLine("" + cs.ChangesetId + ',' + cs.CommitterDisplayName + ',' + teamName + ',' + cs.CreationDate + ',' + branch + ',' + linkid + ',' + linktype + ',' + Regex.Replace(cs.Comment, @"[\n\r]", ""));
                                        
            }
                                              			
            sw.Close();
			fs.Close();
            System.Console.WriteLine("      done!");
            System.Console.WriteLine("**********end**********");

            // 统计信息
            FileStream cfs = new FileStream(Environment.CurrentDirectory + "\\" + startdate + "---" + enddate + "_Count.csv", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
            cfs.SetLength(0);
            StreamWriter csw = new StreamWriter(cfs);
            csw.WriteLine("Owner,Group,Unlinked,Task Linked");

            Dictionary<string, ArrayList>.KeyCollection keyUnlinked = countUnlinked.Keys;
            Dictionary<string, ArrayList>.KeyCollection keyTasklinked = countTasklinked.Keys;

            IEnumerable<string> union = keyUnlinked.Concat(keyTasklinked);
            HashSet<string> unionSet = new HashSet<string>(union);
            foreach(string name in unionSet)
            {
                int unLinked = 0;
                int taskLinked = 0;
                if (countUnlinked.ContainsKey(name))
                {
                    unLinked = countUnlinked[name].Count;
                }
                if (countTasklinked.ContainsKey(name))
                {
                    taskLinked = countTasklinked[name].Count;
                }
                string teamName = "others";
                foreach (KeyValuePair<string, ArrayList> groupInfo in groupMap)
                {
                    if (groupInfo.Value.Contains(name))
                    {
                        teamName = groupInfo.Key;
                    }
                }
                csw.WriteLine("" + name + ',' + teamName + ',' + unLinked + ',' + taskLinked);
            }
            csw.Close();
            cfs.Close();
            System.Console.WriteLine("**********Count Over**********");
           
			Thread.Sleep(5000);
            
        }
        
    }

}
