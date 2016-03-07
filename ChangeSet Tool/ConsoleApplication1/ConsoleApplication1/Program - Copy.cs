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
        /*BinarySearch----------------
        public static int BinarySearch(int start, int end)
        {

            int mid = (start + end) / 2;
            Changeset cs = null;
            try
            {
                while (true)
                {
                    if (version.GetChangeset(mid) != null)
                    {
                        cs = version.GetChangeset(mid);
                        break;
                    }
                    else
                    {
                        mid++;
                    }
                }
                if (cs.CreationDate.CompareTo(DateStart) < 0)
                {

                    BinarySearch(mid, end);

                }
                if (cs.CreationDate.CompareTo(DateEnd) > 0)
                {

                    BinarySearch(start, mid);
                }


                return mid;




            }

            catch (ResourceAccessException)
            {
                return BinarySearch(start + 1, end);
            }


        }
        */
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
            sw.WriteLine("changesetID,owner,group,createDate,branch,linkID,linktype,comment");

            //服务器连接，相关对象的获取
            TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri("https://tfs.slb.com/tfs/Real-Time_Collection/"));
            //版本控制
            version = tpc.GetService(typeof(VersionControlServer)) as VersionControlServer;
            //--------------------------team group-------------------------------------------

            TeamProject[] tepro = version.GetAllTeamProjects(true);
            TeamProject mypro = tepro[0];
            string rootFolder = mypro.ArtifactUri.AbsoluteUri;

            IGroupSecurityService gss = (IGroupSecurityService)tpc.GetService(typeof(IGroupSecurityService));
            Identity[] appGroups = gss.ListApplicationGroups(rootFolder);
            ArrayList RtGroup = new ArrayList();
            
            for (int k = 0; k < appGroups.Length; k++)
            {                
                Identity group = appGroups[k];
                //Console.WriteLine(k +"  " +group.DisplayName + "**************************");
                if ("Real time".Equals(group.DisplayName))
                {
                    Identity[] groupMembers = gss.ReadIdentities(SearchFactor.Sid, new string[] { group.Sid }, QueryMembership.Expanded);
                    foreach (Identity member in groupMembers)
                    {
                        if (member.Members != null)
                        {
                            //Console.WriteLine(member.Members.Length);
                            //Thread.Sleep(3000);
                            foreach (string memberSid in member.Members)
                            {
                                Identity memberInfo = gss.ReadIdentity(SearchFactor.Sid, memberSid, QueryMembership.Expanded);
                                //Console.WriteLine(memberInfo.DisplayName);
                                RtGroup.Add(memberInfo.DisplayName);
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
            
            System.Console.WriteLine("-----------");
            foreach (Changeset changeset in changesets)
            {
                Changeset cs = changeset;
                System.Console.WriteLine(cs.ChangesetId + " *** " + cs.CommitterDisplayName);
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
                System.Console.WriteLine(branch);

                //associatedWork-----------------------------------
                AssociatedWorkItemInfo[] work = cs.AssociatedWorkItems;
                string linkid = "";
                string linktype = "";
                if (work.Length == 1)
                {
                    linkid = (work[0].Id).ToString();
                    linktype = work[0].WorkItemType;
                }
                else
                {
                    for (int k = 0; k < work.Length; k++)
                    {
                        if (!("Task".Equals(work[k].WorkItemType)))
                        {
                            linkid = linkid + "[" + (k + 1) + "] " + (work[k].Id).ToString() + "  ";
                            linktype = linktype + "[" + (k + 1) + "] " + work[k].WorkItemType + "  ";
                        }
                        if (k == work.Length - 1 && "".Equals(linkid))
                        {
                            linkid = (work[k].Id).ToString();
                            linktype = work[k].WorkItemType;
                        }

                    }
                }

                //teamInfo-----------------------------------
                string teamName = "othres";
                if (RtGroup.Contains(cs.CommitterDisplayName))
                {
                    teamName = "Real time";
                }
                //-----------------------------------

                sw.WriteLine("" + cs.ChangesetId + ',' + cs.CommitterDisplayName + ',' + teamName + ',' + cs.CreationDate + ',' + branch + ',' + linkid + ',' + linktype + ',' + Regex.Replace(cs.Comment, @"[\n\r]", ""));
                        
                //foreach (Change change in changeset.Changes)
                //{
                    //string fileName = change.Item.ServerItem;
                //}

            }
           
            /*原本查询方式
            // changeset 相关信息的获取--------------------------------------
            int latestId = version.GetLatestChangesetId();
            Changeset cs;
            for (int i = latestId; i>0; i--)
            {
                try
                {                 
                    cs = version.GetChangeset(i);
                    if (cs == null)
                    {
                        System.Console.WriteLine("No result");
                        continue;
                    }
                    if ((DateStart.CompareTo(cs.CreationDate)<=0) && ((DateEnd.CompareTo(cs.CreationDate)>=0)))
                    {
                        System.Console.WriteLine("changesetid:" + i + "!!");
                        //branch
                        Change[] thchan = cs.Changes;
                        string branch = thchan[0].Item.ServerItem;
                        String[] braPath = branch.Split('/');
                        if (braPath.Length > 4)
                        {
                            branch ="";
                            for (int k = 1; k < 4; k++)
                            {
                                branch = branch + braPath[k] + "/";
                            }
                            if ("Release".Equals(braPath[3]))
                            {
                                branch = branch + braPath[4];
                            }
                        }
                        System.Console.WriteLine(branch);
                         
                        //associatedWork-----------------------------------
                        AssociatedWorkItemInfo[] work = cs.AssociatedWorkItems;
                        string linkid = "";
                        string linktype = "";
                        if (work.Length == 1)
                        {
                            linkid = (work[0].Id).ToString();
                            linktype = work[0].WorkItemType;
                        }
                        else
                        {
                            for (int k = 0; k < work.Length; k++)
                            {
                                if (!("Task" .Equals(work[k].WorkItemType)))
                                {
                                    linkid = linkid + "[" + (k + 1) + "] " + (work[k].Id).ToString();
                                    linktype = linktype + "[" + (k + 1) + "] " + work[k].WorkItemType;
                                }
                                if (k == work.Length - 1 && "".Equals(linkid))
                                {
                                    linkid = (work[k].Id).ToString();
                                    linktype = work[k].WorkItemType;
                                }

                            }
                        }

                        //teamInfo-----------------------------------
                        string teamName = "othres";
                        if (RtGroup.Contains(cs.CommitterDisplayName))
                        {
                            teamName = "Real time";
                        }
                        //-----------------------------------

                        sw.WriteLine("" + i + ',' + cs.CommitterDisplayName + ',' + teamName + ',' + cs.CreationDate + ',' + branch + ','+ linkid + ',' + linktype + ',' + Regex.Replace(cs.Comment, @"[\n\r]", ""));
                        
                    }
                    if ((DateStart.CompareTo(cs.CreationDate) > 0))
                    {
                        break;
                    }
                }
                catch (ResourceAccessException)
                {
                    continue;
                }
            }
            */
                       			
            sw.Close();
			fs.Close();
            System.Console.WriteLine("******end*****");
			Thread.Sleep(5000);
            
        }
        
    }

}
