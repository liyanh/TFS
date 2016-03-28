using System;
using System.Threading;
using System.Collections.Generic;
using System.Text;
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
using System.Web;
using System.Web.Script.Serialization;
using Newtonsoft.Json.Linq;



namespace Prism_Commits
{
    class Program
    {

        static DateTime DateStart = DateTime.Now;
        static DateTime DateEnd = DateTime.Now;
        static string startdate = "";
        static string enddate = "";        
        static string linkedId = "unlinked";
        static string linkedType = "unlinked";

        static void Main(string[] args)
        {
            Console.WriteLine("123!");
            Boolean startDateBoo = true;
            while (startDateBoo)
            {
                try
                {
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
            FileStream fs = new FileStream(Environment.CurrentDirectory + "\\" + "Prism_" + startdate + "---" + enddate + "_Result.csv", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
            fs.SetLength(0);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine("Group,ChangesetID,Owner,CommitDate,LinkedId,LinkedType,Comment");

            string[] serbra = {"Core.Library.Protocol","Core.Library.Sabre","Core.Service.AuditLogging","Core.Service.ChannelData.Listener",
                                "Core.Service.ChannelData.Reader","Core.Service.ChannelData.Writer","Core.Service.Command","Core.Service.DemoApp",
                                "Core.Service.Entitlement","Core.Service.Listener","Core.Service.StreamRouter","Core.Service.SurveryData.Writer",
                                "Core.Service.SurveyData.Listener","Core.Service.SurveyData.Reader"};
            //string[] serbra = { "Core.Service.AuditLogging" };

            foreach (string team in serbra)
            {
                List<List<string>> teamcommit = commitQuery(team);
                foreach (List<string> commit in teamcommit)
                {
                    sw.Write(team + ',');
                    foreach (string val in commit)
                    {
                        sw.Write(val + ',');
                    }
                    sw.WriteLine();
                }
            }
                
            sw.Close();
            fs.Close();
            System.Console.WriteLine("**********end**********");
            Thread.Sleep(5000);
        }

        public static List<List<string>> commitQuery(string team)
        {
            Console.WriteLine(team);
            string reqUri = "https://tfs.slb.com/tfs/SLB1/Prism/_apis/git/repositories/" + team + "/commits?fromDate=" + startdate + "&toDate=" + enddate + "&api-version=1.0";
            HttpWebRequest request = WebRequest.Create(reqUri) as HttpWebRequest;
            request.Credentials = new NetworkCredential("YLi102", "Lifor#900516");
            var responseValue = string.Empty;
            using (var response = request.GetResponse() as HttpWebResponse)
            {

                if (response.StatusCode != HttpStatusCode.OK)
                {
                    var message = String.Format("Request failed. Received HTTP {0}", response.StatusCode);
                    throw new ApplicationException(message);
                }

                // grab the response
                using (var responseStream = response.GetResponseStream())
                {
                    if (responseStream != null)
                        using (var reader = new StreamReader(responseStream))
                        {
                            responseValue = reader.ReadToEnd();
                        }
                }
            }

            JObject respJson = JObject.Parse(responseValue);
            Console.WriteLine("Commits Num: " + respJson["count"]);
            JArray arr = JArray.Parse(respJson["value"].ToString());

            List<List<string>> teamCommit = new List<List<string>>();
            foreach (JObject comjob in arr)
            {
                List<string> cominfo = new List<string>();

                //Console.WriteLine(comjob["commitId"].ToString());                
                cominfo.Add(comjob["commitId"].ToString());
                cominfo.Add(comjob["author"]["name"].ToString());
                cominfo.Add(comjob["committer"]["date"].ToString());
                string comment = comjob["comment"].ToString();

                Regex itemId = new Regex("#\\d+");
                Match m = itemId.Match(comment);
                if (m.Success)
                {
                    string newlinkedId = m.Value.Substring(1);
                    if (newlinkedId.Equals(linkedId) == false)
                    {
                        linkedTypeM(newlinkedId);
                    }
                    
                }
                else
                {
                    linkedId = "unlinked";
                    linkedType = "unlinked";
                }
                cominfo.Add(linkedId);
                cominfo.Add(linkedType);
                comment = comment.Replace(",", ".").Replace("\n",".").Replace("\t",".");
                cominfo.Add(comment);
                teamCommit.Add(cominfo);
            }
            return teamCommit;
        }

        public static void linkedTypeM(string newlinkedId)
        {
            linkedId = newlinkedId;
            string commitInfoUri = "https://tfs.slb.com/tfs/SLB1/_apis/wit/workitems/" + linkedId + "?api-version=1.0";
            //Console.WriteLine(commitInfoUri);
            HttpWebRequest commitReq = WebRequest.Create(commitInfoUri) as HttpWebRequest;
            commitReq.Credentials = new NetworkCredential("YLi102", "Lifor#900516");
            var commitInfo = string.Empty;
            try
            {
                using (var commitResp = commitReq.GetResponse() as HttpWebResponse)
                {
                    if (commitResp.StatusCode != HttpStatusCode.OK)
                    {
                        var message = String.Format("Request failed. Received HTTP {0}", commitResp.StatusCode);
                        throw new ApplicationException(message);
                    }
                    // grab the response
                    using (var responseStream = commitResp.GetResponseStream())
                    {
                        if (responseStream != null)
                        using (var reader = new StreamReader(responseStream))
                        {
                            commitInfo = reader.ReadToEnd();
                        }
                    }
                }
                JObject itemJson = JObject.Parse(commitInfo);
                linkedType = itemJson["fields"]["System.WorkItemType"].ToString();
            }
            catch (Exception)
            {
                Console.WriteLine(newlinkedId);
                linkedId = "error";
                linkedType = "error";
            }
        }               
    }
}
https://tfs.slb.com/tfs/SLB1/Prism/_apis/build/builds?api-version=1.0