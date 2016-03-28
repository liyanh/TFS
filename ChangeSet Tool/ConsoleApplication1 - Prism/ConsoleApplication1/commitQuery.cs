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

namespace ConsoleApplication1
{
    class commitQuery
    {
        public static void commitQuery(JObject comjob)
        {

            string comId = "";
            string aurhou = "";
            string audate = "";
            string committer = "";
            string comdate = "";
            string comment = "";
            string linkedId = "unlinked";
            string linkedType = "unlinked";

            Console.WriteLine(comjob["commitId"].ToString());
            comId = comjob["commitId"].ToString();
            aurhou = comjob["author"]["name"].ToString();
            audate = comjob["author"]["date"].ToString();
            committer = comjob["committer"]["name"].ToString();
            comdate = comjob["committer"]["date"].ToString();
            comment = comjob["comment"].ToString();

            Dictionary<string, string> itemMap = new Dictionary<string, string>();
            Regex itemId = new Regex("#\\d+");
            Match m = itemId.Match(comment);
            if (m.Success)
            {
                string newlinkedId = m.Value.Substring(1);
                if (newlinkedId.Equals(linkedId) == false)
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

                        Console.WriteLine(commitInfo);
                        JObject itemJson = JObject.Parse(commitInfo);
                        linkedType = itemJson["fields"]["System.WorkItemType"].ToString();
                    }
                    catch (Exception)
                    {
                        Console.WriteLine(comId);
                        linkedType = "error";
                    }
                }
            }
            else
            {
                linkedId = "unlinked";
                linkedType = "unlinked";
            }

            if ("unlinked".Equals(linkedId) == false)
            {
                Boolean commitBe = false;
                string usHistoryUri = "https://tfs.slb.com/tfs/SLB1/_apis/wit/workitems/" + linkedId + "/history";
                //Console.WriteLine(commitInfoUri);
                HttpWebRequest historyReq = WebRequest.Create(usHistoryUri) as HttpWebRequest;
                historyReq.Credentials = new NetworkCredential("YLi102", "Lifor#900516");
                var historyInfo = string.Empty;
                try
                {
                    using (var historyResp = historyReq.GetResponse() as HttpWebResponse)
                    {

                        if (historyResp.StatusCode != HttpStatusCode.OK)
                        {
                            var message = String.Format("Request failed. Received HTTP {0}", historyResp.StatusCode);
                            throw new ApplicationException(message);
                        }

                        // grab the response
                        using (var responseStream = historyResp.GetResponseStream())
                        {
                            if (responseStream != null)
                                using (var reader = new StreamReader(responseStream))
                                {
                                    historyInfo = reader.ReadToEnd();
                                }
                        }
                    }

                    //Console.WriteLine(commitInfo);
                    JObject itemJson = JObject.Parse(historyInfo);
                }
                catch (Exception)
                {
                    Console.WriteLine(comId);
                    linkedType = "error";
                }

                sw.WriteLine("" + comId + ',' + aurhou + ',' + team + ',' + audate + ',' + linkedType + ',' + linkedId + ',' + Regex.Replace(comment, @"[\n\r,]", " "));
            }
        }
    }
}
