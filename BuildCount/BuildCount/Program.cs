using System;
using System.Threading;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Collections;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;
using Newtonsoft.Json.Linq;

namespace BuildCount
{
    class Program
    {
        static DateTime DateStart = DateTime.Now;
        static DateTime DateEnd = DateTime.Now;    
        static void Main(string[] args)
        {
            Console.WriteLine("123!");
            Boolean startDateBoo = true;
            while (startDateBoo)
            {
                try
                {
                    System.Console.Write("Input the start time (yyyy-mm-dd): ");
                    string startdate = Console.ReadLine();
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
                string enddate = Console.ReadLine();
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
            string allReqUri = "https://tfs.slb.com/tfs/SLB1/Prism/_apis/build/builds?api-version=2.0&minfinishtime=" + DateStart + "&maxfinishtime=" + DateEnd;
            string failedReqUri = "https://tfs.slb.com/tfs/SLB1/Prism/_apis/build/builds?api-version=2.0&minfinishtime=" + DateStart + "&maxfinishtime=" + DateEnd + "&resultFilter=failed";
            string succeededReqUri = "https://tfs.slb.com/tfs/SLB1/Prism/_apis/build/builds?api-version=2.0&minfinishtime=" + DateStart + "&maxfinishtime=" + DateEnd + "&resultFilter=succeeded";
            //string canceledReqUri = "https://tfs.slb.com/tfs/SLB1/Prism/_apis/build/builds?api-version=2.0&minfinishtime=" + DateStart + "&maxfinishtime=" + DateEnd + "&resultFilter=canceled";


            Console.WriteLine("All Builds: " + countBuilds(allReqUri));
            Console.WriteLine("Failed Builds: " + countBuilds(failedReqUri));
            Console.WriteLine("Succeeded Builds: " + countBuilds(succeededReqUri));
            //Console.WriteLine("Canceled Builds: " + countBuilds(canceledReqUri));
            Thread.Sleep(50000);
        }

        private static string countBuilds(string reqUri)
        {           
            HttpWebRequest request = WebRequest.Create(reqUri) as HttpWebRequest;           
            request.Credentials = new NetworkCredential("YLi102", "Lifor#900516");
            var result = request.GetResponse() as HttpWebResponse;
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
            return respJson["count"].ToString();
        }
    }
}
