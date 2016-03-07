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
        static VersionControlServer version=null;
        static string starttime = "";
        static string endtime = "";
        public static int BinarySearch(int start, int end)
        {
            
            int mid = (start + end)/2;
            Changeset cs=null; 
            try
            {
                while (true)
                {
                    if (version.GetChangeset(mid) != null)
                    {
                        cs = version.GetChangeset(mid);
                        break;
                    }
                    else { 
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
                return BinarySearch(start+1,end);
            }
  
   
        }      
        
        static void Main(string[] args)
        {

            
            Console.WriteLine("123!");
          //  StreamWriter sw = File.AppendText(System.Environment.CurrentDirectory + "\\AnalysisResult.csv");
            FileStream fs = new FileStream(Environment.CurrentDirectory + "\\AnalysisResult.csv", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
            fs.SetLength(0);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine("changesetID,owner,createDate,linkID,linktype,comment");
           
            System.Console.Write("Input the start time (yyyy-mm-dd): ");
            starttime= Console.ReadLine();

            try
            {
                //DateTime.TryParse(string.Format("{yyyy-MM-DD}:00", starttime),out dt);

                DateStart = DateTime.ParseExact(starttime, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                Console.WriteLine("Unable to convert to date");
                System.Console.Write("Input the start time (yyyy-mm-dd): ");
                starttime = Console.ReadLine();
                

            }
            System.Console.Write("Input the end time (yyyy-mm-dd): ");
            endtime = Console.ReadLine();
            try
            {
                DateEnd = DateTime.ParseExact(endtime, "yyyy-MM-dd", null);
            }
            catch (FormatException)
            {
                Console.WriteLine("Unable to convert to date");
                System.Console.Write("Input the end time (yyyy-mm-dd): ");
                endtime = Console.ReadLine();
            }

            TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri("https://tfs.slb.com/tfs/Real-Time_Collection/"));
            //版本控制
            
            version = tpc.GetService(typeof(VersionControlServer)) as VersionControlServer;
            int latestId = version.GetLatestChangesetId();
      //      int mid=BinarySearch(1, latestId);
            List<Changeset> changesetList = new List<Changeset>();
       /*     int j=0;
            for (int i = mid; version.GetChangeset(i).CreationDate >= DateStart;i-- )
            {
                j = i;
            }
        * */
            
            for (int i = latestId; i>0; i--)
            {
                
                try
                {
                    Changeset cs;
                    while(true)
                    {
                        if (version.GetChangeset(i) != null)
                        {
                            cs = version.GetChangeset(i);

                            break;
                        }
                        else
                        {
                            i--;
                        }
                    }
                    
                    if ((DateStart.CompareTo(cs.CreationDate)<=0) && (DateEnd.CompareTo(cs.CreationDate)>=0))
                    {
                        // System.Console.WriteLine("info:"+cs.AssociatedWorkItems);
                        AssociatedWorkItemInfo[] work = cs.AssociatedWorkItems;
                        
                        int linkid=0;
                        string linktype="";
                        for (int k = 0; k < work.Length; k++)
                        {
                            linkid = work[k].Id;
                          

                            linktype = work[k].WorkItemType;
                           
                        }
                        if (work.Length == 0)
                        {
                            System.Console.WriteLine("changesetid:" + i + "!!");
                            
                            string input = "" + i + ',' + cs.CommitterDisplayName + ',' + cs.CreationDate + ',' + "" + ',' + "" + ',' + Regex.Replace(cs.Comment,@"[\n\r]","");
                            sw.WriteLine(input);
                            

                        }
                        else
                        {
                            sw.WriteLine("" + i + ',' + cs.CommitterDisplayName + ',' + cs.CreationDate + ',' + linkid + ',' + linktype + ',' + Regex.Replace(cs.Comment, @"[\n\r]", ""));
                        }
                    }
                    if ((DateStart.CompareTo(cs.CreationDate) > 0))
                    {
                        break;
                    }

                }
                
                catch (ResourceAccessException)
                {
                    System.Console.WriteLine("ResourceAccessException");
                    System.Console.WriteLine("Can't connect Server!!!");
                    break;
                }
                
            }
            sw.Close();
            fs.Close();
            
            Thread.Sleep(5000);
        }
        
    }

    
}
