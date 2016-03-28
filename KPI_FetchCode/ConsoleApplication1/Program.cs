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
using System.Web.Script.Serialization;
using Newtonsoft.Json.Linq;
using System.Diagnostics;



namespace ConsoleApplication1
{
    class Program
    {
        static string cmdPath = "cmd.exe";
        static DateTime DateStart = DateTime.Now;
        static DateTime DateEnd = DateTime.Now;
       
        static void Main(string[] args)
        {                       
            Console.WriteLine("123!");                                            
            
            string reqUriFolder = "https://tfs.slb.com/tfs/SLB1/Prism/_apis/git/repositories?api-version=1.0";
            HttpWebRequest requestFolder = WebRequest.Create(reqUriFolder) as HttpWebRequest;
            requestFolder.Credentials = new NetworkCredential("YLi102", "Lifor#900516");
            var responseFolderValue = string.Empty;
            using (var response = requestFolder.GetResponse() as HttpWebResponse)
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
                            responseFolderValue = reader.ReadToEnd();
                        }
                }
            }
            JObject folderJson = JObject.Parse(responseFolderValue);
            JArray folderArray = JArray.Parse(folderJson["value"].ToString());
            List<string> folderList = new List<string>();
            foreach (JObject folder in folderArray)
            {
                Console.WriteLine(folder["name"]);
                string folderName = folder["name"].ToString();
                if (folderName.StartsWith("Core"))
                {
                    folderList.Add(folderName);
                }
            }
            Console.WriteLine("************************************************");
            foreach (string fn in folderList)
            {
                Console.WriteLine(fn);
                fetchCode(fn);
            }
            //Console.WriteLine(responseFolderValue);            
            System.Console.WriteLine("      done!");
            System.Console.WriteLine("**********end**********");
          
			Thread.Sleep(5000);
            
        }

        public static void fetchCode(string name)
        {
          
            using (Process p = new Process())
            {
                Boolean newFolder = false;

                p.StartInfo.FileName = cmdPath;

                p.StartInfo.UseShellExecute = false;        //是否使用操作系统shell启动

                p.StartInfo.RedirectStandardInput = true;   //接受来自调用程序的输入信息

                p.StartInfo.RedirectStandardOutput = true;  //由调用程序获取输出信息

                p.StartInfo.RedirectStandardError = false;   //重定向标准错误输出

                p.StartInfo.CreateNoWindow = false;          //不显示程序窗口

                string path = "D:\\Prism_Source\\" + name;

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    newFolder = true;
                }
                p.StartInfo.WorkingDirectory = "d:\\Prism_Source";
                
                p.Start();//启动程序
                if (newFolder)
                {
                    p.StandardInput.WriteLine("cd " + name + "&&git clone https://tfs.slb.com/tfs/SLB1/Prism/_git/" + name);
                }
                p.WaitForExit();
                p.Close();
            }
        }
    }

}
