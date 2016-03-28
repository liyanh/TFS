using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Threading;
using System.Collections;
using System.IO;

namespace Complexity_Method
{
    class Program
    {
        static int overMethod = 0;
        static int allMethod = 0;

        static void Main(string[] args)
        {
            Console.WriteLine("123started");
            FileStream fs = new FileStream(Environment.CurrentDirectory + "\\" + "Complexity_Method.csv", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
            fs.SetLength(0);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine("File Name,Method Name,Complexity");

            string [] fileLocation = new string[]{"\\\\ia-env\\PrismMetrics\\SM_csharp_notest.xml", "\\\\ia-env\\PrismMetrics\\jscr.xml"};
            foreach(string path in fileLocation)
            {
                try
                {
                    string myXMLFilePath = path;
                    XmlDocument myXmlDoc = new XmlDocument();
                    myXmlDoc.Load(myXMLFilePath);
                    XmlNodeList methodList;
                    if("\\\\ia-env\\PrismMetrics\\SM_csharp_notest.xml".Equals(path))
                    {
                        methodList = myXmlDoc.GetElementsByTagName("method");
                        allMethod = allMethod + methodList.Count;
                        foreach (XmlNode node in methodList)
                        {
                            if (int.Parse(node.FirstChild.InnerText) > 15)
                            {
                                overMethod = overMethod + 1;
                                sw.WriteLine("" + "," + node.Attributes["name"].Value + "," + node.FirstChild.InnerText);
                                Console.WriteLine(node.Attributes["name"].Value + "," + node.FirstChild.InnerText);
                            } 
                        }
                    }
                    else
                    {
                        methodList = myXmlDoc.GetElementsByTagName("function");
                        allMethod = allMethod + methodList.Count;
                        foreach (XmlNode node in methodList)
                        {
                            foreach (XmlNode comNode in node.ChildNodes)
                            {
                                if ("cyclomatic".Equals(comNode.Name) && int.Parse(comNode.InnerText) > 15)
                                {
                                    overMethod = overMethod + 1;
                                    sw.WriteLine(node.ParentNode.Attributes["path"].Value + "," + node.Attributes["name"].Value + "," + comNode.InnerText);
                                    Console.WriteLine(node.Attributes["name"].Value + "," + comNode.InnerText);
                                }
                            }
                        }
                    }                                      
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }           
            }
            sw.Close();
            fs.Close();
            Console.WriteLine("allMethod: " + allMethod + "    overMethod: " + overMethod);
            Console.WriteLine("*********************end*********************");
            Thread.Sleep(5000);
        }
             
    }
}
