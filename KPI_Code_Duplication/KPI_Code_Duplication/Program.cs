using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Threading;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;

namespace KPI_Code_Duplication
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("123started");
            string myXMLFilePath = "\\\\ia-env\\PrismMetrics\\src\\simianout.xml";
            //遍历xml文件的信息
            try
            {                    
                XmlDocument myXmlDoc = new XmlDocument();
                myXmlDoc.Load(myXMLFilePath);

                FileStream fs = new FileStream(Environment.CurrentDirectory + "\\" + "Code_Duplication.csv", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
                fs.SetLength(0);
                StreamWriter sw = new StreamWriter(fs);
                sw.WriteLine("LineCount,Group,DuplicateFile,StartLine,EndLine");

                XmlNodeList setNode = myXmlDoc.GetElementsByTagName("set");
                foreach (XmlNode set in setNode)
                {
                    foreach (XmlNode block in set)
                    {
                        string sf = block.Attributes["sourceFile"].Value;
                        string[] ss = sf.Split('\\');                                                                                             
                        sw.WriteLine(set.Attributes["lineCount"].Value + "," + ss[3] + "," + block.Attributes["sourceFile"].Value + "," + block.Attributes["startLineNumber"].Value + "," + block.Attributes["endLineNumber"].Value);
                    }
                }
                sw.Close();
                fs.Close();                      
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }                                   
            Console.WriteLine("*********************end*********************");
            Thread.Sleep(5000);
        }        

    }
}
