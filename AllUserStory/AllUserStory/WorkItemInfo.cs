using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

namespace AllUserStory
{
    class WorkItemInfo
    {
        public static List<WorkItem> getWorkItemChild(WorkItemStore workItemStore, WorkItem wi)
        {
            List<WorkItem> linkedItem = new List<WorkItem>();
            WorkItemLinkCollection links = wi.WorkItemLinks;
            foreach (WorkItemLink wil in links)
            {
                //search for the correct link
                if (wil.LinkTypeEnd.Name.Equals("Child"))
                {
                    WorkItem childItem = workItemStore.GetWorkItem(wil.TargetId);
                    linkedItem.Add(childItem);
                }
            }
            return linkedItem;
        }

        public static Revision getWorkItemRevision(WorkItem item)
        {
            int revNum = item.Revision;
            Revision revis = item.Revisions[revNum - 1];
            return revis;
        }
        /*
        foreach (Field fi in childItem.Fields)
        {
            Console.WriteLine(fi.Name);
            Console.WriteLine(revChild.Fields[fi.Name].Value);
       }
       */
    }
}
