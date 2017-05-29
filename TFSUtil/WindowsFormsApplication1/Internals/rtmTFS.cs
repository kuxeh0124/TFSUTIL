using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using TFSUtil.Internals;

namespace TFSUtil.Internals
{
    class rtmTFS
    {
        static string strWiql="";
        static string specWiql = "";
        public List<string> getRTMFields = new List<string>();
        public List<string> getnewRTMFields = new List<string>();
        public Dictionary<string, string> getRTMMappedFields = new Dictionary<string, string>();
        public void loadRTMFields()
        {
            WorkItem dumWi = new WorkItem(reqWiType);
            foreach(Field rField in dumWi.Fields)
            {
                getRTMFields.Add(rField.Name);
            }
            getRTMFields.Add("Linked Test Case IDs and Name");
            getRTMFields.Add("Linked Child Requirement");
            getRTMFields.Add("Linked Other WorkItems");
            getRTMFields.Add("Ignore");
        }

        public void loadNewFields()
        {
            foreach(KeyValuePair<string,string> entry in getRTMMappedFields)
            {
                if (entry.Value == "Ignore")
                {
                    getnewRTMFields.Add(entry.Key);
                }
                else
                {
                    getnewRTMFields.Add(entry.Value);
                }                
            }            
        }
        public void loadRequirements()
        {
            string wiqlDefault = "Select * from WorkItems where [System.WorkItemType]='Requirement'" +
                        " and [System.TeamProject]='" + Globals.getProjectCol.Name.ToString() + "'";

            strWiql = wiqlDefault;
            loadNewFields();
            int snoCtr = 0;        
            Dictionary<string, string> rtmDic = new Dictionary<string, string>();
            foreach (WorkItem wi in wiCol)
            {
                string currentString = "";
                foreach (string destFld in getnewRTMFields)
                {
                    switch (destFld)
                    {
                        case "SNO": case "S/No": case "Sno":
                            snoCtr++;
                            if (currentString.Length == 0) { currentString = Convert.ToString(snoCtr); }
                            else { currentString = currentString + "<nxtData>" + Convert.ToString(snoCtr); }
                            rtmDic[destFld] = currentString;
                            break;
                        case "ID": case "History":
                            int maxRevisions = wi.Revisions.Count - 1;
                            string getString = "";
                            for (int x = maxRevisions; x >= 0; x--)
                            {
                                if (Convert.ToString(wi.Revisions[x].Fields["History"].Value).Length > 0)
                                {
                                    getString = " - " + Convert.ToString(wi.Revisions[x].Fields["History"].Value);
                                }
                            }
                            if (currentString.Length == 0) { currentString = Convert.ToString(wi.Id + getString); }
                            else { currentString = currentString + "<nxtData>" + Convert.ToString(wi.Id + getString); }
                            rtmDic[destFld] = currentString;
                            break;
                        case "Title":
                            if (currentString.Length == 0) { currentString = Convert.ToString(wi.Title); }
                            else { currentString = currentString + "<nxtData>" + Convert.ToString(wi.Title); }
                            rtmDic[destFld] = currentString;
                            break;
                        case "Linked Test Case IDs and Name":
                            string getFullTCInfo = "";
                            if (wi.Links.Count > 0)
                            {
                                string getTestCaseID = "";
                                string getTestCaseTitle = "";
                                foreach (RelatedLink lc in wi.Links)
                                {
                                    //only taking requirement children and not parents
                                    specWiql = "Select * from WorkItems where [ID] = " + lc.RelatedWorkItemId;
                                    if (Convert.ToString(getAllWIDetail(lc.RelatedWorkItemId).Type.Name) == "Test Case")
                                    {
                                        getTestCaseID = Convert.ToString(getAllWIDetail(lc.RelatedWorkItemId).Fields["Test Case ID"].Value);
                                        getTestCaseTitle = Convert.ToString(getAllWIDetail(lc.RelatedWorkItemId).Fields["Title"].Value);
                                    }
                                    if (getFullTCInfo.Length == 0) { getFullTCInfo = getTestCaseID + " - " + getTestCaseTitle; }
                                    else { getFullTCInfo = getFullTCInfo + "\n" + getTestCaseID + " - " + getTestCaseTitle; }
                                }
                            }
                            if (currentString.Length == 0) { currentString = Convert.ToString(getFullTCInfo); }
                            else { currentString = currentString + "<nxtData>" + Convert.ToString(getFullTCInfo); }
                            rtmDic[destFld] = currentString;
                            break;
                        case "Linked Child Requirement":
                            string getFullReqInfo = "";
                            if (wi.Links.Count > 0)
                            {
                                string getReqID = "";
                                string getReqTitle = "";
                                foreach (RelatedLink lc in wi.Links)
                                {
                                    //only taking requirement children and not parents
                                    specWiql = "Select * from WorkItems where [ID] = " + lc.RelatedWorkItemId;
                                    if (Convert.ToString(getAllWIDetail(lc.RelatedWorkItemId).Type.Name) == "Test Case")
                                    {
                                        getReqID = Convert.ToString(getAllWIDetail(lc.RelatedWorkItemId).Fields["Test Case ID"].Value);
                                        getReqTitle = Convert.ToString(getAllWIDetail(lc.RelatedWorkItemId).Fields["Title"].Value);
                                    }
                                    if (getFullReqInfo.Length == 0) { getFullTCInfo = getReqID + " - " + getReqTitle; }
                                    else { getFullReqInfo = getFullReqInfo + "\n" + getReqID + " - " + getReqTitle; }
                                }
                            }
                            if (currentString.Length == 0) { currentString = Convert.ToString(getFullReqInfo); }
                            else { currentString = currentString + "<nxtData>" + Convert.ToString(getFullReqInfo); }
                            rtmDic[destFld] = currentString;
                            break;
                        default:
                            if (currentString.Length == 0) { currentString = Convert.ToString("<nxtData>"); }
                            else { currentString = currentString + "<nxtData>" + Convert.ToString("<nxtData>"); }
                            rtmDic[destFld] = currentString;
                            break;
                    }
                }
            }
        }
        
        public WorkItem getAllWIDetail(int getWorkItemID)
        {
            try
            {
                return specCol[0];
            }
            catch
            {
                return null;
            }
        }

        public void loadReferenceFile()
        {

        }

        public static WorkItemCollection specCol
        {
            get
            {
                return Globals.workItemStore.Query(specWiql);
            }
        }

        public static WorkItemType reqWiType
        {
            get
            {
                return Globals.workItemTypes["Requirement"];
            }
        }

        public static WorkItemType tcWiType
        {
            get
            {
                return Globals.workItemTypes["Test Case"];
            }
        }

        public static WorkItemCollection wiCol
        {
            get
            {
                return Globals.workItemStore.Query(strWiql);
            }
        }
    }
}