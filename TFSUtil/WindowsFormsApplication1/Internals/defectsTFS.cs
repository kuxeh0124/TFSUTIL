using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFSUtil.Internals;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections;
using TFSUtil;
using System.IO;

namespace TFSUtil.Internals
{
    class defectsTFS
    {
        connectTFS con = new connectTFS();
        string projectName = connectTFS.tfsTeamProject.TeamProjectName.ToString();
        public List<string> xmlDefectFields = new List<string>();
        public List<string> xmlDefectFieldsVal = new List<string>();
        public Dictionary<string, string[]> fieldAllowedValues = new Dictionary<string, string[]>();
        public Dictionary<string, string[]> extractDefectValues = new Dictionary<string, string[]>();
        public List<string> listQueries = new List<string>();
        public List<string> listWorkItemValues = new List<string>();
        public Dictionary<string, string[]> reportDic = new Dictionary<string, string[]>();
        public static Dictionary<string, string> queryValue = new Dictionary<string, string>();
        public static string getSuccess { get; set; }
        public static string getTotalUpload { get; set; }
        

        /// <summary>
        /// Gets all the defect fields from the xml template
        /// </summary>
        public void getTFSDefectFields()
        {
            WorkItem workItem = new WorkItem(workItemType);
            List<string> allowedVals = new List<string>();

            loadXMLDefectFields();

            foreach (Field dField in workItem.Fields)
            {
                allowedVals.Clear();
                if (xmlDefectFields.Contains(dField.Name))
                {

                    if (dField.AllowedValues.Count > 0 || dField.Name == "Area Path" || dField.Name == "Iteration Path")
                    {
                        for (int av = 0; av <= dField.AllowedValues.Count - 1; av++)
                        {
                            allowedVals.Add(dField.AllowedValues[av].ToString());
                        }
                        processCustomListItems(dField, allowedVals);
                        fieldAllowedValues[dField.Name] = allowedVals.ToArray();
                    }
                }
            }

            QueryFolder queryFolder = qHeirarchy as QueryFolder;
            QueryItem queryItem = queryFolder["My Queries"];
            queryFolder = queryItem as QueryFolder;
            foreach (var item in queryFolder)
            {
                Guid queryId = FindQuery(queryFolder, item.Name);
                QueryDefinition queryDefinition = Globals.workItemStore.GetQueryDefinition(queryId);
                if (queryDefinition.QueryText.Contains("[System.WorkItemType] = 'Bug'"))
                {
                    if (!listQueries.Contains(item.Name))
                    {
                        listQueries.Add(item.Name);
                    }
                    if (!queryValue.ContainsKey(item.Name))
                    {
                        queryValue.Add(item.Name, queryDefinition.QueryText);
                    }                    
                }
            }
        }

        /// <summary>
        /// Gets the query from the query folder of the user that is using the tool
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="queryName"></param>
        /// <returns></returns>
        private static Guid FindQuery(QueryFolder folder, string queryName)
        {
            foreach (var item in folder)
            {
                if (item.Name.Equals(queryName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return item.Id;
                }

                var itemFolder = item as QueryFolder;
                if (itemFolder != null)
                {
                    var result = FindQuery(itemFolder, queryName);
                    if (!result.Equals(Guid.Empty))
                    {
                        return result;
                    }
                }
            }
            return Guid.Empty;
        }

        /// <summary>
        /// Adds custom LOVs for the template excel file
        /// </summary>
        /// <param name="dField"></param>
        /// <param name="vals"></param>
        private void processCustomListItems(Field dField, List<string> vals)
        {
            if (dField.Name == "State")
            {
                vals.Clear();
                if (vals.Count < 6)
                {
                    vals.Add("New");
                    vals.Add("Assigned");
                    vals.Add("OnHold");
                    vals.Add("Rejected");
                    vals.Add("Resolved");
                    vals.Add("Closed");
                    vals.Add("Reopened");
                }
            }

            if (dField.Name == "Area Path")
            {
                vals.Clear();
                foreach (Node getArea in Globals.getProjectCol.AreaRootNodes)
                {
                    if (getArea.HasChildNodes)
                    {
                        foreach (Node children in getArea.ChildNodes)
                        {
                            vals.Add(children.Path.ToString());
                        }
                    }
                    else
                    {
                        vals.Add(getArea.Path.ToString());
                    }
                }
            }

            if (dField.Name == "Iteration Path")
            {
                vals.Clear();
                foreach (Node getIteration in Globals.getProjectCol.IterationRootNodes)
                {
                    vals.Add(getIteration.Path.ToString());
                }
            }
        }

        /// <summary>
        /// Loads the defect fields into list
        /// </summary>
        public void loadXMLDefectFields()
        {
            XDocument xdoc = XDocument.Load(@"References\" + Globals.getDefectFieldsFromSetting + ".xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;
            
            foreach (XElement r in xRows)
            {
                if (!xmlDefectFields.Contains(r.Value))
                {
                    xmlDefectFields.Add(r.Value);
                }                    
            }
        }

        /// <summary>
        /// Loads the defect fields into a list to be used for validation
        /// </summary>
        public void loadXMLDefectFieldsForValidation()
        {
            XDocument xdoc = XDocument.Load(@"References\" + Globals.getDefectFieldsFromSetting + ".xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;
            xmlDefectFieldsVal.Add("SNo");
            xmlDefectFieldsVal.Add("ID");
            foreach (XElement r in xRows)
            {
                if (!xmlDefectFieldsVal.Contains(r.Value))
                {
                    xmlDefectFieldsVal.Add(r.Value);
                }
            }
        }

        
        /// <summary>
        /// Loads the excel file into TFS
        /// </summary>
        /// <param name="fileName"></param>
        public void loadIntoTFS(string fileName)
        {
            //Dictionary<string[], string[]> dicDefectDetails = new Dictionary<string[], string[]>();
            ExcelProcessing xlProc = new ExcelProcessing();
            xlProc.archiveAndCopy(fileName, "Defect");
            xlProc.readExcelData(fileName, "Defect");
            int getTotalitems = xlProc.dicExData["ID"].Count() - 1;
            bool addDefect = false;
            string dTID = "";
            int getId = 0;
            int successCtr = 0;
            WorkItem workItem = null;
            try
            {
                for (int i = 0; i <= getTotalitems; i++)
                {
                    foreach (KeyValuePair<string, string[]> xlDefects in xlProc.dicExData)
                    {
                        if (xlDefects.Key != "SNo")
                        {
                            if (xlDefects.Key == "ID")
                            {
                                if (xlDefects.Value[i] == null)
                                {
                                    addDefect = true;
                                }
                                if (addDefect)
                                {
                                    workItem = new WorkItem(workItemType);
                                    addDefect = false;
                                }
                                else
                                {
                                    dTID = xlDefects.Value[i];
                                    workItem = Globals.workItemStore.GetWorkItem(Int32.Parse(dTID));
                                    addDefect = false;
                                }
                            }
                            else
                            {

                                if (Convert.ToString(workItem.Fields[xlDefects.Key].Value) != Convert.ToString(xlDefects.Value[i]))
                                {
                                    if (xlDefects.Key == "State")
                                    {
                                        if (Convert.ToString(workItem.Fields[xlDefects.Key].Value) == "Closed")
                                        {
                                            successCtr++;
                                            break;
                                        }
                                    }
                                    if (xlDefects.Key == "History")
                                    {
                                        int maxRevisions = workItem.Revisions.Count - 1;
                                        if (maxRevisions >= 0)
                                        {
                                            if (Convert.ToString(workItem.Revisions[maxRevisions].Fields[xlDefects.Key].Value) !=
                                                Convert.ToString(xlDefects.Value[i]))
                                            {
                                                workItem.Fields[xlDefects.Key].Value = xlDefects.Value[i];
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (Convert.ToString(workItem.Fields[xlDefects.Key].Name).Contains("Date") &&
                                            !String.IsNullOrEmpty(Convert.ToString(xlDefects.Value[i])))
                                        {
                                            if (!Convert.ToString(workItem.Fields[xlDefects.Key].Name).Contains("Resolved"))
                                            {
                                                DateTime getUndDate = DateTime.FromOADate(Convert.ToDouble(xlDefects.Value[i]));
                                                //string newDateTime = getUndDate.ToString("M/d/yyyy");
                                                workItem.Fields[xlDefects.Key].Value = getUndDate;
                                            }
                                        }
                                        else
                                        {
                                            if(!String.IsNullOrEmpty(Convert.ToString(xlDefects.Value[i])))
                                            {
                                                workItem.Fields[xlDefects.Key].Value = Convert.ToString(xlDefects.Value[i]).Replace("\n", "<br>").Replace("\r\n", "<br>");
                                            }                                            
                                        }

                                    }
                                }
                                else
                                {
                                    if (xlDefects.Key == "State")
                                    {
                                        if (Convert.ToString(workItem.Fields[xlDefects.Key].Value) == "Closed")
                                        {                                            
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //ArrayList result = workItem.Validate();
                    if (validateWorkItem(workItem))
                    {
                        successCtr++;
                        workItem.Save();
                        getId = Convert.ToInt32(workItem.Fields["ID"].Value);
                        extractInformationFromDefect(workItem, fileName, xlProc.dicExData, xlProc, i);                        
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                getSuccess = Convert.ToString(successCtr);
                getTotalUpload = Convert.ToString(getTotalitems + 1);
                xlProc.updateReport(Globals.getReportPath, reportDic);
            }
        }

        private bool validateWorkItem(WorkItem wi)
        {
            ArrayList result = wi.Validate();
            if (result.Count == 0)
            {
                if (!reportDic.ContainsKey(Convert.ToString(wi.Id)))
                {
                    reportDic.Add(Convert.ToString(wi.Id), new string[] { Convert.ToString(wi.Id), "Update Successful" });
                }
                else
                {
                    reportDic[Convert.ToString(wi.Id)] = new string[] { Convert.ToString(wi.Id), "Update Successful" };
                }
                return true;
            }
            else
            {
                int ctr = 1;
                List<string> arrGetVals = new List<string>();
                foreach(Field getResult in result)
                {
                    arrGetVals.Add(Convert.ToString(wi.Id));
                    arrGetVals.Add(Convert.ToString(getResult.Name));
                    arrGetVals.Add(Convert.ToString(getResult.OriginalValue));
                    arrGetVals.Add(Convert.ToString(getResult.Value));
                    arrGetVals.Add(Convert.ToString(getResult.Status));
                    reportDic.Add(Convert.ToString(wi.Id) + "-" + ctr, arrGetVals.ToArray());
                    arrGetVals.Clear();
                    ctr++;
                }
                return false;
            }
        }

        //Used for Extraction based on defect filters and my queries
        public bool extractInformationFromDefect(string strWiql, string path, string[] filterType=null, string[] filterValue=null)
        {
            ExcelProcessing xlProc = new ExcelProcessing();
            string actualFileName = "Defect" + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".xlsx";
            string fullFilePath = path + "\\" + actualFileName;
            string[] arrSNOID = { "SNo", "ID" };
            xlProc.createNewExcelTemplate("Defect");
            xlProc.generateDefectTemplate();
            xlProc.updateAndCopy(fullFilePath, "Defect", true, actualFileName);

            string wiqlDefault = "Select * from WorkItems where [System.WorkItemType]='Bug'" +
                                    " and [System.TeamProject]='" + Globals.getProjectCol.Name.ToString() + "'";
            int snoCtr = 1;
            string toAdd = "";
            int getFV = 0;
            try
            {
                if (strWiql.Length == 0)
                {
                    strWiql = wiqlDefault;
                }                
                if (filterType.Count() > 0)
                {
                    foreach (string ft in filterType)
                    {
                        switch(ft)
                        {
                            case "combo_TestingPhase":
                                toAdd = toAdd + " and [Microsoft.VSTS.Common.CrimsonLogic.TestingPhase] = '" + filterValue[getFV] + "'";
                                break;
                            case "combo_User":
                                toAdd = toAdd + " and [System.AssignedTo] = '" + filterValue[getFV] + "'";
                                break;
                            case "combo_State":
                                toAdd = toAdd + " and [System.State] = '" + filterValue[getFV] + "'";
                                break;
                        }
                        getFV++;
                    }
                    strWiql = strWiql + toAdd;
                }                
                loadXMLDefectFields();
                strWiql = parseWiql(strWiql);
                WorkItemCollection witCollection =  Globals.workItemStore.Query(strWiql);

                foreach (string getField in xmlDefectFields)
                {
                    foreach (WorkItem workItem in witCollection)
                    {
                        Globals.AddToDictionary(extractDefectValues, getField, Convert.ToString(workItem.Fields[getField].Value), listWorkItemValues);                        
                    }
                    listWorkItemValues.Clear();
                }
                foreach (string snoID in arrSNOID)
                {
                    foreach (WorkItem workItem in witCollection)
                    {
                        if (snoID == "SNo")
                        {
                            Globals.AddToDictionary(extractDefectValues, snoID, snoCtr.ToString(), listWorkItemValues);
                            snoCtr++;
                        }
                        else
                        {
                            Globals.AddToDictionary(extractDefectValues, snoID, Convert.ToString(workItem.Fields[snoID].Value), listWorkItemValues);
                        }

                    }
                    listWorkItemValues.Clear();
                }
                string compDestPath = xlProc.destPath + "Defect\\" + actualFileName;
                File.Copy(fullFilePath, compDestPath, true);
                xlProc.updateExcelData(fullFilePath, "Defect", extractDefectValues);
                xlProc.updateAndCopy(fullFilePath, "Defect", false, actualFileName);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return false;
            }
        }

        //Used in conjunction with Upload so that it will auto update
        public void extractInformationFromDefect(WorkItem wi, string fileName, Dictionary<string, string[]> getDic, 
            ExcelProcessing xlProc, int valCtr = 0)
        {
            try
            {
                loadXMLDefectFields();
                getDic["ID"][valCtr] = Convert.ToString(Convert.ToInt32(wi.Fields["ID"].Value));
                for (int x = 0; x <= xmlDefectFields.Count-1; x++)
                {
                    getDic[xmlDefectFields[x]][valCtr] = Convert.ToString(wi.Fields[xmlDefectFields[x]].Value);
                }               
                xlProc.updateExcelData(fileName, "Defect", getDic, valCtr);
                xlProc.updateAndCopy(fileName, "Defect");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }            
        }

        //Used for Refresh
        public void extractInformationFromDefect(string fileName)
        {
            ExcelProcessing xlProc = new ExcelProcessing();
            xlProc.getDefectTFSIDs(fileName, "Defect");
            int getTotalId = xlProc.getAllID.Count-1;
            loadXMLDefectFields();
            List<WorkItem> getWorkItems = new List<WorkItem>();
            int snoCtr = 0;
            string[] arrSNOID = { "SNo", "ID" };
            try
            {
                for (int valCtr = 0; valCtr <= getTotalId; valCtr++)
                {
                    getWorkItems.Add(Globals.workItemStore.GetWorkItem(xlProc.getAllID[valCtr]));
                }

                foreach (string getField in xmlDefectFields)
                {
                    foreach(WorkItem workItem in getWorkItems)
                    {                        
                        Globals.AddToDictionary(extractDefectValues, getField, Convert.ToString(workItem.Fields[getField].Value), listWorkItemValues);
                    }
                    listWorkItemValues.Clear();
                }           

                foreach (string snoID in arrSNOID)
                {
                    foreach (WorkItem workItem in getWorkItems)
                    {                     
                        if (snoID == "SNo")
                        {
                            Globals.AddToDictionary(extractDefectValues, snoID, snoCtr.ToString(), listWorkItemValues);
                            snoCtr++;
                        }
                        else
                        {
                            Globals.AddToDictionary(extractDefectValues, snoID, Convert.ToString(workItem.Fields[snoID].Value), listWorkItemValues);
                        }

                    }
                    listWorkItemValues.Clear();
                }           
                xlProc.updateExcelData(fileName, "Defect", extractDefectValues);
                xlProc.updateAndCopy(fileName, "Defect");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
        private string parseWiql(string wiql)
        {
            string newString = "";
            newString = wiql;
            newString = wiql.Replace("@project", "'" + projectName + "'");
            return newString;
        }
        
        public static WorkItemType workItemType
        {
            get
            {
                return Globals.workItemTypes["Bug"];
            }
        }

        public static QueryHierarchy qHeirarchy
        {
            get
            {
                return Globals.getProjectCol.QueryHierarchy;
            }
        }

        public static Dictionary<string, string> getMyQueryText
        {
            get
            {
                return queryValue;
            }
        }

    }
}
