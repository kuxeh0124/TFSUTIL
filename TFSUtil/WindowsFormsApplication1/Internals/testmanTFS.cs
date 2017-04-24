using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using TFSUtil.Internals;

namespace TFSUtil.Internals
{    
    class testmanTFS
    {
        public Dictionary<string, string> getSuiteList = new Dictionary<string, string>();                       
        public Dictionary<string, string> getTCListFromSuite = new Dictionary<string, string>();
        public Dictionary<string, Dictionary<string,string>> getTCInfo = new Dictionary<string, Dictionary<string, string>>();
        List<Dictionary<string, string>> getDicList = new List<Dictionary<string, string>>();
        
        public List<string> xmlTestCaseFields = new List<string>();
        public Dictionary<string, string[]> tcFieldAllowedValues = new Dictionary<string, string[]>();
        static string getSuiteQry = "Select * From TestSuite";
        public void GetTestSuites(string getSuiteNum="")
        {
            if (getSuiteNum.Length > 0)
            {
                getSuiteQry = "Select * From TestSuite where Id=" + getSuiteNum;
            }
            else
            {
                getSuiteQry = "Select * From TestSuite";
            }
            foreach (ITestSuiteBase ts in tsCollection)
            {
                List<string> currentDirectory = new List<string>();
                ITestSuiteBase getParent = null;
                if (!ts.IsRoot)
                {                
                    currentDirectory.Add(ts.Title);                    
                    getParent = ts.Parent;
                    currentDirectory.Add(getParent.Title);
                    while (getParent.Title != ts.Plan.Name)
                    {
                        getParent = getParent.Parent;
                        currentDirectory.Add(getParent.Title);
                    }
                }
                if (ts.TestCaseCount > 0)
                {
                    string suitePath = "";
                    for (int x = currentDirectory.Count-1; x >= 0; x--)
                    {
                        suitePath = suitePath + currentDirectory[x] + @"\";                       
                    }
                    getSuiteList.Add(Convert.ToString(ts.Id), suitePath.Remove(suitePath.LastIndexOf('\\')));
                }
            }
        }

        public void GetTestCaseInformation(string tcID)
        {
            loadXMLTCFields();
            //getTCFields.Clear();            

            ITestCase testCase = connectTFS.tfsTeamProject.TestCases.Find(Convert.ToInt32(tcID));
            Dictionary<string, string> getTCFields = new Dictionary<string, string>();
            try
            {                
                foreach (string getField in xmlTestCaseFields)
                {
                    if(getField != "Test Plan")
                    {                    
                        if (!getTCFields.ContainsKey(getField))
                        {
                            getTCFields.Add(getField, Convert.ToString(testCase.WorkItem.Fields[getField].Value));
                        }
                    }
                }
            }
            catch(FieldDefinitionNotExistException)
            {
                int stepCtr = 1;
                string fullStep = "";
                foreach (ITestAction testStep in testCase.Actions)
                {                   
                    ITestStep ts = (ITestStep)testStep;
                    string getStep = ts.Title.ToString();
                    string getExpRes = ts.ExpectedResult.ToString();
                    if(fullStep.Length==0)
                    {
                        fullStep = stepCtr.ToString() + "<+++>" + getStep + "<--->" + getExpRes;
                    }
                    else
                    {
                        fullStep = fullStep + "<...>" + stepCtr.ToString() + "<+++>" + getStep + "<--->" + getExpRes;
                    }
                    
                    stepCtr++;
                }
                getTCFields.Add("getSteps", fullStep);
            }
            getDicList.Add(getTCFields);
            getTCInfo.Add(tcID, getDicList[getDicList.Count-1]);
        }

        public bool CreateTestCaseExtractFile(string path)
        {
            try
            {
                ExcelProcessing xlProc = new ExcelProcessing();
                string actualFileName = "TestCase" + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".xlsx";
                string fullFilePath = path + "\\" + actualFileName;
                string[] arrSNOID = { "SNo", "ID" };
                xlProc.createNewExcelTemplate("TestCase");
                xlProc.generateTestCaseTemplate();
                xlProc.updateAndCopy(fullFilePath, "TestCase", true, actualFileName);
                //Start Populating excel
                string compDestPath = xlProc.destPath + "TestCase\\" + actualFileName;
                File.Copy(fullFilePath, compDestPath, true);
                xlProc.updateTestCaseExcelData(fullFilePath, "TestCase", getTCInfo);
                xlProc.updateAndCopy(fullFilePath, "TestCase", false, actualFileName);
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        public bool CreateTestCaseExtractFile(string path, bool isSpecial)
        {
            try
            {
                ExcelProcessing xlProc = new ExcelProcessing();
                string actualFileName = "TestCase" + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".xlsx";
                string fullFilePath = path + "\\" + actualFileName;
                string[] arrSNOID = { "SNo", "ID" };
                xlProc.createNewExcelTemplate("TestCase");
                new testmanTFS().getTFSTestCaseFields();
                xlProc.updateAndCopy(fullFilePath, "TestCase", true, actualFileName);
                //Start Populating excel
                string compDestPath = xlProc.destPath + "TestCase\\" + actualFileName;
                File.Copy(fullFilePath, compDestPath, true);
                xlProc.updateTestCaseExcelDataSpecFormat(fullFilePath, "TestCase", getTCInfo);
                xlProc.updateAndCopy(fullFilePath, "TestCase", false, actualFileName);
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        public void GetTestCases(string suitWkItmId)
        {
            ITestSuiteEntryCollection allTestCase = connectTFS.tfsTeamProject.TestSuites.Find(Convert.ToInt32(suitWkItmId)).TestCases;
            foreach(ITestSuiteEntry getTC in allTestCase)
            {
                getTCListFromSuite.Add(Convert.ToString(getTC.Id), Convert.ToString(getTC.Id) + " - " + Convert.ToString(getTC.Title));
            }
        }

        public void getTFSTestCaseFields()
        {
            WorkItem workItem = new WorkItem(testCaseWorkItemType);
            List<string> allowedVals = new List<string>();

            loadXMLTCFields();

            foreach (Field dField in workItem.Fields)
            {
                allowedVals.Clear();
                if (xmlTestCaseFields.Contains(dField.Name))
                {

                    if (dField.AllowedValues.Count > 0 || dField.Name == "Area Path" || dField.Name == "Iteration Path")
                    {
                        for (int av = 0; av <= dField.AllowedValues.Count - 1; av++)
                        {
                            allowedVals.Add(dField.AllowedValues[av].ToString());
                        }
                        processCustomListItems(dField, allowedVals);
                        tcFieldAllowedValues[dField.Name] = allowedVals.ToArray();
                    }
                }
            }
        }

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

        public void loadXMLTCFields()
        {
            XDocument xdoc = XDocument.Load(@"References\TestCaseFields.xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;

            foreach (XElement r in xRows)
            {
                if (!xmlTestCaseFields.Contains(r.Value))
                {
                    xmlTestCaseFields.Add(r.Value);
                }
            }
        }

        public static WorkItemType testCaseWorkItemType
        {
            get
            {
                return Globals.workItemTypes["Test Case"];
            }
        }

        public static ITestSuiteCollection tsCollection
        {
            get
            {
                return getTestSuites.Query(getSuiteQry);
            }
        }
        public static ITestSuiteHelper getTestSuites
        {
            get
            {
                return connectTFS.tfsTeamProject.TestSuites;
            }
        }

        public static ITestPlanCollection getTestPlans
        {
            get
            {
                return connectTFS.tfsTeamProject.TestPlans.Query("Select * from TestPlan");
            }
        }
    }
}
