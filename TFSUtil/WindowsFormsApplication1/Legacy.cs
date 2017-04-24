namespace FindTestCase
{
    using Microsoft.TeamFoundation.Client;
    using Microsoft.TeamFoundation.TestManagement.Client;
    using System;
    using System.Collections.Generic;    
    using excelInterop = Microsoft.Office.Interop.Excel;
    using Internals;
    using System.Windows.Forms;
    using UI_Design;
    

    class FindTestCase
    {
        
        public static string fName = "DWH ETL Test Cases Extract";
        public static string defPath = "F:\\ExtractTFS\\";
        //public static String fName = "API Blackbox Test Cases";
        public static int stepStart = 1;
        //public static TFSUDMainWindow mainWin = new TFSUDMainWindow();
        //Test Values
        //IPOS Data Warehouse Test 
        //Suite ID: 32372     
        static List<int> lstTCID = new List<int>();
        static List<string> lstTCStep = new List<string>();
        static List<string> lstTCExpRes = new List<string>();
        static List<string> lstTCSummary = new List<string>();
        static List<string> lstTCPrecondition = new List<string>();
        static List<string> lstTCTextField1 = new List<string>();
        static List<string> lstTCStepNo = new List<string>();
        static List<string> lstTCIDUD = new List<string>();
        static List<int> lstSingID = new List<int>();
        static List<string> lstDateTested = new List<string>();
        static List<string> lstTCOutcome = new List<string>();
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TFSUtilUI());
            //string getTestSuites = "Select * from TestSuite"; 
            connectTFS.connectToTFS();
            //ITestSuiteCollection getAllSuites = connectTFS.tfsTeamProject.TestSuites.Query(getTestSuites);
            //defectsTFS.loadIntoTFS("F:\\TFS\\KARLDefects.xlsx");       
            ExcelProcessing.archiveAndCopy("F:\\TFS\\KARLDefects.xlsx", "Defect");
            defectsTFS.extractInformationFromDefect("F:\\TFS\\KARLDefects.xlsx");
            //ExcelProcessing.createNewExcelTemplate("Defect");
            //ExcelProcessing.generateDefectTemplate();
            //defectsTFS.loadXMLDefectFields();
            //ExcelProcessing.createNewExcelTemplate("Defect");
            //ExcelProcessing.generateDefectTemplate();
            //ExcelProcessing.closeExcelApp();
        

            //call XAML or Win Forms from here
            //TFSUDMainWindow mainWin = new TFSUDMainWindow();
            //mainWin.ShowDialog();
            //getProjectName();
            //This is for extract=================================================================================
            //Aggregate Change
            //ExtractToExcel("IPOS Data Warehouse", 31795, "IPOSDASH - DataWarehouse Aggregate Change v1.1", null, 2);

            //API Data Verification
            // ExtractToExcel("IPOS Data Warehouse", 31859, "IPOSDASH - API Data Verification", null, 1);
            //ExtractToExcel("IPOS Data Warehouse", 31859, "IPOSDASH - API Data Verification", new int[] { 32188 });

            //DWH ETL
            //ExtractToExcel("IPOS Data Warehouse", 31796, "IPOSDASH - DWH ETL Test Cases v0.3", null, 1);
            //ExtractToExcel("IPOS Data Warehouse", 31859, "IPOSDASH - API Data Verification", new int[] { 32188 });
            //This is for extract=================================================================================


            //This is for upload==================================================================================

            //Aggregate Change*************************************************
            //GetExcelDetails("IPOSDASH - DataWarehouse Aggregate Change v1.1");
            //UpdateTestCaseFromExcel("IPOSDASH - DataWarehouse Aggregate Change v1.1", 
            //    "IPOS Data Warehouse", lstSingID.ToArray(), 31795);
            //UpdateTestCaseFromExcel("UpdatedUpload 1320171730", "IPOS Data Warehouse Test", new int[] { 32369, 32373 }, 32372);
            //Aggregate Change*************************************************

            //API Data Verification********************************************
            //GetExcelDetails("IPOSDASH - API Data Verification for upload V 0.1");
            //UpdateTestCaseFromExcel("IPOSDASH - API Data Verification for upload V 0.1",
            //    "IPOS Data Warehouse", lstSingID.ToArray(), 31859);          
            //API Data Verification********************************************

            //DWH ETL**********************************************************
            //GetExcelDetails("IPOS DASH -  DWH ETL Test Cases Extract for UPLOAD v0.4");
            //UpdateTestCaseFromExcel("IPOS DASH -  DWH ETL Test Cases Extract for UPLOAD v0.4",
            //    "IPOS Data Warehouse", lstSingID.ToArray(), 31796);
            //DWH ETL**********************************************************

            //ETE ETL**********************************************************
            //GetExcelDetails("IPOS DASH - End to End Test cases v0.3");
            //UpdateTestCaseFromExcel("IPOS DASH - End to End Test cases v0.3",
            //    "IPOS Data Warehouse", lstSingID.ToArray(), 32695);
            //ETE ETL**********************************************************

            //Test upload****************TESTING PURPOSES**********************
            //GetExcelDetails("IPOS DASH -  DWH ETL Test Cases Extract for UPLOAD");
            //UpdateTestCaseFromExcel("IPOS DASH -  DWH ETL Test Cases Extract for UPLOAD",
            //    "IPOS Data Warehouse Test", lstSingID.ToArray(), 32372);
            //UpdateTestCaseFromExcel("IPOS DASH -  DWH ETL Test Cases Extract for UPLOAD", "IPOS Data Warehouse Test", new int[] { 32647 }, 32372);
            //Test upload****************TESTING PURPOSES**********************

            //This is for upload==================================================================================

        }

        public static void GetExcelDetails(string fName, int[] specTCID = null)
        {
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(defPath + fName + ".xlsx");
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];
            int tcCtr = 0;
            int iSrch = 0;
            lstTCID.Clear();
            lstTCStep.Clear();
            lstTCExpRes.Clear();
            lstTCSummary.Clear();
            lstTCPrecondition.Clear();
            lstTCTextField1.Clear();
            lstTCStepNo.Clear();
            lstTCIDUD.Clear();
            lstSingID.Clear();
            String currentTC = "";
            if (specTCID == null)
            {
                for (iSrch = 1; iSrch <= 90000; iSrch++)
                {
                    if (ws.get_Range("O" + iSrch).Value != null)
                    {
                        try
                        {
                            lstSingID.Add(Int32.Parse(Convert.ToString(ws.get_Range("O" + iSrch).Value)));
                        }
                        catch(FormatException)
                        {

                        }
                    }
                    if (ws.get_Range("M" + iSrch).Value == null)
                    {
                        break;
                    }
                }
                specTCID = lstSingID.ToArray();
            }

            if (specTCID != null)
            {
                foreach (int tcid in specTCID)
                {
                    for (iSrch = 1; iSrch <= 90000; iSrch++)
                    {
                        if (ws.get_Range("O" + iSrch).Value != null)
                        {
                            try
                            {
                                if (Int32.Parse(Convert.ToString(ws.get_Range("O" + iSrch).Value)) == tcid)
                                {
                                    //lstSingID.Add(Int32.Parse(Convert.ToString(ws.get_Range("O" + iSrch).Value)));
                                    currentTC = Convert.ToString(ws.get_Range("F" + iSrch).Value);
                                    tcCtr = iSrch;
                                    break;
                                }
                            }

                            catch (FormatException)
                            {

                            }

                        }
                    }
                    if (currentTC != "")
                    {
                        for (int o = iSrch; o <= 90000; o++)
                        {
                            if (currentTC == Convert.ToString(ws.get_Range("K" + o).Value))
                            {
                                lstTCID.Add(Int32.Parse(Convert.ToString(ws.get_Range("O" + iSrch).Value)));
                                lstTCStep.Add(Convert.ToString(ws.get_Range("M" + o).Value));
                                lstTCExpRes.Add(Convert.ToString(ws.get_Range("N" + o).Value));
                                lstTCSummary.Add(Convert.ToString(ws.get_Range("H" + iSrch).Value));
                                lstTCPrecondition.Add(Convert.ToString(ws.get_Range("I" + iSrch).Value));
                                lstTCTextField1.Add(Convert.ToString(ws.get_Range("J" + iSrch).Value));
                                lstTCStepNo.Add(Convert.ToString(ws.get_Range("L" + o).Value));
                                lstTCIDUD.Add(Convert.ToString(ws.get_Range("K" + o).Value));
                                lstDateTested.Add(Convert.ToString(ws.get_Range("P" + o).Value));
                                lstTCOutcome.Add(Convert.ToString(ws.get_Range("Q" + o).Value));
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
                //flow if for all test cases
            }
            wb.Save();
            wb.Close();
            xlApp.Quit();
        }

        public static void UpdateTestCaseFromExcel(string fName, string tfsProject, int[] specTCID, int suitWkItmId)
        {
            TfsTeamProjectCollection tfctc = new TfsTeamProjectCollection(new Uri("https://dlm.crimsonlogic.com/tfs/Government"));
            ITestManagementService testmanagementService = tfctc.GetService<ITestManagementService>();
            var teamproject = testmanagementService.GetTeamProject(tfsProject);
            var allTestCase = teamproject.TestSuites.Find(suitWkItmId).TestCases;
            Boolean tcFound = false;
            int[] arrTCID = lstTCID.ToArray();
            string[] arrTCStep = lstTCStep.ToArray();
            string[] arrExpRes = lstTCExpRes.ToArray();
            string[] arrSummary = lstTCSummary.ToArray();
            string[] arrPrecond = lstTCPrecondition.ToArray();
            string[] arrTextField1 = lstTCTextField1.ToArray();
            string[] arrStepNo = lstTCStepNo.ToArray();
            string[] arrTCIDUD = lstTCIDUD.ToArray();
            string[] arrDateTested = lstDateTested.ToArray();
            string[] arrTCOutcome = lstTCOutcome.ToArray();
            int getInitialPoint = 0;
            int getLastStepNo = 0;
            int i = 0;
            foreach (var tc in allTestCase)
            {
                foreach (int tcUpdateId in specTCID)
                {
                    if (tc.Id == tcUpdateId)
                    {
                        ITestCase testCase = teamproject.TestCases.Find(tc.Id);
                        for (int x = 0; x <= arrTCID.Length - 1; x++)
                        {
                            if (tc.Id == arrTCID[x] && Int32.Parse(arrStepNo[x]) == 1)
                            {
                                //var testcaseResult = teamproject.TestResults.Query("Select * from TestResult where TestCaseID='" + tc.Id + "'");
                                var testcaseResult = teamproject.TestResults.ByTestId(tc.Id);
                                testCase.Description = arrSummary[x];
                                testCase.CustomFields["Test Case ID"].Value = arrTCIDUD[x];
                                testCase.CustomFields["Pre-Condition"].Value = arrPrecond[x];
                                testCase.CustomFields["Text Field 1"].Value = arrTextField1[x];
                                Console.WriteLine("Updating: " + arrTCIDUD[x]);
                                try
                                {
                                    //testcaseResult[testcaseResult.Count - 1].DateStarted = Convert.ToDateTime(arrDateTested[x]);
                                }
                                catch (ArgumentOutOfRangeException)
                                {
                                    Console.WriteLine("There are no test results for this test case");
                                }

                                getInitialPoint = x;
                                i = getInitialPoint;
                            }
                            if (tc.Id==arrTCID[x])
                            {
                                try
                                {
                                    if (Int32.Parse(arrStepNo[x + 1]) == 1)
                                    {
                                        getLastStepNo = Int32.Parse(arrStepNo[x]);
                                        break;
                                    }
                                }
                                catch (Exception)
                                {
                                    getLastStepNo = Int32.Parse(arrStepNo[x]);
                                }                                
                            }                        }
                        if (testCase.Actions.Count != getLastStepNo)
                        {

                            int getStepDiff = getLastStepNo - testCase.Actions.Count;
                            if (getStepDiff < 0)
                            {
                                int newStepDiff = 0;
                                newStepDiff = testCase.Actions.Count - getLastStepNo;
                                for (int tsd = 1; tsd <= newStepDiff; tsd++)
                                {
                                    testCase.Actions.RemoveAt(tsd - 1);
                                }                              
                            }
                            else
                            {
                                for (int tsd = 1; tsd <= getStepDiff; tsd++)
                                {
                                    ITestStep newStep = testCase.CreateTestStep();
                                    testCase.Actions.Add(newStep);
                                }
                            }

                        }
                        if (testCase.Actions.Count == getLastStepNo)
                        {
                            foreach (ITestAction testStep in testCase.Actions)
                            {
                                ITestStep ts = (ITestStep)testStep;
                                ts.Title = arrTCStep[i];
                                ts.ExpectedResult = arrExpRes[i];
                                i++;
                            }
                        }


                        testCase.Save();
                        tcFound = true;
                    }
                }
            }
            if (!tcFound)
            {
                Console.WriteLine("Cannot find test case in test suite");
            }
        }
        public static void ExtractToExcel(string tfsProject, int suitWkItmId,
            string fName = "DWH ETL Test Cases Extract", int[] specTCID = null, int resultSet=1)
        {
            //Projects
            //IPOS Data Warehouse
            //IPOS Data Warehouse Test

            //Test Suites
            //31796
            TfsTeamProjectCollection tfctc = new TfsTeamProjectCollection(new Uri("https://dlm.crimsonlogic.com/tfs/Government"));
            ITestManagementService testmanagementService = tfctc.GetService<ITestManagementService>();
            var teamproject = testmanagementService.GetTeamProject(tfsProject);
            var allTestCase = teamproject.TestSuites.Find(suitWkItmId).TestCases;
            //var allTestCase = teamproject.TestSuites.Find(31859).TestCases;              
            List<ITestCase> testrunInPlan = new List<ITestCase>();
            string dc = "";
            string oc = "";
            CreateExcel(fName);

            //Extract for single test case is not yet 
            int tcCtr = 1;
            foreach (var tc in allTestCase)
            {
                //foreach (int specID in specTCID)
                //{
                //    if (specID == tc.Id)
                //    {
                        ITestCase testCase = teamproject.TestCases.Find(tc.Id);
                        var testcaseResult = teamproject.TestResults.Query("Select * from TestResult where TestCaseID='" + tc.Id + "'");
                        string getSummary = testCase.Description.ToString();
                        string getTitle = testCase.Title.ToString();
                        string getPrecond = testCase.CustomFields["Pre-Condition"].Value.ToString();
                        string getRemark1 = testCase.CustomFields["Text Field 1"].Value.ToString();
                        string getTestCaseID = testCase.CustomFields["Test Case ID"].Value.ToString();
                        string getArea = testCase.Area.ToString();
                        Console.WriteLine("Extracting: " + tc.Id.ToString() + ":" + getTestCaseID);
                        int stepCtr = 1;
                        try
                        {
                            dc = testcaseResult[testcaseResult.Count - resultSet].DateCompleted.ToString();
                            oc = testcaseResult[testcaseResult.Count - resultSet].Outcome.ToString();
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            dc = "";
                            oc = "";
                        }

                        List<string> allSteps = new List<string>();
                        List<string> allExpRes = new List<string>();
                        List<int> allStepCt = new List<int>();
                        List<string> allSummary = new List<string>();
                        List<string> allTcId = new List<string>();
                        List<string> allPreCond = new List<string>();
                        List<string> allTextField1 = new List<string>();
                        List<string> allTFSID = new List<string>();
                        List<string> allDateTest = new List<string>();
                        List<string> allOutCome = new List<string>();
                        List<string> allTitle = new List<string>();

                        foreach (ITestAction testStep in testCase.Actions)
                        {
                            ITestStep ts = (ITestStep)testStep;
                            string getStep = StripHTML(ts.Title.ToString());
                            string getExpRes = StripHTML(ts.ExpectedResult.ToString());
                            if (stepCtr == 1)
                            {
                                allSummary.Add(getSummary);
                                allTcId.Add(getTestCaseID);
                                allPreCond.Add(getPrecond);
                                allTextField1.Add(getRemark1);
                                allTFSID.Add(tc.Id.ToString());
                                allDateTest.Add(dc);
                                allOutCome.Add(oc);
                                allTitle.Add(getTitle);
                            }
                            else
                            {
                                allSummary.Add("");
                                allTcId.Add("");
                                allPreCond.Add("");
                                allTextField1.Add("");
                                allTFSID.Add("");
                                allDateTest.Add("");
                                allOutCome.Add("");
                                allTitle.Add("");
                            }
                            allSteps.Add(getStep);
                            allExpRes.Add(getExpRes);
                            allStepCt.Add(stepCtr);
                            stepCtr++;
                        }
                        string[] arrSteps = allSteps.ToArray();
                        string[] arrExpRes = allExpRes.ToArray();
                        int[] arrStepCount = allStepCt.ToArray();
                        string[] arrSummary = allSummary.ToArray();
                        string[] arrTcId = allTcId.ToArray();
                        string[] arrPreCond = allPreCond.ToArray();
                        string[] arrTextField1 = allTextField1.ToArray();
                        string[] arrTFSID = allTFSID.ToArray();
                        string[] arrDateTest = allDateTest.ToArray();
                        string[] arrOutcome = allOutCome.ToArray();
                        string[] arrTitle = allTitle.ToArray();

                        OpenAndUpdate(fName, stepStart + 1, arrStepCount,
                            arrSteps, arrExpRes, arrSummary, arrTcId, arrPreCond,
                            arrTextField1, arrTFSID, arrDateTest, arrOutcome, arrTitle);
                        tcCtr++;

                        allSteps.Clear();
                        allExpRes.Clear();
                        allStepCt.Clear();
                        allSummary.Clear();
                        allTcId.Clear();
                        allPreCond.Clear();
                        allTextField1.Clear();
                //    }
                //}
            }
        }

        static void CreateExcel(String fName)
        {
            excelInterop.Application xlApp = new excelInterop.Application();
            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = false;

            excelInterop.Workbook wb = xlApp.Workbooks.Add(excelInterop.XlWBATemplate.xlWBATWorksheet);
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            ws.get_Range("A1").Value = "Work Item Type";
            ws.get_Range("B1").Value = "Assigned To";
            ws.get_Range("C1").Value = "Iteration";
            ws.get_Range("D1").Value = "Area Path";
            ws.get_Range("E1").Value = "Suite Name (Test that belongs to this suite)";
            ws.get_Range("F1").Value = "Test Case ID";
            ws.get_Range("G1").Value = "Test Scenario Name";
            ws.get_Range("H1").Value = "Test Scenario Description";
            ws.get_Range("I1").Value = "Pre-Condition";
            ws.get_Range("J1").Value = "Text Field 1";
            ws.get_Range("K1").Value = "Text Field 2";
            ws.get_Range("L1").Value = "Test Step No";
            ws.get_Range("M1").Value = "Test Step Name";
            ws.get_Range("N1").Value = "Test Step Expected Result";
            ws.get_Range("O1").Value = "TFS ID";
            ws.get_Range("P1").Value = "Date Tested";

            wb.SaveAs(defPath + fName + ".xlsx");
            wb.Close();
            xlApp.Quit();
        }

        static void UpdateExcelHeaders(String fName, int sRow, string[] getValues)
        {
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(defPath + fName + ".xlsx");
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];
            //System Name

            wb.Save();
            wb.Close();
            xlApp.Quit();
        }

        static void OpenAndUpdate(string fName, int rowNum, int[] stepNo,
            string[] testStep, string[] expRes, string[] summary, string[] tcId,
            string[] preCondition, string[] textField1, string[] tfsID, 
            string[] dateTest, string[] outCome, string[] title)
        {
            excelInterop.Application xlApp = new excelInterop.Application();

            excelInterop.Workbook wb = xlApp.Workbooks.Open(defPath + fName + ".xlsx");
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];
            int i = 0;
            for (int x = rowNum; x < testStep.Length + rowNum; x++)
            {
                if (stepNo[i] == 1)
                {
                    ws.get_Range("A" + x).Value = "Test Case";
                    ws.get_Range("B" + x).Value = "Mei Yien, Wong";
                    ws.get_Range("C" + x).Value = "\\Iteration 1";
                    ws.get_Range("D" + x).Value = "\\BI";
                    ws.get_Range("E" + x).Value = "Test Plan\\Phrase 1b\\Data Warehouse\\ETL Flow and Merging";
                }
                else
                {
                    ws.get_Range("A" + x).Value = "";
                    ws.get_Range("B" + x).Value = "";
                    ws.get_Range("C" + x).Value = "";
                    ws.get_Range("D" + x).Value = "";
                    ws.get_Range("E" + x).Value = "";
                    ws.get_Range("O" + x).Value = "";
                    ws.get_Range("P" + x).Value = "";
                    ws.get_Range("Q" + x).Value = "";
                }

                if (summary[i].Length<1)
                {
                    summary[i] = title[i];
                }
                ws.get_Range("F" + x).Value = tcId[i];
                ws.get_Range("G" + x).Value = title[i];
                ws.get_Range("H" + x).Value = summary[i];
                ws.get_Range("I" + x).Value = preCondition[i];
                ws.get_Range("J" + x).Value = textField1[i];
                ws.get_Range("K" + x).Value = tcId[0];
                ws.get_Range("L" + x).Value = stepNo[i];
                ws.get_Range("M" + x).Value = testStep[i];
                ws.get_Range("N" + x).Value = expRes[i];
                ws.get_Range("O" + x).Value = tfsID[i];
                ws.get_Range("P" + x).Value = dateTest[i];
                ws.get_Range("Q" + x).Value = outCome[i];

                i++;
                stepStart++;
            }

            //stepStart = stepStart + 1;

            wb.Save();
            wb.Close();
            xlApp.Quit();

        }

        public static string StripHTML(string input)
        {
            input = input.Replace("<DIV><P>", "");
            input = input.Replace("<DIV><P />", "");
            input = input.Replace("<P>", "\n");
            input = input.Replace("<P />", "\n");
            input = input.Replace("</P>", "");
            input = input.Replace("&lt;", "<");
            input = input.Replace("&gt;", ">");
            input = input.Replace("</DIV>", "");
            input = input.Replace("</P >", "");
            input = input.Replace("<BR/>", "\n");
            input = input.Replace("<BR />", "\n");        
            return input;
        }



 
    }
}

