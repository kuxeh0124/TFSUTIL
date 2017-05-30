using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFSUtil.Internals;
using System.IO;
using excelInterop = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace TFSUtil
{
    public partial class TFSUtilUI : Form
    {
        bool isConnected = false;
        bool isExtactAll = false;
        Dictionary<string, string> dicTCToExtract = new Dictionary<string, string>();

        public TFSUtilUI()
        {
            InitializeComponent();
            Globals.loadSettings();
        }

        private void newConnectionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isConnected)
            {
                MessageBox.Show("Already Connected to " + connectTFS.tfsTeamProject.ToString() + "\n" 
                    + "Please switch projects if you need to connect to another project", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                tfsConnectionHandler();
            }                                  
        }

        private void cb_testingPhase_CheckedChanged(object sender, EventArgs e)
        {
            if (combo_TestingPhase.Enabled)
            {
                combo_TestingPhase.Enabled = false;
            }
            else
            {
                combo_TestingPhase.Enabled = true;
            }
        }

        private void cb_User_CheckedChanged(object sender, EventArgs e)
        {
            if (combo_User.Enabled)
            {
                combo_User.Enabled = false;
            }
            else
            {
                combo_User.Enabled = true;
            }
        }

        private void cb_State_CheckedChanged(object sender, EventArgs e)
        {
            if (combo_State.Enabled)
            {
                combo_State.Enabled = false;
            }
            else
            {
                combo_State.Enabled = true;
            }
        }

        private void cb_Query_CheckedChanged(object sender, EventArgs e)
        {
            if(!cb_Query.Checked)
            {
                combo_MyQuery.Enabled = false;
                cb_User.Enabled = true;
                cb_State.Enabled = true;
                cb_testingPhase.Enabled = true;
                combo_User.Enabled = false;
                combo_State.Enabled = false;
                combo_TestingPhase.Enabled = false;
            }
            else
            {
                combo_MyQuery.Enabled = true;
                
                cb_User.Enabled = false;
                cb_User.Checked = false;               
                cb_State.Enabled = false;
                cb_State.Checked = false;                
                cb_testingPhase.Enabled = false;
                cb_testingPhase.Checked = false;
                combo_User.Enabled = false;
                combo_TestingPhase.Enabled = false;
                combo_State.Enabled = false;
            }
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_extractAll.Checked)
            {
                combo_MyQuery.Enabled = false;
                cb_Query.Checked = false;
                cb_Query.Enabled = false;
                combo_User.Enabled = false;
                cb_User.Enabled = false;
                cb_User.Checked = false;
                combo_State.Enabled = false;
                cb_State.Enabled = false;
                cb_State.Checked = false;
                combo_TestingPhase.Enabled = false;
                cb_testingPhase.Enabled = false;
                cb_testingPhase.Checked = false;
                isExtactAll = true;
            }
            else
            {                
                cb_Query.Enabled = true;                
                cb_User.Enabled = true;                                
                cb_State.Enabled = true;                                
                cb_testingPhase.Enabled = true;                
            }
        }
        
        private void tfsConnectionHandler()
        {
            connectTFS.connectToTFS();
            if (connectTFS.tfsTeamProject.ToString().Length > 0)
            {
                statusLbl_Connection.Text = "Connected";
                statusLbl_ConnectionTM.Text = "Connected";
                isConnected = true;
                //Load Dropdowns
                try
                {
                    defectsTFS deftfs = new defectsTFS();                    
                    deftfs.getTFSDefectFields();
                    combo_TestingPhase.DataSource = deftfs.fieldAllowedValues["Testing Phase"].ToList();
                    combo_User.DataSource = deftfs.fieldAllowedValues["Assigned To"].ToList();
                    combo_State.DataSource = deftfs.fieldAllowedValues["State"].ToList();
                    combo_MyQuery.DataSource = deftfs.listQueries;                    
                    MessageBox.Show("Successfully Connected to Project: " + connectTFS.tfsTeamProject.ToString());
                }
                catch (Exception err)
                {
                    Console.WriteLine(err.ToString());
                    MessageBox.Show("There was a problem connecting to Project: " + connectTFS.tfsTeamProject.ToString() + "\r\n" +
                        "Please contact administrator", "Connection Unsuccessful", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void switchProjectsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isConnected)
            {
                tfsConnectionHandler();

            }
            else
            {
                MessageBox.Show("Please connect to TFS before switching", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_BrowseForExtract_Click(object sender, EventArgs e)
        {
            int size = -1;                        
             // Show the dialog.
            FolderBrowserDialog fldg = folderBrowserDialog1;
            //fldg.Filter = "Excel files (*.xls)|*.xls|Excel files (*.xlsx)|*.xlsx";
            DialogResult result = fldg.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                //string file = folderBrowserDialog1.FileName;
                try
                {
                    //string text = File.ReadAllText(file);
                    //size = text.Length;
                    
                    Globals.getExtractDrive = fldg.SelectedPath.ToString();
                    txt_extractDirectoryPath.Text = Globals.getExtractDrive;
                }
                catch (IOException)
                {
                }
            }
            //Console.WriteLine(size); // <-- Shows file size in debugging mode.
            //Console.WriteLine(result); // <-- For debugging use.
        }

        private void btn_ExtractDefects_Click(object sender, EventArgs e)
        {
            ProcessExtract();            
        }
        
        private void myWaitForm()
        {
            Form waitForm = new Form();
            waitForm.ControlBox = false;
            waitForm.Size = new Size(200, 80);

            Label waitLabel = new Label();
            waitLabel.Text = "TESTING";
            waitForm.Controls.Add(waitLabel);
            StartPosition = FormStartPosition.CenterScreen;
            waitForm.ShowDialog();
        }

        void ProcessExtract()
        {
            Thread getThread = new Thread(myWaitForm);
            getThread.Start();
            bool getFalse = false;
            int i = 0;
            try
            {
                defectsTFS defTfs = new defectsTFS();
                List<string> getComboData = new List<string>();
                List<string> getComboValue = new List<string>();
                string wiql = "";
                if (txt_extractDirectoryPath.Text.Length != 0)
                {
                    if (!combo_MyQuery.Enabled)
                    {
                        if (combo_TestingPhase.Enabled)
                        {
                            getComboData.Add(combo_TestingPhase.Name.ToString());
                            getComboValue.Add(combo_TestingPhase.SelectedValue.ToString());
                        }
                        if (combo_User.Enabled)
                        {
                            getComboData.Add(combo_User.Name.ToString());
                            getComboValue.Add(combo_User.SelectedValue.ToString());
                        }
                        if (combo_State.Enabled)
                        {
                            getComboData.Add(combo_State.Name.ToString());
                            getComboValue.Add(combo_State.SelectedValue.ToString());
                        }
                    }
                    else
                    {
                        wiql = defectsTFS.queryValue[combo_MyQuery.SelectedValue.ToString()];
                    }
                    getFalse = defTfs.extractInformationFromDefect(wiql, txt_extractDirectoryPath.Text.ToString(),
                        getComboData.ToArray(), getComboValue.ToArray());
                    if (getFalse)
                    {
                        getThread.Abort();
                        MessageBox.Show("Extract completed!");
                    }
                    else
                    {
                        getThread.Abort();
                        MessageBox.Show("There was a problem with the extraction.");
                    }
                }
                else
                {
                    getThread.Abort();
                    MessageBox.Show("Please specify and extract file location", "Missing File Location", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception err)
            {
                getThread.Abort();
                MessageBox.Show("Extract failed due to:");
            }
        }
        void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // The progress percentage is a property of e
            extractProgess.Value = e.ProgressPercentage;
        }
        private void btn_BroweUpload_Click(object sender, EventArgs e)
        {
            int size = -1;
            // Show the dialog.
            OpenFileDialog fldg = openFileDialog1;
            fldg.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"; ;
            DialogResult result = fldg.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {                
                try
                {
                    string file = openFileDialog1.FileName;
                    txt_FileForUpDown.Text = file;
                    Globals.getFileName = file;
                }
                catch (IOException)
                {
                }
            }
        }

        private void btn_Open_Click(object sender, EventArgs e)
        {
            if (txt_FileForUpDown.Text.Length > 0)
            {
                new ExcelProcessing().openShowWorkbook(Globals.getFileName);
            }
            else
            {
                MessageBox.Show("Please select a file to open", "No File to Open", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void btn_Refresh_Click(object sender, EventArgs e)
        {

            if (txt_FileForUpDown.Text.Length > 0)
            {
                DialogResult sureToRefresh = MessageBox.Show("Are you sure you want to refresh the data." +
                    "\r\n\r\nPlease note that all your changes will be removed once you agree. \r\n\r\nContinue?",
                    "Continue Refresh?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (sureToRefresh == DialogResult.Yes)
                {
                    ExcelProcessing xlProc = new ExcelProcessing();
                    defectsTFS defTfs = new defectsTFS();
                    defTfs.loadXMLDefectFieldsForValidation();
                    if (xlProc.validateFileFormat(defTfs.xmlDefectFieldsVal, txt_FileForUpDown.Text))
                    {
                        tfsConnectionHandler();
                        defTfs.extractInformationFromDefect(txt_FileForUpDown.Text);
                        xlProc.openShowWorkbook();
                        MessageBox.Show("Successfully Refreshed the excel file.", "Refresh Success", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a file!");
            }
        }

        private void btn_Upload_Click(object sender, EventArgs e)
        {
            if (isConnected)
            {
                processUpload();
            }
            else
            {
                MessageBox.Show("Please connect to a project first");
            }
        }
        private void processUpload()
        {
            ExcelProcessing xlProc = new ExcelProcessing();
            defectsTFS defTfs = new defectsTFS();
            defTfs.loadXMLDefectFieldsForValidation();
            if (txt_FileForUpDown.Text.Length > 0)
            {
                if (xlProc.validateFileFormat(defTfs.xmlDefectFieldsVal, txt_FileForUpDown.Text))
                {
                    xlProc.createNewReport();
                    defTfs.loadIntoTFS(txt_FileForUpDown.Text);
                    MessageBox.Show("Upload completed! \r\n\r\nUploaded " + defectsTFS.getSuccess + " / " + defectsTFS.getTotalUpload + " Successfully",
                        "Upload Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DialogResult getReportResult = MessageBox.Show("Do you want to view the report?", "View Report", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (getReportResult == DialogResult.Yes)
                    {
                        xlProc.openShowWorkbook(Globals.getReportPath);                                                                      
                    }
                    else 
                    {
                        xlProc.openShowWorkbook();
                    }
                }
                else
                {
                    MessageBox.Show("File validation failed. Please check your template before uploading");
                }
            }
            else
            {
                MessageBox.Show("Please select a file!");
            }
        }

        private void btn_BrowseGenerate_Click(object sender, EventArgs e)
        {
            int size = -1;
            // Show the dialog.
            FolderBrowserDialog fldg = folderBrowserDialog1;
            //fldg.Filter = "Excel files (*.xls)|*.xls|Excel files (*.xlsx)|*.xlsx";
            DialogResult result = fldg.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                //string file = folderBrowserDialog1.FileName;
                try
                {
                    //string text = File.ReadAllText(file);
                    //size = text.Length;

                    Globals.getTemplateDrive = fldg.SelectedPath.ToString();
                    txt_GenTemplate.Text = Globals.getTemplateDrive;
                }
                catch (IOException)
                {
                }
            }
            //Console.WriteLine(size); // <-- Shows file size in debugging mode.
            //Console.WriteLine(result); // <-- For debugging use.
        }

        private void btn_Generate_Click(object sender, EventArgs e)
        {
            if (txt_GenTemplate.Text.Length > 0)
            {
                ExcelProcessing xlProc = new ExcelProcessing();
                xlProc.createNewExcelTemplate("Defect");
                xlProc.generateDefectTemplate(txt_GenTemplate.Text);
                xlProc.openShowWorkbook();
                MessageBox.Show("Template successfully created in: " + txt_GenTemplate.Text, "Generation Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Please select a folder to save in", "No Folder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Test Mangement Objects
        private void btn_loadSuites_Click(object sender, EventArgs e)
        {
            if (isConnected)
            {
                testmanTFS tmTfs = new testmanTFS();
                tmTfs.GetTestSuites(txt_SuiteNumber.Text);
                combo_TestSuites.Enabled = true;
                combo_TestSuites.DataSource = new BindingSource(tmTfs.getSuiteList, null);
                combo_TestSuites.DisplayMember = "Value";
                combo_TestSuites.ValueMember = "Key";
            }
            else
            {
                MessageBox.Show("Please connect to TFS before loading the test suites.");
            }
        }

        private void combo_TestSuites_SelectedIndexChanged(object sender, EventArgs e)
        {
            testmanTFS tmTfs = new testmanTFS();
            dicTCToExtract.Clear();
            try
            {
                if(dicTCToExtract.Count==0)
                {
                    list_TCToExtract.DataSource = new BindingSource();
                }
                //list_TCToExtract.DataSource = new BindingSource(dicTCToExtract, null);
                //list_TCToExtract.DisplayMember = "Value";
                //list_TCToExtract.ValueMember = "Key";
            }
            catch { }

            string getValue = "";
            try
            {
                getValue = ((KeyValuePair<string, string>)combo_TestSuites.SelectedValue).Key.ToString();                
            }
            catch(InvalidCastException)
            {
                getValue = Convert.ToString(combo_TestSuites.SelectedValue);                
            }
            Globals.getTestPlan = Convert.ToString(combo_TestSuites.SelectedItem).Replace(getValue + ",", "").Replace("[", "").Replace("]", "").Trim();
            tmTfs.GetTestCases(getValue);
            list_LoadedTC.DataSource = new BindingSource(tmTfs.getTCListFromSuite, null);
            list_LoadedTC.DisplayMember = "Value";
            list_LoadedTC.ValueMember = "Key";
        }

        private void btn_addTC_Click(object sender, EventArgs e)
        {
            addToExtractList();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            testmanTFS tmTfs = new testmanTFS();
            //For testing purposes
            //ExcelProcessing xlProc = new ExcelProcessing();
            if (combo_uldl.SelectedItem.ToString() == "Download")
            {            
                if (list_TCToExtract.Items.Count != 0)
                {
                    if(txt_UploadDownloadTC.Text.Length>0)
                    {
                        
                        foreach (var item in list_TCToExtract.Items)
                        {
                            string getID = ((KeyValuePair<string, string>)item).Key;
                            tmTfs.GetTestCaseInformation(getID);
                        }
                        if (!cb_SpecFormat.Checked)
                        {
                            if (tmTfs.CreateTestCaseExtractFile(txt_UploadDownloadTC.Text))
                            {
                                MessageBox.Show("Test Case Extraction Complete!", "Test Case Extraction", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Test Case Extraction hit an exception!", "Test Case Extraction", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            if (tmTfs.CreateTestCaseExtractFile(txt_UploadDownloadTC.Text, cb_SpecFormat.Checked))
                            {
                                MessageBox.Show("Test Case Extraction Complete!", "Test Case Extraction", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Test Case Extraction hit an exception!", "Test Case Extraction", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }                    
                    }
                    else
                    {
                        MessageBox.Show("Please select a folder to extract to", "Location Undefined", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Please select test cases to extract", "No test cases to extract", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                if(txt_UploadDownloadTC.Text.Length>0)
                {
                    tmTfs.LoadIntoTFS(txt_UploadDownloadTC.Text);
                }
                else
                {
                    MessageBox.Show("Please select a file to upload", "No test cases to upload", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                
                //xlProc.readExcelDataForTC(txt_UploadDownloadTC.Text, "Testcase");
            }
        }

        private void btn_addTCALL_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < list_LoadedTC.Items.Count; i++)
            {
                list_LoadedTC.SetSelected(i, true);
            }
            addToExtractList();
        }

        private void addToExtractList()
        {
            foreach (var item in list_LoadedTC.SelectedItems)
            {
                if (!dicTCToExtract.ContainsKey(((KeyValuePair<string, string>)item).Key.ToString()))
                {
                    dicTCToExtract.Add(((KeyValuePair<string, string>)item).Key.ToString(), ((KeyValuePair<string, string>)item).Value.ToString());
                }
            }
            list_TCToExtract.DataSource = new BindingSource(dicTCToExtract, null);
            list_TCToExtract.DisplayMember = "Value";
            list_TCToExtract.ValueMember = "Key";
        }
        private void btn_removeTC_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (var item in list_TCToExtract.SelectedItems)
                {
                    if (dicTCToExtract.ContainsKey(((KeyValuePair<string, string>)item).Key.ToString()))
                    {
                        dicTCToExtract.Remove(((KeyValuePair<string, string>)item).Key.ToString());
                    }
                }
                if (dicTCToExtract.Count == 0)
                {
                    list_TCToExtract.DataSource = new BindingSource();
                }
                else
                {
                    list_TCToExtract.DataSource = new BindingSource(dicTCToExtract, null);
                    list_TCToExtract.DisplayMember = "Value";
                    list_TCToExtract.ValueMember = "Key";
                }

            }
            catch { }
        }

        private void btn_TestCaseFileBrowse_Click(object sender, EventArgs e)
        {            
            int size = -1;
            // Show the dialog.
                try
                {
                    if (combo_uldl.SelectedItem.ToString() == "Upload")
                    {
                        OpenFileDialog fldg = openFileDialog1;
                        fldg.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"; ;
                        DialogResult result = fldg.ShowDialog();
                        if (result == DialogResult.OK) // Test result.
                        {
                            try
                            {
                                string file = openFileDialog1.FileName;
                                txt_UploadDownloadTC.Text = file;
                                Globals.testCaseFileName = file;
                            }
                            catch (IOException)
                            {
                            }
                        }
                    }
                    else
                    {
                        FolderBrowserDialog fldg = folderBrowserDialog1;
                        //fldg.Filter = "Excel files (*.xls)|*.xls|Excel files (*.xlsx)|*.xlsx";
                        DialogResult result = fldg.ShowDialog();
                        if (result == DialogResult.OK) // Test result.
                        {
                            //string file = folderBrowserDialog1.FileName;
                            try
                            {
                                //string text = File.ReadAllText(file);
                                //size = text.Length;

                                Globals.getTCExtractDrive = fldg.SelectedPath.ToString();
                                txt_UploadDownloadTC.Text = Globals.getTCExtractDrive;
                            }
                            catch (IOException)
                            {
                            }
                        }
                        //Console.WriteLine(size); // <-- Shows file size in debugging mode.
                        //Console.WriteLine(result); // <-- For debugging use.
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Please select a processing type", "No Processing Type", MessageBoxButtons.OK, MessageBoxIcon.Error);
                } 
        }

        private void btn_GenerateTCTemplate_Click(object sender, EventArgs e)
        {
            if (txt_tcTemplateLoc.Text.Length > 0)
            {
                ExcelProcessing xlProc = new ExcelProcessing();
                xlProc.createNewExcelTemplate("TestCase");
                xlProc.generateTestCaseTemplate(txt_tcTemplateLoc.Text);
                xlProc.openShowWorkbook();
                MessageBox.Show("Template successfully created in: " + txt_tcTemplateLoc.Text, "Generation Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Please select a folder to save in", "No Folder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_BrowseTemplateLocation_Click(object sender, EventArgs e)
        {
            int size = -1;
            // Show the dialog.
            FolderBrowserDialog fldg = folderBrowserDialog1;
            //fldg.Filter = "Excel files (*.xls)|*.xls|Excel files (*.xlsx)|*.xlsx";
            DialogResult result = fldg.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                //string file = folderBrowserDialog1.FileName;
                try
                {
                    //string text = File.ReadAllText(file);
                    //size = text.Length;

                    Globals.getTCTemplateDrive = fldg.SelectedPath.ToString();
                    txt_tcTemplateLoc.Text = Globals.getTCTemplateDrive;
                }
                catch (IOException)
                {
                }
            }
            //Console.WriteLine(size); // <-- Shows file size in debugging mode.
            //Console.WriteLine(result); // <-- For debugging use.
        }

        private void combo_uldl_SelectedIndexChanged(object sender, EventArgs e)
        {
            //txt_UploadDownloadTC.Text = "";
        }

        private void customFieldConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CustomFieldConfig custField = new CustomFieldConfig();
            custField.ShowDialog();
        }

        private void loadRequirements_Click(object sender, EventArgs e)
        {
            rtmTFS rtm = new rtmTFS();
            ExcelProcessing xlProc = new ExcelProcessing();
            int totalRow = MappingTableDataGrid.Rows.Count - 1;
            for (int x = 0; x <= totalRow; x++)
            {
                rtm.getRTMMappedFields[Convert.ToString(MappingTableDataGrid.Rows[x].Cells[0].Value)]
                    = Convert.ToString(MappingTableDataGrid.Rows[x].Cells[1].Value);
            }
            string rtmDestFileName = txt_rtmTemplate.Text;
            rtmDestFileName = rtmDestFileName.Substring(rtmDestFileName.LastIndexOf("\\"));
            Globals.getRtmDestinationFile = txt_rtmDest.Text + rtmDestFileName;
            File.Copy(txt_rtmTemplate.Text, Globals.getRtmDestinationFile, true);
            rtm.loadRequirements();
            xlProc.CreateRTMFromTemplate(Convert.ToInt32(txt_StartHeaderRow.Text), Convert.ToInt32(txt_StartHeaderCol.Text), rtm);
        }

        private void browseRtmTemplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog fldg = openFileDialog1;
            fldg.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"; ;
            DialogResult result = fldg.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                try
                {
                    string file = openFileDialog1.FileName;
                    txt_rtmTemplate.Text = file;
                    Globals.testCaseFileName = file;
                }
                catch (IOException)
                {
                }
            }
        }

        private void loadTemplate_Click(object sender, EventArgs e)
        {
            if (txt_rtmDest.Text.Length < 0 && txt_rtmTemplate.Text.Length < 0)
            {
                MessageBox.Show("Please set a source file and a destination location.", "Source and/Or Destination Missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if(txt_StartHeaderCol.Text.Length<0 && txt_StartHeaderRow.Text.Length<0)
            {
                MessageBox.Show("Please set a header row and column values", "Row and header Values Missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                ExcelProcessing xlProc = new ExcelProcessing();
                rtmTFS rtm = new rtmTFS();

                xlProc.loadRTMExcelFileHeaders(Convert.ToInt32(txt_StartHeaderRow.Text),
                    Convert.ToInt32(txt_StartHeaderCol.Text), txt_rtmTemplate.Text);
                rtm.loadRTMFields();
                DataTable dt = new DataTable();
                dt.Columns.Add("TemplateField", typeof(string));
                dt.Columns.Add("MappingField", typeof(string));
                for (int x = 0; x <= xlProc.rtmHeaders.Count - 1; x++)
                {
                    dt.Rows.Add(new string[] { xlProc.rtmHeaders[x], "Ignore" });
                }

                DataGridViewComboBoxColumn mf = new DataGridViewComboBoxColumn();
                var list11 = rtm.getRTMFields;
                mf.DataSource = list11;
                mf.HeaderText = "MappingField";
                mf.DataPropertyName = "MappingField";
                mf.Width = 230;
                mf.FlatStyle = FlatStyle.Flat;

                DataGridViewTextBoxColumn tf = new DataGridViewTextBoxColumn();
                tf.HeaderText = "TemplateField";
                tf.DataPropertyName = "TemplateField";
                tf.ReadOnly = true;
                tf.Width = 230;

                MappingTableDataGrid.Columns.Add(tf);
                MappingTableDataGrid.Columns.Add(mf);
                MappingTableDataGrid.DataSource = dt;
            }
        }

        private void btn_browseDestRtm_Click(object sender, EventArgs e)
        {
            int size = -1;
            // Show the dialog.
            FolderBrowserDialog fldg = folderBrowserDialog1;
            //fldg.Filter = "Excel files (*.xls)|*.xls|Excel files (*.xlsx)|*.xlsx";
            DialogResult result = fldg.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                //string file = folderBrowserDialog1.FileName;
                try
                {
                    //string text = File.ReadAllText(file);
                    //size = text.Length;

                    Globals.getRTMDrive = fldg.SelectedPath.ToString();
                    txt_rtmDest.Text = Globals.getRTMDrive;
                }
                catch (IOException)
                {
                }
            }
            //Console.WriteLine(size); // <-- Shows file size in debugging mode.
            //Console.WriteLine(result); // <-- For debugging use.
        }
    }
}
