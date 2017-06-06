using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using TFSUtil.Internals;

namespace TFSUtil
{
    public partial class CustomFieldConfig : Form
    {
        Dictionary<string,string> xmlItems = new Dictionary<string, string>();
        public CustomFieldConfig()
        {
            InitializeComponent();
            loadValues();
            loadTemplateXMLs();
        }

        private void loadValues()
        {
            //CustomFieldConfig cfc = new CustomFieldConfig();
            XDocument xdoc = XDocument.Load(@"References\CustomFields.xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;

            foreach (XElement r in xRows)
            {
                if(Convert.ToString(r.Name).Contains("TextField"))
                {
                    this.groupBox1.Controls["txt_" + r.Name].Text = r.Value;
                }           
                else
                {
                    this.groupBox2.Controls["txt_" + r.Name].Text = r.Value;
                }
            }            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //var textBoxes = groupBox1.Controls.OfType<TextBox>();
            XDocument xdoc = XDocument.Load(@"References\CustomFields.xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;
            foreach (XElement r in xRows)
            {
                if (Convert.ToString(r.Name).Contains("TextField"))
                {
                    r.SetValue(this.groupBox1.Controls["txt_" + r.Name].Text);
                }
            }
            xdoc.Save(@"References\CustomFields.xml");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //var textBoxes = groupBox1.Controls.OfType<TextBox>();
            XDocument xdoc = XDocument.Load(@"References\CustomFields.xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;
            foreach (XElement r in xRows)
            {
                if (Convert.ToString(r.Name).Contains("Remarks"))
                {                    
                    r.SetValue(this.groupBox2.Controls["txt_" + r.Name].Text);
                }
            }
            xdoc.Save(@"References\CustomFields.xml");
        }

        private void CustomFieldConfig_Load(object sender, EventArgs e)
        {
            lst_Otptions.SelectedIndex = 0;

            //combo_tcTemplatelist.SelectedIndex = 0;
        }

        private void loadTestCaseFields()
        {
            dgv_TestCaseFields.DataSource = null;
            dgv_TestCaseFields.Rows.Clear();
            dgv_TestCaseFields.Refresh();
            testmanTFS tm = new testmanTFS();
            tm.loadXMLTCFields(combo_tcTemplatelist.Text);
            //Load fields here
            //DataTable dt = new DataTable();
            //dt.Columns.Add("TestCaseFields", typeof(string));
            //for (int x = 0; x <= tm.xmlTestCaseFields.Count - 1; x++)
            //{
            //    dt.Rows.Add(new string[] { tm.xmlTestCaseFields[x] });
            //}

            DataGridViewTextBoxColumn tf = new DataGridViewTextBoxColumn();
            tf.HeaderText = "Fields";
            tf.DataPropertyName = "Fields";
            tf.ReadOnly = true;
            tf.Width = 255;
            tf.SortMode = DataGridViewColumnSortMode.NotSortable;

            dgv_TestCaseFields.Columns.Add(tf);

            DataGridViewRowCollection rows = dgv_TestCaseFields.Rows;

            for (int x = 0; x <= tm.xmlTestCaseFields.Count - 1; x++)
            {
                rows.Add(new string[] { tm.xmlTestCaseFields[x] });
            }

            //dgv_TestCaseFields.DataSource = dt;
        }


        private void lst_Otptions_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (lst_Otptions.SelectedItem.ToString())
            {
                case "Defect and Test Case":                    
                    customFieldsPanel.Dock = DockStyle.None;
                    panel_TestCaseFields.Dock = DockStyle.Fill;
                    panel_TestCaseFields.Visible = true;
                    customFieldsPanel.Visible = false;
                    //showCurrentLabel();
                    loadTemplateComboBox();
                    break;
                case "Custom Fields":
                    customFieldsPanel.Dock = DockStyle.Fill;
                    panel_TestCaseFields.Dock = DockStyle.None;
                    customFieldsPanel.Visible = true;
                    panel_TestCaseFields.Visible = false;                    
                    break;
            }
        }

        private void showCurrentLabel()
        {
            //Label currentLabel = new Label();
            //panel_TestCaseFields.Controls.Add(currentLabel);
            //currentLabel.Size = new System.Drawing.Size(150, 13);
            //currentLabel.Location = new Point(6, 35);
            //currentLabel.Text = "Current Test Case Template: ";
            //currentLabel.Name = "LabelThat";
            
            //Label currentLabelDef = new Label();            
            //currentLabelDef.Location = new Point(6, 60);
            //currentLabelDef.Size = new System.Drawing.Size(150, 13);
            //currentLabelDef.Text = "Current Defect Template: ";
            //currentLabelDef.Name = "LabelThatDef";
            //panel_TestCaseFields.Controls.Add(currentLabelDef);
            //AddCurrentLabels();
        }

        private void AddCurrentLabels()
        {
            //Label tryLabel = new Label();
            
            //tryLabel.Name = "LabelThis";
            //tryLabel.Location = new Point(155, 35);
            //tryLabel.Size = new System.Drawing.Size(300, 13);
            //tryLabel.Text = Internals.Globals.getTestCaseFieldsFromSetting + ".xml";
            //panel_TestCaseFields.Controls.Add(tryLabel);

            //Label tryLabelDef = new Label();            
            //tryLabelDef.Name = "LabelThisDef";
            //tryLabelDef.Location = new Point(155, 60);
            //tryLabelDef.Size = new System.Drawing.Size(300, 13);
            //tryLabelDef.Text = Internals.Globals.getDefectFieldsFromSetting + ".xml";
            //panel_TestCaseFields.Controls.Add(tryLabelDef);
        }

        private void loadTemplateComboBox()
        {
            combo_tcTemplatelist.DataSource = new BindingSource(xmlItems, null);
            combo_tcTemplatelist.DisplayMember = "Value";
            combo_tcTemplatelist.ValueMember = "Key";
        }

        private void loadTemplateXMLs()
        {
            xmlItems.Clear();
            
            IEnumerable<string> dirs = Directory.GetFiles(@"References\").Where(file => Regex.IsMatch(file, "[Defect|TestCase]Fields.*"));
            //            string[] dirs = Directory.GetFiles(@"References\", "DefectFields*.xml, TestCaseFields*.xml");           
            foreach (string dir in dirs)
            {
                xmlItems.Add(dir, dir.Substring(dir.LastIndexOf("\\")+1));
            }
        }

        private void txt_TextField2_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_TextField3_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_TextField4_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_TextField1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btn_applyUseTemplate_Click(object sender, EventArgs e)
        {
            string getComboValue = combo_tcTemplatelist.SelectedValue.ToString();
            string nameLabel = "";
            getComboValue = getComboValue.Substring(getComboValue.LastIndexOf("\\") + 1);
            //Internals.Globals.getTestCaseFieldsFromSetting = getComboValue.Substring(0, getComboValue.Length-4);
            if (getComboValue.Contains("TestCase"))
            {
                Internals.Globals.getTestCaseFieldsFromSetting = getComboValue.Substring(0, getComboValue.Length - 4);
                nameLabel = "TestCase";
            }
            else if (getComboValue.Contains("Defect"))
            {
                Internals.Globals.getDefectFieldsFromSetting = getComboValue.Substring(0, getComboValue.Length - 4);
                nameLabel = "Defect";
            }

            //removeLabels(nameLabel);
            //AddCurrentLabels();
        }

        private void removeLabels(string testCaseOrDefect)
        {
            string getName = "";
            if (testCaseOrDefect.Contains("TestCase"))
            {
                getName = "LabelThis";
            }
            else if (testCaseOrDefect.Contains("Defect"))
            {
                getName = "LabelThisDef";
            }
            foreach (Control control in panel_TestCaseFields.Controls)
            {
                if (control.Name==getName)
                {
                    panel_TestCaseFields.Controls.Remove(control);
                    break;
                }
            }
        }

        private void combo_tcTemplatelist_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void btn_dgvtc_Edit_Click(object sender, EventArgs e)
        {
            if(dgv_TestCaseFields.RowCount!=0)
            {
                dgv_TestCaseFields.EditMode = DataGridViewEditMode.EditOnKeystroke;
                btn_doneEdit.Visible = true;
                btn_dgvtc_up.Visible = true;
                btn_dgvtc_down.Visible = true;
                btn_dgvtc_delete.Visible = true;
                btn_dgv_addRow.Visible = true;
                btn_dgvtc_Edit.Visible = false;
                foreach (DataGridViewRow row in dgv_TestCaseFields.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (!Convert.ToString(cell.Value).Contains("Step"))
                        {
                            cell.ReadOnly = false;
                        }
                        else
                        {
                            cell.ReadOnly = true;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Please load the template file before editing");
            }
            
        }

        private void btn_LoadTCTemplate_Click(object sender, EventArgs e)
        {
            loadTestCaseFields();
        }

        private void btn_doneEdit_Click(object sender, EventArgs e)
        {
            dgv_TestCaseFields.EditMode = DataGridViewEditMode.EditProgrammatically;
            btn_doneEdit.Visible = false;
            btn_dgvtc_up.Visible = false;
            btn_dgvtc_down.Visible = false;
            btn_dgvtc_delete.Visible = false;
            btn_dgv_addRow.Visible = false;
            btn_dgvtc_Edit.Visible = true;
            disableAllEditing();
        }

        private void disableStepEditing()
        {
            foreach (DataGridViewRow row in dgv_TestCaseFields.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!Convert.ToString(cell.Value).Contains("Step"))
                    {
                        cell.ReadOnly = false;
                    }
                    else
                    {
                        cell.ReadOnly = true;
                    }
                }
            }
        }

        private void disableAllEditing()
        {
            foreach (DataGridViewRow row in dgv_TestCaseFields.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!Convert.ToString(cell.Value).Contains("Step"))
                    {
                        cell.ReadOnly = true;
                    }
                    else
                    {
                        cell.ReadOnly = true;
                    }
                }
            }
        }

        private void btn_dgvtc_up_Click(object sender, EventArgs e)
        {                            
            dgv_TestCaseFields.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            DataGridViewRowCollection rows = dgv_TestCaseFields.Rows;            
            int getIndex = dgv_TestCaseFields.SelectedCells[0].OwningRow.Index;
            if (getIndex != 0)
            {
                DataGridViewRow rowToRemove = rows[getIndex - 1];
                if (!dgv_TestCaseFields.Rows[getIndex].Cells[0].Value.ToString().Contains("Step"))
                {
                    rows.Remove(rowToRemove);
                    rows.Insert(getIndex, rowToRemove);
                }
            }
            disableStepEditing();
        }

        private void btn_dgvtc_down_Click(object sender, EventArgs e)
        {
            dgv_TestCaseFields.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            DataGridViewRowCollection rows = dgv_TestCaseFields.Rows;
            int getIndex = dgv_TestCaseFields.SelectedCells[0].OwningRow.Index;
            if (getIndex != rows.Count)
            {
                DataGridViewRow rowToRemove = rows[getIndex + 1];
                if (!dgv_TestCaseFields.Rows[getIndex].Cells[0].Value.ToString().Contains("Step"))
                {
                    if(!dgv_TestCaseFields.Rows[getIndex+1].Cells[0].Value.ToString().Contains("Step"))
                    {
                        rows.Remove(rowToRemove);
                        rows.Insert(getIndex, rowToRemove);
                    }
                }                
            }
            disableStepEditing();
        }

        private void btn_dgvtc_delete_Click(object sender, EventArgs e)
        {
            dgv_TestCaseFields.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            DataGridViewRowCollection rows = dgv_TestCaseFields.Rows;
            int getIndex = dgv_TestCaseFields.SelectedCells[0].OwningRow.Index;            
            if (!dgv_TestCaseFields.Rows[getIndex].Cells[0].Value.ToString().Contains("Step"))
            {
                rows.Remove(rows[getIndex]);
            }
            disableStepEditing();
        }

        private void btn_dgv_addRow_Click(object sender, EventArgs e)
        {
            dgv_TestCaseFields.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            DataGridViewRowCollection rows = dgv_TestCaseFields.Rows;
            int getIndex = dgv_TestCaseFields.SelectedCells[0].OwningRow.Index;
            if (getIndex != 0)
            {
                DataGridViewRow rowToRemove = rows[getIndex + 1];
                if (!dgv_TestCaseFields.Rows[getIndex].Cells[0].Value.ToString().Contains("Step"))
                {
                    //rows.Remove(rowToRemove);
                    rows.Insert(getIndex, "");                   
                }
            }
            disableStepEditing();
        }
    }
}

