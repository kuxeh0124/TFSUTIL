using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
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
            DataTable dt = new DataTable();
            dt.Columns.Add("TestCaseFields", typeof(string));
            for (int x = 0; x <= tm.xmlTestCaseFields.Count - 1; x++)
            {
                dt.Rows.Add(new string[] { tm.xmlTestCaseFields[x] });
            }

            DataGridViewTextBoxColumn tf = new DataGridViewTextBoxColumn();
            tf.HeaderText = "TestCaseFields";
            tf.DataPropertyName = "TestCaseFields";
            tf.ReadOnly = true;
            tf.Width = 310;
            tf.SortMode = DataGridViewColumnSortMode.NotSortable;

            dgv_TestCaseFields.Columns.Add(tf);
            dgv_TestCaseFields.DataSource = dt;
        }
       

        private void lst_Otptions_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (lst_Otptions.SelectedItem.ToString())
            {
                case "Test Case Fields":                    
                    customFieldsPanel.Dock = DockStyle.None;
                    panel_TestCaseFields.Dock = DockStyle.Fill;
                    panel_TestCaseFields.Visible = true;
                    customFieldsPanel.Visible = false;
                    showCurrentLabel();
                    loadTemplateComboBox();
                    break;
                case "Defect Fields":
                    panel_TestCaseFields.Dock = DockStyle.None;
                    customFieldsPanel.Dock = DockStyle.None;
                    panel_TestCaseFields.Visible = false;
                    customFieldsPanel.Visible = false;                    
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
            Label tryLabel = new Label();
            panel_TestCaseFields.Controls.Add(tryLabel);
            tryLabel.Name = "LabelThis";
            tryLabel.Location = new Point(105, 35);
            tryLabel.Size = new System.Drawing.Size(300, 13);
            tryLabel.Text = Internals.Globals.getTestCaseFieldsFromSetting + ".xml";
            Label currentLabel = new Label();
            panel_TestCaseFields.Controls.Add(currentLabel);
            currentLabel.Location = new Point(6, 35);
            currentLabel.Text = "Current Template: ";
            currentLabel.Name = "LabelThat";
        }

        private void loadTemplateComboBox()
        {
            combo_tcTemplatelist.DataSource = new BindingSource(xmlItems, null);
            combo_tcTemplatelist.DisplayMember = "Value";
            combo_tcTemplatelist.ValueMember = "Value";
        }

        private void loadTemplateXMLs()
        {
            xmlItems.Clear();
            string[] dirs = Directory.GetFiles(@"References\", "TestCase*.xml");           
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
            getComboValue = getComboValue.Substring(getComboValue.LastIndexOf("\\") + 1);
            Internals.Globals.getTestCaseFieldsFromSetting = getComboValue.Substring(0, getComboValue.Length-4);
            removeLabels();
            showCurrentLabel();
        }

        private void removeLabels()
        {
            foreach (Control control in panel_TestCaseFields.Controls)
            {
                if (control.Name == "LabelThis")
                {
                    panel_TestCaseFields.Controls.Remove(control);
                    break;
                }
            }
        }

        private void combo_tcTemplatelist_SelectedIndexChanged(object sender, EventArgs e)
        {
            loadTestCaseFields();
        }

        private void btn_dgvtc_Edit_Click(object sender, EventArgs e)
        {
            dgv_TestCaseFields.EditMode = DataGridViewEditMode.EditOnKeystroke;
            btn_doneEdit.Visible = true;
            btn_dgvtc_up.Visible = true;
            btn_dgvtc_down.Visible = true;
            btn_dgvtc_delete.Visible = true;
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
    }
}

