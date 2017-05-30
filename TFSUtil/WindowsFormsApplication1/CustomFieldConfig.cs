using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace TFSUtil
{
    public partial class CustomFieldConfig : Form
    {
        public CustomFieldConfig()
        {
            InitializeComponent();
            loadValues();
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
                    Label tryLabel = new Label();
                    panel_TestCaseFields.Controls.Add(tryLabel);
                    tryLabel.Location = new Point(235, 15);
                    tryLabel.Text = Internals.Globals.getTestCaseFieldsFromSetting + ".xml";
                    Label currentLabel = new Label();
                    panel_TestCaseFields.Controls.Add(currentLabel);
                    currentLabel.Location = new Point(190, 15);
                    currentLabel.Text = "Current: ";
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
    }
}
