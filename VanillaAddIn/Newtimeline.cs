using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Application = Microsoft.Office.Interop.OneNote.Application;  // Conflicts with System.Windows.Forms


namespace MyApplication.VanillaAddIn
{
    public partial class chosetime : Form
    {
        private Application onenoteApplication;
        private XmlDocument xml;
        private OperateOnenote operate;
        private int cycletype;//1 for cycle task, 0 for one time task
        public string newtimeline;

        public chosetime(Application application)
        {
            InitializeComponent();
            custmoninitial();
            onenoteApplication = application;
            operate = new OperateOnenote(onenoteApplication);
            this.Load += Form1_Load;
        }

       

        private void Form1_Load(object sender, EventArgs e)
        {
            operate = new OperateOnenote(this.onenoteApplication);
           
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = "$MONTH " + dataGridView2.SelectedCells[0].Value.ToString() + "$";
            cycletype = 1;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = "$WEEK " + dataGridView1.SelectedCells[0].Value.ToString() + "$";
            cycletype = 1;
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = "$YEAR " + dataGridView3.SelectedCells[0].Value.ToString() + "$";
            cycletype = 1;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            textBox1.Text = "$" + dateTimePicker1.Value.ToLongDateString()+ "$";
            cycletype = 0;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            newtimeline = textBox1.Text;
            Hide();
            if (cycletype == 0)
            {
                Clipboard.SetText(newtimeline);
                SendKeys.SendWait("^v");                
            }
            else if (cycletype == 1)
            {
                
                if (textBox3.Text == "" || textBox2.Text == "")
                {
                    MessageBox.Show("未填写事项或次数！");
                    return;
                }
                int times = Convert.ToInt32(textBox2.Text);
                string pageid=operate.getpageid(textBox3.Text);
                XmlDocument pagexml;
                if (pageid == null)
                {
                    onenoteApplication.CreateNewPage(onenoteApplication.Windows.CurrentWindow.CurrentSectionId, out pageid);
                    pagexml = operate.GetPageContent(pageid);
                    operate.setpagename(pagexml, textBox3.Text);
                }
                else
                {
                    pagexml = operate.GetPageContent(pageid);
                }
                operate.AddPageline(pagexml, textBox3.Text+textBox1.Text, OperateOnenote.linetype.tagline);
                while (times > 1)
                {
                    operate.AddPageline(pagexml, textBox3.Text,OperateOnenote.linetype.tagline);
                    times--;
                }
                onenoteApplication.UpdatePageContent(pagexml.InnerXml, DateTime.MinValue);
            }
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "$DAY " + "0" + "$";
            cycletype = 1;
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }
    }
}
