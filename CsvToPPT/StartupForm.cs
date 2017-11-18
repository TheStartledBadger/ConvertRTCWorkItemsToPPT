using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace CsvToPPT
{
    static class Constants
    {
        public const int HEIGHT_PER_COST = 20;
        public static int WIDTH = 100;
    }

    public partial class StartupForm : Form
    {
        public StartupForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "CSV (*.csv)|*.csv| All files (*.*) |*.*";
            DialogResult result = openFileDialog1.ShowDialog();
            if(result == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // fetch the work items
            ExcelReader xr = new ExcelReader(openFileDialog1.FileName);
            List<WorkItemInfo> workItems = xr.getWorkItems();

            // Generate the ppt file.
            PowerPointWriter ppw = new PowerPointWriter();
            ppw.makePresentation(workItems);
        }



    }
}
