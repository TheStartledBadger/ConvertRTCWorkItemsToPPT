using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;

using System.Runtime.InteropServices;

namespace CsvToPPT
{
    static class Constants
    {
        public const int HEIGHT_PER_COST = 20;
        public static int WIDTH = 100;
    }

    public partial class Form1 : Form
    {
        public Form1()
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
            List<WorkItemInfo> workItems = fetchWorkItems(openFileDialog1.FileName);

            // Generate the ppt file.
            buildPPT(workItems);
        }

        private void buildPPT(List<WorkItemInfo> workItems)
        {
            PowerPoint.Application objApp;
            PowerPoint._Presentation objPres;
            PowerPoint.Slides objSlides;
            PowerPoint._Slide objSlide;
            PowerPoint.TextRange objTextRng;

            //Create a new presentation based on a template.
            objApp = new PowerPoint.Application();
            objPres = objApp.Presentations.Add(MsoTriState.msoTrue);
            objApp.Visible = MsoTriState.msoTrue;
            objSlides = objPres.Slides;

            //Build Slide #1:
            objSlide = objSlides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);

            foreach(var info in workItems)
            {
                addObject(objSlide, info);
            }

            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Font.Size = 8;
        }

        private void addObject(PowerPoint._Slide slide, WorkItemInfo info)
        {
            PowerPoint.Shape shp = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 50, 50, Constants.WIDTH, info.Cost * Constants.HEIGHT_PER_COST);
            shp.TextFrame.TextRange.Text = info.Id + "\n" + info.Summary;
            shp.TextFrame.TextRange.Font.Size = 6;
        }

        private List<WorkItemInfo> fetchWorkItems(string filename)
        {
            List<WorkItemInfo> results = new List<WorkItemInfo>();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            for(int i=1;i<= rowCount; i++)
            {
                string id = getCellValue(xlRange, i,1);
                string summary = getCellValue(xlRange, i, 2);
                string owner = getCellValue(xlRange, i, 3);
                string costString = getCellValue(xlRange, i, 4);
                int cost = Convert.ToInt32(costString);

                results.Add(new WorkItemInfo(id, summary, owner, cost));
            }

            xlWorkbook.Close();
            xlApp.Quit();
            return results;
        }

        private string getCellValue(Excel.Range xlRange, int row, int col)
        {
            if (xlRange.Cells[row, col] != null && xlRange.Cells[row, col].Value2 != null)
            {
                return xlRange.Cells[row, col].Value2.ToString();
            } else
            {
                return "<undefined>";
            }
        }
    }
}
