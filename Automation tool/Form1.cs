using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Automation_tool
{
    public partial class Form1 : Form
    {
        string filePath = string.Empty;
        string fileName = string.Empty;
        public Form1()
        {
            InitializeComponent();
            button2.Enabled = false;
            textBox1.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            
            
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;
                fileName = Path.GetFileName(filePath);
                textBox1.Text = filePath;                
                button2.Enabled = true;
            }            
        }
        bool excelCheck(string excelPath, string excelFileName)
        {
            return excelPath != null && (excelPath.EndsWith(".xls", StringComparison.Ordinal) || excelPath.EndsWith(".xlsx", StringComparison.Ordinal));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (excelCheck(filePath, fileName))
            {
                ReadExcel excel = new ReadExcel(filePath, fileName);
                excel.openExcel();
                excel.closeWorkbook();
                //MessageBox.Show(value);
            }
            else
            {
                textBox1.Clear();
                MessageBox.Show("Please select a valid Excel file", "Error", MessageBoxButtons.OK);
            }
        }
    }
}
