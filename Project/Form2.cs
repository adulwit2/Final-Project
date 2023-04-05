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
using OfficeOpenXml;

namespace Project
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(@"C:\Users\Admin\Desktop\Project\Project\Project\bin\Debug\shop.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int lastUsedRow = worksheet.Dimension.End.Row;
                for (int row = 1; row <= lastUsedRow; row++)
                {
                    string value = worksheet.Cells[row, 1].Value?.ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        this.dataGridView2.Rows.Add(value);
                    }
                }

            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

    }
}
