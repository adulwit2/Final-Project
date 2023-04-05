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
    public partial class Form1 : Form
    {
        public string price;
        public string name;
        public string quantity;
        public string size;
        public string length;
        public string color;
        public Form1()
        {
            InitializeComponent();
            comboBox1.Items.Add("255");
            comboBox1.Items.Add("260");
            comboBox1.Items.Add("265");
            comboBox1.Items.Add("270");
            comboBox1.Items.Add("275");
            comboBox1.Items.Add("280");
            comboBox2.Items.Add("255");
            comboBox2.Items.Add("260");
            comboBox2.Items.Add("265");
            comboBox2.Items.Add("270");
            comboBox2.Items.Add("275");
            comboBox2.Items.Add("280");
            comboBox3.Items.Add("255");
            comboBox3.Items.Add("260");
            comboBox3.Items.Add("265");
            comboBox3.Items.Add("270");
            comboBox3.Items.Add("275");
            comboBox3.Items.Add("280");
            comboBox4.Items.Add("long socks");
            comboBox4.Items.Add("Racing half socks");
            comboBox5.Items.Add("Black");
            comboBox5.Items.Add("White");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            name = this.label2.Text;
            price = this.label7.Text;
            size = this.comboBox1.SelectedItem.ToString();
            quantity = this.textBox1.Text;
            Shoe shoe = new Shoe(name, price, size, quantity);

            var message = $"สั่งซื้อ {shoe.getName()} \nราคา {shoe.getPrice()} \nไซส์ {shoe.getSize()} \nจำนวน {shoe.getQuantity()}";
            
            using (var package = new ExcelPackage(new FileInfo(@"C:\Users\Admin\Desktop\Project\Project\Project\bin\Debug\shop.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                if (worksheet == null)
                {
                    throw new Exception("Worksheet does not exist");
                }
                int lastUsedRow = worksheet.Dimension.End.Row;

                ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                cell.Value = message;

                package.Save();
            }
            this.DialogResult = DialogResult.OK;
            this.Hide();
            MessageBox.Show(message);
            Form2 form2= new Form2();
            form2.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            name = this.label3.Text;
            price = this.label8.Text;
            size = this.comboBox2.SelectedItem.ToString();
            quantity = this.textBox2.Text;
            Shoe shoe = new Shoe(name, price, size, quantity);

            var message = $"สั่งซื้อ {shoe.getName()} \nราคา {shoe.getPrice()} \nไซส์ {shoe.getSize()} \nจำนวน {shoe.getQuantity()}";

            using (var package = new ExcelPackage(new FileInfo(@"C:\Users\Admin\Desktop\Project\Project\Project\bin\Debug\shop.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                if (worksheet == null)
                {
                    throw new Exception("Worksheet does not exist");
                }
                int lastUsedRow = worksheet.Dimension.End.Row;

                ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                cell.Value = message;

                package.Save();
            }
            this.DialogResult = DialogResult.OK;
            this.Hide();
            MessageBox.Show(message);
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            name = this.label4.Text;
            price = this.label9.Text;
            size = this.comboBox3.SelectedItem.ToString();
            quantity = this.textBox3.Text;
            Shoe shoe = new Shoe(name, price, size, quantity);

            var message = $"สั่งซื้อ {shoe.getName()} \nราคา {shoe.getPrice()} \nไซส์ {shoe.getSize()} \nจำนวน {shoe.getQuantity()}";

            using (var package = new ExcelPackage(new FileInfo(@"C:\Users\Admin\Desktop\Project\Project\Project\bin\Debug\shop.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                if (worksheet == null)
                {
                    throw new Exception("Worksheet does not exist");
                }
                int lastUsedRow = worksheet.Dimension.End.Row;

                ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                cell.Value = message;

                package.Save();
            }
            this.DialogResult = DialogResult.OK;
            this.Hide();
            MessageBox.Show(message);
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            name = this.label5.Text;
            price = this.label10.Text;
            length = this.comboBox4.SelectedItem.ToString();
            quantity = this.textBox4.Text;
            size = null;
            Sock sock = new Sock(name, price, length, quantity, size);

            var message = $"สั่งซ์้อ {sock.getName()} \nราคา {sock.getPrice()} \nความยาว{sock.getLength()} \nจำนวน{sock.getQuantity()}";

            using (var package = new ExcelPackage(new FileInfo(@"C:\Users\Admin\Desktop\Project\Project\Project\bin\Debug\shop.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                if (worksheet == null)
                {
                    throw new Exception("Worksheet does not exist");
                }
                int lastUsedRow = worksheet.Dimension.End.Row;

                ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                cell.Value = message;

                package.Save();
            }
            this.DialogResult = DialogResult.OK;
            this.Hide();
            MessageBox.Show(message);
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            name = this.label6.Text;
            price = this.label11.Text;
            color = this.comboBox5.SelectedItem.ToString();
            quantity = this.textBox5.Text;
            size = null;
            length = null;
            Band band = new Band(name, price, color, quantity, size, length);

            var message = $"สั่งซื้อ {band.getName()} \nราคา {band.getPrice()} \nสี{band.getColor()} \nจำวน{band.getQuantity()}";

            using (var package = new ExcelPackage(new FileInfo(@"C:\Users\Admin\Desktop\Project\Project\Project\bin\Debug\shop.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                if (worksheet == null)
                {
                    throw new Exception("Worksheet does not exist");
                }
                int lastUsedRow = worksheet.Dimension.End.Row;

                ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                cell.Value = message;

                package.Save();
            }
            this.DialogResult = DialogResult.OK;
            this.Hide();
            MessageBox.Show(message);
            Form2 form2 = new Form2();
            form2.ShowDialog();

        }
    }
}
