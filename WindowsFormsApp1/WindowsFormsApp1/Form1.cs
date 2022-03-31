using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = dialog.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            string[] allFoundFiles = Directory.GetFiles(textBox1.Text, "*.jpg", SearchOption.AllDirectories);

            string str = textBox1.Text + "\\отчет.xlsx";
            int leng = allFoundFiles.Length;
            progressBar1.Maximum = leng;
            label3.Text = "0 %";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook;
            Excel.Worksheet workSheet;

            workBook = excelApp.Workbooks.Add();
            workBook.Application.DisplayAlerts = false;
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
            workSheet.Name = "Отчет";

            workSheet.Cells[1, 1] = "Номер";
            workSheet.Cells[1, 2] = "Картинка";
            workSheet.Cells[1, 3] = "...";
            workSheet.Cells[1, 4] = "Построение";
            workSheet.Columns[2].ColumnWidth = 30;
            workSheet.Columns[4].ColumnWidth = 20;

            Excel.Range rangeStyle = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[leng + 2, 4]];
            rangeStyle.Cells.Font.Name = "Times New Roman";
            rangeStyle.Cells.Font.Size = 10;
            rangeStyle.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rangeStyle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            for (int i = 1; i <= leng; i++)
            {
                workSheet.Rows[i+1].RowHeight = 120;
                workSheet.Cells[i + 1, 1] = i;
                var image = new Bitmap(allFoundFiles[i - 1]);
                Excel.Range imageRange = (Excel.Range)workSheet.Cells[i+1, 2];
                float left = (float)((double)imageRange.Left);
                float top = (float)((double)imageRange.Top);
                const float imageSize = 100;
                workSheet.Shapes.AddPicture(allFoundFiles[i - 1], Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, imageSize, imageSize);
                label3.Text = (i / leng) * 100 + " %";
            }
            workSheet.Range["D2", "D" + (leng + 1)].NumberFormat = "#'###.00 руб";



            workSheet.Cells[leng + 2, 4].FormulaLocal = "=Сумм(D2:D" + (leng+1) + ")";
            workSheet.Cells[leng + 2, 3].FormulaLocal = "Сумма";
            excelApp.Application.ActiveWorkbook.SaveAs(str);
            workBook.Close(true);
            excelApp.Quit();

            label3.Text = "Готово";

            allFoundFiles = null;
            rangeStyle = null;
            workSheet = null;
            workBook = null;
            excelApp = null;
        }
    }
}
