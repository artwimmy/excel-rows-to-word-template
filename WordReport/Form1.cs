using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

// com objects
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace WordReport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
/*            string res = "C:\\Users\\twimy\\OneDrive\\Desktop\\FINKI\\iSolve\\C# Word templates\\Form.xlsx";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(res, 0, true, 5, "", "", true);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            */
                                             
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //dokumenti excel per numrimin e reshtave...
            //string exceldoc = "C:\\Users\\twimy\\OneDrive\\Desktop\\FINKI\\iSolve\\C# Word templates\\Form.xlsx";
            string exceldoc = textBox1.Text;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(exceldoc);
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            Excel.Range userRange = x.UsedRange;
            int countRecords = userRange.Rows.Count;
            //int add = countRecords +1;
            //x.Cells[add, 1] = "Total Rows "+countRecords;

            //progress bar
            this.progressBar1.Maximum = 100;

            //dokumenti template
            Microsoft.Office.Interop.Word.Document doc = null;
            object fileName = textBox2.Text;
            //object fileName = "C:\\Users\\twimy\\OneDrive\\Desktop\\FINKI\\iSolve\\C# Word templates\\Template.docx";
            object missing = Type.Missing;

            for(int i=0; i<countRecords; i++)
            {
                doc = app.Documents.Open(fileName, missing, missing);
                app.Selection.Find.ClearFormatting();
                app.Selection.Find.Replacement.ClearFormatting();

                string[] tmp = new string[4];
                tmp = readExcel(i);

                app.Selection.Find.Execute("<ID>", missing, missing, missing, missing, missing, missing, missing, missing, tmp[0],2);
                app.Selection.Find.Execute("<name>", missing, missing, missing, missing, missing, missing, missing, missing, tmp[1],2);
                app.Selection.Find.Execute("<sex>", missing, missing, missing, missing, missing, missing, missing, missing, tmp[2],2);
                app.Selection.Find.Execute("<age>", missing, missing, missing, missing, missing, missing, missing, missing, tmp[3],2);


                object SaveAsFile = (object) textBox3.Text + tmp[0] + ".doc";
                doc.SaveAs2(SaveAsFile, missing, missing, missing);
                this.progressBar1.Value += (100 / countRecords);
            }
            this.progressBar1.Value = 100;
            MessageBox.Show("Fajllat u krijuan me sukses!");
            doc.Close(false, missing, missing);
            app.Quit(false, false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }

        // read data from the excel
        private string[] readExcel(int index)
        {
            //string res = "C:\\Users\\twimy\\OneDrive\\Desktop\\FINKI\\iSolve\\C# Word templates\\Form.xlsx";
            string res = textBox1.Text;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            //int rowCount = xlWorkSheet.Rows.Count();

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(res, 0, true, 5, "", "", true);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            index += 2;

            string[] data = new string[4]; // {ID, EMRI, GJINIA, MOSHA
            data[0] = xlWorkSheet.get_Range("A" + index.ToString()).Text; // ID
            data[1] = xlWorkSheet.get_Range("B" + index.ToString()).Value; // emri
            data[2] = xlWorkSheet.get_Range("C" + index.ToString()).Value; // gjinia
            data[3] = xlWorkSheet.get_Range("D" + index.ToString()).Text; // mosha
            
            xlWorkBook.Close(false);
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            return data;
        }

        OpenFileDialog ofd = new OpenFileDialog();
        private void button2_Click(object sender, EventArgs e)
        {
            ofd.Filter = "EXCEL|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;
            }
        }

        OpenFileDialog wrd = new OpenFileDialog();
        private void button3_Click(object sender, EventArgs e)
        {
            wrd.Filter = "WORD|*.docx";
            if(wrd.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = wrd.FileName;
            }
        }

        FolderBrowserDialog reps = new FolderBrowserDialog();
        private void button4_Click(object sender, EventArgs e)
        {
            if (reps.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = reps.SelectedPath;
            }
        }
    }
}
