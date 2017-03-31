using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using BLL;
using System.IO;
using ExportApplication.Properties;

namespace ExportApplication
{
    public partial class GUI_Print : Form
    {
        BLL_Print bll_print = new BLL_Print();
        BLL_HandleFunc bll_handle = new BLL_HandleFunc();
        DataTable dt = new DataTable();

        Excel.Application xlApp = null;
        Excel.Workbook xlWorkBook = null;
        Excel.Worksheet xlWorkSheet = null;
       
        public GUI_Print(string getName)
        {
            InitializeComponent();
            dt = bll_print.GetDataToPrint(getName);
            //dt.Rows[0].Field<string>("RomajiName");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    nyushanaiyousho();
                    break;
                case 1:
                    String path = dt.Rows[0].Field<string>("RomajiName");

                    MessageBox.Show(path);
                    break;
                case 2:
                    MessageBox.Show("2");
                    break;
                case 3:
                    MessageBox.Show("3");
                    break;
                case -1:
                    MessageBox.Show("印刷したい書類を選んでください！");
                    break;

            }
        }

        private void nyushanaiyousho()
        {
            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            try
            {
                
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\GA-TOK147\Desktop\プロジェクト\template.xls", 
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[10, "X"] = dt.Rows[0].Field<string>("RomajiName");
                xlWorkSheet.Cells[8, "X"] = dt.Rows[0].Field<string>("FuriganaName");
                xlWorkSheet.Cells[10, "AY"] = dt.Rows[0].Field<string>("Sex");
                xlWorkSheet.Cells[7, "DK"] = dt.Rows[0].Field<string>("Birth");
                xlWorkSheet.Cells[11, "DK"] = dt.Rows[0].Field<string>("InCompanyDate");
                //
                if (dt.Rows[0].Field<string>("Nationality") != string.Empty)
                {
                    xlWorkSheet.Cells[11, "H"] = "日本以外は国名記入";
                    xlWorkSheet.Cells[11, "H"] = dt.Rows[0].Field<string>("Nationality");
                }
                else {
                    xlWorkSheet.Cells[11, "H"] = "日本";
                }
                //
               
                //if (dt.Rows[0].Field<string>("CardType") == "定住者")
                //{
                Excel.Shapes shp = xlWorkSheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 441, 57, 117.75, 90.75);
                //}
                string temp_zairyuukigen = bll_handle.ConvertJapaneseCalendar(dt.Rows[0].Field<string>("CardTimeOut"));

                

                //cho nay de xu ly may in default
                var printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters;
                int printerIndex = 0;
                foreach (String s in printers)
                {
                    if (s.Equals("白黒　SHARP MX-2650FN SPDL2-c"))
                    {
                        break;
                    }
                    printerIndex++;
                }

                // Print out 1 copy to the default printer:
                xlWorkSheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                     Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // Cleanup:
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.FinalReleaseComObject(xlWorkSheet);

                xlWorkBook.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(xlWorkBook);

                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
                MessageBox.Show("印刷完了");
            }
            catch
            {
                MessageBox.Show("印刷できません");

            }




        }
    }
}
