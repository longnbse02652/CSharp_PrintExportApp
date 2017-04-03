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
                    keiyakusho();
                    break;
                case 2:
                    string zipcode = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
                    string temp1_zipcode = zipcode.Substring(0,3);
                    string temp2_zipcode = zipcode.Substring(3,4);

                    MessageBox.Show(temp1_zipcode + "-" + temp2_zipcode);
                    break;
                case 3:
                    string joindate = "1980/08/07";
                    string[] convert1 = joindate.Split('/');
                    string convert_birth = bll_handle.ConvertJapaneseCalendar(joindate);
                    string nienhieu = convert_birth.Substring(0,2);
                    string year = convert_birth.Substring(2,convert_birth.IndexOf("年")-2);
                    string month = convert1[1];
                    string day = convert1[2];

                    MessageBox.Show(nienhieu + year +"年"+ month +"月"+ day+"日");
                    break;
                case -1:
                    MessageBox.Show("印刷したい書類を選んでください！");
                    break;

            }
        }

        private object ValueOrDBNullIfZero(int val)
        {
            if (val == 0) return DBNull.Value;
            return val;
        }

        private void nyushanaiyousho()
        {
            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            try
            {
                
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(path+@"\File\template.xls", 
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
                //xlWorkSheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 441, 57, 117, 90);
                //}
                string temp_zairyuukigen = dt.Rows[0].Field<string>("CardTimeOut");
                string[] temps = temp_zairyuukigen.Split('/');
                xlWorkSheet.Cells[12, "BL"] = (Convert.ToInt32(temps[0]) - 1988).ToString();
                xlWorkSheet.Cells[12, "BP"] = temps[1];
                xlWorkSheet.Cells[12, "BT"] = temps[2];

                xlWorkSheet.Cells[12, "BX"] = dt.Rows[0].Field<string>("OutTime");
                xlWorkSheet.Cells[17, "Y"] = dt.Rows[0].Field<string>("CompanyName");
                xlWorkSheet.Cells[19, "BI"] = dt.Rows[0].Field<string>("WorkType");
                xlWorkSheet.Cells[18, "BU"] = dt.Rows[0].Field<string>("ClosingDate");
                xlWorkSheet.Cells[24, "U"] = dt.Rows[0].Field<int?>("HakenRyokin");
                xlWorkSheet.Cells[24, "AH"] = dt.Rows[0].Field<string>("HakenRyokinType");
                xlWorkSheet.Cells[24, "AY"] = dt.Rows[0].Field<string>("ShiharaiType");
                xlWorkSheet.Cells[24, "BU"] = dt.Rows[0].Field<string>("Tax");
                xlWorkSheet.Cells[28, "AE"] = dt.Rows[0].Field<string>("SalaryType");
                xlWorkSheet.Cells[28, "AK"] = dt.Rows[0].Field<int?>("BasicSalary");
                xlWorkSheet.Cells[29, "AE"] = dt.Rows[0].Field<int?>("SeikinTeate");
                xlWorkSheet.Cells[30, "AE"] = dt.Rows[0].Field<int?>("GaikinTeate");
                xlWorkSheet.Cells[31, "AE"] = dt.Rows[0].Field<int?>("GijutsuTeate");
                xlWorkSheet.Cells[32, "AE"] = dt.Rows[0].Field<int?>("ShikakuTeate");
                xlWorkSheet.Cells[33, "AE"] = dt.Rows[0].Field<int?>("YakushokuTeate");
                xlWorkSheet.Cells[34, "AE"] = dt.Rows[0].Field<int?>("EigyoTeate");

                xlWorkSheet.Cells[35, "AE"] = dt.Rows[0].Field<int?>("KazokuTeate");
                xlWorkSheet.Cells[36, "AE"] = dt.Rows[0].Field<int?>("JutakuTeate");
                xlWorkSheet.Cells[37, "AE"] = dt.Rows[0].Field<int?>("BekkyoTeate");
                xlWorkSheet.Cells[38, "AM"] = dt.Rows[0].Field<int?>("TsukinTeate");

                xlWorkSheet.Cells[30, "BU"] = dt.Rows[0].Field<int?>("Park");
                xlWorkSheet.Cells[31, "BU"] = dt.Rows[0].Field<int?>("DormitoryFee");
                xlWorkSheet.Cells[32, "BU"] = dt.Rows[0].Field<int?>("WaterFee");

                xlWorkSheet.Cells[41, "G"] = dt.Rows[0].Field<string>("EmployStatus");
                if (dt.Rows[0].Field<string>("EmployStatus") != "正社員")
                {
                    string temp_time1 = dt.Rows[0].Field<string>("EmployTime1");
                    string[] Time1_temps = temp_time1.Split('/');
                    xlWorkSheet.Cells[41, "AG"] = (Convert.ToInt32(Time1_temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[41, "AL"] = Time1_temps[1];
                    xlWorkSheet.Cells[41, "AP"] = Time1_temps[2];

                    string temp_time2 = dt.Rows[0].Field<string>("EmployTime2");
                    string[] Time2_temps = temp_time2.Split('/');
                    xlWorkSheet.Cells[41, "AX"] = (Convert.ToInt32(Time2_temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[41, "BC"] = Time2_temps[1];
                    xlWorkSheet.Cells[41, "BG"] = Time2_temps[2];
                }

                xlWorkSheet.Cells[53, "C"] = dt.Rows[0].Field<string>("BankName");
                xlWorkSheet.Cells[53, "AB"] = dt.Rows[0].Field<string>("BankNameType");
                xlWorkSheet.Cells[53, "AG"] = dt.Rows[0].Field<string>("BranchName");
                xlWorkSheet.Cells[53, "BB"] = dt.Rows[0].Field<string>("BranchNameType");
                xlWorkSheet.Cells[53, "BG"] = dt.Rows[0].Field<string>("AccountName");
                xlWorkSheet.Cells[56, "C"] = dt.Rows[0].Field<string>("BankCode");
                xlWorkSheet.Cells[56, "AG"] = dt.Rows[0].Field<string>("BranchCode");

                xlWorkSheet.Cells[56, "BI"] = dt.Rows[0].Field<string>("AccountCode1");
                xlWorkSheet.Cells[56, "BL"] = dt.Rows[0].Field<string>("AccountCode2");
                xlWorkSheet.Cells[56, "BO"] = dt.Rows[0].Field<string>("AccountCode3");
                xlWorkSheet.Cells[56, "BR"] = dt.Rows[0].Field<string>("AccountCode4");
                xlWorkSheet.Cells[56, "BU"] = dt.Rows[0].Field<string>("AccountCode5");
                xlWorkSheet.Cells[56, "BX"] = dt.Rows[0].Field<string>("AccountCode6");
                xlWorkSheet.Cells[56, "CA"] = dt.Rows[0].Field<string>("AccountCode7");
                xlWorkSheet.Cells[56, "CD"] = dt.Rows[0].Field<string>("AccountCode8");

                if (dt.Rows[0].Field<string>("TravelType") == "通勤")
                {
                    xlWorkSheet.Cells[59, "F"] = "☑";
                }
                else
                {
                    xlWorkSheet.Cells[61, "R"] = dt.Rows[0].Field<string>("HouseName");
                    xlWorkSheet.Cells[61, "AY"] = dt.Rows[0].Field<string>("Room");
                    string inhousedate = dt.Rows[0].Field<string>("InHouseDate");
                    string[] inhousedate_split = inhousedate.Split('/');
                    xlWorkSheet.Cells[61, "BR"] = (Convert.ToInt32(inhousedate_split[0]) - 1988).ToString();
                    xlWorkSheet.Cells[61, "BW"] = inhousedate_split[1];
                    xlWorkSheet.Cells[61, "CB"] = inhousedate_split[2];
                }

                xlWorkSheet.Cells[65, "AH"] = dt.Rows[0].Field<string>("Kouyouhoken");
                xlWorkSheet.Cells[67, "AH"] = dt.Rows[0].Field<string>("Shakaihoken");
                xlWorkSheet.Cells[67, "BC"] = dt.Rows[0].Field<int?>("DependentPeople");
                xlWorkSheet.Cells[67, "BM"] = dt.Rows[0].Field<int?>("ResidentPeople");
                xlWorkSheet.Cells[67, "BW"] = dt.Rows[0].Field<int?>("HealthInsurancePeople");

                xlWorkSheet.Cells[4, "BG"] = dt.Rows[0].Field<string>("CreatePeople");
                xlWorkSheet.Cells[3, "BG"] = dt.Rows[0].Field<string>("Position");

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
                                     printers[printerIndex], Type.Missing, Type.Missing, Type.Missing);

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
            catch (Exception e)
            {
                // Cleanup Memory
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.FinalReleaseComObject(xlWorkSheet);

                xlWorkBook.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(xlWorkBook);

                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
                MessageBox.Show(e.Message, "Error Message");
            }
        }

        private void keiyakusho(){
            String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            try
            {
                
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(path+@"\File\template.xls", 
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

                xlWorkSheet.Cells[5, "G"] = dt.Rows[0].Field<string>("RomajiName");
                xlWorkSheet.Cells[4, "G"] = dt.Rows[0].Field<string>("FuriganaName");
                xlWorkSheet.Cells[5, "T"] = dt.Rows[0].Field<string>("Sex");
                //Birthday
                string birth = dt.Rows[0].Field<string>("Birth");
                string[] convert1 = birth.Split('/');
                string convert_birth = bll_handle.ConvertJapaneseCalendar(birth);
                xlWorkSheet.Cells[4, "Z"] = convert_birth.Substring(0, 2); 
                xlWorkSheet.Cells[4, "AB"] = convert_birth.Substring(2, convert_birth.IndexOf("年")-2);
                xlWorkSheet.Cells[4, "AD"] = convert1[1];
                xlWorkSheet.Cells[4, "AG"] = convert1[2];

                //Zipcode
                string zipcode = (dt.Rows[0].Field<int?>("ZipCode")).ToString();
                if(zipcode.Length == 7){
                    string temp1_zipcode = zipcode.Substring(0, 3);
                    string temp2_zipcode = zipcode.Substring(3, 4);
                    xlWorkSheet.Cells[6, "H"] = temp1_zipcode;
                    xlWorkSheet.Cells[6, "K"] = temp2_zipcode;
                }
                //Address
                xlWorkSheet.Cells[7, "G"] = dt.Rows[0].Field<string>("Address");

                //Mobiphone
                string mobiphone = dt.Rows[0].Field<string>("MobliePhone") ;
                if(mobiphone.Length == 11){
                    string mobi1 = mobiphone.Substring(0, 3);
                    string mobi2 = mobiphone.Substring(3, 4);
                    string mobi3 = mobiphone.Substring(7, 4);
                    xlWorkSheet.Cells[8, "W"] = mobi1;
                    xlWorkSheet.Cells[8, "AA"] = mobi2;
                    xlWorkSheet.Cells[8, "AD"] = mobi3;
                }
                //Phone
                string phone = dt.Rows[0].Field<string>("Phone");
                if(phone.Length == 10){
                    string phone1 = phone.Substring(0, 2);
                    string phone2 = phone.Substring(2, 4);
                    string phone3 = phone.Substring(6, 4);
                    xlWorkSheet.Cells[8, "I"] = phone1;
                    xlWorkSheet.Cells[8, "L"] = phone2;
                    xlWorkSheet.Cells[8, "O"] = phone3;
                }
                //Join company Date
                string joindate = dt.Rows[0].Field<string>("InCompanyDate");
                string[] joindate_temps = joindate.Split('/');
                xlWorkSheet.Cells[10, "I"] = (Convert.ToInt32(joindate_temps[0]) - 1988).ToString();
                xlWorkSheet.Cells[10, "K"] = joindate_temps[1];
                xlWorkSheet.Cells[10, "M"] = joindate_temps[2];
                //keiyaku time
                if (dt.Rows[0].Field<string>("EmployStatus") != "正社員")
                {
                    xlWorkSheet.Cells[10, "Q"] = "□";
                    xlWorkSheet.Cells[10, "V"] = "☑";
                    string temp_time2 = dt.Rows[0].Field<string>("EmployTime2");
                    string[] Time2_temps = temp_time2.Split('/');
                    xlWorkSheet.Cells[10, "AB"] = (Convert.ToInt32(Time2_temps[0]) - 1988).ToString();
                    xlWorkSheet.Cells[10, "AD"] = Time2_temps[1];
                    xlWorkSheet.Cells[10, "AF"] = Time2_temps[2];
                }
                //ContracType
                //ContractRequire

                //My company
                xlWorkSheet.Cells[14, "N"] = dt.Rows[0].Field<string>("MyCompany");
                xlWorkSheet.Cells[15, "G"] = dt.Rows[0].Field<string>("WorkContent");
                //Worktime
                xlWorkSheet.Cells[16, "L"] = dt.Rows[0].Field<string>("WorkTime1");
                xlWorkSheet.Cells[16, "N"] = dt.Rows[0].Field<string>("WorkTime2");
                xlWorkSheet.Cells[16, "U"] = dt.Rows[0].Field<string>("WorkTime3");
                xlWorkSheet.Cells[16, "W"] = dt.Rows[0].Field<string>("WorkTime4");
                xlWorkSheet.Cells[16, "AG"] = dt.Rows[0].Field<string>("RelaxTime");
                //賃金
                xlWorkSheet.Cells[24, "O"] = dt.Rows[0].Field<int?>("BasicSalary");
                xlWorkSheet.Cells[32, "P"] = dt.Rows[0].Field<int?>("SeikinTeate");
                xlWorkSheet.Cells[32, "AC"] = dt.Rows[0].Field<int?>("GaikinTeate");
                xlWorkSheet.Cells[33, "P"] = dt.Rows[0].Field<int?>("GijutsuTeate");
                xlWorkSheet.Cells[33, "AC"] = dt.Rows[0].Field<int?>("ShikakuTeate");
                xlWorkSheet.Cells[34, "P"] = dt.Rows[0].Field<int?>("YakushokuTeate");
                xlWorkSheet.Cells[34, "AC"] = dt.Rows[0].Field<int?>("EigyoTeate");
                xlWorkSheet.Cells[35, "P"] = dt.Rows[0].Field<int?>("KazokuTeate");
                xlWorkSheet.Cells[35, "AC"] = dt.Rows[0].Field<int?>("JutakuTeate");
                xlWorkSheet.Cells[36, "P"] = dt.Rows[0].Field<int?>("BekkyoTeate");
                xlWorkSheet.Cells[36, "AE"] = dt.Rows[0].Field<int?>("TsukinTeate");
                //寮費
                xlWorkSheet.Cells[40, "P"] = dt.Rows[0].Field<int?>("DormitoryFee");
                xlWorkSheet.Cells[42, "N"] = dt.Rows[0].Field<string>("ClosingDate");
                xlWorkSheet.Cells[40, "P"] = dt.Rows[0].Field<int?>("DormitoryFee");

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
                                     printers[printerIndex], Type.Missing, Type.Missing, Type.Missing);

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
            catch (Exception e)
            {
                // Cleanup Memory
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.FinalReleaseComObject(xlWorkSheet);

                xlWorkBook.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(xlWorkBook);

                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
                MessageBox.Show(e.Message, "エラー");
            }
        }
    }
}
