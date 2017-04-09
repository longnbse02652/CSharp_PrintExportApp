using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BLL
{
    public class BLL_HandleFunc
    {
        //ham nay la de chuyen dinh dang cua Nhat ve YYMMDD de them vao database
        public string ConvertFromDatetimePicker_ToYYMMDD(DateTimePicker dtPicker) {
            string date = dtPicker.Value.ToString("yyyy/MM/dd");
            return date;
        }
        //hàm này là để chuyển định dang datime thông thường sang datetime của Nhật de hien thi len
        public string ConvertJapaneseCalendar(string datetime)
        {
            //string datetime = "1993/10/14";
            DateTime dt = Convert.ToDateTime(datetime);
            JapaneseCalendar myCal = new JapaneseCalendar();

            switch (myCal.GetEra(dt).ToString())
            {
                case "1":
                    datetime = "明治" + myCal.GetYear(dt) + "年" + myCal.GetMonth(dt) + "月" + myCal.GetDayOfMonth(dt) + "日";
                    break;
                case "2":
                    datetime = "大正" + myCal.GetYear(dt) + "年" + myCal.GetMonth(dt) + "月" + myCal.GetDayOfMonth(dt) + "日";
                    break;
                case "3":
                    datetime = "昭和" + myCal.GetYear(dt) + "年" + myCal.GetMonth(dt) + "月" + myCal.GetDayOfMonth(dt) + "日";
                    break;
                case "4":
                    datetime = "平成" + myCal.GetYear(dt) + "年" + myCal.GetMonth(dt) + "月" + myCal.GetDayOfMonth(dt) + "日";
                    break;
            }
            return datetime;
        }

        //ham nay de xuat hien Error neu user ko nhap du lieu vao textbox
        public bool ValidateControls(Control control, ErrorProvider errorProvider)
        {
            bool bStatus = true;
            if (control.Text == "" || control.Text == null)
            {
                errorProvider.SetError(control, "ここに入力してください");
                bStatus = false;
               // control.Focus();
            }
            else
                errorProvider.SetError(control, "");
            return bStatus;
        }

        //ham nay de show dia chi thong qua zipcode
        public void AutoShowAddress(TextBox tb_ZipCode, TextBox tb_Address1, ComboBox cb_Address2, TextBox tb_Address3,
                                    ComboBox cb_Address4, TextBox tb_Address5)
        {

            string sKey = tb_ZipCode.Text;
            if (sKey.Length == 7)
            {
               
                Cursor.Current = Cursors.WaitCursor;
                sKey = sKey.Trim();
                //  sKey = Strings.StrConv(sKey, VbStrConv.Narrow, 0);
                String path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                try
                {
                    StreamReader sr = new StreamReader(path + @"\File\KEN_ALL.txt", Encoding.Default);

                    string dat;
                    while ((dat = sr.ReadLine()) != null)
                    {
                        string tmpZip;
                        string[] sbuf = dat.Split(',');
                        tmpZip = sbuf[2].Trim();

                        if (sKey == tmpZip)
                        {
                            tb_Address1.Text = sbuf[6].Substring(0, sbuf[6].Length - 1);
                            cb_Address2.Text = sbuf[6].Last().ToString();
                            tb_Address3.Text = sbuf[7].Substring(0, sbuf[7].Length - 1);
                            cb_Address4.Text = sbuf[7].Last().ToString();
                            tb_Address5.Text = sbuf[8].Trim();
                            break;
                        }
                        Application.DoEvents();
                    }

                    sr.Close();
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message, "ファイルエラー",
                                   MessageBoxButtons.OK,
                                   MessageBoxIcon.Error);
                    return; 
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }
            }

        }

    }
}
