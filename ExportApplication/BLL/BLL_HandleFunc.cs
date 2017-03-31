using System;
using System.Collections.Generic;
using System.Globalization;
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
            //string datetime = "1993-10-14";
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

    }
}
