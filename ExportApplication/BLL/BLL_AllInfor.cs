using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL;
using DTO;

namespace BLL
{
    public class BLL_AllInfor
    {
        DAL_AllInfor dal_infor = new DAL_AllInfor();
        public DataTable GetToListView() {
            return dal_infor.GetDataToView();
        }
        public bool Insert(DTO_AllInfor dto_AllInfor) {
            return dal_infor.Insert(dto_AllInfor);
        }
        //test commit
    }
}
