using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL;

namespace BLL
{
    public class BLL_Infor
    {
        DAL_Infor dal_infor = new DAL_Infor();
        public DataTable GetToListView() {
            return dal_infor.GetToListView();
        }
    }
}
