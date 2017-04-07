using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL;
using System.Data;

namespace BLL
{
    public class BLL_View
    {
        DAL_View dal_view = new DAL_View();

        public DataTable GetDataToView(string name)
        {
            return dal_view.GetDataToView(name);
        }
    }
}
