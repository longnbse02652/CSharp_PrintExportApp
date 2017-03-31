using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL;
using System.Data;

namespace BLL
{
    public class BLL_Print
    {
        DAL_Print dal_print = new DAL_Print();

        public DataTable GetDataToPrint(string name) {
            return dal_print.GetDataToPrint(name);
        }
    }
}
