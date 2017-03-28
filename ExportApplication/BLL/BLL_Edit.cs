using DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public class BLL_Edit
    {
        DAL_Edit dal_edit = new DAL_Edit();
        public DataTable EditForm()
        {
            return dal_edit.EditForm();
        }
    }
}
