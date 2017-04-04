using DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;

namespace BLL
{
    public class BLL_Edit
    {
        DAL_Edit dal_edit = new DAL_Edit();
        public DataTable EditForm(string name)
        {
            return dal_edit.EditForm(name);
        }

        public bool Insert(DTO_Edit dto_Edit)
        {
            return dal_edit.Update(dto_Edit);
        }
    }
}
