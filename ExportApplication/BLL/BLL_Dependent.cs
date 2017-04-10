using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using DAL;
using System.Data;

namespace BLL
{
   public class BLL_Dependent
    {
        DAL_Dependent dal_dependent = new DAL_Dependent();
        public DataTable DependentForm(string nameRomaji)
        {
            return dal_dependent.DependentForm(nameRomaji);
        }

        public bool Insert(DTO_Dependent dto_dependent)
        {
            return dal_dependent.Update(dto_dependent);
        }
       //test
    }
}
