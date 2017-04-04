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
   public class BLL_Travel
    {
        DAL_Travel dal_travel = new DAL_Travel();
        public DataTable TravelForm(string nameRomaji)
        {
            return dal_travel.TravelForm(nameRomaji);
        }

        public bool Insert(DTO_Travel dto_travel)
        {
            return dal_travel.Update(dto_travel);
        }
    }
}
