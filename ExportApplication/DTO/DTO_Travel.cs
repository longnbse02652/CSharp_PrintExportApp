using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
   public class DTO_Travel
    {
       private string _Trainsportation1;
       public string Trainsportation1 { get { return _Trainsportation1; } set { _Trainsportation1 = value; } }

       private string _BeginTrain1;
       public string BeginTrain1 { get { return _BeginTrain1; } set { _BeginTrain1 = value; } }

       private string _EndTrain1;
       public string EndTrain1 { get { return _EndTrain1; } set { _EndTrain1 = value; } }

       private string _Trainsportation2;
       public string Trainsportation2 { get { return _Trainsportation2; } set { _Trainsportation2 = value; } }

       private string _BeginTrain2;
       public string BeginTrain2 { get { return _BeginTrain2; } set { _BeginTrain2 = value; } }

       private string _EndTrain2;
       public string EndTrain2 { get { return _EndTrain2; } set { _EndTrain2 = value; } }

       private string _Trainsportation3;
       public string Trainsportation3 { get { return _Trainsportation3; } set { _Trainsportation3 = value; } }

       private string _BeginTrain3;
       public string BeginTrain3 { get { return _BeginTrain3; } set { _BeginTrain3 = value; } }

       private string _EndTrain3;
       public string EndTrain3 { get { return _EndTrain3; } set { _EndTrain3 = value; } }

       private string _Trainsportation4;
       public string Trainsportation4 { get { return _Trainsportation4; } set { _Trainsportation4 = value; } }

       private string _BeginTrain4;
       public string BeginTrain4 { get { return _BeginTrain4; } set { _BeginTrain4 = value; } }

       private string _EndTrain4;
       public string EndTrain4 { get { return _EndTrain4; } set { _EndTrain4 = value; } }

       private int _MonthRegular1;
       public int MonthRegular1 { get { return _MonthRegular1; } set { _MonthRegular1 = value; } }
       
       private int _MonthRegular2;
       public int MonthRegular2 { get { return _MonthRegular2; } set { _MonthRegular2 = value; } }

       private int _MonthRegular3;
       public int MonthRegular3 { get { return _MonthRegular3; } set { _MonthRegular3 = value; } }

       private int _MonthRegular4;
       public int MonthRegular4 { get { return _MonthRegular4; } set { _MonthRegular4 = value; } }

       private string _Carkm;
       public string Carkm { get { return _Carkm; } set { _Carkm = value; } }

       private string _RomajiName;
       public string RomajiName { get { return _RomajiName; } set { _RomajiName = value; } }
       private int _CarMoney;
       public int CarMoney { get { return _CarMoney; } set { _CarMoney = value; } }

       private int _TotalMoneyTrans;
       public int TotalMoneyTrans { get { return _TotalMoneyTrans; } set { _TotalMoneyTrans = value; } }
  
       public DTO_Travel(string eTrainsportation1, string eBeginTrain1, string eEndTrain1,
       string eTrainsportation2, string eBeginTrain2, string eEndTrain2,string eTrainsportation3, string eBeginTrain3, string eEndTrain3,
       string eTrainsportation4, string eBeginTrain4, string eEndTrain4,int eMonthRegular1,int eMonthRegular2, int eMonthRegular3,int eMonthRegular4,
       string eCarKm, int eCarMoney, int eTotalMoneyTrans, string eRomajiName)
        {
            this._Trainsportation1 = eTrainsportation1;
            this._Trainsportation2 = eTrainsportation2;
            this._Trainsportation3 = eTrainsportation3;
            this._Trainsportation4 = eTrainsportation4;
            this._BeginTrain1 = eBeginTrain1;
            this._BeginTrain2 = eBeginTrain2;
            this._BeginTrain3 = eBeginTrain3;
            this._BeginTrain4 = eBeginTrain4;
            this._EndTrain1 = eEndTrain1;
            this._EndTrain2 = eEndTrain2;
            this._EndTrain3 = eEndTrain3;
            this._EndTrain4 = eEndTrain4;
            this._MonthRegular1 = eMonthRegular1;
            this._MonthRegular2 = eMonthRegular2;
            this._MonthRegular3 = eMonthRegular3;
            this._MonthRegular4 = eMonthRegular4;
            this._Carkm = eCarKm;
            this._CarMoney = eCarMoney;
            this._TotalMoneyTrans = eTotalMoneyTrans;
            this._RomajiName = eRomajiName;
        }
   }
}
