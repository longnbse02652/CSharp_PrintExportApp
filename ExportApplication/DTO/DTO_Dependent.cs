using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
   public  class DTO_Dependent
    {
       private string _DependentPeopleKana1;
       public string DependentPeopleKana1 { get { return _DependentPeopleKana1; } set { _DependentPeopleKana1 = value; } }

       private string _DependentPeopleKana2;
       public string DependentPeopleKana2 { get { return _DependentPeopleKana2; } set { _DependentPeopleKana2 = value; } }
      
       private string _DependentPeopleKana3;
       public string DependentPeopleKana3 { get { return _DependentPeopleKana3; } set { _DependentPeopleKana3 = value; } }
      
       private string _DependentPeopleKana4;
       public string DependentPeopleKana4 { get { return _DependentPeopleKana4; } set { _DependentPeopleKana4 = value; } }
       
       private string _DependentPeopleKana5;
       public string DependentPeopleKana5 { get { return _DependentPeopleKana5; } set { _DependentPeopleKana5 = value; } }

       private string _DependentPeopleKana6;
       public string DependentPeopleKana6 { get { return _DependentPeopleKana6; } set { _DependentPeopleKana6 = value; } }

       private string _DependentPeopleShimei1;
       public string DependentPeopleShimei1 { get { return _DependentPeopleShimei1; } set { _DependentPeopleShimei1 = value; } }

       private string _DependentPeopleShimei2;
       public string DependentPeopleShimei2 { get { return _DependentPeopleShimei2; } set { _DependentPeopleShimei2 = value; } }
      
       private string _DependentPeopleShimei3;
       public string DependentPeopleShimei3 { get { return _DependentPeopleShimei3; } set { _DependentPeopleShimei3 = value; } }
      
       private string _DependentPeopleShimei4;
       public string DependentPeopleShimei4 { get { return _DependentPeopleShimei4; } set { _DependentPeopleShimei4 = value; } }
       
       private string _DependentPeopleShimei5;
       public string DependentPeopleShimei5 { get { return _DependentPeopleShimei5; } set { _DependentPeopleShimei5 = value; } }

       private string _DependentPeopleShimei6;
       public string DependentPeopleShimei6 { get { return _DependentPeopleShimei6; } set { _DependentPeopleShimei6 = value; } }

        private string _Relationship1;
       public string Relationship1 { get { return _Relationship1; } set { _Relationship1 = value; } }

       private string _Relationship2;
       public string Relationship2 { get { return _Relationship2; } set { _Relationship2 = value; } }
      
       private string _Relationship3;
       public string Relationship3 { get { return _Relationship3; } set { _Relationship3 = value; } }
      
       private string _Relationship4;
       public string Relationship4 { get { return _Relationship4; } set { _Relationship4 = value; } }
       
       private string _Relationship5;
       public string Relationship5 { get { return _Relationship5; } set { _Relationship5 = value; } }

       private string _Relationship6;
       public string Relationship6 { get { return _Relationship6; } set { _Relationship6 = value; } }

        private string _DependentPeopleBirth1;
       public string DependentPeopleBirth1 { get { return _DependentPeopleBirth1; } set { _DependentPeopleBirth1 = value; } }

       private string _DependentPeopleBirth2;
       public string DependentPeopleBirth2 { get { return _DependentPeopleBirth2; } set { _DependentPeopleBirth2 = value; } }
      
       private string _DependentPeopleBirth3;
       public string DependentPeopleBirth3 { get { return _DependentPeopleBirth3; } set { _DependentPeopleBirth3 = value; } }
      
       private string _DependentPeopleBirth4;
       public string DependentPeopleBirth4 { get { return _DependentPeopleBirth4; } set { _DependentPeopleBirth4 = value; } }
       
       private string _DependentPeopleBirth5;
       public string DependentPeopleBirth5 { get { return _DependentPeopleBirth5; } set { _DependentPeopleBirth5 = value; } }

       private string _DependentPeopleBirth6;
       public string DependentPeopleBirth6 { get { return _DependentPeopleBirth6; } set { _DependentPeopleBirth6 = value; } }

       private string _Living1;
       public string Living1 { get { return _Living1; } set { _Living1 = value; } }
       private string _Living2;
       public string Living2 { get { return _Living2; } set { _Living2 = value; } }
       private string _Living3;
       public string Living3 { get { return _Living3; } set { _Living3 = value; } }
       private string _Living4;
       public string Living4 { get { return _Living4; } set { _Living4 = value; } }
       private string _Living5;
       public string Living5 { get { return _Living5; } set { _Living5 = value; } }
       private string _Living6;
       public string Living6 { get { return _Living6; } set { _Living6 = value; } }
       private string _RomajiName;
       public string RomajiName { get { return _RomajiName; } set { _RomajiName = value; } }

       private int _DependentPeople;
       public int DependentPeople { get { return _DependentPeople; } set { _DependentPeople = value; } }

       private int _ResidentPeople;
       public int ResidentPeople { get { return _ResidentPeople; } set { _ResidentPeople = value; } }

       private int _HealthInsurancePeople;
       public int HealthInsurancePeople { get { return _HealthInsurancePeople; } set { _HealthInsurancePeople = value; } }

       public DTO_Dependent(string eDependentPeopleKana1, string eDependentPeopleKana2, string eDependentPeopleKana3,
       string eDependentPeopleKana4, string eDependentPeopleKana5,string eDependentPeopleKana6, string eDependentPeopleShimei1, string eDependentPeopleShimei2, string eDependentPeopleShimei3, string eDependentPeopleShimei4,
       string eDependentPeopleShimei5, string eDependentPeopleShimei6, string eRelationship1, string eRelationship2, string eRelationship3, string eRelationship4, string eRelationship5,string eRelationship6, string eDependentPeopleBirth1,
       string eDependentPeopleBirth2, string eDependentPeopleBirth3, string eDependentPeopleBirth4, string eDependentPeopleBirth5,string eDependentPeopleBirth6, string eLiving1, string eLiving2, string eLiving3, string eLiving4, string eLiving5,string eLiving6,string eRomajiName,
           int eDependentPeople, int eResidentPeople, int eHealthInsurancePeople)
        {
            this._DependentPeopleKana1 = eDependentPeopleKana1;
            this._DependentPeopleKana2 = eDependentPeopleKana2;
            this._DependentPeopleKana3 = eDependentPeopleKana3;
            this._DependentPeopleKana4 = eDependentPeopleKana4;
            this._DependentPeopleKana5 = eDependentPeopleKana5;
            this._DependentPeopleKana6 = eDependentPeopleKana6;
            this._DependentPeopleShimei1 = eDependentPeopleShimei1;
            this._DependentPeopleShimei2 = eDependentPeopleShimei2;
            this._DependentPeopleShimei3 = eDependentPeopleShimei3;
            this._DependentPeopleShimei4 = eDependentPeopleShimei4;
            this._DependentPeopleShimei5 = eDependentPeopleShimei5;
            this._DependentPeopleShimei6 = eDependentPeopleShimei6;
            this._Relationship1 = eRelationship1;
            this._Relationship2 = eRelationship2;
            this._Relationship3 = eRelationship3;
            this._Relationship4 = eRelationship4;
            this._Relationship5 = eRelationship5;
            this._Relationship6 = eRelationship6;
            this._DependentPeopleBirth1 = eDependentPeopleBirth1;
            this._DependentPeopleBirth2 = eDependentPeopleBirth2;
            this._DependentPeopleBirth3 = eDependentPeopleBirth3;
            this._DependentPeopleBirth4 = eDependentPeopleBirth4;
            this._DependentPeopleBirth5 = eDependentPeopleBirth5;
            this._DependentPeopleBirth6 = eDependentPeopleBirth6;
            this._Living1 = eLiving1;
            this._Living2 = eLiving2;
            this._Living3 = eLiving3;
            this._Living4 = eLiving4;
            this._Living5 = eLiving5;
            this._Living6 = eLiving6;
            this._RomajiName = eRomajiName;
            this._DependentPeople = eDependentPeople;
            this._ResidentPeople = eResidentPeople;
            this._HealthInsurancePeople = eHealthInsurancePeople;
        }
    }
}
