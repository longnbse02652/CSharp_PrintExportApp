using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class DTO_AllInfor
    {
        private string _idCode;
        public string idCode    {get { return _idCode; }set { _idCode = value; }}

        private string _romaji;
        public string romaji    {get { return _romaji; }set { _romaji = value; }}

        private string _furigana;
        public string furigana  {get { return _furigana; }set { _furigana = value; }}

        private string _sex;
        public string sex   {get { return _sex; }set { _sex = value; }}

        private int _age;
        public int age   {get { return _age; }set { _age = value; }}

        private string _birth;
        public string birth {get { return _birth; }set { _birth = value; }}

        private string _nationality;
        public string nationality { get { return _nationality; } set { _nationality = value; } }

        private string _inCompanyDate;
        public string inCompanyDate { get { return _inCompanyDate; } set { _inCompanyDate = value; } }

        private string _cardType;
        public string cardType { get { return _cardType; } set { _cardType = value; } }

        private string _cardTime;
        public string cardTime { get { return _cardTime; } set { _cardTime = value; } }

        private string _cardTimeOut;
        public string cardTimeOut { get { return _cardTimeOut; } set { _cardTimeOut = value; } }

        private string _outTime;
        public string outTime { get { return _outTime; } set { _outTime = value; } }

        private string _companyCode;
        public string companyCode { get { return _companyCode; } set { _companyCode = value; } }

        private string _companyName;
        public string companyName { get { return _companyName; } set { _companyName = value; } }

        private string _workType;
        public string workType { get { return _workType; } set { _workType = value; } }

        private string _closingDate;
        public string closingDate { get { return _closingDate; } set { _closingDate = value; } }

        private int _zipCode;
        public int zipCode { get { return _zipCode; } set { _zipCode = value; } }

        private string _address;
        public string address { get { return _address; } set { _address = value; } }

        private string _mobliePhone;
        public string mobliePhone { get { return _mobliePhone; } set { _mobliePhone = value; } }

        private string _phone;
        public string phone { get { return _phone; } set { _phone = value; } }

        private string _createPeople;
        public string createPeople { get { return _createPeople; } set { _createPeople = value; } }

        private string _position;
        public string position { get { return _position; } set { _position = value; } }



        public DTO_AllInfor(string pIdcode,string pRomaji, string pFurigana, string pSex, int pAge,string pBirth, string pNationality,
            string pInCompanyDate, string pCardType, string pCardTime, string pCardTimeOut, string pOutTime, string pCompanyCode,
            string pCompanyName, string pWorkType, string pClosingDate, int pZipCode,string pAddress,string pMobliePhone,
            string pPhone, string pCreatePeople, string pPosition)
        {
            this._idCode = pIdcode;
            this._romaji = pRomaji;
            this._furigana = pFurigana;
            this._sex = pSex;
            this._age = pAge;
            this._birth = pBirth;
            this._nationality = pNationality;
            this._inCompanyDate = pInCompanyDate;
            this._cardType = pCardType;
            this._cardTime = pCardTime;
            this._cardTimeOut = pCardTimeOut;
            this._outTime = pOutTime;
            this._companyCode = pCompanyCode;
            this._companyName = pCompanyName;
            this._workType = pWorkType;
            this._closingDate = pClosingDate;
            this._zipCode = pZipCode;
            this._address = pAddress;
            this._mobliePhone = pMobliePhone;
            this._phone = pPhone;
            this._createPeople = pCreatePeople;
            this._position = pPosition;
        }
    }
}
