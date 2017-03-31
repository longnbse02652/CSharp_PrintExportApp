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

        private string _romaji = string.Empty;
        public string romaji    {get { return _romaji; }set { _romaji = value; }}

        private string _furigana = string.Empty;
        public string furigana  {get { return _furigana; }set { _furigana = value; }}

        private string _sex;
        public string sex   {get { return _sex; }set { _sex = value; }}

        private string _birth;
        public string birth {get { return _birth; }set { _birth = value; }}

        private string _nationality;
        public string nationality { get { return _nationality; } set { _nationality = value; } }

        private string _inCompanyDate;
        public string inCompanyDate { get { return _inCompanyDate; } set { _inCompanyDate = value; } }

        private string _cardType;
        public string cardType { get { return _cardType; } set { _cardType = value; } }

        private string _cardTimeStart;
        public string cardTimeStart { get { return _cardTimeStart; } set { _cardTimeStart = value; } }

        private string _cardTimeOver;
        public string cardTimeOver { get { return _cardTimeOver; } set { _cardTimeOver = value; } }

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

        private string _hakenRyokin;
        public string hakenRyokin { get { return _hakenRyokin; } set { _hakenRyokin = value; } }

        private string _hakenRyokinType;
        public string hakenRyokinType { get { return _hakenRyokinType; } set { _hakenRyokinType = value; } }

        private string _shiharaiType;
        public string shiharaiType { get { return _shiharaiType; } set { _shiharaiType = value; } }

        private string _tax;
        public string tax { get { return _tax; } set { _tax = value; } }

        private string _salaryType;
        public string salaryType { get { return _salaryType; } set { _salaryType = value; } }

        private int _basicSalary;
        public int basicSalary { get { return _basicSalary; } set { _basicSalary = value; } }

        private int _seikinTeate;
        public int seikinTeate { get { return _seikinTeate; } set { _seikinTeate = value; } }

        private int _gaikinTeate;
        public int gaikinTeate { get { return _gaikinTeate; } set { _gaikinTeate = value; } }

        private int _gijutsuTeate;
        public int gijutsuTeate { get { return _gijutsuTeate; } set { _gijutsuTeate = value; } }

        private int _shikakuTeate;
        public int shikakuTeate { get { return _shikakuTeate; } set { _shikakuTeate = value; } }

        private int _yakushokuTeate;
        public int yakushokuTeate { get { return _yakushokuTeate; } set { _yakushokuTeate = value; } }

        private int _eigyoTeate;
        public int eigyoTeate { get { return _eigyoTeate; } set { _eigyoTeate = value; } }

        private int _kazokuTeate;
        public int kazokuTeate { get { return _kazokuTeate; } set { _kazokuTeate = value; } }

        private int _jutakuTeate;
        public int jutakuTeate { get { return _jutakuTeate; } set { _jutakuTeate = value; } }

        private int _bekkyoTeate;
        public int bekkyoTeate { get { return _bekkyoTeate; } set { _bekkyoTeate = value; } }

        private int _tsukinTeate;
        public int tsukinTeate { get { return _tsukinTeate; } set { _tsukinTeate = value; } }

        private int _park;
        public int park { get { return _park; } set { _park = value; } }

        private int _dormitoryFee;
        public int dormitoryFee { get { return _dormitoryFee; } set { _dormitoryFee = value; } }

        private int _waterFee;
        public int waterFee { get { return _waterFee; } set { _waterFee = value; } }

        private string _employStatus;
        public string employStatus { get { return _employStatus; } set { _employStatus = value; } }

        private string _employTime1;
        public string employTime1 { get { return _employTime1; } set { _employTime1 = value; } }

        private string _employTime2;
        public string employTime2 { get { return _employTime2; } set { _employTime2 = value; } }

        private string _bankName;
        public string bankName { get { return _bankName; } set { _bankName = value; } }

        private string _bankNameType;
        public string bankNameType { get { return _bankNameType; } set { _bankNameType = value; } }

        private string _branchName;
        public string branchName { get { return _branchName; } set { _branchName = value; } }

        private string _branchNameType;
        public string branchNameType { get { return _branchNameType; } set { _branchNameType = value; } }

        private string _accountName;
        public string accountName { get { return _accountName; } set { _accountName = value; } }

        private string _bankCode;
        public string bankCode { get { return _bankCode; } set { _bankCode = value; } }

        private string _branchCode;
        public string branchCode { get { return _branchCode; } set { _branchCode = value; } }

        private string _accountCode1;
        public string accountCode1 { get { return _accountCode1; } set { _accountCode1 = value; } }

        private string _accountCode2;
        public string accountCode2 { get { return _accountCode2; } set { _accountCode2 = value; } }

        private string _accountCode3;
        public string accountCode3 { get { return _accountCode3; } set { _accountCode3 = value; } }

        private string _accountCode4;
        public string accountCode4 { get { return _accountCode4; } set { _accountCode4 = value; } }

        private string _accountCode5;
        public string accountCode5 { get { return _accountCode5; } set { _accountCode5 = value; } }

        private string _accountCode6;
        public string accountCode6 { get { return _accountCode6; } set { _accountCode6 = value; } }

        private string _accountCode7;
        public string accountCode7 { get { return _accountCode7; } set { _accountCode7 = value; } }

        private string _accountCode8;
        public string accountCode8 { get { return _accountCode8; } set { _accountCode8 = value; } }

        private string _travelType;
        public string travelType { get { return _travelType; } set { _travelType = value; } }

        private string _houseName;
        public string houseName { get { return _houseName; } set { _houseName = value; } }

        private string _room;
        public string room { get { return _room; } set { _room = value; } }

        private string _inHouseDate;
        public string inHouseDate { get { return _inHouseDate; } set { _inHouseDate = value; } }

        private string _kouyouhoken;
        public string kouyouhoken { get { return _kouyouhoken; } set { _kouyouhoken = value; } }

        private string _shakaihoken;
        public string shakaihoken { get { return _shakaihoken; } set { _shakaihoken = value; } }

        private int _dependentPeople;
        public int dependentPeople { get { return _dependentPeople; } set { _dependentPeople = value; } }

        private int _residentPeople;
        public int residentPeople { get { return _residentPeople; } set { _residentPeople = value; } }

        private int _healthInsurancePeople;
        public int healthInsurancePeople { get { return _healthInsurancePeople; } set { _healthInsurancePeople = value; } }

        private string _contractType;
        public string contractType { get { return _contractType; } set { _contractType = value; } }

        private string _contractRequire;
        public string contractRequire { get { return _contractRequire; } set { _contractRequire = value; } }

        private string _myCompany;
        public string myCompany { get { return _myCompany; } set { _myCompany = value; } }

        private string _workContent;
        public string workContent { get { return _workContent; } set { _workContent = value; } }

        private int _workTime1;
        public int workTime1 { get { return _workTime1; } set { _workTime1 = value; } }

        private int _workTime2;
        public int workTime2 { get { return _workTime2; } set { _workTime2 = value; } }

        private int _workTime3;
        public int workTime3 { get { return _workTime3; } set { _workTime3 = value; } }

        private int _workTime4;
        public int workTime4 { get { return _workTime4; } set { _workTime4 = value; } }

        private int _relaxTime;
        public int relaxTime { get { return _relaxTime; } set { _relaxTime = value; } }

        private string _insureCard;
        public string insureCard { get { return _insureCard; } set { _insureCard = value; } }

        private string _pastCompany1;
        public string pastCompany1 { get { return _pastCompany1; } set { _pastCompany1 = value; } }

        private string _nienhieu1;
        public string nienhieu1 { get { return _nienhieu1; } set { _nienhieu1 = value; } }

        private int _beginYear1;
        public int beginYear1 { get { return _beginYear1; } set { _beginYear1 = value; } }

        private int _beginMonth1;
        public int beginMonth1 { get { return _beginMonth1; } set { _beginMonth1 = value; } }

        private int _endYear1;
        public int endYear1 { get { return _endYear1; } set { _endYear1 = value; } }

        private int _endMonth1;
        public int endMonth1 { get { return _endMonth1; } set { _endMonth1 = value; } }

        private string _pastCompany2;
        public string pastCompany2 { get { return _pastCompany2; } set { _pastCompany2 = value; } }

        private string _nienhieu2;
        public string nienhieu2 { get { return _nienhieu2; } set { _nienhieu2 = value; } }

        private int _beginYear2;
        public int beginYear2 { get { return _beginYear2; } set { _beginYear2 = value; } }

        private int _beginMonth2;
        public int beginMonth2 { get { return _beginMonth2; } set { _beginMonth2 = value; } }

        private int _endYear2;
        public int endYear2 { get { return _endYear2; } set { _endYear2 = value; } }

        private int _endMonth2;
        public int endMonth2 { get { return _endMonth2; } set { _endMonth2 = value; } }

        private string _pensionBook;
        public string pensionBook { get { return _pensionBook; } set { _pensionBook = value; } }

        private string _dependentPeopleKana1;
        public string dependentPeopleKana1 { get { return _dependentPeopleKana1; } set { _dependentPeopleKana1 = value; } }

        private string _dependentPeopleShimei1;
        public string dependentPeopleShimei1 { get { return _dependentPeopleShimei1; } set { _dependentPeopleShimei1 = value; } }

        private string _dependentPeopleBirth1;
        public string dependentPeopleBirth1 { get { return _dependentPeopleBirth1; } set { _dependentPeopleBirth1 = value; } }

        private string _relationship1;
        public string relationship1 { get { return _relationship1; } set { _relationship1 = value; } }

        private string _living1;
        public string living1 { get { return _living1; } set { _living1 = value; } }

        private string _dependentPeopleKana2;
        public string dependentPeopleKana2 { get { return _dependentPeopleKana2; } set { _dependentPeopleKana2 = value; } }

        private string _dependentPeopleShimei2;
        public string dependentPeopleShimei2 { get { return _dependentPeopleShimei2; } set { _dependentPeopleShimei2 = value; } }

        private string _dependentPeopleBirth2;
        public string dependentPeopleBirth2 { get { return _dependentPeopleBirth2; } set { _dependentPeopleBirth2 = value; } }

        private string _relationship2;
        public string relationship2 { get { return _relationship2; } set { _relationship2 = value; } }

        private string _living2;
        public string living2 { get { return _living2; } set { _living2 = value; } }

        private string _dependentPeopleKana3;
        public string dependentPeopleKana3 { get { return _dependentPeopleKana3; } set { _dependentPeopleKana3 = value; } }

        private string _dependentPeopleShimei3;
        public string dependentPeopleShimei3 { get { return _dependentPeopleShimei3; } set { _dependentPeopleShimei3 = value; } }

        private string _dependentPeopleBirth3;
        public string dependentPeopleBirth3 { get { return _dependentPeopleBirth3; } set { _dependentPeopleBirth3 = value; } }

        private string _relationship3;
        public string relationship3 { get { return _relationship3; } set { _relationship3 = value; } }

        private string _living3;
        public string living3 { get { return _living3; } set { _living3 = value; } }

        private string _dependentPeopleKana4;
        public string dependentPeopleKana4 { get { return _dependentPeopleKana4; } set { _dependentPeopleKana4 = value; } }

        private string _dependentPeopleShimei4;
        public string dependentPeopleShimei4 { get { return _dependentPeopleShimei4; } set { _dependentPeopleShimei4 = value; } }

        private string _dependentPeopleBirth4;
        public string dependentPeopleBirth4 { get { return _dependentPeopleBirth4; } set { _dependentPeopleBirth4 = value; } }

        private string _relationship4;
        public string relationship4 { get { return _relationship4; } set { _relationship4 = value; } }

        private string _living4;
        public string living4 { get { return _living4; } set { _living4 = value; } }

        private string _dependentPeopleKana5;
        public string dependentPeopleKana5 { get { return _dependentPeopleKana5; } set { _dependentPeopleKana5 = value; } }

        private string _dependentPeopleShimei5;
        public string dependentPeopleShimei5 { get { return _dependentPeopleShimei5; } set { _dependentPeopleShimei5 = value; } }

        private string _dependentPeopleBirth5;
        public string dependentPeopleBirth5 { get { return _dependentPeopleBirth5; } set { _dependentPeopleBirth5 = value; } }

        private string _relationship5;
        public string relationship5 { get { return _relationship5; } set { _relationship5 = value; } }

        private string _living5;
        public string living5 { get { return _living5; } set { _living5 = value; } }

        private string _dependentPeopleKana6;
        public string dependentPeopleKana6 { get { return _dependentPeopleKana6; } set { _dependentPeopleKana6 = value; } }

        private string _dependentPeopleShimei6;
        public string dependentPeopleShimei6 { get { return _dependentPeopleShimei6; } set { _dependentPeopleShimei6 = value; } }

        private string _dependentPeopleBirth6;
        public string dependentPeopleBirth6 { get { return _dependentPeopleBirth6; } set { _dependentPeopleBirth6 = value; } }

        private string _relationship6;
        public string relationship6 { get { return _relationship6; } set { _relationship6 = value; } }

        private string _living6;
        public string living6 { get { return _living6; } set { _living6 = value; } }

        private string _trainsportation1;
        public string trainsportation1 { get { return _trainsportation1; } set { _trainsportation1 = value; } }

        private string _beginTrain1;
        public string beginTrain1 { get { return _beginTrain1; } set { _beginTrain1 = value; } }

        private string _endTrain1;
        public string endTrain1 { get { return _endTrain1; } set { _endTrain1 = value; } }

        private int _monthRegular1;
        public int monthRegular1 { get { return _monthRegular1; } set { _monthRegular1 = value; } }

        private string _trainsportation2;
        public string trainsportation2 { get { return _trainsportation2; } set { _trainsportation2 = value; } }

        private string _beginTrain2;
        public string beginTrain2 { get { return _beginTrain2; } set { _beginTrain2 = value; } }

        private string _endTrain2;
        public string endTrain2 { get { return _endTrain2; } set { _endTrain2 = value; } }

        private int _monthRegular2;
        public int monthRegular2 { get { return _monthRegular2; } set { _monthRegular2 = value; } }

        private string _trainsportation3;
        public string trainsportation3 { get { return _trainsportation3; } set { _trainsportation3 = value; } }

        private string _beginTrain3;
        public string beginTrain3 { get { return _beginTrain3; } set { _beginTrain3 = value; } }

        private string _endTrain3;
        public string endTrain3 { get { return _endTrain3; } set { _endTrain3 = value; } }

        private int _monthRegular3;
        public int monthRegular3 { get { return _monthRegular3; } set { _monthRegular3 = value; } }

        private string _trainsportation4;
        public string trainsportation4 { get { return _trainsportation4; } set { _trainsportation4 = value; } }

        private string _beginTrain4;
        public string beginTrain4 { get { return _beginTrain4; } set { _beginTrain4 = value; } }

        private string _endTrain4;
        public string endTrain4 { get { return _endTrain4; } set { _endTrain4 = value; } }

        private int _monthRegular4;
        public int monthRegular4 { get { return _monthRegular4; } set { _monthRegular4 = value; } }

        private string _carkm;
        public string carkm { get { return _carkm; } set { _carkm = value; } }

        private int _carMoney;
        public int carMoney { get { return _carMoney; } set { _carMoney = value; } }

        private int _totalMoneyTrans;
        public int totalMoneyTrans { get { return _totalMoneyTrans; } set { _totalMoneyTrans = value; } }

        private string _reason;
        public string reason { get { return _reason; } set { _reason = value; } }

        private string _changeDateFrom;
        public string changeDateFrom { get { return _changeDateFrom; } set { _changeDateFrom = value; } }

        private string _changeDate;
        public string changeDate { get { return _changeDate; } set { _changeDate = value; } }

        private double _genkaritsu;
        public double genkaritsu { get { return _genkaritsu; } set { _genkaritsu = value; } }

        private int _teateGaku;
        public int teateGaku { get { return _teateGaku; } set { _teateGaku = value; } }

        private string _accountCode;
        public string accountCode { get { return _accountCode; } set { _accountCode = value; } }

        private int _chingin;
        public int chingin { get { return _chingin; } set { _chingin = value; } }

        private string _chinginType;
        public string chinginType { get { return _chinginType; } set { _chinginType = value; } }

        private int _kyuyoKojoGaku;
        public int kyuyoKojoGaku { get { return _kyuyoKojoGaku; } set { _kyuyoKojoGaku = value; } }

        private int _workTime;
        public int workTime { get { return _workTime; } set { _workTime = value; } }

        public DTO_AllInfor(string pIdcode,string pRomaji, string pFurigana, string pSex,string pBirth, string pNationality,
            string pInCompanyDate, string pCardType, string pCardTimeStart, string pCardTimeOver, string pOutTime, string pCompanyCode,
            string pCompanyName, string pWorkType, string pClosingDate, int pZipCode,string pAddress,string pMobliePhone,
            string pPhone, string pCreatePeople, string pPosition, string HakenRyokin, string HakenRyokinType, string ShiharaiType
            , string Tax, string SalaryType, int BasicSalary, int SeikinTeate, int GaikinTeate, int GijutsuTeate
            , int ShikakuTeate, int YakushokuTeate, int EigyoTeate, int KazokuTeate, int JutakuTeate, int BekkyoTeate
            , int TsukinTeate, int Park, int DormitoryFee, int WaterFee, string EmployStatus, string EmployTime1
            , string EmployTime2, string BankName, string BankNameType, string BranchName, string BranchNameType, string AccountName
            , string BankCode, string BranchCode, string AccountCode1, string AccountCode2, string AccountCode3, string AccountCode4
            , string AccountCode5, string AccountCode6, string AccountCode7, string AccountCode8, string TravelType, string HouseName
            , string Room, string InHouseDate, string Kouyouhoken, string Shakaihoken, int DependentPeople, int ResidentPeople
            , int HealthInsurancePeople, string ContractType, string ContractRequire, string MyCompany, string WorkContent, int WorkTime1
            , int WorkTime2, int WorkTime3, int WorkTime4, int RelaxTime, string InsureCard, string PastCompany1, string Nienhieu1
            , int BeginYear1, int BeginMonth1, int EndYear1, int EndMonth1, string PastCompany2, string Nienhieu2, int BeginYear2
            , int BeginMonth2, int EndYear2, int EndMonth2, string PensionBook, string DependentPeopleKana1, string DependentPeopleShimei1
            , string DependentPeopleBirth1, string Relationship1, string Living1,string DependentPeopleKana2, string DependentPeopleShimei2
            , string DependentPeopleBirth2, string Relationship2, string Living2,string DependentPeopleKana3, string DependentPeopleShimei3
            , string DependentPeopleBirth3, string Relationship3, string Living3,string DependentPeopleKana4, string DependentPeopleShimei4
            , string DependentPeopleBirth4, string Relationship4, string Living4,string DependentPeopleKana5, string DependentPeopleShimei5
            , string DependentPeopleBirth5, string Relationship5, string Living5,string DependentPeopleKana6, string DependentPeopleShimei6
            , string DependentPeopleBirth6, string Relationship6, string Living6, string Trainsportation1, string BeginTrain1, string EndTrain1
            , int MonthRegular1, string Trainsportation2, string BeginTrain2, string EndTrain2, int MonthRegular2
            , string Trainsportation3, string BeginTrain3, string EndTrain3, int MonthRegular3
            , string Trainsportation4, string BeginTrain4, string EndTrain4, int MonthRegular4, string Carkm, int CarMoney, int TotalMoneyTrans
            , string Reason, string ChangeDateFrom, string ChangeDate, double Genkaritsu, int TeateGaku, string AccountCode
            , int Chingin, string ChinginType, int KyuyoKojoGaku, int WorkTime)
        {
            this._idCode = pIdcode;
            this._romaji = pRomaji;
            this._furigana = pFurigana;
            this._sex = pSex;
            this._birth = pBirth;
            this._nationality = pNationality;
            this._inCompanyDate = pInCompanyDate;
            this._cardType = pCardType;
            this._cardTimeStart = pCardTimeStart;
            this._cardTimeOver = pCardTimeOver;
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
            this._hakenRyokin = HakenRyokin;
            this._hakenRyokinType = HakenRyokinType;
            this._shiharaiType = ShiharaiType;
            this._tax = Tax;
            this._salaryType = SalaryType;
            this._basicSalary = BasicSalary;
            this._seikinTeate = SeikinTeate;
            this._gaikinTeate = GaikinTeate;
            this._gijutsuTeate = GijutsuTeate;
            this._shikakuTeate = ShikakuTeate;
            this._yakushokuTeate = YakushokuTeate;
            this._eigyoTeate = EigyoTeate;
            this._kazokuTeate = KazokuTeate;
            this._jutakuTeate = JutakuTeate;
            this._bekkyoTeate = BekkyoTeate;
            this._tsukinTeate = TsukinTeate;
            this._park = Park;
            this._dormitoryFee = DormitoryFee;
            this._waterFee = WaterFee;
            this._employStatus = EmployStatus;
            this._employTime1 = EmployTime1;
            this._employTime2 = EmployTime2;
            this._bankName = BankName;
            this._bankNameType = BankNameType;
            this._branchName = BranchName;
            this._branchNameType = BranchNameType;
            this._accountName = AccountName;
            this._bankCode = BankCode;
            this._branchCode = BranchCode;
            this._accountCode1 = AccountCode1;
            this._accountCode2 = AccountCode2;
            this._accountCode3 = AccountCode3;
            this._accountCode4 = AccountCode4;
            this._accountCode5 = AccountCode5;
            this._accountCode6 = AccountCode6;
            this._accountCode7 = AccountCode7;
            this._accountCode8 = AccountCode8;
            this._travelType = TravelType;
            this._houseName = HouseName;
            this._room = Room;
            this._inHouseDate = InHouseDate;
            this._kouyouhoken = Kouyouhoken;
            this._shakaihoken = Shakaihoken;
            this._dependentPeople = DependentPeople;
            this._residentPeople = ResidentPeople;
            this._healthInsurancePeople = HealthInsurancePeople;
            this._contractType = ContractType;
            this._contractRequire = ContractRequire;
            this._myCompany = MyCompany;
            this._workContent = WorkContent;
            this._workTime1 = WorkTime1;
            this._workTime2 = WorkTime2;
            this._workTime3 = WorkTime3;
            this._workTime4 = WorkTime4;
            this._relaxTime = RelaxTime;
            this._insureCard = InsureCard;
            this._pastCompany1 = PastCompany1;
            this._nienhieu1 = Nienhieu1;
            this._beginYear1 = BeginYear1;
            this._beginMonth1 = BeginMonth1;
            this._endYear1 = EndYear1;
            this._endMonth1 = EndMonth1;
            this._pastCompany2 = PastCompany2;
            this._nienhieu2 = Nienhieu2;
            this._beginYear2 = BeginYear2;
            this._beginMonth2 = BeginMonth2;
            this._endYear2 = EndYear2;
            this._endMonth2 = EndMonth2;
            this._pensionBook = PensionBook;
            this._dependentPeopleKana1 = DependentPeopleKana1;
            this._dependentPeopleShimei1 = DependentPeopleShimei1;
            this._dependentPeopleBirth1 = DependentPeopleBirth1;
            this._relationship1 = Relationship1;
            this._living1 = Living1;
            this._dependentPeopleKana2 = DependentPeopleKana2;
            this._dependentPeopleShimei2 = DependentPeopleShimei2;
            this._dependentPeopleBirth2 = DependentPeopleBirth2;
            this._relationship2 = Relationship2;
            this._living2 = Living2;
            this._dependentPeopleKana3 = DependentPeopleKana3;
            this._dependentPeopleShimei3 = DependentPeopleShimei3;
            this._dependentPeopleBirth3 = DependentPeopleBirth3;
            this._relationship3 = Relationship3;
            this._living3 = Living3;
            this._dependentPeopleKana4 = DependentPeopleKana4;
            this._dependentPeopleShimei4 = DependentPeopleShimei4;
            this._dependentPeopleBirth4 = DependentPeopleBirth4;
            this._relationship4 = Relationship4;
            this._living4 = Living4;
            this._dependentPeopleKana5 = DependentPeopleKana5;
            this._dependentPeopleShimei5 = DependentPeopleShimei5;
            this._dependentPeopleBirth5 = DependentPeopleBirth5;
            this._relationship5 = Relationship5;
            this._living5 = Living5;
            this._dependentPeopleKana6 = DependentPeopleKana6;
            this._dependentPeopleShimei6 = DependentPeopleShimei6;
            this._dependentPeopleBirth6 = DependentPeopleBirth6;
            this._relationship6 = Relationship6;
            this._living6 = Living6;
            this._trainsportation1 = Trainsportation1;
            this._beginTrain1 = BeginTrain1;
            this._endTrain1 = EndTrain1;
            this._monthRegular1 = MonthRegular1;
            this._trainsportation2 = Trainsportation2;
            this._beginTrain2 = BeginTrain2;
            this._endTrain2 = EndTrain2;
            this._monthRegular2 = MonthRegular2;
            this._trainsportation3 = Trainsportation3;
            this._beginTrain3 = BeginTrain3;
            this._endTrain3 = EndTrain3;
            this._monthRegular3 = MonthRegular3;
            this._trainsportation4 = Trainsportation4;
            this._beginTrain4 = BeginTrain4;
            this._endTrain4 = EndTrain4;
            this._monthRegular4 = MonthRegular4;
            this._carkm = Carkm;
            this._carMoney = CarMoney;
            this._totalMoneyTrans = TotalMoneyTrans;
            this._reason = Reason;
            this._changeDateFrom = ChangeDateFrom;
            this._changeDate = ChangeDate;
            this._genkaritsu = Genkaritsu;
            this._teateGaku = TeateGaku;
            this._accountCode = AccountCode;
            this._chingin = Chingin;
            this._chinginType = ChinginType;
            this._kyuyoKojoGaku = KyuyoKojoGaku;
            this._workTime = WorkTime;

        }
    }
}
