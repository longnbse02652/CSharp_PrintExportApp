using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class DTO_Infor
    {
        private string _furigana;
        public string furigana { get;set; }

        private string _romaji;
        public string romaji { get; set; }

        private string _birth;
        public string birth { get; set; }

        public DTO_Infor(string pFurigana,string pRomaji, string pBirth){
            this._furigana = pFurigana;
            this._romaji = pRomaji;
            this._birth = pBirth;
        }
    }
}
