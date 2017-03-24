using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class DTO_AllInfor
    {
        private string _furigana;
        public string furigana
        {
            get { return _furigana; }
            set { _furigana = value; }
        }

        private string _romaji;
        public string romaji
        {
            get { return _romaji; }
            set { _romaji = value; }
        }

        private string _birth;
        public string birth {
            get { return _birth; }
            set { _birth = value; }
        }

        public DTO_AllInfor(string pFurigana, string pRomaji, string pBirth)
        {
            this._furigana = pFurigana;
            this._romaji = pRomaji;
            this._birth = pBirth;
        }
    }
}
