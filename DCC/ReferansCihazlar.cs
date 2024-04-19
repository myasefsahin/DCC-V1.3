using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    public class ReferansCihazlar
    {
        public int ReferansCihazId { get; set; }
        public string SiparisNo { get; set; }
        public string MakineAdi { get; set; } //Cihaz
        public string Imalatci { get; set; } //Üretici Adı
        public string Tip { get; set; } //Model
        public string SeriNumarasi { get; set; } //Seri No
        public string Izlenebilirlik { get; set; }
    }

}
