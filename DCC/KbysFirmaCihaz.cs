using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    public partial class KbysFirmaCihaz
    {
        public int FirmaCihazId { get; set; }
        public string CihazNo { get; set; }

        public string CihazAdi { get; set; } = null!;

        public string? CihazAdiEng { get; set; }

        public byte? CihazGrubu { get; set; }

        public byte? UreticiKodu { get; set; }

        public string? Model { get; set; }

        public string? SeriNo { get; set; }

        public byte? KalPeriyodu { get; set; }
    }

}
