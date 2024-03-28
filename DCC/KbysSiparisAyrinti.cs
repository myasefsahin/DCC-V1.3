using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    public partial class KbysSiparisAyrinti
    {
        public int KbysSiparisAyrintiId { get; set; }
        public decimal SprAyrintiSirano { get; set; }

        public string SiparisNo { get; set; } = null!;

        public decimal SprChzBilgiSirano { get; set; }

        public decimal BirimSirano { get; set; }

        public decimal EskiBirimSirano { get; set; }

        public string KalKodu { get; set; } = null!;

        public string? KalDurumKodu { get; set; }
    }

}
