using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    public class AdministrativeData
    {
        public AdministrativeData()
        {
            SiparisCihazlari = new List<KbysSiparisCihazBilgi>();
        }

        [Key]
        public int AdministrativeDataId { get; set; }
        public decimal BirimSiraNo { get; set; }
        public string SertifikaYili { get; set; }
        public string? Birimkodu { get; set; }
        public int SertifikaNo { get; set; }
        public string FirmaAdi { get; set; }
        public string? Adres { get; set; }
        public string SiparisNo { get; set; }
        public string CihazAdi { get; set; }
        public string? CihazAdiEng { get; set; }
        public string? UreticiAdi { get; set; }
        public int? UreticiKodu { get; set; }
        public string? CihazModel { get; set; }
        public string? CihazSeriNo { get; set; }
        public string CihazKalibrasyonTrh { get; set; }

        public List<KbysSiparisCihazBilgi> SiparisCihazlari { get; set; }


    }
}
