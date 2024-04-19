using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    public class ApiData
    {
        public string SiparisNo { get; set; }
        public string FirmaAdi { get; set; }
        public string FirmaAdres { get; set; }
        public string FirmaPostaKodu { get; set; }
        public string FirmaIl { get; set; }
        public string FirmaIlce { get; set; }
        public string MakineAdi { get; set; } //Cihaz
        public string MakineAdiEng { get; set; } //Cihaz adı ingilizce
        public string Imalatci { get; set; } //Üretici Adı
        public string Tip { get; set; } //Model
        public string SeriNumarasi { get; set; } //Seri No
        public DateTime KalibrasyonBaslangicTarihi { get; set; } //KAL_BAS_TRH
        public DateTime KalibrasyonBitisTarihi { get; set; } //KAL_BIT_TRH
        public DateTime DosyaBaskiTarihi { get; set; } //DOSYA_BASKI_TRH_SAAT
        public string KalibrasyonunYapildigiYer { get; set; } //KAL_YERI
        public DateTime CihazinLaboratuvaraKabulTarihi { get; set; } //KABUL_TRH
        public List<ReferansCihazlar> ReferansCihazlars { get; set; }
        public List<KalibrasyonuYapanlar> KalibrasyonuYapanlars { get; set; }
        public string LaboratuvarSorumlusu { get; set; }
    }
}
