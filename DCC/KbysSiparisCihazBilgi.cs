using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    public class KbysSiparisCihazBilgi
    {
        [Key]
        public int KbysSiparisCihazBilgiId { get; set; }
        public string SiparisNo { get; set; }
        public int? SprChzBilgiSirano { get; set; }
        public string CihazNo { get; set; }
        public string SiparisDurumKodu { get; set; }
        public int? SertifikaBirimSirano { get; set; }
        public int? SertifikaEskiBirimSirano { get; set; }
        public string SertifikaYili { get; set; }
        public int? SertifikaNo { get; set; }
        public string SipTipKodu { get; set; }
        public string CihazSorumlusu { get; set; }
        public string SorumluUnvan { get; set; }
        public string SorumluTel { get; set; }
        public string SorumluFax { get; set; }
        public string OlcumAraligi { get; set; }
        public DateTime? FirmaGonderTrh { get; set; }
        public DateTime? GerceklesenGirisTrh { get; set; }
        public DateTime? GerceklesenCikisTrh { get; set; }
        public string TeknikBilgi { get; set; }
        public string Gelis { get; set; }
        public string Gidis { get; set; }
        public DateTime? FaturaTrh { get; set; }
        public string FaturaNo { get; set; }
        public DateTime? SertifikaCikisTrh { get; set; }
        public double? GerceklesenFiyatiUsd { get; set; }
        public double? GerceklesenFiyatiTl { get; set; }
        public string KalYeri { get; set; }
        public string DeneyTanimi { get; set; }
        public int? GirisYapanSclno { get; set; }
        public DateTime? GirisYapanTrhSaat { get; set; }
        public string MailGonder { get; set; }
        public string UmeGelis { get; set; }
        public string UmeGidis { get; set; }
        public DateTime? EtiketBasimTrhSaat { get; set; }
        public int? EtiketBasanSclno { get; set; }
        public string SorumluEmail { get; set; }
        public string FirmaMailGonderDurum { get; set; }
        public string AkrediteSertDurum { get; set; }
        public string Masraf { get; set; }
        public string CihazLabTeslimDurum { get; set; }
        public DateTime? FatOdemeTrh { get; set; }
        public DateTime? FatGonderimTrh { get; set; }
        public DateTime? YrdLabKalBaslangicTrh { get; set; }
        public DateTime? YrdLabKalBitisTrh { get; set; }
        public int? SertHazirlamaSuresi { get; set; }
        public string FatGonderBarkodNo { get; set; }
        public string NedenSirano { get; set; }
        public DateTime? CihazLabTeslimTrh { get; set; }
        public DateTime? CihazDepoTeslimTrh { get; set; }
        public string SertGidisSekli { get; set; }
        public string DeneyTanimiEng { get; set; }
        public string FirmaOlcumAraligi { get; set; }
        public string FirmaTeknikBilgi { get; set; }
        public int? EtiketSayisi { get; set; }
        public string ParaCinsi { get; set; }

        public virtual KbysFirmaCihaz? CihazNoNavigation { get; set; }

        public virtual ICollection<KbysSiparisAyrinti> KbysSiparisAyrintis { get; set; } = new List<KbysSiparisAyrinti>();
    }

}
