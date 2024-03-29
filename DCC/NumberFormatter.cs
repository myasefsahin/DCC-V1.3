
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DCC;
namespace DCC
{
    
    public class NumberFormatter
    {
        

        public static CalculateEntity deneme(CalculateEntity calculate)
        {
            

            (decimal formatli_belirsizlik, int basamak_sayisi) = IlkAdim(calculate.measurent, calculate.uncertainty);

            if (formatli_belirsizlik != 0)
            {
                calculate.uncertainty = formatli_belirsizlik;
                string formatli_olcum = IkinciAdim(calculate.measurent, basamak_sayisi, calculate.uncertainty);
                calculate.measurent = Convert.ToDecimal(formatli_olcum);

                return calculate;
            }
            else
            {
                return calculate;
            }
        }


        static (decimal, int) IlkAdim(decimal measurent, decimal uncertainty)
        {
            if (uncertainty == 0)
            {
                XmlDocument temp11 = new XmlDocument();
                CertificateForm hatarefresh = new CertificateForm(temp11);
                string errorMessage = "Hata oluştu!\n\n";
                errorMessage += "Belirsizlik değeri sıfır olamaz.\n";
                errorMessage += "Lütfen Excel dosyanızı kontrol ediniz.";

                hatarefresh.refresh();

                MessageBox.Show(errorMessage, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            int basamakSayisi = 0;

            if (uncertainty > 0 && uncertainty < 1)
            {
                
                // İkinci sayıyı virgülden sonraki ilk sıfırdan büyük basamağa kadar yuvarla

                decimal temp = uncertainty;

                while (Math.Floor(temp) == 0)
                {
                    temp *= 10;
                    basamakSayisi++;
                }


                // Bir sonraki basamak
                basamakSayisi++;

                // Virgülden sonraki ilk sıfırdan büyük basamağa kadar yuvarla

                decimal uncertainty_formatli = Math.Round(uncertainty, basamakSayisi, MidpointRounding.ToPositiveInfinity);
                string temp1 = uncertainty_formatli.ToString();
                char sonIndex = temp1[temp1.Length - 1];
                char sonOncekiIndex = temp1[temp1.Length - 2];


                if (sonIndex != '0' && sonOncekiIndex == '0')
                {
                    temp1 = temp1 + "0";
                    Decimal formatliTemp1 = Convert.ToDecimal(temp1);
                    
                    formatliTemp1 = Math.Round(formatliTemp1, basamakSayisi+1, MidpointRounding.ToPositiveInfinity);
                    uncertainty_formatli = formatliTemp1;

                }

                if (sonIndex != '0' && (sonOncekiIndex == '.' || sonOncekiIndex == ','))
                {
                    temp1 = temp1 + "0";
                    Decimal formatliTemp1 = Convert.ToDecimal(temp1);
                    
                    formatliTemp1 = Math.Round(formatliTemp1, basamakSayisi+1, MidpointRounding.ToPositiveInfinity);
                    uncertainty_formatli = formatliTemp1;

                }

                return (uncertainty_formatli, basamakSayisi);
            }
            else if (uncertainty >= 10)
            {   
                decimal uncertainty_formatli = Math.Round(uncertainty, MidpointRounding.ToPositiveInfinity);
                Math.Round(measurent, MidpointRounding.ToPositiveInfinity);
                return (uncertainty_formatli, 0); // Virgül sonrası basamak sayısını 0 olarak döndür
            }
            else if (uncertainty > 1 && uncertainty < 10)
            {

                decimal uncertainty_formatli = (decimal)Math.Round(uncertainty, 1, MidpointRounding.ToPositiveInfinity);


                return (uncertainty_formatli, 1);


            }
            else
            {
                return (0, 0); // İkinci sayı belirtilen aralıkta değil
            }
        }


        static string IkinciAdim(decimal measurent, int virgul_sonrasi_basamak_sayisi, decimal belirsizlik)
        {
           

            if (belirsizlik > 1 && belirsizlik < 10)
            {
                
                measurent = (decimal)Math.Round(measurent, 1, MidpointRounding.ToPositiveInfinity);
                string formatli_measurent1 = measurent.ToString();
                return formatli_measurent1;

            }
            // Birinci sayıyı belirtilen formata göre düzenle
            string formatli_measurent = measurent.ToString("0." + new string('0', virgul_sonrasi_basamak_sayisi));
            return formatli_measurent;



        }
    }

}

