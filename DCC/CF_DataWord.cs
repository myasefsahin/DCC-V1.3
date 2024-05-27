using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCC
{
    class CF_DataWord
    {
        // Excel sütun adlarının listesi
        //Sütun isimleri kullanıcının datagridview üzerinden seçilen hücrenin sütununun indexini bulmak için kullanılmıştır.
        //Burada kullanılan sütunların Excel üzerindeki tabloların fazlalığına göre arttırılmalıdır. 

        List<string> columnName = new List<string>(104) {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
            "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ","AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
            "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ","BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT","BU", "BV", "BW", "BX", "BY", "BZ",
            "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ","CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT","CU", "CV", "CW", "CX", "CY", "CZ"
        };

        // Ana metod

        //Kullanıcın Datagridview1 üzerinden seçtiği hücrenin satır ve sütun değerleri burada parametre olarak alınır ve bu değerlere göre veriler Excel sayfasından çekilir. 
        //Kullanıcının seçtiği Excel dosyasının yolu ve seçtiği Excel sayfasının ismi burada parametre olarak  getirilir ve işlemler bu veriler üzerinden devam eder.
        public void main(string ExcelDosyaYolu, string pageName, int satır, string sütun)
        {
            // Verilen sütun adının indeksini bul
            int harfIndex = columnName.IndexOf(sütun);
            // EPPlus kütüphanesi için lisanslama
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Excel dosyasını açma
            using (var package = new ExcelPackage(new FileInfo(ExcelDosyaYolu)))
            {
                // Belirtilen sayfayı seç
                ExcelWorksheet worksheet = package.Workbook.Worksheets[pageName];
                // Sayfadaki toplam satır sayısını al
                int rowCount = worksheet.Dimension.Rows;

                // Hücre değerlerini saklamak için dizi oluştur
                string[] cellValue = new string[rowCount];

                // Frekans değerlerini çekme
                for (int i = satır; i <= rowCount; i++)
                {
                    cellValue[i - satır] = Convert.ToString(worksheet.Cells[sütun + i].Value);
                    if (!string.IsNullOrEmpty(cellValue[i - satır]))
                    {
                        CF_ArrayFrekans.Add(cellValue[i - satır]);
                    }
                }

                // Calibration Factor Ölçüm değerlerinin Excel dosyasından çekimi ve formatlanması
                for (int i = satır; i < CF_ArrayFrekans.Count + satır; i++)
                {
                    NumberFormatter formatter = new NumberFormatter();
                    CalculateEntity calculateEntity = new CalculateEntity();

                    //Burada ölçülen ve belirsizlik değeri Excel sayfasındaki yerlerine göre belirsizlik NumberFormatter Classı kullanılarak İki anlamlı digit formnatına getirilir ve ölçülen değer de bu formata göre formatlanır.
                    //Her iki değer de Excel sayfasından çekildikten sonra buradan formatlanarak ilgili arraylistlerin içerisine aktarılır.
                    //Calibration Factor ölçüm tipinin ilk tablosu burada formatlanır. 
                    // Buradaki harfindex değişkeni kullanıcının seçtiği hücrenin sütunun yukarıda bulunan sütun adları listesindeki indexine karşılık gelir. 
                    //harfindex değişkeninin bulunduğu yer frenkansın bulunduğu sütuna denk gelir ve diğer sütunlar da bu sütun baz alınarak bir sonraki sütun +1 olacak şekilde ayarlanmıştır. 
                    //Tablolar arasındaki boşluklar ve frekansların tekrarladığı kısımlar boş gibi sayılarak indexler düzenlenmiştrir. 
                    
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 1] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 2] + i].Value);
                    CalculateEntity formattedEntity = NumberFormatter.deneme(calculateEntity);
                    CF_Array.Add(formattedEntity.measurent);
                    CF_ArrayCFUnc.Add(formattedEntity.uncertainty);

                    
                    //Calibration factor 2. tablosunun verilerinin çekilmesi 
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 5] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 6] + i].Value);
                    CalculateEntity formattedEntity1 = NumberFormatter.deneme(calculateEntity);
                    CF_ArrayReel.Add(formattedEntity1.measurent);
                    CF_ArrayReelUnc.Add(formattedEntity1.uncertainty);

                    
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 7] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 8] + i].Value);
                    CalculateEntity formattedEntity2 = NumberFormatter.deneme(calculateEntity);
                    CF_ArrayComplex.Add(formattedEntity2.measurent);
                    CF_ArrayComplexUnc.Add(formattedEntity2.uncertainty);

                    
                    calculateEntity.measurent = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 9] + i].Value);
                    calculateEntity.uncertainty = Convert.ToDecimal(worksheet.Cells[columnName[harfIndex + 10] + i].Value);
                    CalculateEntity formattedEntity3 = NumberFormatter.deneme(calculateEntity);
                    CF_YK.Add(formattedEntity3.measurent);
                    CF_YK_Unc.Add(formattedEntity3.uncertainty);
                }

                // Tablo adlarını çekme
                //Tablo adları her ölçüm tipi için oluşturulmuş Excel dosyalarının sertifika sayfalarında bulunan tabloların il sütununun en üst kısmına göre alınmıştır.
                //oluşturulan yeni Excel dosyalarında da bu özellik dikkate alınarak kod geliştirilmeli veya Excel dosyaları buna göre hazırlanmalıdır. 
                tableName1 = Convert.ToString(worksheet.Cells[columnName[harfIndex] + (satır - 3)].Value);
                tableName2 = Convert.ToString(worksheet.Cells[columnName[harfIndex + 4] + (satır - 3)].Value);
            }
        }

        // Değişken tanımlamaları
        public string tableName1;
        public string tableName2;
        public ArrayList CF_ArrayFrekans { get; set; }
        public ArrayList CF_Array { get; set; }
        public ArrayList CF_ArrayCFUnc { get; set; }
        public ArrayList CF_ArrayReel { get; set; }
        public ArrayList CF_ArrayReelUnc { get; set; }
        public ArrayList CF_ArrayComplex { get; set; }
        public ArrayList CF_ArrayComplexUnc { get; set; }
        public ArrayList CF_YK { get; set; }
        public ArrayList CF_YK_Unc { get; set; }

        // Cihaz bilgileri
        public string OrderNumber { get; set; }
        public string DeviceName { get; set; }
        public string SerialNumber { get; set; }

        // Yapıcı metod
        public CF_DataWord()
        {
            CF_ArrayFrekans = new ArrayList();
            CF_Array = new ArrayList();
            CF_ArrayCFUnc = new ArrayList();
            CF_ArrayReel = new ArrayList();
            CF_ArrayReelUnc = new ArrayList();
            CF_ArrayComplex = new ArrayList();
            CF_ArrayComplexUnc = new ArrayList();
            CF_YK = new ArrayList();
            CF_YK_Unc = new ArrayList();
        }

        // Verileri temizleme metodu
        public void ClearData()
        {
            CF_ArrayFrekans.Clear();
            CF_Array.Clear();
            CF_ArrayCFUnc.Clear();
            CF_ArrayReel.Clear();
            CF_ArrayReelUnc.Clear();
            CF_ArrayComplex.Clear();
            CF_ArrayComplexUnc.Clear();
            CF_YK.Clear();
            CF_YK_Unc.Clear();
        }
    }
}
