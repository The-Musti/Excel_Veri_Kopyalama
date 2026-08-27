using ClosedXML.Excel;
using System.Net.Sockets;
using System.Threading.Tasks;
using Excel_Veri_Kopyalama.Enums;

namespace Excel_Veri_Kopyalama.Services
{
    public class ExcelCopyService
    {
        /// <summary>
        /// MB şablonu için seçilen opsiyona göre kaynak Excel'den verileri okur ve hedefe yazar.
        /// </summary>
        /// <param name="sourceFilePath">Okunacak kaynak dosyanın yolu.</param>
        /// <param name="targetFilePath">Verilerin yazılacağı hedef dosyanın yolu.</param>
        /// <param name="option">Arayüzden seçilen alt işlem türü (Noktasal, Lineer vb.).</param>
        public async Task RunTemplateMBAsync(string sourceFilePath, string targetFilePath, OptionType option)
        {
            await Task.Run(() =>
            {
                // Kaynak ve hedef Excel dosyalarını hafızaya alır
                using var sourceWb = new XLWorkbook(sourceFilePath);
                using var targetWb = new XLWorkbook(targetFilePath);

                // İşlem yapılacak ilk sayfalarını seçer
                var sourceWs = sourceWb.Worksheet(1);
                var targetWs = targetWb.Worksheet(1);

                // Kaynak dosyadaki son satırı ve hedef dosyadaki ilk boş satırı bulur
                int lastSourceRow = sourceWs.LastRowUsed().RowNumber();
                int targetStartRow = targetWs.LastRowUsed()?.RowNumber() + 1 ?? 2;

                int targetRow = targetStartRow;

                // Kaynak dosyadaki verileri satır satır dönerek okur
                for (int row = 2; row <= lastSourceRow; row++)
                {
                    // Noktasal Yangın için veri formatlama ve kopyalama kuralları
                    if (option == OptionType.NoktasalYangın)
                    {
                        string ipAdress = sourceWs.Cell($"B{row}").GetString();
                        string mbReg = sourceWs.Cell($"C{row}").GetString();
                        string compname = sourceWs.Cell($"D{row}").GetString();
                        string comptype = sourceWs.Cell($"E{row}").GetString();

                        targetWs.Cell($"A{targetRow}").Value = ipAdress;
                        targetWs.Cell($"B{targetRow}").Value = $"{compname}.{comptype}";
                        targetWs.Cell($"C{targetRow}").Value = compname;
                        targetWs.Cell($"D{targetRow}").Value = comptype;
                        targetWs.Cell($"E{targetRow}").Value = "H";
                        targetWs.Cell($"F{targetRow}").Value = mbReg;
                        targetWs.Cell($"G{targetRow}").Value = "1";

                        targetRow++;
                    }
                    // Lineer Yangın için veri formatlama kuralları buraya eklenecek
                    if (option == OptionType.LineerYangın)
                    {

                    }
                    // Sulu Söndürme için veri formatlama kuralları buraya eklenecek
                    if (option == OptionType.SuluSondurme)
                    {

                    }
                    // VTS için veri formatlama kuralları buraya eklenecek
                    if (option == OptionType.VTS)
                    {

                    }
                    // VMS için veri formatlama kuralları buraya eklenecek
                    if (option == OptionType.VMS)
                    {

                    }


                }
                // Değişiklikleri hedef dosyaya kaydeder
                targetWb.Save();
            });
        }
        /// <summary>
        /// OPC şablonu için seçilen opsiyona göre kaynak Excel'den verileri okur ve hedefe yazar.
        /// </summary>
        /// <param name="sourceFilePath">Okunacak kaynak dosyanın yolu.</param>
        /// <param name="targetFilePath">Verilerin yazılacağı hedef dosyanın yolu.</param>
        /// <param name="option">Arayüzden seçilen yazılım opsiyonu.</param>
        public async Task RunTemplateOPCAsync(string sourceFilePath, string targetFilePath, OptionType option)
        {
            await Task.Run(() =>
            {
                // Kaynak ve hedef Excel dosyalarını hafızaya alır
                using var sourceWb = new XLWorkbook(sourceFilePath);
                using var targetWb = new XLWorkbook(targetFilePath);

                // İşlem yapılacak ilk sayfaları seçer
                var sourceWs = sourceWb.Worksheet(1);
                var targetWs = targetWb.Worksheet(1);

                // Kepware için veri formatlama kuralları buraya eklenecek
                if (option == OptionType.Kepware)
                {

                }
                // Rslinx için veri formatlama kuralları buraya eklenecek
                if (option == OptionType.Rslinx)
                {

                }
                // FactoryTalk için veri formatlama kuralları buraya eklenecek
                if (option == OptionType.FactoryTalk)
                {

                }
                // Değişiklikleri hedef dosyaya kaydeder
                targetWb.Save();
            });
        }
    }
}