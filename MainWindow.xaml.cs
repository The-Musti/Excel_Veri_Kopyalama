using Excel_Veri_Kopyalama.Enums;
using Excel_Veri_Kopyalama.Services;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace Excel_Veri_Kopyalama
{
    public partial class MainWindow : Window
    {
        // Kullanıcının arayüzden yaptığı seçimleri tutacağımız değişkenler
        private TemplateType _selectedTemplate = TemplateType.None;
        private OptionType _selectedOption = OptionType.None;

        // Kopyalama işlemlerini yürütecek servis sınıfının örneği
        private readonly ExcelCopyService _copyService = new ExcelCopyService();

        // Hangi Şablon + Opsiyon kombinasyonunun hangi dosyaya yazılacağını tutan liste
        private readonly Dictionary<(TemplateType, OptionType), string> _templateTargetPaths;

        // Hangi şablon seçildiğinde, sağ üstte hangi seçeneklerin çıkacağını tutan liste.
        private readonly Dictionary<TemplateType, List<OptionType>> _templateOptionsMapping;

        public MainWindow()
        {
            InitializeComponent();

            // Uygulama farklı bilgisayarlarda da çalışabilsin diye dinamik masaüstü yolu alınıyor.
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Hangi sol menü şablonunun, hangi alt opsiyonlara sahip olduğunun tanımlanması.
            _templateOptionsMapping = new Dictionary<TemplateType, List<OptionType>>
            {
               { TemplateType.MB, new List<OptionType> { OptionType.NoktasalYangın, OptionType.LineerYangın, OptionType.SuluSondurme, OptionType.VTS, OptionType.VMS } },
               { TemplateType.OPC, new List<OptionType> { OptionType.Kepware, OptionType.Rslinx, OptionType.FactoryTalk  } },
               
                // Şu anlık bu şablonların opsiyonu yok, 'None' placeholder. 
               { TemplateType.ORT, new List<OptionType> { OptionType.None } },
               { TemplateType.ANADOLU, new List<OptionType> { OptionType.None } },
               { TemplateType.SMMP, new List<OptionType> { OptionType.None } }
            };


            // Seçilen şablon ve opsiyon ikilisine göre verilerin hangi Excel dosyasına yazılacağının tanımlanması.
            _templateTargetPaths = new Dictionary<(TemplateType, OptionType), string>
            {
                // MB dosyaları
                { (TemplateType.MB, OptionType.NoktasalYangın), System.IO.Path.Combine(desktopPath, @"task\MBDriverTagConf (44).xlsx") },
                { (TemplateType.MB, OptionType.LineerYangın), System.IO.Path.Combine(desktopPath, @"task\HEDEF_Lineer_Yangın.xlsx") },
                { (TemplateType.MB, OptionType.SuluSondurme), System.IO.Path.Combine(desktopPath, @"task\HEDEF_Sulu_Sondurme).xlsx") },
                { (TemplateType.MB, OptionType.VTS), System.IO.Path.Combine(desktopPath, @"task\HEDEF_VTS.xlsx") },
                { (TemplateType.MB, OptionType.VMS), System.IO.Path.Combine(desktopPath, @"task\HEDEF_VMS.xlsx") },

                // OPC dosyaları
                { (TemplateType.OPC, OptionType.Kepware), System.IO.Path.Combine(desktopPath, @"task\HEDEF_Kepware.xlsx") },
                { (TemplateType.OPC, OptionType.Rslinx), System.IO.Path.Combine(desktopPath, @"task\HEDEF_Rslinx.xlsx") },
                { (TemplateType.OPC, OptionType.FactoryTalk), System.IO.Path.Combine(desktopPath, @"task\HEDEF_FactoryTalk.xlsx") },

                // ORT dosyaları
                { (TemplateType.ORT, OptionType.None), System.IO.Path.Combine(desktopPath, @"task\HEDEF_None.xlsx") },

                // ANADOLU dosyaları
                { (TemplateType.ANADOLU, OptionType.None), System.IO.Path.Combine(desktopPath, @"task\HEDEF_None.xlsx") },

                // SMMP dosyaları
                { (TemplateType.SMMP, OptionType.None), System.IO.Path.Combine(desktopPath, @"task\HEDEF_None.xlsx") },

            };
        }

        /// <summary>
        /// Sol menüden bir şablon seçtiğinde tetiklenir.
        /// </summary>
        private void TemplateRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var rb = sender as RadioButton;
            if (rb == null) return;

            // XAML tarafında butona verdiğimiz 'Tag' numarasına göre Enum değerini belirliyoruz.
            _selectedTemplate = (TemplateType)int.Parse(rb.Tag.ToString());

            // Sağ üstteki seçenekleri yeni şablona göre güncellemesi için metodu çağırıyoruz
            UpdateOptionsUI(_selectedTemplate);
        }

        /// <summary>
        /// Seçilen şablona göre sağ üstteki RadioButton'ları silip yeniden çizer.
        /// </summary>
        private void UpdateOptionsUI(TemplateType selectedTemplate)
        {
           
            if (OptionsPanel != null)
            {
                // Ekranı temizle ve önceki seçimi sıfırla
                OptionsPanel.Children.Clear();
                _selectedOption = OptionType.None; // Önceki seçimi sıfırla

                // Şablonun bir opsiyon listesi var mı diye kontrol edilmesi.
                if (_templateOptionsMapping.TryGetValue(selectedTemplate, out var availableOptions))
                {
                    foreach (var option in availableOptions)
                    {
                        // Her bir seçenek için dinamik buton oluştur
                        RadioButton optRb = new RadioButton
                        {
                            Content = option.ToString(),
                            Tag = (int)option,
                            Margin = new Thickness(0, 0, 20, 0),
                            VerticalAlignment = VerticalAlignment.Center
                        };

                        // Oluşturulan butona tıklandığında çalışacak olayı bağla.
                        optRb.Checked += OptionRadioButton_Checked;

                        // Butonu ekrana ekle.
                        OptionsPanel.Children.Add(optRb);
                    }
                }
            }
        }


        /// <summary>
        /// Sağ üstten bir opsiyon seçildiğinde çalışır ve seçimi hafızaya alır.
        /// </summary>
        private void OptionRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var rb = sender as RadioButton;
            if (rb == null) return;

            _selectedOption = (OptionType)int.Parse(rb.Tag.ToString());
        }

        /// <summary>
        /// Kullanıcı Excel dosyasını sürükleyip kutunun üzerine getirdiğinde arkaplanı mavi yapar.
        /// </summary>
        private void DropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                DropZoneBorder.Background = System.Windows.Media.Brushes.LightBlue;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// Fare kutunun dışına çıktığında arkaplanı eski haline çevirir.
        /// </summary>
        private void DropZone_DragLeave(object sender, DragEventArgs e)
        {
            DropZoneBorder.Background = System.Windows.Media.Brushes.WhiteSmoke;
        }

        /// <summary>
        /// Dosya kutunun içine bırakıldığında çalışır ve dosya yolunu alıp işlemeye başlar.
        /// </summary>
        private void DropZone_Drop(object sender, DragEventArgs e)
        {
            DropZoneBorder.Background = System.Windows.Media.Brushes.WhiteSmoke;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    HandleSelectedFile(files[0]);
                }
            }
        }

        /// <summary>
        /// Kullanıcı kutuya tıkladığında standart Windows dosya seçme penceresini açar.
        /// </summary>
        private void DropZoneText_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls" };
            if (dlg.ShowDialog() == true)
            {
                HandleSelectedFile(dlg.FileName);
            }
        }

        /// <summary>
        /// Seçilen dosyayı alır, gerekli kontrolleri yapar ve uygun kopyalama senaryosunu başlatır.
        /// </summary>
        private async void HandleSelectedFile(string filePath)
        {
            // Sistemin çökmemesi için sadece Excel formatlarına izin veriyoruz.
            if (!filePath.EndsWith(".xlsx") && !filePath.EndsWith(".xls"))
            {
                MessageBox.Show("Lütfen bir Excel dosyası seçin.", "Geçersiz Dosya", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Sol taraftan şablon seçilmiş mi onun kontrolü.
            if (_selectedTemplate == TemplateType.None)
            {
                MessageBox.Show("Lütfen önce sol taraftan bir şablon seçin.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Seçilen şablonun opsiyonu varsa, sağ üstten bir seçenek seçilmiş mi onun kontolü.
            if (_templateOptionsMapping.ContainsKey(_selectedTemplate) && _templateOptionsMapping[_selectedTemplate].Count > 0 && _selectedOption == OptionType.None)
            {
                MessageBox.Show("Lütfen sağ üst bölümden bir seçenek seçin.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Arayüzdeki yazıyı güncelliyoruz
            DropZoneText.Text = $"Seçilen dosya: {System.IO.Path.GetFileName(filePath)}";

            // Yapılan seçimlere uygun bir hedef dosya tanımlama kontrolü.
            if (!_templateTargetPaths.TryGetValue((_selectedTemplate, _selectedOption), out string targetPath))
            {
                MessageBox.Show("Seçilen şablon ve opsiyon kombinasyonu için bir hedef dosya tanımlanmamış.", "Dosya Yolu Eksik", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Hedef yolda belirtilen Excel dosyasının var olup olmama kontrolü.
            if (!System.IO.File.Exists(targetPath))
            {
                MessageBox.Show($"Hedef şablon dosyası bulunamadı.\nBeklenen yol: {targetPath}", "Dosya Eksik", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Kopyalama işleminin başlatılması.
            try
            {
                
                switch (_selectedTemplate)
                {
                    case TemplateType.MB:
                        await _copyService.RunTemplateMBAsync(filePath, targetPath, _selectedOption);
                        MessageBox.Show("Kopyalama tamamlandı!", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information);
                        break;

                    case TemplateType.OPC:
                        await _copyService.RunTemplateOPCAsync(filePath, targetPath, _selectedOption);
                        MessageBox.Show("Kopyalama tamamlandı!", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information);
                        break;

                    case TemplateType.ORT:
                        // İleride buraya ORT metodu eklenecek
                        break;

                    case TemplateType.ANADOLU:
                        // İleride buraya ANADOLU metodu eklenecek
                        break; 

                    case TemplateType.SMMP:
                        // İleride buraya SMMP metodu eklenecek
                        break;
                }
            }
            catch (System.IO.IOException)
            {
                // Dosya o an Excel uygulamasında açıksa arka planda veri yazılamaz, bu yüzden uyarı veriyoruz.
                MessageBox.Show("Hedef veya kaynak dosyası şu an başka bir programda açık. Lütfen dosyayı kapatıp tekrar deneyin.", "Dosya Kullanımda", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hata oluştu: {ex.Message}", "Beklenmeyen Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}