# Excel Veri Kopyalayıcı/ Excel Mapper

Belirli şablonlara ve kurallara göre kaynak Excel dosyalarından veri okuyup, hedef Excel dosyalarına (mapping) aktaran bir WPF otomasyon aracıdır. Bu uygulama, tekrarlayan veri taşıma işlemlerini otomatize etmek ve insan hatasını en aza indirmek için tasarlanmıştır.

## Özellikler
- Sürükle-bırak desteği ile kolay dosya seçimi.
- Seçilen ana şablona göre dinamik olarak güncellenen arayüz seçenekleri.
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) kütüphanesi kullanılarak arka planda hızlı ve Excel'in kurulu olmasından bağımsız veri okuma/yazma işlemleri.

## Kurulum ve Gereksinimler
- Proje bağımlılıklarını yüklemek için projenin dizininde terminali açıp `dotnet restore` komutunu çalıştırın veya Visual Studio üzerinden NuGet paketlerini geri yükleyin.
- Uygulamanın düzgün çalışabilmesi için hedef şablon Excel dosyalarının, kullanıcının masaüstünde `task` adlı bir klasörde (örn: `C:\Users\KullaniciAdi\Desktop\task\`) bulunması gerekir. (Gerekli dosya isimleri ve yolları `MainWindow.xaml.cs` içerisinden değiştirilebilir).

## Kullanım
1. Arayüzün sol tarafındaki menüden işlemi yapılacak ana şablonu (MB, OPC vb.) seçin.
2. Şablon seçildiğinde sağ üstte beliren alt opsiyonlardan (Noktasal Yangın, Kepware vb.) birini işaretleyin.
3. Kopyalanacak verileri içeren kaynak Excel dosyanızı (.xlsx veya .xls) ekrandaki sürükle-bırak alanına bırakın veya tıklayarak seçin.
4. Kopyalama işlemi otomatik başlar. Tamamlandığında ekranda bilgi mesajı gösterilir ve hedef dosyanız güncellenir.

## Geliştirici Kılavuzu

Uygulamanın arayüzü, yeni kurallar ve seçenekler eklenirken XAML tarafında minimum değişiklik gerektirecek şekilde dinamik tasarlanmıştır.

### Yeni Şablon Eklemek
1. `Enums.cs` dosyasındaki `TemplateType` içerisine yeni şablonu ekleyin.
2. `MainWindow.xaml` dosyasındaki `StackPanel` içerisine yeni şablon için bir `RadioButton` ekleyin ve `Tag` özelliğine Enum'daki sayısal karşılığını yazın.
3. `MainWindow.xaml.cs` dosyasındaki `_templateOptionsMapping` sözlüğüne şablonun alt opsiyonlarını tanımlayın (Opsiyon yoksa `OptionType.None` olarak belirtin).
4. `_templateTargetPaths` sözlüğüne bu şablonun kopyalama yapacağı hedef dosya yolunu ekleyin.

### Yeni Alt Opsiyon Eklemek
Arayüz opsiyonlar için dinamik çalıştığından dolayı yeni bir opsiyon eklerken XAML tasarım dosyasında değişiklik yapılmasına gerek yoktur:
1. `Enums.cs` dosyasındaki `OptionType` içerisine yeni opsiyonu ekleyin.
2. `MainWindow.xaml.cs` dosyasındaki `_templateOptionsMapping` listesinde, ilgili şablonun listesine bu opsiyonu ekleyin.
3. `_templateTargetPaths` sözlüğüne bu yeni opsiyon kombinasyonunun hedef dosya yolunu ekleyin.

### Mapping (Hücre Kopyalama) Kurallarını Değiştirmek
Tüm veri okuma ve formatlama mantığı `Services/ExcelCopyService.cs` dosyası üzerinden yürütülür.
- İlgili şablona ait kopyalama metodunu (örn: `RunTemplateMBAsync`) bulun.
- Seçilen opsiyonu yakalayan `if (option == OptionType...)` bloğu içerisindeki hücre atama kurallarını (`targetWs.Cell($"A{targetRow}").Value = ...`) ihtiyacınıza göre düzenleyin.
