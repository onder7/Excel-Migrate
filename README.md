# Excel Sekme Birleştirici

Excel dosyalarındaki tüm sekmeleri tek bir Excel dosyasında birleştirmenize olanak sağlayan kullanıcı dostu bir grafiksel arayüz uygulaması.

## 🚀 Özellikler

- 📂 Excel dosyası seçme ve kaydetme
- 📊 Tüm sekmeleri otomatik birleştirme
- 📋 Kaynak sekme bilgisi ekleme
- 🔄 İlerleme durumu gösterimi
- ⚠️ Hata kontrolü ve bilgilendirme

## 💻 Gereksinimler

```python
pip install pandas
pip install tkinter
```

## 🛠️ Kurulum

1. Repoyu klonlayın:
```bash
git clone https://github.com/onder7/Excel-Migrate.git
```

2. Gerekli paketleri yükleyin:
```bash
pip install -r requirements.txt
```

3. Uygulamayı çalıştırın:
```bash
python excel.py
```

## 📝 Kullanım

1. **Excel Dosyası Seçme**
   - "Excel Dosyası Seç" butonuna tıklayın
   - Birleştirmek istediğiniz sekmeleri içeren Excel dosyasını seçin

2. **Kayıt Yeri Belirleme**
   - "Kayıt Yerini Seç" butonuna tıklayın
   - Birleştirilmiş dosyanın kaydedileceği yeri ve ismini belirleyin

3. **Birleştirme**
   - "Sekmeleri Birleştir" butonuna tıklayın
   - İşlem durumunu ilerleme çubuğundan takip edin

## ⚙️ Teknik Detaylar

### Ana Bileşenler

```python
class ExcelBirlestiriciGUI:
    def __init__(self, root):
        # GUI bileşenleri
        self.dosya_cerceve      # Dosya işlemleri çerçevesi
        self.kayit_cerceve      # Kayıt işlemleri çerçevesi
        self.ilerleme          # İlerleme çubuğu
        self.durum_label       # Durum mesajı etiketi
```

### Temel Fonksiyonlar

```python
def dosya_sec(self):
    # Excel dosyası seçimi
    # Desteklenen formatlar: .xlsx, .xls

def kayit_yeri_sec(self):
    # Kayıt yeri ve dosya adı belirleme
    # Varsayılan format: .xlsx

def sekmeleri_birlestir(self):
    # Sekmeleri birleştirme işlemi
    # Her sekmeye kaynak bilgisi ekleme
    # İlerleme durumu güncelleme
```

## 📊 Veri İşleme Süreci

1. **Dosya Okuma**
   ```python
   excel = pd.ExcelFile(self.excel_dosya_yolu)
   ```

2. **Sekme İşleme**
   ```python
   for sayfa in excel.sheet_names:
       df = pd.read_excel(...)
       df['Kaynak_Sekme'] = sayfa
   ```

3. **Birleştirme**
   ```python
   birlestirilmis_df = pd.concat(tum_dataframeler)
   ```

4. **Kaydetme**
   ```python
   birlestirilmis_df.to_excel(self.kayit_dosya_yolu)
   ```

## 🔍 Hata Yönetimi

- Dosya seçim kontrolü
- Format uyumluluk kontrolü
- İşlem süreci hata yakalama
- Kullanıcı bilgilendirme mesajları

## 🛠️ Geliştirme Önerileri

### Performans İyileştirmeleri
- [ ] Büyük dosyalar için chunk-based okuma
- [ ] Bellek optimizasyonu
- [ ] Çoklu işlem desteği

### Yeni Özellikler
- [ ] Çoklu dosya desteği
- [ ] Özel sekme seçimi
- [ ] Veri filtreleme
- [ ] Önizleme özelliği

### Arayüz İyileştirmeleri
- [ ] Tema seçenekleri
- [ ] Dil desteği
- [ ] Detaylı ilerleme bilgisi
- [ ] Özelleştirilebilir arayüz

## ⚠️ Bilinen Sınırlamalar

- Çok büyük Excel dosyaları için bellek kullanımı
- Tek dosya işleme sınırlaması
- Temel hata yakalama

## 👥 Katkıda Bulunma

1. Fork yapın
2. Feature branch oluşturun (`git checkout -b feature/AmazingFeature`)
3. Commit yapın (`git commit -m 'Add some AmazingFeature'`)
4. Branch'i push yapın (`git push origin feature/AmazingFeature`)
5. Pull Request açın

## 📝 Lisans

Bu proje [MIT](LICENSE) lisansı altında lisanslanmıştır.

## 📞 İletişim

Önder AKÖZ - [@onderakoz](https://linkedin.com/in/onderakoz)

Proje Linki: https://github.com/onder7/Excel-Migrate/
