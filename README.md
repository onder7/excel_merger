# Excel Data Merger & Filler Tool
<img width="1366" height="735" alt="image" src="https://github.com/user-attachments/assets/7be0ca0c-e3b0-4d32-be5a-d66ca880e150" />

A powerful Python GUI application that intelligently merges two Excel files and fills missing data based on a common column.

## 🚀 Features

- **Smart Column Matching**: Automatically detects and matches common columns
- **Data Filling**: Fills empty cells in the primary file with data from the secondary file
- **Visual Feedback**: Color-coded results showing filled, new, and unchanged records
- **User-Friendly GUI**: Modern interface with progress tracking
- **Flexible Column Selection**: Choose any column as the merge key

## 🛠️ Installation

1. Clone the repository:
```bash
git clone https://github.com/onder7/excel_merger.git
cd excel-merger-tool
```

2. Install required dependencies:
```bash
pip install pandas openpyxl
```

3. Run the application:
```bash
python excel_merger.py
```

## 📖 How It Works

1. **Select Files**: Choose two Excel files to merge
2. **Choose Columns**: Select the common column (e.g., Computer Name, ID)
3. **Auto-Match**: The tool automatically suggests matching columns
4. **Merge & Fill**: Primary file's empty cells are filled with data from secondary file
5. **Export**: Get a complete, merged Excel file with color-coded changes

## 🎨 Color Coding

- 🟡 **Yellow**: New records added from secondary file
- 🟢 **Green**: Records where data was filled
- ⚪ **White**: Unchanged records

## 📊 Use Cases

- Merge employee data from different systems
- Fill missing information in inventory lists
- Combine computer asset information
- Update customer records with missing details

## 🔧 Requirements

- Python 3.6+
- pandas
- openpyxl
- tkinter (included with Python)

```

## 🤝 Contributing

Feel free to submit issues and enhancement requests!

## 📄 License

This project is open source and available under the MIT License.

---

# Excel Veri Birleştirme ve Doldurma Aracı
<img width="1366" height="735" alt="image" src="https://github.com/user-attachments/assets/6755be06-201e-4201-8da7-5ecb2bd7c82e" />

Ortak sütun temelinde iki Excel dosyasını akıllıca birleştiren ve eksik verileri dolduran güçlü bir Python GUI uygulaması.

## 🚀 Özellikler

- **Akıllı Sütun Eşleştirme**: Ortak sütunları otomatik olarak algılar ve eşleştirir
- **Veri Doldurma**: Ana dosyadaki boş hücreleri ikinci dosyadan gelen verilerle doldurur
- **Görsel Geri Bildirim**: Doldurulan, yeni ve değişmeden kalan kayıtları renk kodlarıyla gösterir
- **Kullanıcı Dostu Arayüz**: İlerleme takibi olan modern arayüz
- **Esnek Sütun Seçimi**: Herhangi bir sütunu birleştirme anahtarı olarak seçebilirsiniz

## 🛠️ Kurulum

1. Projeyi klonlayın:
```bash
git clone https://github.com/onder7/excel_merger.git
cd excel-merger-tool
```

2. Gerekli kütüphaneleri yükleyin:
```bash
pip install pandas openpyxl
```

3. Uygulamayı çalıştırın:
```bash
python excel_merger.py
```

## 📖 Nasıl Çalışır

1. **Dosya Seçimi**: Birleştirilecek iki Excel dosyasını seçin
2. **Sütun Seçimi**: Ortak sütunu seçin (örn: Computer Name, ID)
3. **Otomatik Eşleştirme**: Araç otomatik olarak eşleşen sütunları önerir
4. **Birleştir ve Doldur**: Ana dosyanın boş hücreleri ikinci dosyadan doldurulur
5. **Dışa Aktar**: Renk kodlu değişikliklerle tam, birleştirilmiş Excel dosyası alın

## 🎨 Renk Kodları

- 🟡 **Sarı**: İkinci dosyadan eklenen yeni kayıtlar
- 🟢 **Yeşil**: Veri doldurma yapılan kayıtlar
- ⚪ **Beyaz**: Değişmeden kalan kayıtlar

## 📊 Kullanım Alanları

- Farklı sistemlerden çalışan verilerini birleştirme
- Envanter listelerindeki eksik bilgileri doldurma
- Bilgisayar varlık bilgilerini birleştirme
- Müşteri kayıtlarını eksik detaylarla güncelleme

## 🔧 Gereksinimler

- Python 3.6+
- pandas
- openpyxl
- tkinter (Python ile birlikte gelir)


## 🤝 Katkıda Bulunma

Sorunları ve geliştirme önerilerini bildirmekten çekinmeyin!

## 📄 Lisans

Bu proje açık kaynak kodludur ve MIT Lisansı altında kullanılabilir.
