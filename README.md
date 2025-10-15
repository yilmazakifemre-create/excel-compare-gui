# Excel Compare — Renkli Arayüz

Basit bir masaüstü uygulama (Tkinter + openpyxl) ile iki Excel dosyasındaki belirli aralıkları karşılaştırır:
- Aynı olan hücreleri **yeşil**, farklı/hitap edemeyenleri **kırmızı** ile boyar (kopya dosyalara kaydeder).
- Aynı olanları yan yana **result_matches.xlsx** dosyasına yazar (örneğin C ve D sütunları).

## Kullanım
1. Python 3.8+ kurulu olmalı.
2. Gerekli paketleri yükleyin:
```bash
pip install -r requirements.txt
```
3. `main.py` çalıştırın:
```bash
python main.py
```

## Notlar
- Aralık formatı: `A1:A100` veya `B2:B200` veya `A:A` gibi.
- Uygulama, aktif çalışma dizinine (çalıştırdığınız klasöre) şu dosyaları kaydeder:
  - `<file1_basename>_colored.xlsx`
  - `<file2_basename>_colored.xlsx`
  - `result_matches.xlsx`

## Lisans
MIT
