import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class ExcelBirlestiriciGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Sekme Birleştirici @by Önder AKÖZ")
        self.root.geometry("600x400")
        self.root.configure(pady=20, padx=20)

        # Stil ayarları
        style = ttk.Style()
        style.configure('TButton', padding=10)
        style.configure('TLabel', padding=5)

        # Dosya seçim çerçevesi
        self.dosya_cerceve = ttk.LabelFrame(root, text="Dosya İşlemleri", padding=10)
        self.dosya_cerceve.pack(fill='x', padx=10, pady=5)

        # Dosya yolu etiketi
        self.dosya_yolu_label = ttk.Label(self.dosya_cerceve, text="Seçili Dosya: Henüz dosya seçilmedi")
        self.dosya_yolu_label.pack(fill='x', pady=5)

        # Dosya seçim butonu
        self.dosya_sec_btn = ttk.Button(self.dosya_cerceve, 
                                      text="Excel Dosyası Seç",
                                      command=self.dosya_sec)
        self.dosya_sec_btn.pack(pady=5)

        # Kayıt yeri çerçevesi
        self.kayit_cerceve = ttk.LabelFrame(root, text="Kayıt İşlemleri", padding=10)
        self.kayit_cerceve.pack(fill='x', padx=10, pady=5)

        # Kayıt yolu etiketi
        self.kayit_yolu_label = ttk.Label(self.kayit_cerceve, text="Kayıt Yeri: Henüz seçilmedi")
        self.kayit_yolu_label.pack(fill='x', pady=5)

        # Kayıt yeri seçim butonu
        self.kayit_sec_btn = ttk.Button(self.kayit_cerceve, 
                                      text="Kayıt Yerini Seç",
                                      command=self.kayit_yeri_sec)
        self.kayit_sec_btn.pack(pady=5)

        # İlerleme çubuğu
        self.ilerleme = ttk.Progressbar(root, length=300, mode='determinate')
        self.ilerleme.pack(pady=20)

        # Birleştirme butonu
        self.birlestir_btn = ttk.Button(root, 
                                      text="Sekmeleri Birleştir",
                                      command=self.sekmeleri_birlestir)
        self.birlestir_btn.pack(pady=10)

        # Durum mesajı
        self.durum_label = ttk.Label(root, text="")
        self.durum_label.pack(pady=10)

        # Dosya yollarını saklamak için değişkenler
        self.excel_dosya_yolu = None
        self.kayit_dosya_yolu = None

    def dosya_sec(self):
        dosya_yolu = filedialog.askopenfilename(
            title="Excel Dosyası Seç",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls")]
        )
        if dosya_yolu:
            self.excel_dosya_yolu = dosya_yolu
            self.dosya_yolu_label.config(text=f"Seçili Dosya: {os.path.basename(dosya_yolu)}")

    def kayit_yeri_sec(self):
        kayit_yolu = filedialog.asksaveasfilename(
            title="Kayıt Yerini Seç",
            defaultextension=".xlsx",
            filetypes=[("Excel Dosyası", "*.xlsx")]
        )
        if kayit_yolu:
            self.kayit_dosya_yolu = kayit_yolu
            self.kayit_yolu_label.config(text=f"Kayıt Yeri: {os.path.basename(kayit_yolu)}")

    def sekmeleri_birlestir(self):
        if not self.excel_dosya_yolu or not self.kayit_dosya_yolu:
            messagebox.showerror("Hata", "Lütfen hem kaynak dosyayı hem de kayıt yerini seçin!")
            return

        try:
            # İlerleme çubuğunu sıfırla
            self.ilerleme['value'] = 0
            self.durum_label.config(text="Birleştirme işlemi başlıyor...")
            self.root.update()

            # Excel dosyasını oku
            excel = pd.ExcelFile(self.excel_dosya_yolu)
            tum_dataframeler = []
            
            # Her sekme için işlem yap
            toplam_sekme = len(excel.sheet_names)
            for idx, sayfa in enumerate(excel.sheet_names):
                df = pd.read_excel(self.excel_dosya_yolu, sheet_name=sayfa)
                df['Kaynak_Sekme'] = sayfa
                tum_dataframeler.append(df)
                
                # İlerleme çubuğunu güncelle
                ilerleme = (idx + 1) / toplam_sekme * 100
                self.ilerleme['value'] = ilerleme
                self.durum_label.config(text=f"İşleniyor: {sayfa}")
                self.root.update()

            # Birleştirme işlemi
            self.durum_label.config(text="Veriler birleştiriliyor...")
            self.root.update()
            birlestirilmis_df = pd.concat(tum_dataframeler, ignore_index=True)

            # Dosyayı kaydet
            self.durum_label.config(text="Dosya kaydediliyor...")
            self.root.update()
            birlestirilmis_df.to_excel(self.kayit_dosya_yolu, index=False)

            # İşlem tamamlandı
            self.ilerleme['value'] = 100
            self.durum_label.config(text="İşlem başarıyla tamamlandı!")
            messagebox.showinfo("Başarılı", "Birleştirme işlemi tamamlandı!")

        except Exception as e:
            messagebox.showerror("Hata", f"Bir hata oluştu:\n{str(e)}")
            self.durum_label.config(text="Hata oluştu!")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelBirlestiriciGUI(root)
    root.mainloop()
