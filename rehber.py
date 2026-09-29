import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import re
import os

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Numara Karşılaştırıcı")
        self.root.geometry("500x430")
        self.root.resizable(False, False)

        tk.Label(root, text="Numara Karşılaştırma Programı", font=("Arial", 14, "bold")).pack(pady=10)

        # 1. TXT Seçimi
        tk.Button(root, text="1. Not Defteri (TXT) Seç", command=self.sec_txt, width=30, bg="#e1e1e1").pack(pady=5)
        self.lbl_txt = tk.Label(root, text="Seçilmedi", fg="gray")
        self.lbl_txt.pack()

        # 2. Excel Seçimi
        tk.Button(root, text="2. Excel Dosyası (.xlsx) Seç", command=self.sec_excel, width=30, bg="#e1e1e1").pack(pady=5)
        self.lbl_excel = tk.Label(root, text="Seçilmedi", fg="gray")
        self.lbl_excel.pack()

        # 3. Sütun Seçimi (Açılır Kutulu)
        tk.Label(root, text="Excel Telefon Sütununu Seçin:").pack(pady=(10,0))
        self.combo_sutun = ttk.Combobox(root, width=28, state="readonly")
        self.combo_sutun.pack()

        # 4. Çalıştır Butonu
        tk.Button(root, text="İŞLEMİ BAŞLAT VE KAYDET", command=self.baslat, bg="#28a745", fg="white", font=("Arial", 11, "bold"), height=2).pack(pady=20)

        self.txt_path = ""
        self.excel_path = ""

    def sec_txt(self):
        filename = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if filename:
            self.txt_path = filename
            self.lbl_txt.config(text=os.path.basename(filename), fg="green")

    def sec_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            self.excel_path = filename
            self.lbl_excel.config(text=os.path.basename(filename), fg="green")
            
            # Excel sütunlarını otomatik çekip açılır kutuya ekle
            try:
                df = pd.read_excel(filename, nrows=1)  # Sadece başlıkları hızlıca okur
                sutunlar = list(df.columns)
                self.combo_sutun['values'] = sutunlar
                
                # İçinde "telefon", "tel" geçen sütunu varsayılan olarak otomatik seç
                varsayilan_sutun = None
                for col in sutunlar:
                    if "telefon" in str(col).lower() or "tel" in str(col).lower():
                        varsayilan_sutun = col
                        break
                
                if varsayilan_sutun:
                    self.combo_sutun.set(varsayilan_sutun)
                elif sutunlar:
                    self.combo_sutun.current(0)
            except Exception as e:
                messagebox.showerror("Hata", f"Excel sütunları okunamadı:\n{str(e)}")

    def numara_sadelestir(self, val):
        if pd.isna(val) or val is None:
            return ""
        s = str(val).split('.')[0]
        rakamlar = re.sub(r'\D', '', s)
        if len(rakamlar) >= 10:
            return rakamlar[-10:]
        elif len(rakamlar) >= 7:
            return rakamlar
        return ""

    def baslat(self):
        if not self.txt_path or not self.excel_path:
            messagebox.showerror("Hata", "Lütfen hem TXT hem de Excel dosyasını seçin!")
            return

        sutun = self.combo_sutun.get().strip()
        if not sutun:
            messagebox.showerror("Hata", "Lütfen karşılaştırılacak Excel sütununu seçin!")
            return

        try:
            # TXT Oku
            rehber = set()
            with open(self.txt_path, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    sade = self.numara_sadelestir(line)
                    if sade:
                        rehber.add(sade)

            # Excel Oku
            df = pd.read_excel(self.excel_path)
            if sutun not in df.columns:
                messagebox.showerror("Hata", f"Seçilen '{sutun}' sütunu Excel dosyasında bulunamadı!")
                return

            df["Eşleşme Durumu"] = df[sutun].apply(lambda x: "Var" if self.numara_sadelestir(x) in rehber else "")

            kayit_yeri = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile="Sonuc_Liste.xlsx")
            if kayit_yeri:
                df.to_excel(kayit_yeri, index=False)
                messagebox.showinfo("Başarılı", "İşlem tamamlandı ve dosya kaydedildi!")

        except Exception as e:
            messagebox.showerror("Hata", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    
    # Logo yükleme desteği
    if os.path.exists("logo.ico"):
        try:
            root.iconbitmap("logo.ico")
        except Exception:
            pass
            
    app = App(root)
    root.mainloop()