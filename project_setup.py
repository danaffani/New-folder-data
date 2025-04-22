import os
import shutil
import glob

def create_folder_structure():
    # Membuat folder jika belum ada
    folders = ['raw_input', 'input', 'output']
    for folder in folders:
        if not os.path.exists(folder):
            os.makedirs(folder)
            print(f"Folder '{folder}' berhasil dibuat")
        else:
            print(f"Folder '{folder}' sudah ada")

def move_txt_files():
    # Memindahkan semua file tabel_*.txt ke folder raw_input
    txt_files = glob.glob("tabel_*.txt")
    for file in txt_files:
        dest_path = os.path.join("raw_input", file)
        shutil.copy2(file, dest_path)
        print(f"File '{file}' berhasil disalin ke folder 'raw_input'")

if __name__ == "__main__":
    print("Menyiapkan struktur proyek...")
    create_folder_structure()
    move_txt_files()
    print("Struktur proyek berhasil dibuat!") 