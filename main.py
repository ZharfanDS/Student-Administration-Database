import tkinter as tk
from tkinter import ttk, messagebox, filedialog # Tambahkan filedialog
import pyodbc
import bcrypt
import pandas as pd # Import library pandas

# --- KONFIGURASI KONEKSI DATABASE ---
SERVER = r'' # input your server
DB_SISWA = r'DB_ADMINISTRASI' # input your database for student. Mine is DB_ADMINISTRASI (Make sure you create first using Microsoft SQL Server Management Studio)
DB_ADMIN = r'Admin_DB' # input your database for admin. Mine is Admin_DB (Make sure you create first using Microsoft SQL Server Management Studio)
# USERNAME = 'sa' # change with your username
# PASSWORD = 'PasswordAnda123' # change with your password

# --- KELAS UNTUK JENDELA LOGIN ---
class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Login Aplikasi")
        self.root.geometry("350x180")
        self.root.resizable(False, False)
        
        self.login_successful = False

        frame = ttk.Frame(root, padding="20")
        frame.pack(expand=True, fill="both")

        ttk.Label(frame, text="Username:").grid(row=0, column=0, sticky="w", pady=5)
        self.user_entry = ttk.Entry(frame, width=30)
        self.user_entry.grid(row=0, column=1, pady=5)

        ttk.Label(frame, text="Password:").grid(row=1, column=0, sticky="w", pady=5)
        self.pass_entry = ttk.Entry(frame, width=30, show="*")
        self.pass_entry.grid(row=1, column=1, pady=5)

        login_button = ttk.Button(frame, text="Login", command=self.check_login)
        login_button.grid(row=2, column=0, columnspan=2, pady=15)
        
        self.user_entry.focus_set()
        self.root.bind('<Return>', lambda event=None: login_button.invoke())

    def get_connection(self, db_name):
        try:
            conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER};DATABASE={db_name};Trusted_Connection=yes;'
            # Jika pakai SQL Server Authentication (username & password)
            # conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};'
            conn = pyodbc.connect(conn_str)
            return conn
        except pyodbc.Error as ex:
            messagebox.showerror("Koneksi Gagal", f"Gagal terhubung ke database:\n{ex}")
            return None

    def check_login(self):
        username = self.user_entry.get()
        password = self.pass_entry.get().encode('utf-8')

        if not username or not password:
            messagebox.showerror("Login Gagal", "Username dan Password tidak boleh kosong.")
            return

        conn = self.get_connection(DB_ADMIN)
        if conn:
            cursor = conn.cursor()
            cursor.execute("SELECT PasswordHash FROM Users WHERE Username=?", username)
            row = cursor.fetchone()
            cursor.close()
            conn.close()

            if row:
                stored_hash = row.PasswordHash.encode('utf-8')
                if bcrypt.checkpw(password, stored_hash):
                    messagebox.showinfo("Login Berhasil", f"Selamat datang, {username}!")
                    self.login_successful = True
                    self.root.destroy()
                else:
                    messagebox.showerror("Login Gagal", "Username atau Password salah.")
            else:
                messagebox.showerror("Login Gagal", "Username atau Password salah.")

# --- KELAS UTAMA APLIKASI ---
class AplikasiSiswa:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplikasi Administrasi Siswa")
        self.root.geometry("820x600") # Sedikit perbesar untuk tombol baru
        
        self.user_logged_out = False

        self.frame_tombol = ttk.LabelFrame(root, text="Menu Aksi")
        self.frame_tombol.pack(pady=10, padx=10, fill="x")
        ttk.Button(self.frame_tombol, text="Tambah Siswa Baru...", command=self.buka_dialog_tambah).pack(side="left", padx=10, pady=5)
        ttk.Button(self.frame_tombol, text="Update Data Siswa...", command=self.buka_dialog_update).pack(side="left", padx=10, pady=5)
        ttk.Button(self.frame_tombol, text="Hapus Data Siswa...", command=self.buka_dialog_hapus).pack(side="left", padx=10, pady=5)
        
        style = ttk.Style()
        style.configure("logout.TButton", foreground="blue")
        ttk.Button(self.frame_tombol, text="Logout", command=self.logout, style="logout.TButton").pack(side="right", padx=10, pady=5)
        
        self.frame_filter = ttk.LabelFrame(root, text="Pencarian dan Filter")
        self.frame_filter.pack(pady=5, padx=10, fill="x")
        
        ttk.Label(self.frame_filter, text="Cari Nama:").pack(side="left", padx=(10, 5), pady=5)
        self.search_entry = ttk.Entry(self.frame_filter, width=25)
        self.search_entry.pack(side="left", padx=5, pady=5)
        ttk.Button(self.frame_filter, text="Cari", command=self.cari_siswa).pack(side="left", padx=5, pady=5)
        ttk.Button(self.frame_filter, text="Tampilkan Semua", command=self.tampilkan_semua).pack(side="left", padx=5, pady=5)
        
        self.frame_sort = ttk.LabelFrame(root, text="Pengurutan dan Ekspor")
        self.frame_sort.pack(pady=5, padx=10, fill="x")
        ttk.Button(self.frame_sort, text="Urutkan Nama A-Z", command=self.urutkan_nama_az).pack(side="left", padx=(10, 5), pady=5)
        ttk.Button(self.frame_sort, text="Urutkan Nama Z-A", command=self.urutkan_nama_za).pack(side="left", padx=5, pady=5)
        ttk.Button(self.frame_sort, text="Urutkan ID ↑", command=self.urutkan_id_asc).pack(side="left", padx=(15, 5), pady=5)
        ttk.Button(self.frame_sort, text="Urutkan ID ↓", command=self.urutkan_id_desc).pack(side="left", padx=5, pady=5)
        
        # --- TOMBOL EKSPOR BARU ---
        style.configure("export.TButton", foreground="green")
        ttk.Button(self.frame_sort, text="Ekspor ke Excel", command=self.ekspor_ke_excel, style="export.TButton").pack(side="right", padx=10, pady=5)

        self.frame_tabel = ttk.LabelFrame(root, text="Data Siswa")
        self.frame_tabel.pack(pady=10, padx=10, fill="both", expand=True)
        self.tree = ttk.Treeview(self.frame_tabel, columns=("ID", "Nama", "Kelas", "Alamat"), show="headings")
        self.tree.heading("ID", text="ID"); self.tree.column("ID", width=50, anchor="center")
        self.tree.heading("Nama", text="Nama"); self.tree.column("Nama", width=200)
        self.tree.heading("Kelas", text="Kelas"); self.tree.column("Kelas", width=100)
        self.tree.heading("Alamat", text="Alamat"); self.tree.column("Alamat", width=300)
        scrollbar = ttk.Scrollbar(self.frame_tabel, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.muat_ulang_data()

    # --- FUNGSI BARU UNTUK EKSPOR KE EXCEL ---
    def ekspor_ke_excel(self):
        # Cek jika tabel kosong
        if not self.tree.get_children():
            messagebox.showinfo("Informasi", "Tidak ada data untuk diekspor.")
            return

        # Buka dialog "Save As" untuk memilih lokasi dan nama file
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Simpan sebagai file Excel"
        )

        # Jika pengguna membatalkan dialog, hentikan fungsi
        if not file_path:
            return

        try:
            # Ambil header kolom dari Treeview
            columns = [self.tree.heading(col)['text'] for col in self.tree['columns']]
            
            # Ambil semua data baris dari Treeview
            data_list = []
            for item_id in self.tree.get_children():
                data_list.append(list(self.tree.item(item_id, 'values')))

            # Buat DataFrame pandas dari data dan header
            df = pd.DataFrame(data_list, columns=columns)

            # Simpan DataFrame ke file Excel
            # index=False agar tidak ada kolom nomor baris tambahan dari pandas
            df.to_excel(file_path, index=False)

            messagebox.showinfo("Sukses", f"Data berhasil diekspor ke:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan saat mengekspor data:\n{e}")

    def logout(self):
        if messagebox.askyesno("Konfirmasi Logout", "Apakah Anda yakin ingin logout?"):
            self.user_logged_out = True
            self.root.destroy()

    def get_connection(self, db_name):
        try:
            conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER};DATABASE={db_name};Trusted_Connection=yes;'
            # Jika pakai SQL Server Authentication (username & password)
            # conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};'
            conn = pyodbc.connect(conn_str)
            return conn
        except pyodbc.Error as ex:
            messagebox.showerror("Koneksi Gagal", f"Gagal terhubung ke database:\n{ex}")
            return None

    def muat_ulang_data(self, search_term=None, sort_column='ID', sort_direction='ASC'):
        for i in self.tree.get_children(): self.tree.delete(i)
        conn = self.get_connection(DB_SISWA)
        if conn:
            cursor = conn.cursor()
            params = []
            valid_columns = ['ID', 'Nama']
            if sort_column not in valid_columns: sort_column = 'ID'
            sql_query = f"SELECT ID, Nama, Kelas, Alamat FROM Siswa"
            if search_term and search_term.strip() != "":
                sql_query += " WHERE Nama LIKE ?"
                params.append(f'%{search_term}%')
            sql_query += f" ORDER BY {sort_column} {sort_direction}"
            cursor.execute(sql_query, tuple(params))
            for row in cursor.fetchall(): self.tree.insert("", "end", values=list(row))
            cursor.close(); conn.close()

    def cari_siswa(self): self.muat_ulang_data(search_term=self.search_entry.get())
    def tampilkan_semua(self): self.search_entry.delete(0, "end"); self.muat_ulang_data()
    def urutkan_nama_az(self): self.muat_ulang_data(search_term=self.search_entry.get(), sort_column='Nama', sort_direction='ASC')
    def urutkan_nama_za(self): self.muat_ulang_data(search_term=self.search_entry.get(), sort_column='Nama', sort_direction='DESC')
    def urutkan_id_asc(self): self.muat_ulang_data(search_term=self.search_entry.get(), sort_column='ID', sort_direction='ASC')
    def urutkan_id_desc(self): self.muat_ulang_data(search_term=self.search_entry.get(), sort_column='ID', sort_direction='DESC')

    # ... (Sisa fungsi dialog tambah, update, hapus tetap sama)
    def buka_dialog_tambah(self):
        dialog = tk.Toplevel(self.root); dialog.title("Tambah Siswa Baru"); dialog.geometry("350x200")
        dialog.transient(self.root); dialog.grab_set()
        form_tambah = ttk.Frame(dialog); form_tambah.pack(padx=15, pady=15)
        ttk.Label(form_tambah, text="Nama:").grid(row=0, column=0, sticky="w", pady=3)
        nama_entry = ttk.Entry(form_tambah, width=35); nama_entry.grid(row=0, column=1, pady=3)
        ttk.Label(form_tambah, text="Kelas:").grid(row=1, column=0, sticky="w", pady=3)
        kelas_entry = ttk.Entry(form_tambah, width=35); kelas_entry.grid(row=1, column=1, pady=3)
        ttk.Label(form_tambah, text="Alamat:").grid(row=2, column=0, sticky="w", pady=3)
        alamat_entry = ttk.Entry(form_tambah, width=35); alamat_entry.grid(row=2, column=1, pady=3)
        def do_tambah():
            nama = nama_entry.get(); kelas = kelas_entry.get(); alamat = alamat_entry.get()
            if not nama or not kelas or not alamat: messagebox.showwarning("Input Kosong", "Semua kolom harus diisi.", parent=dialog); return
            conn = self.get_connection(DB_SISWA)
            if conn:
                cursor = conn.cursor()
                cursor.execute("INSERT INTO Siswa (Nama, Kelas, Alamat) VALUES (?, ?, ?)", nama, kelas, alamat)
                conn.commit(); cursor.close(); conn.close()
                messagebox.showinfo("Sukses", "Data siswa berhasil ditambahkan.", parent=dialog)
                dialog.destroy(); self.muat_ulang_data()
        ttk.Button(dialog, text="Simpan Data", command=do_tambah).pack(pady=10); nama_entry.focus_set()

    def buka_dialog_update(self):
        conn = self.get_connection(DB_SISWA)
        if not conn: return
        cursor = conn.cursor(); cursor.execute("SELECT ID, Nama, Kelas, Alamat FROM Siswa ORDER BY Nama")
        data_siswa = {f"{row.Nama} (ID: {row.ID})": row for row in cursor.fetchall()}
        cursor.close(); conn.close()
        if not data_siswa: messagebox.showinfo("Informasi", "Tidak ada data siswa untuk di-update."); return
        dialog = tk.Toplevel(self.root); dialog.title("Update Data Siswa"); dialog.geometry("400x250")
        dialog.transient(self.root); dialog.grab_set()
        ttk.Label(dialog, text="1. Pilih Siswa:").pack(padx=10, pady=(10,0))
        combo_siswa = ttk.Combobox(dialog, values=list(data_siswa.keys()), state="readonly", width=40); combo_siswa.pack(padx=10, pady=5)
        form_update = ttk.Frame(dialog); form_update.pack(padx=10, pady=10)
        ttk.Label(form_update, text="Nama:").grid(row=0, column=0, sticky="w", pady=2)
        nama_update_entry = ttk.Entry(form_update, width=35); nama_update_entry.grid(row=0, column=1, pady=2)
        ttk.Label(form_update, text="Kelas:").grid(row=1, column=0, sticky="w", pady=2)
        kelas_update_entry = ttk.Entry(form_update, width=35); kelas_update_entry.grid(row=1, column=1, pady=2)
        ttk.Label(form_update, text="Alamat:").grid(row=2, column=0, sticky="w", pady=2)
        alamat_update_entry = ttk.Entry(form_update, width=35); alamat_update_entry.grid(row=2, column=1, pady=2)
        def on_siswa_select(event):
            siswa_terpilih = data_siswa[combo_siswa.get()]
            nama_update_entry.delete(0, "end"); nama_update_entry.insert(0, siswa_terpilih.Nama)
            kelas_update_entry.delete(0, "end"); kelas_update_entry.insert(0, siswa_terpilih.Kelas)
            alamat_update_entry.delete(0, "end"); alamat_update_entry.insert(0, siswa_terpilih.Alamat)
        combo_siswa.bind("<<ComboboxSelected>>", on_siswa_select)
        def do_update():
            pilihan = combo_siswa.get()
            if not pilihan: messagebox.showwarning("Peringatan", "Anda harus memilih siswa.", parent=dialog); return
            id_siswa = data_siswa[pilihan].ID
            conn_upd = self.get_connection(DB_SISWA)
            if conn_upd:
                cursor_upd = conn_upd.cursor()
                cursor_upd.execute("UPDATE Siswa SET Nama=?, Kelas=?, Alamat=? WHERE ID=?", nama_update_entry.get(), kelas_update_entry.get(), alamat_update_entry.get(), id_siswa)
                conn_upd.commit(); cursor_upd.close(); conn_upd.close()
                messagebox.showinfo("Sukses", "Data berhasil diupdate.", parent=dialog)
                dialog.destroy(); self.muat_ulang_data()
        ttk.Button(dialog, text="Simpan Perubahan", command=do_update).pack(pady=10)

    def buka_dialog_hapus(self):
        conn = self.get_connection(DB_SISWA)
        if not conn: return
        cursor = conn.cursor(); cursor.execute("SELECT ID, Nama FROM Siswa ORDER BY Nama")
        data_siswa = {f"{row.Nama} (ID: {row.ID})": row.ID for row in cursor.fetchall()}
        cursor.close(); conn.close()
        if not data_siswa: messagebox.showinfo("Informasi", "Tidak ada data untuk dihapus."); return
        dialog = tk.Toplevel(self.root); dialog.title("Hapus Data Siswa"); dialog.geometry("350x150")
        dialog.transient(self.root); dialog.grab_set()
        ttk.Label(dialog, text="Pilih Siswa yang akan dihapus:").pack(padx=10, pady=(10,0))
        combo_siswa = ttk.Combobox(dialog, values=list(data_siswa.keys()), state="readonly", width=40); combo_siswa.pack(padx=10, pady=5)
        def do_delete():
            pilihan = combo_siswa.get()
            if not pilihan: messagebox.showwarning("Peringatan", "Anda harus memilih siswa.", parent=dialog); return
            id_siswa = data_siswa[pilihan]
            nama_siswa = pilihan.split(" (ID:")[0]
            if messagebox.askyesno("Konfirmasi Akhir", f"ANDA YAKIN ingin menghapus permanen data:\n\n{nama_siswa}?", icon='warning', parent=dialog):
                conn_del = self.get_connection(DB_SISWA)
                if conn_del:
                    cursor_del = conn_del.cursor()
                    cursor_del.execute("DELETE FROM Siswa WHERE ID=?", id_siswa)
                    conn_del.commit(); cursor_del.close(); conn_del.close()
                    messagebox.showinfo("Sukses", "Data berhasil dihapus.", parent=dialog)
                    dialog.destroy(); self.muat_ulang_data()
        style = ttk.Style(); style.configure("danger.TButton", foreground="red")
        ttk.Button(dialog, text="Hapus Data Ini", command=do_delete, style="danger.TButton").pack(pady=10)

# --- ALUR UTAMA ---
if __name__ == "__main__":
    while True:
        login_root = tk.Tk()
        login_app = LoginWindow(login_root)
        login_root.mainloop()

        if login_app.login_successful:
            main_root = tk.Tk()
            main_app = AplikasiSiswa(main_root)
            main_root.mainloop()
            if main_app.user_logged_out:
                continue
            else:
                break
        else:
            break
