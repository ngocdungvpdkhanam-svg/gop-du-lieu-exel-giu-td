import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Công cụ Gộp File Excel - Tùy chỉnh Tiêu đề")
        self.root.geometry("600x400")

        self.file_paths = []

        # UI Elements
        label = tk.Label(root, text="Công cụ Gộp File Excel", font=("Arial", 16, "bold"))
        label.pack(pady=10)

        btn_select = tk.Button(root, text="Bước 1: Chọn các file Excel", command=self.select_files, bg="#2196F3", fg="white")
        btn_select.pack(pady=5)

        self.listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=70, height=8)
        self.listbox.pack(pady=10, padx=10)

        # Tùy chọn tiêu đề
        self.keep_header_var = tk.BooleanVar(value=True)
        self.check_header = tk.Checkbutton(root, text="Giữ lại hàng tiêu đề (Chỉ lấy từ file đầu tiên)", variable=self.keep_header_var)
        self.check_header.pack(pady=5)

        btn_merge = tk.Button(root, text="Bước 2: Gộp và Lưu File", command=self.merge_files, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        btn_merge.pack(pady=10)

        self.status_label = tk.Label(root, text="Chờ chọn file...", fg="gray")
        self.status_label.pack(pady=5)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn các file Excel để gộp",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if files:
            self.file_paths = list(files)
            self.listbox.delete(0, tk.END)
            for f in self.file_paths:
                self.listbox.insert(tk.END, os.path.basename(f))
            self.status_label.config(text=f"Đã chọn {len(files)} file.", fg="blue")

    def merge_files(self):
        if not self.file_paths:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một file Excel!")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Lưu file gộp tại..."
        )

        if not save_path:
            return

        try:
            combined_data = []
            
            for i, file in enumerate(self.file_paths):
                # Đọc file
                # header=0 nghĩa là lấy hàng đầu tiên làm tiêu đề
                if self.keep_header_var.get():
                    df = pd.read_excel(file)
                else:
                    # Nếu không muốn giữ tiêu đề (đọc dữ liệu thuần túy)
                    df = pd.read_excel(file, header=None)
                
                combined_data.append(df)

            # Gộp các dataframe
            final_df = pd.concat(combined_data, ignore_index=True)

            # Xuất file
            # index=False để không lưu cột số thứ tự của pandas
            final_df.to_excel(save_path, index=False, header=self.keep_header_var.get())

            messagebox.showinfo("Thành công", f"Đã gộp {len(self.file_paths)} file thành công!\nLưu tại: {save_path}")
            self.status_label.config(text="Hoàn thành!", fg="green")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
