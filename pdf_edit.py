import customtkinter as ctk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter
import os

# 設定佈景主題
ctk.set_appearance_mode("dark")  # "dark" 或 "light"
ctk.set_default_color_theme("blue") # "blue", "green", "dark-blue"


class ModernPDFApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        

        self.title("PDF Master Pro")
        self.geometry("700x550")
        
        self.selected_files = []

        # 設定 Grid 版面
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- 左側側邊欄 (Sidebar) ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        self.logo_label = ctk.CTkLabel(self.sidebar, text="PDF mini Toolbox", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.pack(pady=20, padx=20)

        self.add_btn = ctk.CTkButton(self.sidebar, text="＋ Add Files", command=self.add_files, fg_color="#2ecc71", hover_color="#27ae60")
        self.add_btn.pack(pady=10, padx=20)

        self.clear_btn = ctk.CTkButton(self.sidebar, text="🗑 Empty List", command=self.clear_files, fg_color="#e74c3c", hover_color="#c0392b")
        self.clear_btn.pack(pady=10, padx=20)

        self.appearance_label = ctk.CTkLabel(self.sidebar, text="Visualization", anchor="w")
        self.appearance_label.pack(side="bottom", padx=20, pady=(0, 5))
        self.appearance_option = ctk.CTkOptionMenu(self.sidebar, values=["Dark", "Light", "System"], command=self.change_appearance)
        self.appearance_option.pack(side="bottom", padx=20, pady=(0, 20))

        # --- 右側主面板 (Main Content) ---
        self.main_content = ctk.CTkFrame(self, corner_radius=15, fg_color="transparent")
        self.main_content.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        # 檔案列表顯示區
        self.file_list_label = ctk.CTkLabel(self.main_content, text="Waiting List", font=ctk.CTkFont(size=14))
        self.file_list_label.pack(anchor="w", pady=(0, 10))
        
        self.file_display = ctk.CTkTextbox(self.main_content, height=150)
        self.file_display.pack(fill="x", pady=(0, 20))
        self.file_display.insert("0.0", "There is no file...")

        # 頁數設定區
        self.page_frame = ctk.CTkFrame(self.main_content)
        self.page_frame.pack(fill="x", pady=(0, 20), padx=5)
        
        self.page_label = ctk.CTkLabel(self.page_frame, text="Setting pages (ex: 1, 3-5):", font=ctk.CTkFont(size=12))
        self.page_label.pack(side="left", padx=15, pady=15)
        
        self.page_entry = ctk.CTkEntry(self.page_frame, placeholder_text="Enter the page", width=200)
        self.page_entry.pack(side="right", padx=15, pady=15, fill="x", expand=True)

        # 功能按鈕區 (Grid 佈局)
        self.action_frame = ctk.CTkFrame(self.main_content, fg_color="transparent")
        self.action_frame.pack(fill="both", expand=True)
        
        buttons = [
            ("Combine all pdfs", self.run_merge, "#3498db"),
            ("Remove the pages", self.run_remove, "#95a5a6"),
            ("Extract the pages", self.run_extract, "#9b59b6"),
            ("Extract and save individually", self.run_extract_split, "#f1c40f"), 
            ("Resize to Standard A4", self.run_resize_to_a4, "#1abc9c")
        ]

        for i, (text, cmd, color) in enumerate(buttons):
            btn = ctk.CTkButton(self.action_frame, text=text, command=cmd, 
                                 fg_color=color, hover_color=self.darken_color(color),
                                 height=45, font=ctk.CTkFont(weight="bold"))
            btn.grid(row=i//2, column=i%2, padx=10, pady=10, sticky="nsew")
        
        self.action_frame.grid_columnconfigure((0, 1), weight=1)

    # --- 功能邏輯 (與之前相同但更新介面) ---

    def darken_color(self, hex_color):
        # 簡單的顏色加深邏輯（模擬 Hover 效果）
        return hex_color # 這裡可以寫更複雜的運算，或是直接手動設定

    def change_appearance(self, mode):
        ctk.set_appearance_mode(mode)

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            self.selected_files.extend(files)
            self.update_file_display()

    def clear_files(self):
        self.selected_files = []
        self.update_file_display()

    def update_file_display(self):
        self.file_display.delete("0.0", "end")
        if not self.selected_files:
            self.file_display.insert("0.0", "There is no file...")
        else:
            display_text = "\n".join([f"{i+1}. {os.path.basename(f)}" for i, f in enumerate(self.selected_files)])
            self.file_display.insert("0.0", display_text)

    def parse_pages(self):
        raw = self.page_entry.get().replace(" ", "")
        if not raw: return None
        pages = set()
        try:
            for part in raw.split(','):
                if '-' in part:
                    s, e = map(int, part.split('-'))
                    pages.update(range(s, e + 1))
                else: pages.add(int(part))
            return sorted(list(pages))
        except: return None

    # (合併、移除、提取的邏輯函數與之前版本一致，只需保留即可)
    # 為節省長度，這裡縮減 run_ 開頭的邏輯，直接沿用舊版邏輯即可
    def run_merge(self):
        if len(self.selected_files) < 2: return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not save_path: return
        writer = PdfWriter()
        for p in self.selected_files:
            for pg in PdfReader(p).pages: writer.add_page(pg)
        with open(save_path, "wb") as f: writer.write(f)
        messagebox.showinfo("Finish", "Combine sucessfully!")

    def run_remove(self):
        if not self.selected_files: return
        p_list = self.parse_pages()
        if p_list is None: return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not save_path: return
        reader = PdfReader(self.selected_files[0])
        writer = PdfWriter()
        remove_idx = [p-1 for p in p_list]
        for i, pg in enumerate(reader.pages):
            if i not in remove_idx: writer.add_page(pg)
        with open(save_path, "wb") as f: writer.write(f)
        messagebox.showinfo("Finish", "Remove sucessfully!")

    def run_extract(self):
        if not self.selected_files: return
        p_list = self.parse_pages()
        if p_list is None: return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not save_path: return
        reader = PdfReader(self.selected_files[0])
        writer = PdfWriter()
        for p in p_list:
            if 0 <= p-1 < len(reader.pages): writer.add_page(reader.pages[p-1])
        with open(save_path, "wb") as f: writer.write(f)
        messagebox.showinfo("Finish", "Extract sucessfully!")

    def run_extract_split(self):
        if not self.selected_files: return
        p_list = self.parse_pages()
        if p_list is None: return
        target_dir = filedialog.askdirectory()
        if not target_dir: return
        reader = PdfReader(self.selected_files[0])
        base = os.path.splitext(os.path.basename(self.selected_files[0]))[0]
        for p in p_list:
            if 0 <= p-1 < len(reader.pages):
                writer = PdfWriter()
                writer.add_page(reader.pages[p-1])
                with open(f"{target_dir}/{base}_page_{p}.pdf", "wb") as f: writer.write(f)
        messagebox.showinfo("Finish", "Save one by one sucessfully!")
    
    def run_resize_to_a4(self):
        if not self.selected_files:
            messagebox.showwarning("Warning", "Please add files first!")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not save_path: return

        # A4 標準尺寸 (Points)
        A4_W = 595
        A4_H = 842

        writer = PdfWriter()
        
        # 這裡我們處理第一個選中的檔案
        reader = PdfReader(self.selected_files[0])
        
        for page in reader.pages:
            # 獲取原始大小
            orig_w = float(page.mediabox.width)
            orig_h = float(page.mediabox.height)
            
            # 計算縮放比例 (取寬高縮放中較小的那個，以確保內容完全放入)
            scale_factor = min(A4_W / orig_w, A4_H / orig_h)
            
            # 進行縮放
            page.scale_by(scale_factor)
            
            # 建立一個新的 A4 頁面並將縮放後的內容放上去
            # 或是直接強制修改 mediabox (最直接的方法)
            page.mediabox.lower_left = (0, 0)
            page.mediabox.upper_right = (A4_W, A4_H)
            
            writer.add_page(page)

        with open(save_path, "wb") as f:
            writer.write(f)
        
        messagebox.showinfo("Success", "All pages resized to standard A4!")

if __name__ == "__main__":
    app = ModernPDFApp()
    app.mainloop()