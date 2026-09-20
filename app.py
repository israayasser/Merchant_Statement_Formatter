import threading
from pathlib import Path
from tkinter import filedialog, messagebox
import zipfile
import customtkinter as ctk

from converter import process_files

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class MerchantCleanerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Merchant Statement Cleaner")
        self.geometry("1040x700")
        self.minsize(900, 620)
        self.files = []
        self.output_dir = Path.home() / "Downloads"

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, corner_radius=0, fg_color="#0B1118")
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(header, text="Merchant Statement Cleaner", font=ctk.CTkFont(size=24, weight="bold"), text_color="#FFFFFF").grid(row=0, column=0, padx=28, pady=(22, 2), sticky="w")
        ctk.CTkLabel(header, text="Clean PDF-converted Excel statements into a normal, analysis-ready workbook", font=ctk.CTkFont(size=13), text_color="#B9C7D6").grid(row=1, column=0, padx=28, pady=(0, 22), sticky="w")

        body = ctk.CTkFrame(self, fg_color="#EEF3F8", corner_radius=0)
        body.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(2, weight=1)

        top = ctk.CTkFrame(body, fg_color="transparent")
        top.grid(row=0, column=0, padx=28, pady=24, sticky="ew")
        top.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(top, text="Add Excel Files", height=42, command=self.add_files).grid(row=0, column=1, padx=6)
        ctk.CTkButton(top, text="Add ZIP", height=42, command=self.add_zip).grid(row=0, column=2, padx=6)
        ctk.CTkButton(top, text="Clear", height=42, fg_color="#FFFFFF", text_color="#0B1118", hover_color="#DDE6EF", command=self.clear_files).grid(row=0, column=3, padx=6)

        settings = ctk.CTkFrame(body, fg_color="#FFFFFF", corner_radius=14)
        settings.grid(row=1, column=0, padx=28, pady=(0, 18), sticky="ew")
        settings.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(settings, text="Output mode", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=18, pady=16, sticky="w")
        self.mode = ctk.StringVar(value="separate")
        ctk.CTkRadioButton(settings, text="Separate file for each input", variable=self.mode, value="separate").grid(row=0, column=1, padx=8, pady=16, sticky="w")
        ctk.CTkRadioButton(settings, text="Merge all into one", variable=self.mode, value="merge").grid(row=0, column=2, padx=8, pady=16, sticky="w")
        ctk.CTkRadioButton(settings, text="Group files", variable=self.mode, value="group", command=self.toggle_group).grid(row=0, column=3, padx=8, pady=16, sticky="w")
        self.group_entry = ctk.CTkEntry(settings, width=70, placeholder_text="2", state="disabled")
        self.group_entry.grid(row=0, column=4, padx=(4, 18), pady=16)

        card = ctk.CTkFrame(body, fg_color="#FFFFFF", corner_radius=14)
        card.grid(row=2, column=0, padx=28, pady=(0, 18), sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)
        self.status = ctk.CTkLabel(card, text="No files selected", text_color="#607080")
        self.status.grid(row=0, column=0, padx=18, pady=14, sticky="w")
        self.file_box = ctk.CTkTextbox(card, height=240, fg_color="#F7F9FB", border_width=0)
        self.file_box.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        self.file_box.configure(state="disabled")

        footer = ctk.CTkFrame(body, fg_color="transparent")
        footer.grid(row=3, column=0, padx=28, pady=(0, 26), sticky="ew")
        footer.grid_columnconfigure(0, weight=1)
        self.progress = ctk.CTkProgressBar(footer, height=10)
        self.progress.grid(row=0, column=0, padx=(0, 16), sticky="ew")
        self.progress.set(0)
        self.run_btn = ctk.CTkButton(footer, text="Clean & Export", width=180, height=44, command=self.start)
        self.run_btn.grid(row=0, column=1)

    def toggle_group(self):
        self.group_entry.configure(state="normal")

    def add_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xls *.xlsx *.xlsm")])
        if paths:
            self.files.extend(Path(p) for p in paths)
            self.refresh()

    def add_zip(self):
        path = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")])
        if path:
            self.files.append(Path(path))
            self.refresh()

    def clear_files(self):
        self.files.clear()
        self.refresh()
        self.progress.set(0)

    def refresh(self):
        self.file_box.configure(state="normal")
        self.file_box.delete("1.0", "end")
        for p in self.files:
            self.file_box.insert("end", f"• {p.name}\n")
        self.file_box.configure(state="disabled")
        self.status.configure(text=f"{len(self.files)} input file(s) selected")

    def start(self):
        if not self.files:
            messagebox.showwarning("No files", "Please add at least one Excel file or ZIP.")
            return
        mode = self.mode.get()
        group_size = 2
        if mode == "group":
            try:
                group_size = max(1, int(self.group_entry.get() or "2"))
            except ValueError:
                messagebox.showerror("Invalid group size", "Enter a valid number, such as 2 or 5.")
                return
        out = filedialog.askdirectory(initialdir=str(self.output_dir), title="Choose output folder")
        if not out:
            return
        self.output_dir = Path(out)
        self.run_btn.configure(state="disabled")
        self.progress.set(0)
        threading.Thread(target=self.worker, args=(mode, group_size), daemon=True).start()

    def worker(self, mode, group_size):
        try:
            results = process_files(self.files, self.output_dir, mode=mode, group_size=group_size, progress=self.update_progress)
            self.after(0, lambda: self.finished(results))
        except Exception as exc:
            self.after(0, lambda: self.failed(exc))

    def update_progress(self, value, text):
        self.after(0, lambda: (self.progress.set(value), self.status.configure(text=text)))

    def finished(self, results):
        self.run_btn.configure(state="normal")
        self.progress.set(1)
        self.status.configure(text=f"Done — {len(results)} output file(s) created")
        messagebox.showinfo("Completed", f"Cleaning completed successfully.\n\nFiles created: {len(results)}\nFolder: {self.output_dir}")

    def failed(self, exc):
        self.run_btn.configure(state="normal")
        self.status.configure(text="Processing failed")
        messagebox.showerror("Processing error", str(exc))


if __name__ == "__main__":
    MerchantCleanerApp().mainloop()
