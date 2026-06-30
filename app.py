from __future__ import annotations

import platform
import subprocess
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from converter import excel_to_accurate_xml, get_sheet_names, read_excel_rows


APP_TITLE = "Generate Excel to XML Accurate"
DOCUMENT_TYPES = [
    "SALES_INVOICE",
    "SALES_RECEIPT",
    "OTHER_PAYMENT",
]


class AccurateXmlApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1100x720")
        self.minsize(920, 620)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.excel_path: Path | None = None
        self.sheet_names: list[str] = []

        self.file_var = ctk.StringVar(value="Belum ada file dipilih")
        self.sheet_var = ctk.StringVar(value="")
        self.document_type_var = ctk.StringVar(value=DOCUMENT_TYPES[0])
        self.output_var = ctk.StringVar(value="")
        self.status_var = ctk.StringVar(value="Pilih file Excel untuk mulai.")
        self.last_output_path: Path | None = None

        self._build_layout()

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(self, corner_radius=0, fg_color="#f4f6f8")
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        title = ctk.CTkLabel(
            header,
            text=APP_TITLE,
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#1f2937",
        )
        title.grid(row=0, column=0, sticky="w", padx=20, pady=(18, 4))

        subtitle = ctk.CTkLabel(
            header,
            text="Aplikasi desktop Windows untuk membaca Excel dan menghasilkan file XML.",
            text_color="#4b5563",
        )
        subtitle.grid(row=1, column=0, sticky="w", padx=20, pady=(0, 18))

        controls = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=0)
        controls.grid(row=1, column=0, sticky="ew", padx=20, pady=16)
        controls.grid_columnconfigure(1, weight=1)
        controls.grid_columnconfigure(3, weight=1)

        ctk.CTkButton(controls, text="Pilih Excel", command=self.select_excel).grid(
            row=0, column=0, padx=(0, 10), pady=10, sticky="w"
        )
        ctk.CTkLabel(controls, textvariable=self.file_var, anchor="w").grid(
            row=0, column=1, columnspan=3, padx=(0, 10), pady=10, sticky="ew"
        )

        ctk.CTkLabel(controls, text="Sheet").grid(row=1, column=0, sticky="w", pady=10)
        self.sheet_menu = ctk.CTkOptionMenu(
            controls,
            variable=self.sheet_var,
            values=[""],
            command=lambda _: self.load_preview(),
        )
        self.sheet_menu.grid(row=1, column=1, sticky="ew", padx=(0, 18), pady=10)

        ctk.CTkLabel(controls, text="Jenis XML").grid(row=1, column=2, sticky="w", pady=10)
        self.document_type_menu = ctk.CTkOptionMenu(
            controls,
            variable=self.document_type_var,
            values=DOCUMENT_TYPES,
        )
        self.document_type_menu.grid(row=1, column=3, sticky="ew", padx=(10, 0), pady=10)

        ctk.CTkButton(controls, text="Download XML / Save As", command=self.select_output).grid(
            row=2, column=0, padx=(0, 10), pady=10, sticky="w"
        )
        ctk.CTkEntry(controls, textvariable=self.output_var).grid(
            row=2, column=1, padx=(0, 10), pady=10, sticky="ew"
        )
        ctk.CTkButton(controls, text="Buka Folder", command=self.open_output_folder).grid(
            row=2, column=2, padx=(0, 10), pady=10, sticky="ew"
        )
        ctk.CTkButton(controls, text="Export XML", command=self.generate_xml).grid(
            row=2, column=3, padx=(10, 0), pady=10, sticky="ew"
        )

        preview_frame = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=0)
        preview_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 12))
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            preview_frame,
            text="Preview Data",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=2, pady=(0, 8))

        table_holder = ctk.CTkFrame(preview_frame, fg_color="#ffffff", corner_radius=0)
        table_holder.grid(row=1, column=0, sticky="nsew")
        table_holder.grid_columnconfigure(0, weight=1)
        table_holder.grid_rowconfigure(0, weight=1)

        self.preview_table = ttk.Treeview(table_holder, show="headings", height=16)
        self.preview_table.grid(row=0, column=0, sticky="nsew")

        vertical_scroll = ttk.Scrollbar(table_holder, orient="vertical", command=self.preview_table.yview)
        vertical_scroll.grid(row=0, column=1, sticky="ns")
        horizontal_scroll = ttk.Scrollbar(table_holder, orient="horizontal", command=self.preview_table.xview)
        horizontal_scroll.grid(row=1, column=0, sticky="ew")
        self.preview_table.configure(yscrollcommand=vertical_scroll.set, xscrollcommand=horizontal_scroll.set)

        footer = ctk.CTkFrame(self, corner_radius=0, fg_color="#f4f6f8")
        footer.grid(row=3, column=0, sticky="ew")
        ctk.CTkLabel(footer, textvariable=self.status_var, anchor="w", text_color="#374151").grid(
            row=0, column=0, sticky="ew", padx=20, pady=10
        )

    def select_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Pilih file Excel",
            filetypes=[
                ("Excel files", "*.xls *.xlsx *.xlsm *.xltx *.xltm"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        self.excel_path = Path(path)
        self.file_var.set(str(self.excel_path))
        self.output_var.set(str(self._default_output_path()))

        try:
            self.sheet_names = get_sheet_names(self.excel_path)
        except Exception as exc:
            messagebox.showerror("Gagal membaca Excel", str(exc))
            self.status_var.set("Gagal membaca file Excel.")
            return

        self.sheet_var.set(self.sheet_names[0] if self.sheet_names else "")
        self.sheet_menu.configure(values=self.sheet_names or [""])
        self.load_preview()

    def select_output(self) -> str | None:
        default_path = self._default_output_path()

        path = filedialog.asksaveasfilename(
            title="Simpan XML",
            defaultextension=".xml",
            initialdir=str(default_path.parent),
            initialfile=default_path.name,
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        if path:
            self.output_var.set(path)
            return path
        return None

    def load_preview(self) -> None:
        if not self.excel_path:
            return

        try:
            headers, rows = read_excel_rows(self.excel_path, sheet_name=self.sheet_var.get(), max_rows=50)
        except Exception as exc:
            messagebox.showerror("Gagal preview", str(exc))
            self.status_var.set("Gagal menampilkan preview.")
            return

        self.preview_table.delete(*self.preview_table.get_children())
        self.preview_table["columns"] = headers

        for header in headers:
            self.preview_table.heading(header, text=header)
            self.preview_table.column(header, width=150, minwidth=90, stretch=False)

        for row in rows:
            values = ["" if row.get(header) is None else row.get(header) for header in headers]
            self.preview_table.insert("", "end", values=values)

        self.status_var.set(f"Preview {len(rows)} baris pertama dari sheet {self.sheet_var.get()}.")

    def generate_xml(self) -> None:
        if not self.excel_path:
            messagebox.showwarning("File belum dipilih", "Pilih file Excel terlebih dahulu.")
            return

        output_path = self.output_var.get().strip()
        if not output_path:
            output_path = self.select_output()
            if not output_path:
                return

        try:
            total = excel_to_accurate_xml(
                self.excel_path,
                output_path,
                sheet_name=self.sheet_var.get(),
                document_type=self.document_type_var.get(),
            )
        except Exception as exc:
            messagebox.showerror("Generate gagal", str(exc))
            self.status_var.set("Generate XML gagal.")
            return

        self.last_output_path = Path(output_path)
        self.status_var.set(f"Berhasil generate {total} baris ke {output_path}")
        should_open = messagebox.askyesno(
            "Selesai",
            f"XML berhasil dibuat:\n{output_path}\n\nBuka folder hasil export sekarang?",
        )
        if should_open:
            self.open_output_folder()

    def open_output_folder(self) -> None:
        output_text = self.output_var.get().strip()
        folder = Path(output_text).expanduser().parent if output_text else self._downloads_dir()
        folder.mkdir(parents=True, exist_ok=True)

        try:
            system = platform.system()
            if system == "Darwin":
                subprocess.run(["open", str(folder)], check=False)
            elif system == "Windows":
                subprocess.run(["explorer", str(folder)], check=False)
            else:
                subprocess.run(["xdg-open", str(folder)], check=False)
        except Exception as exc:
            messagebox.showerror("Gagal buka folder", str(exc))

    def _default_output_path(self) -> Path:
        filename = "hasil.xml"
        if self.excel_path:
            filename = f"{self.excel_path.stem}.xml"
        return self._downloads_dir() / filename

    def _downloads_dir(self) -> Path:
        downloads = Path.home() / "Downloads"
        return downloads if downloads.exists() else Path.home()


if __name__ == "__main__":
    app = AccurateXmlApp()
    app.mainloop()
