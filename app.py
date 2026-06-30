from __future__ import annotations

import platform
import subprocess
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from converter import excel_to_accurate_xml, get_sheet_names, read_excel_rows


APP_TITLE = "Generate Excel to XML Accurate"
WINDOW_BG = "#eef2f6"
PANEL_BG = "#ffffff"
TEXT_MAIN = "#17202a"
TEXT_MUTED = "#667085"
ACCENT = "#2563eb"
ACCENT_DARK = "#1d4ed8"
SUCCESS = "#0f766e"
DOCUMENT_TYPES = [
    "SALES_INVOICE",
    "SALES_RECEIPT",
    "OTHER_PAYMENT",
]


class AccurateXmlApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x760")
        self.minsize(980, 640)
        self.configure(fg_color=WINDOW_BG)

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
        self._configure_tree_style()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, corner_radius=0, fg_color=WINDOW_BG)
        header.grid(row=0, column=0, sticky="ew", padx=28, pady=(22, 10))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="Accurate XML Generator",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=TEXT_MAIN,
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            header,
            text="Convert Excel ke XML Accurate untuk Sales Invoice, Sales Receipt, dan Other Payment.",
            font=ctk.CTkFont(size=13),
            text_color=TEXT_MUTED,
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        app_body = ctk.CTkFrame(self, fg_color="transparent")
        app_body.grid(row=1, column=0, sticky="nsew", padx=28, pady=(8, 20))
        app_body.grid_columnconfigure(0, minsize=360)
        app_body.grid_columnconfigure(1, weight=1)
        app_body.grid_rowconfigure(0, weight=1)

        controls = ctk.CTkFrame(app_body, fg_color=PANEL_BG, corner_radius=8, border_width=1, border_color="#d9e2ec")
        controls.grid(row=0, column=0, sticky="nsw", padx=(0, 18))
        controls.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            controls,
            text="Setup Export",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=TEXT_MAIN,
        ).grid(row=0, column=0, sticky="w", padx=22, pady=(22, 4))

        ctk.CTkLabel(
            controls,
            text="Pilih sumber data dan format XML.",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED,
        ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 18))

        ctk.CTkButton(
            controls,
            text="Pilih File Excel",
            height=42,
            corner_radius=6,
            fg_color=ACCENT,
            hover_color=ACCENT_DARK,
            command=self.select_excel,
        ).grid(row=2, column=0, padx=22, pady=(0, 12), sticky="ew")

        file_box = ctk.CTkFrame(controls, fg_color="#f8fafc", corner_radius=6, border_width=1, border_color="#e2e8f0")
        file_box.grid(row=3, column=0, padx=22, pady=(0, 18), sticky="ew")
        file_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            file_box,
            textvariable=self.file_var,
            anchor="w",
            justify="left",
            wraplength=300,
            font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED,
        ).grid(row=0, column=0, padx=12, pady=12, sticky="ew")

        ctk.CTkLabel(controls, text="Sheet", text_color=TEXT_MAIN, font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=4, column=0, sticky="w", padx=22, pady=(0, 6)
        )
        self.sheet_menu = ctk.CTkOptionMenu(
            controls,
            variable=self.sheet_var,
            values=[""],
            height=38,
            corner_radius=6,
            fg_color="#f8fafc",
            button_color="#dbeafe",
            button_hover_color="#bfdbfe",
            text_color=TEXT_MAIN,
            command=lambda _: self.load_preview(),
        )
        self.sheet_menu.grid(row=5, column=0, sticky="ew", padx=22, pady=(0, 16))

        ctk.CTkLabel(controls, text="Jenis XML", text_color=TEXT_MAIN, font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=6, column=0, sticky="w", padx=22, pady=(0, 6)
        )
        self.document_type_menu = ctk.CTkOptionMenu(
            controls,
            variable=self.document_type_var,
            values=DOCUMENT_TYPES,
            height=38,
            corner_radius=6,
            fg_color="#f8fafc",
            button_color="#dbeafe",
            button_hover_color="#bfdbfe",
            text_color=TEXT_MAIN,
        )
        self.document_type_menu.grid(row=7, column=0, sticky="ew", padx=22, pady=(0, 16))

        ctk.CTkLabel(controls, text="Output XML", text_color=TEXT_MAIN, font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=8, column=0, sticky="w", padx=22, pady=(0, 6)
        )
        ctk.CTkEntry(
            controls,
            textvariable=self.output_var,
            height=38,
            corner_radius=6,
            border_color="#cbd5e1",
            fg_color="#f8fafc",
        ).grid(row=9, column=0, padx=22, pady=(0, 10), sticky="ew")

        ctk.CTkButton(
            controls,
            text="Save As",
            height=38,
            corner_radius=6,
            fg_color="#475569",
            hover_color="#334155",
            command=self.select_output,
        ).grid(row=10, column=0, padx=22, pady=(0, 10), sticky="ew")

        ctk.CTkButton(
            controls,
            text="Export XML",
            height=46,
            corner_radius=6,
            fg_color=SUCCESS,
            hover_color="#115e59",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.generate_xml,
        ).grid(row=11, column=0, padx=22, pady=(2, 10), sticky="ew")

        ctk.CTkButton(
            controls,
            text="Buka Folder Output",
            height=38,
            corner_radius=6,
            fg_color="#e2e8f0",
            hover_color="#cbd5e1",
            text_color=TEXT_MAIN,
            command=self.open_output_folder,
        ).grid(row=12, column=0, padx=22, pady=(0, 22), sticky="ew")

        status_card = ctk.CTkFrame(controls, fg_color="#eff6ff", corner_radius=6, border_width=1, border_color="#bfdbfe")
        status_card.grid(row=13, column=0, padx=22, pady=(0, 22), sticky="ew")
        status_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            status_card,
            textvariable=self.status_var,
            anchor="w",
            justify="left",
            wraplength=300,
            font=ctk.CTkFont(size=12),
            text_color="#1e3a8a",
        ).grid(row=0, column=0, padx=12, pady=12, sticky="ew")

        preview_frame = ctk.CTkFrame(app_body, fg_color=PANEL_BG, corner_radius=8, border_width=1, border_color="#d9e2ec")
        preview_frame.grid(row=0, column=1, sticky="nsew")
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            preview_frame,
            text="Preview Data",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=TEXT_MAIN,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=22, pady=(22, 4))

        ctk.CTkLabel(
            preview_frame,
            text="Menampilkan maksimal 50 baris pertama dari sheet yang dipilih.",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 14))

        table_holder = ctk.CTkFrame(preview_frame, fg_color="#ffffff", corner_radius=6, border_width=1, border_color="#e2e8f0")
        table_holder.grid(row=2, column=0, sticky="nsew", padx=22, pady=(0, 22))
        table_holder.grid_columnconfigure(0, weight=1)
        table_holder.grid_rowconfigure(0, weight=1)

        self.preview_table = ttk.Treeview(table_holder, show="headings", height=18, style="Modern.Treeview")
        self.preview_table.grid(row=0, column=0, sticky="nsew")

        vertical_scroll = ttk.Scrollbar(table_holder, orient="vertical", command=self.preview_table.yview)
        vertical_scroll.grid(row=0, column=1, sticky="ns")
        horizontal_scroll = ttk.Scrollbar(table_holder, orient="horizontal", command=self.preview_table.xview)
        horizontal_scroll.grid(row=1, column=0, sticky="ew")
        self.preview_table.configure(yscrollcommand=vertical_scroll.set, xscrollcommand=horizontal_scroll.set)

    def _configure_tree_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Modern.Treeview",
            background="#ffffff",
            foreground=TEXT_MAIN,
            fieldbackground="#ffffff",
            borderwidth=0,
            rowheight=32,
            font=("Arial", 12),
        )
        style.configure(
            "Modern.Treeview.Heading",
            background="#f1f5f9",
            foreground="#334155",
            borderwidth=0,
            relief="flat",
            font=("Arial", 12, "bold"),
        )
        style.map("Modern.Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", TEXT_MAIN)])

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
