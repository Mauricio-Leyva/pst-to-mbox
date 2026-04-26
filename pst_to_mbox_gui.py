#!/usr/bin/env python3
"""
Interfaz grafica simple para convertir archivos PST a Mbox.

Reutiliza la logica de pst_to_mbox2.py y permite generar un EXE portable
con PyInstaller para compartir con usuarios no tecnicos.
"""

import contextlib
import queue
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pst_to_mbox2 as converter


class QueueWriter:
    """Redirige stdout/stderr a una cola para mostrar logs en la GUI."""

    def __init__(self, out_queue: queue.Queue):
        self.out_queue = out_queue

    def write(self, text: str):
        if not text:
            return
        self.out_queue.put(text.replace("\r", "\n"))

    def flush(self):
        return None


class PSTToMboxApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PST a Mbox Portable")
        self.root.geometry("860x560")
        self.root.minsize(760, 500)

        self.log_queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None

        self.pst_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar(value=str(self._default_output_dir()))
        self.max_mb_var = tk.IntVar(value=converter.MAX_MB_DEFAULT)
        self.zip_var = tk.BooleanVar(value=True)

        self._build_ui()
        self.root.after(120, self._poll_log_queue)

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        frame = ttk.Frame(self.root, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(7, weight=1)

        ttk.Label(frame, text="Archivo PST:").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.pst_entry = ttk.Entry(frame, textvariable=self.pst_path_var)
        self.pst_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8), padx=(8, 8))
        self.pst_button = ttk.Button(frame, text="Examinar...", command=self._select_pst)
        self.pst_button.grid(row=0, column=2, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Carpeta de salida:").grid(row=1, column=0, sticky="w", pady=(0, 8))
        self.output_entry = ttk.Entry(frame, textvariable=self.output_dir_var)
        self.output_entry.grid(row=1, column=1, sticky="ew", pady=(0, 8), padx=(8, 8))
        self.output_button = ttk.Button(frame, text="Elegir...", command=self._select_output_dir)
        self.output_button.grid(row=1, column=2, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Max MB por archivo:").grid(row=2, column=0, sticky="w", pady=(0, 8))
        self.max_mb_spin = ttk.Spinbox(frame, from_=50, to=5000, increment=10, textvariable=self.max_mb_var)
        self.max_mb_spin.grid(row=2, column=1, sticky="w", pady=(0, 8), padx=(8, 8))

        self.zip_check = ttk.Checkbutton(frame, text="Comprimir cada Mbox en ZIP", variable=self.zip_var)
        self.zip_check.grid(row=3, column=1, sticky="w", pady=(0, 8), padx=(8, 8))

        self.run_button = ttk.Button(frame, text="Convertir", command=self._start_conversion)
        self.run_button.grid(row=4, column=1, sticky="w", padx=(8, 8), pady=(0, 8))

        ttk.Separator(frame, orient="horizontal").grid(row=5, column=0, columnspan=3, sticky="ew", pady=(4, 10))

        ttk.Label(frame, text="Log:").grid(row=6, column=0, sticky="w")

        log_container = ttk.Frame(frame)
        log_container.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_container, wrap="word", height=16)
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(log_container, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _set_controls_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for widget in (
            self.pst_entry,
            self.pst_button,
            self.output_entry,
            self.output_button,
            self.max_mb_spin,
            self.zip_check,
            self.run_button,
        ):
            widget.configure(state=state)

    @staticmethod
    def _default_output_dir() -> Path:
        downloads_dir = Path.home() / "Downloads"
        if downloads_dir.exists():
            return downloads_dir / "mbox_output"
        return Path.home() / "mbox_output"

    @staticmethod
    def _normalize_out_dir(raw_path: str) -> Path:
        out_dir = Path(raw_path).expanduser()
        if not out_dir.is_absolute():
            out_dir = (Path.cwd() / out_dir).resolve()
        return out_dir

    @staticmethod
    def _ensure_out_dir_writable(out_dir: Path):
        out_dir.mkdir(parents=True, exist_ok=True)

        probe_file = out_dir / ".pst2mbox_write_test.tmp"
        with open(probe_file, "wb") as fh:
            fh.write(b"ok")

        try:
            probe_file.unlink()
        except Exception:
            pass

    def _select_pst(self):
        selected = filedialog.askopenfilename(
            title="Seleccionar archivo PST",
            filetypes=[("Outlook PST", "*.pst"), ("Todos los archivos", "*.*")],
        )
        if not selected:
            return
        self.pst_path_var.set(selected)

    def _select_output_dir(self):
        selected = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if selected:
            self.output_dir_var.set(selected)

    def _append_log(self, text: str):
        self.log_text.insert("end", text)
        self.log_text.see("end")

    def _poll_log_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and item and item[0] == "__done__":
                    success = bool(item[1])
                    error_details = item[2]
                    self._finish_conversion(success, error_details)
                    continue

                self._append_log(str(item))
        except queue.Empty:
            pass
        finally:
            self.root.after(120, self._poll_log_queue)

    def _validate_inputs(self):
        pst_value = self.pst_path_var.get().strip().strip('"').strip("'")
        out_value = self.output_dir_var.get().strip().strip('"').strip("'")

        if not pst_value:
            messagebox.showerror("Falta archivo", "Selecciona un archivo .pst")
            return None

        pst_path = Path(pst_value)
        if not pst_path.exists() or not pst_path.is_file():
            messagebox.showerror("Archivo invalido", f"No existe:\n{pst_path}")
            return None

        if pst_path.suffix.lower() != ".pst":
            ask = messagebox.askyesno(
                "Extension no esperada",
                "El archivo no termina en .pst.\n\nDeseas continuar de todas formas?",
            )
            if not ask:
                return None

        if not out_value:
            messagebox.showerror("Falta carpeta", "Selecciona una carpeta de salida")
            return None

        try:
            out_dir = self._normalize_out_dir(out_value)
            self._ensure_out_dir_writable(out_dir)
        except Exception as exc:
            messagebox.showerror(
                "Carpeta sin acceso",
                "No se puede crear o escribir en la carpeta de salida.\n\n"
                "Prueba con una carpeta local como:\n"
                "  C:\\Users\\<usuario>\\Downloads\\mbox_output\n\n"
                "Si esta activa la proteccion de carpetas de Windows Defender,\n"
                "autoriza esta app o elige otra carpeta.\n\n"
                f"Detalle tecnico:\n{exc}",
            )
            return None

        max_mb = self.max_mb_var.get()
        if max_mb <= 0:
            messagebox.showerror("Valor invalido", "El maximo de MB debe ser mayor a 0")
            return None

        return str(pst_path), str(out_dir), int(max_mb), bool(self.zip_var.get())

    def _start_conversion(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("En progreso", "Ya hay una conversion en ejecucion")
            return

        payload = self._validate_inputs()
        if payload is None:
            return

        pst_path, output_dir, max_mb, do_zip = payload

        self.log_text.delete("1.0", "end")
        self._append_log("Iniciando conversion...\n")
        self._append_log(f"PST: {pst_path}\n")
        self._append_log(f"Salida: {output_dir}\n")
        self._append_log(f"Max MB: {max_mb}\n")
        self._append_log(f"ZIP: {'Si' if do_zip else 'No'}\n\n")

        self._set_controls_enabled(False)
        self.worker = threading.Thread(
            target=self._run_conversion,
            args=(pst_path, output_dir, max_mb * converter.BYTES_PER_MB, do_zip),
            daemon=True,
        )
        self.worker.start()

    def _run_conversion(self, pst_path: str, output_dir: str, max_bytes: int, do_zip: bool):
        writer = QueueWriter(self.log_queue)
        try:
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                converter.pst_to_mbox(
                    pst_path=pst_path,
                    output_dir=output_dir,
                    max_bytes=max_bytes,
                    do_zip=do_zip,
                )
            self.log_queue.put("\nProceso finalizado correctamente.\n")
            self.log_queue.put(("__done__", True, ""))
        except Exception as exc:
            details = traceback.format_exc()
            self.log_queue.put("\nOcurrio un error durante la conversion.\n")
            if isinstance(exc, (FileNotFoundError, PermissionError)):
                self.log_queue.put(
                    "Sugerencia: cambia la carpeta de salida a una ruta local (por ejemplo, Downloads) "
                    "o autoriza la app en Windows Defender.\n\n"
                )
            self.log_queue.put(details)
            self.log_queue.put(("__done__", False, details))

    def _finish_conversion(self, success: bool, error_details: str):
        self._set_controls_enabled(True)
        if success:
            messagebox.showinfo("Listo", "Conversion completada")
        else:
            short_error = "Se produjo un error. Revisa el log para detalles."
            if error_details:
                short_error = short_error + "\n\nUltima linea:\n" + error_details.strip().splitlines()[-1]
            messagebox.showerror("Error", short_error)


def main():
    root = tk.Tk()
    PSTToMboxApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
