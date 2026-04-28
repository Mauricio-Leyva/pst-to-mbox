#!/usr/bin/env python3
"""Graphical interface for converting PST files to Mbox format.

Reuses the logic from pst_to_mbox2.py and allows generating a portable EXE
with PyInstaller for sharing with non-technical users.
"""

from __future__ import annotations

import contextlib
import queue
import threading
import traceback
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pst_to_mbox2 as converter


class QueueWriter:
    """Redirects stdout/stderr to a queue for displaying logs in the GUI."""

    def __init__(self, out_queue: queue.Queue[str]) -> None:
        self.out_queue = out_queue

    def write(self, text: str) -> None:
        if not text:
            return
        self.out_queue.put(text.replace("\r", "\n"))

    def flush(self) -> None:
        return None


class PSTToMboxApp:
    """Main application window for PST to Mbox conversion."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PST a Mbox Portable")
        self.root.geometry("860x600")
        self.root.minsize(760, 540)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.progress_queue: queue.Queue[converter.ProgressInfo] = queue.Queue()
        self.worker: Optional[threading.Thread] = None

        self.pst_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar(value=str(self._default_output_dir()))
        self.max_mb_var = tk.IntVar(value=converter.MAX_MB_DEFAULT)
        self.zip_var = tk.BooleanVar(value=True)

        self.folder_counts: dict[str, int] = {}
        self.total_messages: int = 0

        self._build_ui()
        self.root.after(120, self._poll_log_queue)
        self.root.after(120, self._poll_progress_queue)

    def _build_ui(self) -> None:
        """Builds the user interface widgets."""
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        frame = ttk.Frame(self.root, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(8, weight=1)

        ttk.Label(frame, text="Archivo PST:").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        self.pst_entry = ttk.Entry(frame, textvariable=self.pst_path_var)
        self.pst_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8), padx=(8, 8))
        self.pst_button = ttk.Button(
            frame, text="Examinar...", command=self._select_pst
        )
        self.pst_button.grid(row=0, column=2, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Carpeta de salida:").grid(
            row=1, column=0, sticky="w", pady=(0, 8)
        )
        self.output_entry = ttk.Entry(frame, textvariable=self.output_dir_var)
        self.output_entry.grid(
            row=1, column=1, sticky="ew", pady=(0, 8), padx=(8, 8)
        )
        self.output_button = ttk.Button(
            frame, text="Elegir...", command=self._select_output_dir
        )
        self.output_button.grid(row=1, column=2, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Max MB por archivo:").grid(
            row=2, column=0, sticky="w", pady=(0, 8)
        )
        self.max_mb_spin = ttk.Spinbox(
            frame,
            from_=50,
            to=5000,
            increment=10,
            textvariable=self.max_mb_var,
        )
        self.max_mb_spin.grid(row=2, column=1, sticky="w", pady=(0, 8), padx=(8, 8))

        self.zip_check = ttk.Checkbutton(
            frame,
            text="Comprimir cada Mbox en ZIP",
            variable=self.zip_var,
        )
        self.zip_check.grid(row=3, column=1, sticky="w", pady=(0, 8), padx=(8, 8))

        self.run_button = ttk.Button(
            frame, text="Convertir", command=self._start_conversion
        )
        self.run_button.grid(row=4, column=1, sticky="w", padx=(8, 8), pady=(0, 8))

        ttk.Separator(frame, orient="horizontal").grid(
            row=5, column=0, columnspan=3, sticky="ew", pady=(4, 10)
        )

        self._build_progress_section(frame, row=6)

        ttk.Label(frame, text="Log:").grid(row=7, column=0, sticky="w")

        log_container = ttk.Frame(frame)
        log_container.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_container, wrap="word", height=16)
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(
            log_container, orient="vertical", command=self.log_text.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _build_progress_section(self, parent: ttk.Frame, row: int) -> None:
        """Builds the progress bar section.

        Args:
            parent: Parent frame.
            row: Grid row position.
        """
        progress_frame = ttk.Frame(parent)
        progress_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        progress_frame.columnconfigure(1, weight=1)

        self.progress_label = ttk.Label(progress_frame, text="Listo")
        self.progress_label.grid(row=0, column=0, columnspan=2, sticky="w")

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode="determinate",
            maximum=100,
            value=0,
        )
        self.progress_bar.grid(row=1, column=0, columnspan=2, sticky="ew")

        self.folder_label = ttk.Label(
            progress_frame, text="", font=("TkDefaultFont", 9, "italic")
        )
        self.folder_label.grid(row=2, column=0, columnspan=2, sticky="w")

    def _set_controls_enabled(self, enabled: bool) -> None:
        """Enables or disables all input controls.

        Args:
            enabled: True to enable controls, False to disable.
        """
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
        """Returns the default output directory.

        Returns:
            Path to the default output directory (Downloads/mbox_output).
        """
        downloads_dir = Path.home() / "Downloads"
        if downloads_dir.exists():
            return downloads_dir / "mbox_output"
        return Path.home() / "mbox_output"

    @staticmethod
    def _normalize_out_dir(raw_path: str) -> Path:
        """Normalizes and resolves the output directory path.

        Args:
            raw_path: The raw path string from user input.

        Returns:
            Resolved absolute Path object.
        """
        out_dir = Path(raw_path).expanduser()
        if not out_dir.is_absolute():
            out_dir = (Path.cwd() / out_dir).resolve()
        return out_dir

    @staticmethod
    def _ensure_out_dir_writable(out_dir: Path) -> None:
        """Ensures the output directory is writable.

        Args:
            out_dir: The directory to check.

        Raises:
            Exception: If the directory cannot be created or written to.
        """
        out_dir.mkdir(parents=True, exist_ok=True)

        probe_file = out_dir / ".pst2mbox_write_test.tmp"
        with open(probe_file, "wb") as fh:
            fh.write(b"ok")

        try:
            probe_file.unlink()
        except Exception:
            pass

    def _select_pst(self) -> None:
        """Opens a file dialog to select a PST file."""
        selected = filedialog.askopenfilename(
            title="Seleccionar archivo PST",
            filetypes=[("Outlook PST", "*.pst"), ("Todos los archivos", "*.*")],
        )
        if not selected:
            return
        self.pst_path_var.set(selected)

    def _select_output_dir(self) -> None:
        """Opens a directory dialog to select output directory."""
        selected = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if selected:
            self.output_dir_var.set(selected)

    def _append_log(self, text: str) -> None:
        """Appends text to the log text widget.

        Args:
            text: Text to append.
        """
        self.log_text.insert("end", text)
        self.log_text.see("end")

    def _poll_log_queue(self) -> None:
        """Polls the log queue and updates the log text widget."""
        try:
            while True:
                item = self.log_queue.get_nowait()
                self._append_log(str(item))
        except queue.Empty:
            pass
        finally:
            self.root.after(120, self._poll_log_queue)

    def _poll_progress_queue(self) -> None:
        """Polls the progress queue and updates the progress bar."""
        try:
            while True:
                info = self.progress_queue.get_nowait()
                self._update_progress(info)
        except queue.Empty:
            pass
        finally:
            self.root.after(120, self._poll_progress_queue)

    def _update_progress(self, info: converter.ProgressInfo) -> None:
        """Updates the progress bar based on progress info.

        Args:
            info: Progress information from the converter.
        """
        if info.total > 0:
            percentage = (info.current / info.total) * 100
            self.progress_bar["value"] = percentage
            self.progress_label["text"] = info.message
        else:
            self.progress_bar["value"] = 0
            self.progress_label["text"] = info.message

    def _validate_inputs(self) -> Optional[tuple[str, str, int, bool]]:
        """Validates user inputs.

        Returns:
            Tuple of (pst_path, output_dir, max_mb, do_zip) if valid, None otherwise.
        """
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

    def _start_conversion(self) -> None:
        """Starts the PST to Mbox conversion process."""
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
        self.progress_bar["value"] = 0
        self.progress_label["text"] = "Escaneando archivo PST..."
        self.folder_label["text"] = ""

        self.worker = threading.Thread(
            target=self._run_conversion,
            args=(pst_path, output_dir, max_mb, do_zip),
            daemon=True,
        )
        self.worker.start()

    def _run_conversion(
        self,
        pst_path: str,
        output_dir: str,
        max_mb: int,
        do_zip: bool,
    ) -> None:
        """Runs the conversion in a background thread.

        Args:
            pst_path: Path to the PST file.
            output_dir: Output directory path.
            max_mb: Maximum size per output file in MB.
            do_zip: Whether to compress output as ZIP.
        """
        writer = QueueWriter(self.log_queue)

        def progress_callback(info: converter.ProgressInfo) -> None:
            """Callback to report progress from the converter."""
            self.progress_queue.put(info)

        try:
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                self.folder_counts = converter.scan_pst(
                    pst_path=pst_path,
                    progress_callback=progress_callback,
                )
                self.total_messages = sum(self.folder_counts.values())
                folder_list = ", ".join(
                    f"{name} ({count})" for name, count in self.folder_counts.items()
                )
                self.log_queue.put(f"\nCarpetas detectadas: {len(self.folder_counts)}\n")
                self.log_queue.put(f"Mensajes totales: {self.total_messages}\n")
                self.log_queue.put(f"Detalle: {folder_list}\n\n")

                self.progress_queue.put(converter.ProgressInfo(
                    stage="convert",
                    current=0,
                    total=self.total_messages,
                    message="Iniciando conversion...",
                ))

                converter.pst_to_mbox(
                    pst_path=pst_path,
                    output_dir=output_dir,
                    max_bytes=max_mb * converter.BYTES_PER_MB,
                    do_zip=do_zip,
                    progress_callback=progress_callback,
                    folder_counts=self.folder_counts,
                )
            self.log_queue.put("\nProceso finalizado correctamente.\n")
            self.progress_queue.put(converter.ProgressInfo(
                stage="done",
                current=self.total_messages,
                total=self.total_messages,
                message="Completado",
            ))
        except Exception as exc:
            details = traceback.format_exc()
            self.log_queue.put("\nOcurrio un error durante la conversion.\n")
            if isinstance(exc, (FileNotFoundError, PermissionError)):
                self.log_queue.put(
                    "Sugerencia: cambia la carpeta de salida a una ruta local "
                    "(por ejemplo, Downloads) o autoriza la app en Windows Defender.\n\n"
                )
            self.log_queue.put(details)
            self.progress_queue.put(converter.ProgressInfo(
                stage="error",
                current=0,
                total=self.total_messages,
                message="Error",
            ))

    def _finish_conversion(self) -> None:
        """Called when conversion completes to reset UI state."""
        self._set_controls_enabled(True)


def main() -> None:
    """Main entry point for the application."""
    root = tk.Tk()
    PSTToMboxApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
