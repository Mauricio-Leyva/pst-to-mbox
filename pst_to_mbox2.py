#!/usr/bin/env python3
"""
pst_to_mbox.py - Convierte .pst de Outlook a Mbox por carpeta para Roundcube
Incluye: cuerpo HTML, imagenes embebidas (inline) y adjuntos (pdf, docx, etc.)

INSTALACION (Windows):
    pip install pywin32

Uso:
    python pst_to_mbox.py correo.pst
    python pst_to_mbox.py correo.pst --zip
    python pst_to_mbox.py correo.pst --max-mb 400 --zip
"""

import argparse
import base64
import mailbox
import mimetypes
import os
import sys
import tempfile
import zipfile
import email
import email.utils
import email.message
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.mime.image
import email.encoders
from pathlib import Path
from datetime import datetime

BYTES_PER_MB   = 1024 * 1024
MAX_MB_DEFAULT = 490

FOLDER_NAMES = {
    "inbox":         "Entrada",
    "sent items":    "Enviados",
    "sent":          "Enviados",
    "deleted items": "Eliminados",
    "drafts":        "Borradores",
    "junk email":    "Spam",
    "junk":          "Spam",
    "outbox":        "Bandeja de salida",
    "archive":       "Archivo",
}


def normalize_folder_name(raw: str) -> str:
    clean = raw.strip().lower()
    clean = FOLDER_NAMES.get(clean, raw.strip())
    for ch in r'\/:*?"<>|':
        clean = clean.replace(ch, "_")
    return clean


# ─────────────────────────────────────────────────────────────────────
#  BACKEND: Outlook COM
# ─────────────────────────────────────────────────────────────────────
def outlook_available():
    try:
        import win32com.client  # noqa
        return True
    except ImportError:
        return False


def iter_messages_outlook(pst_path: str):
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    pst_path = str(Path(pst_path).resolve())
    print(f"  Ruta absoluta: {pst_path}")

    for i in range(1, ns.Folders.Count + 1):
        try:
            f = ns.Folders.Item(i)
            if pst_path.lower() in (f.FolderPath or "").lower():
                try:
                    ns.RemoveStore(f)
                except Exception:
                    pass
                break
        except Exception:
            pass

    ns.AddStoreEx(pst_path, 3)
    pst_root = ns.Folders.Item(ns.Folders.Count)

    def walk(folder):
        folder_name = normalize_folder_name(folder.Name)
        items = folder.Items
        for idx in range(1, items.Count + 1):
            try:
                item = items.Item(idx)
                if item.Class != 43:
                    continue
                yield folder_name, build_message(item, folder_name)
            except Exception as exc:
                print(f"  SKIP item {idx} en '{folder_name}': {exc}")
        for j in range(1, folder.Folders.Count + 1):
            try:
                yield from walk(folder.Folders.Item(j))
            except Exception as exc:
                print(f"  SKIP subcarpeta: {exc}")

    yield from walk(pst_root)

    try:
        ns.RemoveStore(pst_root)
    except Exception:
        pass
    pythoncom.CoUninitialize()


def build_message(item, folder_name: str) -> email.message.Message:
    """
    Construye un mensaje MIME completo con:
    - Cuerpo HTML (o texto plano)
    - Imagenes embebidas como partes inline (Content-ID)
    - Adjuntos como partes attachment
    """
    # ── Recopilar adjuntos via COM ─────────────────────────────────
    inline_parts = []   # (cid, filename, data, mime_type)
    attach_parts = []   # (filename, data, mime_type)

    try:
        att_count = item.Attachments.Count
    except Exception:
        att_count = 0

    with tempfile.TemporaryDirectory() as tmpdir:
        for i in range(1, att_count + 1):
            try:
                att  = item.Attachments.Item(i)
                fname = att.FileName or f"adjunto_{i}"
                tmp_path = os.path.join(tmpdir, fname)
                att.SaveAsFile(tmp_path)

                with open(tmp_path, "rb") as fh:
                    data = fh.read()

                mime_type, _ = mimetypes.guess_type(fname)
                if not mime_type:
                    mime_type = "application/octet-stream"

                # Tipo 6 = inline (olByValue embebido), tipo 5 = OLE
                # Detectar si es inline por el PropertyAccessor (PR_ATTACH_CONTENT_ID)
                cid = None
                try:
                    cid = att.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
                    )
                except Exception:
                    pass

                if cid:
                    inline_parts.append((cid, fname, data, mime_type))
                else:
                    attach_parts.append((fname, data, mime_type))

            except Exception as exc:
                print(f"    SKIP adjunto {i}: {exc}")

        # ── Construir estructura MIME ──────────────────────────────
        body_html  = ""
        body_plain = safe(item, "Body", "")
        try:
            body_html = item.HTMLBody or ""
        except Exception:
            pass

        if inline_parts:
            # multipart/related: HTML + imagenes inline
            related = email.mime.multipart.MIMEMultipart("related")

            # Parte de texto
            if body_html:
                related.attach(email.mime.text.MIMEText(body_html, "html", "utf-8"))
            else:
                related.attach(email.mime.text.MIMEText(body_plain, "plain", "utf-8"))

            # Partes inline
            for cid, fname, data, mime_type in inline_parts:
                maintype, subtype = mime_type.split("/", 1)
                part = email.mime.base.MIMEBase(maintype, subtype)
                part.set_payload(data)
                email.encoders.encode_base64(part)
                part.add_header("Content-ID", f"<{cid}>")
                part.add_header("Content-Disposition", "inline", filename=fname)
                related.attach(part)

            if attach_parts:
                # Si hay adjuntos normales, envolver en multipart/mixed
                mixed = email.mime.multipart.MIMEMultipart("mixed")
                mixed.attach(related)
                root_msg = mixed
            else:
                root_msg = related

        elif attach_parts:
            root_msg = email.mime.multipart.MIMEMultipart("mixed")
            if body_html:
                root_msg.attach(email.mime.text.MIMEText(body_html, "html", "utf-8"))
            else:
                root_msg.attach(email.mime.text.MIMEText(body_plain, "plain", "utf-8"))

        else:
            # Sin adjuntos ni inline — mensaje simple
            if body_html:
                root_msg = email.mime.text.MIMEText(body_html, "html", "utf-8")
            else:
                root_msg = email.mime.text.MIMEText(body_plain, "plain", "utf-8")

        # Adjuntos normales
        for fname, data, mime_type in attach_parts:
            maintype, subtype = mime_type.split("/", 1)
            part = email.mime.base.MIMEBase(maintype, subtype)
            part.set_payload(data)
            email.encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment", filename=fname)
            root_msg.attach(part)

    # ── Cabeceras del sobre ────────────────────────────────────────
    root_msg["From"]       = safe(item, "SenderEmailAddress", "unknown@unknown")
    root_msg["To"]         = safe(item, "To", "")
    root_msg["CC"]         = safe(item, "CC", "")
    root_msg["Subject"]    = safe(item, "Subject", "(sin asunto)")
    root_msg["Message-ID"] = safe(item, "InternetMessageId",
                                   f"<gen-{id(item)}@pst2mbox>")
    root_msg["X-PST-Folder"] = folder_name

    for attr in ("ReceivedTime", "SentOn"):
        try:
            dt = getattr(item, attr)
            if dt:
                root_msg["Date"] = email.utils.format_datetime(
                    datetime(dt.year, dt.month, dt.day,
                             dt.hour, dt.minute, dt.second)
                )
                break
        except Exception:
            pass

    return root_msg


def safe(obj, attr, default=""):
    try:
        v = getattr(obj, attr)
        return str(v) if v is not None else default
    except Exception:
        return default


# ─────────────────────────────────────────────────────────────────────
#  BACKEND: libratom (Linux / macOS)
# ─────────────────────────────────────────────────────────────────────
def libratom_available():
    try:
        from libratom.lib.pff import PffArchive  # noqa
        return True
    except ImportError:
        return False


def iter_messages_libratom(pst_path: str):
    from libratom.lib.pff import PffArchive
    with PffArchive(pst_path) as archive:
        for folder in archive.folders():
            folder_name = normalize_folder_name(getattr(folder, "name", "inbox"))
            for msg in folder.sub_messages:
                try:
                    headers = msg.transport_headers or ""
                    body    = msg.plain_text_body or b""
                    if isinstance(body, str):
                        body = body.encode("utf-8", errors="replace")
                    raw = headers.encode("utf-8", errors="replace") + b"\r\n" + body
                    em  = email.message_from_bytes(raw)
                    em["X-PST-Folder"] = folder_name
                    yield folder_name, em
                except Exception as exc:
                    print(f"  SKIP: {exc}")


# ─────────────────────────────────────────────────────────────────────
#  Seleccion de backend
# ─────────────────────────────────────────────────────────────────────
def get_iter_fn():
    if outlook_available():
        print("Backend: Outlook COM (win32com)")
        return iter_messages_outlook
    if libratom_available():
        print("Backend: libratom")
        return iter_messages_libratom
    print(
        "\nERROR: No hay backend disponible.\n"
        "Windows: pip install pywin32\n"
        "Linux:   pip install setuptools && pip install libratom\n"
    )
    sys.exit(1)


# ─────────────────────────────────────────────────────────────────────
#  Escritor Mbox con division automatica por tamano
# ─────────────────────────────────────────────────────────────────────
class FolderWriter:
    def __init__(self, folder_name, out_path, max_bytes):
        self.folder_name = folder_name
        self.out_path    = out_path
        self.max_bytes   = max_bytes
        self.part        = 1
        self.size        = 0
        self.count       = 0
        self.parts       = []
        self._mbox       = None
        self._open_new()

    def _open_new(self):
        fname = (f"{self.folder_name}.mbox" if self.part == 1
                 else f"{self.folder_name}_parte{self.part:03d}.mbox")
        p = self.out_path / fname
        if p.exists():
            p.unlink()
        self._mbox = mailbox.mbox(str(p))
        self.parts.append(p)

    def add(self, em):
        raw = em.as_bytes()
        sz  = len(raw)
        if self.size + sz > self.max_bytes and self.size > 0:
            self._mbox.flush()
            self._mbox.close()
            self.part += 1
            self.size  = 0
            self._open_new()
        self._mbox.add(em)
        self.size  += sz
        self.count += 1
        if self.count % 100 == 0:
            self._mbox.flush()

    def close(self):
        if self._mbox:
            self._mbox.flush()
            self._mbox.close()
            self._mbox = None


# ─────────────────────────────────────────────────────────────────────
#  Conversion principal
# ─────────────────────────────────────────────────────────────────────
def pst_to_mbox(pst_path, output_dir, max_bytes, do_zip):
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    for stale in out_path.glob("*.lock"):
        try:
            stale.unlink()
        except Exception:
            pass

    writers    = {}
    total      = 0
    iter_fn    = get_iter_fn()

    print(f"\nLeyendo: {pst_path}")
    print(f"Max por archivo: {max_bytes // BYTES_PER_MB} MB\n")

    for folder_name, em in iter_fn(pst_path):
        if folder_name not in writers:
            writers[folder_name] = FolderWriter(folder_name, out_path, max_bytes)
            print(f"  Carpeta detectada: {folder_name}")
        writers[folder_name].add(em)
        total += 1
        if total % 100 == 0:
            print(f"  ... {total} mensajes procesados", end="\r")

    print(f"\n\nTotal: {total} mensajes en {len(writers)} carpeta(s)\n")

    all_parts = []
    for folder_name, w in writers.items():
        w.close()
        for p in w.parts:
            size_mb = p.stat().st_size // BYTES_PER_MB
            print(f"  {folder_name:<25} {w.count:>6} msgs  {size_mb:>5} MB  ->  {p.name}")
            all_parts.append(p)

    if do_zip:
        print("\nComprimiendo...")
        zipped = []
        for p in all_parts:
            zp = p.with_suffix(".zip")
            with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
                zf.write(p, arcname=p.name)
            p.unlink()
            zipped.append(zp)
            print(f"  {zp.name}  ({zp.stat().st_size // BYTES_PER_MB} MB)")
        all_parts = zipped

    print("\n" + "="*60)
    print("INSTRUCCIONES PARA IMPORTAR EN ROUNDCUBE:")
    print("="*60)
    print("Para cada archivo:")
    print("  1. Selecciona la carpeta destino en Roundcube")
    print("  2. Ve a Ajustes -> Importar y sube el archivo")
    print()
    for p in all_parts:
        carpeta = p.stem.split("_parte")[0]
        print(f"  {p.name:<45} -> carpeta '{carpeta}'")


# ─────────────────────────────────────────────────────────────────────
#  CLI
# ─────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Convierte .pst a Mbox por carpeta para Roundcube"
    )
    parser.add_argument("pst_file", help="Ruta al archivo .pst")
    parser.add_argument("--output-dir", "-o", default="./mbox_output",
                        help="Carpeta de salida (default: ./mbox_output)")
    parser.add_argument("--max-mb", "-m", type=int, default=MAX_MB_DEFAULT,
                        help=f"Max MB por archivo (default: {MAX_MB_DEFAULT})")
    parser.add_argument("--zip", "-z", action="store_true",
                        help="Comprimir cada Mbox en ZIP")
    args = parser.parse_args()

    if not os.path.isfile(args.pst_file):
        print(f"ERROR: No se encontro '{args.pst_file}'")
        sys.exit(1)

    pst_to_mbox(
        pst_path   = args.pst_file,
        output_dir = args.output_dir,
        max_bytes  = args.max_mb * BYTES_PER_MB,
        do_zip     = args.zip,
    )


if __name__ == "__main__":
    main()
