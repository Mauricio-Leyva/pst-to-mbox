# PST a Mbox Portable

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-blue.svg)](#requisitos)

Interfaz gráfica para convertir archivos **PST** (Outlook) a formato **Mbox**, sin necesidad de usar la terminal.

Genera un `.exe` portable que podés compartir con cualquier compañero, sin que necesite Python ni Outlook instalado — solo Outlook de escritorio.

---

## Capturas de pantalla

```
┌──────────────────────────────────────────────────────┐
│  PST a Mbox Portable                                 │
├──────────────────────────────────────────────────────┤
│  Archivo PST:    [________________________] [Examinar]│
│  Carpeta salida: [________________________] [Elegir]  │
│  Max MB archivo: [100]                              │
│  [x] Comprimir cada Mbox en ZIP                     │
│               [Convertir]                           │
│  ─────────────────────────────────────────────────  │
│  Log:                                               │
│  > Conversion completada. 5 archivos generados.     │
└──────────────────────────────────────────────────┘
```

---

## Características

- **Interfaz gráfica** con Tkinter — sin terminal
- Conversión batch de archivos PST a Mbox
- División por tamaño máximo (configurable en MB)
- Compresión ZIP opcional por archivo generado
- Log en tiempo real dentro de la ventana
- **EXE portable** — compartilo sin instalar nada

---

## Requisitos

- **Windows** (el EXE es portable y no necesita instalación)
- **Outlook de escritorio** instalado en la PC donde se ejecuta la conversión
- Sin Python ni dependencias adicionales requeridos por el usuario final

---

## Uso

1. Descargá `PST2MboxPortable.exe` de la carpeta `dist/`
2. Ejecutalo en cualquier PC con Windows y Outlook
3. Seleccioná el archivo `.pst`
4. Elegí la carpeta de salida
5. Ajustá el tamaño máximo por archivo si es necesario
6. (Opcional) Activá la compresión ZIP
7. Click en **Convertir**

---

## Desarrollo

### Requisitos del entorno de desarrollo

- Python 3.8+
- `pst_to_mbox2.py` (lógica de conversión)
- `pst_to_mbox_gui.py` (interfaz gráfica)

### Generar el EXE portable

```powershell
.\build_portable_exe.bat
```

El ejecutable se genera en `dist\PST2MboxPortable.exe`.

### Personalizar metadatos del EXE

Editá `version_info.txt` antes de compilar:

| Campo              | Descripción                       |
|--------------------|-----------------------------------|
| `CompanyName`      | Nombre de la empresa/autor        |
| `FileDescription`  | Descripción de la app             |
| `ProductName`      | Nombre del producto               |
| `FileVersion`      | Versión del archivo               |
| `LegalCopyright`   | Texto de copyright                |

---

## Estructura del proyecto

```
├── pst_to_mbox2.py          # Lógica de conversión (reutilizable)
├── pst_to_mbox_gui.py       # Interfaz gráfica (Tkinter)
├── build_portable_exe.bat   # Script para generar el .exe
├── version_info.txt         # Metadatos del EXE
├── PST2MboxPortable.spec    # Spec file para PyInstaller
├── dist/                    # Carpeta de salida del .exe
├── build/                   # Archivos temporales de build
└── README.md                # Este archivo
```

---

## Notas

- Si el PST es grande, la conversión puede tardar varios minutos.
- En algunos equipos, **Windows Defender** puede bloquear la escritura en `Documentos`. Usá una carpeta local como `Downloads\mbox_output`.
- El log de la ventana muestra el progreso y errores en tiempo real.

---

## Licencia

MIT © 2026 Mauricio Leyva
