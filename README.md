# PST a Mbox Portable

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-blue.svg)](#requisitos)

Interfaz gráfica para convertir archivos **PST** (Outlook) a formato **Mbox**, sin necesidad de utilizar la terminal.

Genera un ejecutable (`.exe`) portable que puedes compartir con cualquier usuario, sin que necesite tener Python instalado.

---

## Características

- **Interfaz gráfica (GUI)** basada en Tkinter, fácil de utilizar.
- Conversión por lotes (batch) de carpetas dentro de archivos PST a formato Mbox.
- División automática de archivos por tamaño máximo (configurable en MB).
- Compresión ZIP opcional para cada archivo Mbox generado.
- Registro de eventos (log) en tiempo real integrado en la ventana.
- **Ejecutable portable**: no requiere instalación.

---

## 🚀 Descarga e Instalación (Recomendado)

Para la mayoría de los usuarios, la forma más sencilla de utilizar esta herramienta es descargando el ejecutable precompilado:

1. Ve a la sección de **Releases** del repositorio.
2. Descarga la versión más reciente del archivo `PST2MboxPortable.exe`.
3. Guárdalo en tu computadora y ejecútalo (no requiere instalación).

*Nota: Se requiere tener **Microsoft Outlook de escritorio** instalado en la máquina donde se realizará la conversión.*

---

## Uso de la aplicación

1. Ejecuta `PST2MboxPortable.exe`.
2. Haz clic en **Seleccionar archivo PST** y elige el archivo que deseas convertir.
3. Selecciona la carpeta de destino donde se guardarán los archivos Mbox.
4. (Opcional) Ajusta el tamaño máximo por archivo si necesitas fragmentar los resultados.
5. (Opcional) Activa la casilla de compresión ZIP si deseas ahorrar espacio.
6. Haz clic en **Convertir** y espera a que el proceso finalice consultando el panel de registro (log).

---

## 🛠️ Desarrollo y Compilación

Si deseas modificar el código o compilar el ejecutable por tu cuenta, sigue estos pasos:

### Requisitos previos

- Python 3.8 o superior.
- Microsoft Outlook de escritorio instalado.
- Dependencias (asegúrate de tener `pyinstaller` si deseas generar el `.exe`).

### Generar el ejecutable portable

En la raíz del proyecto, ejecuta el script de compilación:

```powershell
.\build_portable_exe.bat
```

El ejecutable final se generará en la ruta `dist\PST2MboxPortable.exe`.

### Personalizar metadatos del ejecutable

Puedes editar el archivo `version_info.txt` antes de compilar para modificar los metadatos del `.exe`:

| Campo              | Descripción                       |
|--------------------|-----------------------------------|
| `CompanyName`      | Nombre de la empresa o autor      |
| `FileDescription`  | Descripción de la aplicación      |
| `ProductName`      | Nombre del producto               |
| `FileVersion`      | Versión del archivo               |
| `LegalCopyright`   | Texto de derechos de autor (copyright) |

---

## 📁 Estructura del proyecto

```text
├── pst_to_mbox2.py          # Lógica principal de conversión
├── pst_to_mbox_gui.py       # Lógica de la interfaz gráfica
├── build_portable_exe.bat   # Script automatizado para compilar
├── version_info.txt         # Configuración de metadatos para el ejecutable
└── README.md                # Documentación del proyecto
```

---

## ⚠️ Notas importantes

- **Tiempo de procesamiento:** Los archivos PST de gran tamaño pueden tomar varios minutos en procesarse. Se paciente.
- **Problemas de permisos:** En algunos equipos, los antivirus o Windows Defender pueden bloquear la escritura en carpetas como `Documentos`. Si experimentas errores, intenta usar tu carpeta de `Descargas` o una ruta en `C:\`.

---

## Licencia

Este proyecto está distribuido bajo la licencia [MIT](LICENSE) © 2026 Mauricio Leyva.
