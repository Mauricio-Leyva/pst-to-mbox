# PST a Mbox - App portable (.exe)

Este proyecto incluye una interfaz grafica para usar el convertidor sin terminal.

## Archivos nuevos

- `pst_to_mbox_gui.py`: interfaz grafica (Tkinter)
- `build_portable_exe.bat`: script para generar el `.exe`
- `version_info.txt`: campos que aparecen en Propiedades -> Detalles

## Personalizar Detalles del EXE

Antes de compilar, edita `version_info.txt` y cambia estos campos:

- `CompanyName`: tu nombre o empresa
- `FileDescription`: descripcion de la app
- `ProductName`: nombre del producto
- `FileVersion` y `ProductVersion`: version visible
- `LegalCopyright`: texto de copyright

## Como generar el EXE

1. Abre PowerShell en esta carpeta.
2. Ejecuta:

```powershell
.\build_portable_exe.bat
```

3. Al terminar, toma este archivo:

- `dist\PST2MboxPortable.exe`

Ese archivo es el que puedes compartir con tus companeros.

## Uso para usuario final

1. Abrir `PST2MboxPortable.exe`
2. Elegir archivo `.pst`
3. Elegir carpeta de salida
4. Ajustar `Max MB por archivo` si hace falta
5. (Opcional) dejar marcada la opcion ZIP
6. Click en `Convertir`

## Notas importantes

- Requisito recomendado en la PC destino: Outlook de escritorio instalado.
- Si el PST es grande, puede tardar varios minutos.
- El log de la ventana muestra el progreso y errores.
- Si aparece error de carpeta, selecciona salida en `Downloads\\mbox_output`.
- En algunos equipos, Windows Defender puede bloquear escritura en Documentos.
