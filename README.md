# Manual de Usuario - VORTEX

Manual de usuario del sistema **VORTEX**, gestión de taller de bombas de agua.

| Archivo | Contenido |
|---|---|
| `Manual_de_Usuario_VORTEX.docx` | Manual final, listo para abrir en Word |
| `generar_manual.py` | Script que genera el `.docx` |
| `capturas/` | Capturas de pantalla que se incrustan en el manual |

## Contenido del manual

1. Introducción — propósito del sistema y perfiles de usuario
2. Módulo Cliente: Solicitar servicio
3. Módulo Administrador — Inicio, Clientes y solicitudes, Inventario, Reportes
4. Flujo de trabajo completo
5. Glosario

Portada sin numerar, encabezado en todas las páginas, pie con `Página X de Y` e índice automático.

> Al abrir el documento, Word preguntará si desea actualizar el índice. Acepte, o
> seleccione la tabla de contenido y presione **F9** para generar los números de página.

## Estado de las capturas

Las cinco capturas **aún no están incluidas**. El documento tiene marcos reservados
del mismo tamaño que la imagen final (6.3 × 3.95 pulgadas), con el título de cada figura.

Para incorporarlas hay dos caminos:

**Opción A — regenerar el documento (recomendada)**

Coloque los archivos en `capturas/` con estos nombres y vuelva a ejecutar el script:

| Archivo | Pantalla |
|---|---|
| `figura1.png` | Módulo Cliente — Solicitar servicio |
| `figura2.png` | Administrador — Inicio (panel de resumen) |
| `figura3.png` | Administrador — Clientes y solicitudes |
| `figura4.png` | Administrador — Inventario de piezas |
| `figura5.png` | Administrador — Reportes y estadísticas |

```bash
pip install python-docx
python3 generar_manual.py
```

El script imprime la asignación de cada imagen para confirmar que quedaron en su lugar.
Si los nombres no son `figuraN.png`, toma las imágenes de la carpeta en orden alfabético.

**Opción B — pegarlas en Word**

Abrir el `.docx` y pegar cada captura dentro del marco gris que le corresponde,
guiándose por el pie de figura que ya está escrito debajo.
