# Manual de Usuario - VORTEX / VORTEX User Manual

Manual de usuario **bilingüe** del sistema **VORTEX**, gestión de taller de bombas de agua.
Un solo documento con la versión en español y, a continuación, la misma versión en inglés.

*Bilingual user manual for the **VORTEX** water pump workshop management system.
A single document containing the Spanish version followed by the same content in English.*

| Archivo | Contenido |
|---|---|
| `Manual_de_Usuario_VORTEX.docx` | Manual final, listo para abrir en Word |
| `generar_manual.py` | Script que genera el `.docx` |
| `capturas/` | Capturas de pantalla que se incrustan en el manual |

## Contenido del manual

El documento contiene las dos versiones, en este orden:

```
Portada
├── Contenido            (índice de la versión en español)
├── VERSIÓN EN ESPAÑOL
│    1. Introducción — propósito del sistema y perfiles de usuario
│    2. Módulo Cliente: Solicitar servicio
│    3. Módulo Administrador — Inicio, Clientes y solicitudes, Inventario, Reportes
│    4. Flujo de trabajo completo
│    5. Glosario
├── Portadilla en inglés
├── Contents             (índice de la versión en inglés)
└── ENGLISH VERSION
     1. Introduction — purpose of the system and user profiles
     2. Client Module: Request a service
     3. Administrator Module — Home, Clients and requests, Inventory, Reports
     4. Complete workflow
     5. Glossary
```

Ambas versiones son idénticas en estructura: 28 títulos y 26 tablas cada una, con las
mismas cinco figuras. El contenido se define una sola vez en `generar_manual.py`, con
cada cadena como par `(español, inglés)`, de modo que las dos versiones no pueden
desincronizarse.

En la versión en inglés los nombres de botones y campos **se conservan en español**,
tal como aparecen en pantalla, con la traducción entre paréntesis —
por ejemplo *the Enviar solicitud (Submit request) button*.

## Formato

- Portada sin numerar; numeración corrida en todo el contenido
- Pie de página bilingüe: `Página X de Y · Page X of Y`
- Encabezado en todas las páginas: `VORTEX | Manual de Usuario · User Manual`
- Dos índices automáticos independientes, cada uno delimitado con un marcador
  (`VERSION_ES` y `VERSION_EN`) para que liste únicamente su propio idioma

> Al abrir el documento, Word preguntará si desea actualizar los índices. Acepte, o
> seleccione cada tabla de contenido y presione **F9** para generar los números de página.
>
> *When opening the document, Word will ask to update the indexes. Accept, or select
> each table of contents and press **F9** to build the page numbers.*

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
