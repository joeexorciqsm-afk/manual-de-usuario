"""
Generador del Manual de Usuario de VORTEX
Sistema de Gestion para Taller de Bombas de Agua

Uso:
    python3 generar_manual.py

Si existen las capturas en ./capturas/figura1.png ... figura5.png se incrustan
automaticamente. Si no existen, se dibuja un marco reservado con su titulo.
"""

import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------------- paleta
NAVY = RGBColor(0x0F, 0x22, 0x33)
NAVY_HEX = "0F2233"
BLUE = RGBColor(0x15, 0x84, 0xD8)
BLUE_HEX = "1584D8"
GRAY = RGBColor(0x5A, 0x6A, 0x78)
LIGHT_HEX = "EEF2F6"
ZEBRA_HEX = "F7F9FB"
BORDER_HEX = "C9D4DE"

CAPTURAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "capturas")


# ------------------------------------------------------------- utilidades
def set_cell_bg(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def set_cell_borders(cell, hex_color=BORDER_HEX, sz=4):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:color"), hex_color)
        borders.append(el)
    tcPr.append(borders)
    return cell


def add_field(paragraph, code):
    """Inserta un campo de Word (PAGE, SECTIONPAGES, TOC...)."""
    run = paragraph.add_run()
    f1 = OxmlElement("w:fldChar")
    f1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = code
    f2 = OxmlElement("w:fldChar")
    f2.set(qn("w:fldCharType"), "end")
    run._r.append(f1)
    run._r.append(instr)
    run._r.append(f2)
    return run


def restart_page_numbering(section, start=1):
    pg = OxmlElement("w:pgNumType")
    pg.set(qn("w:start"), str(start))
    section._sectPr.append(pg)


def horizontal_rule(paragraph, hex_color=BLUE_HEX, sz=18):
    pPr = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), hex_color)
    borders.append(bottom)
    pPr.append(borders)


def p(doc, text="", size=10.5, bold=False, italic=False, color=None,
      align=None, space_after=6, space_before=0, indent=None):
    par = doc.add_paragraph()
    par.paragraph_format.space_after = Pt(space_after)
    par.paragraph_format.space_before = Pt(space_before)
    if indent is not None:
        par.paragraph_format.left_indent = Inches(indent)
    if align is not None:
        par.alignment = align
    if text:
        r = par.add_run(text)
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        r.font.color.rgb = color if color else RGBColor(0x20, 0x2A, 0x33)
    return par


def bullets(doc, items, style="List Bullet"):
    for it in items:
        par = doc.add_paragraph(style=style)
        par.paragraph_format.space_after = Pt(3)
        if isinstance(it, tuple):
            r = par.add_run(it[0])
            r.bold = True
            r.font.size = Pt(10.5)
            r2 = par.add_run(" " + it[1])
            r2.font.size = Pt(10.5)
        else:
            r = par.add_run(it)
            r.font.size = Pt(10.5)


def data_table(doc, rows, widths=None, font_size=8.5, header_hex=NAVY_HEX):
    """Tabla con encabezado oscuro, texto blanco y filas alternas."""
    tbl = doc.add_table(rows=len(rows), cols=len(rows[0]))
    tbl.style = "Table Grid"
    tbl.autofit = False
    for i, fila in enumerate(rows):
        for j, valor in enumerate(fila):
            cell = tbl.cell(i, j)
            cell.text = ""
            par = cell.paragraphs[0]
            par.paragraph_format.space_after = Pt(2)
            par.paragraph_format.space_before = Pt(2)
            run = par.add_run(str(valor))
            run.font.size = Pt(font_size)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i == 0:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                set_cell_bg(cell, header_hex)
            else:
                if i % 2 == 0:
                    set_cell_bg(cell, ZEBRA_HEX)
            if widths:
                cell.width = Inches(widths[j])
    return tbl


def listar_capturas():
    """Imagenes disponibles en ./capturas, en orden natural de nombre."""
    if not os.path.isdir(CAPTURAS):
        return []
    validas = (".png", ".jpg", ".jpeg")
    archivos = [f for f in os.listdir(CAPTURAS)
                if f.lower().endswith(validas)]
    return sorted(archivos, key=lambda s: s.lower())


DISPONIBLES = listar_capturas()
MAPEO = []


def buscar_captura(numero):
    """Resuelve la imagen de la figura N.

    1) Nombre explicito: figura<N>.png / .jpg / .jpeg
    2) Respaldo: la N-esima imagen de la carpeta en orden alfabetico.
    """
    for ext in ("png", "jpg", "jpeg", "PNG", "JPG", "JPEG"):
        cand = os.path.join(CAPTURAS, f"figura{numero}.{ext}")
        if os.path.exists(cand):
            return cand
    if len(DISPONIBLES) >= numero:
        return os.path.join(CAPTURAS, DISPONIBLES[numero - 1])
    return None


def figura(doc, numero, titulo, alto=3.95):
    """Incrusta la captura si existe; si no, reserva un marco con su titulo."""
    ruta = buscar_captura(numero)
    MAPEO.append((numero, os.path.basename(ruta) if ruta else "-- marco reservado --"))

    if ruta:
        par = doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_before = Pt(6)
        par.paragraph_format.space_after = Pt(2)
        par.add_run().add_picture(ruta, width=Inches(6.3))
    else:
        tbl = doc.add_table(rows=1, cols=1)
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.4)
        set_cell_bg(cell, LIGHT_HEX)
        set_cell_borders(cell, BORDER_HEX, sz=8)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        row = tbl.rows[0]
        row.height = Inches(alto)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        par = cell.paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = par.add_run(f"[ Figura {numero} ]")
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = GRAY
        par2 = cell.add_paragraph()
        par2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = par2.add_run(titulo)
        r2.italic = True
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = GRAY

    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(4)
    cap.paragraph_format.space_after = Pt(12)
    r = cap.add_run(f"Figura {numero}. ")
    r.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = NAVY
    r2 = cap.add_run(titulo)
    r2.font.size = Pt(9)
    r2.italic = True
    r2.font.color.rgb = GRAY


def nota(doc, texto, etiqueta="Nota:"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.4)
    set_cell_bg(cell, LIGHT_HEX)
    set_cell_borders(cell, BLUE_HEX, sz=6)
    par = cell.paragraphs[0]
    par.paragraph_format.space_after = Pt(2)
    par.paragraph_format.space_before = Pt(2)
    r = par.add_run(etiqueta + " ")
    r.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = NAVY
    r2 = par.add_run(texto)
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(0x20, 0x2A, 0x33)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)


# =====================================================================
doc = Document()

doc.core_properties.title = "Manual de Usuario - VORTEX"
doc.core_properties.subject = "Sistema de Gestion para Taller de Bombas de Agua"
doc.core_properties.author = "VORTEX"

# --- estilos base
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")

for nombre, tam, color in (("Heading 1", 17, NAVY), ("Heading 2", 13.5, BLUE),
                           ("Heading 3", 11.5, NAVY)):
    st = doc.styles[nombre]
    st.font.name = "Calibri"
    st.font.size = Pt(tam)
    st.font.bold = True
    st.font.color.rgb = color
    st.paragraph_format.space_before = Pt(14 if nombre == "Heading 1" else 10)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.keep_with_next = True

# --- margenes
for s in doc.sections:
    s.top_margin = Inches(1.0)
    s.bottom_margin = Inches(0.9)
    s.left_margin = Inches(1.0)
    s.right_margin = Inches(1.0)

# =====================================================  PORTADA
portada = doc.sections[0]

for _ in range(5):
    doc.add_paragraph()

t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
t.paragraph_format.space_after = Pt(0)
r = t.add_run("VORTEX")
r.font.size = Pt(58)
r.bold = True
r.font.color.rgb = BLUE

st = doc.add_paragraph()
st.alignment = WD_ALIGN_PARAGRAPH.CENTER
st.paragraph_format.space_after = Pt(24)
r = st.add_run("Sistema de Gestion para Taller de Bombas de Agua")
r.font.size = Pt(13)
r.font.color.rgb = GRAY

rule = doc.add_paragraph()
rule.alignment = WD_ALIGN_PARAGRAPH.CENTER
horizontal_rule(rule, BLUE_HEX, 24)
rule.paragraph_format.space_after = Pt(22)

mt = doc.add_paragraph()
mt.alignment = WD_ALIGN_PARAGRAPH.CENTER
mt.paragraph_format.space_after = Pt(6)
r = mt.add_run("MANUAL DE USUARIO")
r.font.size = Pt(26)
r.bold = True
r.font.color.rgb = NAVY

sb = doc.add_paragraph()
sb.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sb.add_run("Guia de funcionamiento de los modulos Cliente y Administrador")
r.font.size = Pt(11)
r.italic = True
r.font.color.rgb = GRAY

for _ in range(8):
    doc.add_paragraph()

info = doc.add_table(rows=3, cols=2)
info.alignment = WD_ALIGN_PARAGRAPH.CENTER
etiquetas = [("Documento", "Manual de usuario"),
             ("Version", "1.0"),
             ("Fecha", "Julio de 2026")]
for i, (k, v) in enumerate(etiquetas):
    c0, c1 = info.cell(i, 0), info.cell(i, 1)
    c0.width, c1.width = Inches(1.5), Inches(2.6)
    pr = c0.paragraphs[0].add_run(k)
    pr.bold = True
    pr.font.size = Pt(10)
    pr.font.color.rgb = NAVY
    c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pv = c1.paragraphs[0].add_run("   " + v)
    pv.font.size = Pt(10)
    pv.font.color.rgb = GRAY

# ==============================================  SECCION DE CONTENIDO
cuerpo = doc.add_section(WD_SECTION.NEW_PAGE)
cuerpo.top_margin = Inches(1.0)
cuerpo.bottom_margin = Inches(0.9)
cuerpo.left_margin = Inches(1.0)
cuerpo.right_margin = Inches(1.0)
restart_page_numbering(cuerpo, 1)

# --- encabezado
cuerpo.header.is_linked_to_previous = False
hp = cuerpo.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r = hp.add_run("VORTEX")
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = BLUE
r2 = hp.add_run("  |  Manual de Usuario")
r2.font.size = Pt(9)
r2.font.color.rgb = GRAY
horizontal_rule(hp, BORDER_HEX, 6)

# --- pie con numeracion de paginas
cuerpo.footer.is_linked_to_previous = False
fp = cuerpo.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
rf = fp.add_run("Pagina ")
rf.font.size = Pt(9)
rf.font.color.rgb = GRAY
add_field(fp, "PAGE").font.size = Pt(9)
rf2 = fp.add_run(" de ")
rf2.font.size = Pt(9)
rf2.font.color.rgb = GRAY
add_field(fp, "SECTIONPAGES").font.size = Pt(9)
for run in fp.runs:
    run.font.size = Pt(9)
    run.font.color.rgb = GRAY

# ---------------------------------------------------- TABLA DE CONTENIDO
doc.add_heading("Contenido", level=1)

toc_par = doc.add_paragraph()
run = toc_par.add_run()
f1 = OxmlElement("w:fldChar")
f1.set(qn("w:fldCharType"), "begin")
f1.set(qn("w:dirty"), "true")
instr = OxmlElement("w:instrText")
instr.set(qn("xml:space"), "preserve")
instr.text = r'TOC \o "1-2" \h \z \u'
fsep = OxmlElement("w:fldChar")
fsep.set(qn("w:fldCharType"), "separate")
ftxt = OxmlElement("w:t")
ftxt.text = "Seleccione esta tabla y presione F9 para generar el indice con numeros de pagina."
fend = OxmlElement("w:fldChar")
fend.set(qn("w:fldCharType"), "end")
run._r.append(f1)
run._r.append(instr)
run._r.append(fsep)
run._r.append(ftxt)
run._r.append(fend)
run.font.size = Pt(10)
run.font.color.rgb = GRAY

doc.add_page_break()

# ================================================== 1. INTRODUCCION
doc.add_heading("1. Introduccion", level=1)
p(doc, "VORTEX es el sistema de gestion de un taller de bombas de agua. Reune en una "
        "sola aplicacion la recepcion de solicitudes de servicio de los clientes, el "
        "seguimiento de esas solicitudes, el control del inventario de piezas y la "
        "consulta de reportes de ventas.")
p(doc, "El sistema se organiza en dos entornos de trabajo independientes:")
bullets(doc, [
    ("Modulo Cliente.", "Pantalla publica donde una persona registra su solicitud de "
     "servicio sin necesidad de tener cuenta ni contrasena."),
    ("Modulo Administrador.", "Area privada, protegida por inicio de sesion, donde el "
     "personal del taller gestiona solicitudes, inventario y reportes."),
])
p(doc, "Este manual describe el funcionamiento de cada pantalla, el significado de sus "
        "datos y las acciones disponibles en cada una.", space_after=10)

doc.add_heading("1.1 Perfiles de usuario", level=2)
data_table(doc, [
    ["Perfil", "Acceso", "Que puede hacer"],
    ["Cliente", "Libre, sin credenciales",
     "Registrar una solicitud de servicio y obtener su numero de pedido."],
    ["Administrador", "Requiere iniciar sesion",
     "Consultar el resumen del taller, administrar clientes y solicitudes, "
     "mantener el inventario de piezas y revisar reportes de ventas."],
], widths=[1.2, 1.6, 3.6], font_size=9.5)
doc.add_paragraph().paragraph_format.space_after = Pt(4)

# ================================================== 2. MODULO CLIENTE
doc.add_page_break()
doc.add_heading("2. Modulo Cliente: Solicitar servicio", level=1)
p(doc, "Es la pantalla con la que inicia el sistema. Su unico proposito es capturar la "
        "solicitud de servicio de un cliente. La pantalla se divide en un menu lateral "
        "oscuro a la izquierda y el formulario de captura a la derecha.")

figura(doc, 1, "Pantalla Solicitar servicio del modulo Cliente")

doc.add_heading("2.1 Menu lateral", level=2)
p(doc, "La barra lateral identifica el modulo activo y concentra la navegacion:")
bullets(doc, [
    ("VORTEX / Cliente.", "Encabezado que confirma que se esta trabajando en el "
     "entorno de cliente."),
    ("Solicitar servicio.", "Unica opcion del menu y por eso aparece siempre "
     "resaltada en azul: indica la seccion en la que se encuentra."),
    ("Iniciar sesion como admin.", "Boton al pie del menu, precedido por la pregunta "
     "\u00bfEres admin? Da paso a la pantalla de acceso del personal del taller."),
])

doc.add_heading("2.2 Campos del formulario", level=2)
p(doc, "El texto guia bajo el titulo indica al cliente que debe llenar sus datos y "
        "describir el problema, y le anticipa que al enviar recibira un numero de pedido. "
        "Los campos solicitados son:")
data_table(doc, [
    ["Campo", "Tipo", "Que se registra"],
    ["Nombre", "Linea de texto",
     "Nombre de la persona o del negocio que solicita el servicio. Es el nombre con el "
     "que la solicitud aparecera en el modulo de administracion."],
    ["Numero (telefono)", "Linea de texto",
     "Telefono de contacto para confirmar la visita o avisar del avance."],
    ["Direccion", "Linea de texto",
     "Domicilio donde se prestara el servicio; orienta al tecnico que sera asignado."],
    ["Motivo por el que solicita el servicio", "Area de texto amplia",
     "Descripcion del problema de la bomba. Al ser un campo grande admite varias lineas, "
     "de modo que el cliente puede detallar sintomas, antecedentes o el modelo del equipo."],
], widths=[1.7, 1.1, 3.6], font_size=9.5)
doc.add_paragraph().paragraph_format.space_after = Pt(4)

doc.add_heading("2.3 Envio de la solicitud", level=2)
p(doc, "Para registrar una solicitud:")
pasos = [
    "Escribir el nombre completo en el campo Nombre.",
    "Capturar un telefono de contacto valido en Numero (telefono).",
    "Indicar la direccion donde se encuentra el equipo en Direccion.",
    "Describir la falla en el area Motivo por el que solicita el servicio.",
    "Pulsar el boton Enviar solicitud.",
]
for i, paso in enumerate(pasos, 1):
    par = doc.add_paragraph(style="List Number")
    par.paragraph_format.space_after = Pt(3)
    par.add_run(paso).font.size = Pt(10.5)

p(doc, "Al confirmar el envio el sistema genera un numero de pedido, que es el "
        "identificador con el que el cliente puede dar seguimiento a su servicio. Ese "
        "mismo consecutivo es el que el administrador vera despues en la columna ID de "
        "las pantallas de gestion.", space_before=6)

nota(doc, "La solicitud entra al sistema con el estado PENDIENTE. A partir de ese "
          "momento queda visible para el administrador tanto en el panel de Inicio como "
          "en la pantalla Clientes y solicitudes.")

# ================================================== 3. MODULO ADMIN
doc.add_page_break()
doc.add_heading("3. Modulo Administrador", level=1)
p(doc, "Tras iniciar sesion se abre el area de administracion. Todas sus pantallas "
        "comparten el mismo menu lateral, lo que permite cambiar de seccion sin perder "
        "el contexto de trabajo.")

doc.add_heading("3.1 Menu lateral del administrador", level=2)
p(doc, "El menu muestra el encabezado VORTEX con la leyenda Administrador y cuatro "
        "opciones de navegacion. La opcion activa se resalta en azul:")
data_table(doc, [
    ["Opcion", "Pantalla a la que conduce"],
    ["Inicio", "Panel de resumen con los indicadores generales del taller y los "
               "proximos servicios pendientes."],
    ["Clientes", "Listado completo de clientes y solicitudes, con las acciones de "
                 "gestion sobre cada una."],
    ["Inventario / Piezas", "Catalogo de piezas con precios y existencias."],
    ["Reportes", "Estadisticas de ventas, piezas mas vendidas y alertas de stock bajo."],
], widths=[1.7, 4.7], font_size=9.5)
p(doc, "", space_after=2)
p(doc, "En la parte inferior el menu identifica al usuario conectado como Administrador "
        "e incluye el boton Cerrar sesion, que finaliza la sesion y devuelve el sistema "
        "a la pantalla del cliente.", space_before=4)

# ---------------------------------------------- 3.2 INICIO
doc.add_page_break()
doc.add_heading("3.2 Inicio: panel de resumen", level=2)
p(doc, "Es la pantalla de bienvenida del administrador. Presenta el estado general del "
        "taller en un solo vistazo mediante cuatro indicadores y una tabla de servicios "
        "proximos.")

figura(doc, 2, "Panel de Inicio del modulo Administrador")

doc.add_heading("3.2.1 Indicadores generales", level=3)
p(doc, "Las cuatro tarjetas superiores resumen la operacion. Cada una se distingue por "
        "el color de su franja lateral:")
data_table(doc, [
    ["Indicador", "Significado", "Ejemplo"],
    ["Clientes registrados", "Numero total de clientes dados de alta en el sistema.", "30"],
    ["Tipos de pieza", "Cantidad de piezas distintas del catalogo y, entre parentesis, "
     "la suma de todas las existencias.", "149 (3311 en stock)"],
    ["Ventas realizadas", "Total de ventas registradas historicamente.", "12"],
    ["Ingresos totales", "Suma del importe de todas las ventas registradas.", "$16,735.00"],
], widths=[1.5, 3.6, 1.3], font_size=9.5)
p(doc, "", space_after=2)
p(doc, "La lectura conjunta de estas tarjetas permite evaluar el tamano de la cartera de "
        "clientes, la profundidad del inventario y el desempeno comercial acumulado.",
  space_before=4)

doc.add_heading("3.2.2 Proximos servicios (pendientes)", level=3)
p(doc, "Debajo de los indicadores se muestra la tabla Proximos servicios, cuyo titulo "
        "incluye el numero de solicitudes pendientes de atender. Solo lista las "
        "solicitudes que aun no se han resuelto, ordenadas por fecha y hora, de manera "
        "que la primera fila corresponde al servicio mas inmediato. Sus columnas son:")
data_table(doc, [
    ["Columna", "Contenido"],
    ["ID", "Numero de pedido de la solicitud, el mismo que recibio el cliente al enviarla."],
    ["Cliente", "Nombre de la persona o negocio solicitante."],
    ["Telefono", "Numero de contacto capturado en la solicitud."],
    ["Direccion", "Domicilio donde debe prestarse el servicio."],
    ["Fecha", "Dia programado o de registro del servicio."],
    ["Hora", "Hora prevista de atencion."],
], widths=[1.1, 5.3], font_size=9.5)
p(doc, "", space_after=2)
p(doc, "Esta tabla funciona como agenda de trabajo del dia: reune el dato de contacto y "
        "la ubicacion junto al horario, que es la informacion necesaria para organizar "
        "las visitas tecnicas.", space_before=4)

# ---------------------------------------------- 3.3 CLIENTES
doc.add_page_break()
doc.add_heading("3.3 Clientes y solicitudes", level=2)
p(doc, "Concentra la gestion completa de las solicitudes. A diferencia de la tabla de "
        "Inicio, aqui se listan todas las solicitudes sin importar su estado, e incorpora "
        "la barra de acciones que permite operar sobre ellas.")

figura(doc, 3, "Pantalla Clientes y solicitudes")

doc.add_heading("3.3.1 Columnas de la tabla", level=3)
data_table(doc, [
    ["Columna", "Contenido"],
    ["ID", "Numero de pedido de la solicitud."],
    ["Nombre", "Cliente solicitante."],
    ["Telefono", "Numero de contacto."],
    ["Direccion", "Domicilio del servicio."],
    ["Estado", "Etapa en la que se encuentra la solicitud: PENDIENTE, EN PROCESO o ATENDIDA."],
    ["Fecha", "Dia asociado a la solicitud."],
    ["Hora", "Hora asociada a la solicitud."],
], widths=[1.1, 5.3], font_size=9.5)
p(doc, "", space_after=2)
p(doc, "El listado se agrupa por estado, presentando primero las solicitudes PENDIENTE, "
        "despues las que estan EN PROCESO y al final las ATENDIDA. Asi el trabajo por "
        "hacer queda siempre en la parte superior de la pantalla.", space_before=4)

doc.add_heading("3.3.2 Estados de una solicitud", level=3)
p(doc, "El estado describe el avance del servicio y es el eje del seguimiento:")
data_table(doc, [
    ["Estado", "Significado"],
    ["PENDIENTE", "La solicitud fue recibida pero el taller aun no comienza a atenderla. "
                  "Es el estado con el que nace toda solicitud enviada por un cliente."],
    ["EN PROCESO", "El servicio ya se esta ejecutando: se asigno un tecnico, se "
                   "diagnostico el equipo o se esta realizando la reparacion."],
    ["ATENDIDA", "El servicio concluyo. La solicitud se conserva en el historial pero ya "
                 "no requiere accion."],
], widths=[1.3, 5.1], font_size=9.5)
p(doc, "", space_after=2)

doc.add_heading("3.3.3 Consultar el detalle de una solicitud", level=3)
p(doc, "El texto guia bajo el titulo lo indica: al hacer doble clic sobre un cliente de "
        "la tabla se abre el detalle de su solicitud. Ahi puede consultarse la "
        "informacion completa del registro, incluido el motivo del servicio que el "
        "cliente describio al enviarlo, que por su extension no cabe en el listado.")

doc.add_heading("3.3.4 Barra de acciones", level=3)
p(doc, "Los botones de la parte superior derecha operan sobre la solicitud previamente "
        "seleccionada en la tabla, con excepcion de Actualizar:")
data_table(doc, [
    ["Boton", "Funcion"],
    ["Cotizar", "Genera la cotizacion del servicio de la solicitud seleccionada, "
                "incorporando las piezas del inventario que se requieran y su precio."],
    ["Imprimir ticket", "Emite el comprobante del servicio con los datos del cliente y "
                        "los conceptos cobrados, para entregarlo como constancia."],
    ["Cambiar estado", "Hace avanzar la solicitud a la siguiente etapa, por ejemplo de "
                       "PENDIENTE a EN PROCESO y despues a ATENDIDA."],
    ["Eliminar", "Borra del sistema la solicitud seleccionada. Se usa para descartar "
                 "registros duplicados o capturados por error."],
    ["Actualizar", "Vuelve a leer la informacion y refresca la tabla para mostrar las "
                   "solicitudes nuevas que hayan llegado desde el modulo Cliente."],
], widths=[1.3, 5.1], font_size=9.5)
p(doc, "", space_after=2)
nota(doc, "Antes de pulsar Cotizar, Imprimir ticket, Cambiar estado o Eliminar es "
          "necesario seleccionar la fila del cliente correspondiente, ya que estas "
          "acciones se aplican al registro activo.")

# ---------------------------------------------- 3.4 INVENTARIO
doc.add_page_break()
doc.add_heading("3.4 Inventario de piezas", level=2)
p(doc, "Es el catalogo de refacciones y equipos del taller. Cumple dos funciones: sirve "
        "como lista de precios al momento de cotizar un servicio y como control de "
        "existencias para saber de que material se dispone.")

figura(doc, 4, "Pantalla Inventario de piezas")

doc.add_heading("3.4.1 Columnas del catalogo", level=3)
data_table(doc, [
    ["Columna", "Contenido"],
    ["Codigo", "Clave corta que identifica la pieza de forma unica. Se forma con un "
               "prefijo por familia y un sufijo por variante, por ejemplo BOM-100 para "
               "una bomba centrifuga de 1 HP o ARR-TER para un arrancador termico."],
    ["Nombre", "Denominacion comercial de la pieza, normalmente con su capacidad o "
               "potencia, como Bomba centrifuga 1.5 HP."],
    ["Descripcion", "Detalle tecnico complementario que precisa material, uso o "
                    "caracteristicas del componente."],
    ["Precio", "Importe unitario de venta. Es el valor que se traslada a la cotizacion."],
    ["Stock", "Unidades disponibles en el almacen."],
], widths=[1.2, 5.2], font_size=9.5)
p(doc, "", space_after=2)
p(doc, "El catalogo se presenta ordenado por nombre, lo que agrupa de forma natural las "
        "piezas de una misma familia y facilita localizarlas al recorrer la lista. "
        "Ejemplos de registros:", space_before=4)
data_table(doc, [
    ["Codigo", "Nombre", "Precio", "Stock"],
    ["ABR-MAN", "Abrazadera para manguera", "$18.00", "110"],
    ["ARR-TER", "Arrancador termico", "$540.00", "11"],
    ["BAT-100", "Bateria ciclo profundo", "$4,300.00", "5"],
    ["BOM-100", "Bomba centrifuga 1 HP", "$2,450.00", "4"],
    ["BOM-SOL2", "Bomba solar sumergible", "$14,500.00", "1"],
], widths=[1.2, 3.0, 1.2, 1.0], font_size=9.5)
p(doc, "", space_after=2)

doc.add_heading("3.4.2 Acciones sobre el inventario", level=3)
data_table(doc, [
    ["Boton", "Funcion"],
    ["+ Agregar", "Da de alta una pieza nueva en el catalogo. Requiere definir su "
                  "codigo, nombre, descripcion, precio y existencia inicial."],
    ["Editar", "Modifica los datos de la pieza seleccionada. Es la via para actualizar "
               "un precio o corregir la cantidad en stock."],
    ["Eliminar", "Retira del catalogo la pieza seleccionada."],
], widths=[1.3, 5.1], font_size=9.5)
p(doc, "", space_after=2)
nota(doc, "Mantener el stock y los precios al dia es lo que garantiza que las "
          "cotizaciones y los reportes reflejen la realidad del taller, ya que ambos se "
          "alimentan de este catalogo.")

# ---------------------------------------------- 3.5 REPORTES
doc.add_page_break()
doc.add_heading("3.5 Reportes y estadisticas", level=2)
p(doc, "Reune la informacion de analisis del taller. Combina cuatro indicadores de "
        "ventas con tres tablas que responden a preguntas concretas: que se vende mas, "
        "como evolucionan las ventas y que material esta por agotarse.")

figura(doc, 5, "Pantalla Reportes y estadisticas")

doc.add_heading("3.5.1 Indicadores de venta", level=3)
data_table(doc, [
    ["Indicador", "Significado", "Ejemplo"],
    ["Ventas totales", "Numero de ventas registradas desde el inicio de operacion.", "12"],
    ["Ingresos totales", "Importe acumulado de todas esas ventas.", "$16,735.00"],
    ["Ventas de la semana", "Importe vendido en la semana en curso y, entre parentesis, "
     "el numero de operaciones.", "$0.00 (0)"],
    ["Ventas del mes", "Importe vendido en el mes en curso y el numero de operaciones "
     "correspondiente.", "$10,785.00 (7)"],
], widths=[1.5, 3.6, 1.3], font_size=9.5)
p(doc, "", space_after=2)
p(doc, "Los indicadores semanal y mensual permiten comparar el ritmo reciente con el "
        "acumulado. Un valor en cero en la semana, con ventas presentes en el mes, "
        "senala simplemente que en los ultimos dias no se ha registrado ninguna "
        "operacion.", space_before=4)

doc.add_heading("3.5.2 Piezas mas vendidas", level=3)
p(doc, "Ordena las piezas segun su rotacion. Presenta el nombre de la Pieza, las "
        "Unidades vendidas y los Ingresos que ha generado cada una. Es la referencia "
        "para decidir que material conviene mantener siempre disponible.")
data_table(doc, [
    ["Pieza", "Unidades vendidas", "Ingresos"],
    ["Impulsor 1/2 HP", "2", "$460.00"],
    ["Presostato automatico", "2", "$360.00"],
    ["Cinta teflon", "2", "$30.00"],
], widths=[3.0, 1.9, 1.5], font_size=9.5)
p(doc, "", space_after=2)
p(doc, "Comparar unidades contra ingresos distingue las piezas de alta rotacion y bajo "
        "importe de aquellas que, con pocas unidades, aportan mas dinero.", space_before=4)

doc.add_heading("3.5.3 Ventas por mes", level=3)
p(doc, "Agrupa las ventas por periodo mensual, mostrando el Mes en formato ano-mes, el "
        "N. de ventas y el Total facturado. Permite observar la tendencia del negocio a "
        "lo largo del tiempo.")
data_table(doc, [
    ["Mes", "N. de ventas", "Total"],
    ["2026-07", "7", "$10,785.00"],
    ["2026-06", "2", "$3,455.00"],
], widths=[2.0, 2.2, 2.2], font_size=9.5)
p(doc, "", space_after=2)

doc.add_heading("3.5.4 Piezas con stock bajo", level=3)
p(doc, "Lista unicamente las piezas cuya existencia es igual o menor a 10 unidades, "
        "umbral indicado en el propio titulo de la seccion. Muestra el Codigo, la Pieza "
        "y el Stock restante. Funciona como alerta de reabastecimiento: las piezas que "
        "aparecen aqui son las que deben pedirse al proveedor antes de quedarse sin "
        "material.")
data_table(doc, [
    ["Codigo", "Pieza", "Stock"],
    ["MOT-750", "Motor 7.5 HP", "1"],
    ["BOM-SUM5", "Bomba sumergible 5 HP", "1"],
], widths=[1.5, 3.4, 1.5], font_size=9.5)
p(doc, "", space_after=2)

doc.add_heading("3.5.5 Actualizar", level=3)
p(doc, "El boton Actualizar de la esquina superior derecha recalcula todos los "
        "indicadores y tablas de la pantalla. Conviene pulsarlo despues de registrar "
        "ventas o de modificar el inventario para consultar cifras vigentes.")

# ================================================== 4. FLUJO COMPLETO
doc.add_page_break()
doc.add_heading("4. Flujo de trabajo completo", level=1)
p(doc, "Las pantallas descritas se articulan en un recorrido unico, desde que el cliente "
        "reporta la falla hasta que el servicio queda cerrado y reflejado en los reportes:")

data_table(doc, [
    ["Paso", "Pantalla", "Que ocurre"],
    ["1", "Cliente / Solicitar servicio",
     "El cliente captura nombre, telefono, direccion y motivo, y envia la solicitud. "
     "Recibe su numero de pedido."],
    ["2", "Administrador / Inicio",
     "La solicitud aparece como PENDIENTE en Proximos servicios y suma al contador de "
     "pendientes."],
    ["3", "Administrador / Clientes",
     "El administrador abre el detalle con doble clic, revisa el motivo del servicio y "
     "genera la cotizacion con Cotizar."],
    ["4", "Administrador / Inventario",
     "Se consultan precios y existencias de las piezas necesarias y se ajusta el stock "
     "con Editar."],
    ["5", "Administrador / Clientes",
     "Con Cambiar estado la solicitud pasa a EN PROCESO mientras se realiza el trabajo, "
     "y a ATENDIDA al concluirlo. Con Imprimir ticket se entrega el comprobante."],
    ["6", "Administrador / Reportes",
     "La operacion se incorpora a los indicadores de ventas, a las piezas mas vendidas y "
     "a las alertas de stock bajo."],
], widths=[0.55, 1.85, 4.0], font_size=9.5)
p(doc, "", space_after=2)

# ================================================== 5. GLOSARIO
doc.add_heading("5. Glosario", level=1)
data_table(doc, [
    ["Termino", "Definicion"],
    ["Solicitud de servicio", "Peticion registrada por un cliente para que el taller "
                              "revise o repare una bomba de agua."],
    ["Numero de pedido", "Identificador consecutivo que el sistema asigna a cada "
                         "solicitud al enviarse. Corresponde a la columna ID."],
    ["Estado", "Etapa de avance de una solicitud: PENDIENTE, EN PROCESO o ATENDIDA."],
    ["Codigo de pieza", "Clave corta y unica que identifica una refaccion en el "
                        "inventario, como BOM-100."],
    ["Stock", "Numero de unidades de una pieza disponibles en el almacen."],
    ["Stock bajo", "Condicion de una pieza cuya existencia es menor o igual a 10 "
                   "unidades; el sistema la reporta como alerta de reabastecimiento."],
    ["Cotizacion", "Documento que estima el costo de un servicio a partir de las piezas "
                   "y su precio de catalogo."],
    ["Ticket", "Comprobante impreso del servicio realizado y de los conceptos cobrados."],
], widths=[1.6, 4.8], font_size=9.5)

# ---------------------------------------------------------------- guardar
salida = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                      "Manual_de_Usuario_VORTEX.docx")
doc.save(salida)

print("Documento generado:", salida)
print()
print("Asignacion de capturas:")
titulos = {
    1: "Modulo Cliente - Solicitar servicio",
    2: "Administrador - Inicio (panel de resumen)",
    3: "Administrador - Clientes y solicitudes",
    4: "Administrador - Inventario de piezas",
    5: "Administrador - Reportes y estadisticas",
}
for num, archivo in MAPEO:
    print(f"  Figura {num}  {titulos.get(num, ''):<45} <-  {archivo}")
faltan = [n for n, a in MAPEO if a.startswith("--")]
if faltan:
    print()
    print(f"Figuras sin imagen: {faltan}. Coloque los archivos en ./capturas/")
    print("con los nombres figura1.png ... figura5.png y ejecute de nuevo.")
