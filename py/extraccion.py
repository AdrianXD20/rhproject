import fitz  # PyMuPDF
from jinja2 import Template
import base64

def extract_pdf_content(pdf_path):
    doc = fitz.open(pdf_path)
    content = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        # Extraer texto
        for block in page.get_text("dict")["blocks"]:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        # Manejar el color
                        color = span.get("color", 0)
                        if isinstance(color, (list, tuple)):  # Color como tupla RGB
                            rgb = tuple(int(c * 255) for c in color[:3])
                        else:  # Color como entero (por ejemplo, 0xRRGGBB)
                            r = (color >> 16) & 255
                            g = (color >> 8) & 255
                            b = color & 255
                            rgb = (r, g, b)
                        content.append({
                            'type': 'text',
                            'text': span["text"],
                            'x0': span["bbox"][0] * 1.33,  # Escala a píxeles
                            'y0': span["bbox"][1] * 1.33,
                            'size': span["size"],
                            'font': span["font"],
                            'color': f"rgb{rgb}"
                        })
        # Extraer imágenes
        for img in page.get_images(full=True):
            xref = img[0]
            base_image = doc.extract_image(xref)
            content.append({
                'type': 'image',
                'data': base64.b64encode(base_image["image"]).decode('utf-8'),
                'x0': img[2] * 1.33,  # Índice correcto para x0
                'y0': img[3] * 1.33,  # Índice correcto para y0
                'width': img[2] * 1.33,  # Índice correcto para ancho
                'height': img[3] * 1.33  # Índice correcto para altura
            })
        # Extraer campos de formulario (si los hay)
        for annot in page.annots():
            if annot.type[1] == 7:  # 7 corresponde a 'Widget' (campo de formulario)
                content.append({
                    'type': 'field',
                    'name': annot.info.get("id", "unknown"),
                    'x0': annot.rect[0] * 1.33,
                    'y0': annot.rect[1] * 1.33,
                    'width': (annot.rect[2] - annot.rect[0]) * 1.33,
                    'height': (annot.rect[3] - annot.rect[1]) * 1.33
                })
        # Extraer líneas de tablas (gráficos vectoriales)
        for path in page.get_drawings():
            for item in path["items"]:
                if item[0] == "l":  # Línea
                    x0, y0 = item[1]  # Punto inicial
                    x1, y1 = item[2]  # Punto final
                    # Escalar coordenadas
                    x0, y0, x1, y1 = x0 * 1.33, y0 * 1.33, x1 * 1.33, y1 * 1.33
                    # Determinar si es horizontal o vertical
                    if abs(y0 - y1) < 1:  # Línea horizontal
                        content.append({
                            'type': 'line',
                            'x0': min(x0, x1),
                            'y0': y0,
                            'width': abs(x1 - x0),
                            'height': 1,
                            'color': 'black'  # Puedes extraer el color si está disponible
                        })
                    elif abs(x0 - x1) < 1:  # Línea vertical
                        content.append({
                            'type': 'line',
                            'x0': x0,
                            'y0': min(y0, y1),
                            'width': 1,
                            'height': abs(y1 - y0),
                            'color': 'black'
                        })
    return content

html_template = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Formulario PDF a HTML</title>
    <style>
        body { font-family: Arial, sans-serif; position: relative; margin: 0; }
        .pdf-text { position: absolute; }
        .pdf-field { position: absolute; border: 1px solid #ccc; box-sizing: border-box; }
        .pdf-image { position: absolute; }
        .pdf-line { position: absolute; } /* Color se define inline */
    </style>
</head>
<body>
    {% for item in content %}
        {% if item.type == 'image' %}
            <img class="pdf-image" src="data:image/png;base64,{{ item.data }}"
                 style="left: {{ item.x0 }}px; top: {{ item.y0 }}px; width: {{ item.width }}px; height: {{ item.height }}px;">
        {% elif item.type == 'text' %}
            <div class="pdf-text" style="left: {{ item.x0 }}px; top: {{ item.y0 }}px; font-size: {{ item.size }}px; font-family: '{{ item.font | replace('Helvetica', 'Helvetica, Arial, sans-serif') | replace('Times', 'Times New Roman, serif') }}', Arial; color: {{ item.color }};">
                {{ item.text }}
            </div>
        {% elif item.type == 'field' %}
            <input type="text" name="{{ item.name }}" class="pdf-field"
                   style="left: {{ item.x0 }}px; top: {{ item.y0 }}px; width: {{ item.width }}px; height: {{ item.height }}px;">
        {% elif item.type == 'line' %}
            <div class="pdf-line" style="left: {{ item.x0 }}px; top: {{ item.y0 }}px; width: {{ item.width }}px; height: {{ item.height }}px; background-color: {{ item.color }};"></div>
        {% endif %}
    {% endfor %}
</body>
</html>
"""

def generate_html(content, output_path):
    template = Template(html_template)
    html_content = template.render(content=content)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

# Ejemplo de uso
pdf_path = "C:/Users/carav/Downloads/Quinto Cuatrimestre/rh/py/formato.pdf"  # Ajusta la ruta si es necesario
output_html = "lol.html"
content = extract_pdf_content(pdf_path)
generate_html(content, output_html)
print(f"HTML generado en: {output_html}")