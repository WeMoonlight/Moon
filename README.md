from flask import Flask, render_template, request, send_file, after_this_request
from docx import Document
from docx.shared import Pt
from datetime import datetime, timedelta
import os

app = Flask(__name__)

def set_times_new_roman(run):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        nombre = request.form["nombre"].upper()
        curp = request.form["curp"].upper()
        categoria = request.form["categoria"].upper()
        fecha_inicio = request.form["fecha_inicio"]

        fecha_dt = datetime.strptime(fecha_inicio, "%Y-%m-%d")
        fecha_vencimiento = fecha_dt + timedelta(days=365)

        fecha_inicio_str = fecha_dt.strftime("%Y%m%d")
        fecha_vencimiento_str = fecha_vencimiento.strftime("%Y%m%d")

        doc = Document("DC3_ICA_FLUOR_BASE.docx")

        # Reemplazar campos en todas las tablas de todas las hojas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            # Nombre
                            if "{NOMBRE}" in run.text:
                                run.text = run.text.replace("{NOMBRE}", nombre)
                                set_times_new_roman(run)
                            # Categoría
                            if "{CATEGORIA}" in run.text:
                                run.text = run.text.replace("{CATEGORIA}", categoria)
                                set_times_new_roman(run)

                        # CURP (campo por campo)
                        for i, c in enumerate(curp[:18]):
                            curp_key = f"{{curp {i+1}}}"
                            if curp_key in paragraph.text:
                                paragraph.text = paragraph.text.replace(curp_key, c)

                        # Fechas inicio y vencimiento (cada dígito)
                        for i in range(8):
                            fi_key = f"{{FECHA DE INICIO {i+1}}}"
                            fv_key = f"{{FECHA DE VENCIMIENTO {i+1}}}"
                            if fi_key in paragraph.text:
                                paragraph.text = paragraph.text.replace(fi_key, fecha_inicio_str[i])
                            if fv_key in paragraph.text:
                                paragraph.text = paragraph.text.replace(fv_key, fecha_vencimiento_str[i])

        output_docx = f"DC3_{nombre.replace(' ', '_')}_MULTI.docx"
        doc.save(output_docx)

        @after_this_request
        def cleanup(response):
            try:
                os.remove(output_docx)
            except Exception:
                pass
            return response

        return send_file(output_docx, as_attachment=True)

    return render_template("form.html")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")
