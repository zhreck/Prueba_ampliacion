import io
from datetime import datetime

from flask import Flask, render_template, request, send_file, flash, redirect, url_for

import engine
import correlativo
import salida_sap

app = Flask(__name__)
app.secret_key = "cambiar-esta-clave-en-produccion"


@app.route("/", methods=["GET"])
def index():
    tipos = engine.listar_tipos()
    return render_template("index.html", tipos=tipos)


@app.route("/procesar", methods=["POST"])
def procesar():
    tipo_id = request.form.get("tipo")
    archivo = request.files.get("archivo")
    generar_sap = request.form.get("generar_sap") == "on"  # Checkbox para habilitar SAP

    if not tipo_id:
        flash("Selecciona un tipo de ampliación.", "error")
        return redirect(url_for("index"))

    if not archivo or archivo.filename == "":
        flash("Sube un archivo Excel con el input.", "error")
        return redirect(url_for("index"))

    try:
        resultado_df = engine.procesar(tipo_id, archivo)
    except (engine.TipoNoEncontradoError, engine.InputInvalidoError) as e:
        flash(str(e), "error")
        return redirect(url_for("index"))

    avisos = resultado_df.attrs.get("avisos", [])
    
    # Si se solicita generar formato SAP
    df_salida = resultado_df
    nombre_sheet = "AMPLIADO"
    
    if generar_sap:
        try:
            # Obtener la filial del input
            if "FILIAL CODIGO" in resultado_df.columns:
                filiales_unicas = resultado_df["FILIAL CODIGO"].unique()
                if len(filiales_unicas) > 1:
                    flash("⚠️ Entrada con múltiples filiales. Usando la primera.", "warning")
                filial = str(filiales_unicas[0]).strip()
            else:
                flash("⚠️ Columna FILIAL CODIGO no encontrada. No se puede generar SAP.", "warning")
                generar_sap = False
            
            if generar_sap:
                # Mapear filial a tipo de material SAP
                tipo_material_sap = salida_sap.filial_a_tipo_material(filial)
                
                # Cargar tabla de referencia centro->cebe si existe para VC00
                centro_cebe_lookup = None
                if filial in ["VC00", "VD00", "VE00"]:
                    try:
                        ref_df = engine.pd.read_excel(
                            engine.BASE_DIR / "data" / "reference" / "centros_cebe_vc00.xlsx"
                        )
                        centro_cebe_lookup = dict(zip(ref_df["CENTRO"], ref_df["COD_CEBE"]))
                    except Exception as e:
                        avisos.append(f"⚠️ No se pudo cargar tabla de referencia CEBE: {e}")
                
                # Aplicar plantilla SAP
                df_salida, metadatos_sap = salida_sap.aplicar_plantilla_sap(
                    df_ampliado=resultado_df,
                    tipo_material=tipo_material_sap,
                    centro_cebe_lookup=centro_cebe_lookup,
                )
                
                nombre_sheet = f"SAP_{tipo_material_sap}"
                
                # Avisos sobre campos pendientes
                if metadatos_sap.get("campos_pendientes"):
                    campos_pend = ", ".join(metadatos_sap["campos_pendientes"][:3])
                    avisos.append(f"ℹ️ Campos pendientes de lookup: {campos_pend}... (Ver documentación)")
                
                if metadatos_sap.get("columnas_vacias"):
                    cols_vacias_count = len(metadatos_sap["columnas_vacias"])
                    avisos.append(f"ℹ️ {cols_vacias_count} columnas vacías en formato SAP (esperado)")
                
        except salida_sap.FormatoSAPError as e:
            flash(f"Error en conversión SAP: {e}", "error")
            generar_sap = False
    
    for a in avisos:
        flash(a, "warning")

    buffer = io.BytesIO()
    df_salida.to_excel(buffer, index=False, sheet_name=nombre_sheet)
    buffer.seek(0)

    tipo_sufijo = "SAP" if generar_sap else "ampliado"
    nombre_salida = f"{tipo_id}_{tipo_sufijo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return send_file(
        buffer,
        as_attachment=True,
        download_name=nombre_salida,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/historial")
def historial():
    tipos_cfg = {t["id"]: engine.cargar_config(t["id"]) for t in engine.listar_tipos()}
    estados = []
    for tipo_id, cfg in tipos_cfg.items():
        corr_cfg = cfg.get("correlativo", {})
        if corr_cfg.get("enabled"):
            try:
                estado = correlativo.estado_contador(
                    corr_cfg["nombre"], corr_cfg.get("range_min"), corr_cfg.get("range_max")
                )
                estados.append((tipo_id, estado))
            except Exception:
                pass  # Skip si no se puede obtener estado
    
    registros = correlativo.historial(limit=200)
    return render_template("historial.html", estados=estados, registros=registros)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)

