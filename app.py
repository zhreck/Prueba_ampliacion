import io
from datetime import datetime

import pandas as pd
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


@app.route("/plantilla/<tipo_id>")
def plantilla(tipo_id):
    try:
        wb = engine.generar_plantilla_vacia(tipo_id)
    except engine.TipoNoEncontradoError as e:
        flash(str(e), "error")
        return redirect(url_for("index"))

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"plantilla_{tipo_id}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


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
    df_pendientes = None

    cfg_tipo = engine.cargar_config(tipo_id)
    if generar_sap and not cfg_tipo.get("sap_disponible", True):
        flash(f"⚠️ {cfg_tipo.get('sap_no_disponible_motivo', 'Formato SAP no disponible para este tipo.')}", "error")
        return redirect(url_for("index"))

    if generar_sap:
        if "FILIAL CODIGO" not in resultado_df.columns:
            flash("⚠️ Columna FILIAL CODIGO no encontrada. No se puede generar SAP.", "warning")
            generar_sap = False
        else:
            filiales_unicas = resultado_df["FILIAL CODIGO"].unique()
            if len(filiales_unicas) > 1:
                flash("⚠️ Entrada con múltiples filiales en el mismo archivo. Solo se puede generar SAP para una filial a la vez.", "error")
                return redirect(url_for("index"))
            filial = str(filiales_unicas[0]).strip()

            try:
                tipo_material_sap = salida_sap.filial_a_tipo_material(filial)
                df_salida, metadatos_sap = salida_sap.aplicar_plantilla_sap(
                    df_ampliado=resultado_df,
                    tipo_material=tipo_material_sap,
                )
                nombre_sheet = f"SAP_{tipo_material_sap}"

                pendientes = metadatos_sap.get("campos_pendientes", {})
                if pendientes:
                    avisos.append(f"ℹ️ {len(pendientes)} columnas pendientes de negocio (ver hoja PENDIENTES).")
                    df_pendientes = pd.DataFrame(
                        {"Columna SAP pendiente": list(pendientes.keys()), "Motivo / nota del negocio": list(pendientes.values())}
                    )

                obligatorios_vacios = metadatos_sap.get("columnas_obligatorias_vacias", [])
                faltantes_no_pendientes = [c for c in obligatorios_vacios if c not in pendientes]
                if faltantes_no_pendientes:
                    avisos.append(
                        f"⚠️ Sin dato (no pendiente de negocio, revisar tabla de referencia): "
                        f"{', '.join(faltantes_no_pendientes)}"
                    )
            except salida_sap.FormatoSAPError as e:
                flash(f"Error en conversión SAP: {e}", "error")
                return redirect(url_for("index"))

    for a in avisos:
        flash(a, "warning")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_salida.to_excel(writer, index=False, sheet_name=nombre_sheet)
        if df_pendientes is not None:
            df_pendientes.to_excel(writer, index=False, sheet_name="PENDIENTES")
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

