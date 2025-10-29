import streamlit as st
import pandas as pd
import time, re
from datetime import date, datetime
from dateutil import parser
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options


# ===============================================================
# 🔹 CONFIGURA AQUÍ TUS TABLEROS (nombre_tablero, url)
# ===============================================================

REPORTS = [
    {"nombre_tablero":"1.EIB", "url":"https://app.powerbi.com/view?r=eyJrIjoiM2Y0MTgyNDUtYWJiZi00N2E0LTk4NDEtYTljYTk5MjUzYzA3IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"2.Programa Formativo para el uso flexible de los espacios de aprendizaje","url":"https://app.powerbi.com/view?r=eyJrIjoiZGM5MGYxYTYtOGI2NC00YjFmLTg2Y2QtZGU0OWFmZmEyYmRkIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"3.PDP","url":"https://app.powerbi.com/view?r=eyJrIjoiZGM5MGYxYTYtOGI2NC00YjFmLTg2Y2QtZGU0OWFmZmEyYmRkIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"4.PDP-ACTUALIZACION","url":"https://app.powerbi.com/view?r=eyJrIjoiNTNiZTFhOTItOTkwMy00M2NiLTljMTQtMzE0OTQxYzlhMzc1IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"5.CONDORCANQUI","url":"https://app.powerbi.com/view?r=eyJrIjoiYTk5MzIyNmEtYWRiOC00M2Q0LTk0YzMtZDk0Zjk0NDY0ZDQ0IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"6.MULTIGRADO-NIVEL 1","url":"https://app.powerbi.com/view?r=eyJrIjoiNTRlY2M1NDQtNzA0Mi00M2I4LTljMzQtM2QzZjg5NDQxMjExIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"7.MULTIGRADO-NIVEL 2","url":"https://app.powerbi.com/view?r=eyJrIjoiYjNhMGQ2YWMtYjk5MC00MTBjLWJmN2YtMjhhOGQ3ZTBhYWFlIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"8.SANTA ROSA","url":"https://app.powerbi.com/view?r=eyJrIjoiZDYxMTFjNzEtNDRjMC00ZDQ2LWJmYzUtNmY3ZjQ0ZjU2ZjBjIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"9.BICENTENARIO","url":"https://app.powerbi.com/view?r=eyJrIjoiMGNkMThmZDYtYmQzZS00YWUyLTg1NTYtOTZiZWRkNmE2ZGRkIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"10.PID","url":"https://app.powerbi.com/view?r=eyJrIjoiOWJhZGI5MDctNTUwOS00NDAxLTkxMDUtYTUzOGQxYjFiM2Y3IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"11.DIPLOMADO","url":"https://app.powerbi.com/view?r=eyJrIjoiMzY2ODcxNmItY2UxZi00Njg4LWEzMTQtYzVkMTkyMTk1ZTJlIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"12.BIAE","url":"https://app.powerbi.com/view?r=eyJrIjoiNmUwZWQ5MmEtMDZlMy00OGJkLTk0ODEtOTFlNzQxZTJiNWMzIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"13.PROGR DE FORM EN COMPET LECT Y CONV CON APORTES EN NEUROEDUCACION","url":"https://app.powerbi.com/view?r=eyJrIjoiYmEyZGE2ZjEtOWE2Ni00NDQxLWExNGQtMDdmZDA4MDYxMjdkIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"14.PROGR DE FORM EN COMPET LECT Y CONV CON APORTES EN NEUROEDUCACION - DRE CALLAO","url":"https://app.powerbi.com/view?r=eyJrIjoiODdhMDMwYzYtNDM5Yi00YTQxLTk5MWQtNDZkZmUwOTQyN2Y5IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"15.PAD DIBRED","url":"https://app.powerbi.com/view?r=eyJrIjoiYjkyYTU0NGUtY2ZmZi00ZjQ1LWIyMDMtZjVmYmJmYTczNDFlIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"16.PAD OAM","url":"https://app.powerbi.com/view?r=eyJrIjoiYTdmNDNmNWYtZWJkYy00ZjkxLThhZGItNTkxMDI3M2EyZDY5IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"17.OAM - APLICANDO LOS PRINCIPIOS DEL DISEÑO UNIVERSAL PARA EL APRENDIZAJE 2.2 AL 3.0 DUA","url":"https://app.powerbi.com/view?r=eyJrIjoiZmFjN2FmZTYtYTc1NS00NjU5LWJhMGEtYzAyMThlN2QxY2ZlIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"18.OAM - INDECOPI","url":"https://app.powerbi.com/view?r=eyJrIjoiZTBiNTFlMjgtYjczOS00MDJjLTg2ZmItNGY5ODBkMmM5NTE0IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"19.OAM - EDUCACION FINANCIERA SBS","url":"https://app.powerbi.com/view?r=eyJrIjoiNjM1NmIzNGMtMTlkMS00MmFmLWE4OTktZGI5MDBjMjIzY2FkIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"20.OAM - IA EN LA PRACTICA DOCENTE","url":"https://app.powerbi.com/view?r=eyJrIjoiZjIxY2YyN2EtM2Y1Ny00ZWExLWE1YjItZDkzYzcyOWVmZGE4IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"21.OAM - APRENDIZAJE A NIVEL REAL - INICIAL","url":"https://app.powerbi.com/view?r=eyJrIjoiMWNlYWViZjEtYzdiOC00MGRkLWJmNDctZmU0NjdhMWNkMDE2IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"22.OAM - APRENDIZAJE A NIVEL REAL - PRIMARIA","url":"https://app.powerbi.com/view?r=eyJrIjoiYzliYTcwMmUtYmE3YS00MDMyLTkxM2UtNmU5MmZjYTNmYTc2IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"23.OAM - APRENDIZAJE A NIVEL REAL - SECUNDARIA MATEMATICA","url":"https://app.powerbi.com/view?r=eyJrIjoiZGJhNDY1OWUtNzc0OC00NjgzLWEzYTItNmU2ZjhkYmFhMzYzIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"24.OAM - APRENDIZAJE A NIVEL REAL - SECUNDARIA COMUNICACION","url":"https://app.powerbi.com/view?r=eyJrIjoiMWM1ZTVhYWQtMDQ2Yi00YWQzLTk5YzItOTU5YWUzNWEwNTY1IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"25.OAM - APRENDIZAJE A NIVEL REAL - SECUNDARIA CYT","url":"https://app.powerbi.com/view?r=eyJrIjoiYjYwZWU0YzEtNjFjMC00NzFhLTg3MzMtMTdkMzk5NGNkYzcyIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"26.OAM - APRENDIZAJE A NIVEL REAL - SECUNDARIA DPCC","url":"https://app.powerbi.com/view?r=eyJrIjoiM2RlY2RhN2MtZDJhNy00YWJmLThiODQtMDY2ZjUxZjlkNDYwIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"27.OAM - LECTURA Y ESCRITURA EN LOS PRIMEROS AÑOS DE ESCOLARIDAD","url":"https://app.powerbi.com/view?r=eyJrIjoiZGExMjZhMzUtZDdiYi00OWM4LWEyNjQtNGE0ZjMyN2EyNWFiIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"28.OAM - LAS COMPETENCIAS MATEMATICAS EN LOS PRIMEROS AÑOS DE ESCOLARIDAD","url":"https://app.powerbi.com/view?r=eyJrIjoiOWUyNjYyNjItZGM4Zi00MmFhLWFmNGMtZWFhMGYyNjQ2NTBmIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"29.OAM - EDUCACION INCLUSIVA PARA ATENDER LA DIVERSIDAD","url":"https://app.powerbi.com/view?r=eyJrIjoiYzE1MmJlOGUtMDA2NC00MzFjLWE5YjItODkxODk4ZTU1NzY2IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"30.OAM - DISEÑO UNIVERSAL PARA EL APRENDIZAJE Y AJUSTES RAZONABLES","url":"https://app.powerbi.com/view?r=eyJrIjoiOTIxNmQxOTMtNzMzZC00MDI0LThmZmQtYThkNWQ0MDZiZDEwIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"31.REGIONALES - APURIMAC","url":"https://app.powerbi.com/view?r=eyJrIjoiNDRlOTEwYWMtNDFjMi00MTdlLWIyNWEtMTdmZTYzY2NiZmY5IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"32.REGIONALES - DRE PIURA","url":"https://app.powerbi.com/view?r=eyJrIjoiMmNmMzk2NTAtYzlkYS00OTA5LTkzNzktMzU4MDQzYWFkMTcxIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"33.REGIONALES - RECUAY ANCASH","url":"https://app.powerbi.com/view?r=eyJrIjoiOGEwNzFlZDAtZDdkYS00NzY2LTg3ZWEtN2RkZmVkY2RjYzhmIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"34.REGIONALES - DRE LIMA","url":"https://app.powerbi.com/view?r=eyJrIjoiYjUzMjc0NjAtY2ZjZC00ZmRkLTg2NjQtZWYxMDJiZmJkZGIzIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"35.FORMADOR DE FORMADORES","url":"https://app.powerbi.com/view?r=eyJrIjoiMTJhMjdmY2UtMWZlZC00YzUxLWJhM2QtMjNjY2FiMWVmMWE4IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"36.CFI - NIVEL 1 - ESCALAMIENTO","url":"https://app.powerbi.com/view?r=eyJrIjoiNjg4Yjg1NTMtN2YzZi00ZTkwLWJhYTEtYTgwMDgxNDVmYTE1IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"37.CFI - NIVEL 2 - CONTINUIDAD","url":"https://app.powerbi.com/view?r=eyJrIjoiY2MyZGNmMDUtNjc1MC00ZmFiLWIzMWUtYjgwMDUzZmQzYjdhIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"38.EBA","url":"https://app.powerbi.com/view?r=eyJrIjoiNWU2Nzc0MTItZWM1Zi00MjEzLWIxNzAtMWIyMGY2OTgwM2Y3IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"39.EDUCUNA","url":"https://app.powerbi.com/view?r=eyJrIjoiNzA3ZmIwYjktMTQzZC00ZGRiLWE3NjctYjBlZDRhZGVlOTkyIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"40.PROGRAMA DE FORMACION PARA LA INTEGRACION DE LAS TECNOLOGIAS DIGITALES","url":"https://app.powerbi.com/view?r=eyJrIjoiYTRiNmVkZjYtODVhMi00MzExLWE5NDYtY2YxMzg2NzczNjQzIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"41.PRONOEI - FOCALIZADO","url":"https://app.powerbi.com/view?r=eyJrIjoiMzdhNmE3MDItYWNhOC00NWU0LTkzODMtZWY3NTg2NjdkMjA2IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"42.PRONOEI - ESP DRE/UGEL","url":"https://app.powerbi.com/view?r=eyJrIjoiYThmYmZmYzktMGFlYy00YzIzLWE3NDAtM2U5YTUwYTc2MjAyIiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
    {"nombre_tablero":"43.PRONOEI - NO FOCALIZADO","url":"https://app.powerbi.com/view?r=eyJrIjoiYWI0Y2EwYzMtNjJhMC00NmFkLTkxM2UtYThhNmJjNzAxNzY3IiwidCI6IjE3OWJkZGE4LWQ5NjQtNDNmZi1hZDNiLTY3NDE4NmEyZmEyOCIsImMiOjR9"},
]
# ===============================================================


# ========================= FUNCIONES BASE =========================

def inicializar_navegador():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)


def buscar_fecha_en_fuente(html_total):
    """Busca fechas tipo 10/28/2025 6:22:44 AM en el HTML/SVG renderizado."""
    patron = r"(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(AM|PM|A\.M\.|P\.M\.|a\. m\.|p\. m\.)?)"
    m = re.search(patron, html_total)
    if not m:
        return None
    candidata = (
        m.group(1)
        .replace("a. m.", "AM")
        .replace("p. m.", "PM")
        .replace("A.M.", "AM")
        .replace("P.M.", "PM")
    )
    try:
        return parser.parse(candidata, dayfirst=False, fuzzy=True)
    except Exception:
        try:
            return parser.parse(candidata, dayfirst=True, fuzzy=True)
        except Exception:
            return None


def auditar_tablero(nombre, url):
    """Abre el dashboard, busca la fecha y devuelve su estado."""
    hoy = date.today()
    driver = inicializar_navegador()
    resultado = {
        "nombre_tablero": nombre,
        "url": url,
        "fecha_tablero": None,
        "estado": "NO SE PUDO LEER ⚠",
        "audit_runtime": datetime.now().isoformat(timespec="seconds"),
    }

    if not isinstance(url, str) or not url.startswith("http"):
        resultado["estado"] = "ERROR ⚠ URL no válida"
        return resultado

    try:
        driver.get(url)
        time.sleep(15)  # Espera para que cargue Power BI
        html_total = driver.page_source
        for svg in driver.find_elements(By.TAG_NAME, "svg"):
            html_total += svg.get_attribute("outerHTML") or ""
        dt_detectada = buscar_fecha_en_fuente(html_total)
        if dt_detectada:
            fecha_sola = dt_detectada.date()
            resultado["fecha_tablero"] = str(fecha_sola)
            resultado["estado"] = "AL DÍA ✅" if fecha_sola == hoy else "DESACTUALIZADO ❌"
    except Exception as e:
        resultado["estado"] = f"ERROR ⚠ {e}"
    finally:
        driver.quit()
    return resultado


# ========================= INTERFAZ STREAMLIT =========================

st.set_page_config(page_title="Auditoría Power BI", page_icon="📊", layout="wide")

st.title("📊 Auditoría Automática de Tableros Power BI")
st.caption("Verifica automáticamente si tus tableros publicados están actualizados a la fecha de hoy.")

# Mostrar lista de tableros cargados desde el código
st.subheader("📋 Lista de tableros configurados")
df_tableros = pd.DataFrame(REPORTS)
st.dataframe(df_tableros, use_container_width=True)

if st.button("🚀 Ejecutar auditoría"):
    resultados = []
    progreso = st.progress(0)
    total = len(REPORTS)

    for i, r in enumerate(REPORTS):
        nombre = r["nombre_tablero"]
        url = r["url"]
        resultado = auditar_tablero(nombre, url)
        resultados.append(resultado)
        progreso.progress((i + 1) / total)
        st.write(f"🔍 {nombre}: {resultado['estado']}")

    st.success("✅ Auditoría completada.")
    df_res = pd.DataFrame(resultados)
    df_res = df_res[["nombre_tablero", "fecha_tablero", "estado", "audit_runtime"]]

    st.subheader("📊 Resultados finales")
    st.dataframe(df_res, use_container_width=True)

 # ====== Exportar resultados ======
    import io
    from openpyxl import Workbook

    # Crear un archivo Excel en memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_res.to_excel(writer, sheet_name="Auditoria", index=False)
    datos_excel = output.getvalue()

    # Mostrar botones de descarga
    st.download_button(
        label="⬇️ Descargar resultados en Excel (.xlsx)",
        data=datos_excel,
        file_name="auditoria_resultados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # También mantener la versión CSV (opcional)
    csv = df_res.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️ Descargar resultados en CSV",
        data=csv,
        file_name="auditoria_resultados.csv",
        mime="text/csv"
    )
else:
    st.info("Presiona el botón 🚀 para verificar el estado de actualización de los tableros.")
