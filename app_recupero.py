import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Procesador de Facturas", page_icon="📄", layout="wide")

st.title("📄 Procesador de Facturas")
st.markdown("Subí el Excel exportado de Monday y obtené **una fila por factura** lista para Albor.")

def is_valid_especie(especie_str):
    if not especie_str or especie_str in ("Especie", "nan"):
        return False
    lower = especie_str.lower()
    if lower.startswith("20") or "to 20" in lower:
        return False
    return True

def parse_invoices(file):
    df = pd.read_excel(file, sheet_name=None, header=None)
    sheet = list(df.values())[0]
    invoices = []
    current_invoice = None
    in_subitems = False

    for i in range(3, len(sheet)):
        row = sheet.iloc[i]
        col0 = row.iloc[0]
        col0_str = str(col0).strip() if pd.notna(col0) else ""

        if col0_str == "Subitems":
            in_subitems = True
            continue

        if pd.isna(col0):
            if current_invoice is not None and in_subitems:
                especie = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
                campana = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
                campo   = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
                if is_valid_especie(especie):
                    if especie not in current_invoice["_especies"]:
                        current_invoice["_especies"].append(especie)
                if campana and campana not in ("Campaña", "nan", "Campo/Establecimiento"):
                    if not (campana.startswith("20") and len(campana) > 8):
                        if campana not in current_invoice["_campanas"]:
                            current_invoice["_campanas"].append(campana)
                if campo and campo not in ("Campo/Establecimiento", "nan"):
                    if campo not in current_invoice["_campos"]:
                        current_invoice["_campos"].append(campo)
            continue

        fecha_carga = row.iloc[2]
        cuit        = row.iloc[3]
        try:
            if float(str(cuit).replace(",", "").strip()) == 0:
                continue
        except (ValueError, TypeError):
            continue

        if pd.notna(fecha_carga) and pd.notna(cuit):
            if current_invoice is not None:
                current_invoice["Especie"]   = " / ".join(current_invoice.pop("_especies"))
                current_invoice["_campanas"] = current_invoice.pop("_campanas")
                current_invoice["_campos"]   = current_invoice.pop("_campos")
                invoices.append(current_invoice)

            in_subitems = False

            # Fecha: usar col13 (fecha vencimiento) que es la que aparece en el Excel destino
            fecha_raw = row.iloc[13]

            # Subelementos (col1) — el detalle con CPEs separados por coma
            subelementos = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""

            current_invoice = {
                "Proveedor":        str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else "",
                "Socio":            str(row.iloc[10]).strip() if pd.notna(row.iloc[10]) else "",
                "Registro Albor":   str(row.iloc[11]).strip() if pd.notna(row.iloc[11]) else "",
                "Fecha":            fecha_raw,
                "Monto":            "",   # columna vacía
                "Monto Total":      "",   # columna vacía
                "Especie":          "",   # se completa al final
                "Servicio":         str(row.iloc[6]).strip() if pd.notna(row.iloc[6]) else "",
                "Detalle":          subelementos,
                "_especies":        [],
                "_campanas":        [],
                "_campos":          [],
            }

    if current_invoice is not None:
        current_invoice["Especie"]   = " / ".join(current_invoice.pop("_especies"))
        current_invoice.pop("_campanas")
        current_invoice.pop("_campos")
        invoices.append(current_invoice)

    return pd.DataFrame(invoices)


def create_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Facturas para Albor"

    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill  = PatternFill("solid", start_color="1F4E79")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin         = Side(style="thin", color="CCCCCC")
    cell_border  = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt_fill     = PatternFill("solid", start_color="EBF3FB")

    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center_align
        cell.border    = cell_border
    ws.row_dimensions[1].height = 30

    for row_idx, row_data in enumerate(df.itertuples(index=False), 2):
        fill = alt_fill if row_idx % 2 == 0 else None
        for col_idx, value in enumerate(row_data, 1):
            if isinstance(value, (list, tuple)):
                value = ", ".join(str(v) for v in value)
            if isinstance(value, str) and value.strip() == "":
                value = None
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font      = Font(name="Arial", size=10)
            cell.border    = cell_border
            cell.alignment = left_align
            if fill:
                cell.fill = fill

    col_widths = {
        "Proveedor":      30,
        "Socio":          16,
        "Registro Albor": 20,
        "Fecha":          14,
        "Monto":          14,
        "Monto Total":    14,
        "Especie":        16,
        "Servicio":       22,
        "Detalle":        60,
    }
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 18)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── UI ───────────────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader("📂 Subí el archivo Excel de Monday", type=["xlsx"])

if uploaded_file:
    with st.spinner("Procesando..."):
        try:
            result_df = parse_invoices(uploaded_file)
            st.success(f"✅ Se procesaron **{len(result_df)} facturas** correctamente.")

            st.subheader("Vista previa")
            st.dataframe(result_df, use_container_width=True, height=420)

            excel_bytes = create_excel(result_df)
            st.download_button(
                label="⬇️ Descargar Excel para Albor",
                data=excel_bytes,
                file_name="facturas_albor.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
            st.exception(e)
else:
    st.info("👆 Subí el Excel exportado de Monday para comenzar.")

st.markdown("---")
st.caption("Procesador de Facturas · Bioceres")
