import streamlit as st
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TEAL  = "FF006B6B"
LTEAL = "FFE0F4F4"
WHITE = "FFFFFFFF"
RED   = "FFC00000"

COLS = [
    "N° Reporte", "N° Póliza",
    "Marca", "Tipo", "Modelo (Año)",
    "Aplica Deducible", "Placas",
    "Nombre", "Teléfono", "E-mail",
    "Fecha", "Descripción de Daños",
    "Hora", "Color",
]

MARCAS = (
    "HONDA|NISSAN|MAZDA|FORD|CHEVROLET|KIA|TOYOTA|VW|VOLKSWAGEN|CHRYSLER|DODGE|JEEP|"
    "RAM|SEAT|RENAULT|MITSUBISHI|SUZUKI|SUBARU|VOLVO|BMW|MERCEDES|AUDI|PEUGEOT|FIAT|"
    "ACURA|INFINITI|LEXUS|CADILLAC|BUICK|GMC|LINCOLN|HYUNDAI|CHERY|MG|BYD|MINI|LAND"
)
COLORES = (
    "NEGRO|BLANCO|ROJO|AZUL|GRIS|PLATA|PLATEADO|VERDE|AMARILLO|NARANJA|"
    "CAFE|CAFÉ|MORADO|BEIGE|VINO|DORADO|ROSA|GUINDA|BRONCE|PERLA"
)

def clean(txt):
    return re.sub(r"\s+", " ", txt or "").strip()

def first_match(pattern, text, group=1, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return clean(m.group(group)) if m else ""

def parse_automoviles(text, lines):
    fecha = hora = poliza = ""
    for line in lines:
        mf = re.search(r"(\d{2}/\d{2}/\d{4})", line)
        mh = re.search(r"(\d{1,2}:\d{2})\s*HRS", line, re.IGNORECASE)
        if mf and mh:
            fecha = mf.group(1)
            hora  = mh.group(1) + " HRS"
            mp = re.search(r"\s(\d{10})\s+\d{6}\s+\d{4}", line)
            if mp:
                poliza = mp.group(1)
            break

    reporte = first_match(r"N[°º]\.\s*REPORTE\s+(\d+)", text)

    nombre = tel = ""
    for i, line in enumerate(lines):
        if "NOMBRE O RAZ" in line.upper() and "CLIENTE" in line.upper():
            nxt = lines[i + 1] if i + 1 < len(lines) else ""
            mt = re.search(r"(\d{2}\s\d{3,4}\s\d{4})\s*$", nxt)
            if mt:
                tel    = re.sub(r"\s", "", mt.group(1))
                nombre = clean(nxt[:mt.start()])
            else:
                nombre = clean(nxt)
            break

    email = ""
    for i, line in enumerate(lines):
        if "E-MAIL" in line.upper() and i+1 < len(lines):
            nxt = lines[i+1].strip()
            m = re.search(r"([\w.\-+]+@[\w.\-]+\.\w+)", nxt)
            if m and "qualitas" not in m.group(1).lower():
                email = m.group(1)
                break

    marca = tipo = modelo = color = ""
    for line in lines:
        m = re.match(
            rf"^({MARCAS})\s+(.+?)\s+(\d{{4}})\s*$",
            line.strip(), re.IGNORECASE
        )
        if m:
            marca  = m.group(1).upper()
            tipo_raw = clean(m.group(2))
            # Quitar la marca si se repite al inicio del tipo (ej: "SUZUKI SWIFT" -> "SWIFT BOOSTERJET")
            if tipo_raw.upper().startswith(marca):
                tipo = tipo_raw[len(marca):].strip()
            else:
                tipo = tipo_raw
            modelo = m.group(3)
            break

    for line in lines:
        if "COLOR" in line.upper():
            m = re.search(rf"COLOR\s+(?:/COLOR\s+)?({COLORES})\b", line, re.IGNORECASE)
            if m:
                color = m.group(1).upper()
                break

    desc = ""
    for i, line in enumerate(lines):
        if "DESCRIPCI" in line.upper() and "REPARAR" in line.upper():
            parts = []
            for j in range(i + 1, min(i + 6, len(lines))):
                l = re.sub(r"\*\*.*?\*\*", "", lines[j]).strip()
                if not l or "VÁLIDO" in l.upper() or "VALID" in l.upper():
                    break
                parts.append(l)
            desc = " ".join(parts)
            break

    # Placas Automóviles: "SERIE PLACAS /LICENSE PLATES XYZ123 AUTOMATIC MANUAL"
    placas = ""
    for line in lines:
        if "PLACAS" in line.upper() and "LICENSE" in line.upper():
            m = re.search(r"PLATES\s+([A-Z0-9]{1,10})\s", line, re.IGNORECASE)
            if m:
                placas = m.group(1).upper()
                break

    # Aplica Deducible — Automóviles: detectar por línea de monto
    # "$ 267,000 5.00 % $ 13350" → SI (monto + cantidad final)
    # "$ V. COM. 3.00 % $"       → SI (valor comercial con %)
    # "$ V. 5.00 % $"            → NO (valor factura sin cantidad)
    # "$ % $"                    → NO (vacío)
    aplica_ded = ""
    for line in lines:
        s = line.strip()
        if not s.startswith("$"): continue
        if re.match(r"^\$\s+\d[\d,\.]+\s+[\d,\.]+\s*%\s+\$\s+[\d,\.]+", s):
            aplica_ded = "SI"; break
        if re.search(r"V\.\s*COM\.?\s+[\d,\.]+\s*%", s):
            aplica_ded = "SI"; break
        if re.match(r"^\$\s+V[\.\s]", s) and re.search(r"%\s*\$\s*$", s):
            aplica_ded = "NO"; break
        if re.match(r"^\$\s+%\s+\$\s*$", s):
            aplica_ded = "NO"; break

    return {"Fecha": fecha, "Hora": hora, "N° Reporte": reporte,
            "N° Póliza": poliza, "Nombre": nombre, "Teléfono": tel,
            "E-mail": email, "Marca": marca, "Tipo": tipo,
            "Modelo (Año)": modelo, "Color": color, "Placas": placas,
            "Aplica Deducible": aplica_ded, "Descripción de Daños": desc}

def parse_express(text, lines):
    fecha = ""
    m = re.search(r"FECHA\s+(\d{4}-\d{2}-\d{2})", text)
    if m:
        p = m.group(1).split("-")
        fecha = f"{p[2]}/{p[1]}/{p[0]}"

    reporte = first_match(r"N[°º]\.\s*REPORTE\s+(\d+)", text)

    nombre = ""
    for i, line in enumerate(lines):
        if "FIRMA DEL CONDUCTOR" in line.upper():
            # El nombre está en la misma línea antes de la firma del ajustador
            # Formato: "DIANA RODRIGUEZ BERISTAIN CARLOS ALFREDO MARTINEZ MARTINEZ"
            # Tomamos la línea anterior que sea solo nombre del conductor
            cand = lines[i - 1] if i > 0 else ""
            # Si la línea tiene más de ~35 chars puede mezclar dos nombres
            # Intentar extraer solo el nombre del conductor de la misma línea "FIRMA DEL CONDUCTOR"
            # que no existe, así que buscamos en la línea anterior
            if re.search(r"[A-ZÁÉÍÓÚÑ]{3,}\s+[A-ZÁÉÍÓÚÑ]{3,}", cand):
                words = cand.strip().split()
                # Nombre típico mexicano: 3 palabras (nombre + 2 apellidos)
                nombre = " ".join(words[:3]) if len(words) > 3 else cand.strip()
            break

    tel = ""
    for line in lines:
        if "CON DIRECCION" in line.upper() or "TELEFONO" in line.upper():
            mt = re.search(r"(?:TELEFONO|TEL)[^\d]*([\d\s()\-\.ext]+\d)", line, re.IGNORECASE)
            if mt:
                digits = re.sub(r"[^\d]", "", mt.group(1))
                tel = digits[:10]  # máximo 10 dígitos
                break

    marca = tipo = modelo = color = ""
    for line in lines:
        m2 = re.match(
            rf"^({MARCAS})\s+(.+?)\s+(\d{{4}})\s+\d",
            line.strip(), re.IGNORECASE
        )
        if m2:
            marca  = m2.group(1).upper()
            tipo_raw = clean(m2.group(2))
            if tipo_raw.upper().startswith(marca):
                tipo = tipo_raw[len(marca):].strip()
            else:
                tipo = tipo_raw
            modelo = m2.group(3)
            break

    for line in lines:
        m3 = re.search(
            rf"[A-Z0-9]{{10,}}\s+({COLORES})\s+[A-Z0-9]+",
            line, re.IGNORECASE
        )
        if m3:
            color = m3.group(1).upper()
            break

    desc = ""
    for i, line in enumerate(lines):
        if "DESCRIPCI" in line.upper() and "DA" in line.upper():
            parts = []
            for j in range(i + 1, min(i + 5, len(lines))):
                l = lines[j].strip()
                if not l or "PREEX" in l.upper():
                    break
                parts.append(l)
            desc = " ".join(parts)
            break

    # Placas Express: línea siguiente a "N°. DE SERIE COLOR PLACAS TRANSMISIÓN"
    placas = ""
    for i, line in enumerate(lines):
        if "PLACAS" in line.upper() and "SERIE" in line.upper() and i+1 < len(lines):
            data_line = lines[i+1]
            # Formato: "SERIE COLOR PLACAS TRANSMISION"
            # COLOR es 1 palabra, PLACAS es alfanumérico 5-8 chars, TRANSMISION al final
            m = re.search(
                rf"[A-Z0-9]{{10,}}\s+(?:{COLORES})\s+([A-Z0-9]{{3,10}})\s",
                data_line, re.IGNORECASE
            )
            if m:
                placas = m.group(1).upper()
            break

    # Aplica Deducible — Express: X explícita en texto
    # "NO SI X FIJO % ADM X" → SI aplica
    # "NO X SI FIJO % X ADM" → NO aplica
    aplica_ded = ""
    for line in lines:
        if ("NO" in line.upper() and "SI" in line.upper() and
                ("FIJO" in line.upper() or "ADM" in line.upper())):
            if re.search(r"NO\s+X\s+SI", line, re.IGNORECASE):
                aplica_ded = "NO"
            elif re.search(r"NO\s+SI\s+X", line, re.IGNORECASE):
                aplica_ded = "SI"
            break

    return {"Fecha": fecha, "Hora": "", "N° Reporte": reporte,
            "N° Póliza": "", "Nombre": nombre, "Teléfono": tel,
            "E-mail": "", "Marca": marca, "Tipo": tipo,
            "Modelo (Año)": modelo, "Color": color, "Placas": placas,
            "Aplica Deducible": aplica_ded, "Descripción de Daños": desc}

def extract_from_pdf(uploaded_file):
    with pdfplumber.open(uploaded_file) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    lines = text.split("\n")
    is_express = bool(re.search(r"AJUSTE\s*EXPRESS", text, re.IGNORECASE))
    data = parse_express(text, lines) if is_express else parse_automoviles(text, lines)
    data["_tipo"] = "Express" if is_express else "Automóviles"
    return data

def make_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Admisiones Quálitas"
    hdr_fill = PatternFill("solid", fgColor=TEAL)
    hdr_font = Font(bold=True, color="FFFFFFFF", name="Arial", size=10)
    row_even = PatternFill("solid", fgColor=LTEAL)
    row_odd  = PatternFill("solid", fgColor=WHITE)
    border   = Border(
        left=Side(style="thin", color="FFCCCCCC"),
        right=Side(style="thin", color="FFCCCCCC"),
        top=Side(style="thin", color="FFCCCCCC"),
        bottom=Side(style="thin", color="FFCCCCCC"),
    )
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    ws.merge_cells("A1:N1")
    tc = ws["A1"]
    tc.value     = "Órdenes de Admisión — Quálitas"
    tc.font      = Font(bold=True, color="FFFFFFFF", name="Arial", size=13)
    tc.fill      = PatternFill("solid", fgColor=RED)
    tc.alignment = center
    ws.row_dimensions[1].height = 28

    for ci, col in enumerate(COLS, 1):
        c = ws.cell(row=2, column=ci, value=col)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = center; c.border = border
    ws.row_dimensions[2].height = 22

    for ri, row in enumerate(rows, 3):
        fill = row_even if ri % 2 == 0 else row_odd
        for ci, col in enumerate(COLS, 1):
            c = ws.cell(row=ri, column=ci, value=row.get(col, ""))
            c.fill = fill; c.border = border
            c.font = Font(name="Arial", size=9)
            c.alignment = left if ci in (8, 10, 12) else center

    for i, w in enumerate([15,16,12,24,14,14,12,28,16,28,12,50,10,12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── UI ────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Admisiones Quálitas", page_icon="🚗", layout="wide")
st.markdown("""
<style>
.main{background-color:#F5FAFA}h1{color:#006B6B}
.stButton>button{background-color:#006B6B;color:white;border-radius:6px;font-weight:bold;padding:.4rem 1.2rem}
.stButton>button:hover{background-color:#004F4F}.block-container{padding-top:2rem}
</style>""", unsafe_allow_html=True)

c1, c2 = st.columns([1,6])
with c1: st.markdown("## 🚗")
with c2:
    st.title("Extractor de Órdenes de Admisión — Quálitas")
    st.caption("Sube una o más órdenes en PDF (Automóviles o Ajuste Express) y descarga los datos en Excel")
st.divider()

uploaded = st.file_uploader(
    "Arrastra aquí tus PDFs de Quálitas", type=["pdf"],
    accept_multiple_files=True,
    help="Formato Orden de Admisión Automóviles y Ajuste Express"
)

if uploaded:
    rows, errors = [], []
    with st.spinner("Procesando PDFs…"):
        for f in uploaded:
            try:
                data = extract_from_pdf(f)
                tipo = data.pop("_tipo","")
                row  = {col: data.get(col,"") for col in COLS}
                row["_tipo"] = tipo
                rows.append(row)
            except Exception as e:
                errors.append(f"**{f.name}**: {e}")

    for err in errors: st.error(err)

    if rows:
        st.success(f"✅ {len(rows)} PDF(s) procesados correctamente")
        import pandas as pd
        df = pd.DataFrame(rows)
        tipo_col = df.pop("_tipo")
        st.markdown("#### Vista previa")
        st.dataframe(df, use_container_width=True, height=min(200+len(rows)*38,500))

        with st.expander("🔍 Ver detalle por orden"):
            for i, row in enumerate(rows):
                badge = "🔵 Automóviles" if tipo_col.iloc[i]=="Automóviles" else "🟣 Ajuste Express"
                st.markdown(f"**{badge} — Reporte {row.get('N° Reporte',i+1)}**")
                c3 = st.columns(3)
                for j,(k,v) in enumerate([(k,v) for k,v in row.items() if k!="_tipo"]):
                    c3[j%3].markdown(f"- **{k}:** {v or '—'}")
                st.divider()

        excel_buf = make_excel(rows)
        st.download_button(
            label="⬇️  Descargar Excel", data=excel_buf,
            file_name="admisiones_qualitas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
else:
    st.info("👆 Sube uno o más PDFs de Quálitas para comenzar.")
    st.markdown("""
**Formatos compatibles:**
- 📄 Orden de Admisión Automóviles
- 📄 Orden de Admisión Ajuste Express

**Campos extraídos:**
`Fecha · Hora · N° Reporte · N° Póliza · Nombre · Teléfono · E-mail · Marca · Tipo · Modelo · Color · Descripción de Daños`
""")

