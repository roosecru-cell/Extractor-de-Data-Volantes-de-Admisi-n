# Extractor de Órdenes de Admisión — Quálitas

Aplicación Streamlit que extrae campos clave de las Órdenes de Admisión de Quálitas (formato **Automóviles** y **Ajuste Express**) y los exporta a Excel.

## Campos extraídos
| Campo | Descripción |
|---|---|
| Fecha | Fecha del siniestro |
| Hora | Hora del reporte (Automóviles) |
| N° Reporte | Número de reporte Quálitas |
| N° Póliza | Número de póliza (Automóviles) |
| Nombre | Nombre del asegurado |
| Teléfono | Teléfono del cliente |
| E-mail | Correo electrónico |
| Marca | Marca del vehículo |
| Tipo | Tipo/submarca del vehículo |
| Modelo (Año) | Año del modelo |
| Color | Color del vehículo |
| Descripción de Daños | Daños a reparar |

## Uso local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy en Streamlit Cloud
1. Fork o sube este repo a GitHub
2. Entra a [share.streamlit.io](https://share.streamlit.io)
3. Selecciona el repo → rama `main` → archivo `app.py`
4. Clic en **Deploy**
