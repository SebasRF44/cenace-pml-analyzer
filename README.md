# ⚡ CENACE PML Analyzer

Aplicación web para descarga y análisis de Precios Marginales Locales (PML) del Sistema Eléctrico Nacional mexicano (CENACE).

**Recurrent Energy / Canadian Solar — México**
Sebastian Roldan (SRF)

## Funcionalidades

- 📊 Descarga masiva de datos PML del CENACE (paralelizada con ThreadPoolExecutor)
- 📑 Generación de Excel con formato corporativo Recurrent Energy
- 🎯 Soporte para sistemas SIN, BCA, BCS y procesos MTR/MDA
- 📈 (Próximamente) Análisis BESS scoring por caso de uso
- 🗺️ (Próximamente) Geocodificación de nodos con OpenStreetMap

## Demo en línea

Despliegue: **[cenace-pml-srf.streamlit.app](https://cenace-pml-srf.streamlit.app)**

## Tecnología

- **Frontend**: Streamlit
- **Backend**: Python (pandas, openpyxl, requests)
- **Hosting**: Streamlit Community Cloud
- **Datos**: API SW-PML del CENACE

## Estructura

```
.
├── app.py                       # App Streamlit principal
├── requirements.txt             # Librerías Python
├── .streamlit/config.toml       # Tema visual RE
└── data/catalogo_nodos.xlsx     # Catálogo CENACE oficial
```

## Uso local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre http://localhost:8501

---

*Subsidiary of Canadian Solar*
